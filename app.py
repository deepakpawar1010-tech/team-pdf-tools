"""Local Flask server for Team PDF Tools."""
from __future__ import annotations

import io
import os
import re
import sys
import uuid
import shutil
import logging
import zipfile
import base64
import gc
import tempfile
import threading
import subprocess
from pathlib import Path

import pymupdf
fitz = pymupdf
from PIL import Image
from flask import Flask, jsonify, render_template, request, send_file, after_this_request
from pypdf import PdfReader, PdfWriter

logger = logging.getLogger(__name__)
word_lock = threading.Lock()

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

app = Flask(__name__)
# Allow large streaming uploads without Flask rejecting them
app.config["MAX_CONTENT_LENGTH"] = None


@app.errorhandler(413)
def handle_too_large(e):
    return jsonify({"error": "File exceeds upload limit (413 Request Entity Too Large)."}), 413


def valid_pdf(upload):
    if not upload or not upload.filename.lower().endswith(".pdf"):
        raise ValueError("Please upload a PDF file.")


def valid_docx(upload):
    if not upload or not (upload.filename.lower().endswith(".docx") or upload.filename.lower().endswith(".doc")):
        raise ValueError("Please upload a Microsoft Word (.docx) document.")


def output_name(filename: str, suffix: str) -> str:
    stem = Path(filename).stem
    safe_stem = re.sub(r"[^\w.-]+", "-", stem).strip("-") or "document"
    return f"{safe_stem}-{suffix}.pdf"


def output_docx_name(filename: str, suffix: str = "converted") -> str:
    stem = Path(filename).stem
    safe_stem = re.sub(r"[^\w.-]+", "-", stem).strip("-") or "document"
    return f"{safe_stem}-{suffix}.docx"


@app.get("/")
def home():
    return render_template("home.html")


COMING_SOON = {
    "pdf-to-markdown": {
        "title": "PDF to Markdown",
        "description": "Extract structured text, headers, lists, and tables into clean, LLM-ready Markdown (.md) documents.",
        "badge": "Markdown / AI Ready",
        "icon": "markdown"
    },
    "pdf-to-excel": {
        "title": "PDF to Excel",
        "description": "Automatically detect tabular data and financial statements in PDFs and export them directly to Excel (.xlsx) spreadsheets.",
        "badge": "Spreadsheet OCR",
        "icon": "excel"
    }
}


@app.get("/healthz")
@app.get("/api/health")
def healthz():
    return jsonify({"status": "healthy", "service": "team-pdf-tools"}), 200


@app.get("/katex-viewer")
@app.get("/katex")
def katex_page():
    return render_template("katex.html")


@app.get("/<tool>")
def tool_page(tool: str):
    if tool in {"katex-viewer", "katex"}:
        return render_template("katex.html")
    if tool in {"merge", "split", "compress", "pdf-to-word", "word-to-pdf", "worksheet-splitter", "oly-ete-splitter", "markdown-to-pdf", "md-to-pdf"}:
        return render_template("tool.html", tool=tool)
    if tool in COMING_SOON:
        return render_template("coming_soon.html", tool=tool, info=COMING_SOON[tool])
    return "Page not found", 404


@app.post("/api/info")
def info():
    input_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            input_path = in_tmp.name
        upload.save(input_path)
        doc = fitz.open(input_path)
        page_count = len(doc)
        doc.close()
        return jsonify({"pages": page_count, "name": upload.filename})
    except Exception as error:
        return jsonify({"error": str(error)}), 400
    finally:
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass


@app.post("/api/worksheet-info")
def worksheet_info():
    input_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        preset = request.form.get("preset", "olympiad").lower()
        stop_at_synopsis = request.form.get("stop_at_synopsis", "true").lower() == "true"
        include_key = request.form.get("include_key", "true").lower() == "true"

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            upload.save(in_tmp.name)
            input_path = in_tmp.name

        from worksheet_splitter import get_worksheet_info
        data = get_worksheet_info(
            Path(input_path),
            stop_at_synopsis=stop_at_synopsis,
            preset=preset,
            include_key=include_key,
        )
        data["name"] = upload.filename
        return jsonify(data)
    except Exception as error:
        return jsonify({"error": str(error)}), 400
    finally:
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass


@app.post("/api/previews")
def previews():
    """Render lightweight page thumbnails for the browser UI.

    The PDF is kept only in memory for this request. For merge mode only the
    first page is rendered for each file; for split mode requested page
    numbers can be supplied as a comma-separated list (or all pages by
    default).
    """
    try:
        uploads = request.files.getlist("files") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads:
            raise ValueError("Please upload a PDF file.")
        mode = request.form.get("mode", "split")
        requested = request.form.get("pages", "")
        requested_pages = None
        if requested.strip():
            requested_pages = {int(x.strip()) for x in requested.split(",") if x.strip()}
            if any(p < 1 for p in requested_pages):
                raise ValueError("Invalid page number.")

        output = []
        for file_index, upload in enumerate(uploads):
            valid_pdf(upload)
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    upload.save(tmp_file.name)
                    tmp_path = tmp_file.name

                document = fitz.open(tmp_path)
                page_count = document.page_count
                if mode == "merge":
                    page_numbers = [1] if page_count else []
                elif requested_pages is None:
                    page_numbers = list(range(1, min(13, page_count + 1)))
                else:
                    page_numbers = sorted(p for p in requested_pages if p <= page_count)

                for page_number in page_numbers:
                    page = document.load_page(page_number - 1)
                    rect = page.rect
                    longest = max(rect.width, rect.height) or 1
                    scale = min(1.0, 240 / longest)
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    jpeg = pixmap.tobytes("jpeg", jpg_quality=65)
                    del pixmap
                    output.append({
                        "fileIndex": file_index,
                        "page": page_number,
                        "pages": page_count,
                        "name": upload.filename,
                        "data": "data:image/jpeg;base64," + base64.b64encode(jpeg).decode("ascii"),
                    })
                document.close()
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass

        return jsonify({"previews": output})
    except Exception as error:
        return jsonify({"error": str(error)}), 400


@app.post("/api/split")
def split_pdf():
    input_path = None
    output_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        ranges = request.form.get("ranges", "")

        # Stream upload directly to disk using 1MB chunks (supports 500MB+ files with zero RAM bloating)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            input_path = in_tmp.name
            shutil.copyfileobj(upload.stream, in_tmp, length=1024 * 1024)

        # Open file directly from disk via OS memory-mapping (virtually 0 MB RAM)
        doc = fitz.open(input_path)
        total_pages = len(doc)

        selected_ranges = []
        for piece in ranges.split(","):
            piece = piece.strip()
            if not piece:
                continue
            if "-" in piece:
                parts = piece.split("-", 1)
                start = int(parts[0].strip())
                end = int(parts[1].strip())
            else:
                start = end = int(piece)
            if start > end:
                start, end = end, start
            if start < 1 or end > total_pages:
                raise ValueError(f"Use page ranges between 1 and {total_pages}.")
            selected_ranges.append((start, end))

        if not selected_ranges:
            raise ValueError("Please provide at least one valid page range.")

        combine = request.form.get("combine", "true").lower() == "true"
        if combine:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                output_path = out_tmp.name

            out_doc = fitz.open()
            for start, end in selected_ranges:
                out_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
            # Ultra-fast save: garbage=1 cleans xrefs without quadratic duplicate-stream scanning;
            # deflate=False copies existing compressed streams directly in C in milliseconds!
            out_doc.save(output_path, garbage=1, deflate=False)
            out_doc.close()
            doc.close()
            del out_doc
            del doc
            gc.collect()

            @after_this_request
            def cleanup_split_combined(response):
                try:
                    if input_path and os.path.exists(input_path):
                        os.remove(input_path)
                    if output_path and os.path.exists(output_path):
                        os.remove(output_path)
                except Exception:
                    pass
                return response

            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_name(upload.filename, "selected-pages"),
                mimetype="application/pdf",
            )

        # Unmerged: Create individual PDF files bundled into a ZIP archive directly in memory
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
            output_path = zip_tmp.name

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_STORED) as zf:
            for number, (start, end) in enumerate(selected_ranges, start=1):
                sub_doc = fitz.open()
                sub_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
                sub_bytes = sub_doc.tobytes(garbage=1, deflate=False)
                sub_doc.close()
                del sub_doc
                suffix = f"page-{start}" if start == end else f"pages-{start}-{end}"
                file_entry_name = output_name(upload.filename, f"{number:02d}_{suffix}")
                zf.writestr(file_entry_name, sub_bytes)

        doc.close()
        del doc
        gc.collect()

        @after_this_request
        def cleanup_split_zip(response):
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name(upload.filename, "split-files").replace(".pdf", ".zip"),
            mimetype="application/zip",
        )
    except Exception as error:
        logger.exception("Split PDF failed: %s", error)
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        return jsonify({"error": str(error)}), 400


@app.post("/api/merge")
def merge_pdf():
    input_paths = []
    output_path = None
    try:
        uploads = request.files.getlist("files")
        if len(uploads) < 2:
            raise ValueError("Choose at least two PDF files to merge.")

        out_doc = fitz.open()
        for upload in uploads:
            valid_pdf(upload)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                tmp_path = in_tmp.name
            upload.save(tmp_path)
            input_paths.append(tmp_path)
            sub_doc = fitz.open(tmp_path)
            out_doc.insert_pdf(sub_doc)
            sub_doc.close()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
            output_path = out_tmp.name

        out_doc.save(output_path, garbage=3, deflate=True)
        out_doc.close()
        del out_doc
        gc.collect()

        @after_this_request
        def cleanup_merge(response):
            for p in input_paths:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
            try:
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name="merged.pdf",
            mimetype="application/pdf",
        )
    except Exception as error:
        logger.exception("Merge PDF failed: %s", error)
        for p in input_paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        return jsonify({"error": str(error)}), 400


def optimize_pdf_images(doc: fitz.Document, dpi_target: int = 220, quality: int = 85) -> int:
    """
    Optimizes embedded raster images inside a PDF:
    - Downsamples images whose DPI exceeds dpi_target on the page using Lanczos resampling.
    - Preserves 100% of vector text, fonts, mathematical symbols, equations, and annotations.
    - Replaces image streams via page.replace_image(xref, stream=...).
    - Shared images across pages are deduplicated and processed only once.
    Returns the count of optimized images.
    """
    processed_xrefs = set()
    optimized_count = 0

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        image_infos = page.get_image_info(xrefs=True)
        if not image_infos:
            continue

        for img_info in image_infos:
            xref = img_info.get("xref")
            if not xref or xref in processed_xrefs:
                continue
            processed_xrefs.add(xref)

            bbox = img_info.get("bbox")
            if not bbox:
                continue

            # Display dimensions in inches (72 points = 1 inch)
            disp_w_in = max((bbox[2] - bbox[0]) / 72.0, 0.1)
            disp_h_in = max((bbox[3] - bbox[1]) / 72.0, 0.1)

            orig_w = img_info.get("width", 0)
            orig_h = img_info.get("height", 0)
            if orig_w <= 0 or orig_h <= 0:
                continue

            # Current effective DPI of the image as displayed on this page
            current_dpi = max(orig_w / disp_w_in, orig_h / disp_h_in)

            try:
                extracted = doc.extract_image(xref)
                if not extracted or not extracted.get("image"):
                    continue

                raw_bytes = extracted["image"]
                ext = extracted.get("ext", "").lower()

                # If current DPI is already within target range and image is already a compressed JPEG, keep it
                needs_downsample = current_dpi > (dpi_target * 1.15)
                if not needs_downsample and ext in ("jpeg", "jpg"):
                    continue

                pil_img = Image.open(io.BytesIO(raw_bytes))

                if needs_downsample:
                    target_w = min(max(int(disp_w_in * dpi_target), 50), orig_w)
                    target_h = min(max(int(disp_h_in * dpi_target), 50), orig_h)
                    pil_img = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)

                out_buf = io.BytesIO()
                has_alpha = pil_img.mode in ("RGBA", "LA") or (pil_img.mode == "P" and "transparency" in pil_img.info)
                if has_alpha:
                    pil_img.save(out_buf, format="PNG", optimize=True)
                else:
                    if pil_img.mode != "RGB":
                        pil_img = pil_img.convert("RGB")
                    pil_img.save(out_buf, format="JPEG", quality=quality, optimize=True)

                new_bytes = out_buf.getvalue()
                if len(new_bytes) < len(raw_bytes):
                    page.replace_image(xref, stream=new_bytes)
                    optimized_count += 1
            except Exception as ex:
                logger.debug("Skipping image optimization for xref %s: %s", xref, ex)

    return optimized_count


@app.post("/api/compress")
def compress_pdf():
    input_paths = []
    output_path = None
    temp_compressed_files = []
    try:
        # Support both multiple files ('files') and single file ('file')
        uploads = request.files.getlist("files") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads or not any(getattr(u, "filename", None) for u in uploads):
            raise ValueError("Please upload at least one PDF file to compress.")

        for upload in uploads:
            valid_pdf(upload)

        preset = request.form.get("quality", "best").lower()
        settings = {
            "small": (120, 68),
            "balanced": (160, 78),
            "best": (220, 85),
            "laptop": (220, 85),
            "ultra": (300, 92),
        }
        dpi_target, quality = settings.get(preset, settings["best"])

        # Case 1: Single file upload
        if len(uploads) == 1:
            upload = uploads[0]
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                input_path = in_tmp.name
                shutil.copyfileobj(upload.stream, in_tmp, length=1024 * 1024)
            input_paths.append(input_path)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                output_path = out_tmp.name

            doc = fitz.open(input_path)
            optimize_pdf_images(doc, dpi_target=dpi_target, quality=quality)
            doc.save(output_path, garbage=3, deflate=True)
            doc.close()
            del doc
            gc.collect()

            @after_this_request
            def cleanup_single(response):
                for p in input_paths:
                    try:
                        if os.path.exists(p):
                            os.remove(p)
                    except Exception:
                        pass
                try:
                    if output_path and os.path.exists(output_path):
                        os.remove(output_path)
                except Exception:
                    pass
                return response

            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_name(upload.filename, "compressed"),
                mimetype="application/pdf"
            )

        # Case 2: Batch conversion (Multiple PDF files) -> Output as ZIP archive
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
            output_path = zip_tmp.name

        seen_names = {}
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for upload in uploads:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                    in_path = in_tmp.name
                    shutil.copyfileobj(upload.stream, in_tmp, length=1024 * 1024)
                input_paths.append(in_path)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as comp_tmp:
                    comp_path = comp_tmp.name
                temp_compressed_files.append(comp_path)

                doc = fitz.open(in_path)
                optimize_pdf_images(doc, dpi_target=dpi_target, quality=quality)
                doc.save(comp_path, garbage=3, deflate=True)
                doc.close()
                del doc

                base_entry_name = output_name(upload.filename, "compressed")
                if base_entry_name in seen_names:
                    seen_names[base_entry_name] += 1
                    stem = Path(base_entry_name).stem
                    entry_name = f"{stem}_{seen_names[base_entry_name]}.pdf"
                else:
                    seen_names[base_entry_name] = 1
                    entry_name = base_entry_name

                zf.write(comp_path, arcname=entry_name)
                try:
                    os.remove(comp_path)
                except Exception:
                    pass

        gc.collect()

        @after_this_request
        def cleanup_batch(response):
            for p in input_paths:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
            for p in temp_compressed_files:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
            try:
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name="compressed-pdfs-bundle.zip",
            mimetype="application/zip"
        )
    except Exception as error:
        logger.exception("Compress PDF failed: %s", error)
        for p in input_paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        for p in temp_compressed_files:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        return jsonify({"error": str(error)}), 400


@app.post("/api/pdf-to-word")
def pdf_to_word():
    input_path = None
    output_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)

        if Converter is None:
            return jsonify({"error": "pdf2docx is not installed on the server."}), 500

        # Stream upload to disk temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            upload.save(in_tmp.name)
            input_path = in_tmp.name

        # Create output .docx temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as out_tmp:
            output_path = out_tmp.name

        # Parse optional page range
        pages_param = request.form.get("pages", "").strip()
        pages_list = None
        if pages_param:
            try:
                doc = fitz.open(input_path)
                total_pages = doc.page_count
                doc.close()
                indices = set()
                for part in pages_param.split(","):
                    p = part.strip()
                    if "-" in p:
                        s, e = p.split("-", 1)
                        s, e = int(s.strip()), int(e.strip())
                        if s > e:
                            s, e = e, s
                        for i in range(max(1, s), min(total_pages, e) + 1):
                            indices.add(i - 1)
                    elif p.isdigit():
                        val = int(p)
                        if 1 <= val <= total_pages:
                            indices.add(val - 1)
                if indices:
                    pages_list = sorted(list(indices))
            except Exception as parse_err:
                print("Page range parse warning:", parse_err)

        language = request.form.get("language", "auto").strip().lower()
        handled_by_regional = False
        if language != "english":
            try:
                from regional_converters import convert_regional_pdf
                handled_by_regional = convert_regional_pdf(
                    input_path, output_path, language=language, original_filename=upload.filename
                )
            except Exception as reg_err:
                logger.warning("Regional converter exception: %s", reg_err)

        if not handled_by_regional:
            if not Converter:
                raise RuntimeError("PDF to Word conversion module (pdf2docx) is not installed.")
            cv = Converter(input_path)
            if pages_list:
                cv.convert(output_path, pages=pages_list)
            else:
                cv.convert(output_path)
            cv.close()
            del cv
            gc.collect()

        @after_this_request
        def cleanup(response):
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_docx_name(upload.filename, "converted"),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as error:
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        return jsonify({"error": str(error)}), 400


def find_libreoffice_bin():
    candidates = [
        "libreoffice",
        "soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]
    for cmd in candidates:
        if shutil.which(cmd) or (os.path.exists(cmd) and os.path.isfile(cmd)):
            return cmd
    return None


def convert_docx_with_word_com(input_docx: str, output_pdf: str) -> bool:
    """Uses native Microsoft Word background COM automation on Windows (File -> Save As -> PDF)."""
    if sys.platform != "win32":
        return False

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False

    with word_lock:
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            abs_in = str(Path(input_docx).resolve())
            abs_out = str(Path(output_pdf).resolve())
            doc = word.Documents.Open(abs_in, ReadOnly=True)
            # wdFormatPDF = 17
            doc.SaveAs2(abs_out, FileFormat=17)
            doc.Close(False)
            doc = None
            word.Quit()
            word = None
            return True
        except Exception as err:
            logger.warning("Word COM automation error: %s", err)
            raise
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def convert_docx_with_libreoffice(input_docx: str, output_pdf: str) -> bool:
    """Uses headless LibreOffice export for Linux/Docker/Render servers."""
    lo_bin = find_libreoffice_bin()
    if not lo_bin:
        return False

    out_dir = Path(output_pdf).parent.resolve()
    unique_id = uuid.uuid4().hex
    profile_dir = Path(tempfile.gettempdir()) / f"lo_profile_{unique_id}"
    profile_dir.mkdir(parents=True, exist_ok=True)
    profile_uri = profile_dir.as_uri()

    cmd = [
        lo_bin,
        f"-env:UserInstallation={profile_uri}",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(out_dir),
        str(Path(input_docx).resolve())
    ]

    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        stem = Path(input_docx).stem
        expected_pdf = out_dir / f"{stem}.pdf"
        if expected_pdf.exists() and expected_pdf.stat().st_size > 0:
            if str(expected_pdf.resolve()) != str(Path(output_pdf).resolve()):
                shutil.move(str(expected_pdf), str(output_pdf))
            return True
        logger.warning("LibreOffice returned code %s, stderr: %s", res.returncode, res.stderr)
        return False
    except Exception as err:
        logger.warning("LibreOffice conversion failed: %s", err)
        return False
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def convert_docx_to_pdf_high_fidelity(input_docx: str, output_pdf: str):
    """
    Multi-tier conversion engine:
    1. Microsoft Word COM ('Save As PDF') on Windows if MS Office is installed
    2. LibreOffice headless on Linux / Docker / Render / Windows
    3. PyMuPDF fallback
    """
    # 1. Try Microsoft Word native Save As PDF (100% fidelity)
    if sys.platform == "win32":
        try:
            if convert_docx_with_word_com(input_docx, output_pdf):
                if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
                    return
        except Exception as e:
            logger.warning("Word COM failed, trying fallback: %s", e)

    # 2. Try LibreOffice headless
    try:
        if convert_docx_with_libreoffice(input_docx, output_pdf):
            if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
                return
    except Exception as e:
        logger.warning("LibreOffice failed, trying fallback: %s", e)

    # 3. Fallback: PyMuPDF
    try:
        doc = fitz.open(input_docx)
        pdf_bytes = doc.convert_to_pdf()
        doc.close()
        del doc
        with open(output_pdf, "wb") as f:
            f.write(pdf_bytes)
        gc.collect()
        if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
            return
    except Exception as e:
        raise RuntimeError(f"All Word to PDF conversion engines failed: {e}")


@app.post("/api/word-to-pdf")
def word_to_pdf():
    input_path = None
    output_path = None
    try:
        upload = request.files.get("file")
        valid_docx(upload)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as in_tmp:
            upload.save(in_tmp.name)
            input_path = in_tmp.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
            output_path = out_tmp.name

        convert_docx_to_pdf_high_fidelity(input_path, output_path)

        gc.collect()

        @after_this_request
        def cleanup(response):
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name(upload.filename, "converted"),
            mimetype="application/pdf"
        )
    except Exception as error:
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        return jsonify({"error": str(error)}), 400



@app.post("/api/worksheet-info")
def worksheet_info_api():
    temp_files: list[str] = []
    try:
        uploads = request.files.getlist("files") or request.files.getlist("file") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads:
            raise ValueError("Please upload at least one PDF file.")

        preset = request.form.get("preset", "olympiad").lower()
        if preset == "olympiad":
            any_cbse = any("cbse" in (u.filename or "").lower() for u in uploads)
            any_ssc = any(("ssc" in (u.filename or "").lower() or "ap_" in (u.filename or "").lower() or "ts_" in (u.filename or "").lower()) for u in uploads)
            if any_cbse:
                preset = "cbse"
            elif any_ssc:
                preset = "ssc"

        stop_at_synopsis = request.form.get("stop_at_synopsis", "true").lower() == "true"
        include_key = request.form.get("include_key", "false").lower() == "true"

        from worksheet_splitter import get_worksheet_info

        results = []
        aggregate_pages = 0
        all_worksheets = []

        for upload in uploads:
            valid_pdf(upload)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                upload.save(in_tmp.name)
                temp_files.append(in_tmp.name)

            info = get_worksheet_info(
                Path(in_tmp.name),
                stop_at_synopsis=stop_at_synopsis,
                preset=preset,
                include_key=include_key,
            )
            aggregate_pages += info["total_pages"]
            results.append({
                "filename": upload.filename,
                "total_pages": info["total_pages"],
                "worksheets": info["worksheets"],
            })
            for ws in info["worksheets"]:
                ws_copy = dict(ws)
                ws_copy["source_file"] = upload.filename
                all_worksheets.append(ws_copy)

        return jsonify({
            "total_files": len(uploads),
            "total_pages": aggregate_pages,
            "worksheets": all_worksheets,
            "files": results,
        })
    except Exception as error:
        return jsonify({"error": str(error)}), 400
    finally:
        for tf in temp_files:
            if tf and os.path.exists(tf):
                try:
                    os.remove(tf)
                except Exception:
                    pass


@app.post("/api/worksheet-splitter")
@app.post("/api/oly-ete-splitter")
def worksheet_splitter_api():
    temp_files: list[str] = []
    try:
        uploads = request.files.getlist("files") or request.files.getlist("file") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads:
            raise ValueError("Please upload at least one PDF file.")

        preset = request.form.get("preset", "olympiad").lower()
        if preset == "olympiad":
            any_cbse = any("cbse" in (u.filename or "").lower() for u in uploads)
            any_ssc = any(("ssc" in (u.filename or "").lower() or "ap_" in (u.filename or "").lower() or "ts_" in (u.filename or "").lower()) for u in uploads)
            if any_cbse:
                preset = "cbse"
            elif any_ssc:
                preset = "ssc"

        stop_at_synopsis = request.form.get("stop_at_synopsis", "true").lower() == "true"
        include_key = request.form.get("include_key", "false").lower() == "true"
        crop_top = request.form.get("crop_top", "true").lower() == "true"

        pdf_inputs: list[tuple[str, Path]] = []
        for upload in uploads:
            valid_pdf(upload)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                upload.save(in_tmp.name)
                temp_files.append(in_tmp.name)
                pdf_inputs.append((upload.filename, Path(in_tmp.name)))

        from worksheet_splitter import split_pdfs_to_zip
        zip_buf, worksheets = split_pdfs_to_zip(
            pdf_inputs,
            stop_at_synopsis=stop_at_synopsis,
            preset=preset,
            include_key=include_key,
            crop_top=crop_top,
        )
        if not worksheets and preset == "olympiad":
            # Fallback to CBSE in case user did not switch tab
            zip_buf, worksheets = split_pdfs_to_zip(
                pdf_inputs,
                stop_at_synopsis=stop_at_synopsis,
                preset="cbse",
                include_key=include_key,
                crop_top=crop_top,
            )
            if worksheets:
                preset = "cbse"

        if not worksheets:
            if preset == "ssc":
                raise ValueError("No Multiple Choice Questions (MCQs) or Objective Question sections were detected.")
            elif preset == "cbse":
                raise ValueError("No CBSE Objective Exercises, Assessment Sheets, or Multiple Choice Questions were detected.")
            else:
                raise ValueError("No worksheet banners (CUQ or WORKSHEET headings) were detected.")

        gc.collect()

        suffix = "cbse-sheets" if preset == "cbse" else ("mcqs" if preset == "ssc" else "worksheets")
        if len(uploads) == 1:
            dl_name = output_name(uploads[0].filename, suffix).replace(".pdf", ".zip")
        else:
            dl_name = f"split_{suffix}_bundle.zip"

        return send_file(
            zip_buf,
            as_attachment=True,
            download_name=dl_name,
            mimetype="application/zip"
        )
    except Exception as error:
        return jsonify({"error": str(error)}), 400
    finally:
        for tf in temp_files:
            if tf and os.path.exists(tf):
                try:
                    os.remove(tf)
                except Exception:
                    pass


@app.post("/api/markdown-to-pdf")
@app.post("/api/md-to-pdf")
def markdown_to_pdf_api():
    try:
        markdown_text = ""
        filename = "document.md"

        upload = request.files.get("file")
        if upload and upload.filename:
            filename = upload.filename
            raw_bytes = upload.read()
            if filename.lower().endswith(".pdf") or upload.mimetype == "application/pdf":
                from md_to_pdf import extract_latex_text_from_pdf
                markdown_text = extract_latex_text_from_pdf(raw_bytes)
            else:
                markdown_text = raw_bytes.decode("utf-8", errors="replace")
        elif "markdown_text" in request.form and request.form["markdown_text"].strip():
            markdown_text = request.form["markdown_text"]
        else:
            raise ValueError("Please provide a PDF or Markdown file (.pdf, .md, .txt, .tex) or paste Markdown content.")

        preset = request.form.get("preset", "exam").lower()
        paper_size = request.form.get("paper_size", "A4").upper()
        render_math = request.form.get("render_math", "true").lower() == "true"
        font_family = request.form.get("font_family", "sans").lower()

        from md_to_pdf import render_markdown_to_pdf
        pdf_bytes = render_markdown_to_pdf(
            markdown_text=markdown_text,
            preset=preset,
            paper_size=paper_size,
            render_math=render_math,
            font_family=font_family,
        )

        dl_name = output_name(filename, "rendered") if filename != "document.md" else "rendered-document.pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=dl_name,
            mimetype="application/pdf",
        )
    except Exception as error:
        logger.exception("Markdown to PDF conversion failed: %s", error)
        return jsonify({"error": str(error)}), 400


@app.post("/api/extract-markdown-from-pdf")
def extract_markdown_from_pdf_api():
    try:
        upload = request.files.get("file")
        if not upload or not upload.filename:
            raise ValueError("No PDF file provided.")
        raw_bytes = upload.read()
        from md_to_pdf import extract_latex_text_from_pdf
        text = extract_latex_text_from_pdf(raw_bytes)
        import fitz
        doc = fitz.open(stream=raw_bytes, filetype="pdf")
        return jsonify({
            "text": text,
            "page_count": len(doc),
            "filename": upload.filename,
        })
    except Exception as error:
        logger.exception("PDF text extraction failed: %s", error)
        return jsonify({"error": str(error)}), 400


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
