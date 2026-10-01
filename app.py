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
import time
import tempfile
import threading
import subprocess
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from PIL import Image
import pymupdf
fitz = pymupdf
from flask import Flask, jsonify, render_template, request, send_file, after_this_request
from pypdf import PdfReader, PdfWriter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
word_lock = threading.Lock()
compress_jobs = {}
compress_jobs_lock = threading.Lock()

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


@app.get("/<tool>")
def tool_page(tool: str):
    if tool in {"merge", "split", "compress", "pdf-to-word", "word-to-pdf", "worksheet-splitter", "oly-ete-splitter"}:
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


COMPRESSION_PRESETS = {
    "balanced": {"dpi": 110, "quality": 66, "max_mb": 15.0},
    "high": {"dpi": 160, "quality": 78, "max_mb": 30.0},
    "ultra": {"dpi": 220, "quality": 86, "max_mb": 45.0},
}


def rasterize_pdf_to_limit(
    input_path: str,
    output_path: str,
    max_size_bytes: int,
    start_dpi: int,
    start_quality: int,
) -> int:
    """Last-resort size reduction for image-heavy PDFs; returns final bytes or 0."""
    dpi = start_dpi
    quality = start_quality
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            temp_path = tmp.name

        for _ in range(7):
            source = fitz.open(input_path)
            raster_doc = fitz.open()
            try:
                raster_doc.set_metadata(source.metadata)
                scale = dpi / 72.0
                for page in source:
                    rect = page.rect
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(scale, scale),
                        colorspace=fitz.csRGB,
                        alpha=False,
                        annots=True,
                    )
                    jpeg = pix.tobytes("jpeg", jpg_quality=quality)
                    out_page = raster_doc.new_page(width=rect.width, height=rect.height)
                    out_page.insert_image(out_page.rect, stream=jpeg)
                raster_doc.save(temp_path, garbage=3, deflate=False)
            finally:
                raster_doc.close()
                source.close()

            size = os.path.getsize(temp_path)
            if size <= max_size_bytes:
                shutil.move(temp_path, output_path)
                temp_path = None
                return size

            if dpi <= 40:
                break
            ratio = (max_size_bytes / size) ** 0.5
            dpi = max(40, int(dpi * min(0.84, max(0.55, ratio * 0.92))))
            quality = max(42, quality - 7)
        return 0
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def compress_single_pdf_optimized(
    input_path: str,
    output_path: str,
    preset: str = "balanced",
    max_ceiling_mb: float = 45.0,
) -> dict:
    """Intelligently compresses a PDF using fast xref metadata filtering, multithreaded
    image optimization, and PyMuPDF stream compaction.
    Guarantees that selectable text, vector artwork, annotations, and page layout are 100% preserved.
    """
    t_start = time.perf_counter()
    input_size = os.path.getsize(input_path)
    input_mb = input_size / (1024 * 1024)

    preset = preset if preset in COMPRESSION_PRESETS else "balanced"
    cfg = COMPRESSION_PRESETS[preset]
    max_ceiling_mb = min(max_ceiling_mb, cfg["max_mb"])
    dpi_target = cfg["dpi"]
    quality = cfg["quality"]

    t_open_start = time.perf_counter()
    doc = fitz.open(input_path)
    total_pages = len(doc)
    t_open = time.perf_counter() - t_open_start

    # Fast path: Empty document
    if total_pages == 0:
        doc.save(output_path, garbage=0, deflate=False)
        doc.close()
        return {
            "input_mb": input_mb,
            "output_mb": input_mb,
            "saved_pct": 0.0,
            "total_time": time.perf_counter() - t_start,
            "images_found": 0,
            "images_opt": 0,
            "attempts": 1,
            "open_time": t_open,
            "analysis_time": 0.0,
            "opt_time": 0.0,
            "save_time": 0.0,
        }

    t_analysis_start = time.perf_counter()
    # 1. Discover unique image xrefs and their displaying page dimensions
    image_xrefs = {}
    page_dims = {}
    for pno in range(total_pages):
        page = doc[pno]
        rect = page.rect
        page_dims[pno] = (rect.width, rect.height)
        for img_info in page.get_images():
            xref = img_info[0]
            if xref not in image_xrefs:
                image_xrefs[xref] = pno

    total_images = len(image_xrefs)

    # 2. Fast candidate filter via xref Length key (skips thousands of tiny glyphs/icons)
    min_stream_len = 35 * 1024 if preset != "ultra" else 45 * 1024
    candidates = []
    for xref, pno in image_xrefs.items():
        try:
            t, val = doc.xref_get_key(xref, "Length")
            if t == "int" and int(val) >= min_stream_len:
                # Replacing a base image without its soft mask can turn transparent
                # content black in PDF viewers. Leave masked images untouched.
                smask_type, smask_value = doc.xref_get_key(xref, "SMask")
                mask_type, mask_value = doc.xref_get_key(xref, "Mask")
                if (smask_type == "xref" and smask_value != "0 0 R") or (
                    mask_type not in ("null", "none") and mask_value not in ("null", "[]")
                ):
                    continue
                candidates.append((xref, pno))
        except Exception:
            pass

    t_analysis = time.perf_counter() - t_analysis_start

    # If no candidate images, perform fast stream compaction with garbage=2, deflate=True
    if not candidates:
        t_save_start = time.perf_counter()
        doc.save(output_path, garbage=2, deflate=True)
        doc.close()
        t_save = time.perf_counter() - t_save_start
        out_size = os.path.getsize(output_path)
        if out_size >= input_size:
            shutil.copyfile(input_path, output_path)
            out_size = input_size
        out_mb = out_size / (1024 * 1024)
        if out_mb > max_ceiling_mb:
            raster_size = rasterize_pdf_to_limit(
                input_path, output_path, int(max_ceiling_mb * 1024 * 1024), dpi_target, quality
            )
            if not raster_size:
                raise ValueError(
                    f"This PDF could not be compressed below {max_ceiling_mb:g} MB with the selected preset. "
                    "Choose a lower size preset or a smaller source PDF."
                )
            out_size = raster_size
            out_mb = out_size / (1024 * 1024)
        pct_saved = (1.0 - (out_size / input_size)) * 100
        return {
            "input_mb": input_mb,
            "output_mb": out_mb,
            "saved_pct": pct_saved,
            "total_time": time.perf_counter() - t_start,
            "images_found": total_images,
            "images_opt": 0,
            "attempts": 1,
            "open_time": t_open,
            "analysis_time": t_analysis,
            "opt_time": 0.0,
            "save_time": t_save,
        }

    # 3. Extract candidate image streams
    t_opt_start = time.perf_counter()
    extracted = []
    for xref, pno in candidates:
        try:
            b = doc.extract_image(xref)
            if b and len(b.get("image", b"")) >= min_stream_len:
                pw, ph = page_dims.get(pno, (595.0, 842.0))
                extracted.append((xref, pno, b, pw, ph))
        except Exception:
            pass

    # 4. Multithreaded Pillow optimization
    def process_candidate_image(item):
        xref, pno, b, pw, ph = item
        img_bytes = b["image"]
        orig_len = len(img_bytes)
        orig_w = b.get("width", 0)
        orig_h = b.get("height", 0)
        ext = b.get("ext", "").lower()

        max_w = max(1, int(round((pw / 72.0) * dpi_target)))
        max_h = max(1, int(round((ph / 72.0) * dpi_target)))
        scale = min(1.0, max_w / max(1, orig_w), max_h / max(1, orig_h))

        # If already fits target resolution and is JPEG in high/ultra quality, skip re-encoding
        if scale >= 0.96 and ext in ("jpeg", "jpg") and quality >= 82:
            return None

        new_w = max(1, int(round(orig_w * scale)))
        new_h = max(1, int(round(orig_h * scale)))

        try:
            cs = b.get("colorspace", 3)
            cs_name = b.get("cs-name", "")
            if cs == 4 or cs_name == "DeviceCMYK":
                pix = fitz.Pixmap(doc, xref)
                if pix.colorspace != fitz.csRGB:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            else:
                pil_img = Image.open(io.BytesIO(img_bytes))
                if pil_img.mode != "RGB":
                    pil_img = pil_img.convert("RGB")

            if scale < 0.95:
                pil_img = pil_img.resize((new_w, new_h), Image.Resampling.BILINEAR)

            out_buf = io.BytesIO()
            pil_img.save(out_buf, format="JPEG", quality=quality)

            new_bytes = out_buf.getvalue()
            if len(new_bytes) < orig_len * 0.92:
                return (xref, pno, new_bytes)
        except Exception:
            pass
        return None

    workers = min(8, os.cpu_count() or 4)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        opt_results = list(pool.map(process_candidate_image, extracted))

    # 5. Replace optimized image streams in PyMuPDF doc
    images_opt = 0
    for res in opt_results:
        if res:
            xref, pno, new_bytes = res
            try:
                doc[pno].replace_image(xref, stream=new_bytes)
                images_opt += 1
            except Exception:
                pass

    t_opt = time.perf_counter() - t_opt_start

    # 6. Save optimized document using garbage=2, deflate=True (fast and clean)
    t_save_start = time.perf_counter()
    doc.save(output_path, garbage=2, deflate=True)
    t_save = time.perf_counter() - t_save_start
    doc.close()

    out_size = os.path.getsize(output_path)
    # Never produce an output larger than original
    if out_size >= input_size:
        shutil.copyfile(input_path, output_path)
        out_size = input_size
    out_mb = out_size / (1024 * 1024)
    if out_mb > max_ceiling_mb:
        raster_size = rasterize_pdf_to_limit(
            input_path, output_path, int(max_ceiling_mb * 1024 * 1024), dpi_target, quality
        )
        if not raster_size:
            raise ValueError(
                f"This PDF could not be compressed below {max_ceiling_mb:g} MB with the selected preset. "
                "Choose a lower size preset or a smaller source PDF."
            )
        out_size = raster_size
        out_mb = out_size / (1024 * 1024)
    total_time = time.perf_counter() - t_start
    pct_saved = (1.0 - (out_size / input_size)) * 100

    return {
        "input_mb": input_mb,
        "output_mb": out_mb,
        "saved_pct": pct_saved,
        "total_time": total_time,
        "images_found": total_images,
        "images_opt": images_opt,
        "attempts": 1,
        "open_time": t_open,
        "analysis_time": t_analysis,
        "opt_time": t_opt,
        "save_time": t_save,
    }


@app.post("/api/split")
def split_pdf():
    t_start = time.perf_counter()
    input_path = None
    output_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        ranges = request.form.get("ranges", "")

        t_open_start = time.perf_counter()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            input_path = in_tmp.name
        upload.save(input_path)
        input_size = os.path.getsize(input_path)

        doc = fitz.open(input_path)
        total_pages = len(doc)
        t_open = time.perf_counter() - t_open_start

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
        t_proc_start = time.perf_counter()

        if combine:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                output_path = out_tmp.name

            out_doc = fitz.open()
            for start, end in selected_ranges:
                out_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
            t_proc = time.perf_counter() - t_proc_start

            t_save_start = time.perf_counter()
            # Fast page copy: no expensive garbage collection, no stream re-deflation
            out_doc.save(output_path, garbage=0, deflate=False)
            t_save = time.perf_counter() - t_save_start
            out_doc.close()
            doc.close()

            out_size = os.path.getsize(output_path)
            total_time = time.perf_counter() - t_start

            logger.info(
                "[SPLIT COMBINED]\n"
                "Input: %.2f MB\n"
                "Pages: %d\n"
                "Ranges: %s\n"
                "Open: %.4fs\n"
                "Processing: %.4fs\n"
                "Save: %.4fs\n"
                "Output: %.2f MB\n"
                "Total: %.4fs",
                input_size / (1024 * 1024),
                total_pages,
                ranges,
                t_open,
                t_proc,
                t_save,
                out_size / (1024 * 1024),
                total_time,
            )

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

        # Unmerged: Extract individual PDF files directly in-memory into a ZIP archive
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
            output_path = zip_tmp.name

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for number, (start, end) in enumerate(selected_ranges, start=1):
                sub_doc = fitz.open()
                sub_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
                pdf_bytes = sub_doc.tobytes(garbage=0, deflate=False)
                sub_doc.close()
                suffix = f"page-{start}" if start == end else f"pages-{start}-{end}"
                file_entry_name = output_name(upload.filename, f"{number:02d}_{suffix}")
                zf.writestr(file_entry_name, pdf_bytes)

        t_proc = time.perf_counter() - t_proc_start
        doc.close()

        out_size = os.path.getsize(output_path)
        total_time = time.perf_counter() - t_start

        logger.info(
            "[SPLIT ZIP]\n"
            "Input: %.2f MB\n"
            "Pages: %d\n"
            "Files Extracted: %d\n"
            "Processing & Zip: %.4fs\n"
            "Output: %.2f MB\n"
            "Total: %.4fs",
            input_size / (1024 * 1024),
            total_pages,
            len(selected_ranges),
            t_proc,
            out_size / (1024 * 1024),
            total_time,
        )

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
    t_start = time.perf_counter()
    input_paths = []
    output_path = None
    try:
        uploads = request.files.getlist("files")
        if len(uploads) < 2:
            raise ValueError("Choose at least two PDF files to merge.")

        total_input_bytes = 0
        out_doc = fitz.open()

        t_open_start = time.perf_counter()
        for upload in uploads:
            valid_pdf(upload)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                tmp_path = in_tmp.name
            upload.save(tmp_path)
            input_paths.append(tmp_path)
            total_input_bytes += os.path.getsize(tmp_path)

            sub_doc = fitz.open(tmp_path)
            out_doc.insert_pdf(sub_doc)
            sub_doc.close()
        t_open_and_insert = time.perf_counter() - t_open_start
        total_pages = len(out_doc)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
            output_path = out_tmp.name

        t_save_start = time.perf_counter()
        # Fast page-copy save: streams are already compressed in valid PDFs
        out_doc.save(output_path, garbage=0, deflate=False)
        t_save = time.perf_counter() - t_save_start
        out_doc.close()

        output_size = os.path.getsize(output_path)
        total_time = time.perf_counter() - t_start

        logger.info(
            "[MERGE]\n"
            "Input: %.2f MB\n"
            "Files: %d\n"
            "Pages: %d\n"
            "Open & Insert: %.4fs\n"
            "Save: %.4fs\n"
            "Output: %.2f MB\n"
            "Total: %.4fs",
            total_input_bytes / (1024 * 1024),
            len(uploads),
            total_pages,
            t_open_and_insert,
            t_save,
            output_size / (1024 * 1024),
            total_time,
        )

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


@app.post("/api/compress")
def compress_pdf():
    input_paths = []
    job_id = uuid.uuid4().hex
    try:
        uploads = request.files.getlist("files") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads or not any(getattr(upload, "filename", None) for upload in uploads):
            raise ValueError("Please upload at least one PDF file to compress.")
        for upload in uploads:
            valid_pdf(upload)
        preset = request.form.get("quality", "balanced").strip().lower()

        with compress_jobs_lock:
            if any(job["status"] in ("queued", "processing") for job in compress_jobs.values()):
                return jsonify({"error": "Another compression job is running. Please wait for it to finish."}), 429
            compress_jobs[job_id] = {"status": "queued", "error": None, "output_path": None,
                                     "input_paths": [], "download_name": None, "mimetype": None}

        file_entries = []
        for upload in uploads:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                input_path = in_tmp.name
            input_paths.append(input_path)
            upload.save(input_path)
            file_entries.append((input_path, upload.filename))

        with compress_jobs_lock:
            compress_jobs[job_id]["input_paths"] = input_paths[:]
        threading.Thread(target=run_compression_job, args=(job_id, file_entries, preset),
                         daemon=True, name=f"compress-{job_id[:8]}").start()
        return jsonify({"job_id": job_id, "status": "queued"}), 202
    except Exception as error:
        logger.exception("Could not start compression job: %s", error)
        for path in input_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass
        with compress_jobs_lock:
            if job_id in compress_jobs:
                compress_jobs[job_id].update(status="failed", error=str(error))
        return jsonify({"error": str(error)}), 400


def run_compression_job(job_id, file_entries, preset):
    output_path = None
    temporary_paths = []
    started = time.perf_counter()
    try:
        with compress_jobs_lock:
            compress_jobs[job_id]["status"] = "processing"
        max_mb = COMPRESSION_PRESETS.get(preset, COMPRESSION_PRESETS["balanced"])["max_mb"]
        if len(file_entries) == 1:
            input_path, filename = file_entries[0]
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                output_path = out_tmp.name
            stats = compress_single_pdf_optimized(input_path, output_path, preset=preset, max_ceiling_mb=max_mb)
            download_name = output_name(filename, "compressed")
            mimetype = "application/pdf"
            logger.info("[COMPRESS JOB] Input %.2f MB, output %.2f MB, preset %s, elapsed %.1fs",
                        stats["input_mb"], stats["output_mb"], preset, time.perf_counter() - started)
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
                output_path = zip_tmp.name
            seen_names = {}
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for input_path, filename in file_entries:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as comp_tmp:
                        comp_path = comp_tmp.name
                    temporary_paths.append(comp_path)
                    compress_single_pdf_optimized(input_path, comp_path, preset=preset, max_ceiling_mb=max_mb)
                    entry_name = output_name(filename, "compressed")
                    if entry_name in seen_names:
                        seen_names[entry_name] += 1
                        entry_name = f"{Path(entry_name).stem}_{seen_names[entry_name]}.pdf"
                    else:
                        seen_names[entry_name] = 1
                    zf.write(comp_path, arcname=entry_name)
                    os.remove(comp_path)
                    temporary_paths.remove(comp_path)
            download_name = "compressed-pdfs-bundle.zip"
            mimetype = "application/zip"
            logger.info("[COMPRESS BATCH JOB] %d PDFs finished in %.1fs", len(file_entries), time.perf_counter() - started)

        with compress_jobs_lock:
            compress_jobs[job_id].update(status="complete", output_path=output_path,
                                          download_name=download_name, mimetype=mimetype)
        output_path = None
    except Exception as error:
        logger.exception("Compression job %s failed: %s", job_id, error)
        with compress_jobs_lock:
            if job_id in compress_jobs:
                compress_jobs[job_id].update(status="failed", error=str(error))
    finally:
        for path in [entry[0] for entry in file_entries] + temporary_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass
        if output_path:
            try:
                os.remove(output_path)
            except OSError:
                pass


@app.get("/api/compress/status/<job_id>")
def compress_job_status(job_id):
    with compress_jobs_lock:
        job = compress_jobs.get(job_id)
        if not job:
            return jsonify({"error": "Compression job not found. Please start it again."}), 404
        return jsonify({"status": job["status"], "error": job["error"]})


@app.get("/api/compress/download/<job_id>")
def download_compressed_job(job_id):
    with compress_jobs_lock:
        job = compress_jobs.get(job_id)
        if not job or job["status"] != "complete" or not job["output_path"]:
            return jsonify({"error": "Compressed file is not ready for download."}), 404
        output_path = job["output_path"]
        download_name = job["download_name"]
        mimetype = job["mimetype"]
    response = send_file(output_path, as_attachment=True, download_name=download_name, mimetype=mimetype)

    def cleanup_downloaded_job():
        with compress_jobs_lock:
            finished = compress_jobs.pop(job_id, None)
        if finished:
            for path in finished["input_paths"] + [output_path]:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except OSError:
                    pass

    response.call_on_close(cleanup_downloaded_job)
    return response

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



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
