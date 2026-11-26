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
import math
import tempfile
import threading
import subprocess
from pathlib import Path
from functools import wraps

from PIL import Image
import pymupdf
fitz = pymupdf
from flask import Flask, jsonify, render_template, request, send_file, after_this_request
from pypdf import PdfReader, PdfWriter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
word_lock = threading.Lock()
heavy_process_lock = threading.Lock()
MAX_RENDER_PIXELS = 8_000_000
MAX_DOCUMENT_PAGES = 600

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

if getattr(sys, "frozen", False):
    exe_dir = os.path.dirname(sys.executable)
    candidates = [
        getattr(sys, "_MEIPASS", None),
        os.path.join(exe_dir, "_internal"),
        exe_dir,
        os.path.dirname(os.path.abspath(__file__)),
    ]
    template_folder = None
    static_folder = None
    for c in candidates:
        if c and not template_folder and os.path.isdir(os.path.join(c, "templates")):
            template_folder = os.path.join(c, "templates")
        if c and not static_folder and os.path.isdir(os.path.join(c, "static")):
            static_folder = os.path.join(c, "static")
    if not template_folder:
        template_folder = os.path.join(exe_dir, "_internal", "templates")
    if not static_folder:
        static_folder = os.path.join(exe_dir, "_internal", "static")
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

# Configure a 1.5 GiB upload limit for large PDF processing.
MAX_UPLOAD_SIZE_BYTES = int(1.5 * 1024 * 1024 * 1024)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE_BYTES


def memory_guarded(view):
    """Serialize memory-heavy document jobs on small hosted instances."""
    @wraps(view)
    def guarded(*args, **kwargs):
        if not heavy_process_lock.acquire(blocking=False):
            return jsonify({"error": "The server is processing another document. Please try again shortly."}), 429
        try:
            return view(*args, **kwargs)
        finally:
            heavy_process_lock.release()
    return guarded


@app.errorhandler(413)
def handle_too_large(e):
    return jsonify({"error": "File exceeds the 1.5 GB upload limit (413 Request Entity Too Large)."}), 413


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


def get_downloads_folder() -> str:
    """Resolve Windows user's true Shell Downloads directory with fallback."""
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
            ) as key:
                dl = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                if os.path.isdir(dl):
                    return dl
        except Exception:
            pass
    fallback = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(fallback, exist_ok=True)
    return fallback


def get_unique_filename(directory: str, filename: str) -> str:
    """Prevent overwriting existing files in Downloads by appending (1), (2), etc."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base} ({counter}){ext}")
        counter += 1
    return candidate


def smart_send_file(file_path_or_buffer, download_name: str, mimetype: str = "application/pdf"):
    """
    Delivers file to client via standard HTTP attachment.
    The browser (Google Chrome) downloads the file directly to the user's Downloads folder
    and displays it in Chrome's download manager.
    """
    resp = send_file(
        file_path_or_buffer,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype,
    )
    resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return resp


def send_temporary_file(file_path: str, download_name: str, mimetype: str, cleanup_paths=()):
    """Stream a generated file and remove temporary files after the response closes."""
    response = smart_send_file(file_path, download_name=download_name, mimetype=mimetype)

    def cleanup():
        for path in cleanup_paths:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass

    response.call_on_close(cleanup)
    return response


@app.get("/")
def home():
    return render_template("home.html")


@app.get("/healthz")
@app.get("/api/health")
def healthz():
    return jsonify({"status": "healthy", "service": "team-pdf-tools"}), 200


@app.post("/api/open-folder")
def open_folder():
    data = request.get_json(silent=True) or {}
    path = data.get("path")
    logger.info(f"[OPEN FOLDER] Request received for path: {path}")
    try:
        folder = get_downloads_folder()
        target = folder
        if path and os.path.exists(path):
            target = path if os.path.isdir(path) else os.path.dirname(path)

        logger.info(f"[OPEN FOLDER] Opening directory in File Explorer: {target}")
        if sys.platform == "win32":
            os.startfile(target)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])
        return jsonify({"status": "ok", "opened": target, "type": "folder"})
    except Exception as e:
        logger.exception("Failed to open folder: %s", e)
        return jsonify({"error": str(e)}), 500


@app.post("/api/save-to-downloads")
def save_to_downloads():
    file = request.files.get("file")
    filename = request.form.get("filename") or "document.pdf"
    if not file:
        return jsonify({"error": "No file provided"}), 400
    try:
        downloads_dir = get_downloads_folder()
        dest = get_unique_filename(downloads_dir, filename)
        file.save(dest)
        return jsonify({
            "status": "ok",
            "path": dest,
            "filename": os.path.basename(dest),
            "folder": downloads_dir
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.get("/<tool>")
def tool_page(tool: str):
    if tool in {"merge", "split", "compress", "pdf-to-word", "word-to-pdf", "worksheet-splitter", "oly-ete-splitter"}:
        return render_template("tool.html", tool=tool)
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
@memory_guarded
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
@memory_guarded
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
    "extreme": {"dpi": 60, "quality": 45, "max_mb": 10.0, "label": "Ultra Low — up to 10 MB"},
    "ultra-low": {"dpi": 60, "quality": 45, "max_mb": 10.0, "label": "Ultra Low — up to 10 MB"},
    "balanced": {"dpi": 72, "quality": 55, "max_mb": 15.0, "label": "Balanced — up to 15 MB"},
    "high": {"dpi": 100, "quality": 68, "max_mb": 30.0, "label": "High Clarity — up to 30 MB"},
    "best": {"dpi": 100, "quality": 68, "max_mb": 30.0, "label": "High Clarity — up to 30 MB"},
    "ultra": {"dpi": 130, "quality": 78, "max_mb": 45.0, "label": "Ultra Clarity — up to 45 MB"},
    "small": {"dpi": 60, "quality": 45, "max_mb": 10.0, "label": "Ultra Low — up to 10 MB"},
}


def compress_single_pdf_optimized(
    input_path: str,
    output_path: str,
    preset: str = "balanced",
    max_ceiling_mb: float = None,
) -> dict:
    """Optimize image streams, then use a searchable raster fallback if needed.

    Raster fallback can alter fine layout details and may reduce editability of
    non-text objects; it is used only when needed to meet the selected size limit.
    """
    t_start = time.perf_counter()
    input_size = os.path.getsize(input_path)
    input_mb = input_size / (1024 * 1024)

    cfg = COMPRESSION_PRESETS.get(preset, COMPRESSION_PRESETS["balanced"])
    preset_limit = float(cfg.get("max_mb", 15.0))
    if max_ceiling_mb is not None:
        target_max_mb = min(preset_limit, float(max_ceiling_mb))
    else:
        target_max_mb = preset_limit

    target_bytes = int(target_max_mb * 1024 * 1024)
    must_shrink = input_size > target_bytes

    t_open_start = time.perf_counter()
    doc = fitz.open(input_path)
    total_pages = len(doc)
    t_open = time.perf_counter() - t_open_start

    if total_pages > MAX_DOCUMENT_PAGES:
        doc.close()
        raise ValueError(f"This PDF exceeds the {MAX_DOCUMENT_PAGES}-page limit for this server.")

    # Fast path: Empty document
    if total_pages == 0:
        doc.save(output_path, garbage=0, deflate=False)
        doc.close()
        return {
            "input_mb": input_mb,
            "output_mb": input_mb,
            "target_mb": target_max_mb,
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

    # Fast deduplication of duplicate CMYK ICC profiles (saves up to 35+ MB on merged/multi-page PDFs)
    icc_xrefs = []
    for xref in range(1, doc.xref_length()):
        try:
            t, val = doc.xref_get_key(xref, "N")
            t2, alt = doc.xref_get_key(xref, "Alternate")
            if t == "int" and int(val) == 4 and alt == "/DeviceCMYK":
                icc_xrefs.append(xref)
        except Exception:
            pass
    if len(icc_xrefs) > 1:
        primary_icc = icc_xrefs[0]
        dup_set = set(icc_xrefs[1:])
        for xref in range(1, doc.xref_length()):
            try:
                obj = doc.xref_object(xref)
                for dup in dup_set:
                    token = f"{dup} 0 R"
                    if token in obj:
                        doc.update_object(xref, obj.replace(token, f"{primary_icc} 0 R"))
            except Exception:
                pass
        for dup in dup_set:
            try:
                doc.update_stream(dup, b"")
            except Exception:
                pass

    # 1. Discover unique image xrefs and their displaying page dimensions
    image_xrefs = {}
    page_dims = {}
    for pno in range(total_pages):
        page = doc[pno]
        rect = page.rect
        page_dims[pno] = (rect.width, rect.height)
        for img_info in page.get_images():
            xref = img_info[0]
            smask = img_info[1]
            if xref not in image_xrefs:
                image_xrefs[xref] = (pno, smask)

    total_images = len(image_xrefs)

    # 2. Candidate filter via xref Length key
    # When file exceeds target ceiling, inspect images down to 4 KB so multi-page PDFs are fully compressed
    min_stream_len = 4096 if must_shrink else (12288 if preset != "ultra" else 20480)
    candidates = []
    for xref, (pno, smask) in image_xrefs.items():
        try:
            t, val = doc.xref_get_key(xref, "Length")
            if t == "int" and int(val) >= min_stream_len:
                candidates.append((xref, pno, smask))
        except Exception:
            pass

    t_analysis = time.perf_counter() - t_analysis_start

    # Start at the selected preset's quality. The earlier file-size ratio
    # calculation over-downsampled large scans even when the ceiling had room.
    base_dpi = cfg["dpi"]
    base_q = cfg["quality"]
    dpi_target = base_dpi
    quality = base_q
    dim_limit_scale = 1.0

    t_opt_start = time.perf_counter()

    # 4. Optimize one decoded image at a time. Keeping all decoded images in a
    # list and processing eight concurrently can exceed Render Free's RAM.
    def optimize_images_in_doc(current_doc, cur_dpi, cur_q, cur_cand, force_scale=1.0, cur_dim_limit=1.0):
        if not cur_cand:
            return 0

        def optimize_one(xref, pno, smask):
            try:
                b = current_doc.extract_image(xref)
                if not b or len(b.get("image", b"")) < min_stream_len:
                    return False
                orig_w = int(b.get("width", 0))
                orig_h = int(b.get("height", 0))
                if orig_w < 1 or orig_h < 1 or orig_w * orig_h > MAX_RENDER_PIXELS:
                    return False
                pw, ph = page_dims.get(pno, (595.0, 842.0))
                cs = b.get("colorspace", 3)
                cs_name = b.get("cs-name", "")
                if cs == 4 or cs_name == "DeviceCMYK":
                    return False
                pix = fitz.Pixmap(current_doc, xref)
                if pix.colorspace.name == "DeviceCMYK":
                    return False
                if pix.colorspace != fitz.csRGB:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pix_data = ("RGB", pix.width, pix.height, pix.samples)

                smask_data = None
                if smask > 0:
                    try:
                        spix = fitz.Pixmap(current_doc, smask)
                        smask_data = (spix.width, spix.height, spix.samples)
                    except Exception:
                        pass
                orig_len = len(b.get("image", b""))
                ext = b.get("ext", "").lower()

                max_w = max(1, int(round((pw / 72.0) * cur_dpi * force_scale)))
                max_h = max(1, int(round((ph / 72.0) * cur_dpi * force_scale)))
                scale = min(1.0, max_w / max(1, orig_w), max_h / max(1, orig_h)) * force_scale
                if must_shrink and cur_dim_limit < 1.0:
                    scale = min(scale, cur_dim_limit * force_scale)
                scale = min(1.0, max(0.05, scale))

                if not must_shrink and force_scale >= 0.99:
                    if scale >= 0.96 and ext in ("jpeg", "jpg") and cur_q >= 82 and smask == 0:
                        return False

                new_w = max(30, int(round(orig_w * scale)))
                new_h = max(30, int(round(orig_h * scale)))
                mode, w, h, samples = pix_data
                pil_img = Image.frombytes(mode, [w, h], samples)
                del samples, pix_data
                if scale < 0.98 or force_scale < 0.99:
                    pil_img = pil_img.resize((new_w, new_h), Image.Resampling.BILINEAR)

                out_buf = io.BytesIO()
                pil_img.save(out_buf, format="JPEG", quality=cur_q, optimize=True)
                new_bytes = out_buf.getvalue()

                new_smask_bytes = None
                if smask_data and (scale < 0.98 or force_scale < 0.99):
                    sw, sh, ssamples = smask_data
                    pil_mask = Image.frombytes("L", [sw, sh], ssamples)
                    del ssamples, smask_data
                    pil_mask = pil_mask.resize((new_w, new_h), Image.Resampling.BILINEAR)
                    sbuf = io.BytesIO()
                    pil_mask.save(sbuf, format="PNG")
                    new_smask_bytes = sbuf.getvalue()

                should_replace = len(new_bytes) < (orig_len if must_shrink or force_scale < 0.99 else orig_len * 0.98)
                if not should_replace:
                    return False
                current_doc.update_stream(xref, new_bytes)
                current_doc.xref_set_key(xref, "Width", str(new_w))
                current_doc.xref_set_key(xref, "Height", str(new_h))
                if smask > 0 and new_smask_bytes:
                    current_doc.update_stream(smask, new_smask_bytes)
                    current_doc.xref_set_key(smask, "Width", str(new_w))
                    current_doc.xref_set_key(smask, "Height", str(new_h))
                return True
            except Exception:
                return False

        replaced = 0
        for xref, pno, smask in cur_cand:
            if optimize_one(xref, pno, smask):
                replaced += 1
        return replaced

    t_opt = 0.0
    images_opt = 0
    if candidates:
        images_opt = optimize_images_in_doc(
            doc,
            dpi_target,
            quality,
            candidates,
            force_scale=1.0,
            cur_dim_limit=dim_limit_scale,
        )
        t_opt = time.perf_counter() - t_opt_start

    # 5. Save optimized document using garbage=2, deflate=True, use_objstms=1 (super fast)
    t_save_start = time.perf_counter()
    doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
    t_save = time.perf_counter() - t_save_start
    out_size = os.path.getsize(output_path)
    attempts = 1

    # 6. Multi-pass iterative refinement if output still exceeds chosen ceiling limit
    while out_size > target_bytes and attempts < 4 and candidates:
        prev_size = out_size
        attempts += 1
        excess_ratio = out_size / target_bytes
        dim_scale = min(0.85, (1.0 / math.sqrt(excess_ratio)) * 0.90)
        dim_scale = max(0.15, dim_scale)
        quality = max(20, int(quality * 0.80))
        dpi_target = max(28, int(dpi_target * dim_scale))

        doc.close()
        doc = fitz.open(output_path)
        cur_cands = []
        for pno in range(len(doc)):
            for img_info in doc[pno].get_images():
                cur_cands.append((img_info[0], pno, img_info[1]))

        images_opt += optimize_images_in_doc(
            doc,
            dpi_target,
            quality,
            cur_cands,
            force_scale=1.0,
            cur_dim_limit=1.0,
        )
        t_save_extra_start = time.perf_counter()
        temp_iter = output_path + f".iter{attempts}.pdf"
        doc.save(temp_iter, garbage=2, deflate=True, use_objstms=1)
        doc.close()
        shutil.move(temp_iter, output_path)
        doc = fitz.open(output_path)
        t_save += (time.perf_counter() - t_save_extra_start)
        out_size = os.path.getsize(output_path)
        # If image optimization yielded negligible further reduction (< 4%),
        # remaining bulk is vector/Form data, so immediately jump to hybrid pass!
        if (prev_size - out_size) < (prev_size * 0.04):
            break

    # 7. HYBRID PASS: If output STILL exceeds target ceiling (e.g. heavy vector curves / Form XObjects),
    # automatically consolidate pages into crisp compressed backgrounds with searchable text on top!
    if out_size > target_bytes:
        attempts += 1
        doc.close()
        src_doc = fitz.open(output_path)
        hybrid_doc = fitz.open()

        if target_max_mb <= 10.0:
            h_dpi = 100
            h_q = 50
        elif target_max_mb <= 15.0:
            h_dpi = 115
            h_q = 55
        elif target_max_mb <= 30.0:
            h_dpi = 135
            h_q = 65
        else:
            h_dpi = 150
            h_q = 72

        page_budget = target_bytes / max(1, len(src_doc))
        budget_dpi = max(35, int(150 * math.sqrt(page_budget / (250 * 1024))))
        h_dpi = min(h_dpi, budget_dpi)

        for pno in range(len(src_doc)):
            page = src_doc[pno]
            rect = page.rect
            page_area = max(1.0, rect.width * rect.height)
            safe_dpi = min(h_dpi, 72 * math.sqrt(MAX_RENDER_PIXELS / page_area))
            pix = page.get_pixmap(dpi=safe_dpi, colorspace=fitz.csRGB, alpha=False)
            pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            pil_img.save(buf, format="JPEG", quality=h_q, optimize=True)
            img_b = buf.getvalue()

            new_page = hybrid_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_b)

            # Preserve searchable / selectable text on top
            blocks = page.get_text("blocks")
            for b in blocks:
                if len(b) >= 5 and b[4].strip():
                    txt = b[4].strip()
                    r = fitz.Rect(b[0], b[1], b[2], b[3])
                    try:
                        new_page.insert_textbox(r, txt, fontsize=10, render_mode=3)
                    except Exception:
                        pass

        num_hybrid_pages = len(src_doc)
        src_doc.close()
        temp_h = output_path + ".hybrid.pdf"
        hybrid_doc.save(temp_h, garbage=2, deflate=True, use_objstms=1)
        hybrid_doc.close()
        shutil.move(temp_h, output_path)
        out_size = os.path.getsize(output_path)
        images_opt += num_hybrid_pages
    else:
        doc.close()

    # Safety: Never produce output larger than input if input was already within ceiling
    if out_size >= input_size and input_size <= target_bytes:
        shutil.copyfile(input_path, output_path)
        out_size = input_size

    if out_size > target_bytes:
        try:
            os.remove(output_path)
        except OSError:
            pass
        raise ValueError(
            f"Could not reduce this PDF below {target_max_mb:g} MB without more aggressive quality loss. "
            "Try the next lower compression preset."
        )

    out_mb = out_size / (1024 * 1024)
    pct_saved = (1.0 - (out_size / input_size)) * 100 if input_size > 0 else 0.0
    total_time = time.perf_counter() - t_start

    return {
        "input_mb": input_mb,
        "output_mb": out_mb,
        "target_mb": target_max_mb,
        "saved_pct": pct_saved,
        "total_time": total_time,
        "images_found": total_images,
        "images_opt": images_opt,
        "attempts": attempts,
        "open_time": t_open,
        "analysis_time": t_analysis,
        "opt_time": t_opt,
        "save_time": t_save,
    }


@app.post("/api/split")
@memory_guarded
def split_pdf():
    t_start = time.perf_counter()
    input_path = None
    output_path = None
    temp_split_paths = []
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
        if total_pages > MAX_DOCUMENT_PAGES:
            doc.close()
            raise ValueError(f"This PDF exceeds the {MAX_DOCUMENT_PAGES}-page limit for this server.")
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

            return smart_send_file(
                output_path,
                download_name=output_name(upload.filename, "selected-pages"),
                mimetype="application/pdf",
            )

        # Unmerged: Extract individual PDF files directly in-memory into a ZIP archive
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
            output_path = zip_tmp.name

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_STORED) as zf:
            for number, (start, end) in enumerate(selected_ranges, start=1):
                sub_doc = fitz.open()
                try:
                    sub_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as sub_tmp:
                        sub_path = sub_tmp.name
                    temp_split_paths.append(sub_path)
                    sub_doc.save(sub_path, garbage=0, deflate=False)
                finally:
                    sub_doc.close()
                suffix = f"page-{start}" if start == end else f"pages-{start}-{end}"
                file_entry_name = output_name(upload.filename, f"{number:02d}_{suffix}")
                zf.write(sub_path, arcname=file_entry_name)
                os.remove(sub_path)
                temp_split_paths.remove(sub_path)

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

        return smart_send_file(
            output_path,
            download_name=output_name(upload.filename, "split-files").replace(".pdf", ".zip"),
            mimetype="application/zip",
        )
    except Exception as error:
        logger.exception("Split PDF failed: %s", error)
        for path in temp_split_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
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
@memory_guarded
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
            if len(out_doc) + len(sub_doc) > MAX_DOCUMENT_PAGES:
                sub_doc.close()
                out_doc.close()
                raise ValueError(f"Merged PDFs cannot exceed {MAX_DOCUMENT_PAGES} pages on this server.")
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

        return smart_send_file(
            output_path,
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
@memory_guarded
def compress_pdf():
    input_paths = []
    output_path = None
    temp_compressed_files = []
    try:
        uploads = request.files.getlist("files") or ([request.files.get("file")] if request.files.get("file") else [])
        if not uploads or not any(getattr(u, "filename", None) for u in uploads):
            raise ValueError("Please upload at least one PDF file to compress.")

        for upload in uploads:
            valid_pdf(upload)

        preset = request.form.get("quality", "balanced").strip().lower()

        # Case 1: Single file upload
        if len(uploads) == 1:
            upload = uploads[0]
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                input_path = in_tmp.name
            input_paths.append(input_path)
            upload.save(input_path)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                output_path = out_tmp.name

            stats = compress_single_pdf_optimized(
                input_path,
                output_path,
                preset=preset,
            )

            logger.info(
                "[COMPRESS]\n"
                "Input: %.2f MB\n"
                "Preset: %s\n"
                "Images Found: %d\n"
                "Images Optimized: %d\n"
                "Attempts: %d\n"
                "Open: %.4fs\n"
                "Analysis: %.4fs\n"
                "Image Opt: %.4fs\n"
                "Save: %.4fs\n"
                "Output: %.2f MB (%.1f%% reduction)\n"
                "Total: %.4fs",
                stats["input_mb"],
                preset,
                stats.get("images_found", 0),
                stats.get("images_opt", 0),
                stats.get("attempts", 1),
                stats.get("open_time", 0.0),
                stats.get("analysis_time", 0.0),
                stats.get("opt_time", 0.0),
                stats.get("save_time", 0.0),
                stats["output_mb"],
                stats["saved_pct"],
                stats["total_time"],
            )

            return send_temporary_file(
                output_path,
                download_name=output_name(upload.filename, "compressed"),
                mimetype="application/pdf",
                cleanup_paths=input_paths + [output_path],
            )

        # Case 2: Batch conversion (Multiple PDF files) -> Output as ZIP archive
        t_batch_start = time.perf_counter()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as zip_tmp:
            output_path = zip_tmp.name

        seen_names = {}
        total_batch_in_bytes = 0
        total_batch_out_bytes = 0

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_STORED) as zf:
            for upload in uploads:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
                    in_path = in_tmp.name
                input_paths.append(in_path)
                upload.save(in_path)
                total_batch_in_bytes += os.path.getsize(in_path)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as comp_tmp:
                    comp_path = comp_tmp.name
                temp_compressed_files.append(comp_path)

                compress_single_pdf_optimized(
                    in_path,
                    comp_path,
                    preset=preset,
                )
                comp_size = os.path.getsize(comp_path)
                total_batch_out_bytes += comp_size

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

        t_batch_total = time.perf_counter() - t_batch_start
        zip_size = os.path.getsize(output_path)

        logger.info(
            "[COMPRESS BATCH]\n"
            "Input Files: %d\n"
            "Input Total: %.2f MB\n"
            "Output ZIP: %.2f MB\n"
            "Total Time: %.4fs",
            len(uploads),
            total_batch_in_bytes / (1024 * 1024),
            zip_size / (1024 * 1024),
            t_batch_total,
        )

        return send_temporary_file(
            output_path,
            download_name="compressed-pdfs-bundle.zip",
            mimetype="application/zip",
            cleanup_paths=input_paths + temp_compressed_files + [output_path],
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
@memory_guarded
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

        return smart_send_file(
            output_path,
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
@memory_guarded
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

        return smart_send_file(
            output_path,
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
@memory_guarded
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
@memory_guarded
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

        return smart_send_file(
            zip_buf,
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
    import waitress
    logger.info(f"Starting Waitress production WSGI server on http://127.0.0.1:{port}")
    waitress.serve(
        app,
        host="0.0.0.0",
        port=port,
        threads=8,
        channel_timeout=600,
        max_request_body_size=MAX_UPLOAD_SIZE_BYTES,
    )
