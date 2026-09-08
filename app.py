"""Local Flask server for Team PDF Tools."""
from __future__ import annotations

import io
import os
import re
import zipfile
import base64
import gc
import tempfile
from pathlib import Path

import pymupdf as fitz
from flask import Flask, jsonify, render_template, request, send_file, after_this_request
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
MAX_UPLOAD_BYTES = 1200 * 1024 * 1024  # 1.2 GB
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES


def valid_pdf(upload):
    if not upload or not upload.filename.lower().endswith(".pdf"):
        raise ValueError("Please upload a PDF file.")


def output_name(filename: str, suffix: str) -> str:
    stem = Path(filename).stem
    safe_stem = re.sub(r"[^\w.-]+", "-", stem).strip("-") or "document"
    return f"{safe_stem}-{suffix}.pdf"


@app.get("/")
def home():
    return render_template("home.html")


COMING_SOON = {
    "pdf-to-word": {
        "title": "PDF to Word",
        "description": "Convert PDF documents into editable Microsoft Word (.docx) files while preserving original typography and formatting.",
        "badge": "DOCX Conversion",
        "icon": "word"
    },
    "word-to-pdf": {
        "title": "Word to PDF",
        "description": "Convert Microsoft Word documents directly into high-fidelity, standardized PDF documents.",
        "badge": "DOCX to PDF",
        "icon": "word-to-pdf"
    },
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


@app.get("/<tool>")
def tool_page(tool: str):
    if tool in {"merge", "split", "compress"}:
        return render_template("tool.html", tool=tool)
    if tool in COMING_SOON:
        return render_template("coming_soon.html", tool=tool, info=COMING_SOON[tool])
    return "Page not found", 404


@app.post("/api/info")
def info():
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        reader = PdfReader(upload.stream)
        return jsonify({"pages": len(reader.pages), "name": upload.filename})
    except Exception as error:
        return jsonify({"error": str(error)}), 400


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
            data = upload.read()
            document = fitz.open(stream=data, filetype="pdf")
            page_count = document.page_count
            if mode == "merge":
                page_numbers = [1] if page_count else []
            elif requested_pages is None:
                page_numbers = list(range(1, page_count + 1))
            else:
                page_numbers = sorted(p for p in requested_pages if p <= page_count)

            for page_number in page_numbers:
                page = document.load_page(page_number - 1)
                rect = page.rect
                longest = max(rect.width, rect.height) or 1
                scale = min(1.0, 240 / longest)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                jpeg = pixmap.tobytes("jpeg", jpg_quality=65)
                output.append({
                    "fileIndex": file_index,
                    "page": page_number,
                    "pages": page_count,
                    "name": upload.filename,
                    "data": "data:image/jpeg;base64," + base64.b64encode(jpeg).decode("ascii"),
                })
            document.close()

        return jsonify({"previews": output})
    except Exception as error:
        return jsonify({"error": str(error)}), 400


@app.post("/api/split")
def split_pdf():
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        ranges = request.form.get("ranges", "")
        reader = PdfReader(upload.stream)
        selected_ranges = []
        for piece in ranges.split(","):
            start, end = (int(value.strip()) for value in piece.split("-", 1))
            if start < 1 or end < start or end > len(reader.pages):
                raise ValueError(f"Use page ranges between 1 and {len(reader.pages)}.")
            selected_ranges.append((start, end))
        if request.form.get("combine", "true") == "true":
            writer = PdfWriter()
            for start, end in selected_ranges:
                for index in range(start - 1, end):
                    writer.add_page(reader.pages[index])
            result = io.BytesIO()
            writer.write(result)
            result.seek(0)
            return send_file(result, as_attachment=True, download_name=output_name(upload.filename, "selected-pages"), mimetype="application/pdf")
        archive = io.BytesIO()
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as output:
            for number, (start, end) in enumerate(selected_ranges, start=1):
                writer = PdfWriter()
                for index in range(start - 1, end):
                    writer.add_page(reader.pages[index])
                pdf = io.BytesIO()
                writer.write(pdf)
                output.writestr(output_name(upload.filename, f"pages-{start}-{end}"), pdf.getvalue())
        archive.seek(0)
        return send_file(archive, as_attachment=True, download_name=output_name(upload.filename, "split-files").replace(".pdf", ".zip"), mimetype="application/zip")
    except Exception as error:
        return jsonify({"error": str(error)}), 400


@app.post("/api/merge")
def merge_pdf():
    try:
        uploads = request.files.getlist("files")
        if len(uploads) < 2:
            raise ValueError("Choose at least two PDF files to merge.")
        writer = PdfWriter()
        for upload in uploads:
            valid_pdf(upload)
            reader = PdfReader(upload.stream)
            for page in reader.pages:
                writer.add_page(page)
        result = io.BytesIO()
        writer.write(result)
        result.seek(0)
        return send_file(result, as_attachment=True, download_name="merged.pdf", mimetype="application/pdf")
    except Exception as error:
        return jsonify({"error": str(error)}), 400


@app.post("/api/compress")
def compress_pdf():
    input_path = None
    output_path = None
    try:
        upload = request.files.get("file")
        valid_pdf(upload)
        preset = request.form.get("quality", "balanced")
        settings = {
            "small": (0.7, 50, 900),
            "balanced": (0.85, 65, 1200),
            "best": (1.0, 80, 1600),
            "ultra": (1.2, 90, 2000),
        }
        scale, quality, max_dim = settings.get(preset, settings["balanced"])

        # Stream upload to disk to avoid blowing up memory on large files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            upload.save(in_tmp.name)
            input_path = in_tmp.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
            output_path = out_tmp.name

        source = fitz.open(input_path)
        result = fitz.open()

        for idx, page in enumerate(source):
            rect = page.rect
            longest = max(rect.width, rect.height) or 1
            eff_scale = min(scale, max_dim / longest)

            pixmap = page.get_pixmap(matrix=fitz.Matrix(eff_scale, eff_scale), alpha=False)
            jpeg_bytes = pixmap.tobytes("jpeg", jpg_quality=quality)
            del pixmap

            new_page = result.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(new_page.rect, stream=jpeg_bytes)
            del jpeg_bytes

            if (idx + 1) % 10 == 0:
                gc.collect()

        result.save(output_path, garbage=4, deflate=True)
        result.close()
        source.close()
        del result, source
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
            download_name=output_name(upload.filename, "compressed"),
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
