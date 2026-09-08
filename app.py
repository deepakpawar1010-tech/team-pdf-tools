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


@app.get("/<tool>")
def tool_page(tool: str):
    if tool in {"merge", "split", "compress", "pdf-to-word", "word-to-pdf"}:
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
            "small": (96, 50),
            "balanced": (144, 68),
            "best": (200, 80),
            "ultra": (280, 90),
        }
        dpi_target, quality = settings.get(preset, settings["balanced"])

        # Stream upload to disk to avoid blowing up memory on large files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_tmp:
            upload.save(in_tmp.name)
            input_path = in_tmp.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
            output_path = out_tmp.name

        doc = fitz.open(input_path)
        try:
            doc.rewrite_images(dpi_target=dpi_target, quality=quality)
        except Exception:
            pass

        doc.save(output_path, garbage=4, deflate=True, clean=True)
        doc.close()
        del doc
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

        doc = fitz.open(input_path)
        pdf_bytes = doc.convert_to_pdf()
        doc.close()
        del doc

        with open(output_path, "wb") as f:
            f.write(pdf_bytes)

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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
