from __future__ import annotations

import io
import logging
import zipfile
from pathlib import Path

import pymupdf as fitz

from .config import WORKSHEET_OUTPUT_DIR
from .cropper import crop_pages_for_worksheet
from .detector import (
    WorksheetCandidate,
    build_worksheet_ranges,
    detect_all_boundaries,
    detect_headings,
    detect_worksheet_candidates,
)

logger = logging.getLogger(__name__)


def _safe_name(name: str) -> str:
    cleaned = "".join(character if character.isalnum() or character in {" ", "-", "_"} else "_" for character in name)
    return cleaned.strip().rstrip(".") or "Untitled"


def _resolve_output_path(output_dir: Path, worksheet_name: str) -> Path:
    primary_path = output_dir / f"{worksheet_name}.pdf"
    if not primary_path.exists():
        return primary_path

    try:
        primary_path.unlink()
        return primary_path
    except PermissionError:
        suffix = 1
        while True:
            candidate_path = output_dir / f"{worksheet_name}_{suffix}.pdf"
            if not candidate_path.exists():
                return candidate_path
            suffix += 1


def get_worksheet_info(pdf_path: Path, stop_at_synopsis: bool = True) -> dict:
    """Detects worksheet candidates and returns preview information."""
    markers = detect_all_boundaries(pdf_path)
    with fitz.open(pdf_path) as source_doc:
        total_pages = source_doc.page_count

    ranges = build_worksheet_ranges(markers, total_pages, stop_at_synopsis=stop_at_synopsis)
    summary = []
    for r in ranges:
        summary.append({
            "name": r.worksheet_name,
            "start_page": r.start_page,
            "end_page": r.end_page,
            "page_count": max(1, r.end_page - r.start_page + 1)
        })
    return {
        "total_pages": total_pages,
        "count": len(summary),
        "worksheets": summary
    }


def split_pdf_to_zip(pdf_path: Path, stop_at_synopsis: bool = True) -> tuple[io.BytesIO, list[dict]]:
    """
    Splits the PDF into individual cropped worksheet PDFs and bundles them
    into an in-memory ZIP archive for instant download.
    """
    markers = detect_all_boundaries(pdf_path)
    has_worksheets = any(m.kind == "worksheet" for m in markers)
    if not has_worksheets:
        return io.BytesIO(), []

    zip_buffer = io.BytesIO()
    worksheets_info = []

    with fitz.open(pdf_path) as source_document:
        worksheet_ranges = build_worksheet_ranges(markers, source_document.page_count, stop_at_synopsis=stop_at_synopsis)

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for worksheet_range in worksheet_ranges:
                start_index = worksheet_range.start_page - 1
                end_index = worksheet_range.end_page - 1
                if end_index < start_index:
                    continue

                ws_doc = fitz.open()
                ws_doc.insert_pdf(source_document, from_page=start_index, to_page=end_index)
                crop_pages_for_worksheet(ws_doc, worksheet_range)

                pdf_bytes = ws_doc.tobytes(garbage=3, deflate=True)
                ws_doc.close()

                filename = f"{worksheet_range.worksheet_name}.pdf"
                zip_file.writestr(filename, pdf_bytes)

                worksheets_info.append({
                    "filename": filename,
                    "start_page": worksheet_range.start_page,
                    "end_page": worksheet_range.end_page
                })

    zip_buffer.seek(0)
    return zip_buffer, worksheets_info


def split_pdf(pdf_path: Path, output_root: Path | None = None, stop_at_synopsis: bool = True) -> dict:
    headings = detect_headings(pdf_path)
    markers = detect_all_boundaries(pdf_path)
    candidates = [
        WorksheetCandidate(m.name, m.page_number, m.trigger_text, m.bbox, m.reason)
        for m in markers if m.kind == "worksheet"
    ]

    target_root = output_root or WORKSHEET_OUTPUT_DIR
    output_dir = target_root / _safe_name(pdf_path.stem)
    output_dir.mkdir(parents=True, exist_ok=True)

    saved_files: list[Path] = []
    if markers:
        with fitz.open(pdf_path) as source_document:
            worksheet_ranges = build_worksheet_ranges(markers, source_document.page_count, stop_at_synopsis=stop_at_synopsis)
            for worksheet_range in worksheet_ranges:
                start_index = worksheet_range.start_page - 1
                end_index = worksheet_range.end_page - 1
                if end_index < start_index:
                    continue

                worksheet_document = fitz.open()
                worksheet_document.insert_pdf(source_document, from_page=start_index, to_page=end_index)
                crop_pages_for_worksheet(worksheet_document, worksheet_range)

                output_path = _resolve_output_path(output_dir, worksheet_range.worksheet_name)
                worksheet_document.save(output_path, garbage=3, deflate=True)
                worksheet_document.close()
                saved_files.append(output_path)

    logger.info("Worksheet splitter processed %s and created %s files", pdf_path.name, len(saved_files))
    return {
        "pdf_name": pdf_path.name,
        "headings": headings,
        "candidates": candidates,
        "output_dir": output_dir,
        "saved_files": saved_files,
    }


def split_pdfs(pdf_paths: list[Path], output_root: Path | None = None, stop_at_synopsis: bool = True) -> list[dict]:
    results: list[dict] = []
    for pdf_path in pdf_paths:
        try:
            results.append(split_pdf(pdf_path, output_root=output_root, stop_at_synopsis=stop_at_synopsis))
        except Exception as error:
            logger.exception("Worksheet splitting failed for %s", pdf_path)
            results.append(
                {
                    "pdf_name": pdf_path.name,
                    "headings": [],
                    "candidates": [],
                    "output_dir": output_root or WORKSHEET_OUTPUT_DIR,
                    "saved_files": [],
                    "error": str(error),
                }
            )
    return results

