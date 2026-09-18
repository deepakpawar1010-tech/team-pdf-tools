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


def get_worksheet_info(
    pdf_path: Path,
    stop_at_synopsis: bool = True,
    preset: str = "olympiad",
    include_key: bool = True,
) -> dict:
    """Detects worksheet candidates and returns preview information."""
    markers = detect_all_boundaries(pdf_path, preset=preset)
    with fitz.open(pdf_path) as source_doc:
        total_pages = source_doc.page_count

    ranges = build_worksheet_ranges(
        markers,
        total_pages,
        stop_at_synopsis=stop_at_synopsis,
        preset=preset,
        include_key=include_key,
    )
    summary = []
    for r in ranges:
        extra = getattr(r, "extra_pages", ())
        base_count = max(1, r.end_page - r.start_page + 1)
        total_count = base_count + len(extra)
        summary.append({
            "name": r.worksheet_name,
            "start_page": r.start_page,
            "end_page": r.end_page,
            "extra_pages": list(extra),
            "page_count": total_count
        })
    return {
        "total_pages": total_pages,
        "count": len(summary),
        "worksheets": summary
    }


def split_pdf_to_zip(
    pdf_path: Path,
    stop_at_synopsis: bool = True,
    preset: str = "olympiad",
    include_key: bool = True,
    crop_top: bool = True,
) -> tuple[io.BytesIO, list[dict]]:
    """Splits a single PDF into individual cropped worksheet PDFs inside an in-memory ZIP."""
    return split_pdfs_to_zip(
        [(pdf_path.name, pdf_path)],
        stop_at_synopsis=stop_at_synopsis,
        preset=preset,
        include_key=include_key,
        crop_top=crop_top,
    )


def split_pdfs_to_zip(
    pdf_inputs: list[tuple[str, Path]],
    stop_at_synopsis: bool = True,
    preset: str = "olympiad",
    include_key: bool = True,
    crop_top: bool = True,
) -> tuple[io.BytesIO, list[dict]]:
    """
    Splits one or more PDFs into individual cropped worksheet PDFs and bundles
    them into a single consolidated in-memory ZIP archive for instant download.
    """
    zip_buffer = io.BytesIO()
    all_worksheets_info: list[dict] = []
    multi_file = len(pdf_inputs) > 1

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for orig_name, pdf_path in pdf_inputs:
            doc_stem = _safe_name(Path(orig_name).stem)
            markers = detect_all_boundaries(pdf_path, preset=preset)
            has_worksheets = any(m.kind == "worksheet" for m in markers)
            if not has_worksheets:
                continue

            with fitz.open(pdf_path) as source_document:
                worksheet_ranges = build_worksheet_ranges(
                    markers,
                    source_document.page_count,
                    stop_at_synopsis=stop_at_synopsis,
                    preset=preset,
                    include_key=include_key,
                )

                for worksheet_range in worksheet_ranges:
                    start_index = worksheet_range.start_page - 1
                    end_index = worksheet_range.end_page - 1
                    if end_index < start_index:
                        continue

                    ws_doc = fitz.open()
                    ws_doc.insert_pdf(source_document, from_page=start_index, to_page=end_index)
                    crop_pages_for_worksheet(ws_doc, worksheet_range, preset=preset, crop_top=crop_top)

                    for ep in getattr(worksheet_range, "extra_pages", ()):
                        ep_idx = ep - 1
                        if 0 <= ep_idx < source_document.page_count:
                            ws_doc.insert_pdf(source_document, from_page=ep_idx, to_page=ep_idx)

                    pdf_bytes = ws_doc.tobytes(garbage=3, deflate=True)
                    ws_doc.close()

                    if multi_file:
                        filename = f"{doc_stem}_{worksheet_range.worksheet_name}.pdf"
                    else:
                        filename = f"{worksheet_range.worksheet_name}.pdf"

                    existing_names = set(zip_file.namelist())
                    counter = 1
                    base_filename = filename
                    while filename in existing_names:
                        stem = base_filename[:-4]
                        filename = f"{stem}_{counter}.pdf"
                        counter += 1

                    zip_file.writestr(filename, pdf_bytes)

                    all_worksheets_info.append({
                        "source_file": orig_name,
                        "filename": filename,
                        "start_page": worksheet_range.start_page,
                        "end_page": worksheet_range.end_page,
                        "extra_pages": list(getattr(worksheet_range, "extra_pages", ())),
                    })

    zip_buffer.seek(0)
    return zip_buffer, all_worksheets_info


def split_pdf(
    pdf_path: Path,
    output_root: Path | None = None,
    stop_at_synopsis: bool = True,
    preset: str = "olympiad",
    include_key: bool = True,
    crop_top: bool = True,
) -> dict:
    headings = detect_headings(pdf_path)
    markers = detect_all_boundaries(pdf_path, preset=preset)
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
            worksheet_ranges = build_worksheet_ranges(
                markers,
                source_document.page_count,
                stop_at_synopsis=stop_at_synopsis,
                preset=preset,
                include_key=include_key,
            )
            for worksheet_range in worksheet_ranges:
                start_index = worksheet_range.start_page - 1
                end_index = worksheet_range.end_page - 1
                if end_index < start_index:
                    continue

                worksheet_document = fitz.open()
                worksheet_document.insert_pdf(source_document, from_page=start_index, to_page=end_index)
                crop_pages_for_worksheet(worksheet_document, worksheet_range, preset=preset, crop_top=crop_top)

                for ep in getattr(worksheet_range, "extra_pages", ()):
                    ep_idx = ep - 1
                    if 0 <= ep_idx < source_document.page_count:
                        worksheet_document.insert_pdf(source_document, from_page=ep_idx, to_page=ep_idx)

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


def split_pdfs(
    pdf_paths: list[Path],
    output_root: Path | None = None,
    stop_at_synopsis: bool = True,
    preset: str = "olympiad",
    include_key: bool = True,
    crop_top: bool = True,
) -> list[dict]:
    results: list[dict] = []
    for pdf_path in pdf_paths:
        try:
            results.append(
                split_pdf(
                    pdf_path,
                    output_root=output_root,
                    stop_at_synopsis=stop_at_synopsis,
                    preset=preset,
                    include_key=include_key,
                    crop_top=crop_top,
                )
            )
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

