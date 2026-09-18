from __future__ import annotations

import base64
import logging
import re
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pymupdf as fitz

logger = logging.getLogger(__name__)

WORKSHEET_PATTERN = re.compile(r"\bWORKSHEET\s*-?\s*(\d{1,3})\b", re.IGNORECASE)
SYNOPSIS_PATTERN = re.compile(r"\bSYNOPSIS\s*-?\s*(\d{1,3})?\b", re.IGNORECASE)
CUQ_PATTERN = re.compile(r"^\s*CUQ\s*$", re.IGNORECASE)

# Compact binary template (24x120 pixels) for graphic SYNOPSIS header badges
B64_SYNOPSIS_TEMPLATE = (
    "eJx9z8FtwzAMBVDSDPp7MKIFimoTe7HCpqIBOkJXYdBD1xDQBQTkkkNghXJy6aU8iHgQQH629qe2"
    "ub+3tbVrb5WCItegoYxcUCicJBsUemATIwySVVTowComwwtnZeWd6buTWUkAyemiR8piyfAD+dXm"
    "PMEpDdw02jsNnbSCZsOTiSI0mJSJhmA+OfjywmV69V/fDJhUrstOGFBkc46d4Qrc5Cs5xTmtwCyf"
    "57ocRTRtEYietpM4bZ6Z/JK6BNEn8aCJc+zETjjfPEmJzlzjuekH29z+rTtGWZhZ"
)
_CACHED_SYNOPSIS_TEMPLATE = None


def _get_synopsis_template():
    global _CACHED_SYNOPSIS_TEMPLATE
    if _CACHED_SYNOPSIS_TEMPLATE is None:
        try:
            import numpy as np
            rec_packed = np.frombuffer(zlib.decompress(base64.b64decode(B64_SYNOPSIS_TEMPLATE)), dtype=np.uint8)
            _CACHED_SYNOPSIS_TEMPLATE = np.unpackbits(rec_packed).reshape((24, 120)) * 255
        except Exception as err:
            logger.debug("Could not unpack synopsis template: %s", err)
            _CACHED_SYNOPSIS_TEMPLATE = False
    return _CACHED_SYNOPSIS_TEMPLATE if _CACHED_SYNOPSIS_TEMPLATE is not False else None


@dataclass(frozen=True)
class HeadingMatch:
    kind: str
    number: str
    page_number: int
    bbox: tuple[float, float, float, float]
    text: str

    @property
    def label(self) -> str:
        return f"{self.kind.title()} {self.number}"


@dataclass(frozen=True)
class LineData:
    text: str
    bbox: tuple[float, float, float, float]
    word_items: tuple[tuple[float, float, float, float, str], ...]
    page_number: int


@dataclass(frozen=True)
class WorksheetCandidate:
    worksheet_name: str
    page_number: int
    trigger_text: str
    bbox: tuple[float, float, float, float]
    reason: str


@dataclass(frozen=True)
class BoundaryMarker:
    kind: str  # "worksheet" or "synopsis"
    name: str
    page_number: int
    bbox: tuple[float, float, float, float]
    trigger_text: str
    reason: str


@dataclass(frozen=True)
class WorksheetRange:
    worksheet_name: str
    start_page: int
    end_page: int
    start_bbox: tuple[float, float, float, float]
    end_bbox: tuple[float, float, float, float] | None = None


def _group_words_into_lines(page: fitz.Page) -> Iterable[LineData]:
    raw_words = page.get_text("words", sort=True)
    grouped: dict[tuple[int, int], list[tuple[float, float, float, float, str, int]]] = {}

    for x0, y0, x1, y1, word, block_no, line_no, word_no in raw_words:
        grouped.setdefault((block_no, line_no), []).append((x0, y0, x1, y1, word, word_no))

    page_number = page.number + 1
    ordered_lines = sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1]))

    for _, words in ordered_lines:
        sorted_words = sorted(words, key=lambda item: item[5])
        text_parts = [word for _, _, _, _, word, _ in sorted_words]
        if not text_parts:
            continue

        x0 = min(item[0] for item in sorted_words)
        y0 = min(item[1] for item in sorted_words)
        x1 = max(item[2] for item in sorted_words)
        y1 = max(item[3] for item in sorted_words)

        yield LineData(
            text=" ".join(text_parts).strip(),
            bbox=(x0, y0, x1, y1),
            word_items=tuple((x0, y0, x1, y1, word) for x0, y0, x1, y1, word, _ in sorted_words),
            page_number=page_number,
        )


def _normalize_number(number: str) -> str:
    return number.zfill(2)


def _build_heading_match(kind: str, line: LineData, matched_text: str, number: str) -> HeadingMatch:
    matched_words = matched_text.split()
    line_words_upper = [word.upper() for _, _, _, _, word in line.word_items]
    matched_words_upper = [word.upper() for word in matched_words]

    start_index = -1
    for index in range(len(line_words_upper) - len(matched_words_upper) + 1):
        if line_words_upper[index : index + len(matched_words_upper)] == matched_words_upper:
            start_index = index
            break

    if start_index == -1:
        bbox = line.bbox
    else:
        selected_words = line.word_items[start_index : start_index + len(matched_words_upper)]
        bbox = (
            min(item[0] for item in selected_words),
            min(item[1] for item in selected_words),
            max(item[2] for item in selected_words),
            max(item[3] for item in selected_words),
        )

    return HeadingMatch(
        kind=kind,
        number=_normalize_number(number),
        page_number=line.page_number,
        bbox=bbox,
        text=matched_text,
    )


def detect_headings(pdf_path: Path) -> list[HeadingMatch]:
    with fitz.open(pdf_path) as document:
        matches: list[HeadingMatch] = []
        for page in document:
            for line in _group_words_into_lines(page):
                worksheet_match = WORKSHEET_PATTERN.search(line.text)
                if worksheet_match:
                    matches.append(
                        _build_heading_match("Worksheet", line, worksheet_match.group(0), worksheet_match.group(1))
                    )
                synopsis_match = SYNOPSIS_PATTERN.search(line.text)
                if synopsis_match:
                    matches.append(
                        _build_heading_match("Synopsis", line, synopsis_match.group(0), synopsis_match.group(1))
                    )
        matches.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0], item.kind))
        return matches


def _detect_synopsis_images_on_page(doc: fitz.Document, page: fitz.Page) -> list[tuple[float, float, float, float]]:
    template = _get_synopsis_template()
    if template is None:
        return []

    try:
        import cv2
        import numpy as np
    except ImportError:
        return []

    matches: list[tuple[float, float, float, float]] = []
    for img in page.get_images():
        xref = img[0]
        rects = page.get_image_rects(xref)
        for r in rects:
            if r.y0 < 250:
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.width < 80 or pix.height < 15:
                        continue
                    if pix.n >= 3:
                        pix_rgb = fitz.Pixmap(fitz.csRGB, pix) if pix.colorspace.name != "DeviceRGB" else pix
                        arr = np.frombuffer(pix_rgb.samples, dtype=np.uint8).reshape((pix_rgb.height, pix_rgb.width, 3))
                        gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
                    else:
                        gray = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width))

                    norm = cv2.resize(gray, (120, 24))
                    _, b = cv2.threshold(norm, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    score = float(np.mean(template == b))
                    if score >= 0.85:
                        matches.append((r.x0, r.y0, r.x1, r.y1))
                except Exception:
                    continue
    return matches


def detect_all_boundaries(pdf_path: Path) -> list[BoundaryMarker]:
    """
    Detects both Worksheet start boundaries and Synopsis start boundaries in document order.
    """
    with fitz.open(pdf_path) as document:
        markers: list[BoundaryMarker] = []
        seen_ws_pages: set[int] = set()
        worksheet_index = 0

        # Pass 1: Look for CUQ lines (Standard for Olympiad / e-Techno materials)
        for page in document:
            for line in _group_words_into_lines(page):
                if line.page_number in seen_ws_pages:
                    continue
                if CUQ_PATTERN.match(line.text):
                    seen_ws_pages.add(line.page_number)
                    worksheet_index += 1
                    markers.append(
                        BoundaryMarker(
                            kind="worksheet",
                            name=f"WS-{worksheet_index}",
                            page_number=line.page_number,
                            bbox=line.bbox,
                            trigger_text=line.text,
                            reason="Found 'CUQ' directly below worksheet banner.",
                        )
                    )

        # Pass 2: If no CUQ found, check for explicit WORKSHEET patterns
        if not markers:
            seen_ws: set[str] = set()
            for page in document:
                for line in _group_words_into_lines(page):
                    match = WORKSHEET_PATTERN.search(line.text)
                    if match:
                        ws_num = match.group(1)
                        ws_key = f"WS-{int(ws_num)}"
                        if ws_key not in seen_ws and line.page_number not in seen_ws_pages:
                            seen_ws.add(ws_key)
                            seen_ws_pages.add(line.page_number)
                            markers.append(
                                BoundaryMarker(
                                    kind="worksheet",
                                    name=ws_key,
                                    page_number=line.page_number,
                                    bbox=line.bbox,
                                    trigger_text=match.group(0),
                                    reason=f"Found explicit header '{match.group(0)}'.",
                                )
                            )

        # Pass 3: Look for SYNOPSIS boundaries (Text lines & Graphic banners)
        synopsis_index = 0
        seen_synopsis_pages: set[int] = set()
        for page in document:
            page_num = page.number + 1
            found_on_page = False

            # 3a. Text lines check
            for line in _group_words_into_lines(page):
                syn_match = SYNOPSIS_PATTERN.search(line.text)
                if syn_match:
                    syn_num = syn_match.group(1) if syn_match.groups() and syn_match.group(1) else ""
                    syn_name = f"Synopsis {syn_num}".strip() if syn_num else f"Synopsis {synopsis_index + 1}"
                    synopsis_index += 1
                    markers.append(
                        BoundaryMarker(
                            kind="synopsis",
                            name=syn_name,
                            page_number=page_num,
                            bbox=line.bbox,
                            trigger_text=line.text,
                            reason="Found 'SYNOPSIS' in text.",
                        )
                    )
                    seen_synopsis_pages.add(page_num)
                    found_on_page = True
                    break

            if found_on_page:
                continue

            # 3b. Graphic image banner check (handles outlined vector titles with image drop-shadows)
            img_bboxes = _detect_synopsis_images_on_page(document, page)
            for bbox in img_bboxes:
                synopsis_index += 1
                markers.append(
                    BoundaryMarker(
                        kind="synopsis",
                        name=f"Synopsis {synopsis_index}",
                        page_number=page_num,
                        bbox=bbox,
                        trigger_text="SYNOPSIS BANNER",
                        reason="Matched graphic SYNOPSIS header banner.",
                    )
                )
                seen_synopsis_pages.add(page_num)
                break

        markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
        return markers


def detect_worksheet_candidates(pdf_path: Path) -> list[WorksheetCandidate]:
    """Returns worksheet candidates for backwards compatibility."""
    boundaries = detect_all_boundaries(pdf_path)
    candidates: list[WorksheetCandidate] = []
    for b in boundaries:
        if b.kind == "worksheet":
            candidates.append(
                WorksheetCandidate(
                    worksheet_name=b.name,
                    page_number=b.page_number,
                    trigger_text=b.trigger_text,
                    bbox=b.bbox,
                    reason=b.reason,
                )
            )
    return candidates


def build_worksheet_ranges(
    candidates_or_markers: list[WorksheetCandidate | BoundaryMarker],
    total_pages: int,
    stop_at_synopsis: bool = True,
) -> list[WorksheetRange]:
    """
    Builds non-overlapping worksheet page ranges.
    When stop_at_synopsis is True, each worksheet stops before any subsequent Synopsis section starts.
    """
    ranges: list[WorksheetRange] = []

    for idx, marker in enumerate(candidates_or_markers):
        if getattr(marker, "kind", "worksheet") != "worksheet":
            continue

        # Look forward for the next relevant boundary
        next_boundary = None
        for candidate in candidates_or_markers[idx + 1 :]:
            c_kind = getattr(candidate, "kind", "worksheet")
            if stop_at_synopsis or c_kind == "worksheet":
                next_boundary = candidate
                break

        name = getattr(marker, "worksheet_name", None) or getattr(marker, "name", "WS")
        start_page = marker.page_number

        if next_boundary is not None:
            if next_boundary.page_number == start_page:
                end_page = start_page
                end_bbox = next_boundary.bbox
            elif next_boundary.bbox[1] < 250:
                # Next boundary starts near the top of the subsequent page: exclude that page completely
                end_page = max(start_page, next_boundary.page_number - 1)
                end_bbox = None
            else:
                # Next boundary starts mid-page: include the page up to the boundary bbox
                end_page = next_boundary.page_number
                end_bbox = next_boundary.bbox
        else:
            end_page = total_pages
            end_bbox = None

        ranges.append(
            WorksheetRange(
                worksheet_name=name,
                start_page=start_page,
                end_page=end_page,
                start_bbox=marker.bbox,
                end_bbox=end_bbox,
            )
        )

    return ranges
