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
SSC_START_PATTERN = re.compile(
    r"\b(?:MULTIPLE\s+CHOICE\s+QUESTIONS|OBJECTIVE\s+TYPE\s+QUESTIONS|OBJECTIVE\s+EXERCISE)\b",
    re.IGNORECASE,
)
SSC_KEY_PATTERN = re.compile(r"^\s*[*#]*\s*(?:KEY|ANSWER\s+KEY)(?:\s*:|\s*$)", re.IGNORECASE)
CBSE_START_PATTERN = re.compile(
    r"\b(?:ASSESSMENT\s+SHEET\s*[-–—]?\s*\d+|MULTIPLE\s+CHOICE\s+QUESTIONS|OBJECTIVE\s+EXERCISE|OBJECTIVE\s+TYPE\s+QUESTIONS)\b",
    re.IGNORECASE,
)
CBSE_KEY_PATTERN = re.compile(
    r"\b(?:(?:OBJECTIVE\s+EXERCISE\s*[-–—]?\s*)?KEY|ANSWER\s+KEY)\b",
    re.IGNORECASE,
)

# Compact binary template (24x120 pixels) for graphic SYNOPSIS header badges
B64_SYNOPSIS_TEMPLATE = (
    "eJx9z8FtwzAMBVDSDPp7MKIFimoTe7HCpqIBOkJXYdBD1xDQBQTkkkNghXJy6aU8iHgQQH629qe2"
    "ub+3tbVrb5WCItegoYxcUCicJBsUemATIwySVVTowComwwtnZeWd6buTWUkAyemiR8piyfAD+dXm"
    "PMEpDdw02jsNnbSCZsOTiSI0mJSJhmA+OfjywmV69V/fDJhUrstOGFBkc46d4Qrc5Cs5xTmtwCyf"
    "57ocRTRtEYietpM4bZ6Z/JK6BNEn8aCJc+zETjjfPEmJzlzjuekH29z+rTtGWZhZ"
)
_CACHED_SYNOPSIS_TEMPLATE = None
_GLOBAL_OCR = None


def _get_ocr():
    global _GLOBAL_OCR
    if _GLOBAL_OCR is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _GLOBAL_OCR = RapidOCR()
        except Exception as err:
            logger.debug("Could not initialize RapidOCR: %s", err)
            _GLOBAL_OCR = False
    return _GLOBAL_OCR if _GLOBAL_OCR is not False else None



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
    kind: str  # "worksheet" or "synopsis" or "key"
    name: str
    page_number: int
    bbox: tuple[float, float, float, float]
    trigger_text: str
    reason: str
    extra_pages: tuple[int, ...] = ()


@dataclass(frozen=True)
class WorksheetRange:
    worksheet_name: str
    start_page: int
    end_page: int
    start_bbox: tuple[float, float, float, float]
    end_bbox: tuple[float, float, float, float] | None = None
    extra_pages: tuple[int, ...] = ()


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


def _detect_key_on_page(page: fitz.Page) -> tuple[float, float, float, float] | None:
    lines = [line.strip() for line in page.get_text().splitlines() if line.strip()]

    # 1. Text header: standalone "KEY" or "ANSWER KEY" or table header containing Q.NO & KEY
    for line in lines:
        if line.upper() in {"KEY", "ANSWER KEY"} or re.match(r"^[*#]*\s*(?:KEY|ANSWER\s+KEY)\s*$", line, re.IGNORECASE):
            rects = page.search_for(line)
            if rects:
                return (rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1)
        if "Q.NO" in line.upper() and "KEY" in line.upper():
            rects = page.search_for("KEY")
            if rects:
                return (rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1)

    # 2. Vector badge "KEY" centered on page (red badge with white outline)
    drawings = page.get_drawings()
    badge_candidates = [
        d for d in drawings
        if 200 <= d["rect"].y0 <= 500 and 25 <= d["rect"].width <= 80 and d.get("fill") and d["fill"][0] > 0.7 and d["fill"][1] < 0.2
    ]
    if badge_candidates:
        r = badge_candidates[0]["rect"]
        return (r.x0, r.y0, r.x1, r.y1)

    # 3. Check for consecutive answer key list: 1), 2), 3), 4), 5), 6)...
    nums: list[int] = []
    first_key_line: str | None = None
    for line in lines:
        m = re.match(r"^(\d{1,3})\)\s*[1-4A-D]?\s*$", line)
        if m:
            n = int(m.group(1))
            if not nums and n == 1:
                nums.append(1)
                first_key_line = line
            elif nums and n == nums[-1] + 1:
                nums.append(n)
        else:
            if len(nums) >= 6:
                break
            nums = []
            first_key_line = None

    if len(nums) >= 6 and first_key_line:
        rects = page.search_for(first_key_line)
        if rects:
            return (rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1)

    return None


def detect_ssc_boundaries(pdf_path: Path) -> list[BoundaryMarker]:
    """
    Detects AP SSC and TS SSC Multiple Choice / Objective Questions boundaries.
    """
    with fitz.open(pdf_path) as document:
        markers: list[BoundaryMarker] = []
        section_index = 0
        seen_start_pages: set[int] = set()

        for page in document:
            page_num = page.number + 1
            for line in _group_words_into_lines(page):
                match = SSC_START_PATTERN.search(line.text)
                if match and page_num not in seen_start_pages:
                    seen_start_pages.add(page_num)
                    section_index += 1
                    markers.append(
                        BoundaryMarker(
                            kind="worksheet",
                            name=f"Objective-Q{section_index}" if section_index > 1 else "Objective-Questions",
                            page_number=page_num,
                            bbox=line.bbox,
                            trigger_text=match.group(0),
                            reason=f"Found '{match.group(0)}' header.",
                        )
                    )
                    break

        if not markers:
            return []

        # Find KEY boundaries following the first start marker
        for page in document:
            page_num = page.number + 1
            if page_num < markers[0].page_number:
                continue
            kb = _detect_key_on_page(page)
            if kb is not None:
                markers.append(
                    BoundaryMarker(
                        kind="key",
                        name="Answer-Key",
                        page_number=page_num,
                        bbox=kb,
                        trigger_text="KEY",
                        reason="Found Answer Key section.",
                    )
                )

        markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
        return markers


def detect_cbse_boundaries(pdf_path: Path) -> list[BoundaryMarker]:
    """
    Detects CBSE Objective Exercise, Assessment Sheets, and Multiple Choice Questions boundaries.
    Optimized for sub-second execution across digital text, image banners, and vector outline documents.
    """
    with fitz.open(pdf_path) as document:
        markers: list[BoundaryMarker] = []
        total_pages = len(document)
        start_scan_page = max(0, int(total_pages * 0.60))

        # Check total text in the last 40% of the book
        scan_text = "".join(document[p].get_text() for p in range(start_scan_page, total_pages))

        if len(scan_text.strip()) > 100:
            # Mode A: Digital text is present (e.g. Chemistry Class 8)
            seen_pages: set[int] = set()
            found_key = False
            for pno in range(start_scan_page + 1, total_pages + 1):
                p = document[pno - 1]
                lines = [l.strip() for l in p.get_text().splitlines() if l.strip()]
                for line in lines:
                    if not found_key and pno > 10:
                        km = CBSE_KEY_PATTERN.search(line)
                        if km and ("KEY" in line.upper()):
                            rects = p.search_for(line)
                            bbox = (rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1) if rects else (0, 0, p.rect.width, 100)
                            markers.append(
                                BoundaryMarker(
                                    kind="key",
                                    name="Answer-Key",
                                    page_number=pno,
                                    bbox=bbox,
                                    trigger_text=line,
                                    reason="Found Answer Key section.",
                                )
                            )
                            found_key = True
                            break

                # Collect candidate header matches on this page, prioritizing ASSESSMENT > OBJECTIVE > MULTIPLE CHOICE
                page_start_matches = []
                for line in lines:
                    sm = CBSE_START_PATTERN.search(line)
                    if sm:
                        raw_name = sm.group(0).strip()
                        prio = 1 if "ASSESSMENT" in raw_name.upper() else (2 if "OBJECTIVE" in raw_name.upper() else 3)
                        page_start_matches.append((prio, raw_name))

                if not found_key and page_start_matches and pno not in seen_pages:
                    seen_pages.add(pno)
                    page_start_matches.sort(key=lambda x: x[0])
                    raw_name = page_start_matches[0][1]
                    if "ASSESSMENT" in raw_name.upper():
                        num_m = re.search(r"\d+", raw_name)
                        sheet_no = num_m.group(0) if num_m else "1"
                        name = f"Assessment-Sheet-{sheet_no}"
                    elif "OBJECTIVE" in raw_name.upper():
                        name = "Objective-Exercise"
                    else:
                        name = "Objective-Questions"

                    rects = p.search_for(raw_name)
                    bbox = (rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1) if rects else (0, 0, p.rect.width, 100)
                    markers.append(
                        BoundaryMarker(
                            kind="worksheet",
                            name=name,
                            page_number=pno,
                            bbox=bbox,
                            trigger_text=raw_name,
                            reason=f"Found '{raw_name}' header.",
                        )
                    )

            markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
            return markers

        # Mode B: Vector drawings (Print to PDF e.g. Chemistry Class 6, Maths Class 6)
        green_sheets = []
        blue_key_page = None
        blue_key_bbox = None
        red_headings = []
        key_table_page = None
        key_table_bbox = None

        for pno in range(start_scan_page + 1, total_pages + 1):
            p = document[pno - 1]
            drawings = p.get_drawings()
            if not drawings:
                continue

            # Check red heading drawings: "V. MULTIPLE CHOICE QUESTIONS" (fill: (0.5, 0.0, 0.0))
            reds = [d for d in drawings if d.get("fill") and abs(d["fill"][0] - 0.5) < 0.05 and d["fill"][1] < 0.05 and d["fill"][2] < 0.05]
            if len(reds) >= 10:
                y0_red = min(d["rect"].y0 for d in reds)
                y1_red = max(d["rect"].y1 for d in reds)
                x0_red = min(d["rect"].x0 for d in reds)
                x1_red = max(d["rect"].x1 for d in reds)
                if y0_red > 200:
                    red_headings.append((pno, (x0_red, y0_red, x1_red, y1_red)))

            for d in drawings:
                if d.get("fill") and d["rect"].width > 120 and d["rect"].height > 18:
                    fill = d["fill"]
                    # Green banner for Chemistry Assessment Sheets: fill=[0, 0.69, 0.31]
                    if abs(fill[0] - 0.0) < 0.05 and abs(fill[1] - 0.69) < 0.05 and abs(fill[2] - 0.31) < 0.05:
                        green_sheets.append((pno, (d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1)))
                        break
                    # Blue ribbon for Maths Key: fill=[0, 0.44, 0.75]
                    elif d["rect"].y0 < 150 and abs(fill[0]) < 0.05 and abs(fill[1] - 0.44) < 0.06 and abs(fill[2] - 0.75) < 0.06:
                        blue_key_page = pno
                        blue_key_bbox = (d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1)

            # Check for Key table on final page of Chemistry (tables with multiple colored cells)
            if pno == total_pages and green_sheets and not key_table_bbox:
                table_cells = [d for d in drawings if d.get("fill") and d["rect"].height > 15 and d["rect"].y0 > 250]
                if len(table_cells) >= 10:
                    min_y0 = min(d["rect"].y0 for d in table_cells)
                    key_table_page = pno
                    key_table_bbox = (0, min_y0, p.rect.width, min_y0 + 50)

        if green_sheets:
            # Chemistry with vector green banners (e.g. Natures treasure 6 CBSE)
            for idx, (pno, bbox) in enumerate(green_sheets):
                markers.append(
                    BoundaryMarker(
                        kind="worksheet",
                        name=f"Assessment-Sheet-{idx+1}",
                        page_number=pno,
                        bbox=bbox,
                        trigger_text=f"ASSESSMENT SHEET-{idx+1}",
                        reason=f"Found green Assessment Sheet-{idx+1} vector banner.",
                    )
                )
            key_p = key_table_page or total_pages
            key_b = key_table_bbox or (0, 265.0, document[-1].rect.width, 350.0)
            markers.append(
                BoundaryMarker(
                    kind="key",
                    name="Answer-Key",
                    page_number=key_p,
                    bbox=key_b,
                    trigger_text="KEY",
                    reason="Found Key page at end of chapter.",
                )
            )
            markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
            return markers

        if red_headings and blue_key_page:
            # Maths with vector outlines (e.g. Data handling and presentation 6 CBSE)
            pno_mcq, bbox_mcq = red_headings[0]
            markers.append(
                BoundaryMarker(
                    kind="worksheet",
                    name="Objective-Questions",
                    page_number=pno_mcq,
                    bbox=bbox_mcq,
                    trigger_text="MULTIPLE CHOICE QUESTIONS",
                    reason="Found vector Multiple Choice Questions heading.",
                )
            )
            markers.append(
                BoundaryMarker(
                    kind="key",
                    name="Answer-Key",
                    page_number=blue_key_page,
                    bbox=blue_key_bbox or (200.0, 62.0, 410.0, 104.0),
                    trigger_text="KEY",
                    reason="Found blue ribbon Key banner.",
                )
            )
            markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
            return markers

        # Mode C: Graphic Image Banners (Physics e.g. Heat Transfer in Nature 7 CBSE)
        # Fast color-signature analysis (< 0.1s) avoiding heavy OCR
        import numpy as np
        found_img_start = False
        found_img_key = False
        fallback_candidates = []

        for pno in range(start_scan_page + 1, total_pages + 1):
            p = document[pno - 1]
            images = p.get_images()
            candidate_imgs = [img for img in images if img[2] > 800 and 2.5 <= img[2] / img[3] <= 6.0]
            for img in candidate_imgs:
                xref = img[0]
                rects = p.get_image_rects(xref)
                for r in rects:
                    if r.width > 220 and 35 <= r.height <= 95 and r.y0 < 300 and 100 <= r.x0 <= 250:
                        try:
                            pix = fitz.Pixmap(document, xref)
                            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                            # Purple banner for Objective Exercise: [165, 72, 163]
                            purple_count = np.sum((arr[:, :, 0] > 140) & (arr[:, :, 1] < 100) & (arr[:, :, 2] > 130))
                            # Cyan banner for Key: [1, 176, 239]
                            cyan_count = np.sum((arr[:, :, 0] < 30) & (arr[:, :, 1] > 150) & (arr[:, :, 2] > 200))

                            if purple_count > 10000 and not found_img_start:
                                markers.append(
                                    BoundaryMarker(
                                        kind="worksheet",
                                        name="Objective-Exercise",
                                        page_number=pno,
                                        bbox=(r.x0, r.y0, r.x1, r.y1),
                                        trigger_text="OBJECTIVE EXERCISE",
                                        reason="Found Objective Exercise banner via color profile.",
                                    )
                                )
                                found_img_start = True
                            elif cyan_count > 10000 and not found_img_key:
                                markers.append(
                                    BoundaryMarker(
                                        kind="key",
                                        name="Answer-Key",
                                        page_number=pno,
                                        bbox=(r.x0, r.y0, r.x1, r.y1),
                                        trigger_text="OBJECTIVE EXERCISE-KEY",
                                        reason="Found Objective Exercise Key banner via color profile.",
                                    )
                                )
                                found_img_key = True
                            else:
                                fallback_candidates.append((pno, xref, r))
                        except Exception:
                            pass
            if found_img_start and found_img_key:
                break

        # Fallback to targeted OCR only if fast color analysis did not find both banners
        if not (found_img_start and found_img_key) and fallback_candidates:
            ocr = _get_ocr()
            if ocr is not None:
                for pno, xref, r in fallback_candidates:
                    try:
                        pix = fitz.Pixmap(document, xref)
                        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                        res, _ = ocr(arr)
                        if res:
                            txt = re.sub(r'[^A-Z0-9]', '', res[0][1].upper())
                            if ("OBJECTIVEEXERCISEKEY" in txt or "ANSWERKEY" in txt) and not found_img_key:
                                markers.append(
                                    BoundaryMarker(
                                        kind="key",
                                        name="Answer-Key",
                                        page_number=pno,
                                        bbox=(r.x0, r.y0, r.x1, r.y1),
                                        trigger_text=res[0][1],
                                        reason="Found Objective Exercise Key banner via OCR fallback.",
                                    )
                                )
                                found_img_key = True
                            elif ("OBJECTIVEEXERCISE" in txt or "MULTIPLECHOICE" in txt) and not found_img_start:
                                markers.append(
                                    BoundaryMarker(
                                        kind="worksheet",
                                        name="Objective-Exercise",
                                        page_number=pno,
                                        bbox=(r.x0, r.y0, r.x1, r.y1),
                                        trigger_text=res[0][1],
                                        reason="Found Objective Exercise banner via OCR fallback.",
                                    )
                                )
                                found_img_start = True
                    except Exception:
                        pass
                    if found_img_start and found_img_key:
                        break

        markers.sort(key=lambda item: (item.page_number, item.bbox[1], item.bbox[0]))
        return markers


def detect_all_boundaries(pdf_path: Path, preset: str = "olympiad") -> list[BoundaryMarker]:
    if preset == "ssc":
        return detect_ssc_boundaries(pdf_path)
    if preset == "cbse":
        return detect_cbse_boundaries(pdf_path)

    # Olympiad / e-Techno preset
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
    preset: str = "olympiad",
    include_key: bool = True,
) -> list[WorksheetRange]:
    """
    Builds non-overlapping worksheet page ranges.
    When preset is 'ssc':
      Each objective section stops at the KEY boundary (inclusive if include_key=True).
    When preset is 'olympiad':
      When stop_at_synopsis is True, each worksheet stops before any subsequent Synopsis section starts.
    """
    ranges: list[WorksheetRange] = []

    for idx, marker in enumerate(candidates_or_markers):
        if getattr(marker, "kind", "worksheet") != "worksheet":
            continue

        name = getattr(marker, "worksheet_name", None) or getattr(marker, "name", "WS")
        start_page = marker.page_number

        if preset in {"ssc", "cbse"}:
            key_markers = [m for m in candidates_or_markers if getattr(m, "kind", "") == "key"]
            global_key = key_markers[0] if key_markers else None

            next_marker = None
            for candidate in candidates_or_markers[idx + 1 :]:
                c_kind = getattr(candidate, "kind", "")
                if c_kind in {"key", "worksheet"}:
                    next_marker = candidate
                    break

            extra_pages = ()
            if next_marker is not None:
                c_kind = getattr(next_marker, "kind", "")
                if c_kind == "key":
                    if not include_key:
                        # Stop before key: if key is at top of page, end on previous page; otherwise bottom-crop right before key
                        if next_marker.bbox[1] < 150.0 and next_marker.page_number > start_page:
                            end_page = next_marker.page_number - 1
                            end_bbox = None
                        else:
                            end_page = next_marker.page_number
                            end_bbox = next_marker.bbox
                    else:
                        # Include key: if multi-page gap (Maths Option A), append end-of-book key table
                        if total_pages - next_marker.page_number >= 2:
                            end_page = next_marker.page_number - 1
                            end_bbox = None
                            extra_pages = (total_pages,)
                        else:
                            end_page = next_marker.page_number
                            end_bbox = None
                else:
                    # Followed by another worksheet (e.g. Assessment Sheet 1 -> 2)
                    end_page = next_marker.page_number
                    end_bbox = next_marker.bbox
                    if include_key and global_key:
                        extra_pages = (global_key.page_number,)
            else:
                end_page = total_pages
                end_bbox = None
        else:
            # Look forward for the next relevant boundary
            next_boundary = None
            for candidate in candidates_or_markers[idx + 1 :]:
                c_kind = getattr(candidate, "kind", "worksheet")
                if stop_at_synopsis or c_kind == "worksheet":
                    next_boundary = candidate
                    break

            if next_boundary is not None:
                if next_boundary.page_number == start_page:
                    end_page = start_page
                    end_bbox = next_boundary.bbox
                elif next_boundary.bbox[1] < 250:
                    end_page = max(start_page, next_boundary.page_number - 1)
                    end_bbox = None
                else:
                    end_page = next_boundary.page_number
                    end_bbox = next_boundary.bbox
            else:
                end_page = total_pages
                end_bbox = None
            extra_pages = ()

        ranges.append(
            WorksheetRange(
                worksheet_name=name,
                start_page=start_page,
                end_page=end_page,
                start_bbox=marker.bbox,
                end_bbox=end_bbox,
                extra_pages=extra_pages,
            )
        )

    return ranges
