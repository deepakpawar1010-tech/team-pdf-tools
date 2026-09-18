from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pymupdf as fitz


WORKSHEET_PATTERN = re.compile(r"\bWORKSHEET\s*-?\s*(\d{1,3})\b", re.IGNORECASE)
SYNOPSIS_PATTERN = re.compile(r"\bSYNOPSIS\s*-?\s*(\d{1,3})\b", re.IGNORECASE)
CUQ_PATTERN = re.compile(r"^\s*CUQ\s*$", re.IGNORECASE)


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


def detect_worksheet_candidates(pdf_path: Path) -> list[WorksheetCandidate]:
    with fitz.open(pdf_path) as document:
        candidates: list[WorksheetCandidate] = []
        seen_pages: set[int] = set()
        worksheet_index = 0

        # Pass 1: Look for CUQ lines (Standard for Olympiad / e-Techno materials)
        for page in document:
            for line in _group_words_into_lines(page):
                if line.page_number in seen_pages:
                    continue
                if CUQ_PATTERN.match(line.text):
                    seen_pages.add(line.page_number)
                    worksheet_index += 1
                    candidates.append(
                        WorksheetCandidate(
                            worksheet_name=f"WS-{worksheet_index}",
                            page_number=line.page_number,
                            trigger_text=line.text,
                            bbox=line.bbox,
                            reason="Found 'CUQ' directly below the worksheet banner, so worksheet names are assigned sequentially.",
                        )
                    )

        # Pass 2: If no CUQ found, check for explicit WORKSHEET patterns
        if not candidates:
            seen_ws: set[str] = set()
            for page in document:
                for line in _group_words_into_lines(page):
                    match = WORKSHEET_PATTERN.search(line.text)
                    if match:
                        ws_num = match.group(1)
                        ws_key = f"WS-{int(ws_num)}"
                        if ws_key not in seen_ws and line.page_number not in seen_pages:
                            seen_ws.add(ws_key)
                            seen_pages.add(line.page_number)
                            candidates.append(
                                WorksheetCandidate(
                                    worksheet_name=ws_key,
                                    page_number=line.page_number,
                                    trigger_text=match.group(0),
                                    bbox=line.bbox,
                                    reason=f"Found explicit header '{match.group(0)}'.",
                                )
                            )

        candidates.sort(key=lambda item: item.page_number)
        return candidates


def build_worksheet_ranges(candidates: list[WorksheetCandidate], total_pages: int) -> list[WorksheetRange]:
    ranges: list[WorksheetRange] = []
    for index, candidate in enumerate(candidates):
        next_candidate = candidates[index + 1] if index + 1 < len(candidates) else None
        end_page = next_candidate.page_number - 1 if next_candidate else total_pages
        end_bbox = next_candidate.bbox if next_candidate and next_candidate.page_number == candidate.page_number else None
        ranges.append(
            WorksheetRange(
                worksheet_name=candidate.worksheet_name,
                start_page=candidate.page_number,
                end_page=end_page,
                start_bbox=candidate.bbox,
                end_bbox=end_bbox,
            )
        )
    return ranges
