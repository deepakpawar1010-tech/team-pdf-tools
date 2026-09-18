from __future__ import annotations

import pymupdf as fitz

from .config import FIRST_PAGE_BANNER_MARGIN, SAME_PAGE_BOTTOM_MARGIN
from .detector import WorksheetRange


def crop_pages_for_worksheet(
    document: fitz.Document,
    worksheet_range: WorksheetRange,
    preset: str = "olympiad",
    crop_top: bool = True,
) -> None:
    if document.page_count == 0:
        return

    first_page = document[0]
    first_page_rect = first_page.rect

    if crop_top:
        if preset in {"ssc", "cbse"}:
            # In SSC / CBSE mode: if the objective banner is mid-page (e.g. y > 120), crop from start_bbox[1] - 16.0
            # This cleanly removes subjective exercises above the MCQ banner (like Exercise 11.3 on TS SSC or Assessment Sheets on CBSE)
            if worksheet_range.start_bbox[1] > 120.0:
                first_page_top = max(0.0, worksheet_range.start_bbox[1] - 16.0)
                first_page.set_cropbox(fitz.Rect(0, first_page_top, first_page_rect.width, first_page_rect.height))
        else:
            # Olympiad mode: crop top banner margin
            first_page_top = max(0.0, worksheet_range.start_bbox[1] - FIRST_PAGE_BANNER_MARGIN)
            first_page.set_cropbox(fitz.Rect(0, first_page_top, first_page_rect.width, first_page_rect.height))

    if worksheet_range.end_bbox is not None:
        first_top = first_page.cropbox.y0
        if document.page_count == 1:
            same_page_rect = document[0].rect
            bottom_edge = max(first_top + 1, worksheet_range.end_bbox[1] - SAME_PAGE_BOTTOM_MARGIN)
            bottom_edge = min(bottom_edge, same_page_rect.height)
            document[0].set_cropbox(fitz.Rect(0, first_top, same_page_rect.width, bottom_edge))
        else:
            last_page = document[-1]
            last_page_rect = last_page.rect
            bottom_edge = max(1.0, worksheet_range.end_bbox[1] - SAME_PAGE_BOTTOM_MARGIN)
            bottom_edge = min(bottom_edge, last_page_rect.height)
            last_page.set_cropbox(fitz.Rect(0, 0, last_page_rect.width, bottom_edge))

