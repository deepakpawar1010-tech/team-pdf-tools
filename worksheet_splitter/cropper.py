from __future__ import annotations

import pymupdf as fitz

from .config import FIRST_PAGE_BANNER_MARGIN, SAME_PAGE_BOTTOM_MARGIN
from .detector import WorksheetRange


def crop_pages_for_worksheet(document: fitz.Document, worksheet_range: WorksheetRange) -> None:
    if document.page_count == 0:
        return

    first_page = document[0]
    first_page_rect = first_page.rect
    first_page_top = max(0.0, worksheet_range.start_bbox[1] - FIRST_PAGE_BANNER_MARGIN)
    first_page.set_cropbox(fitz.Rect(0, first_page_top, first_page_rect.width, first_page_rect.height))

    if worksheet_range.end_bbox is not None:
        if document.page_count == 1:
            same_page_rect = document[0].rect
            bottom_edge = max(first_page_top + 1, worksheet_range.end_bbox[1] - SAME_PAGE_BOTTOM_MARGIN)
            bottom_edge = min(bottom_edge, same_page_rect.height)
            document[0].set_cropbox(fitz.Rect(0, first_page_top, same_page_rect.width, bottom_edge))
        else:
            last_page = document[-1]
            last_page_rect = last_page.rect
            bottom_edge = max(1.0, worksheet_range.end_bbox[1] - SAME_PAGE_BOTTOM_MARGIN)
            bottom_edge = min(bottom_edge, last_page_rect.height)
            last_page.set_cropbox(fitz.Rect(0, 0, last_page_rect.width, bottom_edge))

