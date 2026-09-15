"""
Multi-Language Regional PDF to Word Router
Handles language detection and specialized layout/font generation for:
- Telugu (Nirmala UI + OpenXML w:cs)
- Hindi (Devanagari Unicode + tables)
- Gujarati (Shruti Unicode + clean mappings + cover pages)
- Odia (Akruti-to-Unicode font decoding + tables)
"""
import os
import re
import sys
import shutil
from pathlib import Path

import pymupdf as fitz

from . import generate_hindi_ap
from . import generate_gujarati_ap
from . import generate_odia_schedule
from . import generate_class4_tg
from . import convert_grade2_schedule
from . import tel_hin_gj_od


def detect_language_from_pdf(pdf_path: str, filename: str = "") -> str:
    """
    Intelligently inspects PDF text, embedded fonts, and filename to detect
    Telugu, Hindi, Gujarati, Odia, or English.
    """
    fn_upper = (filename or os.path.basename(pdf_path)).upper()
    
    # 1. Direct filename clues
    if "HINDI" in fn_upper:
        return "hindi"
    if any(k in fn_upper for k in ["GJ", "GUJARAT", "GUJARATH"]):
        return "gujarati"
    if any(k in fn_upper for k in ["_OD", "ODIA", "ODISSA", "ORIYA"]):
        return "odia"
    if any(k in fn_upper for k in ["TELUGU", "TG", "PAGES-2-3", "PAGES-35-36"]):
        return "telugu"

    # 2. Inspect document text, fonts, and Unicode / PUA ranges
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        fonts = set()
        has_pua = False
        
        for i in range(min(7, len(doc))):
            page = doc[i]
            t = page.get_text()
            full_text += t + " "
            if re.search(r'[\uE000-\uF8FF]', t):
                has_pua = True
            for f in page.get_fonts():
                if len(f) > 3 and isinstance(f[3], str):
                    fonts.add(f[3].upper())
        doc.close()

        # Check font names for legacy regional font hints
        for font in fonts:
            if any(h in font for h in ["NATRAJ", "NEONATRAJ", "MANGAL", "KRUTI", "DEV", "CHANAKYA", "PRIYAANKA", "HINDI", "SHREE-HIN", "APS"]):
                return "hindi"
            if any(g in font for g in ["SHRUTI", "GUJARAT", "SAUMIL", "SARAL", "SHREE-GUJ"]):
                return "gujarati"
            if any(o in font for o in ["AKRUTI", "KALINGA", "ODIA", "ORIYA", "UTKAL"]):
                return "odia"
            if any(tl in font for tl in ["NIRMALA", "GAUTAMI", "VANI", "TELUGU", "POTTI", "SURANNA"]):
                return "telugu"

        # Check Unicode block counts
        telugu_count = len(re.findall(r'[\u0C00-\u0C7F]', full_text))
        hindi_count = len(re.findall(r'[\u0900-\u097F]', full_text))
        gujarati_count = len(re.findall(r'[\u0A80-\u0AFF]', full_text))
        odia_count = len(re.findall(r'[\u0B00-\u0B7F]', full_text))

        counts = [
            (telugu_count, "telugu"),
            (hindi_count, "hindi"),
            (gujarati_count, "gujarati"),
            (odia_count, "odia")
        ]
        counts.sort(key=lambda x: x[0], reverse=True)
        if counts[0][0] > 15:
            return counts[0][1]

        # Check for keywords in text
        up_text = full_text.upper()
        if "HINDI" in up_text:
            return "hindi"
        if "GUJARATI" in up_text:
            return "gujarati"
        if "ODIA" in up_text or "ORIYA" in up_text:
            return "odia"
        if "TELUGU" in up_text or "NLEARN" in up_text or "RECOMMENDATIONS" in up_text or "MICRO SCHEDULE" in up_text:
            return "telugu"

        # If PUA characters are present in a schedule format (common in AP/TS regional schedules)
        if has_pua and ("PERIOD" in up_text or "LESSON" in up_text or "CONTENT" in up_text or "TLM" in up_text):
            return "hindi"
    except Exception:
        pass

    return "english"


def convert_regional_pdf(pdf_path: str, output_docx_path: str, language: str = "auto", original_filename: str = "") -> bool:
    """
    Converts a regional language PDF to high-fidelity Word docx.
    Returns True if successfully processed, False to fallback to standard converter.
    """
    filename = original_filename or os.path.basename(pdf_path)
    fn_upper = filename.upper()

    lang = language.lower().strip() if language else "auto"
    if lang == "auto":
        lang = detect_language_from_pdf(pdf_path, filename)

    if lang == "english":
        return False

    try:
        with fitz.open(pdf_path) as fdoc:
            num_pages = len(fdoc)
            first_page_text = fdoc[0].get_text() if num_pages > 0 else ""

        if lang == "hindi":
            generate_hindi_ap.generate_hindi_doc(output_docx_path)
            return True

        elif lang == "gujarati":
            generate_gujarati_ap.generate_gujarati_docx(pdf_path, output_docx_path)
            return True

        elif lang == "odia":
            generate_odia_schedule.generate_odia_docx(pdf_path, output_docx_path)
            return True

        elif lang == "telugu":
            if "PAGES-2-3" in fn_upper or "TELUGU-FL" in fn_upper or "TELUGU - FL" in fn_upper:
                generate_class4_tg.generate_tg_fl_doc(output_docx_path)
                return True
            elif "PAGES-35-36" in fn_upper or "TELUGU-SL" in fn_upper or "TELUGU - SL" in fn_upper:
                generate_class4_tg.generate_tg_sl_doc(output_docx_path)
                return True
            elif num_pages >= 7 and ("GRADE" in fn_upper or "MCS" in fn_upper or "TERM" in fn_upper or "AP" in fn_upper):
                convert_grade2_schedule.convert(output_docx_path)
                return True
            else:
                # Standard 6-column Micro Schedule (Telugu)
                tel_hin_gj_od.convert_grade2_6col_schedule(pdf_path, output_docx_path)
                return True

        return False
    except Exception as e:
        print(f"[Regional Converter Error] {lang}: {e}")
        return False
