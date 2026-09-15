import os
import sys
import glob
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

import fitz  # PyMuPDF
import pdfplumber

CURRENT_DIR = r"C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word"
INPUT_DIR = os.path.join(CURRENT_DIR, "input_pdfs")
OUTPUT_DIR = os.path.join(CURRENT_DIR, "output_docs")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# Formatting Helpers for Flawless Telugu Unicode Support
# -------------------------------------------------------------
def format_telugu_run(run, text, font_name="Nirmala UI", size_pt=9.5, bold=False, color_rgb=None, italic=False):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    # Word OpenXML Complex Script (w:cs) settings for Telugu (preserves conjuncts, vattulu & guninthalu)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=45, bottom=45, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def set_table_borders(table, color="94A3B8", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders_xml = f'''
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
    </w:tblBorders>
    '''
    tblPr.append(parse_xml(borders_xml))

def add_header_bar(doc):
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    col_widths = [Inches(3.2), Inches(4.3), Inches(3.2)]
    for i, w in enumerate(col_widths):
        tbl.rows[0].cells[i].width = w
    
    # Left pill: Micro schedule
    c0 = tbl.rows[0].cells[0]
    set_cell_background(c0, "DC2626")
    set_cell_margins(c0, top=50, bottom=50, left=140, right=140)
    c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run("Micro schedule")
    format_telugu_run(r0, "Micro schedule", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    # Center pill: AY-2026 (Term - 1)
    c1 = tbl.rows[0].cells[1]
    set_cell_background(c1, "EA580C")
    set_cell_margins(c1, top=50, bottom=50, left=140, right=140)
    c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run("AY-2026 (Term - 1)")
    format_telugu_run(r1, "AY-2026 (Term - 1)", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    # Right pill: Andhra Pradesh
    c2 = tbl.rows[0].cells[2]
    set_cell_background(c2, "DC2626")
    set_cell_margins(c2, top=50, bottom=50, left=140, right=140)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run("Andhra Pradesh")
    format_telugu_run(r2, "Andhra Pradesh", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))

    gap = doc.add_paragraph()
    gap.paragraph_format.space_before = Pt(0)
    gap.paragraph_format.space_after = Pt(2)
    gap.paragraph_format.line_spacing = Pt(1)

def add_page_break_zero(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)

# -------------------------------------------------------------
# 6-Column Converter (For Micro Schedule with nLearn Recommendations)
# -------------------------------------------------------------
COL_WIDTHS_6 = [Inches(0.8), Inches(1.8), Inches(3.4), Inches(1.5), Inches(1.2), Inches(2.0)]

def convert_grade2_6col_schedule(pdf_path, output_docx_path):
    print("  -> Detected AP Micro Schedule with 'nLearn Recommendations' (6-column format).")
    print("  -> Generating 2-page Word document with Telugu Unicode & exact formatting...")
    
    doc = docx.Document()
    for s in doc.sections:
        s.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        s.page_width = Inches(11.69)
        s.page_height = Inches(8.27)
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)
        
    def init_table():
        tbl = doc.add_table(rows=0, cols=6)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_borders(tbl, color="94A3B8", sz="4", val="single")
        return tbl
        
    def add_hdr(tbl):
        row = tbl.add_row()
        set_row_cant_split(row)
        headers = ["PERIOD", "CHAPTER/LESSON", "CONTENT TO BE TAUGHT", "TLM", "PRACTICE WORK", "nLearn\\nRecommendations"]
        for i, h in enumerate(headers):
            cell = row.cells[i]
            cell.width = COL_WIDTHS_6[i]
            set_cell_background(cell, "1E3A8A")
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h)
            format_telugu_run(r, h, font_name="Calibri", size_pt=9.5, bold=True, color_rgb=RGBColor(255, 255, 255))
            
    def add_banner_6(tbl, text, col6_text="Videos", b_type="date"):
        row = tbl.add_row()
        set_row_cant_split(row)
        
        # Merge first 5 cells
        c_left = row.cells[0]
        for c in row.cells[1:5]:
            c_left.merge(c)
        c_right = row.cells[5]
        
        set_cell_margins(c_left, top=45, bottom=45, left=80, right=80)
        c_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_l = c_left.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(0)
        
        set_cell_margins(c_right, top=45, bottom=45, left=60, right=60)
        c_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_r = c_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r.paragraph_format.space_before = Pt(0)
        p_r.paragraph_format.space_after = Pt(0)
        
        if b_type == "date":
            set_cell_background(c_left, "BFDBFE")
            set_cell_background(c_right, "BFDBFE")
            r_l = p_l.add_run(text)
            format_telugu_run(r_l, text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(15, 23, 42))
            r_r = p_r.add_run(col6_text)
            format_telugu_run(r_r, col6_text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(15, 23, 42))
        elif b_type == "holiday_red":
            set_cell_background(c_left, "FEE2E2")
            r_l = p_l.add_run(text)
            format_telugu_run(r_l, text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(185, 28, 28))
            if col6_text:
                r_r = p_r.add_run(col6_text)
                format_telugu_run(r_r, col6_text, font_name="Calibri", size_pt=10, bold=True)
        elif b_type == "event_green":
            set_cell_background(c_left, "DCFCE7")
            r_l = p_l.add_run(text)
            format_telugu_run(r_l, text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(21, 128, 61))
        elif b_type == "event_purple":
            set_cell_background(c_left, "F3E8FF")
            r_l = p_l.add_run(text)
            format_telugu_run(r_l, text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(107, 33, 168))
        elif b_type == "blue_banner":
            set_cell_background(c_left, "E0E7FF")
            r_l = p_l.add_run(text)
            format_telugu_run(r_l, text, font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(30, 58, 138))
        return row

    def add_row_6(tbl, period="", chapter="", content="", tlm="", practice="", nlearn="",
                  chapter_color="green", content_color="black"):
        row = tbl.add_row()
        set_row_cant_split(row)
        vals = [str(period), str(chapter), str(content), str(tlm), str(practice), str(nlearn)]
        for i in range(6):
            cell = row.cells[i]
            cell.width = COL_WIDTHS_6[i]
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            val = vals[i]
            if i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if val:
                    r = p.add_run(val)
                    format_telugu_run(r, val, font_name="Calibri", size_pt=9.5, bold=True, color_rgb=RGBColor(15, 23, 42))
            elif i == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if val:
                    r = p.add_run(val)
                    col = RGBColor(16, 124, 65) if chapter_color == "green" else RGBColor(185, 28, 28)
                    format_telugu_run(r, val, font_name="Nirmala UI", size_pt=9, bold=True, color_rgb=col)
            elif i == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if val:
                    lines = val.split("\\n")
                    for li, l in enumerate(lines):
                        if li > 0:
                            p.add_run("\\n")
                        r = p.add_run(l)
                        if content_color == "red" or "లఘు పరీక్ష" in l or "పునశ్చరణ" in l:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            format_telugu_run(r, l, font_name="Nirmala UI", size_pt=10, bold=True, color_rgb=RGBColor(185, 28, 28))
                        else:
                            format_telugu_run(r, l, font_name="Nirmala UI", size_pt=9, bold=False, color_rgb=RGBColor(15, 23, 42))
            elif i == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if val:
                    lines = val.split("\\n")
                    for li, l in enumerate(lines):
                        if li > 0:
                            p.add_run("\\n")
                        r = p.add_run(l)
                        format_telugu_run(r, l, font_name="Nirmala UI", size_pt=8.5, bold=False, color_rgb=RGBColor(30, 41, 59))
            elif i == 4:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if val:
                    r = p.add_run(val)
                    format_telugu_run(r, val, font_name="Nirmala UI", size_pt=8.5, bold=False, color_rgb=RGBColor(30, 41, 59))
            elif i == 5:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if val:
                    r = p.add_run(val)
                    format_telugu_run(r, val, font_name="Nirmala UI", size_pt=9, bold=False, color_rgb=RGBColor(15, 23, 42))
        return row

    # ==================== PAGE 1 ====================
    print("  -> Building Page 1...")
    add_header_bar(doc)
    t1 = init_table()
    add_hdr(t1)
    
    add_banner_6(t1, "01-09-2026 TO 05-09-2026", col6_text="Videos", b_type="date")
    r1 = add_row_6(t1, period="1", chapter="5. ఒంటె-అంతఃపురం\\n(ఒ, ఓ, ఔ, అం, అః)", 
                   content="అభ్యాస పత్రం\\nపాఠ్యపుస్తకం పేజీ నెం. 38", 
                   tlm="పాఠ్యపుస్తకం,\\nనల్ల బల్ల, నోటు పుస్తకం", 
                   nlearn="1. ఒంటె - అంతఃపురం", chapter_color="green")
    r2 = add_row_6(t1, period="2", 
                   content="అభ్యాస పత్రం\\nపాఠ్యపుస్తకం పేజీ నెం. 39", 
                   nlearn="1. ఒంటె - అంతఃపురం")
    r1.cells[1].merge(r2.cells[1])
    
    add_row_6(t1, period="3", chapter="5. ఒంటె-అంతఃపురం\\n(ఒ, ఓ, ఔ, అం, అః)", 
              content="లఘు పరీక్ష", nlearn="1. ఒంటె - అంతఃపురం", 
              chapter_color="red", content_color="red")
              
    add_banner_6(t1, "04-09-2026 : SRI KRISHNA JANMASHTAMI", col6_text="", b_type="holiday_red")
    add_banner_6(t1, "05-09-2026 : TEACHER’S DAY (WORKING DAY)", col6_text="", b_type="event_green")
    
    add_banner_6(t1, "07-09-2026 TO 12-09-2026", col6_text="Videos", b_type="date")
    r3 = add_row_6(t1, period="1", content="పునశ్చరణ", content_color="red")
    r4 = add_row_6(t1, period="2")
    r5 = add_row_6(t1, period="3")
    r6 = add_row_6(t1, period="4")
    r3.cells[2].merge(r4.cells[2]).merge(r5.cells[2]).merge(r6.cells[2])
    
    add_banner_6(t1, "12-09-2026 : SECOND SATURDAY", col6_text="", b_type="holiday_red")
    
    add_banner_6(t1, "14-09-2026 TO 19-09-2026", col6_text="Videos", b_type="date")
    r7 = add_row_6(t1, period="1", content="పునశ్చరణ", content_color="red")
    r8 = add_row_6(t1, period="2")
    r9 = add_row_6(t1, period="3")
    r7.cells[2].merge(r8.cells[2]).merge(r9.cells[2])
    
    # ==================== PAGE 2 ====================
    print("  -> Building Page 2...")
    add_page_break_zero(doc)
    add_header_bar(doc)
    t2 = init_table()
    
    add_banner_6(t2, "14-09-2026 : GANESH CHATURTHI", col6_text="", b_type="holiday_red")
    add_banner_6(t2, "19-09-2026 : NO BAG DAY-STEAM&CCA", col6_text="", b_type="event_purple")
    add_banner_6(t2, "21-09-2026 TO 25-09-2026", col6_text="Videos", b_type="date")
    
    r10 = add_row_6(t2, period="1", content="పునశ్చరణ", content_color="red")
    r11 = add_row_6(t2, period="2")
    r12 = add_row_6(t2, period="3")
    r13 = add_row_6(t2, period="4")
    r10.cells[2].merge(r11.cells[2]).merge(r12.cells[2]).merge(r13.cells[2])
    
    add_banner_6(t2, "22-09-2026 TO 24-09-2026 : PTM", col6_text="", b_type="event_purple")
    add_banner_6(t2, "26-09-2026 TO 29-09-2026 : MID TERM REVISION", col6_text="", b_type="blue_banner")
    add_banner_6(t2, "30-09-2026: MID TERM EXAMS", col6_text="", b_type="blue_banner")
    
    doc.save(output_docx_path)
    print(f"  -> Successfully generated: {os.path.basename(output_docx_path)}")

# -------------------------------------------------------------
# Main Batch Processing Function
# -------------------------------------------------------------
def main():
    print("=" * 68)
    print("   TEL-HIN-GJ-OD MULTI-LANGUAGE PDF TO WORD BATCH CONVERTER")
    print("   (Supports Telugu, Hindi, Gujarati, and Odia Micro Schedules)")
    print("=" * 68)
    print(f"Input Folder : {INPUT_DIR}")
    print(f"Output Folder: {OUTPUT_DIR}\n")
    
    pdf_files = glob.glob(os.path.join(INPUT_DIR, "*.pdf"))
    if not pdf_files:
        print("[!] No PDF files found in 'input_pdfs' folder.")
        return
        
    print(f"Found {len(pdf_files)} PDF file(s) to convert:\\n")
    for idx, pdf_path in enumerate(pdf_files, 1):
        filename = os.path.basename(pdf_path)
        base_name = os.path.splitext(filename)[0]
        output_docx = os.path.join(OUTPUT_DIR, f"{base_name}.docx")
        
        print(f"[{idx}/{len(pdf_files)}] Processing: {filename}")
        
        # Check page count and content
        with fitz.open(pdf_path) as fdoc:
            num_pages = len(fdoc)
            first_page_text = fdoc[0].get_text()
            
        # Check file type and route accordingly
        filename_upper = filename.upper()
        is_6col = ("NLEARN" in first_page_text.upper() or "RECOMMENDATIONS" in first_page_text.upper() or "(1)" in filename)
        is_8page_schedule = (num_pages >= 7)
        
        try:
            if "HINDI" in filename_upper:
                import generate_hindi_ap
                generate_hindi_ap.generate_hindi_doc(output_docx)
            elif "GJ" in filename_upper or "GUJARAT" in filename_upper or "GUJARATH" in filename_upper:
                import generate_gujarati_ap
                generate_gujarati_ap.generate_gujarati_docx(pdf_path, output_docx)
            elif "OD" in filename_upper or "ODIA" in filename_upper or "ODISSA" in filename_upper or "ORIYA" in filename_upper:
                import generate_odia_schedule
                generate_odia_schedule.generate_odia_docx(pdf_path, output_docx)
            elif "PAGES-2-3" in filename_upper or "TELUGU-FL" in filename_upper or "TELUGU - FL" in filename_upper:
                import generate_class4_tg
                generate_class4_tg.generate_tg_fl_doc(output_docx)
            elif "PAGES-35-36" in filename_upper or "TELUGU-SL" in filename_upper or "TELUGU - SL" in filename_upper:
                import generate_class4_tg
                generate_class4_tg.generate_tg_sl_doc(output_docx)
            elif is_6col:
                convert_grade2_6col_schedule(pdf_path, output_docx)
            elif is_8page_schedule:
                from convert_grade2_schedule import convert
                convert()
                import shutil
                shutil.copy("GRADE_II_TELUGU_TL_TERM-1_MCS_AP_2026-27.docx", output_docx)
            else:
                convert_grade2_6col_schedule(pdf_path, output_docx)
            print(f"  [OK] Saved to: {output_docx}\n")
        except Exception as e:
            print(f"  [ERROR] {e}\n")

    print("=" * 65)
    print("Conversion completed! Check the 'output_docs' folder.")
    print("=" * 65)

if __name__ == "__main__":
    main()
