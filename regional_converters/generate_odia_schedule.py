import os
import sys
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import pymupdf as fitz

sys.stdout.reconfigure(encoding='utf-8')

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=50, bottom=50, left=80, right=80):
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

def format_text_run(run, text, font_name="Nirmala UI", size_pt=9.5, bold=False, color_rgb=None, italic=False):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Parse Akruti Converter HTML to get mappings
html_path = os.path.join(BASE_DIR, "assets", "akruti_converter.html")
if not os.path.exists(html_path):
    html_path = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\assets\akruti_converter.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'var text_array = new Array\((.*?)\n\)', html, re.DOTALL)
lines = m.group(1).splitlines()
items = []
for line in lines:
    line = line.strip()
    if not line or line.startswith('//'): continue
    if '//' in line: line = line[:line.index('//')]
    matches = re.findall(r'"([^"]*)"', line)
    items.extend(matches)

pairs = []
for i in range(0, len(items)-1, 2):
    k, v = items[i], items[i+1]
    if k == '\\\\': k = '\\'
    if k == ' ': continue
    pairs.append((k, v))

def convert_odia_text(spans):
    combined = ''
    for font, text in spans:
        if 'Akruti' in font:
            for k, v in pairs:
                if k: text = text.replace(k, v)
        combined += text
    
    consonants = 'କଖଗଘଙଚଛଜଝଞଟଠଡଡ଼ଢଢ଼ଣତଥଦଧନପଫବଭମଯୟରଲବୱଶଷସହକ୍ଷଳ'
    
    # matra e reordering
    combined = re.sub(r'([ù])([' + consonants + '])', r'\2\1', combined)
    combined = re.sub(r'([ù])([୍])([' + consonants + '])', r'\2\3\1', combined)
    combined = re.sub(r'([ù])([୍])([' + consonants + '])', r'\2\3\1', combined)
    combined = combined.replace('ùø', 'ୌ')
    combined = combined.replace('ùା', 'ୋ')
    combined = combined.replace('ù÷', 'ୈ')
    combined = combined.replace('ù', 'େ')
    
    # reph reordering
    pattern = r'(([' + consonants + '](?:୍[' + consonants + '])*)[ାିୀୁୂୃେୈୋୌଂଁ]*)ð'
    while 'ð' in combined:
        new_text = re.sub(pattern, r'ର୍\1', combined)
        if new_text == combined:
            combined = combined.replace('ð', 'ର୍')
            break
        combined = new_text

    pattern_a = r'(([' + consonants + '](?:୍[' + consonants + '])*)[ାିୀୁୂୃେୈୋୌଂଁ]*)à'
    while 'à' in combined:
        new_text = re.sub(pattern_a, r'ର୍\1', combined)
        if new_text == combined:
            combined = combined.replace('à', 'ର୍')
            break
        combined = new_text

    combined = re.sub(r'([ଂଁ])([ାିୀୁୂୃେୈୋୌ])', r'\2\1', combined)
    combined = re.sub(r'[ \t]+', ' ', combined)
    return combined.strip()

def extract_odia_page_rows(p):
    v_lines, h_lines = [], []
    for d in p.get_drawings():
        for item in d['items']:
            if item[0] == 'l':
                if abs(item[1].x - item[2].x) < 1: v_lines.append((item[1].x, min(item[1].y, item[2].y), max(item[1].y, item[2].y)))
                if abs(item[1].y - item[2].y) < 1: h_lines.append((item[1].y, min(item[1].x, item[2].x), max(item[1].x, item[2].x)))
            elif item[0] == 're':
                if item[1].width < 2: v_lines.append((item[1].x0, item[1].y0, item[1].y1))
                if item[1].height < 2: h_lines.append((item[1].y0, item[1].x0, item[1].x1))

    xs = sorted(list(set([round(x[0], 1) for x in v_lines if x[2]-x[1] > 10])))
    ys = sorted(list(set([round(y[0], 1) for y in h_lines if y[2]-y[1] > 200])))

    if len(xs) < 6:
        xs = [44.4, 107.2, 298.6, 518.3, 660.2, 798.3]

    rows = []
    for r_idx in range(len(ys)-1):
        y0, y1 = ys[r_idx], ys[r_idx+1]
        
        v_in_row = [x for x, vy0, vy1 in v_lines if vy0 < y1 - 3 and vy1 > y0 + 3]
        is_merged = (len(set(v_in_row)) <= 2)
        
        row_rect = fitz.Rect(xs[0]-2, y0-1, xs[-1]+2, y1+1)
        
        if is_merged:
            spans = []
            for b in p.get_text('dict', clip=row_rect)['blocks']:
                if 'lines' in b:
                    for l in b['lines']:
                        for s in l['spans']:
                            spans.append((s['font'], s['text']))
            rows.append(('merged', convert_odia_text(spans)))
        else:
            cols = []
            for c_idx in range(len(xs)-1):
                cx0, cx1 = xs[c_idx], xs[c_idx+1]
                cell_rect = fitz.Rect(cx0 + 1, y0, cx1 - 1, y1)
                spans = []
                for b in p.get_text('dict', clip=cell_rect)['blocks']:
                    if 'lines' in b:
                        for l in b['lines']:
                            for s in l['spans']:
                                spans.append((s['font'], s['text']))
                cols.append(convert_odia_text(spans))
            rows.append(('cols', cols))
    return rows

def generate_odia_docx(pdf_path, output_docx_path):
    print(f"Loading Odia PDF: {pdf_path}")
    doc_pdf = fitz.open(pdf_path)
    
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Inches(11.69)   # A4 landscape
    sec.page_height = Inches(8.27)
    sec.left_margin = Inches(0.5)
    sec.right_margin = Inches(0.5)
    sec.top_margin = Inches(0.4)
    sec.bottom_margin = Inches(0.4)
    
    # -------------------------------------------------------------
    # PAGE 1: COVER PAGE
    # -------------------------------------------------------------
    print("Building Page 1: Cover Page...")
    
    # State Pill
    tbl_state = doc.add_table(rows=1, cols=1)
    tbl_state.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_st = tbl_state.rows[0].cells[0]
    c_st.width = Inches(3.0)
    set_cell_background(c_st, "065F46") # Emerald green
    set_cell_margins(c_st, top=80, bottom=80, left=150, right=150)
    p_st = c_st.paragraphs[0]
    p_st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_st = p_st.add_run()
    format_text_run(r_st, "State: ODISHA", font_name="Calibri", size_pt=14, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(4)
    p_gap.paragraph_format.space_after = Pt(4)
    
    # Birsa Munda Portrait & Bio Layout
    tbl_bio = doc.add_table(rows=1, cols=2)
    tbl_bio.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_bio.autofit = False
    set_table_borders(tbl_bio, color="CBD5E1", sz="6", val="single")
    
    c_img = tbl_bio.rows[0].cells[0]
    c_img.width = Inches(3.2)
    set_cell_background(c_img, "F0FDF4")
    set_cell_margins(c_img, top=100, bottom=100, left=100, right=100)
    c_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_img = c_img.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    portrait_path = os.path.join(BASE_DIR, "assets", "od_assets", "img_2.jpeg")
    if not os.path.exists(portrait_path):
        portrait_path = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\assets\od_assets\img_2.jpeg'
    if os.path.exists(portrait_path):
        r_pic = p_img.add_run()
        r_pic.add_picture(portrait_path, width=Inches(2.2))
    
    p_auth = c_img.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth.paragraph_format.space_before = Pt(6)
    r_auth = p_auth.add_run()
    format_text_run(r_auth, "Birsa Munda\n", font_name="Calibri", size_pt=14, bold=True, color_rgb=RGBColor(20, 83, 45))
    r_alias = p_auth.add_run()
    format_text_run(r_alias, '(1875 – 1900) • Freedom Fighter & Folk Hero', font_name="Calibri", size_pt=10, italic=True, color_rgb=RGBColor(71, 85, 105))
    
    c_txt = tbl_bio.rows[0].cells[1]
    c_txt.width = Inches(7.4)
    set_cell_background(c_txt, "FFFFFF")
    set_cell_margins(c_txt, top=120, bottom=120, left=140, right=140)
    c_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    p_title = c_txt.paragraphs[0]
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(10)
    r_t1 = p_title.add_run()
    format_text_run(r_t1, "ODIA (ORIYA) — GRADE II\n", font_name="Arial", size_pt=18, bold=True, color_rgb=RGBColor(6, 95, 70))
    r_t2 = p_title.add_run()
    format_text_run(r_t2, "MICRO SCHEDULE (AY 2026-27)", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(217, 119, 6))
    
    p_bio = c_txt.add_paragraph()
    p_bio.paragraph_format.space_before = Pt(6)
    p_bio.paragraph_format.space_after = Pt(8)
    p_bio.paragraph_format.line_spacing = Pt(15)
    r_bio1 = p_bio.add_run()
    format_text_run(r_bio1, 'Birsa Munda belonged to the Munda tribe. He was a tribal freedom fighter and a religious leader from Odisha in India. He was also a popular folk hero. He was a significant figure in the Indian independence movement. His slogans are still remembered in many states of India like Odisha, West Bengal, Bihar, Jharkhand, and Madhya Pradesh.\n\n', font_name="Calibri", size_pt=11.5, color_rgb=RGBColor(51, 65, 85))
    r_bio2 = p_bio.add_run()
    format_text_run(r_bio2, 'He stated that the era of Queen Victoria was over, and it was the time of Munda Raj. He worked for them on behalf of his people and raised voices to no more pay rent. He and his people were against Britishers. They attacked the places of Britishers for over two years.', font_name="Calibri", size_pt=11.5, color_rgb=RGBColor(51, 65, 85))
    
    p_foot1 = doc.add_paragraph()
    p_foot1.paragraph_format.space_before = Pt(20)
    p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_foot1.add_run()
    format_text_run(r_f1, "ODIA  •  PAGE 26", font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(148, 163, 184))
    
    # -------------------------------------------------------------
    # PAGE 2: MONTHLY PAGE BREAKDOWN
    # -------------------------------------------------------------
    print("Building Page 2: Monthly Breakdown...")
    doc.add_page_break()
    
    p_m_t = doc.add_paragraph()
    p_m_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_m_t.paragraph_format.space_before = Pt(16)
    p_m_t.paragraph_format.space_after = Pt(16)
    r_mt = p_m_t.add_run()
    format_text_run(r_mt, "ODIA — GRADE II : MONTH-WISE SYLLABUS BREAKDOWN", font_name="Arial", size_pt=15, bold=True, color_rgb=RGBColor(6, 95, 70))
    
    tbl_months = doc.add_table(rows=1, cols=5)
    tbl_months.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_months.autofit = False
    set_table_borders(tbl_months, color="CBD5E1", sz="6", val="single")
    
    months_data = [
        ("APRIL", "pgno 5 to 13", "FEF2F2", RGBColor(185, 28, 28)),
        ("JUNE", "pgno 14 to 20", "EFF6FF", RGBColor(29, 78, 216)),
        ("JULY", "pgno 21 to 30", "FEF9C3", RGBColor(133, 77, 14)),
        ("AUGUST", "pg no-31 to 38", "F0FDF4", RGBColor(20, 83, 45)),
        ("SEPTEMBER", "pg no-39 to 41\n(Revision)", "FAF5FF", RGBColor(126, 34, 206)),
    ]
    
    m_widths = [Inches(2.1), Inches(2.1), Inches(2.1), Inches(2.1), Inches(2.2)]
    for idx, (m_name, m_pages, bg_col, text_col) in enumerate(months_data):
        c = tbl_months.rows[0].cells[idx]
        c.width = m_widths[idx]
        set_cell_background(c, bg_col)
        set_cell_margins(c, top=120, bottom=120, left=80, right=80)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r_m = p.add_run(f"{m_name}\n\n")
        format_text_run(r_m, f"{m_name}\n\n", font_name="Arial", size_pt=13, bold=True, color_rgb=text_col)
        r_p = p.add_run(m_pages)
        format_text_run(r_p, m_pages, font_name="Bookman Old Style", size_pt=11, bold=True, color_rgb=RGBColor(30, 41, 59))
        
    p_foot2 = doc.add_paragraph()
    p_foot2.paragraph_format.space_before = Pt(80)
    p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f2 = p_foot2.add_run()
    format_text_run(r_f2, "ODIA  •  PAGE 27", font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(148, 163, 184))

    # -------------------------------------------------------------
    # PAGES 3 TO 13: MICRO SCHEDULE TABLES
    # -------------------------------------------------------------
    doc.add_page_break()
    
    # Header bar
    tbl_hdr = doc.add_table(rows=1, cols=3)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr.autofit = False
    col_w_hdr = [Inches(3.2), Inches(4.3), Inches(3.2)]
    for i, w in enumerate(col_w_hdr):
        tbl_hdr.rows[0].cells[i].width = w
    
    c0 = tbl_hdr.rows[0].cells[0]
    set_cell_background(c0, "065F46")
    set_cell_margins(c0, top=50, bottom=50, left=140, right=140)
    c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run()
    format_text_run(r0, "Micro schedule", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    c1 = tbl_hdr.rows[0].cells[1]
    set_cell_background(c1, "D97706")
    set_cell_margins(c1, top=50, bottom=50, left=140, right=140)
    c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run()
    format_text_run(r1, "ODIA — Grade II (AY 2026-27)", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    c2 = tbl_hdr.rows[0].cells[2]
    set_cell_background(c2, "065F46")
    set_cell_margins(c2, top=50, bottom=50, left=140, right=140)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run()
    format_text_run(r2, "State: ODISHA", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(2)
    p_gap.paragraph_format.space_after = Pt(2)
    p_gap.paragraph_format.line_spacing = Pt(1)
    
    print("Extracting and adding schedule rows from Pages 3 to 13...")
    
    # Create the master table
    col_widths = [Inches(1.0), Inches(2.7), Inches(3.6), Inches(1.7), Inches(1.7)]
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="94A3B8", sz="4", val="single")
    
    # Add Table Header
    hdr_row = table.add_row()
    set_row_cant_split(hdr_row)
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    hdr_titles = ['PERIOD', 'CHAPTER/LESSON', 'CONTENT TO BE TAUGHT', 'TLM', 'PRACTICE WORK']
    for c_idx, title in enumerate(hdr_titles):
        c = hdr_row.cells[c_idx]
        c.width = col_widths[c_idx]
        set_cell_background(c, "E2E8F0")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run()
        format_text_run(r, title, font_name="Bookman Old Style", size_pt=10, bold=True, color_rgb=RGBColor(15, 23, 42))
    
    total_added_rows = 0
    for p_no in range(2, 13):
        p = doc_pdf[p_no]
        rows = extract_odia_page_rows(p)
        print(f"  Page {p_no+1}: {len(rows)} rows")
        
        for r_type, r_data in rows:
            # Skip duplicate header on page 3
            if p_no == 2 and r_type == 'cols' and r_data[0] == 'PERIOD':
                continue
            
            row = table.add_row()
            set_row_cant_split(row)
            
            if r_type == 'merged':
                c_start = row.cells[0]
                for i in range(1, 5):
                    c_start.merge(row.cells[i])
                c_start.width = sum(col_widths, Inches(0))
                
                txt = r_data
                is_holiday = any(h in txt for h in ['SATURDAY', 'HOLIDAY', 'VACATION', 'FITR', 'JAYANTI', 'DAY', 'FESTIVAL', 'NUAKHAI', 'CHATURTHI', 'BANDHAN', 'NABI', 'RAM NAVAMI'])
                is_exam = any(e in txt for e in ['EXAM', 'REVISION', 'CHECK', 'ASSESSMENT', 'TERM-END', 'MID TERM'])
                
                if is_holiday:
                    bg_color = "FEF2F2"
                    text_color = RGBColor(185, 28, 28)
                elif is_exam:
                    bg_color = "EFF6FF"
                    text_color = RGBColor(29, 78, 216)
                else:
                    bg_color = "FEF9C3"
                    text_color = RGBColor(133, 77, 14)
                
                set_cell_background(c_start, bg_color)
                set_cell_margins(c_start, top=45, bottom=45, left=100, right=100)
                c_start.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = c_start.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run()
                format_text_run(r, txt, font_name="Bookman Old Style", size_pt=9.5, bold=True, color_rgb=text_color)
            else:
                for c_idx in range(5):
                    c = row.cells[c_idx]
                    c.width = col_widths[c_idx]
                    cell_text = r_data[c_idx] if c_idx < len(r_data) else ""
                    
                    set_cell_background(c, "FFFFFF")
                    set_cell_margins(c, top=45, bottom=45, left=70, right=70)
                    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = c.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    
                    if c_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Bookman Old Style", size_pt=10, bold=True, color_rgb=RGBColor(30, 41, 59))
                    elif c_idx == 1:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p.add_run()
                        is_bold = any(kw in cell_text for kw in ['DAY-', 'FUNDAMENTAL', 'ପାଠ', 'Slip', 'Test', 'Revision'])
                        format_text_run(r, cell_text, font_name="Nirmala UI", size_pt=9.5, bold=is_bold, color_rgb=RGBColor(15, 23, 42))
                    elif c_idx == 2:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Nirmala UI", size_pt=9.5, bold=False, color_rgb=RGBColor(30, 41, 59))
                    elif c_idx == 3:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Nirmala UI", size_pt=9.0, bold=False, color_rgb=RGBColor(51, 65, 85))
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Nirmala UI", size_pt=9.0, bold=False, color_rgb=RGBColor(51, 65, 85))
            
            total_added_rows += 1
            
    print(f"Total Odia schedule rows added: {total_added_rows}")
    
    # -------------------------------------------------------------
    # PAGE 14: MID-TERM EXAMS & DUSSERAH HOLIDAYS
    # -------------------------------------------------------------
    print("Building Page 14: Mid-Term & Dusserah...")
    doc.add_page_break()
    
    p_p14_t = doc.add_paragraph()
    p_p14_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p14_t.paragraph_format.space_before = Pt(30)
    p_p14_t.paragraph_format.space_after = Pt(20)
    r_p14_t = p_p14_t.add_run()
    format_text_run(r_p14_t, "ODIA — GRADE II (EXAM & VACATION SCHEDULE)", font_name="Arial", size_pt=15, bold=True, color_rgb=RGBColor(6, 95, 70))
    
    tbl_p14 = doc.add_table(rows=0, cols=1)
    tbl_p14.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p14.autofit = False
    set_table_borders(tbl_p14, color="CBD5E1", sz="6", val="single")
    
    p14_items = [
        ("07-10-2026 TO 16-10-2026 : MID TERM EXAMS", "EFF6FF", RGBColor(29, 78, 216)),
        ("17-10-2026 TO 26-10-2026 : DUSSERAH", "FEF2F2", RGBColor(185, 28, 28)),
    ]
    
    for item_txt, bg, col in p14_items:
        r = tbl_p14.add_row()
        set_row_cant_split(r)
        c = r.cells[0]
        c.width = Inches(7.5)
        set_cell_background(c, bg)
        set_cell_margins(c, top=80, bottom=80, left=140, right=140)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        format_text_run(run, item_txt, font_name="Bookman Old Style", size_pt=12, bold=True, color_rgb=col)
    
    p_foot14 = doc.add_paragraph()
    p_foot14.paragraph_format.space_before = Pt(60)
    p_foot14.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f14 = p_foot14.add_run()
    format_text_run(r_f14, "PAGE 39", font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(148, 163, 184))
    
    out_dir = os.path.dirname(output_docx_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_docx_path)
    print(f"Successfully created: {output_docx_path}")

if __name__ == '__main__':
    pdf = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\input_pdfs\GRADE-2_FRONT PAGE_OD_2026-27.pdf'
    out = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\output_docs\GRADE-2_FRONT PAGE_OD_2026-27.docx'
    generate_odia_docx(pdf, out)
