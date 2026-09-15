import os
import sys
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
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

def format_text_run(run, text, font_name="Shruti", size_pt=9.5, bold=False, color_rgb=None, italic=False):
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

guj_corrections = [
    ('પાઠ્યપુસ્ર્ક', 'પાઠ્યપુસ્તક'),
    ('પુસ્ર્ક', 'પુસ્તક'),
    ('પૃષ્ઠ્', 'પૃષ્ઠ'),
    ('પાઠ્', 'પાઠ'),
    ('પઠ્ન', 'પઠન'),
    ('ર્થા', 'તથા'),
    ('શબ્દાથત', 'શબ્દાર્થ'),
    ('સમજૂર્ી', 'સમજૂતી'),
    ('કાર્ત', 'કાર્ડ'),
    ('ચચાત', 'ચર્ચા'),
    ('વકતશીટ', 'વર્કશીટ'),
    ('વગત', 'વર્ગ'),
    ('વિતન', 'વર્તન'),
    ('દીર્ત', 'દીર્ઘ'),
    ('બાળગીર્', 'બાળગીત'),
    ('કવવર્ા', 'કવિતા'),
    ('સંર્ાકૂકર્ી', 'સંતાકૂકડી'),
    ('પ્રાર્તના', 'પ્રાર્થના'),
    ('દદશાઓ', 'દિશાઓ'),
    ('વવવવધ', 'વિવિધ'),
    ('વશક્ષણ', 'શિક્ષણ'),
    ('સ્વરભચહ્નો', 'સ્વરચિહ્નો'),
    ('માદિર્', 'માહિતગાર'),
    ('પદાર્ો', 'પદાર્થો'),
    ('હ્સસ્વ', 'હ્રસ્વ'),
    ('હ્સવ', 'હ્રસ્વ'),
    ('પશ્ર્નોત્તર', 'પ્રશ્નોત્તર'),
    ('અનસુ વ્ ાર', 'અનુસ્વાર'),
    ('મુખય્', 'મુખ્ય'),
    ('અન ે', 'અને'),
    ('તણ્ર', 'ત્રણ'),
    ('ઋતઓુ', 'ઋતુઓ'),
]

def clean_gujarati_text(text):
    if not text:
        return ""
    res = text
    for k, v in guj_corrections:
        res = res.replace(k, v)
    # clean extra spaces around colons/commas
    res = re.sub(r'\s+', ' ', res)
    return res.strip()

def extract_page_table_data(p):
    v_lines, h_lines = [], []
    for d in p.get_drawings():
        for item in d['items']:
            if item[0] == 'l':
                if abs(item[1].x - item[2].x) < 1: v_lines.append((item[1].x, min(item[1].y, item[2].y), max(item[1].y, item[2].y)))
                if abs(item[1].y - item[2].y) < 1: h_lines.append((item[1].y, min(item[1].x, item[2].x), max(item[1].x, item[2].x)))
            elif item[0] == 're':
                if item[1].width < 2: v_lines.append((item[1].x0, item[1].y0, item[1].y1))
                if item[1].height < 2: h_lines.append((item[1].y0, item[1].x0, item[1].x1))
    
    ys = sorted(list(set([round(y[0], 1) for y in h_lines if y[2]-y[1] > 200])))
    xs = sorted(list(set([round(x[0], 1) for x in v_lines if x[2]-x[1] > 10])))
    
    rows = []
    for r_idx in range(len(ys)-1):
        y0, y1 = ys[r_idx], ys[r_idx+1]
        
        # Check internal vertical lines
        v_in_row = [x for x, vy0, vy1 in v_lines if vy0 < y1 - 3 and vy1 > y0 + 3]
        unique_v = sorted(list(set(v_in_row)))
        is_merged = (len(unique_v) <= 2)
        
        row_rect = fitz.Rect(xs[0]-2, y0-1, xs[-1]+2, y1+1)
        words = p.get_text('words', clip=row_rect)
        
        if is_merged:
            words_sorted = sorted(words, key=lambda w: (round(w[1]/4)*4, w[0]))
            text = ' '.join([w[4] for w in words_sorted])
            rows.append(('merged', clean_gujarati_text(text)))
        else:
            cols = []
            for c_idx in range(len(xs)-1):
                cx0, cx1 = xs[c_idx], xs[c_idx+1]
                c_words = [w for w in words if cx0 - 2 <= (w[0]+w[2])/2 < cx1 + 2]
                c_words_sorted = sorted(c_words, key=lambda w: (round(w[1]/4)*4, w[0]))
                col_text = ' '.join([w[4] for w in c_words_sorted])
                cols.append(clean_gujarati_text(col_text))
            rows.append(('cols', cols))
    return rows

def generate_gujarati_docx(pdf_path, output_docx_path):
    print(f"Loading Gujarati PDF: {pdf_path}")
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
    set_cell_background(c_st, "1E3A8A")
    set_cell_margins(c_st, top=80, bottom=80, left=150, right=150)
    p_st = c_st.paragraphs[0]
    p_st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_st = p_st.add_run()
    format_text_run(r_st, "State: GUJARAT", font_name="Calibri", size_pt=14, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(4)
    p_gap.paragraph_format.space_after = Pt(4)
    
    # Author portrait & Bio Layout
    tbl_bio = doc.add_table(rows=1, cols=2)
    tbl_bio.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_bio.autofit = False
    set_table_borders(tbl_bio, color="CBD5E1", sz="6", val="single")
    
    c_img = tbl_bio.rows[0].cells[0]
    c_img.width = Inches(3.2)
    set_cell_background(c_img, "F8FAFC")
    set_cell_margins(c_img, top=100, bottom=100, left=100, right=100)
    c_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_img = c_img.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    portrait_path = os.path.join(BASE_DIR, "assets", "gj_assets", "img_2.jpeg")
    if not os.path.exists(portrait_path):
        portrait_path = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\assets\gj_assets\img_2.jpeg'
    if os.path.exists(portrait_path):
        r_pic = p_img.add_run()
        r_pic.add_picture(portrait_path, width=Inches(2.4))
    
    p_auth = c_img.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth.paragraph_format.space_before = Pt(6)
    r_auth = p_auth.add_run()
    format_text_run(r_auth, "Gaurishankar Joshi\n", font_name="Calibri", size_pt=14, bold=True, color_rgb=RGBColor(30, 41, 59))
    r_alias = p_auth.add_run()
    format_text_run(r_alias, '(1892 – 1965) • Pen Name: "Dhumketu"', font_name="Calibri", size_pt=10, italic=True, color_rgb=RGBColor(71, 85, 105))
    
    c_txt = tbl_bio.rows[0].cells[1]
    c_txt.width = Inches(7.4)
    set_cell_background(c_txt, "FFFFFF")
    set_cell_margins(c_txt, top=120, bottom=120, left=140, right=140)
    c_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    p_title = c_txt.paragraphs[0]
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(10)
    r_t1 = p_title.add_run()
    format_text_run(r_t1, "GUJARATI — GRADE II\n", font_name="Arial", size_pt=18, bold=True, color_rgb=RGBColor(30, 58, 138))
    r_t2 = p_title.add_run()
    format_text_run(r_t2, "MICRO SCHEDULE (AY 2026-27)", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(217, 119, 6))
    
    p_bio = c_txt.add_paragraph()
    p_bio.paragraph_format.space_before = Pt(6)
    p_bio.paragraph_format.space_after = Pt(8)
    p_bio.paragraph_format.line_spacing = Pt(15)
    r_bio1 = p_bio.add_run()
    format_text_run(r_bio1, 'Gaurishankar Govardhanram Joshi, known as “Dhumketu,” was a famous Gujarati author and a pioneer of modern Gujarati short stories. Born in 1892 in Gujarat, he wrote numerous short stories and novels known for their emotional depth and focus on human relationships.\n\n', font_name="Calibri", size_pt=11.5, color_rgb=RGBColor(51, 65, 85))
    r_bio2 = p_bio.add_run()
    format_text_run(r_bio2, 'His well-known work “Post Office” reflects his simple yet powerful style. He was awarded the Narmad Suvarna Chandrak in 1949 and played a key role in shaping Gujarati literature.', font_name="Calibri", size_pt=11.5, color_rgb=RGBColor(51, 65, 85))
    
    p_foot1 = doc.add_paragraph()
    p_foot1.paragraph_format.space_before = Pt(20)
    p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_foot1.add_run()
    format_text_run(r_f1, "GUJARATI  •  PAGE 66", font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(148, 163, 184))
    
    # -------------------------------------------------------------
    # PAGES 2 TO 15: MICRO SCHEDULE TABLES
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
    set_cell_background(c0, "1E3A8A")
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
    format_text_run(r1, "GUJARATI — Grade II (AY 2026-27)", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    c2 = tbl_hdr.rows[0].cells[2]
    set_cell_background(c2, "1E3A8A")
    set_cell_margins(c2, top=50, bottom=50, left=140, right=140)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run()
    format_text_run(r2, "State: GUJARAT", font_name="Arial", size_pt=13, bold=True, color_rgb=RGBColor(255, 255, 255))
    
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(2)
    p_gap.paragraph_format.space_after = Pt(2)
    p_gap.paragraph_format.line_spacing = Pt(1)
    
    print("Extracting and adding schedule rows from Pages 2 to 15...")
    
    # Create the master table
    col_widths = [Inches(1.0), Inches(2.7), Inches(3.6), Inches(1.7), Inches(1.7)]
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="94A3B8", sz="4", val="single")
    
    # Add Table Header
    hdr_row = table.add_row()
    set_row_cant_split(hdr_row)
    # Set tblHeader so it repeats at page breaks
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
    for p_no in range(1, 15):
        p = doc_pdf[p_no]
        rows = extract_page_table_data(p)
        print(f"  Page {p_no+1}: {len(rows)} rows")
        
        for r_type, r_data in rows:
            # Skip duplicate header on page 2
            if p_no == 1 and r_type == 'cols' and r_data[0] == 'PERIOD':
                continue
            
            row = table.add_row()
            set_row_cant_split(row)
            
            if r_type == 'merged':
                # Merge across all 5 cells
                c_start = row.cells[0]
                for i in range(1, 5):
                    c_start.merge(row.cells[i])
                c_start.width = sum(col_widths, Inches(0))
                
                # Check if it's a date range vs holiday vs exam
                txt = r_data
                is_holiday = any(h in txt for h in ['SATURDAY', 'HOLIDAY', 'JAYANTI', 'DAY', 'FESTIVAL', 'EID', 'CHATURTHI', 'DIWALI'])
                is_exam = any(e in txt for e in ['EXAM', 'REVISION', 'CHECK', 'ASSESSMENT', 'TERM-END'])
                
                if is_holiday:
                    bg_color = "FEF2F2" # soft red/rose
                    text_color = RGBColor(185, 28, 28)
                elif is_exam:
                    bg_color = "EFF6FF" # soft blue
                    text_color = RGBColor(29, 78, 216)
                else:
                    bg_color = "FEF9C3" # soft yellow / date range
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
                        is_bold = any(kw in cell_text for kw in ['Chapter', 'Check', 'પુનરાવર્તન'])
                        format_text_run(r, cell_text, font_name="Shruti", size_pt=9.5, bold=is_bold, color_rgb=RGBColor(15, 23, 42))
                    elif c_idx == 2:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Shruti", size_pt=9.5, bold=False, color_rgb=RGBColor(30, 41, 59))
                    elif c_idx == 3:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Shruti", size_pt=9.0, bold=False, color_rgb=RGBColor(51, 65, 85))
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run()
                        format_text_run(r, cell_text, font_name="Shruti", size_pt=9.0, bold=False, color_rgb=RGBColor(51, 65, 85))
            
            total_added_rows += 1
            
    print(f"Total schedule rows added: {total_added_rows}")
    
    # -------------------------------------------------------------
    # PAGE 16: MARCH 2027 HOLIDAYS & TERM-END EXAMS
    # -------------------------------------------------------------
    print("Building Page 16: Term-End Holidays & Exams...")
    doc.add_page_break()
    
    p_p16_t = doc.add_paragraph()
    p_p16_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p16_t.paragraph_format.space_before = Pt(20)
    p_p16_t.paragraph_format.space_after = Pt(14)
    r_p16_t = p_p16_t.add_run()
    format_text_run(r_p16_t, "GUJARATI — GRADE II (TERM-END SCHEDULE)", font_name="Arial", size_pt=15, bold=True, color_rgb=RGBColor(30, 58, 138))
    
    tbl_p16 = doc.add_table(rows=0, cols=1)
    tbl_p16.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p16.autofit = False
    set_table_borders(tbl_p16, color="CBD5E1", sz="6", val="single")
    
    p16_items = [
        ("06-03-2027 : MAHA SHIVARATRI", "FEF2F2", RGBColor(185, 28, 28)),
        ("10-03-2027 : RAMADAN", "FEF2F2", RGBColor(185, 28, 28)),
        ("22-03-2027 : HOLI", "FEF2F2", RGBColor(185, 28, 28)),
        ("26-03-2027 : GOOD FRIDAY", "FEF2F2", RGBColor(185, 28, 28)),
        ("27-03-2027 : FOURTH SATURDAY", "FEF2F2", RGBColor(185, 28, 28)),
        ("08-03-2027 TO 11-03-2027 : TERM-END REVISION", "EFF6FF", RGBColor(29, 78, 216)),
        ("12-03-2027 TO 20-03-2027 : TERM-END EXAMS", "FEF3C7", RGBColor(180, 83, 9)),
        ("23-03-2027 TO 31-03-2027 : TERM END HOLIDAYS", "FEF2F2", RGBColor(185, 28, 28)),
    ]
    
    for item_txt, bg, col in p16_items:
        r = tbl_p16.add_row()
        set_row_cant_split(r)
        c = r.cells[0]
        c.width = Inches(7.5)
        set_cell_background(c, bg)
        set_cell_margins(c, top=70, bottom=70, left=140, right=140)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        format_text_run(run, item_txt, font_name="Bookman Old Style", size_pt=11, bold=True, color_rgb=col)
    
    p_foot16 = doc.add_paragraph()
    p_foot16.paragraph_format.space_before = Pt(40)
    p_foot16.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f16 = p_foot16.add_run()
    format_text_run(r_f16, "PAGE 81", font_name="Calibri", size_pt=10, bold=True, color_rgb=RGBColor(148, 163, 184))
    
    out_dir = os.path.dirname(output_docx_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_docx_path)
    print(f"Successfully created: {output_docx_path}")

if __name__ == '__main__':
    pdf = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\input_pdfs\GRADE-2_FRONT PAGE_GJ_2026-27.pdf'
    out = r'C:\Users\Deepak Pawar\Desktop\Automation Projects\Telugu_PDF_to_Word\output_docs\GRADE-2_FRONT PAGE_GJ_2026-27.docx'
    generate_gujarati_docx(pdf, out)
