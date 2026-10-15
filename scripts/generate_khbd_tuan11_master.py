# -*- coding: utf-8 -*-
"""
generate_khbd_tuan11_master.py
Tạo chuẩn hóa 5 KHBD môn Tin học Tuần 11 (PPCT 10) cho Tiểu học:
- Tiền tiểu học (TT3)
- Lớp 1 (1A1)
- Lớp 2 (2A1)
- Lớp 3 (3A1)
- Lớp 4 (4C1)

Tuân thủ nghiêm ngặt AGENTS.md, CV 5512 (PL4), Thông tư 02/2025 - CV 3456.
"""

import os, sys, copy
from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_PATH = r"D:\UNIGO\KHBD_Tin_học\Tiền_tiểu_học\Tuần_10\KHBD_Tin_hoc_Tiền_tiểu_học_Tiet09_Em_tao_buc_tranh_au_tien.docx"
OUT_BASE_DIR = r"D:\UNIGO\KHBD_Tin_học"

# Style constants
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(13)
LINE_SPACING = 1.35
SPACE_AFTER = Emu(38100) # 3pt

INDENT_0 = 0
INDENT_1 = 457200      # 1.27cm
INDENT_2 = 450215      # 1.25cm
INDENT_BULLET = 540000 # 1.5cm

def afont(run, bold=False, italic=False, size=FONT_SIZE, color=None):
    run.font.name = FONT_NAME
    run.font.size = size
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), FONT_NAME)
    rFonts.set(qn('w:hAnsi'), FONT_NAME)
    rFonts.set(qn('w:cs'), FONT_NAME)
    rFonts.set(qn('w:eastAsia'), FONT_NAME)

def add_p(doc, text="", bold=False, italic=False, first_indent=INDENT_0, left_indent=0,
          align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=SPACE_AFTER, space_before=0, line_spacing=LINE_SPACING):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if first_indent:
        pf.first_line_indent = Emu(first_indent)
    if left_indent:
        pf.left_indent = Emu(left_indent)
    pf.space_after = space_after
    pf.space_before = Emu(space_before)
    pf.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        afont(run, bold=bold, italic=italic)
    return p

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_no_borders(tbl):
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl._tbl.insert(0, tblPr)
    old_b = tblPr.find(qn('w:tblBorders'))
    if old_b is not None:
        tblPr.remove(old_b)
    borders_xml = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders_xml)
    for r in tbl._tbl.findall(qn('w:tr')):
        for c in r.findall(qn('w:tc')):
            tcPr = c.find(qn('w:tcPr'))
            if tcPr is not None:
                tcB = tcPr.find(qn('w:tcBorders'))
                if tcB is not None:
                    tcPr.remove(tcB)

def set_table_borders(tbl, color="000000", sz="4", val="single"):
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl._tbl.insert(0, tblPr)
    old_b = tblPr.find(qn('w:tblBorders'))
    if old_b is not None:
        tblPr.remove(old_b)
    borders_xml = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders_xml)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def create_table_0(doc, lop_name, ngay_day_str):
    """Bảng thông tin 3x2 NO BORDER đúng chuẩn Section IV.4"""
    tbl = doc.add_table(rows=3, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl)
    
    # Row 0
    c00 = tbl.rows[0].cells[0]
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r00 = p00.add_run("Trường: Tiểu học và THCS UNIGO")
    afont(r00, bold=True)
    
    c01 = tbl.rows[0].cells[1]
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r01 = p01.add_run("Ngày soạn: 10/10/2026")
    afont(r01, bold=True)
    
    # Row 1
    c10 = tbl.rows[1].cells[0]
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r10 = p10.add_run("GV: Đậu Đình Nguyên")
    afont(r10, bold=False)
    
    c11 = tbl.rows[1].cells[1]
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r11 = p11.add_run(f"Ngày dạy: {ngay_day_str}")
    afont(r11, bold=False)
    
    # Row 2
    c20 = tbl.rows[2].cells[0]
    p20 = c20.paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r20 = p20.add_run("Tổ: Tổ chuyên môn Tiểu học")
    afont(r20, bold=False)
    
    c21 = tbl.rows[2].cells[1]
    p21 = c21.paragraphs[0]
    p21.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r21 = p21.add_run(f"Lớp: {lop_name}")
    afont(r21, bold=False)
    
    for row in tbl.rows:
        for cell in row.cells:
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.width = Cm(8.0)
    
    # Thêm dòng trống cách
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = 0
    p_sp.paragraph_format.space_after = Emu(76200)

def create_signature_table(doc):
    """Phần Rút kinh nghiệm & Bảng chữ ký 3x3 NO BORDER (Section IV.5)"""
    p_rk = doc.add_paragraph()
    p_rk.paragraph_format.space_before = Emu(76200)
    p_rk.paragraph_format.space_after = Emu(38100)
    r_rk = p_rk.add_run("RÚT KINH NGHIỆM SAU BÀI DẠY:")
    afont(r_rk, bold=True)
    
    p_dot1 = doc.add_paragraph()
    p_dot1.paragraph_format.space_after = Emu(19050)
    r_dot1 = p_dot1.add_run("...........................................................................................................................")
    afont(r_dot1)
    
    p_dot2 = doc.add_paragraph()
    p_dot2.paragraph_format.space_after = Emu(114300)
    r_dot2 = p_dot2.add_run("...........................................................................................................................")
    afont(r_dot2)
    
    # Bảng 3x3
    tbl = doc.add_table(rows=3, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl)
    
    headers = ["DUYỆT CỦA BGH", "DUYỆT CỦA TỔ CM", "NGƯỜI SOẠN"]
    for col_i, text in enumerate(headers):
        cell = tbl.rows[0].cells[col_i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        afont(r, bold=True)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
    
    subtitles = ["(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"]
    for col_i, text in enumerate(subtitles):
        cell = tbl.rows[1].cells[col_i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        afont(r, italic=True)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        
    names = ["\n\n\n", "\n\n\n", "\n\n\nĐậu Đình Nguyên"]
    for col_i, text in enumerate(names):
        cell = tbl.rows[2].cells[col_i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        afont(r, bold=True if "Đậu" in text else False)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)


def build_khbd_doc(lesson_meta):
    doc = Document(TEMPLATE_PATH)
    
    # Giữ w:sectPr
    for child in list(doc.element.body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag != 'sectPr':
            doc.element.body.remove(child)
            
    # Bảo tồn header runs
    if doc.sections:
        sec = doc.sections[0]
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.0)
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
    
    # 1. Bảng 0: Thông tin đầu bài 3x2 NO BORDER
    create_table_0(doc, lesson_meta["lop_name"], lesson_meta["ngay_day"])
    
    # 2. Tiêu đề bài dạy
    p_mon = add_p(doc, "KẾ HOẠCH DẠY HỌC MÔN TIN HỌC", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Emu(38100))
    afont(p_mon.runs[0], size=Pt(14), bold=True)
    
    p_bai = add_p(doc, f"{lesson_meta['ten_bai'].upper()} (Tiết: {lesson_meta['tiet_ppct']} theo PPCT)",
                  bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Emu(114300))
    afont(p_bai.runs[0], size=Pt(13), bold=True)
    
    # 3. Phần I: YÊU CẦU CẦN ĐẠT
    add_p(doc, "I. YÊU CẦU CẦN ĐẠT:", bold=True, first_indent=INDENT_0, space_after=Emu(38100))
    add_p(doc, "- Sau bài học này em sẽ:", first_indent=INDENT_1, space_after=Emu(19050))
    for y in lesson_meta["yccd_list"]:
        add_p(doc, f"+ {y}", left_indent=INDENT_BULLET, space_after=Emu(19050))
        
    # 1. Phát triển phẩm chất (TRƯỚC)
    add_p(doc, "1. Phát triển phẩm chất", bold=True, first_indent=INDENT_1, space_after=Emu(38100), space_before=38100)
    for pc_name, pc_desc in lesson_meta["pham_chat"].items():
        p_pc = add_p(doc, f"- {pc_name}: ", bold=True, left_indent=INDENT_BULLET, space_after=Emu(19050))
        r_desc = p_pc.add_run(pc_desc)
        afont(r_desc, bold=False)
        
    # 2. Phát triển năng lực (SAU)
    add_p(doc, "2. Phát triển năng lực", bold=True, first_indent=INDENT_1, space_after=Emu(38100), space_before=38100)
    
    # 2.1. Năng lực đặc thù
    add_p(doc, "2.1. Năng lực đặc thù (Tin học):", bold=True, first_indent=INDENT_2, space_after=Emu(25400))
    for nl_code, nl_data in lesson_meta["nl_dac_thu"].items():
        p_nl = add_p(doc, f"- {nl_code} ({nl_data['ten']}): ", bold=True, left_indent=INDENT_BULLET, space_after=Emu(19050))
        r_nl = p_nl.add_run(nl_data["bieu_hien"])
        afont(r_nl, bold=False)
        
    # 2.2. Năng lực số (TT 02/2025 - CV 3456)
    add_p(doc, "2.2. Năng lực số (Thông tư 02/2025 – CV 3456):", bold=True, first_indent=INDENT_2, space_after=Emu(25400))
    for nls_item in lesson_meta["nang_luc_so"]:
        p_nls = add_p(doc, f"- {nls_item['mien']} (thành tố {nls_item['thanh_to']} – Bậc {nls_item['bac']}): ",
                      bold=True, left_indent=INDENT_BULLET, space_after=Emu(19050))
        r_nls = p_nls.add_run(nls_item["bieu_hien"])
        afont(r_nls, bold=False)
        
    # 2.3. Năng lực chung
    add_p(doc, "2.3. Năng lực chung:", bold=True, first_indent=INDENT_2, space_after=Emu(25400))
    for nlc_name, nlc_desc in lesson_meta["nl_chung"].items():
        p_nlc = add_p(doc, f"- {nlc_name}: ", bold=True, left_indent=INDENT_BULLET, space_after=Emu(19050))
        r_nlc = p_nlc.add_run(nlc_desc)
        afont(r_nlc, bold=False)
        
    # 4. Phần II: ĐỒ DÙNG DẠY HỌC
    add_p(doc, "II. ĐỒ DÙNG DẠY HỌC :", bold=True, first_indent=INDENT_0, space_after=Emu(38100), space_before=57150)
    add_p(doc, f"1. Giáo viên: {lesson_meta['do_dung_gv']}", first_indent=INDENT_1, space_after=Emu(19050))
    add_p(doc, f"2. Học sinh: {lesson_meta['do_dung_hs']}", first_indent=INDENT_1, space_after=Emu(38100))
    
    # 5. Phần III: PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC
    add_p(doc, "III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC", bold=True, first_indent=INDENT_0, space_after=Emu(38100), space_before=38100)
    add_p(doc, f"- Các phương pháp: {lesson_meta['phuong_phap']}", left_indent=INDENT_1, space_after=Emu(19050))
    add_p(doc, f"- Kĩ thuật dạy học: {lesson_meta['ki_thuat']}", left_indent=INDENT_1, space_after=Emu(38100))
    
    # 6. Phần IV: CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU (Bảng 2 cột chuẩn)
    add_p(doc, "IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU:", bold=True, first_indent=INDENT_0, space_after=Emu(57150), space_before=38100)
    
    tbl_hd = doc.add_table(rows=0, cols=2)
    tbl_hd.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_hd)
    
    # Header row
    hdr_row = tbl_hd.add_row()
    c_hdr0 = hdr_row.cells[0]
    p_h0 = c_hdr0.paragraphs[0]
    p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h0 = p_h0.add_run("HOẠT ĐỘNG CỦA GV – HS")
    afont(r_h0, bold=True)
    set_cell_background(c_hdr0, "F2F2F2")
    set_cell_margins(c_hdr0, top=120, bottom=120, left=150, right=150)
    c_hdr0.width = Cm(10.5)
    
    c_hdr1 = hdr_row.cells[1]
    p_h1 = c_hdr1.paragraphs[0]
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h1 = p_h1.add_run("KẾT QUẢ CẦN ĐẠT")
    afont(r_h1, bold=True)
    set_cell_background(c_hdr1, "F2F2F2")
    set_cell_margins(c_hdr1, top=120, bottom=120, left=150, right=150)
    c_hdr1.width = Cm(5.5)
    
    def add_activity_merged_row(title_text):
        row = tbl_hd.add_row()
        c0 = row.cells[0]
        c1 = row.cells[1]
        tcPr0 = c0._tc.get_or_add_tcPr()
        gs = parse_xml(r'<w:gridSpan %s w:val="2"/>' % nsdecls('w'))
        tcPr0.append(gs)
        row._tr.remove(c1._tc)
        set_cell_background(c0, "E9EEF4")
        set_cell_margins(c0, top=100, bottom=100, left=150, right=150)
        p = c0.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title_text)
        afont(r, bold=True)
        
    def add_activity_content_row(gv_hs_paras, kq_paras):
        row = tbl_hd.add_row()
        c0 = row.cells[0]
        c1 = row.cells[1]
        c0.width = Cm(10.5)
        c1.width = Cm(5.5)
        set_cell_margins(c0, top=100, bottom=100, left=120, right=120)
        set_cell_margins(c1, top=100, bottom=100, left=120, right=120)
        
        # Cột 1
        p0 = c0.paragraphs[0]
        for idx, (p_text, is_bold, is_italic) in enumerate(gv_hs_paras):
            cur_p = p0 if idx == 0 else c0.add_paragraph()
            cur_p.paragraph_format.line_spacing = 1.2
            cur_p.paragraph_format.space_after = Emu(25400)
            r = cur_p.add_run(p_text)
            afont(r, bold=is_bold, italic=is_italic, size=Pt(12.5))
            
        # Cột 2
        p1 = c1.paragraphs[0]
        for idx, (p_text, is_bold, is_italic) in enumerate(kq_paras):
            cur_p = p1 if idx == 0 else c1.add_paragraph()
            cur_p.paragraph_format.line_spacing = 1.2
            cur_p.paragraph_format.space_after = Emu(25400)
            r = cur_p.add_run(p_text)
            afont(r, bold=is_bold, italic=is_italic, size=Pt(12.5))
            
    # Thêm 4 hoạt động
    for hd in lesson_meta["hoat_dong"]:
        add_activity_merged_row(hd["tieu_de"])
        add_activity_content_row(hd["gv_hs"], hd["ket_qua"])
        
    # 7. Phần V: ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY
    add_p(doc, "V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY :", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_after=Emu(38100), space_before=76200)
    add_p(doc, "(GV ghi những nội dung mà mình đã bổ sung ngoài KHBD đã lên...)",
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Emu(38100))
    add_p(doc, "...........................................................................................................................",
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Emu(114300))
    
    # 8. Bảng chữ ký cuối bài
    create_signature_table(doc)
    
    # Save file
    out_dir = os.path.join(OUT_BASE_DIR, lesson_meta["folder_name"], "Tuần_11")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, lesson_meta["file_name"])
    doc.save(out_file)
    print(f"[OK] Generated KHBD: {out_file}")
    return out_file

# ============================================================
# LESSON DATA SPECIFICATIONS FOR ALL 5 CLASSES
# ============================================================
LESSONS_DATA = [
    # 1. TIỀN TIỂU HỌC
    {
        "folder_name": "Tiền_tiểu_học",
        "file_name": "KHBD_Tin_hoc_Tiền_tiểu_học_Tiet10_On_tap_Be_lam_chu_chuot.docx",
        "lop_name": "TT3",
        "ngay_day": "15/10/2026",
        "tiet_ppct": 10,
        "ten_bai": "BÀI 10. ÔN TẬP: BÉ LÀM CHỦ CHUỘT",
        "yccd_list": [
            "Cầm chuột đúng cách bằng tay phải, thả lỏng cổ tay.",
            "Thực hiện thành thạo thao tác di chuyển con trỏ chuột đến đúng vị trí mong muốn.",
            "Phân biệt và thực hiện đúng thao tác nháy nút chuột trái, nháy đúp chuột để mở trò chơi và kéo thả chuột để ghép tranh."
        ],
        "pham_chat": {
            "Chăm chỉ": "Tích cực, kiên trì luyện tập các bài tập di chuyển và kéo thả chuột. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Trách nhiệm": "Biết giữ gìn cẩn thận chuột và máy tính, không giật dây chuột, ngồi ngay ngắn. (Đạt được thông qua Hoạt động 1, Hoạt động 4)",
            "Trung thực": "Tự mình thao tác và hoàn thành các nhiệm vụ ghép hình trò chơi. (Đạt được thông qua Hoạt động 3)"
        },
        "nl_dac_thu": {
            "NLa": {
                "ten": "Sử dụng và quản lí các phương tiện ICT",
                "bieu_hien": "Nhận biết các bộ phận của chuột (nút trái, nút phải, bánh lăn) và cầm chuột đúng tư thế. (Đạt được thông qua Hoạt động 1, Hoạt động 2)"
            },
            "NLd": {
                "ten": "Ứng dụng ICT trong học và tự học",
                "bieu_hien": "Sử dụng chuột linh hoạt để tương tác với trò chơi học tập và tô màu ghép tranh. (Đạt được thông qua Hoạt động 3)"
            }
        },
        "nang_luc_so": [
            {
                "mien": "Miền I. Khai thác dữ liệu và thông tin",
                "thanh_to": "1.4. Xử lý dữ liệu",
                "bac": 1,
                "bieu_hien": "Thao tác chọn đúng đối tượng hình ảnh hiển thị trên màn hình bằng nút chuột trái theo yêu cầu. (Đạt được thông qua Hoạt động 2)"
            },
            {
                "mien": "Miền V. Giải quyết vấn đề",
                "thanh_to": "5.3. Sử dụng sáng tạo công nghệ",
                "bac": 1,
                "bieu_hien": "Kéo thả các mảnh ghép về đúng vị trí để hoàn thành bức tranh hoàn chỉnh trên máy tính. (Đạt được thông qua Hoạt động 3)"
            }
        ],
        "nl_chung": {
            "Tự chủ và tự học": "Tự tin cầm chuột điều khiển con trỏ theo ý muốn cá nhân mà không cần trợ giúp liên tục. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Giao tiếp và hợp tác": "Biết chia sẻ chuột và lượt chơi với bạn bên cạnh trong phòng máy. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "do_dung_gv": "Máy chiếu, máy tính GV, bài giảng điện tử tương tác, phần mềm trò chơi rèn chuột (Mousercise / Ghép tranh gấu con).",
        "do_dung_hs": "Máy tính có chuột rời, miếng lót chuột, bàn ghế vừa tầm vóc.",
        "phuong_phap": "Trực quan mẫu, trò chơi học tập, thực hành luyện tập cá nhân và theo cặp.",
        "ki_thuat": "Động não, làm mẫu từng bước, khích lệ tích cực.",
        "hoat_dong": [
            {
                "tieu_de": "1. Hoạt động MỞ ĐẦU (5 phút)\n*Mục tiêu: Tạo không khí vui tươi, kích hoạt lại tư thế cầm chuột và nhận diện các nút trên chuột máy tính.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV tổ chức trò chơi \"Bàn tay thỏ trắng\":\n  + Đố các con: Chú chuột máy tính có mấy tai nhỉ?\n  + Tai trái là nút nào? Tai phải là nút nào?", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS làm động tác đặt bàn tay lên lưng chú chuột: ngón trỏ đặt lên nút chuột trái, ngón giữa đặt lên nút chuột phải.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV gọi 2 bạn lên trước lớp làm mẫu lại cách cầm chuột.\n- Cả lớp vỗ tay khen ngợi.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chốt: Hôm nay chúng mình sẽ cùng thi xem ai là \"Dũng sĩ điều khiển chuột tài ba\" nhé!", False, False)
                ],
                "ket_qua": [
                    ("- HS hào hứng, nhớ lại cấu tạo chuột:\n  + Nút trái, nút phải, bánh lăn.\n- HS cầm chuột đúng cách bằng tay phải.", False, False)
                ]
            },
            {
                "tieu_de": "2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI / ÔN LUYỆN KỸ NĂNG (15 phút)\n*Mục tiêu: Củng cố 4 kỹ năng dùng chuột: Di chuyển con trỏ, Nháy chuột trái, Nháy đúp chuột, Kéo thả chuột.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV trình chiếu các bức tranh tương ứng với 4 tuyệt chiêu của chuột:\n  + Thao tác 1: Rê chuột (di chuyển con trỏ nhẹ nhàng trên mặt bàn).\n  + Thao tác 2: Nháy chuột trái (nghe tiếng \"tách\" dứt khoát).\n  + Thao tác 3: Nháy đúp (gõ ngón trỏ 2 lần thật nhanh để mở phần mềm).\n  + Thao tác 4: Kéo thả (nhấn giữ nút trái - đưa hình đến nơi mới - thả tay ra).", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS quan sát GV làm mẫu trên màn hình lớn.\n- HS thực hành không chạm (làm động tác trên không trung cùng GV).", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV gọi HS nhắc lại: Khi muốn gắp một quả táo bỏ vào rổ thì ta dùng thao tác gì? (Kéo thả chuột).", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV nhận xét, nhắc HS không ấn quá mạnh làm đau ngón tay và hỏng nút chuột.", False, False)
                ],
                "ket_qua": [
                    ("- HS nắm vững 4 thao tác chuột chuẩn xác:\n  + Di chuyển chuột\n  + Nháy chuột trái\n  + Nháy đúp chuột\n  + Kéo thả đối tượng.", False, False)
                ]
            },
            {
                "tieu_de": "3. HĐ LUYỆN TẬP - THỰC HÀNH (10 phút)\n*Mục tiêu: Bé tự tin điều khiển chuột hoàn thành trò chơi ghép hình và hái quả trên máy tính.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV mở phần mềm trò chơi thực hành ghép hình con vật trên máy HS.\n- Nhiệm vụ: Dùng chuột kéo thả các miếng ghép tròn, vuông để hoàn thiện bức tranh chú cún con.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS ngồi thẳng lưng, tay phải cầm chuột thực hành kéo thả.\n- GV đi từng bàn quan sát, uốn nắn bạn nào cầm chuột sai hoặc tay còn cứng.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS hoàn thành giơ tay để GV và bạn bên cạnh kiểm tra bức tranh.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- Tuyên dương các bạn có đôi tay khéo léo, hoàn thành nhanh và đẹp mắt.", False, False)
                ],
                "ket_qua": [
                    ("- 100% HS di chuyển được con trỏ chuột đến các mảnh ghép.\n- Đa số HS thực hiện thành công thao tác kéo thả hoàn thiện sản phẩm.", False, False)
                ]
            },
            {
                "tieu_de": "4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)\n*Mục tiêu: Rèn luyện ý thức giữ gìn vệ sinh và an toàn thiết bị phòng máy sau khi làm việc.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV đưa hiệu lệnh: \"Chuột về tổ!\". Yêu cầu HS đặt chuột ngay ngắn lên tấm lót chuột, không giật dây.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS đặt chuột vào vị trí quy định, đẩy bàn phím vào gầm bàn, xếp ghế ngay ngắn.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV cùng các tổ trưởng nhí đi chấm điểm xem tổ nào gọn gàng nhất.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV dặn dò: Về nhà các con có thể giúp bố mẹ di chuột chọn bài hát thiếu nhi yêu thích nhé!", False, False)
                ],
                "ket_qua": [
                    ("- Chuột và bàn phím được cất đặt ngăn nắp.\n- HS hình thành thói quen bảo vệ thiết bị công nghệ.", False, False)
                ]
            }
        ]
    },

    # 2. LỚP 1
    {
        "folder_name": "Lớp_1",
        "file_name": "KHBD_Tin_hoc_Lớp_1_Tiet10_On_tap_ky_nang_so.docx",
        "lop_name": "1A1",
        "ngay_day": "12/10/2026",
        "tiet_ppct": 10,
        "ten_bai": "BÀI 10. ÔN TẬP KỸ NĂNG SỐ",
        "yccd_list": [
            "Hệ thống hóa và thực hiện thành thạo thao tác khởi động máy tính, mở phần mềm đồ họa đơn giản (Paint / Tux Paint).",
            "Sử dụng kết hợp chuột máy tính (chọn cọ vẽ, chọn màu, kéo nét) và bàn phím (gõ phím cách, gõ tên mình) để tạo sản phẩm số.",
            "Biết thực hiện thao tác lưu tệp sản phẩm vào thư mục cá nhân và tắt máy an toàn sau khi học."
        ],
        "pham_chat": {
            "Chăm chỉ": "Say mê sáng tạo sản phẩm tranh vẽ trên máy tính, kiên trì thử lại khi thao tác chưa ưng ý. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Trách nhiệm": "Có ý thức giữ gìn sản phẩm của mình, cẩn thận lưu bài vào đúng nơi quy định. (Đạt được thông qua Hoạt động 3, Hoạt động 4)",
            "Trung thực": "Tự tay sáng tạo nét vẽ và gõ tên của chính mình. (Đạt được thông qua Hoạt động 3)"
        },
        "nl_dac_thu": {
            "NLa": {
                "ten": "Sử dụng và quản lí các phương tiện ICT",
                "bieu_hien": "Khởi động phần mềm đúng cách bằng thao tác nháy đúp chuột; thao tác thành thạo nút lưu (Save) và nút đóng (X). (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            "NLd": {
                "ten": "Ứng dụng ICT trong học và tự học",
                "bieu_hien": "Tạo ra một sản phẩm số đầu đời (bức tranh bông hoa/ngôi nhà kèm tên) thể hiện ý tưởng cá nhân. (Đạt được thông qua Hoạt động 3)"
            }
        },
        "nang_luc_so": [
            {
                "mien": "Miền III. Sáng tạo nội dung số",
                "thanh_to": "3.1. Phát triển nội dung số",
                "bac": 1,
                "bieu_hien": "Tạo nội dung đồ họa số đơn giản bằng cách vẽ nét, chọn màu và chèn văn bản tên mình vào tệp tranh. (Đạt được thông qua Hoạt động 3)"
            },
            {
                "mien": "Miền IV. An toàn",
                "thanh_to": "4.1. Bảo vệ thiết bị",
                "bac": 1,
                "bieu_hien": "Thực hiện đúng quy trình tắt máy tính an toàn (Start -> Power -> Shut down), không tự ý ngắt nguồn điện. (Đạt được thông qua Hoạt động 4)"
            }
        ],
        "nl_chung": {
            "Tự chủ và tự học": "Tự thực hiện các bước lưu tệp theo hướng dẫn của giáo viên và bảng quy trình trực quan. (Đạt được thông qua Hoạt động 3)",
            "Giao tiếp và hợp tác": "Tự tin giới thiệu sản phẩm bức tranh số của mình cho bạn bên cạnh và lắng nghe nhận xét. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "do_dung_gv": "Máy chiếu, máy tính GV, video mô phỏng quy trình tạo và lưu sản phẩm tranh vẽ số, phiếu khen thưởng \"Họa sĩ nhí công nghệ\".",
        "do_dung_hs": "Phòng thực hành máy tính có cài sẵn phần mềm Paint hoặc Tux Paint, tai nghe (nếu có).",
        "phuong_phap": "Phương pháp trực quan nêu gương, dạy học thực hành trải nghiệm, trò chơi đóng vai họa sĩ.",
        "ki_thuat": "Kĩ thuật đặt câu hỏi, động não nhanh, trưng bày phòng tranh (Gallery Walk).",
        "hoat_dong": [
            {
                "tieu_de": "1. Hoạt động MỞ ĐẦU (5 phút)\n*Mục tiêu: Khơi gợi niềm tự hào về các sản phẩm số đã học từ đầu năm đến nay; tạo hứng thú bước vào tiết ôn tập tổng hợp.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV chiếu một số bức tranh đẹp mà các bạn học sinh Lớp 1 đã vẽ trong các tiết trước:\n  + Các em có nhận ra đây là bức tranh của ai không?\n  + Bức tranh này được vẽ bằng bút chì hay vẽ trên máy tính?", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS quan sát tranh, hào hứng gọi tên các công cụ: Bút chì màu, Thùng sơn tô màu, Hình tròn, Ngôi sao.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- Đại diện 1 HS nhắc lại: Để vẽ được bức tranh trên máy tính, ta cần dùng chuột và bàn phím.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV dẫn dắt: Tiết học hôm nay chúng mình sẽ cùng trổ tài ôn tập toàn diện kỹ năng số để chuẩn bị cho bài Đánh giá định kỳ 1 nhé!", False, False)
                ],
                "ket_qua": [
                    ("- HS nhận thức được mục tiêu bài học: Tổng hợp các kỹ năng vẽ, gõ chữ và lưu tranh trên máy tính.", False, False)
                ]
            },
            {
                "tieu_de": "2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI / HỆ THỐNG HÓA KỸ NĂNG (15 phút)\n*Mục tiêu: Hệ thống hóa quy trình 4 bước tạo và quản lý sản phẩm số: Mở phần mềm -> Sáng tạo nét vẽ -> Gõ tên -> Lưu tệp.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV trình chiếu sơ đồ 4 bước tạo sản phẩm số chuẩn:\n  + Bước 1: Nháy đúp vào biểu tượng bảng màu Paint trên màn hình nền.\n  + Bước 2: Dùng chuột chọn công cụ Shapes (hình tròn, vuông) và Color để vẽ ngôi nhà/bông hoa.\n  + Bước 3: Chọn biểu tượng chữ A trên thanh công cụ, nháy chuột vào tranh và gõ tên mình (Ví dụ: AN 1A1).\n  + Bước 4: Nhấn biểu tượng đĩa mềm (Save) màu tím ở góc trái để lưu bài.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS quan sát từng thao tác chiếu chậm của GV.\n- HS nhắc lại tên nút Save và vị trí của nó.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV hỏi: Điều gì xảy ra nếu em vẽ xong một bức tranh rất đẹp mà quên bấm nút Save trước khi tắt máy?\n- HS trả lời: Bức tranh sẽ bị biến mất!", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV khắc sâu: Luôn luôn bấm nút Save (hoặc phím tắt Ctrl + S) để lưu giữ tác phẩm của mình.", False, False)
                ],
                "ket_qua": [
                    ("- HS nhớ quy trình 4 bước tạo sản phẩm số.\n- Hiểu rõ tầm quan trọng của việc lưu tệp vào máy tính.", False, False)
                ]
            },
            {
                "tieu_de": "3. HĐ LUYỆN TẬP - THỰC HÀNH (10 phút)\n*Mục tiêu: HS tự tay tạo một bức tranh hoàn chỉnh mang dấu ấn cá nhân, gõ tên và lưu bài thành công.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV giao thử thách: \"Bức tranh ước mơ của em\":\n  + Vẽ một hình yêu thích (ngôi nhà, mặt trời, chiếc xe hoặc bông hoa).\n  + Tô ít nhất 3 màu rực rỡ.\n  + Gõ tên và lớp của em ở góc dưới bức tranh.\n  + Lưu tệp với tên: Ten_1A1 (Ví dụ: Nam_1A1).", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS độc lập thao tác trên máy tính của mình.\n- GV đi bao quát, hỗ trợ những bạn gõ phím còn chậm hoặc lúng túng khi lưu tệp.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS quay sang đổi chỗ cho bạn bên cạnh để chiêm ngưỡng tác phẩm của nhau.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chọn chiếu bài của 3 bạn hoàn thành xuất sắc lên màn hình lớn, cả lớp cùng vỗ tay khen ngợi.", False, False)
                ],
                "ket_qua": [
                    ("- 100% HS tạo được hình vẽ có màu sắc.\n- Đa số HS gõ được tên mình và lưu được tệp vào thư mục theo hướng dẫn.", False, False)
                ]
            },
            {
                "tieu_de": "4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)\n*Mục tiêu: Đóng phần mềm, tắt máy tính đúng quy trình kỹ thuật và thu dọn vị trí học tập.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV hướng dẫn thao tác kết thúc tiết học:\n  + Nháy chuột vào nút dấu X màu đỏ ở góc trên cùng bên phải để đóng Paint.\n  + Nháy nút Start ở góc trái dưới -> Chọn Power -> Chọn Shut down.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS thực hiện tắt máy tính tuần tự, chờ màn hình tối hẳn mới đứng dậy.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV kiểm tra tình trạng tắt máy ở từng dãy bàn.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV khen ngợi cả lớp đã làm chủ các kỹ năng số cơ bản và sẵn sàng cho tiết đánh giá tuần tới.", False, False)
                ],
                "ket_qua": [
                    ("- Toàn bộ máy tính được tắt an toàn, đúng quy trình.\n- Phòng máy gọn gàng, bàn ghế ngăn nắp.", False, False)
                ]
            }
        ]
    },

    # 3. LỚP 2
    {
        "folder_name": "Lớp_2",
        "file_name": "KHBD_Tin_hoc_Lớp_2_Tiet10_On_tap_tep_thu_muc.docx",
        "lop_name": "2A1",
        "ngay_day": "13/10/2026",
        "tiet_ppct": 10,
        "ten_bai": "BÀI 10. ÔN TẬP TỆP – THƯ MỤC",
        "yccd_list": [
            "Phân biệt được rõ ràng giữa tệp (File - tài liệu, bài vẽ, bài hát) và thư mục (Folder - ngăn chứa màu vàng).",
            "Thực hiện thành thạo thao tác tạo thư mục mới (New folder), đổi tên thư mục và mở thư mục trong ổ đĩa học tập (D: hoặc E:).",
            "Lưu sản phẩm học tập vào đúng thư mục của mình và biết cách mở lại tệp đã lưu một cách nhanh chóng, chính xác."
        ],
        "pham_chat": {
            "Chăm chỉ": "Tự giác luyện tập thao tác tổ chức dữ liệu khoa học, kiên nhẫn tìm kiếm tệp bài tập. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Trách nhiệm": "Có ý thức cẩn thận, không xóa nhầm hoặc đổi tên thư mục của bạn khác trong cùng máy tính. (Đạt được thông qua Hoạt động 1, Hoạt động 4)",
            "Trung thực": "Chỉ thao tác và nộp đúng sản phẩm do chính mình tạo ra. (Đạt được thông qua Hoạt động 3)"
        },
        "nl_dac_thu": {
            "NLa": {
                "ten": "Sử dụng và quản lí các phương tiện ICT",
                "bieu_hien": "Biết cách duyệt cây thư mục trong phần mềm quản lý tệp (File Explorer); tạo và đặt tên thư mục đúng quy tắc. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            "NLc": {
                "ten": "Giải quyết vấn đề với sự hỗ trợ của ICT",
                "bieu_hien": "Sắp xếp tệp tranh vẽ và bài tập vào đúng thư mục tương ứng để tìm lại nhanh chóng khi cần. (Đạt được thông qua Hoạt động 3)"
            }
        },
        "nang_luc_so": [
            {
                "mien": "Miền I. Khai thác dữ liệu và thông tin",
                "thanh_to": "1.3. Quản lý dữ liệu",
                "bac": 1,
                "bieu_hien": "Tổ chức và lưu trữ các tệp số vào các thư mục có tên gợi nhớ để quản lý thông tin khoa học. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            {
                "mien": "Miền IV. An toàn",
                "thanh_to": "4.2. Bảo vệ dữ liệu cá nhân",
                "bac": 1,
                "bieu_hien": "Tôn trọng không gian lưu trữ của bạn bè, không tùy tiện sửa đổi hay xóa dữ liệu trong thư mục người khác. (Đạt được thông qua Hoạt động 4)"
            }
        ],
        "nl_chung": {
            "Tự chủ và tự học": "Chủ động kiểm tra lại đường dẫn lưu tệp của bản thân trước khi đóng ứng dụng. (Đạt được thông qua Hoạt động 3)",
            "Giao tiếp và hợp tác": "Hỗ trợ bạn bên cạnh tìm lại tệp bài vẽ khi bạn lỡ lưu nhầm chỗ. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "do_dung_gv": "Máy chiếu, máy tính GV, hình ảnh mô phỏng ngăn tủ đựng sách vở so sánh với thư mục máy tính, sơ đồ cây thư mục trực quan.",
        "do_dung_hs": "Máy tính có ổ đĩa D: hoặc E: để thực hành tạo cây thư mục, vở ghi môn Tin học.",
        "phuong_phap": "Phương pháp so sánh liên hệ thực tiễn, hướng dẫn trực quan từng bước, thực hành dự án mini.",
        "ki_thuat": "Kĩ thuật mảnh ghép, động não, chia sẻ nhóm đôi.",
        "hoat_dong": [
            {
                "tieu_de": "1. Hoạt động MỞ ĐẦU (5 phút)\n*Mục tiêu: Kết nối kiến thức thực tế về ngăn cặp sách với khái niệm thư mục trong máy tính; tạo động lực sắp xếp khoa học.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV đặt tình huống: \"Trong cặp sách của em có nhiều ngăn: ngăn để sách giáo khoa, ngăn để vở viết, ngăn để hộp bút. Vì sao chúng mình phải chia nhiều ngăn như vậy?\"\n- GV chiếu hình ảnh chiếc cặp sách ngăn nắp đối chiếu với cặp sách lộn xộn.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS trao đổi nhanh với bạn bên cạnh: Chia ngăn để tìm đồ dùng nhanh hơn, không bị quăn mép vở.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV gọi 2 HS trả lời.\n- GV dẫn tiếp: Trong máy tính, các bác kỹ sư cũng tạo ra những chiếc \"ngăn cặp\" màu vàng, gọi là gì nhỉ?", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- Cả lớp đồng thanh: THƯ MỤC (Folder)! GV kết luận và giới thiệu bài ôn tập tổng hợp kỹ năng quản lý tệp và thư mục.", False, False)
                ],
                "ket_qua": [
                    ("- HS hiểu được ý nghĩa của việc tổ chức, phân loại tệp và thư mục để tìm kiếm nhanh chóng.", False, False)
                ]
            },
            {
                "tieu_de": "2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI / ÔN TẬP KỸ NĂNG (15 phút)\n*Mục tiêu: Khắc sâu sự khác biệt giữa tệp và thư mục; củng cố các bước tạo, đổi tên thư mục và lưu tệp đúng chỗ.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV trình chiếu bảng so sánh trực quan:\n  + Thư mục (Folder): Biểu tượng kẹp giấy màu vàng, dùng để chứa các thư mục con và các tệp.\n  + Tệp (File): Chứa nội dung cụ thể (tranh vẽ .png, bài văn .docx, bài hát .mp3).\n- GV hướng dẫn thao tác chuẩn 3 bước tạo thư mục:\n  + Bước 1: Mở ổ đĩa D: (hoặc thư mục lớp Lop2A1).\n  + Bước 2: Nháy nút New folder trên thanh công cụ (hoặc nháy chuột phải -> New -> Folder).\n  + Bước 3: Gõ tên của em (Ví dụ: Minh_An) rồi nhấn phím Enter.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS theo dõi GV thao tác mẫu, ghi nhớ các quy tắc đặt tên thư mục không chứa ký tự đặc biệt.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV gọi 1 HS lên máy chủ thao tác tạo thư mục mang tên mình trong ổ đĩa học tập.\n- Cả lớp quan sát và nhận xét.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chuẩn hóa: Sau khi tạo thư mục xong, khi lưu bất kỳ bài tập nào, ta phải nháy đúp mở thư mục đó ra trước rồi mới nhấn Save.", False, False)
                ],
                "ket_qua": [
                    ("- HS phân biệt chính xác tệp và thư mục qua hình dạng biểu tượng.\n- Nắm vững thao tác tạo và đặt tên thư mục mới.", False, False)
                ]
            },
            {
                "tieu_de": "3. HĐ LUYỆN TẬP - THỰC HÀNH (10 phút)\n*Mục tiêu: Tự tạo cây thư mục cá nhân và lưu một tệp tranh vẽ vào đúng thư mục vừa tạo.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV giao nhiệm vụ thực hành trên máy:\n  + Nhiệm vụ 1: Vào ổ đĩa D: -> mở thư mục Lop_2A1 -> tạo thư mục riêng theo cú pháp: Ten_STT (Ví dụ: Lan_15).\n  + Nhiệm vụ 2: Mở phần mềm Paint, vẽ một chiếc lá hoặc bông hoa đơn giản.\n  + Nhiệm vụ 3: Lưu bức tranh vào đúng thư mục Lan_15 với tên tệp là: Chiec_la.\n  + Nhiệm vụ 4: Đóng Paint, vào lại ổ đĩa D: mở thư mục của mình để kiểm tra xem tệp đã nằm trong đó chưa.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS bắt tay vào thực hiện 4 nhiệm vụ tuần tự.\n- GV đi kiểm tra, hướng dẫn các bạn còn nhầm lẫn giữa nháy đơn và nháy đúp chuột khi mở thư mục.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS mở sẵn thư mục chứa tệp trên màn hình để GV đánh giá.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV tuyên dương các bạn thao tác chuẩn xác, nhanh nhẹn; lưu ý cả lớp không lưu bài ra Desktop lộn xộn.", False, False)
                ],
                "ket_qua": [
                    ("- 100% HS tạo được thư mục cá nhân đúng tên.\n- Trên 90% HS lưu tệp tranh vẽ vào đúng thư mục của mình và mở lại thành công.", False, False)
                ]
            },
            {
                "tieu_de": "4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)\n*Mục tiêu: Ý thức đạo đức và văn hóa số trong quản lý dữ liệu dùng chung trên máy tính phòng thực hành.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV đặt câu hỏi tình huống: \"Nếu em thấy trong thư mục Lop_2A1 có thư mục của bạn bên cạnh, em có được phép xóa hoặc đổi tên thư mục của bạn không? Vì sao?\"", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS suy nghĩ và thảo luận nhanh nhóm đôi.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS trả lời: Không được xóa, vì đó là bài làm của bạn, xóa đi bạn sẽ mất hết bài tập và rất buồn.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV kết luận quy tắc văn hóa số: Tôn trọng và bảo vệ tài sản số của người khác; hướng dẫn HS thoát các cửa sổ và tắt máy an toàn.", False, False)
                ],
                "ket_qua": [
                    ("- HS có ý thức bảo vệ dữ liệu của bạn bè.\n- Thu dọn chỗ ngồi, tắt máy trật tự.", False, False)
                ]
            }
        ]
    },

    # 4. LỚP 3
    {
        "folder_name": "Lớp_3",
        "file_name": "KHBD_Tin_hoc_Lớp_3_Tiet10_Bai_4_Lam_viec_voi_may_tinh_Tiet_2.docx",
        "lop_name": "3A1",
        "ngay_day": "15/10/2026",
        "tiet_ppct": 10,
        "ten_bai": "BÀI 4. LÀM VIỆC VỚI MÁY TÍNH (TIẾT 2)",
        "yccd_list": [
            "Nhận biết được màn hình nền (Desktop) và các thành phần chính trên màn hình: biểu tượng phần mềm (Icon), thanh công việc (Taskbar) và nút Start.",
            "Thực hiện thành thạo thao tác khởi động một phần mềm ứng dụng bằng cách nháy đúp chuột vào biểu tượng và đóng cửa sổ phần mềm bằng nút lệnh Close (X).",
            "Nắm vững và thực hiện đúng quy trình tắt máy tính an toàn (Start -> Power -> Shut down); có ý thức tuân thủ các quy tắc an toàn về điện và vệ sinh phòng máy tính."
        ],
        "pham_chat": {
            "Chăm chỉ": "Chủ động quan sát màn hình, tích cực thực hành các thao tác quản lý cửa sổ ứng dụng theo yêu cầu bài học. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Trách nhiệm": "Nghiêm túc tuân thủ quy tắc tắt máy an toàn, có ý thức tiết kiệm điện và bảo vệ thiết bị trong phòng máy. (Đạt được thông qua Hoạt động 3, Hoạt động 4)",
            "Trung thực": "Báo cáo trung thực khi máy tính gặp sự cố hoặc thao tác chưa đúng quy trình. (Đạt được thông qua Hoạt động 3)"
        },
        "nl_dac_thu": {
            "NLa": {
                "ten": "Sử dụng và quản lí các phương tiện ICT",
                "bieu_hien": "Nhận biết các biểu tượng phần mềm trên màn hình nền; thao tác mở, thu nhỏ, phóng to và đóng cửa sổ làm việc; tắt máy tính đúng trình tự. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            "NLc": {
                "ten": "Giải quyết vấn đề với sự hỗ trợ của ICT",
                "bieu_hien": "Lựa chọn đúng biểu tượng phần mềm cần dùng để phục vụ yêu cầu học tập (ví dụ: mở trình duyệt hoặc phần mềm luyện gõ phím). (Đạt được thông qua Hoạt động 2)"
            }
        },
        "nang_luc_so": [
            {
                "mien": "Miền IV. An toàn",
                "thanh_to": "4.1. Bảo vệ thiết bị",
                "bac": 2,
                "bieu_hien": "Giải thích được tác hại của việc tắt máy tính bằng cách rút phích cắm đột ngột và thực hiện thành thạo quy trình Shut down theo đúng hướng dẫn kỹ thuật. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
            },
            {
                "mien": "Miền V. Giải quyết vấn đề",
                "thanh_to": "5.1. Giải quyết vấn đề kỹ thuật",
                "bac": 2,
                "bieu_hien": "Nhận biết trạng thái máy tính đã tắt hoàn toàn (đèn tín hiệu tắt, màn hình tối) trước khi rời khỏi vị trí làm việc. (Đạt được thông qua Hoạt động 4)"
            }
        ],
        "nl_chung": {
            "Tự chủ và tự học": "Tự thực hành mở và đóng các phần mềm theo trình tự các bước trong sách giáo khoa. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Giao tiếp và hợp tác": "Lắng nghe, quan sát và hỗ trợ bạn trong nhóm thực hiện quy trình tắt máy an toàn. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "do_dung_gv": "Máy chiếu, máy tính GV, SGK Tin học 3 (bộ sách Kết nối tri thức với cuộc sống), video minh họa quy trình khởi động và tắt máy tính an toàn.",
        "do_dung_hs": "SGK Tin học 3, vở ghi, máy tính thực hành phòng máy.",
        "phuong_phap": "Phương pháp thuyết trình trực quan, làm mẫu từng bước, thực hành cá nhân kết hợp kiểm tra nhóm đôi.",
        "ki_thuat": "Kĩ thuật đặt câu hỏi, động não, giao nhiệm vụ từng chặng.",
        "hoat_dong": [
            {
                "tieu_de": "1. Hoạt động MỞ ĐẦU (5 phút)\n*Mục tiêu: Nhắc lại kiến thức tư thế ngồi và khởi động máy ở Tiết 1; tạo tình huống khám phá khu vực làm việc trên màn hình máy tính.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV chiếu hình ảnh một chiếc bàn học có sách vở, hộp bút và hình ảnh màn hình máy tính sau khi bật nguồn:\n  + Câu hỏi: Khi ngồi vào bàn học, em cần sắp xếp bàn học như thế nào?\n  + Màn hình máy tính sau khi bật lên trông giống chiếc bàn học ở điểm nào?", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS quan sát hình ảnh và suy nghĩ cá nhân, trao đổi nhanh với bạn bên cạnh.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- Đại diện 2 HS trả lời: Màn hình máy tính cũng có các hình vẽ nhỏ giống như các cuốn sách, đồ dùng để trên bàn.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV nhận xét, chốt kiến thức: Màn hình đầu tiên xuất hiện sau khi khởi động máy tính được gọi là Màn hình nền (Desktop) - nơi chứa các công cụ làm việc của chúng ta. Hôm nay chúng mình cùng học Tiết 2 của Bài 4 nhé!", False, False)
                ],
                "ket_qua": [
                    ("- HS nhận diện được màn hình nền Desktop sau khi bật máy.\n- Sẵn sàng bước vào các hoạt động thực hành làm việc với máy tính.", False, False)
                ]
            },
            {
                "tieu_de": "2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI (15 phút)\n*Mục tiêu: HS nhận biết khu vực Desktop, các biểu tượng phần mềm, cách mở/đóng cửa sổ ứng dụng và quy trình tắt máy an toàn.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV hướng dẫn HS quan sát SGK trang 20 - 22 và màn hình máy chiếu:\n  + Nhiệm vụ 1: Tìm hiểu Màn hình nền (Desktop) và Biểu tượng (Icon). Quan sát thanh Taskbar ở đáy màn hình và nút Start ở góc trái.\n  + Nhiệm vụ 2: Tìm hiểu cách mở phần mềm: Nháy đúp chuột vào biểu tượng phần mềm trên Desktop. Quan sát 3 nút điều khiển ở góc trên bên phải cửa sổ: Thu nhỏ (-), Phóng to/Thu lại (ô vuông), Đóng (X màu đỏ).\n  + Nhiệm vụ 3: Tìm hiểu quy trình tắt máy tính an toàn (Shut down):\n    * Bước 1: Đóng tất cả các cửa sổ đang mở bằng nút X.\n    * Bước 2: Nháy chuột vào nút Start ở góc trái dưới màn hình.\n    * Bước 3: Nháy vào biểu tượng Power (Nguồn).\n    * Bước 4: Nháy chọn lệnh Shut down và chờ máy tắt hẳn.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS theo dõi GV thao tác mẫu trên máy chiếu.\n- Đọc thầm các bước trong SGK và nhắc lại các bước tắt máy cho bạn cùng bàn.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV hỏi: Nếu em tắt máy bằng cách bấm giữ nút nguồn trên thân máy hoặc rút phích cắm thì điều gì sẽ xảy ra?\n- HS trả lời: Máy tính có thể bị hỏng phần cứng, mất dữ liệu và giảm tuổi thọ thiết bị.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chuẩn hóa quy trình, chiếu hộp GHI NHỚ màu vàng trong SGK:\n  + Nháy đúp chuột vào biểu tượng để khởi động phần mềm.\n  + Nháy vào nút X để đóng cửa sổ phần mềm.\n  + Luôn tắt máy tính bằng lệnh Shut down trong bảng chọn Start.", False, False)
                ],
                "ket_qua": [
                    ("- HS gọi tên đúng: Desktop, Icon, Taskbar, nút Start.\n- Nắm chắc chức năng của nút Close (X) và lệnh Shut down.\n- Hiểu lý do vì sao phải tắt máy theo đúng quy trình.", False, False)
                ]
            },
            {
                "tieu_de": "3. HĐ LUYỆN TẬP - THỰC HÀNH (10 phút)\n*Mục tiêu: Tự thực hiện thành thạo thao tác mở, chuyển đổi và đóng cửa sổ phần mềm; thực hành quy trình tắt máy tính an toàn.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV giao nhiệm vụ luyện tập trên máy tính của mỗi HS:\n  + Thao tác 1: Tìm biểu tượng phần mềm trình duyệt web (Google Chrome) hoặc phần mềm Paint trên Desktop, nháy đúp chuột để mở.\n  + Thao tác 2: Nháy nút Phóng to để cửa sổ tràn màn hình, sau đó nháy lại nút đó để thu nhỏ kích thước.\n  + Thao tác 3: Nháy nút Close (X) để đóng cửa sổ phần mềm.\n  + Thao tác 4: Thực hiện lại với một phần mềm khác.\n  + Thao tác 5: Thực hành quy trình tắt máy tính bằng lệnh Shut down khi có hiệu lệnh của GV.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS độc lập thao tác trên máy tính cá nhân.\n- GV đi từng dãy bàn quan sát, uốn nắn thao tác nháy đúp (nhiều bạn nháy quá chậm khiến máy không nhận lệnh).", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS báo cáo kết quả khi đóng mở thành công 2 phần mềm khác nhau.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV nhận xét tinh thần luyện tập của cả lớp, tuyên dương những bạn thao tác chuột dứt khoát và chuẩn xác.", False, False)
                ],
                "ket_qua": [
                    ("- 100% HS mở được ít nhất 1 phần mềm bằng thao tác nháy đúp.\n- 100% HS đóng được cửa sổ bằng nút X.\n- HS nắm vững quy trình các bước chuẩn bị tắt máy.", False, False)
                ]
            },
            {
                "tieu_de": "4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)\n*Mục tiêu: Đánh giá nhanh bằng câu hỏi trắc nghiệm và thực hiện tắt máy đồng loạt, kiểm tra an toàn phòng máy.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV chiếu câu hỏi trắc nghiệm nhanh:\n  \"Để tắt máy tính an toàn, em thực hiện theo trình tự nào sau đây?\"\n  A. Rút phích cắm điện ngay lập tức.\n  B. Nhấn giữ nút nguồn trên thân máy tính trong 5 giây.\n  C. Chọn nút Start -> Power -> Shut down.\n  D. Nhấn nút tắt trên màn hình rồi bỏ về.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- Cả lớp đồng thanh chọn đáp án C.\n- GV phát lệnh: Cả lớp cùng thực hiện thao tác Shut down trên máy của mình.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS quan sát đèn tín hiệu trên thân máy và màn hình, chờ màn hình tắt hẳn.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV kiểm tra lần cuối toàn phòng máy, nhắc nhở HS đẩy ghế, xếp gọn đồ dùng trước khi ra về.", False, False)
                ],
                "ket_qua": [
                    ("- HS chọn đúng đáp án C.\n- Toàn bộ máy tính được tắt an toàn, đúng quy trình kỹ thuật.\n- Phòng máy tính an toàn, gọn gàng.", False, False)
                ]
            }
        ]
    },

    # 5. LỚP 4
    {
        "folder_name": "Lớp_4",
        "file_name": "KHBD_Tin_hoc_Lớp_4_Tiet10_Bai_4_Tim_kiem_thong_tin_tren_Internet_Tiet_2.docx",
        "lop_name": "4C1",
        "ngay_day": "14/10/2026",
        "tiet_ppct": 10,
        "ten_bai": "BÀI 4. TÌM KIẾM THÔNG TIN TRÊN INTERNET (TIẾT 2)",
        "yccd_list": [
            "Xác định được từ khóa tìm kiếm chính xác, ngắn gọn và phù hợp với chủ đề tìm kiếm cho trước.",
            "Thực hiện thành thạo các bước tìm kiếm thông tin dạng văn bản và tìm kiếm hình ảnh bằng máy tìm kiếm (Google).",
            "Biết nháy chuột vào siêu liên kết (Hyperlink) để mở trang web xem thông tin chi tiết; biết lưu hình ảnh phục vụ học tập và có ý thức chọn lọc thông tin đáng tin cậy."
        ],
        "pham_chat": {
            "Chăm chỉ": "Tích cực tìm tòi, chủ động tìm kiếm các tư liệu phong phú phục vụ bài học lịch sử địa phương và danh lam thắng cảnh. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Trách nhiệm": "Có ý thức chọn lọc nguồn tin từ các trang web chính thống, không xem và không chia sẻ các thông tin độc hại, không lành mạnh. (Đạt được thông qua Hoạt động 2, Hoạt động 4)",
            "Trung thực": "Tôn trọng quyền tác giả và nguồn gốc tư liệu khi khai thác hình ảnh từ Internet. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "nl_dac_thu": {
            "NLa": {
                "ten": "Sử dụng và quản lí các phương tiện ICT",
                "bieu_hien": "Khởi động trình duyệt web, truy cập đúng địa chỉ máy tìm kiếm (google.com) và thao tác tải hình ảnh về máy tính. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            "NLc": {
                "ten": "Giải quyết vấn đề với sự hỗ trợ của ICT",
                "bieu_hien": "Biết thu hẹp phạm vi tìm kiếm bằng cách bổ sung từ khóa cụ thể (ví dụ: 'Khuê Văn Các Văn Miếu Quốc Tử Giám') để tìm đúng thông tin mong muốn. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            }
        },
        "nang_luc_so": [
            {
                "mien": "Miền I. Khai thác dữ liệu và thông tin",
                "thanh_to": "1.1. Duyệt, tìm kiếm và lọc dữ liệu",
                "bac": 2,
                "bieu_hien": "Lựa chọn từ khóa tìm kiếm phù hợp, thực hiện tìm kiếm thông tin và hình ảnh trên Internet, lọc kết quả từ các siêu liên kết đáng tin cậy. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
            },
            {
                "mien": "Miền IV. An toàn",
                "thanh_to": "4.2. Bảo vệ dữ liệu cá nhân",
                "bac": 2,
                "bieu_hien": "Nhận biết các cảnh báo an toàn từ trình duyệt khi truy cập trang web lạ, tuyệt đối không nhấn vào các đường link quảng cáo đáng ngờ. (Đạt được thông qua Hoạt động 4)"
            }
        ],
        "nl_chung": {
            "Tự chủ và tự học": "Tự giác thực hành theo các nhiệm vụ trong SGK, ghi chép lại những thông tin bổ ích tìm được vào vở. (Đạt được thông qua Hoạt động 2, Hoạt động 3)",
            "Giao tiếp và hợp tác": "Thảo luận nhóm để chọn lọc ra 3 hình ảnh đẹp nhất về Văn Miếu - Quốc Tử Giám để trình bày trước lớp. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        },
        "do_dung_gv": "Máy chiếu, máy tính GV có kết nối mạng Internet ổn định, SGK Tin học 4 (Kết nối tri thức), phiếu học tập \"Hành trình khám phá di sản\".",
        "do_dung_hs": "SGK Tin học 4, máy tính thực hành có kết nối Internet và trình duyệt web (Google Chrome / Edge), vở ghi.",
        "phuong_phap": "Phương pháp dạy học giải quyết vấn đề, dạy học hợp tác theo nhóm, thực hành trực tuyến có định hướng.",
        "ki_thuat": "Kĩ thuật KWL (Biết - Muốn biết - Đã học), động não, chia sẻ cặp đôi.",
        "hoat_dong": [
            {
                "tieu_de": "1. Hoạt động MỞ ĐẦU (5 phút)\n*Mục tiêu: Ôn lại khái niệm máy tìm kiếm và từ khóa ở Tiết 1; tạo tình huống xuất phát hấp dẫn về chuyến tham quan di tích lịch sử.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV chiếu hình ảnh Khuê Văn Các và đặt câu hỏi:\n  + Các em có biết đây là di tích lịch sử nổi tiếng nào ở Hà Nội không?\n  + Nếu muốn tìm hiểu xem Văn Miếu - Quốc Tử Giám được xây dựng từ năm nào và có những cảnh đẹp gì, chúng ta có thể dùng công cụ nào trên máy tính?", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS nhận ra hình ảnh Văn Miếu - Quốc Tử Giám và trả lời: Dùng máy tìm kiếm Google trên Internet.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV hỏi: Ở Tiết 1, chúng ta đã biết muốn tìm nhanh thì phải có gì?\n- HS đồng thanh: Phải có TỪ KHÓA!", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV dẫn dắt: Đúng vậy! Hôm nay ở Tiết 2, cô trò mình sẽ cùng thực hành đóng vai \"Hướng dẫn viên du lịch nhí\" để tìm kiếm thông tin và hình ảnh tuyệt đẹp về Văn Miếu nhé!", False, False)
                ],
                "ket_qua": [
                    ("- HS xác định được mục tiêu tiết học: Thực hành tìm kiếm thông tin và hình ảnh theo chủ đề trên Internet bằng từ khóa chính xác.", False, False)
                ]
            },
            {
                "tieu_de": "2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI / THỰC HÀNH MẪU (15 phút)\n*Mục tiêu: Hướng dẫn kỹ năng chọn từ khóa tối ưu, kỹ năng tìm kiếm hình ảnh (Images) và nháy chuột vào siêu liên kết.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV hướng dẫn HS quan sát SGK trang 22 - 24 và thao tác mẫu trên máy chiếu:\n  + Nhiệm vụ 1: Lựa chọn từ khóa thông minh:\n    * So sánh các từ khóa: \"Văn Miếu\" vs \"Lịch sử Văn Miếu Quốc Tử Giám\" vs \"Hình ảnh Khuê Văn Các\".\n    * Nhấn mạnh: Từ khóa càng cụ thể thì kết quả nhận được càng chính xác.\n  + Nhiệm vụ 2: Mở siêu liên kết (Hyperlink):\n    * Khi rê con trỏ chuột vào dòng tiêu đề màu xanh, con trỏ biến thành hình bàn tay chỉ ngón trỏ.\n    * Nháy chuột vào siêu liên kết để mở trang web đọc thông tin chi tiết.\n  + Nhiệm vụ 3: Tìm kiếm hình ảnh và lưu ảnh:\n    * Chọn mục \"Hình ảnh\" (Images) trên trang kết quả Google.\n    * Nháy chuột phải vào bức ảnh ưng ý -> Chọn \"Lưu hình ảnh thành...\" (Save image as...) -> Chọn thư mục của mình -> Nhấn Save.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS theo dõi từng bước GV thao tác mẫu, ghi chú 3 thao tác quan trọng vào vở.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- GV gọi 1 HS lên bảng nhắc lại các bước lưu hình ảnh từ Internet về máy tính.\n- GV đặt câu hỏi tương tác: Làm sao biết trang web nào là trang tin chính thống, đáng tin cậy?\n- HS trả lời: Chọn các trang web của cơ quan nhà nước, bảo tàng, hoặc trang web du lịch chính thức.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chốt quy trình và nhấn mạnh quy tắc đạo đức: Chỉ tải hình ảnh phục vụ học tập, không tải tệp lạ không rõ nguồn gốc.", False, False)
                ],
                "ket_qua": [
                    ("- HS nắm vững các bước tìm kiếm hình ảnh chuyên biệt.\n- Biết cách nhận diện siêu liên kết và thao tác lưu hình ảnh về máy.", False, False)
                ]
            },
            {
                "tieu_de": "3. HĐ LUYỆN TẬP - THỰC HÀNH (10 phút)\n*Mục tiêu: HS tự mình tìm kiếm thông tin và tải được ít nhất 2 hình ảnh về Văn Miếu lưu vào thư mục cá nhân.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV giao nhiệm vụ thực hành theo phiếu bài tập:\n  + Thử thách 1: Mở trình duyệt Google Chrome, truy cập google.com.\n  + Thử thách 2: Nhập từ khóa \"Văn Miếu Quốc Tử Giám\", nháy vào 1 trang web uy tín để tìm câu trả lời: Văn Miếu được xây dựng vào năm nào? (Năm 1070).\n  + Thử thách 3: Chuyển sang thẻ Hình ảnh, tìm và lưu 2 bức ảnh (1 ảnh Khuê Văn Các, 1 ảnh Vườn bia Tiến sĩ) vào thư mục riếng của em trong ổ đĩa D:.", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS thực hành độc lập trên máy tính.\n- GV đi bao quát các máy, kịp thời hỗ trợ các em gặp trục trặc về bàn phím tiếng Việt hoặc thao tác nháy chuột phải.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS giơ tay báo cáo khi đã tải xong 2 ảnh vào thư mục và tìm được năm xây dựng Văn Miếu.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV chiếu thư mục bài làm của một số bạn lên màn hình lớn, khen ngợi các bạn đã chọn được ảnh rất đẹp, rõ nét và lưu đúng chỗ.", False, False)
                ],
                "ket_qua": [
                    ("- 100% HS truy cập được máy tìm kiếm và nhập từ khóa tiếng Việt có dấu.\n- Trên 90% HS tìm được câu trả lời năm 1070 và lưu được 2 bức ảnh về thư mục cá nhân.", False, False)
                ]
            },
            {
                "tieu_de": "4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)\n*Mục tiêu: Vận dụng kỹ năng tìm kiếm để khám phá một địa danh tại quê hương mình; rèn ý thức an toàn thông tin trên Internet.",
                "gv_hs": [
                    ("Bước 1: Chuyển giao nhiệm vụ:", True, True),
                    ("- GV giao nhiệm vụ vận dụng mở rộng:\n  \"Về nhà, em hãy cùng bố mẹ dùng máy tìm kiếm để tra cứu thông tin về một danh lam thắng cảnh hoặc di tích lịch sử tại quê hương em.\"\n- GV hỏi nhanh: \"Khi đang lướt web, nếu thấy xuất hiện một cửa sổ lạ nhấp nháy báo trúng thưởng, em sẽ làm gì?\"", False, False),
                    ("Bước 2: Thực hiện nhiệm vụ:", True, True),
                    ("- HS suy nghĩ và thảo luận nhanh.", False, False),
                    ("Bước 3: Báo cáo, thảo luận:", True, True),
                    ("- HS trả lời: Đóng ngay cửa sổ đó lại bằng nút X, tuyệt đối không nháy chuột vào vì có thể chứa virus hoặc lừa đảo.", False, False),
                    ("Bước 4: Kết luận, nhận định:", True, True),
                    ("- GV tuyên dương câu trả lời chính xác, nhắc nhở toàn lớp đóng trình duyệt và tắt máy tính an toàn trước khi kết thúc tiết học.", False, False)
                ],
                "ket_qua": [
                    ("- HS có kỹ năng phòng tránh cạm bẫy quảng cáo và bảo vệ an toàn trên mạng Internet.\n- Máy tính được tắt cẩn thận, phòng máy an toàn.", False, False)
                ]
            }
        ]
    }
]

def main():
    print("=== START GENERATING 5 KHBD FOR TUẦN 11 (PPCT 10) ===")
    for lesson in LESSONS_DATA:
        build_khbd_doc(lesson)
    print("=== ALL 5 KHBD COMPLETED SUCCESSFULLY! ===")

if __name__ == "__main__":
    main()
