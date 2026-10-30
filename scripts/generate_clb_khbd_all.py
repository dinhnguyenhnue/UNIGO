# -*- coding: utf-8 -*-
"""
Script sinh tự động toàn bộ 70 Kế hoạch bài dạy (KHBD) cho CLB Robotics UNIGO
- Phân hệ 1 (Lớp 1 - 2): 35 buổi (Tuần 1-17: OLLO Initiate; Tuần 18-35: OLLO Kinder)
- Phân hệ 2 (Lớp 3 - 7): 35 buổi (Tuần 1-17: OLLO Kinder; Tuần 18-35: OLLO Excel)
Tuân thủ 100% quy chuẩn UNIGO (AGENTS.md).
"""

import sys, os, re, shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

BASE_KHBD_DIR = r"D:\UNIGO\CLB_Robotics\KHBD_CLB_Robotics"
DIR_LOP_1_2 = os.path.join(BASE_KHBD_DIR, "01_Phan_he_Lop_1_2 (Initiate - Kinder)")
DIR_LOP_3_7 = os.path.join(BASE_KHBD_DIR, "02_Phan_he_Lop_3_7 (Kinder - Excel)")

os.makedirs(DIR_LOP_1_2, exist_ok=True)
os.makedirs(DIR_LOP_3_7, exist_ok=True)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def sanitize_filename(name):
    clean = re.sub(r'[\\/*?:"<>|]', "", name)
    clean = clean.replace(" ", "_").replace("–", "-")
    clean = re.sub(r'_+', '_', clean)
    return clean

CALENDAR = [
    ("12/09/2026", "05/09/2026"),
    ("19/09/2026", "12/09/2026"),
    ("26/09/2026", "19/09/2026"),
    ("03/10/2026", "26/09/2026"),
    ("10/10/2026", "03/10/2026"),
    ("17/10/2026", "10/10/2026"),
    ("24/10/2026", "17/10/2026"),
    ("31/10/2026", "24/10/2026"),
    ("07/11/2026", "31/10/2026"), # ĐGĐK 1
    ("14/11/2026", "07/11/2026"),
    ("21/11/2026", "14/11/2026"),
    ("28/11/2026", "21/11/2026"),
    ("05/12/2026", "28/11/2026"),
    ("12/12/2026", "05/12/2026"),
    ("19/12/2026", "12/12/2026"),
    ("26/12/2026", "19/12/2026"),
    ("02/01/2027", "26/12/2026"), # ĐGĐK 2
    # HK2
    ("09/01/2027", "02/01/2027"),
    ("16/01/2027", "09/01/2027"),
    ("23/01/2027", "16/01/2027"),
    ("30/01/2027", "23/01/2027"),
    ("20/02/2027", "13/02/2027"),
    ("27/02/2027", "20/02/2027"),
    ("06/03/2027", "27/02/2027"),
    ("13/03/2027", "06/03/2027"),
    ("20/03/2027", "13/03/2027"),
    ("27/03/2027", "20/03/2027"),
    ("03/04/2027", "27/03/2027"), # ĐGĐK 3
    ("10/04/2027", "03/04/2027"),
    ("17/04/2027", "10/04/2027"),
    ("24/04/2027", "17/04/2027"),
    ("08/05/2027", "01/05/2027"),
    ("15/05/2027", "08/05/2027"),
    ("22/05/2027", "15/05/2027"), # ĐGĐK 4
    ("29/05/2027", "22/05/2027")  # Bế mạc
]

def create_khbd_doc(
    buoi_num,
    ngay_day,
    ngay_soan,
    group_name,       # "CLB Robotics (Khối 1 - 2)" hoặc "CLB Robotics (Khối 3 - 7)"
    to_chuyen_mon,    # "Tổ chuyên môn Tiểu học" hoặc "Tổ chuyên môn THCS"
    chu_diem_kit,     # "OLLO INITIATE", "OLLO KINDER", "OLLO EXCEL"
    bai_so,           # "Bài 1", "Buổi 17"...
    ten_bai,          # "TẬP THỂ DỤC NÀO!"
    yccd_kienthuc,    # List[str] dạng danh từ
    pc_chamdchi,
    pc_trachnhiem,
    pc_trungthuc,
    nl_dacthu_items,  # [(mã NL, tên NL, mô tả, thông qua HĐ)]
    nls_items,        # [(Miền, thành tố, bậc, mô tả, thông qua HĐ)]
    nl_chung_items,   # [(Tên, mô tả, thông qua HĐ)]
    doday_gv,
    doday_hs,
    hd1_title, hd1_gv_hs, hd1_kq,
    hd2_title, hd2_gv_hs, hd2_kq,
    hd3_title, hd3_gv_hs, hd3_kq,
    hd4_title, hd4_gv_hs, hd4_kq,
    out_filepath
):
    doc = docx.Document()

    sec = doc.sections[0]
    sec.top_margin = Inches(0.75)
    sec.bottom_margin = Inches(0.75)
    sec.left_margin = Inches(0.8)
    sec.right_margin = Inches(0.8)

    # 1. Table 0: Thông tin đầu bài (3x2 NO BORDER)
    t0 = doc.add_table(rows=3, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t0)
    w0 = [Inches(3.8), Inches(3.2)]

    row0 = t0.rows[0].cells
    row0[0].text = "Trường: Tiểu học và THCS UNIGO"
    row0[0].paragraphs[0].runs[0].bold = True
    row0[1].text = f"Ngày soạn: {ngay_soan}"
    row0[1].paragraphs[0].runs[0].bold = True

    row1 = t0.rows[1].cells
    row1[0].text = "GV: Đậu Đình Nguyên"
    row1[1].text = f"Ngày dạy: {ngay_day}"

    row2 = t0.rows[2].cells
    row2[0].text = f"Tổ: {to_chuyen_mon}"
    row2[1].text = f"Lớp: {group_name}"

    for r in t0.rows:
        for i, c in enumerate(r.cells):
            c.width = w0[i]
            set_cell_margins(c, top=40, bottom=40, left=60, right=60)

    # 2. Tiêu đề bài dạy
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("KẾ HOẠCH DẠY HỌC CÂU LẠC BỘ ROBOTICS")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0, 32, 96)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"CHỦ ĐIỂM: BỘ THIẾT BỊ: {chu_diem_kit}")
    r.bold = True
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(f"{bai_so.upper()}. {ten_bai.upper()} (Buổi {buoi_num} – Thứ 7: 09h45 - 11h15)")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(192, 0, 0)

    # I. YÊU CẦU CẦN ĐẠT
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("I. YÊU CẦU CẦN ĐẠT:")
    r.bold = True

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("- Sau bài học này em sẽ:")
    r.italic = True

    for item in yccd_kienthuc:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.2)
        r = p.add_run(f"+ {item}")

    # 1. Phát triển phẩm chất
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("1. Phát triển phẩm chất")
    r.bold = True

    doc.add_paragraph(f"- Chăm chỉ: {pc_chamdchi}")
    doc.add_paragraph(f"- Trách nhiệm: {pc_trachnhiem}")
    doc.add_paragraph(f"- Trung thực: {pc_trungthuc}")

    # 2. Phát triển năng lực
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("2. Phát triển năng lực")
    r.bold = True

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("2.1. Năng lực đặc thù (Robotics):")
    r.italic = True
    for nl_code, nl_name, nl_desc, nl_hds in nl_dacthu_items:
        doc.add_paragraph(f"- {nl_code} ({nl_name}): {nl_desc} (Đạt được thông qua {nl_hds})")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("2.2. Năng lực số (Thông tư 02/2025 – CV 3456):")
    r.italic = True
    for mien, tc, bac, desc, nl_hds in nls_items:
        doc.add_paragraph(f"- {mien} (thành tố {tc} – Bậc {bac}): {desc} (Đạt được thông qua {nl_hds})")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("2.3. Năng lực chung:")
    r.italic = True
    for nl_name, nl_desc, nl_hds in nl_chung_items:
        doc.add_paragraph(f"- {nl_name}: {nl_desc} (Đạt được thông qua {nl_hds})")

    # II. ĐỒ DÙNG DẠY HỌC
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("II. ĐỒ DÙNG DẠY HỌC:")
    r.bold = True

    doc.add_paragraph(f"1. Giáo viên: {doday_gv}")
    doc.add_paragraph(f"2. Học sinh: {doday_hs}")

    # III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC:")
    r.bold = True
    doc.add_paragraph("- Phương pháp dạy học: Dạy học qua dự án (PBL), phương pháp bàn tay nặn bột, thực hành trải nghiệm, làm việc nhóm, giải quyết vấn đề kỹ thuật.")
    doc.add_paragraph("- Kĩ thuật dạy học: Kĩ thuật động não (Brainstorming), đặt câu hỏi gợi mở, hướng dẫn mẫu từng bước và đánh giá đồng đẳng.")

    # IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU (THỜI LƯỢNG: 90 PHÚT):")
    r.bold = True

    tbl_proc = doc.add_table(rows=1, cols=2)
    tbl_proc.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_proc)
    w_proc = [Inches(4.2), Inches(2.8)]

    # Header row
    h0 = tbl_proc.rows[0].cells[0]
    h1 = tbl_proc.rows[0].cells[1]
    h0.text = "HOẠT ĐỘNG CỦA GV – HS"
    h1.text = "KẾT QUẢ CẦN ĐẠT"
    h0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    h0.paragraphs[0].runs[0].bold = True
    h1.paragraphs[0].runs[0].bold = True
    set_cell_background(h0, "D9E1F2")
    set_cell_background(h1, "D9E1F2")
    set_cell_margins(h0)
    set_cell_margins(h1)
    h0.width = w_proc[0]
    h1.width = w_proc[1]

    acts = [
        (hd1_title, hd1_gv_hs, hd1_kq),
        (hd2_title, hd2_gv_hs, hd2_kq),
        (hd3_title, hd3_gv_hs, hd3_kq),
        (hd4_title, hd4_gv_hs, hd4_kq)
    ]

    for title, gv_hs, kq in acts:
        # Title row (spanning 2 columns)
        r_t = tbl_proc.add_row()
        c_t = r_t.cells[0]
        r_t.cells[0].merge(r_t.cells[1])
        c_t.text = title
        c_t.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        c_t.paragraphs[0].runs[0].bold = True
        set_cell_background(c_t, "F2F2F2")
        set_cell_margins(c_t)
        c_t.width = Inches(7.0)

        # Content row
        r_c = tbl_proc.add_row()
        c0 = r_c.cells[0]
        c1 = r_c.cells[1]
        c0.text = gv_hs
        c1.text = kq
        set_cell_margins(c0)
        set_cell_margins(c1)
        c0.width = w_proc[0]
        c1.width = w_proc[1]

    # V. RÚT KINH NGHIỆM SAU BÀI DẠY
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("V. RÚT KINH NGHIỆM SAU BÀI DẠY:")
    r.bold = True
    doc.add_paragraph("...........................................................................................................................................................................")
    doc.add_paragraph("...........................................................................................................................................................................")

    # VI. BẢNG CHỮ KÝ 3x3 NO BORDER
    t_sig = doc.add_table(rows=3, cols=3)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t_sig)
    w_sig = [Inches(2.3), Inches(2.3), Inches(2.4)]

    # Row 0
    headers_sig = ["DUYỆT CỦA BGH", "DUYỆT CỦA TỔ CM", "NGƯỜI SOẠN"]
    for i, h in enumerate(headers_sig):
        c = t_sig.cell(0, i)
        c.text = h
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].bold = True
        c.width = w_sig[i]
        set_cell_margins(c)

    # Row 1
    sub_sig = ["(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"]
    for i, s in enumerate(sub_sig):
        c = t_sig.cell(1, i)
        c.text = s
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].italic = True
        c.width = w_sig[i]
        set_cell_margins(c)

    # Row 2
    names_sig = ["\n\n\n\n", "\n\n\n\n", "\n\n\n\nĐậu Đình Nguyên"]
    for i, n in enumerate(names_sig):
        c = t_sig.cell(2, i)
        c.text = n
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if "Đậu Đình Nguyên" in n:
            c.paragraphs[0].runs[0].bold = True
        c.width = w_sig[i]
        set_cell_margins(c)

    # Font Times New Roman 13pt toàn bộ
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.name = "Times New Roman"
            r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
            if not r.font.size:
                r.font.size = Pt(13)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
                        if not r.font.size:
                            r.font.size = Pt(13)

    doc.save(out_filepath)
    return out_filepath

print("Core generator module defined.")
