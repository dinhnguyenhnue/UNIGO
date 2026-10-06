# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════════
  TẠO MA TRẬN ĐỀ + ĐỀ ÔN TẬP (TUẦN 9) & ĐỀ KIỂM TRA (TUẦN 10)
  MÔN TIN HỌC – LỚP 3 ĐẾN LỚP 8 – ĐÁNH GIÁ ĐỊNH KỲ 1
  NĂM HỌC 2026 – 2027
  ─────────────────────────────────────────────────────────────────
  Tuần 9: Ma trận + Đề ôn tập (~40 câu TN + 5 câu TL) 
  Tuần 10: Đề kiểm tra (20 câu TN + 2 câu thực hành)
═══════════════════════════════════════════════════════════════════════
"""
import sys, os, copy
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ═══════════════════ PATHS ═══════════════════
TEMPLATE_THCS = r'D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx'
TEMPLATE_TH = r'D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx'  # Same template
OUT_DIR = r'D:\UNIGO\KHBD_Tin_học'

# ═══════════════════ HELPER FUNCTIONS ═══════════════════
def set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def set_no_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="none"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:bottom w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'  <w:insideH w:val="none"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_font(run, name="Times New Roman", size_pt=13, bold=False, italic=False, color_rgb=None):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{name}" w:hAnsi="{name}" w:cs="{name}"/>')
    rPr.append(rFonts)

def add_para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False,
             size_pt=13, space_before=0, space_after=4, color_rgb=None, line_spacing=1.15, keep_with_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    if text:
        run = p.add_run(text)
        set_font(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
        return p, run
    return p, None

def add_question_options(doc, opt_a, opt_b, opt_c, opt_d, size_pt=12):
    """
    Căn chỉnh đáp án trắc nghiệm A, B, C, D đều đẹp chuẩn văn bản đề thi:
    - 4 cột (1 dòng): nếu tất cả đáp án ngắn (<= 15 ký tự)
    - 2 cột (2 dòng): nếu độ dài vừa phải (<= 35 ký tự và opt_a, opt_c <= 28 ký tự),
      cột 1 (A, C) căn lề trái 0.25 inch, cột 2 (B, D) căn tab stop 3.55 inch thẳng hàng tuyệt đối.
    - 1 cột (4 dòng): nếu có đáp án dài (> 35 ký tự hoặc opt_a/opt_c > 28 ký tự),
      mỗi đáp án 1 dòng riêng thụt lề 0.25 inch thẳng hàng.
    Giữ các dòng trong cùng 1 câu hỏi không bị ngắt trang giữa chừng (keep_with_next).
    """
    opts = [opt_a, opt_b, opt_c, opt_d]
    max_len = max(len(o) for o in opts)

    if max_len <= 15:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.add_tab_stop(Inches(0.25), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(1.85), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.45), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(5.05), WD_TAB_ALIGNMENT.LEFT)
        r = p.add_run(f"\t{opt_a}\t{opt_b}\t{opt_c}\t{opt_d}")
        set_font(r, size_pt=size_pt)
    elif max_len <= 35 and max(len(opt_a), len(opt_c)) <= 28:
        for idx, (o1, o2) in enumerate([(opt_a, opt_b), (opt_c, opt_d)]):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1 if idx == 0 else 2.5)
            p.paragraph_format.line_spacing = 1.15
            if idx == 0:
                p.paragraph_format.keep_with_next = True
            p.paragraph_format.tab_stops.add_tab_stop(Inches(0.25), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(3.55), WD_TAB_ALIGNMENT.LEFT)
            r = p.add_run(f"\t{o1}\t{o2}")
            set_font(r, size_pt=size_pt)
    else:
        for idx, opt in enumerate(opts):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1 if idx < 3 else 2.5)
            p.paragraph_format.line_spacing = 1.15
            if idx < 3:
                p.paragraph_format.keep_with_next = True
            p.paragraph_format.left_indent = Inches(0.25)
            r = p.add_run(opt)
            set_font(r, size_pt=size_pt)

def format_cell_para(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, italic=False, size_pt=12, color_rgb=None):
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    set_font(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
    return run

def add_cell_para(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, italic=False, size_pt=12, color_rgb=None):
    p = cell.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    set_font(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
    return run

def merge_cells(table, row1, col1, row2, col2):
    cell_start = table.cell(row1, col1)
    cell_end = table.cell(row2, col2)
    cell_start.merge(cell_end)

def clean_body(doc):
    """Giữ header/footer (sectPr), xóa toàn bộ body content."""
    body = doc.element.body
    for child in list(body):
        if not child.tag.endswith('sectPr'):
            body.remove(child)

def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)

def build_header_table(doc, lop, de_type="ÔN TẬP", thoi_gian="45 phút"):
    """Tạo bảng header chuẩn UNIGO cho đề thi."""
    is_thcs = int(lop) >= 6
    to_name = "TỔ CHUYÊN MÔN THCS" if is_thcs else "TỔ CHUYÊN MÔN TIỂU HỌC"
    
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_l, cell_r = tbl.rows[0].cells
    cell_l.width = Inches(3.2)
    cell_r.width = Inches(3.8)

    # Left cell
    format_cell_para(cell_l, f"TRƯỜNG TIỂU HỌC & THCS UNIGO", bold=True, size_pt=12)
    add_cell_para(cell_l, to_name, bold=True, size_pt=12)

    # Right cell
    if de_type == "ÔN TẬP":
        format_cell_para(cell_r, f"ĐỀ CƯƠNG & BÀI TẬP ÔN TẬP", bold=True, size_pt=12)
        add_cell_para(cell_r, "ĐÁNH GIÁ ĐỊNH KỲ 1", bold=True, size_pt=12)
    else:
        format_cell_para(cell_r, f"ĐỀ KIỂM TRA ĐÁNH GIÁ ĐỊNH KỲ 1", bold=True, size_pt=12)
    add_cell_para(cell_r, f"NĂM HỌC: 2026 - 2027 | MÔN: TIN HỌC {lop}", bold=True, size_pt=12)
    add_cell_para(cell_r, f"(Thời gian làm bài: {thoi_gian})", italic=True, size_pt=11)

    for cell in [cell_l, cell_r]:
        set_cell_borders(cell, "none", "none", "none", "none")

    # Divider
    add_para(doc, "─" * 70, align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10, color_rgb=(100, 100, 100), space_after=2)

def build_student_info(doc, lop):
    """Bảng thông tin học sinh."""
    tbl = doc.add_table(rows=2, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="666666")

    format_cell_para(tbl.cell(0, 0), "Họ và tên: .........................................................", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12)
    format_cell_para(tbl.cell(0, 1), f"Lớp: {lop}A1", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12)
    format_cell_para(tbl.cell(0, 2), "Điểm", bold=True, size_pt=12)
    format_cell_para(tbl.cell(0, 3), "Nhận xét của giáo viên", bold=True, size_pt=12)
    set_cell_background(tbl.cell(0, 2), "F2F4F8")
    set_cell_background(tbl.cell(0, 3), "F2F4F8")

    format_cell_para(tbl.cell(1, 0), "Ngày kiểm tra: ......./......./2026", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12)
    format_cell_para(tbl.cell(1, 1), "STT: .........", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12)
    format_cell_para(tbl.cell(1, 2), "\n\n", size_pt=12)
    format_cell_para(tbl.cell(1, 3), "\n............................................................\n............................................................", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=11)

    for r in tbl.rows:
        for c in r.cells:
            set_cell_margins(c, 80, 80, 100, 100)

def build_matrix_table(doc, matrix_data, headers):
    """Tạo bảng ma trận đề."""
    n_rows = len(matrix_data) + 1
    n_cols = len(headers)
    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)

    # Header row
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        set_cell_background(cell, "D6E4F0")
        format_cell_para(cell, h, bold=True, size_pt=11)

    # Data rows
    for r_idx, row_data in enumerate(matrix_data, start=1):
        is_total = r_idx == len(matrix_data)
        for c_idx, val in enumerate(row_data):
            cell = tbl.cell(r_idx, c_idx)
            if is_total:
                set_cell_background(cell, "E8D5F5")
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            format_cell_para(cell, val, align=align, bold=(is_total or c_idx == 0), size_pt=11)

    for r in tbl.rows:
        for c in r.cells:
            set_cell_margins(c, 60, 60, 80, 80)

def build_answer_key_table(doc, questions, points_each=0.25):
    """Tạo bảng đáp án trắc nghiệm."""
    n_q = len(questions)
    # Split into rows of 10
    row_size = 10
    chunks = [questions[i:i+row_size] for i in range(0, n_q, row_size)]
    
    for chunk_idx, chunk in enumerate(chunks):
        n = len(chunk)
        tbl = doc.add_table(rows=2, cols=n)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl)

        start_num = chunk_idx * row_size + 1
        for col_idx, (_, _, _, _, _, ans) in enumerate(chunk):
            # Header
            cell_h = tbl.cell(0, col_idx)
            set_cell_background(cell_h, "D6E4F0")
            format_cell_para(cell_h, f"Câu {start_num + col_idx}", bold=True, size_pt=10)
            # Answer
            cell_a = tbl.cell(1, col_idx)
            format_cell_para(cell_a, ans, bold=True, size_pt=11, color_rgb=(180, 0, 0))
        
        for r in tbl.rows:
            for c in r.cells:
                set_cell_margins(c, 50, 50, 60, 60)
        
        if chunk_idx < len(chunks) - 1:
            add_para(doc, "", space_after=2)

def build_signature_table(doc):
    """Bảng ký duyệt cuối file."""
    add_para(doc, "", space_after=6)
    tbl = doc.add_table(rows=3, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_no_borders(tbl)
    
    titles = ["DUYỆT CỦA BAN GIÁM HIỆU", "DUYỆT CỦA TỔ CHUYÊN MÔN", "GIÁO VIÊN RA ĐỀ"]
    subtitles = ["(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"]
    names = ["", "", "Đậu Đình Nguyên"]
    
    for i in range(3):
        format_cell_para(tbl.cell(0, i), titles[i], bold=True, size_pt=12)
        format_cell_para(tbl.cell(1, i), subtitles[i], italic=True, size_pt=12)
        format_cell_para(tbl.cell(2, i), f"\n\n\n{names[i]}", bold=(i == 2), size_pt=12)


# ═══════════════════════════════════════════════════════════════════════
# DATA: NỘI DUNG KIẾN THỨC & CÂU HỎI CHO TỪNG LỚP
# ═══════════════════════════════════════════════════════════════════════

# ──────────────── LỚP 3 ────────────────
# Tuần 1-8: Bài 1 (Thông tin và quyết định), Bài 2 (Xử lí thông tin), 
#           Bài 3 (Máy tính và em), Bài 4 (Làm việc với máy tính - T1)
GRADE_3 = {
    'lop': '3',
    'ten_lop': 'Lớp 3',
    'noi_dung_on': [
        "Bài 1: Thông tin và quyết định – Khái niệm thông tin, vai trò thông tin trong đời sống, thông tin giúp con người ra quyết định đúng đắn.",
        "Bài 2: Xử lí thông tin – Quy trình xử lí thông tin (Thu nhận → Lưu trữ → Xử lí → Truyền), các giác quan thu nhận thông tin.",
        "Bài 3: Máy tính và em – Các bộ phận cơ bản của máy tính (màn hình, chuột, bàn phím, thân máy), phân biệt máy tính bàn và máy tính xách tay.",
        "Bài 4: Làm việc với máy tính – Tư thế ngồi đúng khi dùng máy tính, quy tắc an toàn, khởi động và tắt máy tính đúng cách.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Thông tin và quyết định", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Xử lí thông tin", "4 TN", "4 TN", "2 TN + 1 TL", "1 TL", "10 TN + 2 TL"],
        ["Bài 3. Máy tính và em", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 4. Làm việc với máy tính", "4 TN", "4 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["TỔNG", "18 câu", "14 câu", "7 TN + 4 TL", "1 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Thông tin và quyết định", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 2. Xử lí thông tin", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Máy tính và em", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 4. Làm việc với máy tính", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "-", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "8 câu (2.0 điểm)", "3 TN + 1 TH (2.75 điểm)", "1 TH (2.75 điểm)", "10.0đ"],
    ],
    # ĐỀ ÔN TẬP - 40 câu TN
    'on_tap_tn': [
        # Bài 1: Thông tin và quyết định (10 câu)
        ("Câu 1. Thông tin là gì?",
         "A. Là những đồ vật xung quanh chúng ta",
         "B. Là những hiểu biết về thế giới xung quanh và về chính bản thân mình",
         "C. Là các loại máy tính hiện đại",
         "D. Là sách vở và bút viết", "B"),
        ("Câu 2. Khi nghe thấy tiếng chuông trường reo, em biết được điều gì?",
         "A. Trời sắp mưa",
         "B. Đến giờ vào lớp hoặc ra chơi",
         "C. Có người gõ cửa",
         "D. Xe cứu thương đi qua", "B"),
        ("Câu 3. Biển báo giao thông hình tròn, viền đỏ, có hình người đi bộ gạch chéo cho em biết điều gì?",
         "A. Được phép đi bộ qua đường",
         "B. Cấm người đi bộ",
         "C. Đường dành cho xe máy",
         "D. Khu vực trường học", "B"),
        ("Câu 4. Em nhận biết thông tin bằng cách nào?",
         "A. Chỉ bằng mắt nhìn",
         "B. Chỉ bằng tai nghe",
         "C. Bằng các giác quan: mắt, tai, mũi, lưỡi, da",
         "D. Chỉ bằng tay sờ", "C"),
        ("Câu 5. Thông tin giúp con người làm được điều gì?",
         "A. Chạy nhanh hơn",
         "B. Ra quyết định đúng đắn",
         "C. Bơi giỏi hơn",
         "D. Hát hay hơn", "B"),
        ("Câu 6. Khi xem dự báo thời tiết, em biết trời sẽ mưa. Đây là ví dụ về điều gì?",
         "A. Thông tin giúp em ra quyết định mang áo mưa",
         "B. Thông tin không có ích gì",
         "C. Dự báo luôn luôn sai",
         "D. Em thích xem thời tiết cho vui", "A"),
        ("Câu 7. Khi nhìn thấy đèn giao thông màu đỏ, em cần làm gì?",
         "A. Đi nhanh qua đường",
         "B. Dừng lại chờ đèn xanh",
         "C. Rẽ sang đường khác",
         "D. Nhắm mắt đi tiếp", "B"),
        ("Câu 8. Trong các vật sau, vật nào KHÔNG phải là vật mang tin?",
         "A. Sách giáo khoa",
         "B. Tờ báo",
         "C. Chiếc bàn học",
         "D. Bảng thông báo", "C"),
        ("Câu 9. Bạn An đọc bảng thời khóa biểu và biết hôm nay có tiết Tin học. Bạn An đã thu nhận thông tin bằng giác quan nào?",
         "A. Tai (thính giác)",
         "B. Mắt (thị giác)",
         "C. Mũi (khứu giác)",
         "D. Tay (xúc giác)", "B"),
        ("Câu 10. Vì sao thông tin quan trọng trong cuộc sống?",
         "A. Vì thông tin giúp em chơi game",
         "B. Vì thông tin giúp con người hiểu biết và đưa ra quyết định phù hợp",
         "C. Vì thông tin là đồ chơi",
         "D. Vì thông tin chỉ có trên Internet", "B"),

        # Bài 2: Xử lí thông tin (10 câu)
        ("Câu 11. Quy trình xử lí thông tin gồm mấy bước?",
         "A. 2 bước",
         "B. 3 bước",
         "C. 4 bước: Thu nhận → Lưu trữ → Xử lí → Truyền",
         "D. 5 bước", "C"),
        ("Câu 12. Bước đầu tiên trong quy trình xử lí thông tin là gì?",
         "A. Truyền thông tin",
         "B. Xử lí thông tin",
         "C. Lưu trữ thông tin",
         "D. Thu nhận thông tin", "D"),
        ("Câu 13. Khi bạn Minh nghe cô giáo giảng bài, bạn Minh đang thực hiện bước nào?",
         "A. Thu nhận thông tin",
         "B. Truyền thông tin",
         "C. Xử lí thông tin",
         "D. Lưu trữ thông tin", "A"),
        ("Câu 14. Con người thu nhận thông tin bằng gì?",
         "A. Bằng máy tính",
         "B. Bằng điện thoại",
         "C. Bằng các giác quan",
         "D. Bằng sách vở", "C"),
        ("Câu 15. Não của con người thực hiện bước nào trong quy trình xử lí thông tin?",
         "A. Thu nhận",
         "B. Xử lí",
         "C. Lưu trữ",
         "D. Cả xử lí và lưu trữ", "D"),
        ("Câu 16. Máy tính xử lí thông tin nhanh hơn con người vì sao?",
         "A. Vì máy tính rất đẹp",
         "B. Vì máy tính có bộ xử lí (CPU) tính toán cực nhanh",
         "C. Vì máy tính dùng điện",
         "D. Vì máy tính có màn hình to", "B"),
        ("Câu 17. Khi em ghi bài vào vở, em đang thực hiện bước nào?",
         "A. Thu nhận thông tin",
         "B. Xử lí thông tin",
         "C. Lưu trữ thông tin",
         "D. Truyền thông tin", "C"),
        ("Câu 18. Khi em kể lại câu chuyện cho bạn nghe, em đang thực hiện bước nào?",
         "A. Thu nhận thông tin",
         "B. Xử lí thông tin",
         "C. Lưu trữ thông tin",
         "D. Truyền thông tin", "D"),
        ("Câu 19. Thông tin có thể được lưu trữ ở đâu?",
         "A. Chỉ trong sách vở",
         "B. Chỉ trong máy tính",
         "C. Trong sách, vở, USB, ổ cứng, đám mây và nhiều nơi khác",
         "D. Chỉ trên Internet", "C"),
        ("Câu 20. Ở bước 'Xử lí thông tin', con người sử dụng bộ phận nào để suy nghĩ?",
         "A. Tim",
         "B. Não",
         "C. Tay",
         "D. Mắt", "B"),

        # Bài 3: Máy tính và em (10 câu)
        ("Câu 21. Máy tính để bàn thường gồm những bộ phận chính nào?",
         "A. Màn hình, chuột, bàn phím, thân máy",
         "B. Chỉ có màn hình và chuột",
         "C. Chỉ có thân máy",
         "D. Chỉ có bàn phím và loa", "A"),
        ("Câu 22. Bộ phận nào của máy tính hiển thị hình ảnh, chữ viết?",
         "A. Thân máy",
         "B. Chuột",
         "C. Màn hình",
         "D. Bàn phím", "C"),
        ("Câu 23. Chuột máy tính dùng để làm gì?",
         "A. Để gõ chữ",
         "B. Để chỉ, chọn và di chuyển trên màn hình",
         "C. Để phát âm thanh",
         "D. Để sạc pin", "B"),
        ("Câu 24. Bàn phím máy tính dùng để làm gì?",
         "A. Để hiển thị hình ảnh",
         "B. Để nghe nhạc",
         "C. Để gõ chữ, số và ra lệnh cho máy tính",
         "D. Để chụp ảnh", "C"),
        ("Câu 25. Thân máy (CPU) chứa bộ phận nào quan trọng nhất?",
         "A. Loa phát nhạc",
         "B. Bộ xử lí – 'bộ não' của máy tính",
         "C. Màn hình phụ",
         "D. Pin dự phòng", "B"),
        ("Câu 26. Máy tính xách tay (laptop) khác máy tính để bàn ở điểm nào?",
         "A. Laptop không có bàn phím",
         "B. Laptop gọn nhẹ, dễ mang theo, tích hợp màn hình và bàn phím",
         "C. Laptop không có màn hình",
         "D. Laptop không thể kết nối Internet", "B"),
        ("Câu 27. Thiết bị nào sau đây KHÔNG phải là máy tính?",
         "A. Máy tính bảng (tablet)",
         "B. Điện thoại thông minh",
         "C. Bàn là điện",
         "D. Máy tính xách tay", "C"),
        ("Câu 28. Loa máy tính dùng để làm gì?",
         "A. Để gõ chữ",
         "B. Để phát ra âm thanh",
         "C. Để hiển thị hình ảnh",
         "D. Để di chuyển con trỏ", "B"),
        ("Câu 29. Máy tính có thể giúp em làm những việc gì?",
         "A. Chỉ chơi game",
         "B. Học tập, vẽ tranh, nghe nhạc, tìm kiếm thông tin và nhiều việc khác",
         "C. Chỉ xem phim",
         "D. Chỉ gõ văn bản", "B"),
        ("Câu 30. Webcam là thiết bị dùng để làm gì?",
         "A. Để gõ chữ",
         "B. Để chụp ảnh và quay video",
         "C. Để phát nhạc",
         "D. Để in giấy", "B"),

        # Bài 4: Làm việc với máy tính (10 câu)
        ("Câu 31. Tư thế ngồi đúng khi dùng máy tính là gì?",
         "A. Nằm trên giường, để máy tính trên bụng",
         "B. Ngồi thẳng lưng, mắt cách màn hình khoảng 50 cm, hai chân chạm đất",
         "C. Ngồi xổm trên ghế",
         "D. Cúi sát mặt vào màn hình", "B"),
        ("Câu 32. Khi dùng máy tính, em nên nghỉ mắt sau mỗi bao lâu?",
         "A. Sau mỗi 5 giờ",
         "B. Không cần nghỉ",
         "C. Sau mỗi 30-45 phút",
         "D. Sau mỗi 3 giờ", "C"),
        ("Câu 33. Để tắt máy tính đúng cách, em cần làm gì?",
         "A. Rút phích cắm điện ngay",
         "B. Vào menu Start → chọn Shut Down (Tắt máy)",
         "C. Bấm giữ nút nguồn thật lâu",
         "D. Đóng nắp laptop rồi bỏ đi", "B"),
        ("Câu 34. Khi sử dụng máy tính, em KHÔNG nên làm điều gì?",
         "A. Rửa tay sạch trước khi dùng",
         "B. Ăn uống gần máy tính và đổ nước lên bàn phím",
         "C. Ngồi đúng tư thế",
         "D. Nghỉ mắt thường xuyên", "B"),
        ("Câu 35. Vì sao em cần ngồi đúng tư thế khi dùng máy tính?",
         "A. Để máy tính chạy nhanh hơn",
         "B. Để bảo vệ sức khỏe mắt, cột sống và cổ tay",
         "C. Để máy tính không bị hỏng",
         "D. Để cô giáo khen", "B"),
        ("Câu 36. Để khởi động (bật) máy tính, em cần nhấn nút nào?",
         "A. Nút Reset",
         "B. Nút nguồn (Power)",
         "C. Nút Caps Lock",
         "D. Nút Enter", "B"),
        ("Câu 37. Khi đang dùng máy tính mà mất điện đột ngột, em nên làm gì?",
         "A. Bấm nút nguồn liên tục",
         "B. Bình tĩnh chờ có điện lại, sau đó khởi động máy bình thường",
         "C. Rút hết dây cáp",
         "D. Gõ bàn phím liên tục", "B"),
        ("Câu 38. Khoảng cách từ mắt đến màn hình nên là bao nhiêu?",
         "A. Khoảng 10 cm",
         "B. Khoảng 50 - 70 cm",
         "C. Khoảng 200 cm",
         "D. Càng gần càng tốt", "B"),
        ("Câu 39. Em cần giữ vệ sinh máy tính bằng cách nào?",
         "A. Dùng nước lau màn hình",
         "B. Dùng khăn mềm, khô lau bụi nhẹ nhàng",
         "C. Dùng xà phòng rửa bàn phím",
         "D. Không cần vệ sinh", "B"),
        ("Câu 40. Theo quy tắc an toàn, em nên dùng máy tính bao lâu mỗi ngày?",
         "A. Cả ngày không nghỉ",
         "B. Sử dụng có giới hạn, tối đa 1-2 giờ mỗi ngày đối với trẻ em",
         "C. Chỉ 5 phút mỗi tuần",
         "D. Không giới hạn thời gian", "B"),
    ],
    # ĐỀ ÔN TẬP - 5 câu TL
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Em hãy kể tên 4 bộ phận chính của máy tính để bàn và nêu chức năng của mỗi bộ phận.",
         "- Màn hình: hiển thị hình ảnh, chữ viết (0,5 điểm)\n- Thân máy: chứa bộ xử lí và bộ nhớ (0,5 điểm)\n- Bàn phím: gõ chữ, số, ra lệnh (0,5 điểm)\n- Chuột: chỉ, chọn, di chuyển con trỏ (0,5 điểm)"),
        ("Câu 2 (2,0 điểm)", "Em hãy nêu 4 bước trong quy trình xử lí thông tin. Cho ví dụ minh họa cho từng bước.",
         "- Thu nhận: mắt nhìn, tai nghe (0,5 điểm)\n- Lưu trữ: ghi vào vở, nhớ trong đầu (0,5 điểm)\n- Xử lí: não suy nghĩ, so sánh (0,5 điểm)\n- Truyền: nói cho người khác, viết thư (0,5 điểm)"),
        ("Câu 3 (2,0 điểm)", "Em hãy nêu 4 quy tắc an toàn khi sử dụng máy tính.",
         "- Ngồi đúng tư thế, lưng thẳng (0,5 điểm)\n- Mắt cách màn hình 50-70 cm (0,5 điểm)\n- Nghỉ mắt sau mỗi 30-45 phút (0,5 điểm)\n- Không ăn uống gần máy tính (0,5 điểm)"),
        ("Câu 4 (2,0 điểm)", "Hãy cho 3 ví dụ về thông tin giúp con người ra quyết định trong cuộc sống hàng ngày.",
         "- Nghe dự báo thời tiết → mang áo mưa (0,7 điểm)\n- Xem bảng giờ tàu → đến ga đúng giờ (0,7 điểm)\n- Đọc nhãn thực phẩm → chọn đồ ăn phù hợp (0,6 điểm)"),
        ("Câu 5 (2,0 điểm)", "Vì sao máy tính được coi là công cụ xử lí thông tin vượt trội? Em hãy so sánh giữa máy tính và con người trong việc xử lí thông tin.",
         "- Máy tính tính toán nhanh hơn con người rất nhiều (0,5 điểm)\n- Máy tính có thể lưu trữ lượng lớn thông tin (0,5 điểm)\n- Máy tính làm việc không mệt mỏi (0,5 điểm)\n- Nhưng con người có thể sáng tạo, suy luận linh hoạt hơn máy tính (0,5 điểm)"),
    ],
    # ĐỀ KIỂM TRA - 20 câu TN
    'kiem_tra_tn': [
        ("Câu 1. Thông tin là gì?",
         "A. Là các loại máy móc", "B. Là những hiểu biết về thế giới xung quanh",
         "C. Là sách vở", "D. Là đồ chơi", "B"),
        ("Câu 2. Em thu nhận thông tin bằng cách nào?",
         "A. Chỉ bằng mắt", "B. Chỉ bằng tai",
         "C. Bằng các giác quan", "D. Chỉ bằng mũi", "C"),
        ("Câu 3. Khi nghe tiếng chuông báo hết giờ ra chơi, em biết cần phải làm gì?",
         "A. Tiếp tục chơi", "B. Vào lớp học",
         "C. Về nhà", "D. Đi ngủ", "B"),
        ("Câu 4. Bước đầu tiên trong quy trình xử lí thông tin là gì?",
         "A. Lưu trữ", "B. Xử lí",
         "C. Truyền", "D. Thu nhận", "D"),
        ("Câu 5. Quy trình xử lí thông tin có mấy bước?",
         "A. 2 bước", "B. 3 bước",
         "C. 4 bước", "D. 5 bước", "C"),
        ("Câu 6. Bộ phận nào hiển thị hình ảnh trên máy tính?",
         "A. Chuột", "B. Bàn phím",
         "C. Màn hình", "D. Loa", "C"),
        ("Câu 7. Bàn phím dùng để làm gì?",
         "A. Phát âm thanh", "B. Hiển thị hình ảnh",
         "C. Gõ chữ, số và ra lệnh", "D. Chụp ảnh", "C"),
        ("Câu 8. Thân máy chứa bộ phận quan trọng nào?",
         "A. Loa phát nhạc", "B. Bộ xử lí (bộ não máy tính)",
         "C. Màn hình phụ", "D. Pin sạc", "B"),
        ("Câu 9. Máy tính xách tay khác máy tính bàn ở điểm nào?",
         "A. Không có màn hình", "B. Gọn nhẹ, dễ mang theo",
         "C. Không thể gõ chữ", "D. Không có chuột", "B"),
        ("Câu 10. Thiết bị nào sau đây KHÔNG phải là máy tính?",
         "A. Laptop", "B. Máy tính bảng",
         "C. Tivi thường", "D. Điện thoại thông minh", "C"),
        ("Câu 11. Tư thế ngồi đúng khi dùng máy tính là gì?",
         "A. Nằm ngửa", "B. Ngồi thẳng lưng, mắt cách màn hình 50 cm",
         "C. Cúi sát màn hình", "D. Ngồi xổm", "B"),
        ("Câu 12. Em nên nghỉ mắt sau mỗi bao lâu khi dùng máy tính?",
         "A. 5 giờ", "B. 30-45 phút",
         "C. 3 giờ", "D. Không cần nghỉ", "B"),
        ("Câu 13. Để tắt máy tính đúng cách, em vào đâu?",
         "A. Rút dây điện", "B. Giữ nút nguồn",
         "C. Start → Shut Down", "D. Đóng nắp laptop", "C"),
        ("Câu 14. Em KHÔNG nên làm gì khi dùng máy tính?",
         "A. Rửa tay trước khi dùng", "B. Ăn uống gần máy tính",
         "C. Ngồi đúng tư thế", "D. Nghỉ mắt thường xuyên", "B"),
        ("Câu 15. Nút nào để bật máy tính?",
         "A. Enter", "B. Caps Lock",
         "C. Power (nút nguồn)", "D. Reset", "C"),
        ("Câu 16. Khi em ghi bài vào vở, em đang thực hiện bước nào?",
         "A. Thu nhận", "B. Xử lí",
         "C. Lưu trữ", "D. Truyền", "C"),
        ("Câu 17. Khi em kể chuyện cho bạn nghe, em đang thực hiện bước nào?",
         "A. Thu nhận", "B. Xử lí",
         "C. Lưu trữ", "D. Truyền thông tin", "D"),
        ("Câu 18. Não bộ con người tương ứng với bộ phận nào của máy tính?",
         "A. Màn hình", "B. Bàn phím",
         "C. Bộ xử lí (CPU)", "D. Chuột", "C"),
        ("Câu 19. Webcam dùng để làm gì?",
         "A. Gõ chữ", "B. Chụp ảnh và quay video",
         "C. Phát nhạc", "D. In giấy", "B"),
        ("Câu 20. Khoảng cách từ mắt đến màn hình máy tính nên là bao nhiêu?",
         "A. 10 cm", "B. 50-70 cm",
         "C. 200 cm", "D. Càng gần càng tốt", "B"),
    ],
    # ĐỀ KIỂM TRA - 2 câu thực hành
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành nhận diện", 
         "Em hãy quan sát máy tính trong phòng Tin học và chỉ ra (hoặc ghi tên) 4 bộ phận chính của máy tính để bàn. Với mỗi bộ phận, em hãy nêu 1 chức năng của nó.",
         "- Nhận diện đúng 4 bộ phận: Màn hình, Thân máy, Bàn phím, Chuột (1,0 điểm)\n- Nêu đúng chức năng mỗi bộ phận (1,0 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành thao tác",
         "Em hãy thực hiện các bước sau trên máy tính:\na) Khởi động máy tính đúng cách (1,0 điểm)\nb) Mở một chương trình bất kỳ (ví dụ: Paint) bằng chuột (1,0 điểm)\nc) Tắt máy tính đúng cách qua menu Bắt đầu (Start) rồi chọn Tắt máy (Shut Down) (1,0 điểm)",
         "a) Nhấn nút nguồn (Power), chờ màn hình làm việc xuất hiện (1,0 điểm)\nb) Nhấp đúp chuột vào biểu tượng Paint hoặc tìm qua menu Start (1,0 điểm)\nc) Nhấp Start rồi chọn Shut Down, chờ máy tính tắt hoàn toàn (1,0 điểm)"),
    ],
}

# ──────────────── LỚP 4 ────────────────
GRADE_4 = {
    'lop': '4',
    'ten_lop': 'Lớp 4',
    'noi_dung_on': [
        "Bài 1: Phần cứng và phần mềm máy tính – Phân biệt phần cứng (thiết bị vật lí) và phần mềm (chương trình), ví dụ thiết bị nhập, xuất, xử lí, lưu trữ.",
        "Bài 2: Gõ bàn phím đúng cách – Vị trí các ngón tay trên hàng phím cơ sở, quy tắc gõ 10 ngón, phím Home Row.",
        "Bài 3: Thông tin trên trang Web – Khái niệm trang web, trình duyệt, thanh địa chỉ, liên kết (hyperlink), các dạng thông tin trên web.",
        "Bài 4: Tìm kiếm thông tin trên Internet – Sử dụng công cụ tìm kiếm (Google), nhập từ khóa, chọn lọc kết quả phù hợp.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Phần cứng và phần mềm", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Gõ bàn phím đúng cách", "4 TN", "4 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 3. Thông tin trên trang Web", "5 TN", "3 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["Bài 4. Tìm kiếm trên Internet", "4 TN", "4 TN", "1 TN + 1 TL", "1 TN + 1 TL", "10 TN + 2 TL"],
        ["TỔNG", "18 câu", "14 câu", "6 TN + 4 TL", "2 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Phần cứng và phần mềm", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 2. Gõ bàn phím đúng cách", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Thông tin trên trang Web", "3 TN (0.75 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 4. Tìm kiếm trên Internet", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "7 câu (1.75 điểm)", "3 TN + 1 TH (2.5 điểm)", "1 TN + 1 TH (3.0 điểm)", "10.0đ"],
    ],
    'on_tap_tn': [
        # Bài 1: Phần cứng và phần mềm (10 câu)
        ("Câu 1. Phần cứng máy tính là gì?",
         "A. Là các chương trình cài đặt trong máy tính", "B. Là các bộ phận vật lí có thể nhìn thấy và chạm vào được",
         "C. Là hệ điều hành Windows", "D. Là các trò chơi trên máy tính", "B"),
        ("Câu 2. Phần mềm máy tính là gì?",
         "A. Là các thiết bị ngoại vi", "B. Là bàn phím và chuột",
         "C. Là các chương trình giúp máy tính hoạt động và thực hiện công việc", "D. Là dây cáp kết nối", "C"),
        ("Câu 3. Thiết bị nào sau đây là thiết bị nhập (Input)?",
         "A. Màn hình", "B. Loa",
         "C. Bàn phím", "D. Máy in", "C"),
        ("Câu 4. Thiết bị nào sau đây là thiết bị xuất (Output)?",
         "A. Chuột", "B. Bàn phím",
         "C. Máy quét (Scanner)", "D. Máy in (Printer)", "D"),
        ("Câu 5. CPU là viết tắt của cụm từ nào?",
         "A. Central Processing Unit", "B. Computer Personal Unit",
         "C. Central Power Unit", "D. Computer Program Utility", "A"),
        ("Câu 6. Ổ cứng (Hard Disk) thuộc loại thiết bị nào?",
         "A. Thiết bị nhập", "B. Thiết bị xuất",
         "C. Thiết bị lưu trữ", "D. Thiết bị xử lí", "C"),
        ("Câu 7. Phần mềm nào sau đây là phần mềm soạn thảo văn bản?",
         "A. Paint", "B. Microsoft Word",
         "C. Calculator", "D. Chrome", "B"),
        ("Câu 8. RAM là bộ nhớ dùng để làm gì?",
         "A. Lưu trữ dữ liệu vĩnh viễn", "B. Lưu trữ dữ liệu tạm thời khi máy đang hoạt động",
         "C. Hiển thị hình ảnh", "D. Kết nối Internet", "B"),
        ("Câu 9. Hệ điều hành Windows thuộc loại nào?",
         "A. Phần cứng", "B. Thiết bị ngoại vi",
         "C. Phần mềm hệ thống", "D. Thiết bị lưu trữ", "C"),
        ("Câu 10. Thiết bị nào vừa là thiết bị nhập vừa là thiết bị xuất?",
         "A. Bàn phím", "B. Loa",
         "C. Màn hình cảm ứng", "D. Máy in", "C"),

        # Bài 2: Gõ bàn phím đúng cách (10 câu)
        ("Câu 11. Hàng phím cơ sở (Home Row) gồm những phím nào?",
         "A. Q W E R T Y U I O P", "B. A S D F G H J K L ;",
         "C. Z X C V B N M", "D. 1 2 3 4 5 6 7 8 9 0", "B"),
        ("Câu 12. Khi gõ phím cơ sở, ngón trỏ tay trái đặt ở phím nào?",
         "A. Phím A", "B. Phím S",
         "C. Phím F", "D. Phím D", "C"),
        ("Câu 13. Khi gõ phím cơ sở, ngón trỏ tay phải đặt ở phím nào?",
         "A. Phím J", "B. Phím K",
         "C. Phím L", "D. Phím H", "A"),
        ("Câu 14. Phím nào dùng để viết chữ hoa?",
         "A. Enter", "B. Shift hoặc Caps Lock",
         "C. Tab", "D. Backspace", "B"),
        ("Câu 15. Phím Backspace dùng để làm gì?",
         "A. Xuống dòng mới", "B. Xóa kí tự phía trước con trỏ",
         "C. In đậm chữ", "D. Chuyển ngôn ngữ", "B"),
        ("Câu 16. Phím Enter dùng để làm gì?",
         "A. Xóa kí tự", "B. Viết chữ hoa",
         "C. Xuống dòng mới hoặc xác nhận lệnh", "D. Tắt máy tính", "C"),
        ("Câu 17. Phím Space dùng để làm gì?",
         "A. Xóa chữ", "B. Tạo khoảng trắng giữa các từ",
         "C. In đậm", "D. Gạch chân", "B"),
        ("Câu 18. Khi gõ 10 ngón, hai ngón cái dùng để gõ phím nào?",
         "A. Phím Shift", "B. Phím Enter",
         "C. Phím Space (phím cách)", "D. Phím Tab", "C"),
        ("Câu 19. Lợi ích của việc gõ bàn phím đúng cách là gì?",
         "A. Gõ nhanh hơn và ít mỏi tay", "B. Máy tính chạy nhanh hơn",
         "C. Màn hình sáng hơn", "D. Internet nhanh hơn", "A"),
        ("Câu 20. Khi gõ phím, em nên nhìn vào đâu?",
         "A. Nhìn vào bàn phím liên tục", "B. Nhìn vào màn hình",
         "C. Nhìn vào tay", "D. Nhắm mắt", "B"),

        # Bài 3: Thông tin trên trang Web (10 câu)
        ("Câu 21. Trang web là gì?",
         "A. Là một cuốn sách in trên giấy", "B. Là trang thông tin được hiển thị trên trình duyệt web",
         "C. Là một phần mềm cài trong máy tính", "D. Là một loại máy tính", "B"),
        ("Câu 22. Trình duyệt web nào sau đây phổ biến nhất?",
         "A. Microsoft Word", "B. Google Chrome",
         "C. Paint", "D. Calculator", "B"),
        ("Câu 23. Thanh địa chỉ trên trình duyệt dùng để làm gì?",
         "A. Gõ văn bản", "B. Nhập địa chỉ trang web cần truy cập",
         "C. Vẽ tranh", "D. Tính toán", "B"),
        ("Câu 24. Liên kết (Hyperlink) trên trang web có đặc điểm gì?",
         "A. Luôn có màu đen, không gạch chân", "B. Thường có màu xanh, gạch chân, click vào sẽ chuyển sang trang khác",
         "C. Không thể nhấn vào được", "D. Là hình ảnh không có chữ", "B"),
        ("Câu 25. Thông tin trên trang web có những dạng nào?",
         "A. Chỉ có chữ", "B. Chỉ có hình ảnh",
         "C. Văn bản, hình ảnh, âm thanh, video", "D. Chỉ có video", "C"),
        ("Câu 26. Để truy cập trang web, em cần gì?",
         "A. Chỉ cần có máy tính", "B. Máy tính hoặc điện thoại có kết nối Internet và trình duyệt web",
         "C. Chỉ cần có bàn phím", "D. Chỉ cần có loa", "B"),
        ("Câu 27. Website của trường học thường có địa chỉ kết thúc bằng gì?",
         "A. .com", "B. .edu.vn",
         "C. .xyz", "D. .game", "B"),
        ("Câu 28. Khi truy cập trang web, biểu tượng ổ khóa trên thanh địa chỉ nghĩa là gì?",
         "A. Trang web bị khóa, không truy cập được", "B. Kết nối an toàn, thông tin được bảo mật",
         "C. Trang web có virus", "D. Máy tính bị hỏng", "B"),
        ("Câu 29. Nút Back (←) trên trình duyệt dùng để làm gì?",
         "A. Tắt trình duyệt", "B. Quay lại trang web trước đó",
         "C. Tải trang web nhanh hơn", "D. In trang web", "B"),
        ("Câu 30. Tab trên trình duyệt cho phép em làm gì?",
         "A. Mở nhiều trang web cùng lúc", "B. Tắt Internet",
         "C. Xóa lịch sử duyệt web", "D. Tải game", "A"),

        # Bài 4: Tìm kiếm thông tin trên Internet (10 câu)
        ("Câu 31. Công cụ tìm kiếm phổ biến nhất hiện nay là gì?",
         "A. Yahoo", "B. Google",
         "C. Bing", "D. Ask", "B"),
        ("Câu 32. Để tìm kiếm thông tin, em gõ gì vào ô tìm kiếm?",
         "A. Địa chỉ nhà em", "B. Từ khóa liên quan đến nội dung cần tìm",
         "C. Số điện thoại", "D. Mật khẩu email", "B"),
        ("Câu 33. Khi tìm kiếm 'Thủ đô của Việt Nam', kết quả đúng nhất là gì?",
         "A. TP. Hồ Chí Minh", "B. Đà Nẵng",
         "C. Hà Nội", "D. Huế", "C"),
        ("Câu 34. Để tìm chính xác một cụm từ trên Google, em nên làm gì?",
         "A. Gõ từng chữ một", "B. Đặt cụm từ trong dấu ngoặc kép \" \"",
         "C. Gõ bằng chữ hoa", "D. Bấm Enter nhiều lần", "B"),
        ("Câu 35. Kết quả tìm kiếm trên Google hiển thị những gì?",
         "A. Chỉ có hình ảnh", "B. Tiêu đề trang web, mô tả ngắn và đường liên kết",
         "C. Chỉ có video", "D. Chỉ có số liệu", "B"),
        ("Câu 36. Từ khóa nào phù hợp khi tìm kiếm về động vật quý hiếm ở Việt Nam?",
         "A. 'Game hay nhất 2026'", "B. 'Động vật quý hiếm Việt Nam'",
         "C. 'Thời tiết hôm nay'", "D. 'Công thức nấu phở'", "B"),
        ("Câu 37. Khi tìm kiếm thông tin, em nên chọn kết quả từ nguồn nào?",
         "A. Trang web lạ, không rõ tác giả", "B. Nguồn uy tín như Wikipedia, báo chính thống, trang .edu",
         "C. Trang có nhiều quảng cáo nhấp nháy", "D. Trang yêu cầu tải phần mềm lạ", "B"),
        ("Câu 38. Tìm kiếm bằng hình ảnh trên Google gọi là gì?",
         "A. Google Maps", "B. Google Images (Tìm kiếm hình ảnh)",
         "C. Google Drive", "D. Google Mail", "B"),
        ("Câu 39. Vì sao em cần kiểm tra thông tin từ nhiều nguồn khi tìm kiếm?",
         "A. Để mất thời gian", "B. Để chắc chắn thông tin chính xác và đáng tin cậy",
         "C. Để máy tính chạy nhanh hơn", "D. Để có nhiều trang web hơn", "B"),
        ("Câu 40. Khi gặp trang web yêu cầu nhập thông tin cá nhân, em nên làm gì?",
         "A. Nhập ngay mọi thông tin", "B. Hỏi ý kiến thầy cô hoặc phụ huynh trước khi nhập",
         "C. Tắt máy tính", "D. Chia sẻ link cho bạn bè cùng nhập", "B"),
    ],
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Em hãy phân biệt phần cứng và phần mềm máy tính. Cho 3 ví dụ về phần cứng và 3 ví dụ về phần mềm.",
         "- Phần cứng: bộ phận vật lí (màn hình, bàn phím, chuột) (1,0 điểm)\n- Phần mềm: chương trình (Word, Chrome, Paint) (1,0 điểm)"),
        ("Câu 2 (2,0 điểm)", "Em hãy nêu vị trí đặt các ngón tay trên hàng phím cơ sở khi gõ 10 ngón. Vì sao việc gõ đúng cách lại quan trọng?",
         "- Tay trái: A-S-D-F, tay phải: J-K-L-; (1,0 điểm)\n- Gõ đúng giúp nhanh hơn, ít mỏi, không cần nhìn phím (1,0 điểm)"),
        ("Câu 3 (2,0 điểm)", "Em hãy mô tả các bước để tìm kiếm thông tin trên Google. Khi nào em cần dùng dấu ngoặc kép?",
         "- Mở trình duyệt → Vào google.com → Gõ từ khóa → Nhấn Enter → Chọn kết quả phù hợp (1,0 điểm)\n- Dùng ngoặc kép khi cần tìm chính xác cụm từ (1,0 điểm)"),
        ("Câu 4 (2,0 điểm)", "Em hãy giải thích liên kết (hyperlink) là gì. Khi nhấn vào liên kết trên trang web thì điều gì xảy ra?",
         "- Hyperlink: chữ/hình có thể nhấn vào để chuyển sang trang khác (1,0 điểm)\n- Khi nhấn: trình duyệt mở trang web mới tương ứng (1,0 điểm)"),
        ("Câu 5 (2,0 điểm)", "Khi tìm thấy một bài viết hay trên Internet, em cần lưu ý gì về bản quyền và độ tin cậy?",
         "- Kiểm tra tác giả, nguồn gốc uy tín (1,0 điểm)\n- Ghi rõ nguồn khi sử dụng, không sao chép nguyên văn (1,0 điểm)"),
    ],
    'kiem_tra_tn': [
        ("Câu 1. Phần cứng máy tính là gì?",
         "A. Chương trình máy tính", "B. Các bộ phận vật lí có thể nhìn thấy",
         "C. Hệ điều hành", "D. Trang web", "B"),
        ("Câu 2. Phần mềm máy tính là gì?",
         "A. Thiết bị ngoại vi", "B. Dây cáp kết nối",
         "C. Chương trình giúp máy tính hoạt động", "D. Ổ cứng", "C"),
        ("Câu 3. Bàn phím thuộc loại thiết bị gì?",
         "A. Thiết bị xuất", "B. Thiết bị nhập",
         "C. Thiết bị lưu trữ", "D. Thiết bị xử lí", "B"),
        ("Câu 4. Máy in thuộc loại thiết bị gì?",
         "A. Thiết bị nhập", "B. Thiết bị xuất",
         "C. Thiết bị xử lí", "D. Thiết bị lưu trữ", "B"),
        ("Câu 5. CPU đóng vai trò gì trong máy tính?",
         "A. Hiển thị hình ảnh", "B. Xử lí dữ liệu (bộ não máy tính)",
         "C. Phát âm thanh", "D. Kết nối Internet", "B"),
        ("Câu 6. Ngón trỏ tay trái đặt ở phím nào khi gõ 10 ngón?",
         "A. Phím A", "B. Phím D",
         "C. Phím F", "D. Phím S", "C"),
        ("Câu 7. Ngón trỏ tay phải đặt ở phím nào?",
         "A. Phím H", "B. Phím J",
         "C. Phím K", "D. Phím L", "B"),
        ("Câu 8. Phím Space dùng để làm gì?",
         "A. Xóa chữ", "B. Xuống dòng",
         "C. Tạo khoảng trắng", "D. Viết chữ hoa", "C"),
        ("Câu 9. Trang web là gì?",
         "A. Sách in trên giấy", "B. Trang thông tin hiển thị trên trình duyệt",
         "C. Phần mềm cài đặt", "D. Ổ cứng máy tính", "B"),
        ("Câu 10. Trình duyệt web phổ biến là gì?",
         "A. Word", "B. Excel",
         "C. Chrome", "D. Paint", "C"),
        ("Câu 11. Thanh địa chỉ dùng để làm gì?",
         "A. Gõ văn bản", "B. Nhập địa chỉ trang web",
         "C. Tính toán", "D. Vẽ tranh", "B"),
        ("Câu 12. Hyperlink là gì?",
         "A. Loại phần cứng", "B. Liên kết dẫn đến trang web khác",
         "C. Loại virus", "D. Phần mềm diệt virus", "B"),
        ("Câu 13. Công cụ tìm kiếm phổ biến nhất là gì?",
         "A. Yahoo", "B. Google",
         "C. Bing", "D. DuckDuckGo", "B"),
        ("Câu 14. Từ khóa tìm kiếm là gì?",
         "A. Mật khẩu đăng nhập", "B. Từ hoặc cụm từ liên quan đến nội dung cần tìm",
         "C. Tên đăng nhập", "D. Số điện thoại", "B"),
        ("Câu 15. Phím Backspace dùng để làm gì?",
         "A. Xuống dòng", "B. Xóa kí tự phía trước con trỏ",
         "C. In đậm", "D. Tắt máy", "B"),
        ("Câu 16. RAM dùng để làm gì?",
         "A. Lưu vĩnh viễn", "B. Lưu tạm thời khi máy hoạt động",
         "C. Hiển thị hình", "D. Kết nối mạng", "B"),
        ("Câu 17. Ổ cứng dùng để làm gì?",
         "A. Xử lí dữ liệu", "B. Hiển thị hình ảnh",
         "C. Lưu trữ dữ liệu lâu dài", "D. Phát âm thanh", "C"),
        ("Câu 18. Để tìm chính xác cụm từ, em đặt trong dấu gì?",
         "A. Ngoặc đơn ( )", "B. Ngoặc kép \" \"",
         "C. Ngoặc vuông [ ]", "D. Gạch chéo / /", "B"),
        ("Câu 19. Nguồn thông tin nào đáng tin cậy nhất?",
         "A. Trang web lạ", "B. Tin đồn trên mạng",
         "C. Trang .edu, .gov, báo chính thống", "D. Bài đăng ẩn danh", "C"),
        ("Câu 20. Hệ điều hành Windows là loại gì?",
         "A. Phần cứng", "B. Phần mềm hệ thống",
         "C. Thiết bị ngoại vi", "D. Thiết bị lưu trữ", "B"),
    ],
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành gõ bàn phím",
         "Em hãy thực hiện các yêu cầu sau:\na) Mở phần mềm soạn thảo văn bản (Word hoặc Notepad) (0,5 điểm)\nb) Gõ đoạn văn sau bằng cách đặt tay đúng vị trí hàng phím cơ sở:\n   \"Em yêu trường Tiểu học và THCS UNIGO. Em sẽ cố gắng học tập thật tốt.\" (1,0 điểm)\nc) Lưu tệp với tên 'BaiTap_HoTen.docx' (0,5 điểm)",
         "a) Mở đúng phần mềm soạn thảo (0,5 điểm)\nb) Gõ đúng nội dung, đúng tư thế (1,0 điểm)\nc) Lưu tệp đúng tên (0,5 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành tìm kiếm",
         "Em hãy thực hiện các bước sau:\na) Mở trình duyệt web Google Chrome (0,5 điểm)\nb) Tìm kiếm trên Google với từ khóa: \"Các danh lam thắng cảnh nổi tiếng ở Việt Nam\" (1,0 điểm)\nc) Chọn một kết quả từ nguồn đáng tin cậy và ghi lại tiêu đề trang web cùng với ba danh lam thắng cảnh tìm được (1,5 điểm)",
         "a) Mở đúng trình duyệt Google Chrome (0,5 điểm)\nb) Nhập đúng từ khóa tìm kiếm và nhấn phím Enter (1,0 điểm)\nc) Chọn nguồn thông tin uy tín, ghi đúng ba thắng cảnh (1,5 điểm)"),
    ],
}

# ──────────────── LỚP 5 ────────────────
GRADE_5 = {
    'lop': '5',
    'ten_lop': 'Lớp 5',
    'noi_dung_on': [
        "Bài 1: Em có thể làm gì với máy tính? – Vai trò của máy tính trong học tập, giải trí, giao tiếp; ứng dụng thực tế.",
        "Bài 2: Tìm kiếm thông tin trên website – Sử dụng công cụ tìm kiếm, đánh giá nguồn thông tin, từ khóa hiệu quả.",
        "Bài 3: Tìm kiếm thông tin trong giải quyết vấn đề – Xác định vấn đề, tìm kiếm thông tin phù hợp, đưa ra giải pháp.",
        "Bài 4: Cây thư mục – Khái niệm tệp, thư mục, cây thư mục, cách tổ chức dữ liệu, tạo/đổi tên/xóa thư mục.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Em làm gì với máy tính?", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Tìm kiếm trên website", "4 TN", "4 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["Bài 3. Tìm kiếm GQ vấn đề", "5 TN", "3 TN", "1 TN + 1 TL", "1 TN + 1 TL", "10 TN + 2 TL"],
        ["Bài 4. Cây thư mục", "4 TN", "4 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["TỔNG", "18 câu", "14 câu", "6 TN + 4 TL", "2 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Em làm gì với máy tính?", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 2. Tìm kiếm trên website", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Tìm kiếm GQ vấn đề", "3 TN (0.75 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 4. Cây thư mục", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "7 câu (1.75 điểm)", "4 TN + 1 TH (2.75 điểm)", "1 TH (2.75 điểm)", "10.0đ"],
    ],
    'on_tap_tn': [
        # Bài 1 (10 câu)
        ("Câu 1. Máy tính có thể giúp em làm gì trong học tập?",
         "A. Chỉ chơi game", "B. Soạn bài, tìm kiếm tài liệu, làm bài trình chiếu",
         "C. Chỉ xem phim", "D. Chỉ nghe nhạc", "B"),
        ("Câu 2. Ứng dụng nào giúp em soạn thảo văn bản?",
         "A. Paint", "B. Microsoft Word",
         "C. Calculator", "D. Notepad chỉ gõ được text thuần", "B"),
        ("Câu 3. Máy tính giúp con người giao tiếp qua phương tiện nào?",
         "A. Thư tay", "B. Email, chat, video call",
         "C. Chỉ gọi điện thoại bàn", "D. Chỉ viết thư", "B"),
        ("Câu 4. Phần mềm trình chiếu phổ biến là gì?",
         "A. Microsoft Excel", "B. Microsoft PowerPoint",
         "C. Microsoft Access", "D. Paint", "B"),
        ("Câu 5. Máy tính có thể giúp bác sĩ làm gì?",
         "A. Chỉ chơi game", "B. Chụp X-quang, lưu trữ hồ sơ bệnh nhân, hỗ trợ phẫu thuật",
         "C. Chỉ gõ văn bản", "D. Chỉ nghe nhạc", "B"),
        ("Câu 6. Máy tính có thể thay thế hoàn toàn con người không?",
         "A. Có, máy tính làm được mọi thứ", "B. Không, máy tính là công cụ hỗ trợ, con người vẫn cần sáng tạo và ra quyết định",
         "C. Có, nếu là máy tính đời mới", "D. Máy tính không giúp ích gì", "B"),
        ("Câu 7. Phần mềm vẽ tranh trên máy tính là gì?",
         "A. Word", "B. Paint / Canva",
         "C. Excel", "D. Chrome", "B"),
        ("Câu 8. Internet giúp em làm gì?",
         "A. Chỉ tải game", "B. Tìm kiếm thông tin, học trực tuyến, giao tiếp",
         "C. Chỉ xem quảng cáo", "D. Chỉ đọc truyện", "B"),
        ("Câu 9. Để tạo một bài báo cáo, em dùng phần mềm nào?",
         "A. Paint", "B. Word hoặc Google Docs",
         "C. Calculator", "D. Camera", "B"),
        ("Câu 10. Lưu trữ đám mây (Cloud) có nghĩa là gì?",
         "A. Lưu dữ liệu trên mây thật", "B. Lưu dữ liệu trên máy chủ Internet, truy cập từ mọi nơi",
         "C. Lưu dữ liệu trên USB", "D. Lưu dữ liệu trên giấy", "B"),
        # Bài 2 (10 câu)
        ("Câu 11. Để tìm kiếm thông tin trên website, em cần sử dụng gì?",
         "A. Máy in", "B. Công cụ tìm kiếm (Google, Bing)",
         "C. Máy quét", "D. Loa", "B"),
        ("Câu 12. Từ khóa tìm kiếm hiệu quả cần đảm bảo yêu cầu gì?",
         "A. Càng dài càng tốt", "B. Ngắn gọn, chính xác, liên quan đến nội dung cần tìm",
         "C. Viết bằng tiếng nước ngoài", "D. Gõ ngẫu nhiên", "B"),
        ("Câu 13. Nguồn thông tin nào trên Internet đáng tin cậy?",
         "A. Bài đăng ẩn danh trên mạng xã hội", "B. Trang .edu, .gov, báo chính thống, Wikipedia",
         "C. Video quảng cáo sản phẩm", "D. Tin nhắn chuyển tiếp trên Zalo", "B"),
        ("Câu 14. Khi tìm kiếm thấy nhiều kết quả, em nên làm gì?",
         "A. Chọn kết quả đầu tiên mà không đọc", "B. Đọc tiêu đề và mô tả, chọn kết quả phù hợp nhất từ nguồn uy tín",
         "C. Tải hết tất cả", "D. Bỏ qua tất cả", "B"),
        ("Câu 15. Google Images dùng để làm gì?",
         "A. Gửi email", "B. Tìm kiếm hình ảnh",
         "C. Soạn văn bản", "D. Tạo bảng tính", "B"),
        ("Câu 16. Khi tìm kiếm \"Hồ Gươm\", dấu ngoặc kép giúp gì?",
         "A. Xóa kết quả", "B. Tìm chính xác cụm từ \"Hồ Gươm\"",
         "C. Mở trang web mới", "D. Tắt Internet", "B"),
        ("Câu 17. Phím tắt Ctrl + F dùng để làm gì trên trang web?",
         "A. Tắt trình duyệt", "B. Tìm kiếm từ khóa trong trang web đang mở",
         "C. In trang web", "D. Sao chép trang web", "B"),
        ("Câu 18. Bookmark (Đánh dấu trang) dùng để làm gì?",
         "A. Xóa trang web", "B. Lưu lại địa chỉ trang web yêu thích để truy cập nhanh sau này",
         "C. Tải file", "D. Chụp ảnh màn hình", "B"),
        ("Câu 19. Em không nên làm gì khi tìm kiếm trên Internet?",
         "A. Dùng từ khóa rõ ràng", "B. Nhấp vào quảng cáo lạ hoặc trang web yêu cầu tải phần mềm bất thường",
         "C. Chọn nguồn uy tín", "D. Kiểm tra thông tin từ nhiều nguồn", "B"),
        ("Câu 20. Lịch sử duyệt web (History) lưu lại điều gì?",
         "A. Mật khẩu email", "B. Danh sách các trang web em đã truy cập",
         "C. File tải về", "D. Số điện thoại", "B"),
        # Bài 3 (10 câu)
        ("Câu 21. Bước đầu tiên khi giải quyết vấn đề bằng tìm kiếm thông tin là gì?",
         "A. Tải phần mềm", "B. Xác định rõ vấn đề cần giải quyết",
         "C. Mở game", "D. Tắt máy tính", "B"),
        ("Câu 22. Khi cần tìm hiểu về thời tiết cho chuyến đi dã ngoại, em nên tìm ở đâu?",
         "A. Trang bán hàng online", "B. Trang dự báo thời tiết (weather.com.vn)",
         "C. Trang game", "D. Trang mạng xã hội", "B"),
        ("Câu 23. Quy trình giải quyết vấn đề bằng tìm kiếm gồm những bước nào?",
         "A. Chỉ gõ Google rồi đọc", "B. Xác định vấn đề → Tìm kiếm → Đánh giá → Tổng hợp → Giải quyết",
         "C. Chỉ hỏi bạn bè", "D. Chỉ đọc sách", "B"),
        ("Câu 24. Khi thông tin tìm được mâu thuẫn nhau, em nên làm gì?",
         "A. Tin tất cả", "B. So sánh nhiều nguồn và chọn nguồn uy tín nhất",
         "C. Bỏ qua tất cả", "D. Tự bịa đáp án", "B"),
        ("Câu 25. Tìm kiếm thông tin giúp em phát triển kỹ năng gì?",
         "A. Chỉ kỹ năng gõ phím", "B. Tư duy phản biện, đánh giá và giải quyết vấn đề",
         "C. Chỉ kỹ năng vẽ", "D. Chỉ kỹ năng hát", "B"),
        ("Câu 26. Em cần làm bài thuyết trình về 'Bảo vệ môi trường'. Bước đầu tiên em nên làm gì?",
         "A. Mở PowerPoint ngay", "B. Xác định nội dung cần trình bày và tìm kiếm tài liệu liên quan",
         "C. Vẽ tranh", "D. In ấn", "B"),
        ("Câu 27. Khi tìm được thông tin hữu ích, em nên làm gì?",
         "A. Sao chép nguyên văn", "B. Ghi chú lại nguồn và tóm tắt nội dung chính",
         "C. Xóa đi", "D. Gửi cho tất cả bạn", "B"),
        ("Câu 28. Trích dẫn nguồn tham khảo nghĩa là gì?",
         "A. Xóa tên tác giả", "B. Ghi rõ tên tác giả và đường liên kết khi sử dụng thông tin",
         "C. Đổi tên tác giả", "D. Tự nhận là của mình", "B"),
        ("Câu 29. Sử dụng thông tin trên mạng cần tuân thủ quy tắc gì?",
         "A. Sao chép thoải mái", "B. Tôn trọng bản quyền, trích dẫn nguồn, kiểm chứng thông tin",
         "C. Không cần ghi nguồn", "D. Tin tất cả thông tin trên mạng", "B"),
        ("Câu 30. Kỹ năng tìm kiếm thông tin quan trọng vì sao?",
         "A. Chỉ để chơi game", "B. Giúp em học tập hiệu quả, giải quyết vấn đề trong cuộc sống",
         "C. Chỉ để xem phim", "D. Không quan trọng", "B"),
        # Bài 4 (10 câu)
        ("Câu 31. Tệp (File) là gì?",
         "A. Là thư mục", "B. Là đơn vị lưu trữ dữ liệu trên máy tính (văn bản, hình ảnh, nhạc...)",
         "C. Là ổ cứng", "D. Là màn hình", "B"),
        ("Câu 32. Thư mục (Folder) dùng để làm gì?",
         "A. Hiển thị hình ảnh", "B. Chứa và sắp xếp các tệp cho gọn gàng",
         "C. Phát nhạc", "D. Kết nối Internet", "B"),
        ("Câu 33. Cây thư mục là gì?",
         "A. Một loại cây trong sân trường", "B. Cách tổ chức thư mục theo cấp bậc: thư mục mẹ chứa thư mục con",
         "C. Một trò chơi", "D. Một phần mềm", "B"),
        ("Câu 34. Đuôi tệp .docx cho biết tệp đó là loại gì?",
         "A. Hình ảnh", "B. Tệp văn bản Word",
         "C. Tệp nhạc", "D. Tệp video", "B"),
        ("Câu 35. Đuôi tệp .jpg hoặc .png cho biết tệp đó là loại gì?",
         "A. Văn bản", "B. Hình ảnh",
         "C. Âm thanh", "D. Video", "B"),
        ("Câu 36. Để tạo thư mục mới, em nhấp chuột phải và chọn gì?",
         "A. Delete", "B. New → Folder",
         "C. Copy", "D. Rename", "B"),
        ("Câu 37. Để đổi tên tệp hoặc thư mục, em nhấp chuột phải và chọn gì?",
         "A. Delete", "B. Copy",
         "C. Rename", "D. New", "C"),
        ("Câu 38. Thùng rác (Recycle Bin) dùng để làm gì?",
         "A. Lưu trữ vĩnh viễn", "B. Chứa tạm các tệp/thư mục đã xóa, có thể khôi phục",
         "C. Tạo tệp mới", "D. Kết nối Internet", "B"),
        ("Câu 39. Thư mục gốc trong ổ đĩa C gọi là gì?",
         "A. Thư mục con", "B. Thư mục mẹ (thư mục gốc)",
         "C. Tệp tin", "D. Ổ USB", "B"),
        ("Câu 40. Vì sao cần tổ chức tệp và thư mục gọn gàng?",
         "A. Để máy tính đẹp hơn", "B. Để dễ tìm kiếm và quản lí dữ liệu",
         "C. Để máy tính chạy nhanh hơn", "D. Để ổ cứng rộng hơn", "B"),
    ],
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Em hãy nêu 4 ứng dụng của máy tính trong đời sống hàng ngày. Mỗi ứng dụng cho 1 ví dụ cụ thể.",
         "- Học tập: soạn bài, tìm tài liệu (0,5 điểm)\n- Giải trí: xem phim, nghe nhạc (0,5 điểm)\n- Giao tiếp: email, video call (0,5 điểm)\n- Công việc: thiết kế, tính toán (0,5 điểm)"),
        ("Câu 2 (2,0 điểm)", "Em hãy mô tả cây thư mục để sắp xếp dữ liệu học tập, gồm thư mục gốc 'HOC_TAP' và 3 thư mục con cho 3 môn học.",
         "- Thư mục gốc: HOC_TAP (0,5 điểm)\n- Thư mục con: TOAN, TIENG_VIET, TIN_HOC (0,5 điểm)\n- Mỗi thư mục con chứa: Bai_tap, Tai_lieu (1,0 điểm)"),
        ("Câu 3 (2,0 điểm)", "Khi cần tìm hiểu về 'Tết Nguyên Đán ở Việt Nam' để làm bài thuyết trình, em sẽ thực hiện quy trình tìm kiếm thông tin như thế nào?",
         "- Xác định vấn đề: cần tìm nguồn gốc, phong tục, ý nghĩa (0,5 điểm)\n- Tìm kiếm trên Google với từ khóa phù hợp (0,5 điểm)\n- Chọn nguồn uy tín, ghi chú nội dung (0,5 điểm)\n- Tổng hợp và trình bày (0,5 điểm)"),
        ("Câu 4 (2,0 điểm)", "Em hãy phân biệt tệp và thư mục. Cho 3 ví dụ về tên tệp kèm đuôi mở rộng.",
         "- Tệp: đơn vị lưu trữ dữ liệu (văn bản, ảnh, nhạc) (0,5 điểm)\n- Thư mục: nơi chứa và sắp xếp các tệp (0,5 điểm)\n- Ví dụ: BaiTap.docx, AnhLop.jpg, BaiHat.mp3 (1,0 điểm)"),
        ("Câu 5 (2,0 điểm)", "Giải thích vì sao cần kiểm tra độ tin cậy của thông tin trên Internet. Nêu 3 cách kiểm tra.",
         "- Vì Internet có cả thông tin sai, giả (0,5 điểm)\n- Kiểm tra tác giả/nguồn gốc (0,5 điểm)\n- So sánh nhiều nguồn (0,5 điểm)\n- Chọn trang uy tín (.edu, .gov) (0,5 điểm)"),
    ],
    'kiem_tra_tn': [
        ("Câu 1. Máy tính giúp em làm gì trong học tập?",
         "A. Chỉ chơi game", "B. Soạn bài, tìm tài liệu, làm bài trình chiếu",
         "C. Chỉ xem phim", "D. Chỉ nghe nhạc", "B"),
        ("Câu 2. Phần mềm trình chiếu phổ biến là gì?",
         "A. Excel", "B. PowerPoint",
         "C. Paint", "D. Calculator", "B"),
        ("Câu 3. Internet giúp em giao tiếp qua gì?",
         "A. Chỉ thư tay", "B. Email, chat, video call",
         "C. Chỉ điện thoại bàn", "D. Chỉ tin nhắn SMS", "B"),
        ("Câu 4. Để tìm kiếm trên website, em dùng gì?",
         "A. Máy in", "B. Công cụ tìm kiếm Google",
         "C. Máy quét", "D. Loa", "B"),
        ("Câu 5. Từ khóa tìm kiếm cần đảm bảo gì?",
         "A. Càng dài càng tốt", "B. Ngắn gọn, chính xác, liên quan nội dung",
         "C. Bằng tiếng Anh", "D. Gõ ngẫu nhiên", "B"),
        ("Câu 6. Nguồn thông tin nào đáng tin cậy?",
         "A. Bài đăng ẩn danh", "B. Trang .edu, .gov, Wikipedia",
         "C. Tin đồn Zalo", "D. Video quảng cáo", "B"),
        ("Câu 7. Bước đầu tiên giải quyết vấn đề bằng tìm kiếm là gì?",
         "A. Tải phần mềm", "B. Xác định rõ vấn đề cần giải quyết",
         "C. Mở game", "D. Tắt máy", "B"),
        ("Câu 8. Khi thông tin mâu thuẫn, em nên làm gì?",
         "A. Tin tất cả", "B. So sánh nhiều nguồn, chọn nguồn uy tín",
         "C. Bỏ qua", "D. Tự bịa", "B"),
        ("Câu 9. Trích dẫn nguồn nghĩa là gì?",
         "A. Xóa tên tác giả", "B. Ghi rõ tên tác giả và nguồn khi sử dụng thông tin",
         "C. Đổi tên tác giả", "D. Tự nhận là của mình", "B"),
        ("Câu 10. Tệp (File) là gì?",
         "A. Thư mục", "B. Đơn vị lưu trữ dữ liệu trên máy tính",
         "C. Ổ cứng", "D. Màn hình", "B"),
        ("Câu 11. Thư mục (Folder) dùng để làm gì?",
         "A. Phát nhạc", "B. Chứa và sắp xếp các tệp",
         "C. Hiển thị hình", "D. Kết nối mạng", "B"),
        ("Câu 12. Cây thư mục là gì?",
         "A. Cây trong sân trường", "B. Tổ chức thư mục theo cấp bậc mẹ-con",
         "C. Trò chơi", "D. Phần mềm", "B"),
        ("Câu 13. Đuôi .docx là loại tệp gì?",
         "A. Hình ảnh", "B. Văn bản Word",
         "C. Nhạc", "D. Video", "B"),
        ("Câu 14. Đuôi .jpg là loại tệp gì?",
         "A. Văn bản", "B. Hình ảnh",
         "C. Âm thanh", "D. Video", "B"),
        ("Câu 15. Tạo thư mục mới: chuột phải → chọn gì?",
         "A. Delete", "B. New → Folder",
         "C. Copy", "D. Rename", "B"),
        ("Câu 16. Đổi tên tệp: chuột phải → chọn gì?",
         "A. Delete", "B. Copy",
         "C. Rename", "D. New", "C"),
        ("Câu 17. Recycle Bin dùng để làm gì?",
         "A. Lưu vĩnh viễn", "B. Chứa tạm tệp đã xóa, có thể khôi phục",
         "C. Tạo tệp", "D. Kết nối mạng", "B"),
        ("Câu 18. Máy tính thay thế hoàn toàn con người không?",
         "A. Có", "B. Không, máy tính là công cụ hỗ trợ",
         "C. Có nếu đời mới", "D. Máy tính vô ích", "B"),
        ("Câu 19. Cloud (đám mây) là gì?",
         "A. Mây trên trời", "B. Lưu dữ liệu trên máy chủ Internet",
         "C. USB", "D. Ổ cứng", "B"),
        ("Câu 20. Bookmark dùng để làm gì?",
         "A. Xóa trang web", "B. Lưu trang web yêu thích",
         "C. Tải file", "D. Chụp ảnh", "B"),
    ],
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành tổ chức thư mục",
         "Em hãy thực hiện trên máy tính:\na) Tạo thư mục 'HOC_TAP' trên Desktop (0,5 điểm)\nb) Trong thư mục 'HOC_TAP', tạo 3 thư mục con: 'TOAN', 'TIENG_VIET', 'TIN_HOC' (0,75 điểm)\nc) Tạo 1 tệp văn bản tên 'BaiTap1.docx' trong thư mục 'TIN_HOC' (0,75 điểm)",
         "a) Tạo đúng thư mục HOC_TAP trên Desktop (0,5 điểm)\nb) Tạo đúng 3 thư mục con (0,75 điểm)\nc) Tạo đúng tệp trong thư mục TIN_HOC (0,75 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành tìm kiếm và giải quyết vấn đề",
         "Tình huống: Em cần tìm hiểu về '5 loài động vật quý hiếm ở Việt Nam' để làm bài thuyết trình.\na) Mở trình duyệt Google Chrome, tìm kiếm với từ khóa phù hợp (1,0 điểm)\nb) Chọn một kết quả từ nguồn đáng tin cậy, ghi lại tên năm loài động vật (1,0 điểm)\nc) Ghi rõ nguồn tham khảo gồm tên trang web và đường liên kết (1,0 điểm)",
         "a) Mở Google Chrome, nhập từ khóa tìm kiếm hợp lí (1,0 điểm)\nb) Chọn nguồn đáng tin cậy, ghi đúng năm loài động vật (1,0 điểm)\nc) Ghi đầy đủ nguồn tham khảo (1,0 điểm)"),
    ],
}

# ──────────────── LỚP 6 ────────────────
GRADE_6 = {
    'lop': '6',
    'ten_lop': 'Lớp 6',
    'noi_dung_on': [
        "Bài 1: Thông tin và dữ liệu – Khái niệm thông tin, dữ liệu, vật mang tin, các dạng thông tin cơ bản, tầm quan trọng.",
        "Bài 2: Xử lí thông tin – Quy trình 4 bước: Thu nhận → Lưu trữ → Xử lí → Truyền; vai trò CPU.",
        "Bài 3: Thông tin trong máy tính – Khái niệm bit, biểu diễn thông tin thành dãy bit; đơn vị đo B, KB, MB, GB, TB.",
        "Bài 4: Mạng máy tính – Khái niệm mạng LAN, WAN, các thiết bị mạng (Router, Switch), lợi ích mạng máy tính.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Thông tin và dữ liệu", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Xử lí thông tin", "4 TN", "4 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["Bài 3. Thông tin trong MT", "5 TN", "3 TN", "1 TN + 1 TL", "1 TN + 1 TL", "10 TN + 2 TL"],
        ["Bài 4. Mạng máy tính", "4 TN", "4 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["TỔNG", "18 câu", "14 câu", "6 TN + 4 TL", "2 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Thông tin và dữ liệu", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 2. Xử lí thông tin", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Thông tin trong MT", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 4. Mạng máy tính", "3 TN (0.75 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "7 câu (1.75 điểm)", "3 TN + 1 TH (2.5 điểm)", "1 TN + 1 TH (3.0 điểm)", "10.0đ"],
    ],
    'on_tap_tn': [
        # Bài 1 (10)
        ("Câu 1. Thông tin là gì?", "A. Là thiết bị máy tính", "B. Là những hiểu biết về thế giới xung quanh và bản thân", "C. Là phần mềm", "D. Là ổ cứng", "B"),
        ("Câu 2. Dữ liệu là gì?", "A. Là thông tin được ghi lên vật mang tin dưới dạng cụ thể", "B. Là máy tính", "C. Là Internet", "D. Là phần cứng", "A"),
        ("Câu 3. Vật mang tin là gì?", "A. Là thông tin", "B. Là phương tiện chứa dữ liệu (sách, USB, đĩa CD)", "C. Là phần mềm", "D. Là bàn phím", "B"),
        ("Câu 4. Thông tin có mấy dạng cơ bản?", "A. 1 dạng", "B. 2 dạng", "C. 3 dạng: văn bản, hình ảnh, âm thanh", "D. 5 dạng", "C"),
        ("Câu 5. Ví dụ nào thể hiện thông tin dạng hình ảnh?", "A. Bài hát", "B. Tấm bản đồ", "C. Bài văn", "D. Đoạn hội thoại", "B"),
        ("Câu 6. Khi đọc bảng điểm, em tiếp nhận thông tin dạng gì?", "A. Âm thanh", "B. Hình ảnh", "C. Văn bản (số liệu)", "D. Video", "C"),
        ("Câu 7. Thông tin giúp con người làm gì?", "A. Chơi game", "B. Ra quyết định chính xác", "C. Tăng dung lượng máy", "D. Làm đẹp máy tính", "B"),
        ("Câu 8. USB là ví dụ của gì?", "A. Thông tin", "B. Vật mang tin", "C. Phần mềm", "D. Internet", "B"),
        ("Câu 9. Dữ liệu và thông tin khác nhau thế nào?", "A. Giống nhau hoàn toàn", "B. Dữ liệu là thông tin được biểu diễn cụ thể trên vật mang tin", "C. Thông tin nằm trong dữ liệu", "D. Không liên quan", "B"),
        ("Câu 10. Biển báo cấm hút thuốc mang thông tin dạng gì?", "A. Văn bản", "B. Âm thanh", "C. Hình ảnh (kí hiệu)", "D. Video", "C"),
        # Bài 2 (10)
        ("Câu 11. Quy trình xử lí thông tin gồm mấy bước?", "A. 2", "B. 3", "C. 4: Thu nhận → Lưu trữ → Xử lí → Truyền", "D. 5", "C"),
        ("Câu 12. Bộ phận nào của máy tính xử lí thông tin?", "A. Màn hình", "B. Bàn phím", "C. CPU (bộ vi xử lí)", "D. Loa", "C"),
        ("Câu 13. Bước 'Thu nhận' thông tin ở con người thực hiện qua gì?", "A. Tay chân", "B. Các giác quan", "C. Ổ cứng", "D. Chuột máy tính", "B"),
        ("Câu 14. Thiết bị nào là thiết bị nhập liệu?", "A. Màn hình", "B. Bàn phím", "C. Loa", "D. Máy in", "B"),
        ("Câu 15. Thiết bị nào là thiết bị xuất dữ liệu?", "A. Chuột", "B. Bàn phím", "C. Máy in", "D. Webcam", "C"),
        ("Câu 16. Ổ cứng trong máy tính thực hiện bước nào?", "A. Thu nhận", "B. Xử lí", "C. Lưu trữ", "D. Truyền", "C"),
        ("Câu 17. CPU viết tắt của cụm từ nào?", "A. Central Processing Unit", "B. Computer Personal Unit", "C. Central Power Unit", "D. Computer Program Utility", "A"),
        ("Câu 18. Máy tính xử lí thông tin vượt trội vì sao?", "A. Vì rất đẹp", "B. Vì tính toán nhanh, lưu trữ nhiều, không mệt", "C. Vì dùng điện", "D. Vì có màn hình to", "B"),
        ("Câu 19. Bước 'Truyền thông tin' thể hiện qua ví dụ nào?", "A. Ghi bài vào vở", "B. Gửi email cho bạn", "C. Đọc sách", "D. Nghĩ trong đầu", "B"),
        ("Câu 20. RAM dùng để làm gì?", "A. Lưu trữ lâu dài", "B. Lưu trữ tạm thời khi máy hoạt động", "C. Hiển thị hình", "D. Kết nối mạng", "B"),
        # Bài 3 (10)
        ("Câu 21. Bit là gì?", "A. Đơn vị đo trọng lượng", "B. Đơn vị nhỏ nhất của thông tin trong máy tính, có giá trị 0 hoặc 1", "C. Loại virus", "D. Tên phần mềm", "B"),
        ("Câu 22. 1 Byte bằng bao nhiêu bit?", "A. 4 bit", "B. 8 bit", "C. 16 bit", "D. 32 bit", "B"),
        ("Câu 23. 1 KB bằng bao nhiêu Byte?", "A. 100 Byte", "B. 512 Byte", "C. 1024 Byte", "D. 2048 Byte", "C"),
        ("Câu 24. Thứ tự đúng của đơn vị đo dung lượng từ nhỏ đến lớn là gì?", "A. MB < KB < GB < TB", "B. B < KB < MB < GB < TB", "C. TB < GB < MB < KB", "D. KB < B < MB < GB", "B"),
        ("Câu 25. USB 4GB có thể lưu trữ khoảng bao nhiêu bức ảnh 2MB?", "A. 500", "B. 1000", "C. 2000", "D. 4000", "C"),
        ("Câu 26. Thông tin trong máy tính được biểu diễn dưới dạng gì?", "A. Chữ viết tay", "B. Dãy bit (0 và 1)", "C. Hình vẽ", "D. Âm thanh tự nhiên", "B"),
        ("Câu 27. 1 MB bằng bao nhiêu KB?", "A. 100 KB", "B. 512 KB", "C. 1024 KB", "D. 2048 KB", "C"),
        ("Câu 28. 1 GB bằng bao nhiêu MB?", "A. 100 MB", "B. 512 MB", "C. 1024 MB", "D. 2048 MB", "C"),
        ("Câu 29. Ổ cứng 500 GB lưu được khoảng bao nhiêu bộ phim 2 GB?", "A. 100", "B. 200", "C. 250", "D. 500", "C"),
        ("Câu 30. Chữ 'A' trong bảng mã ASCII được biểu diễn bằng dãy bit nào?", "A. 0100 0001", "B. 1111 1111", "C. 0000 0000", "D. 1010 1010", "A"),
        # Bài 4 (10)
        ("Câu 31. Mạng máy tính là gì?", "A. Một máy tính đơn lẻ", "B. Hệ thống các máy tính được kết nối với nhau để chia sẻ tài nguyên", "C. Một loại phần mềm", "D. Một loại ổ cứng", "B"),
        ("Câu 32. Mạng LAN là gì?", "A. Mạng toàn cầu", "B. Mạng cục bộ, phạm vi nhỏ (phòng, tòa nhà)", "C. Mạng vệ tinh", "D. Mạng điện thoại", "B"),
        ("Câu 33. Mạng WAN là gì?", "A. Mạng cục bộ", "B. Mạng diện rộng, kết nối nhiều LAN ở xa nhau", "C. Mạng Bluetooth", "D. Mạng trong 1 phòng", "B"),
        ("Câu 34. Router (bộ định tuyến) dùng để làm gì?", "A. Hiển thị hình ảnh", "B. Kết nối và điều hướng dữ liệu giữa các mạng", "C. Gõ văn bản", "D. Phát nhạc", "B"),
        ("Câu 35. Switch (bộ chuyển mạch) dùng để làm gì?", "A. Tắt máy tính", "B. Kết nối các máy tính trong cùng mạng LAN", "C. In tài liệu", "D. Quét virus", "B"),
        ("Câu 36. Lợi ích của mạng máy tính là gì?", "A. Chỉ để chơi game online", "B. Chia sẻ dữ liệu, máy in, Internet và làm việc nhóm", "C. Làm máy chạy chậm hơn", "D. Không có lợi ích gì", "B"),
        ("Câu 37. Cáp mạng (Ethernet) dùng để làm gì?", "A. Sạc pin", "B. Kết nối máy tính với mạng qua dây", "C. Phát Wi-Fi", "D. Hiển thị hình ảnh", "B"),
        ("Câu 38. Wi-Fi là gì?", "A. Dây cáp mạng", "B. Công nghệ kết nối mạng không dây", "C. Loại phần mềm", "D. Tên máy tính", "B"),
        ("Câu 39. Server (máy chủ) trong mạng dùng để làm gì?", "A. Chỉ để chơi game", "B. Cung cấp tài nguyên và dịch vụ cho các máy tính khác", "C. Chỉ để soạn văn bản", "D. Chỉ để xem phim", "B"),
        ("Câu 40. Mạng Internet thuộc loại mạng nào?", "A. LAN", "B. MAN", "C. WAN (mạng diện rộng toàn cầu)", "D. PAN", "C"),
    ],
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Em hãy phân biệt thông tin, dữ liệu và vật mang tin. Cho ví dụ minh họa cho mỗi khái niệm.",
         "- Thông tin: hiểu biết về thế giới xung quanh (0,5 điểm)\n- Dữ liệu: thông tin được ghi cụ thể trên vật mang tin (0,5 điểm)\n- Vật mang tin: phương tiện chứa dữ liệu (sách, USB) (0,5 điểm)\n- Ví dụ minh họa phù hợp (0,5 điểm)"),
        ("Câu 2 (2,0 điểm)", "Trình bày 4 bước trong quy trình xử lí thông tin. So sánh vai trò CPU và não người.",
         "- 4 bước: Thu nhận → Lưu trữ → Xử lí → Truyền (1,0 điểm)\n- CPU = bộ não máy tính, tính toán nhanh hơn nhưng không sáng tạo (1,0 điểm)"),
        ("Câu 3 (2,0 điểm)", "Giải thích khái niệm bit và quy đổi: 1 KB = ? Byte, 1 MB = ? KB, 1 GB = ? MB. Tính: USB 8 GB lưu được bao nhiêu bức ảnh 4 MB?",
         "- Bit: đơn vị nhỏ nhất (0 hoặc 1) (0,5 điểm)\n- 1 KB = 1024 B, 1 MB = 1024 KB, 1 GB = 1024 MB (0,5 điểm)\n- 8 GB = 8 × 1024 = 8192 MB → 8192 ÷ 4 = 2048 ảnh (1,0 điểm)"),
        ("Câu 4 (2,0 điểm)", "Phân biệt mạng LAN và WAN. Nêu 3 lợi ích của mạng máy tính trong trường học.",
         "- LAN: mạng cục bộ, phạm vi nhỏ (0,5 điểm)\n- WAN: mạng diện rộng, phạm vi lớn (0,5 điểm)\n- 3 lợi ích: chia sẻ file, dùng chung máy in, truy cập Internet (1,0 điểm)"),
        ("Câu 5 (2,0 điểm)", "Router và Switch khác nhau thế nào? Em hãy mô tả cách mạng LAN trong phòng Tin học của trường hoạt động.",
         "- Router: kết nối giữa các mạng, điều hướng (0,5 điểm)\n- Switch: kết nối các máy trong cùng LAN (0,5 điểm)\n- Mô tả: các máy → cáp → Switch → Router → Internet (1,0 điểm)"),
    ],
    'kiem_tra_tn': [
        ("Câu 1. Thông tin là gì?",
         "A. Là các thiết bị máy tính", "B. Là những hiểu biết về thế giới xung quanh và về chính bản thân mình",
         "C. Là các phần mềm cài trong máy tính", "D. Là ổ cứng lưu trữ dữ liệu", "B"),
        ("Câu 2. Dữ liệu là gì?",
         "A. Là một loại máy tính", "B. Là thông tin được ghi lại cụ thể trên vật mang tin",
         "C. Là mạng Internet", "D. Là linh kiện phần cứng", "B"),
        ("Câu 3. Trong các vật sau, vật nào là ví dụ của vật mang tin?",
         "A. Thông tin", "B. USB, sách, đĩa CD",
         "C. Phần mềm máy tính", "D. Bộ vi xử lí trung tâm", "B"),
        ("Câu 4. Thông tin cơ bản có mấy dạng?",
         "A. 1 dạng", "B. 2 dạng",
         "C. 3 dạng: văn bản, hình ảnh, âm thanh", "D. 5 dạng", "C"),
        ("Câu 5. Thông tin giúp con người làm được điều gì quan trọng?",
         "A. Chơi trò chơi điện tử", "B. Ra quyết định chính xác trong cuộc sống",
         "C. Tăng dung lượng máy tính", "D. Làm cho máy tính đẹp hơn", "B"),
        ("Câu 6. Quy trình xử lí thông tin gồm bao nhiêu bước?",
         "A. 2 bước", "B. 3 bước",
         "C. 4 bước: Thu nhận → Lưu trữ → Xử lí → Truyền", "D. 5 bước", "C"),
        ("Câu 7. Bộ vi xử lí trung tâm (CPU) trong máy tính dùng để làm gì?",
         "A. Hiển thị hình ảnh ra màn hình", "B. Xử lí và tính toán dữ liệu",
         "C. Phát ra âm thanh", "D. Lưu trữ dữ liệu lâu dài", "B"),
        ("Câu 8. Thiết bị nào sau đây là thiết bị nhập dữ liệu?",
         "A. Màn hình", "B. Bàn phím",
         "C. Loa", "D. Máy in", "B"),
        ("Câu 9. Bộ nhớ RAM trong máy tính dùng để làm gì?",
         "A. Lưu trữ dữ liệu lâu dài", "B. Lưu trữ dữ liệu tạm thời khi máy đang hoạt động",
         "C. Hiển thị hình ảnh", "D. Kết nối mạng Internet", "B"),
        ("Câu 10. Vì sao máy tính xử lí thông tin nhanh hơn con người?",
         "A. Vì máy tính rất đẹp", "B. Vì bộ vi xử lí (CPU) có tốc độ tính toán cực nhanh",
         "C. Vì máy tính sử dụng điện", "D. Vì máy tính có màn hình to", "B"),
        ("Câu 11. Bit là gì?",
         "A. Là đơn vị đo trọng lượng", "B. Là đơn vị nhỏ nhất của thông tin trong máy tính, có giá trị 0 hoặc 1",
         "C. Là một loại virus máy tính", "D. Là tên một phần mềm", "B"),
        ("Câu 12. 1 Byte bằng bao nhiêu bit?",
         "A. 4 bit", "B. 8 bit",
         "C. 16 bit", "D. 32 bit", "B"),
        ("Câu 13. 1 KB (Kilobyte) bằng bao nhiêu Byte?",
         "A. 100 Byte", "B. 512 Byte",
         "C. 1024 Byte", "D. 2048 Byte", "C"),
        ("Câu 14. Thứ tự đúng của các đơn vị đo dung lượng thông tin từ nhỏ đến lớn là gì?",
         "A. MB < KB < GB < TB", "B. B < KB < MB < GB < TB",
         "C. TB < GB < MB < KB", "D. KB < B < MB < GB", "B"),
        ("Câu 15. Mạng LAN (Local Area Network) là gì?",
         "A. Là mạng máy tính toàn cầu", "B. Là mạng cục bộ, kết nối các máy tính trong phạm vi nhỏ",
         "C. Là mạng vệ tinh", "D. Là mạng điện thoại di động", "B"),
        ("Câu 16. Bộ định tuyến (Router) dùng để làm gì?",
         "A. Hiển thị hình ảnh trên màn hình", "B. Kết nối và điều hướng dữ liệu giữa các mạng",
         "C. Soạn thảo văn bản", "D. Phát ra âm thanh", "B"),
        ("Câu 17. Wi-Fi là gì?",
         "A. Là một loại dây cáp mạng", "B. Là công nghệ kết nối mạng không dây",
         "C. Là một phần mềm máy tính", "D. Là tên một loại máy tính", "B"),
        ("Câu 18. Lợi ích của mạng máy tính là gì?",
         "A. Chỉ dùng để chơi trò chơi trực tuyến", "B. Chia sẻ dữ liệu, dùng chung máy in và làm việc nhóm hiệu quả",
         "C. Làm cho máy tính chạy chậm hơn", "D. Không có lợi ích gì", "B"),
        ("Câu 19. Internet thuộc loại mạng nào?",
         "A. Mạng LAN (cục bộ)", "B. Mạng MAN (đô thị)",
         "C. Mạng WAN (diện rộng toàn cầu)", "D. Mạng PAN (cá nhân)", "C"),
        ("Câu 20. Một USB có dung lượng 4 GB có thể lưu trữ được khoảng bao nhiêu bức ảnh có dung lượng 2 MB mỗi bức?",
         "A. Khoảng 500 bức ảnh", "B. Khoảng 1000 bức ảnh",
         "C. Khoảng 2000 bức ảnh", "D. Khoảng 4000 bức ảnh", "C"),
    ],
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành quy đổi đơn vị",
         "Em hãy tính toán và ghi kết quả:\na) 1 GB = ? MB = ? KB = ? Byte (1,0 điểm)\nb) Ổ cứng 256 GB có thể lưu tối đa bao nhiêu bộ phim có dung lượng 4 GB mỗi phim? (1,0 điểm)",
         "a) 1 GB = 1024 MB = 1.048.576 KB = 1.073.741.824 Byte (1,0 điểm)\nb) 256 ÷ 4 = 64 bộ phim (1,0 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành nhận diện mạng",
         "Em hãy quan sát phòng Tin học và thực hiện:\na) Chỉ ra thiết bị chuyển mạch (Switch), cáp mạng và bộ định tuyến (Router) (1,0 điểm)\nb) Giải thích cách các máy tính trong phòng kết nối với nhau thành mạng cục bộ (1,0 điểm)\nc) Kiểm tra máy tính có kết nối mạng Internet hay không bằng cách mở trình duyệt và truy cập google.com (1,0 điểm)",
         "a) Nhận diện đúng các thiết bị mạng trong phòng học (1,0 điểm)\nb) Giải thích đúng: máy tính kết nối cáp mạng đến Switch để chia sẻ dữ liệu (1,0 điểm)\nc) Mở trình duyệt web và truy cập thành công trang google.com (1,0 điểm)"),
    ],
}

# ──────────────── LỚP 7 ────────────────
GRADE_7 = {
    'lop': '7',
    'ten_lop': 'Lớp 7',
    'noi_dung_on': [
        "Bài 1: Thiết bị vào - ra – Phân loại thiết bị nhập (Input), thiết bị xuất (Output), thiết bị vừa nhập vừa xuất.",
        "Bài 2: Phần mềm máy tính – Phân loại phần mềm hệ thống và phần mềm ứng dụng; vai trò hệ điều hành.",
        "Bài 3: Quản lý dữ liệu trong máy tính – Tệp, thư mục, cây thư mục, đường dẫn, các thao tác quản lí tệp/thư mục.",
        "Bài 4: Mạng xã hội và kênh trao đổi thông tin – Đặc điểm mạng xã hội, kênh truyền thông, ưu nhược điểm, an toàn.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Thiết bị vào - ra", "5 TN", "3 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Phần mềm máy tính", "4 TN", "4 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["Bài 3. Quản lý dữ liệu", "5 TN", "3 TN", "1 TN + 1 TL", "1 TN + 1 TL", "10 TN + 2 TL"],
        ["Bài 4. Mạng xã hội", "4 TN", "4 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["TỔNG", "18 câu", "14 câu", "6 TN + 4 TL", "2 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Thiết bị vào - ra", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 2. Phần mềm máy tính", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Quản lý dữ liệu", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 4. Mạng xã hội", "3 TN (0.75 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "7 câu (1.75 điểm)", "3 TN + 1 TH (2.5 điểm)", "1 TN + 1 TH (3.0 điểm)", "10.0đ"],
    ],
    'on_tap_tn': [
        # Bài 1 (10)
        ("Câu 1. Thiết bị nhập (Input) dùng để làm gì?", "A. Hiển thị kết quả", "B. Đưa dữ liệu vào máy tính", "C. Phát âm thanh", "D. In tài liệu", "B"),
        ("Câu 2. Thiết bị xuất (Output) dùng để làm gì?", "A. Nhập dữ liệu", "B. Hiển thị hoặc đưa kết quả ra ngoài", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 3. Bàn phím thuộc loại thiết bị gì?", "A. Xuất", "B. Nhập", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 4. Màn hình thuộc loại thiết bị gì?", "A. Nhập", "B. Xuất", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 5. Thiết bị nào vừa nhập vừa xuất?", "A. Bàn phím", "B. Loa", "C. Màn hình cảm ứng", "D. Máy in", "C"),
        ("Câu 6. Webcam thuộc loại thiết bị gì?", "A. Xuất", "B. Nhập", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 7. Máy in thuộc loại thiết bị gì?", "A. Nhập", "B. Xuất", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 8. Micro thuộc loại thiết bị gì?", "A. Xuất", "B. Nhập", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 9. Loa thuộc loại thiết bị gì?", "A. Nhập", "B. Xuất", "C. Lưu trữ", "D. Xử lí", "B"),
        ("Câu 10. Máy quét (Scanner) thuộc loại thiết bị gì?", "A. Xuất", "B. Nhập", "C. Lưu trữ", "D. Xử lí", "B"),
        # Bài 2 (10)
        ("Câu 11. Phần mềm hệ thống là gì?", "A. Game", "B. Phần mềm quản lí và điều khiển phần cứng (hệ điều hành, driver)", "C. Word", "D. Chrome", "B"),
        ("Câu 12. Phần mềm ứng dụng là gì?", "A. Hệ điều hành", "B. Phần mềm phục vụ nhu cầu cụ thể (Word, Excel, Chrome)", "C. Driver", "D. BIOS", "B"),
        ("Câu 13. Windows là loại phần mềm gì?", "A. Ứng dụng", "B. Hệ thống (hệ điều hành)", "C. Game", "D. Tiện ích", "B"),
        ("Câu 14. Microsoft Word là loại phần mềm gì?", "A. Hệ thống", "B. Ứng dụng", "C. Driver", "D. BIOS", "B"),
        ("Câu 15. Hệ điều hành có vai trò gì?", "A. Soạn văn bản", "B. Quản lí phần cứng, phần mềm và cung cấp giao diện", "C. Tạo bảng tính", "D. Chỉ để chơi game", "B"),
        ("Câu 16. Phần mềm diệt virus thuộc loại gì?", "A. Hệ thống", "B. Ứng dụng (tiện ích)", "C. Game", "D. Hệ điều hành", "B"),
        ("Câu 17. Driver (trình điều khiển) dùng để làm gì?", "A. Soạn văn bản", "B. Giúp hệ điều hành giao tiếp với thiết bị phần cứng", "C. Duyệt web", "D. Chơi game", "B"),
        ("Câu 18. Google Chrome là loại phần mềm gì?", "A. Hệ thống", "B. Ứng dụng (trình duyệt web)", "C. Driver", "D. BIOS", "B"),
        ("Câu 19. Khi cài phần mềm mới, em nên tải từ đâu?", "A. Trang web lạ", "B. Trang chính thức của nhà phát triển hoặc cửa hàng ứng dụng uy tín", "C. Link người lạ gửi", "D. Email spam", "B"),
        ("Câu 20. Phần mềm có bản quyền nghĩa là gì?", "A. Miễn phí tùy ý", "B. Được bảo vệ pháp luật, cần mua hoặc xin phép sử dụng", "C. Không cần cài đặt", "D. Chỉ dùng 1 lần", "B"),
        # Bài 3 (10)
        ("Câu 21. Đường dẫn (path) là gì?", "A. Đường đi trên bản đồ", "B. Địa chỉ xác định vị trí tệp/thư mục trong cây thư mục", "C. Tên tệp", "D. Tên phần mềm", "B"),
        ("Câu 22. Đường dẫn C:\\Users\\HocSinh\\BaiTap.docx cho biết gì?", "A. BaiTap.docx nằm trong thư mục HocSinh trên ổ C:", "B. BaiTap.docx nằm trên Desktop", "C. BaiTap.docx là thư mục", "D. BaiTap.docx bị xóa", "A"),
        ("Câu 23. Phím tắt Ctrl+C dùng để làm gì?", "A. Dán", "B. Sao chép", "C. Cắt", "D. Xóa", "B"),
        ("Câu 24. Phím tắt Ctrl+V dùng để làm gì?", "A. Sao chép", "B. Dán", "C. Cắt", "D. Xóa", "B"),
        ("Câu 25. Phím tắt Ctrl+X dùng để làm gì?", "A. Sao chép", "B. Dán", "C. Cắt (di chuyển)", "D. Xóa", "C"),
        ("Câu 26. Phím Delete dùng để làm gì?", "A. Sao chép", "B. Xóa tệp/thư mục vào Recycle Bin", "C. Đổi tên", "D. Dán", "B"),
        ("Câu 27. Shift + Delete dùng để làm gì?", "A. Sao chép", "B. Xóa vĩnh viễn không qua Recycle Bin", "C. Đổi tên", "D. Dán", "B"),
        ("Câu 28. File Explorer dùng để làm gì?", "A. Duyệt web", "B. Quản lí tệp và thư mục trên máy tính", "C. Soạn văn bản", "D. Chơi game", "B"),
        ("Câu 29. Đuôi mở rộng .xlsx là loại tệp gì?", "A. Văn bản Word", "B. Bảng tính Excel", "C. Hình ảnh", "D. Video", "B"),
        ("Câu 30. Để tìm kiếm tệp trong File Explorer, em dùng gì?", "A. Thanh địa chỉ", "B. Ô Search (tìm kiếm) góc phải trên", "C. Nút Start", "D. Taskbar", "B"),
        # Bài 4 (10)
        ("Câu 31. Mạng xã hội phổ biến nào em biết?", "A. Word", "B. Facebook, Zalo, TikTok", "C. Excel", "D. Paint", "B"),
        ("Câu 32. Ưu điểm của mạng xã hội là gì?", "A. Gây nghiện", "B. Kết nối, chia sẻ thông tin, học tập nhanh chóng", "C. Mất thời gian", "D. Lộ thông tin cá nhân", "B"),
        ("Câu 33. Nhược điểm của mạng xã hội là gì?", "A. Kết nối bạn bè", "B. Nguy cơ nghiện, lộ thông tin, bắt nạt trực tuyến", "C. Học tập hiệu quả", "D. Tìm kiếm nhanh", "B"),
        ("Câu 34. Em KHÔNG nên làm gì trên mạng xã hội?", "A. Chia sẻ bài học hay", "B. Đăng ảnh, số điện thoại, địa chỉ nhà cá nhân", "C. Kết bạn với người quen", "D. Hỏi bài thầy cô", "B"),
        ("Câu 35. Bắt nạt trực tuyến (Cyberbullying) là gì?", "A. Giúp đỡ bạn trên mạng", "B. Hành vi đe dọa, xúc phạm, lan truyền tin xấu về người khác qua mạng", "C. Chia sẻ bài học", "D. Tìm kiếm thông tin", "B"),
        ("Câu 36. Khi bị bắt nạt trên mạng, em nên làm gì?", "A. Bắt nạt lại", "B. Báo ngay cho thầy cô, phụ huynh và chặn tài khoản đó", "C. Im lặng chịu đựng", "D. Xóa tài khoản", "B"),
        ("Câu 37. Email là gì?", "A. Mạng xã hội", "B. Thư điện tử gửi qua Internet", "C. Phần mềm game", "D. Trang web", "B"),
        ("Câu 38. Blog là gì?", "A. Mạng cục bộ", "B. Nhật kí trực tuyến, trang web cá nhân chia sẻ bài viết", "C. Phần mềm diệt virus", "D. Hệ điều hành", "B"),
        ("Câu 39. Kênh YouTube dùng để làm gì?", "A. Soạn văn bản", "B. Chia sẻ và xem video trực tuyến", "C. Tạo bảng tính", "D. Gửi email", "B"),
        ("Câu 40. Để bảo vệ tài khoản mạng xã hội, em nên làm gì?", "A. Dùng mật khẩu đơn giản '123456'", "B. Đặt mật khẩu mạnh, bật xác thực 2 bước, không chia sẻ mật khẩu", "C. Đăng mật khẩu lên Facebook", "D. Dùng chung tài khoản với bạn", "B"),
    ],
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Phân biệt thiết bị nhập và thiết bị xuất. Cho 3 ví dụ mỗi loại và 1 thiết bị vừa nhập vừa xuất.",
         "- Thiết bị nhập: bàn phím, chuột, webcam (0,75 điểm)\n- Thiết bị xuất: màn hình, loa, máy in (0,75 điểm)\n- Vừa nhập vừa xuất: màn hình cảm ứng (0,5 điểm)"),
        ("Câu 2 (2,0 điểm)", "Phân biệt phần mềm hệ thống và phần mềm ứng dụng. Nêu vai trò của hệ điều hành.",
         "- Phần mềm hệ thống: quản lí phần cứng máy tính (Windows, macOS) (0,75 điểm)\n- Phần mềm ứng dụng: phục vụ nhu cầu cụ thể của người dùng (Word, Chrome) (0,75 điểm)\n- Hệ điều hành: quản lí tài nguyên và cung cấp giao diện người dùng (0,5 điểm)"),
        ("Câu 3 (2,0 điểm)", "Giải thích đường dẫn: D:\\HocTap\\TinHoc\\Bai1.docx. Nêu 3 phím tắt quản lí tệp.",
         "- Đường dẫn: tệp Bai1.docx trong thư mục TinHoc, con của HocTap, trên ổ D: (1,0 điểm)\n- Ctrl+C (sao chép), Ctrl+V (dán), Ctrl+X (cắt) (1,0 điểm)"),
        ("Câu 4 (2,0 điểm)", "Nêu 3 ưu điểm và 3 nhược điểm của mạng xã hội.",
         "- Ưu: kết nối nhanh, chia sẻ thông tin, học tập trực tuyến (1,0 điểm)\n- Nhược: nghiện, lộ thông tin, bắt nạt trực tuyến (1,0 điểm)"),
        ("Câu 5 (2,0 điểm)", "Tình huống: Bạn Lan nhận được tin nhắn từ người lạ trên mạng xã hội yêu cầu cung cấp địa chỉ nhà và số điện thoại. Em hãy tư vấn cho bạn Lan cách xử lí an toàn.",
         "- Không cung cấp thông tin cá nhân cho người lạ (0,5 điểm)\n- Chặn/Báo cáo tài khoản đó (0,5 điểm)\n- Báo cho phụ huynh/thầy cô (0,5 điểm)\n- Không nhấp vào link lạ (0,5 điểm)"),
    ],
    'kiem_tra_tn': [
        ("Câu 1. Thiết bị nhập (Input) dùng để làm gì?",
         "A. Hiển thị kết quả xử lí ra ngoài", "B. Đưa dữ liệu vào máy tính để xử lí",
         "C. Phát ra âm thanh", "D. In tài liệu ra giấy", "B"),
        ("Câu 2. Bàn phím thuộc loại thiết bị gì?",
         "A. Thiết bị xuất", "B. Thiết bị nhập",
         "C. Thiết bị lưu trữ", "D. Thiết bị xử lí", "B"),
        ("Câu 3. Màn hình máy tính thuộc loại thiết bị gì?",
         "A. Thiết bị nhập", "B. Thiết bị xuất",
         "C. Thiết bị lưu trữ", "D. Thiết bị xử lí", "B"),
        ("Câu 4. Thiết bị nào vừa là thiết bị nhập vừa là thiết bị xuất?",
         "A. Bàn phím", "B. Loa",
         "C. Màn hình cảm ứng", "D. Máy in", "C"),
        ("Câu 5. Webcam thuộc loại thiết bị gì?",
         "A. Thiết bị xuất", "B. Thiết bị nhập",
         "C. Thiết bị lưu trữ", "D. Thiết bị xử lí", "B"),
        ("Câu 6. Phần mềm hệ thống là gì?",
         "A. Là các trò chơi điện tử", "B. Là phần mềm quản lí và điều khiển phần cứng, gồm hệ điều hành và trình điều khiển",
         "C. Là phần mềm Microsoft Word", "D. Là trình duyệt Google Chrome", "B"),
        ("Câu 7. Microsoft Windows thuộc loại phần mềm gì?",
         "A. Phần mềm ứng dụng", "B. Phần mềm hệ thống (hệ điều hành)",
         "C. Trò chơi điện tử", "D. Phần mềm tiện ích", "B"),
        ("Câu 8. Microsoft Word thuộc loại phần mềm gì?",
         "A. Phần mềm hệ thống", "B. Phần mềm ứng dụng",
         "C. Trình điều khiển thiết bị", "D. Chương trình BIOS", "B"),
        ("Câu 9. Hệ điều hành có vai trò gì trong máy tính?",
         "A. Chỉ dùng để soạn thảo văn bản", "B. Quản lí phần cứng, phần mềm và cung cấp giao diện người dùng",
         "C. Chỉ dùng để tạo bảng tính", "D. Chỉ dùng để chơi trò chơi điện tử", "B"),
        ("Câu 10. Trình điều khiển (Driver) dùng để làm gì?",
         "A. Soạn thảo văn bản", "B. Giúp hệ điều hành giao tiếp với các thiết bị phần cứng",
         "C. Duyệt web trên Internet", "D. Chơi trò chơi điện tử", "B"),
        ("Câu 11. Đường dẫn (Path) trong máy tính là gì?",
         "A. Đường đi trên bản đồ", "B. Địa chỉ xác định vị trí của tệp hoặc thư mục trong cây thư mục",
         "C. Tên của một tệp tin", "D. Tên của một phần mềm", "B"),
        ("Câu 12. Tổ hợp phím Ctrl + C dùng để thực hiện thao tác gì?",
         "A. Dán dữ liệu", "B. Sao chép dữ liệu",
         "C. Cắt dữ liệu", "D. Xóa dữ liệu", "B"),
        ("Câu 13. Tổ hợp phím Ctrl + V dùng để thực hiện thao tác gì?",
         "A. Sao chép dữ liệu", "B. Dán dữ liệu",
         "C. Cắt dữ liệu", "D. Xóa dữ liệu", "B"),
        ("Câu 14. Phần mềm File Explorer dùng để làm gì?",
         "A. Duyệt web trên Internet", "B. Quản lí tệp và thư mục trên máy tính",
         "C. Soạn thảo văn bản", "D. Chơi trò chơi điện tử", "B"),
        ("Câu 15. Tệp có đuôi mở rộng .xlsx là loại tệp gì?",
         "A. Tệp văn bản Word", "B. Tệp bảng tính Excel",
         "C. Tệp hình ảnh", "D. Tệp video", "B"),
        ("Câu 16. Kể tên các mạng xã hội phổ biến hiện nay?",
         "A. Microsoft Word, Excel", "B. Facebook, Zalo, TikTok",
         "C. Microsoft Excel, PowerPoint", "D. Paint, Notepad", "B"),
        ("Câu 17. Ưu điểm của mạng xã hội là gì?",
         "A. Gây nghiện cho người sử dụng", "B. Giúp kết nối bạn bè, chia sẻ thông tin và học tập nhanh chóng",
         "C. Làm mất thời gian học tập", "D. Làm lộ thông tin cá nhân", "B"),
        ("Câu 18. Em KHÔNG nên làm gì trên mạng xã hội?",
         "A. Chia sẻ bài học hay cho bạn bè", "B. Đăng ảnh cá nhân, số điện thoại và địa chỉ nhà lên mạng",
         "C. Kết bạn với người quen biết", "D. Hỏi bài thầy cô giáo", "B"),
        ("Câu 19. Khi bị bắt nạt trên mạng, em nên làm gì?",
         "A. Bắt nạt lại người đó", "B. Báo ngay cho thầy cô, phụ huynh và chặn tài khoản người bắt nạt",
         "C. Im lặng chịu đựng", "D. Xóa tài khoản mạng xã hội của mình", "B"),
        ("Câu 20. Để bảo vệ tài khoản mạng xã hội, em nên làm gì?",
         "A. Đặt mật khẩu đơn giản dễ nhớ như 123456", "B. Đặt mật khẩu mạnh và bật xác thực 2 bước",
         "C. Đăng mật khẩu lên Facebook để không bị quên", "D. Dùng chung tài khoản với bạn bè", "B"),
    ],
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành quản lí tệp",
         "Em hãy thực hiện trên máy tính:\na) Mở File Explorer, tạo thư mục 'BAI_KIEM_TRA' trên Desktop (0,5 điểm)\nb) Trong thư mục đó, tạo 2 thư mục con: 'LY_THUYET' và 'THUC_HANH' (0,5 điểm)\nc) Tạo 1 tệp Word tên 'TraLoi.docx' trong thư mục 'LY_THUYET' (0,5 điểm)\nd) Sao chép tệp 'TraLoi.docx' sang thư mục 'THUC_HANH' (0,5 điểm)",
         "a) Mở File Explorer, tạo thư mục đúng (0,5 điểm)\nb) Tạo 2 thư mục con (0,5 điểm)\nc) Tạo tệp Word đúng tên (0,5 điểm)\nd) Sao chép thành công (Ctrl+C, Ctrl+V) (0,5 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành nhận diện phần mềm",
         "Em hãy thực hiện các yêu cầu sau:\na) Liệt kê ba phần mềm hệ thống và ba phần mềm ứng dụng đang có trên máy tính (1,0 điểm)\nb) Mở phần mềm Microsoft Word, soạn một đoạn văn ngắn giới thiệu về trường UNIGO (1,0 điểm)\nc) Lưu tệp với tên 'GioiThieu_UNIGO.docx' vào thư mục 'BAI_KIEM_TRA' (1,0 điểm)",
         "a) Phần mềm hệ thống: Windows, Driver, Antivirus; Phần mềm ứng dụng: Word, Chrome, Paint (1,0 điểm)\nb) Mở Microsoft Word, gõ đoạn văn đúng yêu cầu (1,0 điểm)\nc) Lưu đúng tên tệp và vị trí yêu cầu (1,0 điểm)"),
    ],
}

# ──────────────── LỚP 8 ────────────────
GRADE_8 = {
    'lop': '8',
    'ten_lop': 'Lớp 8',
    'noi_dung_on': [
        "Bài 1: Lược sử công cụ tính toán – Các thế hệ máy tính (1-5), linh kiện đặc trưng, ENIAC.",
        "Bài 2: Thông tin trong môi trường số – Đặc điểm thông tin số, tác động tích cực/tiêu cực, đánh giá độ tin cậy.",
        "Bài 3: Thực hành khai thác thông tin số – Kỹ năng tìm kiếm, trích dẫn, bản quyền.",
        "Bài 4: Đạo đức và văn hóa trong sử dụng công nghệ số – Quy tắc ứng xử, bản quyền, an toàn.",
        "Bài 5: Sử dụng bảng tính giải quyết bài toán thực tế (Tiết 1) – Giới thiệu Excel, ô, hàng, cột, nhập dữ liệu.",
    ],
    'matrix_headers': ["Nội dung", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng"],
    'matrix_on_tap': [
        ["Bài 1. Lược sử công cụ tính toán", "4 TN", "4 TN", "2 TN + 1 TL", "-", "10 TN + 1 TL"],
        ["Bài 2. Thông tin trong MT số", "5 TN", "3 TN", "1 TN + 1 TL", "1 TN", "10 TN + 1 TL"],
        ["Bài 3. Thực hành khai thác TTS", "3 TN", "3 TN", "2 TN + 1 TL", "1 TN + 1 TL", "9 TN + 2 TL"],
        ["Bài 4-5. Đạo đức và bảng tính", "5 TN", "4 TN", "1 TN + 1 TL", "1 TN", "11 TN + 1 TL"],
        ["TỔNG", "17 câu", "14 câu", "6 TN + 4 TL", "3 TN + 1 TL", "40 TN + 5 TL"],
    ],
    'matrix_kiem_tra': [
        ["Bài 1. Lược sử công cụ", "3 TN (0.75 điểm)", "2 TN (0.5 điểm)", "-", "-", "1.25đ"],
        ["Bài 2. Thông tin MT số", "2 TN (0.5 điểm)", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "-", "1.25đ"],
        ["Bài 3. Khai thác TTS", "2 TN (0.5 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1.25đ"],
        ["Bài 4-5. Đạo đức và bản quyền", "3 TN (0.75 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1 TN (0.25 điểm)", "1.5đ"],
        ["Phần thực hành", "-", "-", "1 câu (2.0 điểm)", "1 câu (2.75 điểm)", "4.75đ"],
        ["TỔNG", "10 câu (2.5 điểm)", "6 câu (1.5 điểm)", "3 TN + 1 TH (2.5 điểm)", "2 TN + 1 TH (3.25 điểm)", "10.0đ"],
    ],
    'on_tap_tn': [
        # Bài 1 (10)
        ("Câu 1. Máy tính thế hệ 1 sử dụng linh kiện gì?", "A. Đèn chân không", "B. Transistor", "C. IC", "D. VLSI", "A"),
        ("Câu 2. Máy tính điện tử đầu tiên tên gì?", "A. Apple I", "B. ENIAC", "C. IBM PC", "D. Pascaline", "B"),
        ("Câu 3. Máy tính thế hệ 2 dùng linh kiện gì?", "A. Đèn chân không", "B. Transistor (bóng bán dẫn)", "C. IC", "D. VLSI", "B"),
        ("Câu 4. Máy tính thế hệ 3 dùng linh kiện gì?", "A. Đèn chân không", "B. Transistor", "C. Mạch tích hợp IC", "D. VLSI", "C"),
        ("Câu 5. Máy tính thế hệ 4 gắn liền với gì?", "A. Đèn chân không", "B. Vi xử lý (Microprocessor)", "C. Rơ le", "D. Ống tia âm cực", "B"),
        ("Câu 6. Máy tính thế hệ 5 hướng tới gì?", "A. Tốc độ chậm hơn", "B. Trí tuệ nhân tạo (AI)", "C. Dùng đèn chân không", "D. Không có mạng", "B"),
        ("Câu 7. Bàn tính (Abacus) là gì?", "A. Máy tính điện tử", "B. Công cụ tính toán thủ công cổ xưa nhất", "C. Phần mềm", "D. Robot", "B"),
        ("Câu 8. Máy Pascaline do ai phát minh?", "A. Alan Turing", "B. Blaise Pascal", "C. Steve Jobs", "D. Bill Gates", "B"),
        ("Câu 9. ENIAC ra đời năm nào?", "A. 1920", "B. 1946", "C. 1970", "D. 2000", "B"),
        ("Câu 10. Xu hướng phát triển máy tính là gì?", "A. Ngày càng to hơn", "B. Nhỏ gọn, nhanh hơn, thông minh hơn", "C. Chậm hơn", "D. Đắt hơn và khó dùng", "B"),
        # Bài 2 (10)
        ("Câu 11. Đặc điểm của thông tin trong môi trường số là gì?", "A. Ít và chậm", "B. Khối lượng lớn, đa dạng, lan truyền nhanh", "C. Luôn chính xác 100%", "D. Không thể lưu trữ", "B"),
        ("Câu 12. Đâu KHÔNG phải đặc điểm TT số?", "A. Khối lượng khổng lồ", "B. Đa dạng dạng TT", "C. Hoàn toàn chính xác không cần kiểm chứng", "D. Lan truyền nhanh", "C"),
        ("Câu 13. Tác động tích cực của TT số?", "A. Nghiện mạng", "B. Truy cập tri thức nhanh, học tập hiệu quả", "C. Lộ thông tin cá nhân", "D. Mất thời gian", "B"),
        ("Câu 14. Tác động tiêu cực của TT số?", "A. Học tập hiệu quả", "B. Nghiện mạng, tiếp xúc thông tin xấu", "C. Kết nối bạn bè", "D. Truy cập tri thức", "B"),
        ("Câu 15. Bước đầu tiên đánh giá độ tin cậy TT?", "A. Chia sẻ ngay", "B. Kiểm tra tác giả và nguồn gốc", "C. Like", "D. Tin tuyệt đối", "B"),
        ("Câu 16. Tên miền .gov thuộc loại tổ chức nào?", "A. Thương mại", "B. Chính phủ", "C. Giải trí", "D. Cá nhân", "B"),
        ("Câu 17. Tin giả (Fake News) có đặc điểm gì?", "A. Có tên tác giả rõ ràng", "B. Giật gân, không rõ nguồn, kêu gọi chia sẻ gấp", "C. Từ báo chính thống", "D. Từ Bộ GD&ĐT", "B"),
        ("Câu 18. Kiểm chứng chéo nghĩa là gì?", "A. Tin 1 nguồn duy nhất", "B. So sánh thông tin từ nhiều nguồn khác nhau", "C. Xóa thông tin", "D. Chia sẻ ngay", "B"),
        ("Câu 19. Trang .edu thuộc loại tổ chức nào?", "A. Thương mại", "B. Giáo dục", "C. Giải trí", "D. Quân sự", "B"),
        ("Câu 20. Khi thấy thông tin giật gân trên mạng xã hội, em nên làm gì?", "A. Chia sẻ ngay", "B. Kiểm tra nguồn gốc trước khi tin", "C. Like", "D. Comment khen", "B"),
        # Bài 3 (9)
        ("Câu 21. Khi tìm kiếm trên Google, dùng dấu ngoặc kép để làm gì?", "A. Xóa kết quả", "B. Tìm chính xác cụm từ", "C. Mở tab mới", "D. Tắt Google", "B"),
        ("Câu 22. Trích dẫn nguồn tham khảo nghĩa là?", "A. Xóa tên tác giả", "B. Ghi rõ tên tác giả, nguồn gốc khi sử dụng", "C. Tự nhận là của mình", "D. Không cần ghi", "B"),
        ("Câu 23. Vi phạm bản quyền là hành vi nào?", "A. Ghi rõ nguồn", "B. Sao chép tác phẩm người khác rồi tự nhận", "C. Dùng Creative Commons", "D. Chia sẻ link gốc", "B"),
        ("Câu 24. Creative Commons là gì?", "A. Virus", "B. Giấy phép cho phép sử dụng tác phẩm với điều kiện nhất định", "C. Phần mềm", "D. Mạng xã hội", "B"),
        ("Câu 25. Toán tử site: trên Google dùng để?", "A. Tắt Google", "B. Tìm kết quả chỉ trong 1 trang web cụ thể", "C. Xóa kết quả", "D. Mở email", "B"),
        ("Câu 26. Toán tử - (trừ) trên Google dùng để?", "A. Cộng kết quả", "B. Loại bỏ từ khóa không mong muốn", "C. Nhân kết quả", "D. Chia kết quả", "B"),
        ("Câu 27. Google Scholar dùng để tìm gì?", "A. Game", "B. Tài liệu học thuật, nghiên cứu khoa học", "C. Video giải trí", "D. Nhạc", "B"),
        ("Câu 28. Khi sử dụng ảnh từ Internet, em cần?", "A. Thoải mái không cần xin phép", "B. Kiểm tra bản quyền và ghi rõ nguồn", "C. Xóa tên tác giả", "D. Đổi tên thành của mình", "B"),
        ("Câu 29. Filetype:pdf trên Google dùng để?", "A. Xóa PDF", "B. Tìm chỉ tệp PDF", "C. Mở PDF", "D. In PDF", "B"),
        # Bài 4-5 (11)
        ("Câu 30. Đạo đức số là gì?", "A. Chơi game", "B. Chuẩn mực hành vi khi sử dụng công nghệ số", "C. Phần mềm", "D. Phần cứng", "B"),
        ("Câu 31. Hành vi nào vi phạm đạo đức số?", "A. Ghi nguồn trích dẫn", "B. Hack tài khoản người khác", "C. Dùng phần mềm có bản quyền", "D. Bảo vệ thông tin cá nhân", "B"),
        ("Câu 32. Bản quyền phần mềm nghĩa là gì?", "A. Miễn phí hoàn toàn", "B. Quyền sở hữu trí tuệ của tác giả đối với phần mềm", "C. Không ai sở hữu", "D. Chỉ dùng 1 ngày", "B"),
        ("Câu 33. Phần mềm mã nguồn mở là gì?", "A. Phần mềm không dùng được", "B. Phần mềm cho phép xem, sửa đổi và chia sẻ mã nguồn miễn phí", "C. Phần mềm rất đắt", "D. Phần mềm chỉ chạy trên hệ điều hành Mac", "B"),
        ("Câu 34. Ô trong bảng tính Excel được xác định bởi gì?", "A. Chỉ hàng", "B. Giao điểm của cột và hàng (ví dụ: A1, B2)", "C. Chỉ cột", "D. Tên tệp", "B"),
        ("Câu 35. Cột trong Excel được đánh tên bằng gì?", "A. Số (1, 2, 3)", "B. Chữ cái (A, B, C)", "C. Kí hiệu đặc biệt", "D. Ngày tháng", "B"),
        ("Câu 36. Hàng trong Excel được đánh số bằng gì?", "A. Chữ cái", "B. Số (1, 2, 3...)", "C. Tên tháng", "D. Ký hiệu", "B"),
        ("Câu 37. Thanh công thức trong Excel dùng để?", "A. Vẽ tranh", "B. Nhập và hiển thị nội dung/công thức của ô đang chọn", "C. In tài liệu", "D. Tắt Excel", "B"),
        ("Câu 38. Để nhập dữ liệu vào ô Excel, em cần?", "A. Click vào ô → gõ nội dung → Enter", "B. Rút USB", "C. Tắt máy", "D. Xóa Excel", "A"),
        ("Câu 39. Sheet (trang tính) trong Excel là gì?", "A. Trang web", "B. Một bảng tính chứa các ô dữ liệu", "C. Phần mềm Word", "D. File hình ảnh", "B"),
        ("Câu 40. Để mở Excel, em click vào biểu tượng nào?", "A. Biểu tượng W (Word)", "B. Biểu tượng X (Excel) màu xanh lá", "C. Biểu tượng P (PowerPoint)", "D. Biểu tượng Chrome", "B"),
    ],
    'on_tap_tl': [
        ("Câu 1 (2,0 điểm)", "Trình bày đặc điểm của 5 thế hệ máy tính điện tử (linh kiện đặc trưng, ví dụ đại diện).",
         "- TH1 (1945-1955): đèn chân không, ENIAC (0,4 điểm)\n- TH2 (1955-1965): Transistor (0,4 điểm)\n- TH3 (1965-1971): IC (0,4 điểm)\n- TH4 (1971-nay): Vi xử lý, PC (0,4 điểm)\n- TH5 (tương lai): AI, lượng tử (0,4 điểm)"),
        ("Câu 2 (2,0 điểm)", "Nêu 4 tiêu chí đánh giá độ tin cậy thông tin trên Internet. Cho ví dụ 1 tiêu chí.",
         "- Nguồn gốc/tác giả (0,4 điểm)\n- Tính cập nhật (0,4 điểm)\n- Mục đích bài viết (0,4 điểm)\n- Kiểm chứng chéo (0,4 điểm)\n- Ví dụ phù hợp (0,4 điểm)"),
        ("Câu 3 (2,0 điểm)", "Tình huống: Nam tìm bài viết hay trên mạng nhưng không rõ tác giả. Nam có nên sao chép nguyên văn không? Vì sao? Tư vấn cho Nam cách tìm nguồn tin cậy.",
         "- Không nên sao chép: vi phạm bản quyền, rủi ro tin sai (1,0 điểm)\n- Tư vấn: tìm nguồn .gov/.edu, kiểm chứng, trích dẫn hợp lệ (1,0 điểm)"),
        ("Câu 4 (2,0 điểm)", "Giải thích khái niệm bản quyền phần mềm. Phân biệt phần mềm thương mại, phần mềm miễn phí và phần mềm mã nguồn mở.",
         "- Bản quyền: quyền sở hữu trí tuệ (0,5 điểm)\n- Thương mại: phải mua (Microsoft Office) (0,5 điểm)\n- Miễn phí: dùng free (Chrome) (0,5 điểm)\n- Mã nguồn mở: xem/sửa mã (LibreOffice) (0,5 điểm)"),
        ("Câu 5 (2,0 điểm)", "Mô tả giao diện phần mềm bảng tính Excel: thanh Ribbon, thanh công thức, vùng bảng tính. Giải thích cách xác định địa chỉ ô.",
         "- Thanh Ribbon: chứa các tab lệnh (Home, Insert...) (0,5 điểm)\n- Thanh công thức: hiển thị nội dung ô (0,5 điểm)\n- Vùng bảng tính: gồm ô, hàng, cột (0,5 điểm)\n- Địa chỉ ô: tên cột + số hàng (ví dụ: B3) (0,5 điểm)"),
    ],
    'kiem_tra_tn': [
        ("Câu 1. Máy tính điện tử thế hệ thứ nhất (1945 - 1955) sử dụng linh kiện điện tử nào?",
         "A. Đèn điện tử chân không", "B. Bóng bán dẫn (Transistor)",
         "C. Mạch tích hợp (IC)", "D. Mạch tích hợp cỡ lớn (VLSI)", "A"),
        ("Câu 2. Chiếc máy tính điện tử đầu tiên trên thế giới có tên gọi là gì?",
         "A. Apple I", "B. ENIAC",
         "C. IBM PC", "D. Pascaline", "B"),
        ("Câu 3. Máy tính thế hệ thứ tư (từ năm 1971 đến nay) gắn liền với sự phát triển của linh kiện nào?",
         "A. Đèn điện tử chân không", "B. Vi xử lý (Microprocessor) tích hợp mật độ cao",
         "C. Rơ le điện từ", "D. Ống tia âm cực", "B"),
        ("Câu 4. Đặc điểm của thông tin trong môi trường số là gì?",
         "A. Có khối lượng ít và lan truyền chậm", "B. Có khối lượng lớn, đa dạng và lan truyền nhanh chóng",
         "C. Luôn chính xác 100% không cần kiểm chứng", "D. Không thể lưu trữ được", "B"),
        ("Câu 5. Phát biểu nào sau đây KHÔNG phải là đặc điểm của thông tin số?",
         "A. Có khối lượng khổng lồ và tăng nhanh", "B. Tồn tại ở nhiều dạng khác nhau",
         "C. Hoàn toàn chính xác 100% không cần kiểm chứng", "D. Lan truyền nhanh chóng trên toàn cầu", "C"),
        ("Câu 6. Khi tiếp nhận thông tin trên mạng, bước đầu tiên em cần làm để đánh giá độ tin cậy là gì?",
         "A. Nhấn nút chia sẻ ngay cho bạn bè", "B. Kiểm tra tác giả và nguồn gốc xuất bản của thông tin",
         "C. Bấm nút thích (Like)", "D. Tin tưởng tuyệt đối vì có nhiều lượt xem", "B"),
        ("Câu 7. Tin giả (Fake News) thường có đặc điểm gì?",
         "A. Có tên tác giả rõ ràng và nguồn uy tín", "B. Tiêu đề giật gân, không rõ nguồn gốc, kêu gọi chia sẻ gấp",
         "C. Được đăng trên báo chính thống", "D. Được phát hành từ Bộ Giáo dục và Đào tạo", "B"),
        ("Câu 8. Khi tìm kiếm trên Google, đặt cụm từ trong dấu ngoặc kép có tác dụng gì?",
         "A. Xóa tất cả kết quả tìm kiếm", "B. Tìm kiếm chính xác cụm từ đó",
         "C. Mở thêm một tab trình duyệt mới", "D. Tắt công cụ tìm kiếm Google", "B"),
        ("Câu 9. Hành vi nào sau đây là vi phạm bản quyền?",
         "A. Ghi rõ nguồn tham khảo khi trích dẫn", "B. Sao chép tác phẩm của người khác rồi tự nhận là của mình",
         "C. Sử dụng tài liệu có giấy phép Creative Commons", "D. Chia sẻ đường liên kết bài viết gốc cho người khác", "B"),
        ("Câu 10. Giấy phép Creative Commons là gì?",
         "A. Là một loại virus máy tính", "B. Là giấy phép cho phép sử dụng tác phẩm với điều kiện nhất định",
         "C. Là một phần mềm máy tính", "D. Là một mạng xã hội", "B"),
        ("Câu 11. Hành vi nào sau đây vi phạm đạo đức khi sử dụng công nghệ số?",
         "A. Ghi rõ nguồn khi trích dẫn tài liệu", "B. Xâm nhập trái phép vào tài khoản của người khác",
         "C. Sử dụng phần mềm có bản quyền hợp pháp", "D. Bảo vệ thông tin cá nhân của mình", "B"),
        ("Câu 12. Phần mềm mã nguồn mở là gì?",
         "A. Là phần mềm không thể sử dụng được", "B. Là phần mềm cho phép người dùng xem, sửa đổi và chia sẻ mã nguồn miễn phí",
         "C. Là phần mềm có giá rất đắt", "D. Là phần mềm chỉ chạy được trên hệ điều hành Mac", "B"),
        ("Câu 13. Ô trong phần mềm bảng tính Excel được xác định bởi yếu tố nào?",
         "A. Chỉ bởi số hàng", "B. Bởi giao điểm của tên cột và số hàng (ví dụ: A1, B2)",
         "C. Chỉ bởi tên cột", "D. Bởi tên tệp", "B"),
        ("Câu 14. Cột trong phần mềm bảng tính Excel được đánh tên bằng gì?",
         "A. Bằng số (1, 2, 3)", "B. Bằng chữ cái (A, B, C)",
         "C. Bằng ký hiệu đặc biệt", "D. Bằng ngày tháng", "B"),
        ("Câu 15. Hàng trong phần mềm bảng tính Excel được đánh số bằng gì?",
         "A. Bằng chữ cái", "B. Bằng số (1, 2, 3...)",
         "C. Bằng tên tháng trong năm", "D. Bằng ký hiệu đặc biệt", "B"),
        ("Câu 16. Thanh công thức trong phần mềm bảng tính Excel dùng để làm gì?",
         "A. Dùng để vẽ tranh", "B. Dùng để nhập và hiển thị nội dung hoặc công thức của ô đang chọn",
         "C. Dùng để in tài liệu", "D. Dùng để tắt phần mềm Excel", "B"),
        ("Câu 17. Để nhập dữ liệu vào một ô trong Excel, em cần thực hiện thao tác gì?",
         "A. Nhấp chuột vào ô cần nhập → gõ nội dung → nhấn phím Enter", "B. Rút thiết bị USB ra khỏi máy tính",
         "C. Tắt máy tính", "D. Gỡ bỏ phần mềm Excel khỏi máy", "A"),
        ("Câu 18. Tên miền có đuôi .edu thuộc loại tổ chức nào?",
         "A. Tổ chức thương mại", "B. Tổ chức giáo dục",
         "C. Tổ chức giải trí", "D. Tổ chức quân sự", "B"),
        ("Câu 19. Toán tử site: khi tìm kiếm trên Google dùng để làm gì?",
         "A. Dùng để tắt Google", "B. Dùng để tìm kết quả chỉ trong một trang web cụ thể",
         "C. Dùng để xóa tất cả kết quả tìm kiếm", "D. Dùng để mở hộp thư email", "B"),
        ("Câu 20. Xu hướng phát triển của máy tính hiện đại là gì?",
         "A. Ngày càng to lớn và cồng kềnh hơn", "B. Ngày càng nhỏ gọn, nhanh hơn và thông minh hơn nhờ trí tuệ nhân tạo",
         "C. Ngày càng chạy chậm hơn", "D. Ngày càng đắt đỏ và khó sử dụng", "B"),
    ],
    'kiem_tra_th': [
        ("Câu 1 (2,0 điểm) – Thực hành bảng tính",
         "Em hãy thực hiện trên máy tính:\na) Mở phần mềm Microsoft Excel (0,5 điểm)\nb) Nhập bảng điểm lớp 8A1 gồm các cột: Số thứ tự, Họ tên, Toán, Ngữ văn, Tiếng Anh, Tin học. Nhập dữ liệu cho 5 bạn học sinh (1,0 điểm)\nc) Lưu tệp với tên 'BangDiem_8A1.xlsx' (0,5 điểm)",
         "a) Mở phần mềm Microsoft Excel thành công (0,5 điểm)\nb) Nhập đúng bảng sáu cột và năm hàng dữ liệu (1,0 điểm)\nc) Lưu đúng tên tệp vào máy tính (0,5 điểm)"),
        ("Câu 2 (3,0 điểm) – Thực hành đánh giá thông tin",
         "Em hãy thực hiện các yêu cầu sau:\na) Mở trình duyệt Google Chrome, tìm kiếm với từ khóa: \"Trí tuệ nhân tạo AI trong giáo dục\" (0,5 điểm)\nb) Chọn hai kết quả từ hai nguồn khác nhau, ghi lại tiêu đề và địa chỉ nguồn (1,0 điểm)\nc) So sánh hai nguồn: nội dung giống và khác nhau, đánh giá nguồn nào đáng tin cậy hơn và giải thích lí do (1,5 điểm)",
         "a) Tìm kiếm đúng từ khóa trên trình duyệt (0,5 điểm)\nb) Chọn hai nguồn, ghi đầy đủ tiêu đề và địa chỉ trang web (1,0 điểm)\nc) So sánh hợp lí, đánh giá chính xác dựa trên tiêu chí tin cậy (1,5 điểm)"),
    ],
}

ALL_GRADES = [GRADE_3, GRADE_4, GRADE_5, GRADE_6, GRADE_7, GRADE_8]


# ═══════════════════════════════════════════════════════════════════════
# MAIN GENERATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════

def generate_on_tap(grade_data):
    """Tạo file Đề ôn tập (Tuần 9) cho 1 lớp."""
    lop = grade_data['lop']
    is_thcs = int(lop) >= 6
    template = TEMPLATE_THCS if is_thcs else TEMPLATE_TH
    
    # Tạo thư mục đầu ra
    out_dir = os.path.join(OUT_DIR, f'Lớp_{lop}', 'Tuần_09')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'De_on_tap_DGDK1_Tin_hoc_Lop_{lop}.docx')

    try:
        doc = Document(template)
        clean_body(doc)
    except:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.7)
            sec.bottom_margin = Inches(0.7)
            sec.left_margin = Inches(0.75)
            sec.right_margin = Inches(0.75)

    # 1. Header
    build_header_table(doc, lop, de_type="ÔN TẬP")
    
    # 2. Student info
    build_student_info(doc, lop)
    add_para(doc, "", space_after=4)

    # 3. Ma trận phạm vi kiến thức
    add_para(doc, "A. MA TRẬN PHẠM VI KIẾN THỨC ÔN TẬP", bold=True, size_pt=13, color_rgb=(27, 79, 155), space_before=6)
    for item in grade_data['noi_dung_on']:
        add_para(doc, f"  • {item}", size_pt=12, space_after=2)
    add_para(doc, "", space_after=4)

    # 4. Bảng ma trận ôn tập
    add_para(doc, "BẢNG MA TRẬN ĐỀ ÔN TẬP", bold=True, size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4)
    build_matrix_table(doc, grade_data['matrix_on_tap'], grade_data['matrix_headers'])
    add_para(doc, "", space_after=6)

    # 5. Phần I: Trắc nghiệm
    add_para(doc, "B. NỘI DUNG ĐỀ ÔN TẬP", bold=True, size_pt=14, color_rgb=(27, 79, 155), space_before=8)
    add_para(doc, f"PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (40 câu)", bold=True, size_pt=13, space_before=4)
    add_para(doc, "Khoanh tròn vào chữ cái (A, B, C hoặc D) đứng trước câu trả lời đúng nhất:", italic=True, size_pt=12, space_after=4)

    for q_text, opt_a, opt_b, opt_c, opt_d, _ in grade_data['on_tap_tn']:
        add_para(doc, q_text, bold=True, size_pt=12, space_before=4, space_after=2, keep_with_next=True)
        add_question_options(doc, opt_a, opt_b, opt_c, opt_d, size_pt=12)

    # 6. Phần II: Tự luận
    add_para(doc, f"PHẦN II. TỰ LUẬN (5 câu)", bold=True, size_pt=13, space_before=8)
    for title, content, _ in grade_data['on_tap_tl']:
        add_para(doc, title, bold=True, size_pt=12, space_before=6)
        add_para(doc, f"  {content}", size_pt=12, space_after=4)

    # 7. Đáp án trắc nghiệm
    add_page_break(doc)
    add_para(doc, "C. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM", bold=True, size_pt=14, color_rgb=(27, 79, 155), space_before=6)
    add_para(doc, "I. ĐÁP ÁN TRẮC NGHIỆM", bold=True, size_pt=13, space_before=4)
    build_answer_key_table(doc, grade_data['on_tap_tn'])
    add_para(doc, "*(Mỗi câu trắc nghiệm: 0,25 điểm)*", italic=True, size_pt=11, space_before=4)

    # 8. Hướng dẫn chấm tự luận
    add_para(doc, "II. HƯỚNG DẪN CHẤM TỰ LUẬN", bold=True, size_pt=13, space_before=8)
    for title, _, guide in grade_data['on_tap_tl']:
        add_para(doc, f"{title}:", bold=True, size_pt=12, space_before=4)
        add_para(doc, f"  {guide}", size_pt=11, space_after=4)

    # 9. Ký duyệt
    build_signature_table(doc)

    # Save
    try:
        doc.save(out_path)
        print(f"  ✅ Đã tạo: {out_path}")
    except PermissionError:
        alt_path = out_path.replace('.docx', '_new.docx')
        doc.save(alt_path)
        print(f"  ⚠️ File đang mở, lưu tại: {alt_path}")
    return out_path


def generate_kiem_tra(grade_data):
    """Tạo file Đề kiểm tra (Tuần 10) cho 1 lớp."""
    lop = grade_data['lop']
    is_thcs = int(lop) >= 6
    template = TEMPLATE_THCS if is_thcs else TEMPLATE_TH
    
    out_dir = os.path.join(OUT_DIR, f'Lớp_{lop}', 'Tuần_10')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'De_kiem_tra_DGDK1_Tin_hoc_Lop_{lop}.docx')

    try:
        doc = Document(template)
        clean_body(doc)
    except:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.7)
            sec.bottom_margin = Inches(0.7)
            sec.left_margin = Inches(0.75)
            sec.right_margin = Inches(0.75)

    # 1. Header
    build_header_table(doc, lop, de_type="KIỂM TRA")

    # 2. Student info
    build_student_info(doc, lop)
    add_para(doc, "", space_after=4)

    # 3. Ma trận đề kiểm tra - KHÔNG ghi vào đề kiểm tra (chỉ ghi trong đề ôn tập)

    # 4. Phần I: Trắc nghiệm (20 câu - 5.25đ)
    add_para(doc, "PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (20 câu – 5,0 điểm)", bold=True, size_pt=13, space_before=6)
    add_para(doc, "Khoanh tròn vào chữ cái (A, B, C hoặc D) đứng trước câu trả lời đúng nhất:", italic=True, size_pt=12, space_after=4)

    for q_text, opt_a, opt_b, opt_c, opt_d, _ in grade_data['kiem_tra_tn']:
        add_para(doc, q_text, bold=True, size_pt=12, space_before=4, space_after=2, keep_with_next=True)
        add_question_options(doc, opt_a, opt_b, opt_c, opt_d, size_pt=12)

    # 5. Phần II: Thực hành (2 câu - 4.75đ)
    add_para(doc, "PHẦN II. THỰC HÀNH (2 câu – 5,0 điểm)", bold=True, size_pt=13, space_before=8)
    for title, content, _ in grade_data['kiem_tra_th']:
        add_para(doc, title, bold=True, size_pt=12, space_before=6)
        add_para(doc, f"  {content}", size_pt=12, space_after=4)

    # Kết thúc đề
    add_para(doc, "─── Hết ───", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13, space_before=8)
    add_para(doc, "Giám thị không giải thích gì thêm.", italic=True, size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 6. Đáp án (trang mới)
    add_page_break(doc)
    add_para(doc, "ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM", bold=True, size_pt=14, 
             align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=(27, 79, 155), space_before=6)
    
    add_para(doc, "I. TRẮC NGHIỆM (20 câu – 5,0 điểm. Mỗi câu trả lời đúng được 0,25 điểm)", 
             bold=True, size_pt=13, space_before=4)
    build_answer_key_table(doc, grade_data['kiem_tra_tn'])

    add_para(doc, "II. THỰC HÀNH (2 câu – 5,0 điểm)", bold=True, size_pt=13, space_before=8)
    for title, _, guide in grade_data['kiem_tra_th']:
        add_para(doc, f"{title}:", bold=True, size_pt=12, space_before=4)
        add_para(doc, f"  {guide}", size_pt=11, space_after=4)

    # Ký duyệt
    build_signature_table(doc)

    try:
        doc.save(out_path)
        print(f"  ✅ Đã tạo: {out_path}")
    except PermissionError:
        alt_path = out_path.replace('.docx', '_new.docx')
        doc.save(alt_path)
        print(f"  ⚠️ File đang mở, lưu tại: {alt_path}")
    return out_path


def main():
    print("=" * 70)
    print("  TẠO MA TRẬN ĐỀ & ĐỀ KIỂM TRA ĐÁNH GIÁ ĐỊNH KỲ 1")
    print("  MÔN TIN HỌC – LỚP 3 ĐẾN LỚP 8")
    print("  NĂM HỌC 2026 – 2027")
    print("=" * 70)
    
    all_files = []
    for grade in ALL_GRADES:
        lop = grade['lop']
        print(f"\n{'─'*50}")
        print(f"  📝 LỚP {lop}")
        print(f"{'─'*50}")
        
        # Tuần 9: Đề ôn tập
        print(f"  [Tuần 9] Tạo đề ôn tập...")
        f1 = generate_on_tap(grade)
        all_files.append(f1)
        
        # Tuần 10: Đề kiểm tra
        print(f"  [Tuần 10] Tạo đề kiểm tra...")
        f2 = generate_kiem_tra(grade)
        all_files.append(f2)
    
    print(f"\n{'='*70}")
    print(f"  ✅ HOÀN THÀNH! Đã tạo {len(all_files)} file.")
    print(f"{'='*70}")
    
    for f in all_files:
        print(f"  📄 {f}")

if __name__ == '__main__':
    main()
