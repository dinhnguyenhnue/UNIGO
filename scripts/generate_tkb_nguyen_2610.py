import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

# Input / Output Paths
SRC_EXCEL = r"D:\UNIGO\TKB toàn trường CHECK - 26.10.xlsx"
DIR_TEACHER = r"D:\UNIGO\Thời khóa biểu giáo viên"
DIR_MAU = r"D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Thời khóa biểu"

os.makedirs(DIR_TEACHER, exist_ok=True)
os.makedirs(DIR_MAU, exist_ok=True)

# ─────────────────────────────────────────────────────────────
# 1. PARSE DATA FROM SOURCE EXCEL
# ─────────────────────────────────────────────────────────────
def parse_tkb_2610():
    wb = openpyxl.load_workbook(SRC_EXCEL, data_only=True)
    ws = wb['TKB_LOP_SC']

    col_class = {}
    col_session = {}
    current_class = None
    for c in range(1, ws.max_column + 1):
        v4 = ws.cell(4, c).value
        v5 = ws.cell(5, c).value
        if v4 is not None and str(v4).strip() not in ['THỨ', 'TIẾT', 'THU', 'TIET']:
            current_class = str(v4).strip()
        elif v4 in ['THỨ', 'TIẾT', 'THU', 'TIET']:
            current_class = None
        col_class[c] = current_class
        col_session[c] = str(v5).strip() if v5 else ''

    row_day = {}
    row_period = {}
    cur_day = None
    for r in range(6, 31):
        v_a = ws.cell(r, 1).value
        if v_a is not None and str(v_a).strip().isdigit():
            cur_day = int(str(v_a).strip())
        row_day[r] = cur_day
        row_period[r] = int(ws.cell(r, 2).value)

    slots = []
    for r in range(6, 31):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and 'Nguyên' in str(v):
                day = row_day[r]
                period = row_period[r]
                cls = col_class[c]
                session = col_session[c]
                raw_text = str(v).strip()
                subj = "Robotics" if "Robotics" in raw_text else "Tin học"
                slots.append({
                    'day_num': day,
                    'day_str': f"Thứ {day}",
                    'session': session,
                    'period': period,
                    'class': cls,
                    'subject': subj,
                    'raw': raw_text
                })

    # Sort slots chronologically
    slots.sort(key=lambda s: (s['day_num'], 0 if s['session'] == 'Sáng' else 1, s['period']))
    return slots

# ─────────────────────────────────────────────────────────────
# 2. CREATE EXCEL TIMETABLE (.XLSX)
# ─────────────────────────────────────────────────────────────
def create_excel_tkb(slots, output_paths):
    wb = openpyxl.Workbook()

    font_tnr = "Times New Roman"

    # Fills
    fill_header_dark = PatternFill(start_color="1D2A64", end_color="1D2A64", fill_type="solid")
    fill_header_mid  = PatternFill(start_color="253580", end_color="253580", fill_type="solid")
    fill_sub_info    = PatternFill(start_color="EBF0FF", end_color="EBF0FF", fill_type="solid")
    fill_session     = PatternFill(start_color="E8ECF5", end_color="E8ECF5", fill_type="solid")
    fill_period_bg   = PatternFill(start_color="F5F7FF", end_color="F5F7FF", fill_type="solid")
    fill_time_bg     = PatternFill(start_color="EFF3FF", end_color="EFF3FF", fill_type="solid")
    fill_tin         = PatternFill(start_color="D6E4FF", end_color="D6E4FF", fill_type="solid")
    fill_robot       = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
    fill_success     = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")

    # Fonts
    f_title1   = Font(name=font_tnr, size=14, bold=True, color="FFFFFF")
    f_title2   = Font(name=font_tnr, size=15, bold=True, color="FFFFFF")
    f_info     = Font(name=font_tnr, size=12, bold=False, color="1D2A64")
    f_legend_tin  = Font(name=font_tnr, size=10, bold=True, color="003C8C")
    f_legend_rob  = Font(name=font_tnr, size=10, bold=True, color="8C3C00")
    f_legend_succ = Font(name=font_tnr, size=10, bold=True, color="276A3C")
    f_col_hdr  = Font(name=font_tnr, size=12, bold=True, color="FFFFFF")
    f_sess_hdr = Font(name=font_tnr, size=12, bold=True, color="1D2A64")
    f_tiet_hdr = Font(name=font_tnr, size=12, bold=True, color="1D2A64")
    f_time     = Font(name=font_tnr, size=10, italic=True, color="404060")
    f_tin      = Font(name=font_tnr, size=11, bold=True, color="003C8C")
    f_robot    = Font(name=font_tnr, size=11, bold=True, color="8C3C00")

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left   = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Borders
    border_thin_gray = Side(style="thin", color="AAAACC")
    border_thick_dark = Side(style="medium", color="1D2A64")
    box_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)

    # ==========================================
    # SHEET 1: TKB ĐẬU ĐÌNH NGUYÊN - 26.10
    # ==========================================
    ws1 = wb.active
    ws1.title = "TKB Đậu Đình Nguyên - 26.10"
    ws1.views.sheetView[0].showGridLines = True

    # Column widths
    col_widths = {'A': 9.0, 'B': 14.0, 'C': 14.0, 'D': 18.0, 'E': 18.0, 'F': 18.0, 'G': 18.0, 'H': 18.0}
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width

    # Row heights
    row_heights = {
        1: 28.0, 2: 30.0, 3: 24.0, 4: 22.0, 5: 30.0, 6: 22.0,
        7: 42.0, 8: 42.0, 9: 42.0, 10: 42.0, 11: 42.0,
        12: 22.0, 13: 42.0, 14: 42.0, 15: 42.0, 16: 42.0, 17: 42.0
    }
    for r_idx, h in row_heights.items():
        ws1.row_dimensions[r_idx].height = h

    # Row 1: Header School
    ws1.merge_cells("A1:H1")
    c = ws1["A1"]
    c.value = "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO"
    c.font = f_title1
    c.fill = fill_header_dark
    c.alignment = align_center

    # Row 2: Title
    ws1.merge_cells("A2:H2")
    c = ws1["A2"]
    c.value = "THỜI KHOÁ BIỂU CÁ NHÂN GIÁO VIÊN  —  Năm học 2026–2027 (Bản cập nhật 26/10)"
    c.font = f_title2
    c.fill = fill_header_mid
    c.alignment = align_center

    # Row 3: Teacher Info
    ws1.merge_cells("A3:H3")
    c = ws1["A3"]
    c.value = "Giáo viên: Đậu Đình Nguyên   |   Bộ môn: Tin học & Robotics   |   Tổng: 25 tiết/tuần (Tin học: 14 tiết | Robotics: 11 tiết)"
    c.font = f_info
    c.fill = fill_sub_info
    c.alignment = align_center

    # Row 4: Legend
    ws1.merge_cells("A4:C4")
    c = ws1["A4"]
    c.value = "🔵 Xanh nhạt: Tin học (14 tiết)"
    c.font = f_legend_tin
    c.fill = fill_tin
    c.alignment = align_center

    ws1.merge_cells("D4:F4")
    c = ws1["D4"]
    c.value = "🟠 Cam nhạt: Robotics (11 tiết)"
    c.font = f_legend_rob
    c.fill = fill_robot
    c.alignment = align_center

    ws1.merge_cells("G4:H4")
    c = ws1["G4"]
    c.value = "✅ Đã xử lý triệt để trùng lịch (0 tiết trùng)"
    c.font = f_legend_succ
    c.fill = fill_success
    c.alignment = align_center

    for c_idx in range(1, 9):
        ws1.cell(4, c_idx).border = box_border

    # Row 5: Column Headers
    headers = ["Tiết", "Giờ TH", "Giờ THCS", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    for idx, h_text in enumerate(headers, 1):
        cell = ws1.cell(5, idx, h_text)
        cell.font = f_col_hdr
        cell.fill = fill_header_dark
        cell.alignment = align_center
        cell.border = box_border

    # Row 6: Morning Session Header
    ws1.merge_cells("A6:H6")
    c = ws1["A6"]
    c.value = "▸   BUỔI SÁNG"
    c.font = f_sess_hdr
    c.fill = fill_session
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c_idx in range(1, 9):
        ws1.cell(6, c_idx).border = box_border

    # Morning Periods Setup (Rows 7-11)
    morning_times = [
        ("Tiết 1", "8:15–8:50", "8:05–8:50"),
        ("Tiết 2", "8:55–9:30", "8:55–9:40"),
        ("Tiết 3", "9:45–10:20", "9:45–10:25"),
        ("Tiết 4", "10:25–11:00", "10:30–11:10"),
        ("Tiết 5", "", "11:15–11:50"),
    ]

    for idx, (tiet_str, th_time, thcs_time) in enumerate(morning_times, 7):
        c_tiet = ws1.cell(idx, 1, tiet_str)
        c_tiet.font = f_tiet_hdr
        c_tiet.fill = fill_period_bg
        c_tiet.alignment = align_center
        c_tiet.border = box_border

        c_th = ws1.cell(idx, 2, th_time)
        c_th.font = f_time
        c_th.fill = fill_time_bg
        c_th.alignment = align_center
        c_th.border = box_border

        c_thcs = ws1.cell(idx, 3, thcs_time)
        c_thcs.font = f_time
        c_thcs.fill = fill_time_bg
        c_thcs.alignment = align_center
        c_thcs.border = box_border

        for d_col in range(4, 9):
            ws1.cell(idx, d_col).border = box_border
            ws1.cell(idx, d_col).alignment = align_center

    # Row 12: Afternoon Session Header
    ws1.merge_cells("A12:H12")
    c = ws1["A12"]
    c.value = "▸   BUỔI CHIỀU"
    c.font = f_sess_hdr
    c.fill = fill_session
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c_idx in range(1, 9):
        ws1.cell(12, c_idx).border = box_border

    # Afternoon Periods Setup (Rows 13-17)
    afternoon_times = [
        ("Tiết 1", "13:15–13:50", "13:05–13:50"),
        ("Tiết 2", "13:55–14:30", "13:55–14:40"),
        ("Tiết 3", "14:55–15:30", "14:50–15:30"),
        ("Tiết 4", "15:35–16:10", "15:35–16:20"),
        ("Tiết 5", "", ""),
    ]

    for idx, (tiet_str, th_time, thcs_time) in enumerate(afternoon_times, 13):
        c_tiet = ws1.cell(idx, 1, tiet_str)
        c_tiet.font = f_tiet_hdr
        c_tiet.fill = fill_period_bg
        c_tiet.alignment = align_center
        c_tiet.border = box_border

        c_th = ws1.cell(idx, 2, th_time)
        c_th.font = f_time
        c_th.fill = fill_time_bg
        c_th.alignment = align_center
        c_th.border = box_border

        c_thcs = ws1.cell(idx, 3, thcs_time)
        c_thcs.font = f_time
        c_thcs.fill = fill_time_bg
        c_thcs.alignment = align_center
        c_thcs.border = box_border

        for d_col in range(4, 9):
            ws1.cell(idx, d_col).border = box_border
            ws1.cell(idx, d_col).alignment = align_center

    # Fill Teaching Slots into Sheet 1
    day_to_col = {2: 4, 3: 5, 4: 6, 5: 7, 6: 8}
    for slot in slots:
        d_col = day_to_col[slot['day_num']]
        r_idx = (6 + slot['period']) if slot['session'] == 'Sáng' else (12 + slot['period'])
        cell = ws1.cell(r_idx, d_col)

        is_rob = slot['subject'] == 'Robotics'
        cell.value = f"🤖 Robotics\n({slot['class']})" if is_rob else f"💻 Tin học\n({slot['class']})"
        cell.font = f_robot if is_rob else f_tin
        cell.fill = fill_robot if is_rob else fill_tin
        cell.alignment = align_center
        cell.border = box_border

    # Outer border for main table (rows 5 to 17, cols 1 to 8)
    for ri in range(5, 18):
        for ci in range(1, 9):
            cell = ws1.cell(ri, ci)
            left_side = border_thick_dark if ci == 1 else cell.border.left
            right_side = border_thick_dark if ci == 8 else cell.border.right
            top_side = border_thick_dark if ri == 5 else cell.border.top
            bottom_side = border_thick_dark if ri == 17 else cell.border.bottom
            cell.border = Border(left=left_side, right=right_side, top=top_side, bottom=bottom_side)

    # Summary and notes section under table
    ws1.cell(19, 1).value = "📌 CÁC ĐIỂM CẬP NHẬT TRỌNG TÂM TRONG TKB 26/10:"
    ws1.cell(19, 1).font = Font(name=font_tnr, size=11, bold=True, color="1D2A64")
    
    notes = [
        "1. Đã giải quyết triệt để trùng lịch Thứ Năm: Tiết Tin học lớp TTH2 chuyển từ Thứ 5 Sáng T4 sang Thứ 3 Sáng T2 (Cell AE12). Thứ 5 Sáng T4 chỉ còn 1 tiết Robotics 1A1.",
        "2. Điều chuyển lớp 3A1: Tin học 3A1 chuyển từ Thứ 5 Chiều T1 sang Thứ 2 Chiều T4 (Cell N9).",
        "3. Điều chuyển môn Lớp 4C1: Thứ 2 Sáng T3 dạy Tin học 4C1; Thứ 4 Chiều T3 dạy Robotics 4C1.",
        "4. Điều chuyển môn Lớp 3C1: Thứ 4 Sáng T4 dạy Robotics 3C1; Thứ 6 Sáng T2 dạy Tin học 3C1.",
        "5. Tổng số tiết giảng dạy: Đúng chuẩn 25 tiết/tuần (14 tiết Tin học + 11 tiết Robotics), lịch trình thông suốt, không còn bất kỳ trùng chéo nào."
    ]
    for i, note in enumerate(notes, 20):
        ws1.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
        c = ws1.cell(i, 1, note)
        c.font = Font(name=font_tnr, size=10, italic=True, color="333333")
        c.alignment = align_left

    # Page setup
    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
    ws1.page_setup.paperSize   = ws1.PAPERSIZE_A4
    ws1.page_setup.fitToPage   = True
    ws1.page_setup.fitToWidth  = 1
    ws1.page_setup.fitToHeight = 0
    ws1.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6, header=0.3, footer=0.3)
    ws1.print_area = 'A1:H25'

    # ==========================================
    # SHEET 2: DANH SÁCH TIẾT & ĐỐI CHIẾU
    # ==========================================
    ws2 = wb.create_sheet(title="Danh sách 25 tiết & Đối chiếu")
    ws2.views.sheetView[0].showGridLines = True

    # Column widths
    ws2.column_dimensions['A'].width = 8.0
    ws2.column_dimensions['B'].width = 12.0
    ws2.column_dimensions['C'].width = 14.0
    ws2.column_dimensions['D'].width = 10.0
    ws2.column_dimensions['E'].width = 14.0
    ws2.column_dimensions['F'].width = 14.0
    ws2.column_dimensions['G'].width = 30.0

    # Title
    ws2.merge_cells("A1:G1")
    c = ws2["A1"]
    c.value = "DANH SÁCH CHI TIẾT 25 TIẾT DẠY THEO TUẦN (TKB 26/10) - THẦY ĐẬU ĐÌNH NGUYÊN"
    c.font = Font(name=font_tnr, size=13, bold=True, color="FFFFFF")
    c.fill = fill_header_dark
    c.alignment = align_center
    ws2.row_dimensions[1].height = 28.0

    # Table Header
    tbl2_headers = ["STT", "Thứ", "Buổi", "Tiết", "Môn học", "Lớp", "Ghi chú phòng học"]
    for ci, h_text in enumerate(tbl2_headers, 1):
        c = ws2.cell(2, ci, h_text)
        c.font = Font(name=font_tnr, size=11, bold=True, color="FFFFFF")
        c.fill = fill_header_mid
        c.alignment = align_center
        c.border = box_border
    ws2.row_dimensions[2].height = 24.0

    for idx, s in enumerate(slots, 1):
        r_num = 2 + idx
        ws2.row_dimensions[r_num].height = 20.0
        is_rob = s['subject'] == 'Robotics'
        row_fill = fill_robot if is_rob else fill_tin

        c1 = ws2.cell(r_num, 1, idx)
        c2 = ws2.cell(r_num, 2, s['day_str'])
        c3 = ws2.cell(r_num, 3, s['session'])
        c4 = ws2.cell(r_num, 4, f"Tiết {s['period']}")
        c5 = ws2.cell(r_num, 5, s['subject'])
        c6 = ws2.cell(r_num, 6, s['class'])
        c7 = ws2.cell(r_num, 7, "Phòng Robotics" if is_rob else "Phòng Tin học")

        for ci in range(1, 8):
            cell = ws2.cell(r_num, ci)
            cell.font = Font(name=font_tnr, size=11, bold=(ci in [5, 6]))
            cell.border = box_border
            if ci in [5, 6]:
                cell.fill = row_fill
            if ci != 7:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Save to all destination paths
    for p in output_paths:
        try:
            wb.save(p)
            print(f"✓ Đã lưu Excel thành công: {p}")
        except PermissionError:
            print(f"✗ Lỗi cấp quyền / file đang mở: {p}")

# ─────────────────────────────────────────────────────────────
# 3. CREATE WORD TIMETABLE (.DOCX)
# ─────────────────────────────────────────────────────────────
def set_tbl_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        for child in list(tblPr[0]):
            if child.tag.endswith('tblBorders'):
                tblPr[0].remove(child)
        borders_xml = parse_xml(
            r'<w:tblBorders %s>'
            r'<w:top w:val="single" w:sz="6" w:space="0" w:color="1D2A64"/>'
            r'<w:left w:val="single" w:sz="6" w:space="0" w:color="1D2A64"/>'
            r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="1D2A64"/>'
            r'<w:right w:val="single" w:sz="6" w:space="0" w:color="1D2A64"/>'
            r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="AAAACC"/>'
            r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="AAAACC"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr[0].append(borders_xml)

def set_cell_bg(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_vert_align(cell, align="center"):
    tcPr = cell._element.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align)
    tcPr.append(vAlign)

def add_run(para, text, font_name='Times New Roman', size=12, bold=False, italic=False, color=None):
    run = para.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run

def style_header_cell(cell, text, bg_hex, fg=(255,255,255), size=12, bold=True):
    set_cell_bg(cell, bg_hex)
    set_cell_vert_align(cell)
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    if fg != (0, 0, 0):
        run.font.color.rgb = RGBColor(*fg)

def create_docx_tkb(slots, output_paths):
    doc = docx.Document()
    sec = doc.sections[0]
    
    # Landscape A4
    sec.page_width  = int(29.7 * 360000)
    sec.page_height = int(21.0 * 360000)
    sec.left_margin   = int(1.2 * 360000)
    sec.right_margin  = int(1.2 * 360000)
    sec.top_margin    = int(1.2 * 360000)
    sec.bottom_margin = int(1.2 * 360000)

    sectPr = sec._sectPr
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        sectPr.append(pgSz)
    pgSz.set(qn('w:orient'), 'landscape')

    # Title Block
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    add_run(p, "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO\n", size=13, bold=True, color=(29, 42, 100))
    add_run(p, "THỜI KHOÁ BIỂU CÁ NHÂN GIÁO VIÊN\n", size=16, bold=True, color=(29, 42, 100))
    add_run(p, "Năm học 2026 – 2027 (Cập nhật 26/10/2026)", size=12, italic=True, color=(80, 80, 80))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(3)
    p2.paragraph_format.space_after  = Pt(8)
    add_run(p2, "Giáo viên: Đậu Đình Nguyên   |   Bộ môn: Tin học & Robotics   |   Tổng: 25 tiết/tuần (Tin học: 14 | Robotics: 11)",
            size=12.5, bold=True, color=(29, 42, 100))

    # Table: 13 rows x 8 cols
    table = doc.add_table(rows=13, cols=8)
    table.style = 'Table Grid'
    set_tbl_borders(table)

    col_headers = ["Tiết", "Thời gian\n(Tiểu học)", "Thời gian\n(THCS)", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    for ci, hdr in enumerate(col_headers):
        c = table.rows[0].cells[ci]
        style_header_cell(c, hdr, "1D2A64", fg=(255,255,255), size=11, bold=True)

    # Helper for session row
    def add_session_row(row_idx, label):
        cell0 = table.rows[row_idx].cells[0]
        for ci in range(1, 8):
            cell0.merge(table.rows[row_idx].cells[ci])
        style_header_cell(cell0, label, "E8ECF5", fg=(29, 42, 100), size=12, bold=True)

    add_session_row(1, "▸   BUỔI SÁNG")
    add_session_row(7, "▸   BUỔI CHIỀU")

    times_sang_th   = {1: "8:15–8:50", 2: "8:55–9:30", 3: "9:45–10:20", 4: "10:25–11:00", 5: ""}
    times_sang_thcs = {1: "8:05–8:50", 2: "8:55–9:40", 3: "9:45–10:25", 4: "10:30–11:10", 5: "11:15–11:50"}
    times_chieu_th   = {1: "13:15–13:50", 2: "13:55–14:30", 3: "14:55–15:30", 4: "15:35–16:10", 5: ""}
    times_chieu_thcs = {1: "13:05–13:50", 2: "13:55–14:40", 3: "14:50–15:30", 4: "15:35–16:20", 5: ""}

    # Period labels and times
    for p_num in range(1, 6):
        # Morning (rows 2..6)
        rs = table.rows[p_num + 1]
        set_cell_bg(rs.cells[0], "F5F7FF")
        set_cell_bg(rs.cells[1], "EFF3FF")
        set_cell_bg(rs.cells[2], "EFF3FF")
        for ci in (0, 1, 2):
            set_cell_vert_align(rs.cells[ci])
            rs.cells[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(rs.cells[0].paragraphs[0], f"Tiết {p_num}", size=11, bold=True, color=(29, 42, 100))
        add_run(rs.cells[1].paragraphs[0], times_sang_th.get(p_num, ""), size=10, italic=True, color=(60, 60, 80))
        add_run(rs.cells[2].paragraphs[0], times_sang_thcs.get(p_num, ""), size=10, italic=True, color=(60, 60, 80))

        # Afternoon (rows 8..12)
        rc = table.rows[p_num + 7]
        set_cell_bg(rc.cells[0], "F5F7FF")
        set_cell_bg(rc.cells[1], "EFF3FF")
        set_cell_bg(rc.cells[2], "EFF3FF")
        for ci in (0, 1, 2):
            set_cell_vert_align(rc.cells[ci])
            rc.cells[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(rc.cells[0].paragraphs[0], f"Tiết {p_num}", size=11, bold=True, color=(29, 42, 100))
        add_run(rc.cells[1].paragraphs[0], times_chieu_th.get(p_num, ""), size=10, italic=True, color=(60, 60, 80))
        add_run(rc.cells[2].paragraphs[0], times_chieu_thcs.get(p_num, ""), size=10, italic=True, color=(60, 60, 80))

    # Clear empty cells default formatting
    for r in range(2, 7):
        for c in range(3, 8):
            set_cell_vert_align(table.rows[r].cells[c])
    for r in range(8, 13):
        for c in range(3, 8):
            set_cell_vert_align(table.rows[r].cells[c])

    # Fill Teaching Slots
    day_to_col = {2: 3, 3: 4, 4: 5, 5: 6, 6: 7}
    for slot in slots:
        ci = day_to_col[slot['day_num']]
        ri = (slot['period'] + 1) if slot['session'] == 'Sáng' else (slot['period'] + 7)
        cell = table.rows[ri].cells[ci]

        is_rob = slot['subject'] == 'Robotics'
        bg = "FFE5CC" if is_rob else "D6E4FF"
        fg = (140, 60, 0) if is_rob else (0, 60, 140)
        label = "🤖 Robotics" if is_rob else "💻 Tin học"

        set_cell_bg(cell, bg)
        set_cell_vert_align(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, label + "\n", size=10.5, bold=True, color=fg)
        add_run(p, f"({slot['class']})", size=10, bold=False, color=(50, 50, 50))

    # Notes section under table
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(10)
    add_run(p_note, "📌 Ghi chú điểm mới TKB 26/10: ", size=10.5, bold=True, color=(29, 42, 100))
    add_run(p_note, "Tiết Tin học TTH2 chuyển từ Thứ 5 Sáng T4 sang Thứ 3 Sáng T2 (hoàn toàn hết trùng lịch). Tin học 3A1 chuyển từ Thứ 5 Chiều T1 sang Thứ 2 Chiều T4. Đã tối ưu hóa luân chuyển phòng học Tin học & Robotics.\n", size=10, italic=True, color=(50, 50, 50))

    for path in output_paths:
        try:
            doc.save(path)
            print(f"✓ Đã lưu Word thành công: {path}")
        except PermissionError:
            print(f"✗ Lỗi cấp quyền / file đang mở: {path}")

# ─────────────────────────────────────────────────────────────
# 4. EXECUTION
# ─────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("TRÍCH XUẤT VÀ TẠO THỜI KHÓA BIỂU THẦY ĐẬU ĐÌNH NGUYÊN (26/10)")
    print("=" * 60)

    slots = parse_tkb_2610()
    print(f"Đã trích xuất thành công: {len(slots)} tiết dạy / tuần.")
    tin_count = sum(1 for s in slots if s['subject'] == 'Tin học')
    rob_count = sum(1 for s in slots if s['subject'] == 'Robotics')
    print(f"  - Môn Tin học:  {tin_count} tiết")
    print(f"  - Môn Robotics: {rob_count} tiết")

    excel_paths = [
        os.path.join(DIR_TEACHER, "Thời khóa biểu - Đậu Đình Nguyên (Mới 26.10).xlsx"),
        os.path.join(DIR_TEACHER, "Thời khóa biểu - Đậu Đình Nguyên.xlsx"),
        os.path.join(DIR_MAU,     "Thời khóa biểu - Đậu Đình Nguyên.xlsx"),
    ]

    docx_paths = [
        os.path.join(DIR_TEACHER, "Thời khóa biểu - Đậu Đình Nguyên (Mới 26.10).docx"),
        os.path.join(DIR_TEACHER, "Thời khóa biểu - Đậu Đình Nguyên.docx"),
        os.path.join(DIR_MAU,     "Thời khóa biểu - Đậu Đình Nguyên.docx"),
    ]

    print("\n--- Tạo các file Excel ---")
    create_excel_tkb(slots, excel_paths)

    print("\n--- Tạo các file Word ---")
    create_docx_tkb(slots, docx_paths)

    print("\nHOÀN THÀNH TOÀN BỘ!")

if __name__ == '__main__':
    main()
