import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def create_tkb_nguyen_new():
    wb = openpyxl.Workbook()
    
    # Define styles matching UNIGO aesthetic
    font_tnr = "Times New Roman"
    
    # Fills
    fill_header_dark = PatternFill(start_color="1D2A64", end_color="1D2A64", fill_type="solid")
    fill_header_mid = PatternFill(start_color="253580", end_color="253580", fill_type="solid")
    fill_sub_info = PatternFill(start_color="EBF0FF", end_color="EBF0FF", fill_type="solid")
    fill_session = PatternFill(start_color="E8ECF5", end_color="E8ECF5", fill_type="solid")
    fill_period_bg = PatternFill(start_color="F5F7FF", end_color="F5F7FF", fill_type="solid")
    
    fill_tin = PatternFill(start_color="D6E4FF", end_color="D6E4FF", fill_type="solid")
    fill_robot = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
    fill_conflict = PatternFill(start_color="FFD6D6", end_color="FFD6D6", fill_type="solid")
    fill_opt_fix = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    
    # Fonts
    f_white_title1 = Font(name=font_tnr, size=14, bold=True, color="FFFFFF")
    f_white_title2 = Font(name=font_tnr, size=15, bold=True, color="FFFFFF")
    f_info = Font(name=font_tnr, size=12, bold=False, color="1D2A64")
    f_legend_tin = Font(name=font_tnr, size=10, bold=True, color="003C8C")
    f_legend_rob = Font(name=font_tnr, size=10, bold=True, color="8C3C00")
    f_legend_conf = Font(name=font_tnr, size=10, bold=True, color="990000")
    
    f_col_hdr = Font(name=font_tnr, size=12, bold=True, color="FFFFFF")
    f_sess_hdr = Font(name=font_tnr, size=12, bold=True, color="1D2A64")
    f_tiet_hdr = Font(name=font_tnr, size=12, bold=True, color="1D2A64")
    f_time = Font(name=font_tnr, size=10, color="404060")
    
    f_tin = Font(name=font_tnr, size=11, bold=True, color="003C8C")
    f_robot = Font(name=font_tnr, size=11, bold=True, color="8C3C00")
    f_conflict = Font(name=font_tnr, size=10, bold=True, color="B30000")
    
    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    # Borders
    border_thin_gray = Side(style="thin", color="AAAACC")
    border_thick_dark = Side(style="medium", color="1D2A64")
    border_red_thick = Side(style="medium", color="CC0000")
    
    box_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    conflict_border = Border(left=border_red_thick, right=border_red_thick, top=border_red_thick, bottom=border_red_thick)
    
    # ==========================================
    # SHEET 1: THỜI KHÓA BIỂU CÁ NHÂN (24.10)
    # ==========================================
    ws1 = wb.active
    ws1.title = "TKB Đậu Đình Nguyên - 24.10"
    ws1.views.sheetView[0].showGridLines = True
    
    # Column widths
    col_widths = {'A': 9.0, 'B': 13.0, 'C': 13.0, 'D': 17.0, 'E': 17.0, 'F': 17.0, 'G': 20.0, 'H': 17.0}
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width

    # Row heights
    row_heights = {
        1: 28.0, 2: 30.0, 3: 24.0, 4: 20.0, 5: 30.0, 6: 22.0,
        7: 42.0, 8: 42.0, 9: 42.0, 10: 48.0, 11: 42.0,
        12: 22.0, 13: 42.0, 14: 42.0, 15: 42.0, 16: 42.0, 17: 42.0
    }
    for r_idx, h in row_heights.items():
        ws1.row_dimensions[r_idx].height = h

    # Row 1: UNIGO Header
    ws1.merge_cells("A1:H1")
    c = ws1["A1"]
    c.value = "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO"
    c.font = f_white_title1
    c.fill = fill_header_dark
    c.alignment = align_center

    # Row 2: Title
    ws1.merge_cells("A2:H2")
    c = ws1["A2"]
    c.value = "THỜI KHOÁ BIỂU CÁ NHÂN GIÁO VIÊN  —  Năm học 2026–2027 (Bản cập nhật 24/10)"
    c.font = f_white_title2
    c.fill = fill_header_mid
    c.alignment = align_center

    # Row 3: Teacher Info
    ws1.merge_cells("A3:H3")
    c = ws1["A3"]
    c.value = "Giáo viên: Đậu Đình Nguyên   |   Bộ môn: Tin học & Robotics   |   Tổng: 25 tiết/tuần"
    c.font = f_info
    c.fill = fill_sub_info
    c.alignment = align_center

    # Row 4: Legend
    ws1.merge_cells("A4:C4")
    c = ws1["A4"]
    c.value = "🔵 Xanh nhạt: Tin học"
    c.font = f_legend_tin
    c.fill = fill_tin
    c.alignment = align_center

    ws1.merge_cells("D4:F4")
    c = ws1["D4"]
    c.value = "🟠 Cam nhạt: Robotics"
    c.font = f_legend_rob
    c.fill = fill_robot
    c.alignment = align_center

    ws1.merge_cells("G4:H4")
    c = ws1["G4"]
    c.value = "⚠️ Đỏ nhạt: Cảnh báo trùng T5-T4"
    c.font = f_legend_conf
    c.fill = fill_conflict
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
        c_th.alignment = align_center
        c_th.border = box_border

        c_thcs = ws1.cell(idx, 3, thcs_time)
        c_thcs.font = f_time
        c_thcs.alignment = align_center
        c_thcs.border = box_border

        for d_col in range(4, 9):
            ws1.cell(idx, d_col).border = box_border

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
        c_th.alignment = align_center
        c_th.border = box_border

        c_thcs = ws1.cell(idx, 3, thcs_time)
        c_thcs.font = f_time
        c_thcs.alignment = align_center
        c_thcs.border = box_border

        for d_col in range(4, 9):
            ws1.cell(idx, d_col).border = box_border

    # Populate Schedule Data based on TKB toàn trường 24.10
    # Day columns: Thứ Hai=4, Thứ Ba=5, Thứ Tư=6, Thứ Năm=7, Thứ Sáu=8
    # Morning: rows 7-11 (T1-T5), Afternoon: rows 13-17 (T1-T5)

    schedule_data = {
        # Thứ Hai (Col 4)
        (9, 4):  ("robot", "🤖 Robotics\n(4C1)"),
        (10, 4): ("tin",   "💻 Tin học\n(1A1)"),
        (15, 4): ("tin",   "💻 Tin học\n(2C1)"),

        # Thứ Ba (Col 5)
        (7, 5):  ("tin",   "💻 Tin học\n(1C1)"),
        (10, 5): ("robot", "🤖 Robotics\n(5C1)"),
        (15, 5): ("tin",   "💻 Tin học\n(7A1)"),
        (16, 5): ("robot", "🤖 Robotics\n(7A1)"),

        # Thứ Tư (Col 6)
        (8, 6):  ("robot", "🤖 Robotics\n(3A1)"),
        (10, 6): ("tin",   "💻 Tin học\n(3C1)"),
        (15, 6): ("tin",   "💻 Tin học\n(4C1)"),
        (16, 6): ("robot", "🤖 Robotics\n(2A1)"),

        # Thứ Năm (Col 7)
        (7, 7):  ("tin",   "💻 Tin học\n(TT3)"),
        (8, 7):  ("tin",   "💻 Tin học\n(TTH 1)"),    # MỚI (chuyển từ Chiều T2)
        (10, 7): ("conflict", "⚠️ TRÙNG LỊCH:\n🤖 Robotics (1A1)\n💻 Tin học (TTH2)"), # TRÙNG
        (13, 7): ("tin",   "💻 Tin học\n(3A1)"),
        (14, 7): ("tin",   "💻 Tin học\n(5C1)"),      # MỚI (chuyển từ T3 Sáng T3)
        (15, 7): ("robot", "🤖 Robotics\n(2C1)"),     # MỚI (chuyển từ T5 Sáng T4)
        (16, 7): ("robot", "🤖 Robotics\n(1C1)"),

        # Thứ Sáu (Col 8)
        (7, 8):  ("tin",   "💻 Tin học\n(2A1)"),
        (8, 8):  ("robot", "🤖 Robotics\n(3C1)"),
        (10, 8): ("tin",   "💻 Tin học\n(6A1)"),
        (11, 8): ("robot", "🤖 Robotics\n(6A1)"),
        (13, 8): ("tin",   "💻 Tin học\n(8A1)"),
        (14, 8): ("robot", "🤖 Robotics\n(8A1)"),
    }

    for (r, col), (stype, text) in schedule_data.items():
        cell = ws1.cell(r, col, text)
        cell.alignment = align_center
        if stype == "tin":
            cell.font = f_tin
            cell.fill = fill_tin
            cell.border = box_border
        elif stype == "robot":
            cell.font = f_robot
            cell.fill = fill_robot
            cell.border = box_border
        elif stype == "conflict":
            cell.font = f_conflict
            cell.fill = fill_conflict
            cell.border = conflict_border

    # Ghi chú dưới bảng TKB
    r_notes = 19
    ws1.merge_cells(f"A{r_notes}:H{r_notes}")
    c = ws1[f"A{r_notes}"]
    c.value = "📌 CÁC ĐIỂM THAY ĐỔI TRONG TKB 24/10 & CẢNH BÁO TRÙNG LỊCH"
    c.font = Font(name=font_tnr, size=12, bold=True, color="1D2A64")
    c.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    notes = [
        "1. Tin học 5C1: Chuyển từ [Thứ Ba Sáng Tiết 3] sang [Thứ Năm Chiều Tiết 2].",
        "2. Tin học TTH 1: Chuyển từ [Thứ Năm Chiều Tiết 2] lên [Thứ Năm Sáng Tiết 2].",
        "3. Robotics 2C1: Chuyển từ [Thứ Năm Sáng Tiết 4] xuống [Thứ Năm Chiều Tiết 3].",
        "4. Robotics 1A1: Trong TKB cũ xếp Thứ Năm Sáng Tiết 2; trong TKB mới 24/10 xếp Thứ Năm Sáng Tiết 4.",
        "⚠️ 5. CẢNH BÁO TRÙNG LỊCH: Tại [Thứ Năm Sáng Tiết 4], thầy Nguyên bị trùng 2 tiết cùng lúc: Robotics 1A1 và Tin học TTH2.",
        "   👉 Đề xuất phương án tối ưu: Hoán đổi nội bộ lớp TTH2 giữa Tiết 3 (GVVN - Trang) và Tiết 4 (Tin - Nguyên).",
        "      Khi đó: Thầy Nguyên dạy TTH2 vào Tiết 3 (lúc này thầy Nguyên trống) và dạy 1A1 vào Tiết 4 (cô Trang trống T4)."
    ]

    for idx, n_text in enumerate(notes, r_notes + 1):
        ws1.merge_cells(f"A{idx}:H{idx}")
        cell = ws1[f"A{idx}"]
        cell.value = n_text
        if "⚠️" in n_text:
            cell.font = Font(name=font_tnr, size=11, bold=True, color="CC0000")
            cell.fill = PatternFill(start_color="FFF2F2", end_color="FFF2F2", fill_type="solid")
        elif "👉" in n_text:
            cell.font = Font(name=font_tnr, size=10, bold=True, italic=True, color="003366")
            cell.fill = PatternFill(start_color="F2F8FF", end_color="F2F8FF", fill_type="solid")
        else:
            cell.font = Font(name=font_tnr, size=10, color="202040")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws1.row_dimensions[idx].height = 22.0

    # ========================================================
    # SHEET 2: BẢNG ĐỐI CHIẾU CHI TIẾT & PHƯƠNG ÁN ĐÃ XỬ LÝ
    # ========================================================
    ws2 = wb.create_sheet(title="Phương án kiến nghị & Đối chiếu")
    ws2.views.sheetView[0].showGridLines = True

    ws2.column_dimensions['A'].width = 6.0
    ws2.column_dimensions['B'].width = 16.0
    ws2.column_dimensions['C'].width = 14.0
    ws2.column_dimensions['D'].width = 25.0
    ws2.column_dimensions['E'].width = 28.0
    ws2.column_dimensions['F'].width = 25.0
    ws2.column_dimensions['G'].width = 32.0

    # Header Sheet 2
    ws2.merge_cells("A1:G1")
    c = ws2["A1"]
    c.value = "BẢNG ĐỐI CHIẾU TKB CŨ VÀ TKB 24.10 CỦA THẦY ĐẬU ĐÌNH NGUYÊN"
    c.font = f_white_title1
    c.fill = fill_header_dark
    c.alignment = align_center
    ws2.row_dimensions[1].height = 28.0

    headers_s2 = ["STT", "Thứ", "Thời gian", "TKB Cũ", "TKB Mới 24.10", "Trạng thái", "Ghi chú & Đề xuất"]
    ws2.row_dimensions[2].height = 26.0
    for idx, h in enumerate(headers_s2, 1):
        cell = ws2.cell(2, idx, h)
        cell.font = f_col_hdr
        cell.fill = fill_header_mid
        cell.alignment = align_center
        cell.border = box_border

    comparison_data = [
        (1,  "Thứ Hai", "Sáng Tiết 3",   "Robotics (4C1)",  "Robotics (4C1)",        "Giữ nguyên",  "Phòng Robotics"),
        (2,  "Thứ Hai", "Sáng Tiết 4",   "Tin học (1A1)",   "Tin học (1A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (3,  "Thứ Hai", "Chiều Tiết 3",  "Tin học (2C1)",   "Tin học (2C1)",         "Giữ nguyên",  "Phòng Tin học"),
        (4,  "Thứ Ba",  "Sáng Tiết 1",   "Tin học (1C1)",   "Tin học (1C1)",         "Giữ nguyên",  "Phòng Tin học"),
        (5,  "Thứ Ba",  "Sáng Tiết 3",   "Tin học (5C1)",   "(Trống)",               "Đổi buổi",    "Chuyển sang Thứ 5 Chiều Tiết 2"),
        (6,  "Thứ Ba",  "Sáng Tiết 4",   "Robotics (5C1)",  "Robotics (5C1)",        "Giữ nguyên",  "Phòng Robotics"),
        (7,  "Thứ Ba",  "Chiều Tiết 3",  "Tin học (7A1)",   "Tin học (7A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (8,  "Thứ Ba",  "Chiều Tiết 4",  "Robotics (7A1)",  "Robotics (7A1)",        "Giữ nguyên",  "Phòng Robotics"),
        (9,  "Thứ Tư",  "Sáng Tiết 2",   "Robotics (3A1)",  "Robotics (3A1)",        "Giữ nguyên",  "Phòng Robotics"),
        (10, "Thứ Tư",  "Sáng Tiết 4",   "Tin học (3C1)",   "Tin học (3C1)",         "Giữ nguyên",  "Phòng Tin học"),
        (11, "Thứ Tư",  "Chiều Tiết 3",  "Tin học (4C1)",   "Tin học (4C1)",         "Giữ nguyên",  "Phòng Tin học"),
        (12, "Thứ Tư",  "Chiều Tiết 4",  "Robotics (2A1)",  "Robotics (2A1)",        "Giữ nguyên",  "Phòng Robotics"),
        (13, "Thứ Năm", "Sáng Tiết 1",   "Tin học (TT3)",   "Tin học (TT3)",         "Giữ nguyên",  "Phòng Tin học"),
        (14, "Thứ Năm", "Sáng Tiết 2",   "Robotics (1A1)",  "Tin học (TTH 1)",       "Đổi tiết",    "TTH 1 chuyển từ Chiều T2 lên Sáng T2"),
        (15, "Thứ Năm", "Sáng Tiết 3",   "(Trống)",         "(Trống)",               "Trống",       "Đề xuất xếp TTH2 Tin vào đây để tránh trùng"),
        (16, "Thứ Năm", "Sáng Tiết 4",   "Robotics (2C1)",  "⚠️ 1A1 (Rob) & TTH2 (Tin)", "TRÙNG LỊCH", "Trùng 2 tiết! Đổi TTH2 sang Tiết 3"),
        (17, "Thứ Năm", "Chiều Tiết 1",  "Tin học (3A1)",   "Tin học (3A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (18, "Thứ Năm", "Chiều Tiết 2",  "Tin học (TTH 1)", "Tin học (5C1)",         "Đổi tiết",    "5C1 Tin chuyển từ Thứ 3 Sáng T3 sang"),
        (19, "Thứ Năm", "Chiều Tiết 3",  "Tin học (TTH2)",  "Robotics (2C1)",        "Đổi tiết",    "2C1 Robotics chuyển từ Thứ 5 Sáng T4 sang"),
        (20, "Thứ Năm", "Chiều Tiết 4",  "Robotics (1C1)",  "Robotics (1C1)",        "Giữ nguyên",  "Phòng Robotics"),
        (21, "Thứ Sáu", "Sáng Tiết 1",   "Tin học (2A1)",   "Tin học (2A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (22, "Thứ Sáu", "Sáng Tiết 2",   "Robotics (3C1)",  "Robotics (3C1)",        "Giữ nguyên",  "Phòng Robotics"),
        (23, "Thứ Sáu", "Sáng Tiết 4",   "Tin học (6A1)",   "Tin học (6A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (24, "Thứ Sáu", "Sáng Tiết 5",   "Robotics (6A1)",  "Robotics (6A1)",        "Giữ nguyên",  "Phòng Robotics"),
        (25, "Thứ Sáu", "Chiều Tiết 1",  "Tin học (8A1)",   "Tin học (8A1)",         "Giữ nguyên",  "Phòng Tin học"),
        (26, "Thứ Sáu", "Chiều Tiết 2",  "Robotics (8A1)",  "Robotics (8A1)",        "Giữ nguyên",  "Phòng Robotics"),
    ]

    for row_idx, data_tuple in enumerate(comparison_data, 3):
        ws2.row_dimensions[row_idx].height = 22.0
        for col_idx, val in enumerate(data_tuple, 1):
            cell = ws2.cell(row_idx, col_idx, val)
            cell.font = Font(name=font_tnr, size=11, color="101030")
            cell.border = box_border
            
            # Alignments & conditional styling
            if col_idx in (1, 2, 3, 6):
                cell.alignment = align_center
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
            if "TRÙNG" in str(val):
                cell.fill = fill_conflict
                cell.font = Font(name=font_tnr, size=11, bold=True, color="CC0000")
            elif "Đổi" in str(val):
                cell.fill = PatternFill(start_color="FFF8DC", end_color="FFF8DC", fill_type="solid")
            elif "Giữ nguyên" in str(val):
                cell.font = Font(name=font_tnr, size=11, color="006600")

    output_dir = r"D:\UNIGO\Thời khóa biểu giáo viên"
    output_filename = "Thời khóa biểu - Đậu Đình Nguyên (Mới 24.10).xlsx"
    target_path = os.path.join(output_dir, output_filename)
    
    wb.save(target_path)
    print(f"File created successfully at: {target_path}")

if __name__ == "__main__":
    create_tkb_nguyen_new()
