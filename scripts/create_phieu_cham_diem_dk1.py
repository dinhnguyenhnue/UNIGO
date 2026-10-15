r"""
Tạo Phiếu chấm điểm Đánh giá định kỳ 1 (ĐGĐK1) cho môn Tin học và Robotics (Lớp 6, 7, 8)
Dựa trên form mẫu: C:\Users\bmngu\Downloads\Phiếu chấm điểm ĐGĐK1-Lớp 6-7-8.xlsx
Lưu tại: D:\UNIGO\Check_các_cuộc_thi\Thi_DK1\Phiếu chấm điểm ĐGĐK1 - Tin học & Robotics - Lớp 6-7-8.xlsx
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Danh sách học sinh chính thức
STUDENTS = {
    '6A1': [
        'Dương Minh Đức',
        'Giang Minh Đức',
        'Nguyễn Ánh Dương',
        'Nguyễn Bùi Khởi Minh',
        'Nguyễn Nhật Tiến',
        'Phạm Anh Duy',
        'Phí Thị Như Quỳnh',
        'Đinh Hoàng Tuấn Tú',
    ],
    '7A1': [
        'Lê Ngọc Hân',
        'Đỗ Nguyễn Hà Minh',
        'Trần Hải Phong',
        'Nguyễn Hoàng Hương Trà',
        'Bùi Mai Phương',
        'Trần Đức Tài',
        'Đỗ Trọng Hoàng',
    ],
    '8A1': [
        'Trần Lê Khánh An',
        'Nguyễn Quỳnh Chi',
        'Nguyễn Trần Khánh Dương',
        'Lê Bảo Minh',
        'Nguyễn Quang Hải Nam',
        'Trần An Như',
        'Đỗ Tuệ Linh',
        'Nguyễn Minh Thư',
        'Phan Đức Việt',
        'Nguyễn Nam Phong',
    ],
}

SHEETS_CONFIG = [
    {'title': 'Tin học 6', 'mon': 'Tin học', 'lop': '6A1'},
    {'title': 'Tin học 7', 'mon': 'Tin học', 'lop': '7A1'},
    {'title': 'Tin học 8', 'mon': 'Tin học', 'lop': '8A1'},
    {'title': 'Robotics 6', 'mon': 'Robotics', 'lop': '6A1'},
    {'title': 'Robotics 7', 'mon': 'Robotics', 'lop': '7A1'},
    {'title': 'Robotics 8', 'mon': 'Robotics', 'lop': '8A1'},
]

# Styles
FONT_NAME = 'Times New Roman'
FONT_TITLE_BOLD = Font(name=FONT_NAME, size=12, bold=True)
FONT_BOLD_13 = Font(name=FONT_NAME, size=13, bold=True)
FONT_HEADER_BOLD = Font(name=FONT_NAME, size=12, bold=True)
FONT_REGULAR = Font(name=FONT_NAME, size=12, bold=False)
FONT_ITALIC = Font(name=FONT_NAME, size=12, italic=True)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')

THIN_SIDE = Side(border_style='thin', color='000000')
BORDER_ALL = Border(top=THIN_SIDE, left=THIN_SIDE, right=THIN_SIDE, bottom=THIN_SIDE)

COL_WIDTHS = {
    'A': 4.0,
    'B': 9.5,
    'C': 20.0,
    'D': 2.0,
    'E': 10.0,
    'F': 20.0,
    'G': 15.0,
}

ROW_HEIGHT_DEFAULT = 18.0


def set_cell(ws, row, col, value=None, font=FONT_REGULAR, alignment=ALIGN_CENTER, border=None):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    return cell


def apply_border_box(ws, start_row, start_col, end_row, end_col):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL


def create_sheet(ws, sheet_cfg):
    mon = sheet_cfg['mon']
    lop = sheet_cfg['lop']
    students = STUDENTS[lop]
    n_students = len(students)

    # Set column widths
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Row 1: UỶ BAN NHÂN DÂN XÃ HOÀI ĐỨC
    ws.row_dimensions[1].height = ROW_HEIGHT_DEFAULT
    set_cell(ws, 1, 1, 'UỶ BAN NHÂN DÂN XÃ HOÀI ĐỨC', font=FONT_TITLE_BOLD, alignment=ALIGN_LEFT)

    # Row 2: TRƯỜNG TIỂU HỌC VÀ THCS UNIGO
    ws.row_dimensions[2].height = ROW_HEIGHT_DEFAULT
    set_cell(ws, 2, 1, 'TRƯỜNG TIỂU HỌC VÀ THCS UNIGO', font=FONT_TITLE_BOLD, alignment=ALIGN_LEFT)

    # Row 3: Blank
    ws.row_dimensions[3].height = ROW_HEIGHT_DEFAULT

    # Row 4: PHIẾU CHẤM ĐIỂM THEO PHÁCH
    ws.row_dimensions[4].height = ROW_HEIGHT_DEFAULT
    ws.merge_cells('A4:G4')
    set_cell(ws, 4, 1, 'PHIẾU CHẤM ĐIỂM THEO PHÁCH', font=FONT_BOLD_13, alignment=ALIGN_CENTER)

    # Row 5: Blank
    ws.row_dimensions[5].height = ROW_HEIGHT_DEFAULT

    # Row 6: Thời điểm: Kiểm tra định kì lần 1          Năm học: 2026-2027
    ws.row_dimensions[6].height = ROW_HEIGHT_DEFAULT
    ws.merge_cells('A6:E6')
    set_cell(ws, 6, 1, 'Thời điểm: Kiểm tra định kì lần 1         ', font=FONT_HEADER_BOLD, alignment=ALIGN_LEFT)
    ws.merge_cells('F6:G6')
    set_cell(ws, 6, 6, 'Năm học: 2026-2027', font=FONT_HEADER_BOLD, alignment=ALIGN_LEFT)

    # Row 7: Môn: ...                                   Lớp: ...
    ws.row_dimensions[7].height = ROW_HEIGHT_DEFAULT
    ws.merge_cells('A7:E7')
    set_cell(ws, 7, 1, f'Môn: {mon}', font=FONT_HEADER_BOLD, alignment=ALIGN_LEFT)
    ws.merge_cells('F7:G7')
    set_cell(ws, 7, 6, f'Lớp: {lop}', font=FONT_HEADER_BOLD, alignment=ALIGN_LEFT)

    # Table Header (Row 8 & 9)
    ws.row_dimensions[8].height = ROW_HEIGHT_DEFAULT
    ws.row_dimensions[9].height = ROW_HEIGHT_DEFAULT

    # TT (A8:B9)
    ws.merge_cells('A8:B9')
    set_cell(ws, 8, 1, 'TT', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)

    # Họ và tên (C8:D9)
    ws.merge_cells('C8:D9')
    set_cell(ws, 8, 3, 'Họ và tên', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)

    # Tổng điểm (E8:F8)
    ws.merge_cells('E8:F8')
    set_cell(ws, 8, 5, 'Tổng  điểm', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)

    # Bằng số (E9), Bằng chữ (F9)
    set_cell(ws, 9, 5, 'Bằng số', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)
    set_cell(ws, 9, 6, 'Bằng chữ', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)

    # Ghi chú (G8:G9)
    ws.merge_cells('G8:G9')
    set_cell(ws, 8, 7, 'Ghi chú', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)

    apply_border_box(ws, 8, 1, 9, 7)

    # Student rows (Row 10 to Row 10 + n_students - 1)
    start_student_row = 10
    end_student_row = start_student_row + n_students - 1

    for idx, student_name in enumerate(students, start=1):
        r = start_student_row + idx - 1
        ws.row_dimensions[r].height = ROW_HEIGHT_DEFAULT

        # TT (merge A:B)
        ws.merge_cells(f'A{r}:B{r}')
        set_cell(ws, r, 1, idx, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # Họ và tên (merge C:D)
        ws.merge_cells(f'C{r}:D{r}')
        set_cell(ws, r, 3, student_name, font=FONT_REGULAR, alignment=Alignment(horizontal='left', vertical='center', indent=1))

        # Bằng số (E)
        set_cell(ws, r, 5, None, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # Bằng chữ (F)
        set_cell(ws, r, 6, None, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # Ghi chú (G)
        set_cell(ws, r, 7, None, font=FONT_REGULAR, alignment=ALIGN_LEFT)

        apply_border_box(ws, r, 1, r, 7)

    # Row Tổng số bài & Ngày tháng
    row_sum = end_student_row + 1
    ws.row_dimensions[row_sum].height = ROW_HEIGHT_DEFAULT
    ws.merge_cells(f'A{row_sum}:C{row_sum}')
    set_cell(ws, row_sum, 1, f'Tổng số bài: {n_students}', font=FONT_REGULAR, alignment=ALIGN_LEFT)

    ws.merge_cells(f'E{row_sum}:G{row_sum}')
    set_cell(ws, row_sum, 5, '                Hoài Đức, ngày 09 tháng 10 năm 2026', font=FONT_REGULAR, alignment=ALIGN_RIGHT)

    # Row Số học sinh vắng & Chữ ký
    row_sign = row_sum + 1
    ws.row_dimensions[row_sign].height = ROW_HEIGHT_DEFAULT
    ws.merge_cells(f'A{row_sign}:C{row_sign}')
    set_cell(ws, row_sign, 1, 'Số học sinh vắng: 0', font=FONT_REGULAR, alignment=ALIGN_LEFT)

    ws.merge_cells(f'E{row_sign}:G{row_sign}')
    set_cell(ws, row_sign, 5, '             Cán bộ chấm thi (Kí, ghi rõ họ tên)', font=FONT_TITLE_BOLD, alignment=ALIGN_CENTER)

    # Phổ điểm table
    row_pho_hdr = row_sign + 1
    ws.row_dimensions[row_pho_hdr].height = ROW_HEIGHT_DEFAULT

    # Header phổ điểm: TT | Phổ điểm | Số lượng (C:D) | %
    set_cell(ws, row_pho_hdr, 1, 'TT', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)
    set_cell(ws, row_pho_hdr, 2, 'Phổ điểm', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)
    ws.merge_cells(f'C{row_pho_hdr}:D{row_pho_hdr}')
    set_cell(ws, row_pho_hdr, 3, 'Số lượng', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)
    set_cell(ws, row_pho_hdr, 5, '%', font=FONT_HEADER_BOLD, alignment=ALIGN_CENTER)
    apply_border_box(ws, row_pho_hdr, 1, row_pho_hdr, 5)

    bands = [
        (1, '8-10', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=8", $E${start_student_row}:$E${end_student_row}, "<=10")'),
        (2, '6.5-<8', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=6.5", $E${start_student_row}:$E${end_student_row}, "<8")'),
        (3, '5-<6.5', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=5", $E${start_student_row}:$E${end_student_row}, "<6.5")'),
        (4, '3.5-<5', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=3.5", $E${start_student_row}:$E${end_student_row}, "<5")'),
        (5, '2-<3.5', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=2", $E${start_student_row}:$E${end_student_row}, "<3.5")'),
        (6, '0-<2', f'=COUNTIFS($E${start_student_row}:$E${end_student_row}, ">=0", $E${start_student_row}:$E${end_student_row}, "<2")'),
    ]

    for band_idx, band_label, formula_count in bands:
        r_band = row_pho_hdr + band_idx
        ws.row_dimensions[r_band].height = ROW_HEIGHT_DEFAULT

        # TT (Col A)
        set_cell(ws, r_band, 1, band_idx, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # Phổ điểm (Col B)
        set_cell(ws, r_band, 2, band_label, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # Số lượng (Col C:D merge)
        ws.merge_cells(f'C{r_band}:D{r_band}')
        set_cell(ws, r_band, 3, formula_count, font=FONT_REGULAR, alignment=ALIGN_CENTER)

        # % (Col E)
        formula_pct = f'=IF(COUNT($E${start_student_row}:$E${end_student_row})>0, C{r_band}/COUNT($E${start_student_row}:$E${end_student_row}), 0)'
        cell_pct = set_cell(ws, r_band, 5, formula_pct, font=FONT_REGULAR, alignment=ALIGN_CENTER)
        cell_pct.number_format = '0.0%'

        apply_border_box(ws, r_band, 1, r_band, 5)


def main():
    output_dir = r'D:\UNIGO\Check_các_cuộc_thi\Thi_DK1'
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, 'Phiếu chấm điểm ĐGĐK1 - Tin học & Robotics - Lớp 6-7-8.xlsx')

    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    for i, cfg in enumerate(SHEETS_CONFIG):
        ws = wb.create_sheet(title=cfg['title'])
        create_sheet(ws, cfg)

    wb.remove(default_sheet)
    wb.save(output_file)
    print(f'Successfully created: {output_file}')


if __name__ == '__main__':
    main()
