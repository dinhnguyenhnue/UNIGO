#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
check_cuoc_thi.py - Thống kê cuộc thi học sinh UNIGO
=====================================================
Tự động quét thư mục Check_các_cuộc_thi, đọc kết quả thi từ các lần,
so khớp với danh sách HS toàn trường, và xuất file Excel thống kê.

Usage:
    python scripts/check_cuoc_thi.py
"""

import os
import re
import sys
import unicodedata
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    numbers
)
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = Path(r"D:\UNIGO\Check_các_cuộc_thi")
DS_HS_FILE = BASE_DIR / "DANH SÁCH HỌC SINH NĂM HỌC 2026 - 2027.xlsx"
OUTPUT_FILE = BASE_DIR / "Thống kê cuộc thi.xlsx"

# Folders to ignore (not contests)
IGNORE_FOLDERS = {"Thống kê cuộc thi.xlsx"}
IGNORE_PREFIXES = ("~$",)

# ============================================================
# STYLES
# ============================================================
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
HEADER_FONT = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUBHEADER_FONT = Font(name="Times New Roman", size=11, bold=True)
DATA_FONT = Font(name="Times New Roman", size=11)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FONT = Font(name="Times New Roman", size=11, color="006100")
RED_FONT = Font(name="Times New Roman", size=11, color="9C0006")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Times New Roman", size=14, bold=True, color="2F5496")
CLASS_HEADER_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
CLASS_HEADER_FONT = Font(name="Times New Roman", size=12, bold=True, color="BF8F00")

# ============================================================
# HELPERS
# ============================================================

def normalize_name(name: str) -> str:
    """Normalize Vietnamese name for matching: lowercase, strip, collapse spaces."""
    if not name:
        return ""
    name = str(name).strip()
    # Collapse multiple spaces
    name = re.sub(r"\s+", " ", name)
    # Lowercase
    name = name.lower()
    return name


def normalize_class(cls: str) -> str:
    """Normalize class name: uppercase, strip."""
    if not cls:
        return ""
    return str(cls).strip().upper()


def extract_lan_number(folder_name: str) -> int:
    """Extract round number from folder name like 'Lần 1', 'Lần 2'."""
    m = re.search(r"[Ll]ần\s*(\d+)", folder_name)
    if m:
        return int(m.group(1))
    # Try just digits
    m = re.search(r"(\d+)", folder_name)
    if m:
        return int(m.group(1))
    return 0


def find_header_columns(ws):
    """
    Auto-detect column indices for name and class in a contest result sheet.
    Returns (name_col_idx, class_col_idx, extra_cols) — 0-based indices.
    extra_cols is a dict of {col_name: col_idx} for score, round, time etc.
    """
    name_col = None
    class_col = None
    extra_cols = {}

    # Search first 5 rows for header
    for row in ws.iter_rows(min_row=1, max_row=5, values_only=False):
        for cell in row:
            val = str(cell.value or "").strip().lower()
            if any(kw in val for kw in ["họ và tên", "họ tên", "ho va ten", "ho ten"]):
                name_col = cell.column - 1  # 0-based
            elif val in ["lớp", "lop"]:
                class_col = cell.column - 1
            elif val in ["khối", "khoi"]:
                if class_col is None:  # fallback
                    extra_cols["khối"] = cell.column - 1
            # Capture extra info columns
            if any(kw in val for kw in ["tổng điểm", "điểm thi", "diem thi"]):
                extra_cols["điểm"] = cell.column - 1
            if any(kw in val for kw in ["vòng", "vong"]):
                extra_cols["vòng"] = cell.column - 1
            if any(kw in val for kw in ["thời gian", "thoi gian"]):
                extra_cols["thời_gian"] = cell.column - 1

        if name_col is not None:
            break

    return name_col, class_col, extra_cols


# ============================================================
# STEP 1: Read master student list
# ============================================================

def read_master_students(filepath: Path) -> dict:
    """
    Read master student list.
    Returns: {
        'all_students': [(họ_tên, lớp), ...],  # ordered
        'by_class': {'1A1': [họ_tên, ...], ...},
        'classes_order': ['1A1', '1C1', ...]
    }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    all_students = []
    by_class = {}
    classes_order = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        class_name = normalize_class(sheet_name)
        classes_order.append(class_name)
        by_class[class_name] = []

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            # Column A = STT, Column C = Họ tên hs
            stt = row[0]
            if stt is None:
                continue
            stt_str = str(stt).replace(".0", "").strip()
            if not stt_str.isdigit():
                continue

            ho_ten = str(row[2] or "").strip()
            if not ho_ten:
                continue

            all_students.append((ho_ten, class_name))
            by_class[class_name].append(ho_ten)

    wb.close()
    return {
        "all_students": all_students,
        "by_class": by_class,
        "classes_order": classes_order,
    }


# ============================================================
# STEP 2: Scan contest directories
# ============================================================

def scan_contests(base_dir: Path) -> dict:
    """
    Scan contest directories.
    Returns: {
        'IOE': {
            1: [Path_to_excel, ...],  # Lần 1
            2: [Path_to_excel, ...],  # Lần 2
        },
        ...
    }
    """
    contests = {}

    for item in sorted(base_dir.iterdir()):
        if not item.is_dir():
            continue
        if item.name.startswith(tuple(IGNORE_PREFIXES)):
            continue

        contest_name = item.name
        contests[contest_name] = {}

        for sub in sorted(item.iterdir()):
            if not sub.is_dir():
                continue
            lan_num = extract_lan_number(sub.name)
            if lan_num == 0:
                continue

            excel_files = []
            for f in sub.iterdir():
                if f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~$"):
                    excel_files.append(f)

            if excel_files:
                contests[contest_name][lan_num] = excel_files

    return contests


# ============================================================
# STEP 3: Read contest results
# ============================================================

def read_contest_results(excel_files: list) -> list:
    """
    Read contest result files.
    Returns: [(normalized_name, class_name, {extra_info}), ...]
    """
    results = []

    for fpath in excel_files:
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
        except Exception as e:
            print(f"  ⚠ Không đọc được file: {fpath.name} — {e}")
            continue

        for ws in wb.worksheets:
            name_col, class_col, extra_cols = find_header_columns(ws)

            if name_col is None:
                print(f"  ⚠ Không tìm thấy cột tên HS trong sheet '{ws.title}' của file {fpath.name}")
                continue

            # Find header row
            header_row = 1
            for r in range(1, 6):
                cell_val = str(ws.cell(row=r, column=name_col + 1).value or "").strip().lower()
                if any(kw in cell_val for kw in ["họ và tên", "họ tên"]):
                    header_row = r
                    break

            for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
                if name_col >= len(row):
                    continue
                raw_name = row[name_col]
                if not raw_name or str(raw_name).strip() == "":
                    continue

                ho_ten = str(raw_name).strip()
                norm_name = normalize_name(ho_ten)

                # Get class
                cls = ""
                if class_col is not None and class_col < len(row):
                    cls = normalize_class(str(row[class_col] or ""))

                # Get extra info
                info = {"tên_gốc": ho_ten}
                for key, idx in extra_cols.items():
                    if idx < len(row) and row[idx] is not None:
                        info[key] = row[idx]

                results.append((norm_name, cls, info))

        wb.close()

    return results


# ============================================================
# STEP 4: Match students
# ============================================================

def match_students(master_students: list, contest_results: list) -> dict:
    """
    Match contest results to master student list.
    Returns: {(họ_tên, lớp): {extra_info} or None}
    """
    matched = {}

    # Build lookup from contest results
    # Key: normalized_name -> list of (class, info)
    contest_lookup = {}
    for norm_name, cls, info in contest_results:
        if norm_name not in contest_lookup:
            contest_lookup[norm_name] = []
        contest_lookup[norm_name].append((cls, info))

    for ho_ten, lop in master_students:
        norm = normalize_name(ho_ten)
        key = (ho_ten, lop)

        if norm in contest_lookup:
            # Try to match by class first
            found = None
            for cls, info in contest_lookup[norm]:
                if cls == lop or cls == "":
                    found = info
                    break
            if found is None:
                # Fallback: take first match regardless of class
                found = contest_lookup[norm][0][1]
            matched[key] = found
        else:
            matched[key] = None

    return matched


# ============================================================
# STEP 5: Generate output Excel
# ============================================================

def apply_cell_style(cell, font=None, fill=None, alignment=None, border=None):
    """Apply styles to a cell."""
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border


def auto_fit_columns(ws, min_width=8, max_width=40):
    """Auto-fit column widths."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                # Approximate width for Vietnamese characters
                cell_len = len(str(cell.value))
                max_len = max(max_len, cell_len)
        adjusted = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted


def generate_output(master_data, contests_data, all_matches, output_path):
    """Generate the final Excel report."""
    wb = openpyxl.Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    all_students = master_data["all_students"]
    classes_order = master_data["classes_order"]
    by_class = master_data["by_class"]

    # ========================================
    # SHEET 1: TỔNG HỢP
    # ========================================
    ws_tong = wb.create_sheet("Tổng hợp")

    # Build column structure: STT | Họ tên | Lớp | Contest1_Lan1 | Contest1_Lan2 | ...
    contest_columns = []  # [(contest_name, lan_num), ...]
    for contest_name in sorted(contests_data.keys()):
        for lan_num in sorted(contests_data[contest_name].keys()):
            contest_columns.append((contest_name, lan_num))

    # Title row
    ws_tong.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + len(contest_columns))
    title_cell = ws_tong.cell(row=1, column=1, value="THỐNG KÊ CUỘC THI HỌC SINH UNIGO - NĂM HỌC 2026-2027")
    apply_cell_style(title_cell, font=TITLE_FONT, alignment=CENTER)
    ws_tong.row_dimensions[1].height = 30

    # Header row
    row = 3
    headers = ["STT", "Họ và tên", "Lớp"]
    for contest_name, lan_num in contest_columns:
        headers.append(f"{contest_name}\nLần {lan_num}")

    # Summary column
    if contest_columns:
        headers.append("Tổng cuộc thi\nđã tham gia")

    for col_idx, header in enumerate(headers, 1):
        cell = ws_tong.cell(row=row, column=col_idx, value=header)
        apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=CENTER, border=THIN_BORDER)

    ws_tong.row_dimensions[row].height = 35

    # Data rows — grouped by class
    data_row = row + 1
    stt = 0
    for cls in classes_order:
        # Class separator row
        ws_tong.merge_cells(
            start_row=data_row, start_column=1,
            end_row=data_row, end_column=len(headers)
        )
        cls_cell = ws_tong.cell(row=data_row, column=1, value=f"Lớp {cls}")
        apply_cell_style(cls_cell, font=CLASS_HEADER_FONT, fill=CLASS_HEADER_FILL, alignment=LEFT, border=THIN_BORDER)
        # Apply border to merged area
        for c in range(2, len(headers) + 1):
            apply_cell_style(ws_tong.cell(row=data_row, column=c), border=THIN_BORDER)
        ws_tong.row_dimensions[data_row].height = 22
        data_row += 1

        for ho_ten in by_class[cls]:
            stt += 1
            key = (ho_ten, cls)

            ws_tong.cell(row=data_row, column=1, value=stt)
            apply_cell_style(ws_tong.cell(row=data_row, column=1), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

            ws_tong.cell(row=data_row, column=2, value=ho_ten)
            apply_cell_style(ws_tong.cell(row=data_row, column=2), font=DATA_FONT, alignment=LEFT, border=THIN_BORDER)

            ws_tong.cell(row=data_row, column=3, value=cls)
            apply_cell_style(ws_tong.cell(row=data_row, column=3), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

            total_participated = 0
            for col_offset, (contest_name, lan_num) in enumerate(contest_columns):
                match_key = (contest_name, lan_num)
                match_result = all_matches.get(match_key, {}).get(key)

                col = 4 + col_offset
                cell = ws_tong.cell(row=data_row, column=col)

                if match_result is not None:
                    cell.value = "✅ Đã thi"
                    apply_cell_style(cell, font=GREEN_FONT, fill=GREEN_FILL, alignment=CENTER, border=THIN_BORDER)
                    total_participated += 1
                else:
                    cell.value = "❌ Chưa thi"
                    apply_cell_style(cell, font=RED_FONT, fill=RED_FILL, alignment=CENTER, border=THIN_BORDER)

            # Summary column
            if contest_columns:
                sum_col = 4 + len(contest_columns)
                sum_cell = ws_tong.cell(row=data_row, column=sum_col, value=total_participated)
                apply_cell_style(sum_cell, font=Font(name="Times New Roman", size=11, bold=True),
                                alignment=CENTER, border=THIN_BORDER)

            data_row += 1

    # Statistics row
    data_row += 1
    ws_tong.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=3)
    stat_cell = ws_tong.cell(row=data_row, column=1, value="THỐNG KÊ TỔNG:")
    apply_cell_style(stat_cell, font=Font(name="Times New Roman", size=12, bold=True), alignment=LEFT)

    for col_offset, (contest_name, lan_num) in enumerate(contest_columns):
        match_key = (contest_name, lan_num)
        matches = all_matches.get(match_key, {})
        participated = sum(1 for v in matches.values() if v is not None)
        total = len(matches)
        col = 4 + col_offset
        cell = ws_tong.cell(row=data_row, column=col,
                            value=f"{participated}/{total} ({participated*100//total}%)")
        apply_cell_style(cell, font=Font(name="Times New Roman", size=11, bold=True), alignment=CENTER)

    auto_fit_columns(ws_tong)
    # Override specific columns
    ws_tong.column_dimensions["A"].width = 6
    ws_tong.column_dimensions["B"].width = 28
    ws_tong.column_dimensions["C"].width = 8

    # Freeze panes
    ws_tong.freeze_panes = "D4"

    # ========================================
    # SHEET 2+: Chi tiết từng cuộc thi
    # ========================================
    for contest_name in sorted(contests_data.keys()):
        ws_ct = wb.create_sheet(f"Chi tiết {contest_name}")

        # Title
        ws_ct.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
        title = ws_ct.cell(row=1, column=1,
                           value=f"CHI TIẾT KẾT QUẢ {contest_name.upper()} - NĂM HỌC 2026-2027")
        apply_cell_style(title, font=TITLE_FONT, alignment=CENTER)
        ws_ct.row_dimensions[1].height = 30

        row = 3
        for lan_num in sorted(contests_data[contest_name].keys()):
            match_key = (contest_name, lan_num)
            matches = all_matches.get(match_key, {})

            # Lần header
            ws_ct.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            lan_cell = ws_ct.cell(row=row, column=1, value=f"Lần {lan_num}")
            apply_cell_style(lan_cell, font=SUBHEADER_FONT, fill=SUBHEADER_FILL, alignment=LEFT, border=THIN_BORDER)
            for c in range(2, 9):
                apply_cell_style(ws_ct.cell(row=row, column=c), border=THIN_BORDER)
            row += 1

            # Detail headers
            detail_headers = ["STT", "Họ và tên", "Lớp", "Trạng thái", "Vòng", "Điểm", "Thời gian", "Ghi chú"]
            for col_idx, h in enumerate(detail_headers, 1):
                cell = ws_ct.cell(row=row, column=col_idx, value=h)
                apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=CENTER, border=THIN_BORDER)
            row += 1

            # Data - group by class
            stt_detail = 0
            for cls in classes_order:
                for ho_ten in by_class[cls]:
                    stt_detail += 1
                    key = (ho_ten, cls)
                    info = matches.get(key)

                    ws_ct.cell(row=row, column=1, value=stt_detail)
                    apply_cell_style(ws_ct.cell(row=row, column=1), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

                    ws_ct.cell(row=row, column=2, value=ho_ten)
                    apply_cell_style(ws_ct.cell(row=row, column=2), font=DATA_FONT, alignment=LEFT, border=THIN_BORDER)

                    ws_ct.cell(row=row, column=3, value=cls)
                    apply_cell_style(ws_ct.cell(row=row, column=3), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

                    status_cell = ws_ct.cell(row=row, column=4)
                    if info is not None:
                        status_cell.value = "✅ Đã thi"
                        apply_cell_style(status_cell, font=GREEN_FONT, fill=GREEN_FILL, alignment=CENTER, border=THIN_BORDER)

                        # Extra info
                        vong = info.get("vòng", "")
                        diem = info.get("điểm", "")
                        thoi_gian = info.get("thời_gian", "")

                        ws_ct.cell(row=row, column=5, value=vong)
                        ws_ct.cell(row=row, column=6, value=diem)
                        ws_ct.cell(row=row, column=7, value=thoi_gian)
                    else:
                        status_cell.value = "❌ Chưa thi"
                        apply_cell_style(status_cell, font=RED_FONT, fill=RED_FILL, alignment=CENTER, border=THIN_BORDER)

                    for c in range(5, 9):
                        apply_cell_style(ws_ct.cell(row=row, column=c), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

                    row += 1

            row += 1  # Blank row between rounds

        auto_fit_columns(ws_ct)
        ws_ct.column_dimensions["A"].width = 6
        ws_ct.column_dimensions["B"].width = 28
        ws_ct.freeze_panes = "D4"

    # ========================================
    # SHEET LAST: NHẮC NHỞ GV
    # ========================================
    ws_gv = wb.create_sheet("Nhắc nhở GV")

    # Title
    ws_gv.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title = ws_gv.cell(row=1, column=1,
                       value="NHẮC NHỞ GIÁO VIÊN CHỦ NHIỆM - HỌC SINH CHƯA THI")
    apply_cell_style(title, font=TITLE_FONT, alignment=CENTER)
    ws_gv.row_dimensions[1].height = 30

    row = 3
    for cls in classes_order:
        for contest_name in sorted(contests_data.keys()):
            for lan_num in sorted(contests_data[contest_name].keys()):
                match_key = (contest_name, lan_num)
                matches = all_matches.get(match_key, {})

                # Collect students who haven't participated
                chua_thi = []
                for ho_ten in by_class[cls]:
                    key = (ho_ten, cls)
                    if matches.get(key) is None:
                        chua_thi.append(ho_ten)

                if not chua_thi:
                    continue

                total_class = len(by_class[cls])
                da_thi = total_class - len(chua_thi)

                # Class + contest header
                ws_gv.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
                header_text = (
                    f"Lớp {cls} — {contest_name} Lần {lan_num} "
                    f"({da_thi}/{total_class} đã thi, "
                    f"{len(chua_thi)}/{total_class} chưa thi)"
                )
                h_cell = ws_gv.cell(row=row, column=1, value=header_text)
                apply_cell_style(h_cell, font=CLASS_HEADER_FONT, fill=CLASS_HEADER_FILL, alignment=LEFT, border=THIN_BORDER)
                for c in range(2, 6):
                    apply_cell_style(ws_gv.cell(row=row, column=c), border=THIN_BORDER)
                row += 1

                # Sub-headers
                sub_headers = ["STT", "Họ và tên", "Lớp", "Trạng thái", "Ghi chú"]
                for col_idx, h in enumerate(sub_headers, 1):
                    cell = ws_gv.cell(row=row, column=col_idx, value=h)
                    apply_cell_style(cell, font=SUBHEADER_FONT, fill=SUBHEADER_FILL, alignment=CENTER, border=THIN_BORDER)
                row += 1

                # List students
                for idx, ho_ten in enumerate(chua_thi, 1):
                    ws_gv.cell(row=row, column=1, value=idx)
                    apply_cell_style(ws_gv.cell(row=row, column=1), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

                    ws_gv.cell(row=row, column=2, value=ho_ten)
                    apply_cell_style(ws_gv.cell(row=row, column=2), font=DATA_FONT, alignment=LEFT, border=THIN_BORDER)

                    ws_gv.cell(row=row, column=3, value=cls)
                    apply_cell_style(ws_gv.cell(row=row, column=3), font=DATA_FONT, alignment=CENTER, border=THIN_BORDER)

                    status = ws_gv.cell(row=row, column=4, value="❌ Chưa thi")
                    apply_cell_style(status, font=RED_FONT, fill=RED_FILL, alignment=CENTER, border=THIN_BORDER)

                    apply_cell_style(ws_gv.cell(row=row, column=5), font=DATA_FONT, alignment=LEFT, border=THIN_BORDER)

                    row += 1

                row += 1  # Blank row between sections

    auto_fit_columns(ws_gv)
    ws_gv.column_dimensions["A"].width = 6
    ws_gv.column_dimensions["B"].width = 28
    ws_gv.freeze_panes = "A3"

    # ========================================
    # SAVE
    # ========================================
    try:
        wb.save(output_path)
        print(f"\n✅ Đã lưu file thống kê: {output_path}")
    except PermissionError:
        alt_path = output_path.with_stem(output_path.stem + " (new)")
        wb.save(alt_path)
        print(f"\n⚠ File gốc đang mở, đã lưu bản mới: {alt_path}")

    wb.close()


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("  THỐNG KÊ CUỘC THI HỌC SINH UNIGO")
    print("=" * 60)

    # Step 1: Read master student list
    print(f"\n📋 Đọc danh sách HS: {DS_HS_FILE.name}")
    master = read_master_students(DS_HS_FILE)
    print(f"   → {len(master['all_students'])} học sinh, {len(master['classes_order'])} lớp")
    for cls in master["classes_order"]:
        print(f"     • {cls}: {len(master['by_class'][cls])} HS")

    # Step 2: Scan contests
    print(f"\n🔍 Quét thư mục cuộc thi: {BASE_DIR}")
    contests = scan_contests(BASE_DIR)
    if not contests:
        print("   ⚠ Không tìm thấy cuộc thi nào!")
        return

    for contest_name, rounds in contests.items():
        print(f"   📁 {contest_name}:")
        for lan_num, files in sorted(rounds.items()):
            print(f"      • Lần {lan_num}: {len(files)} file(s)")

    # Step 3 & 4: Read results and match
    print(f"\n📊 Đọc kết quả và so khớp...")
    all_matches = {}  # {(contest_name, lan_num): {(ho_ten, lop): info_or_None}}

    for contest_name, rounds in contests.items():
        for lan_num, files in sorted(rounds.items()):
            print(f"   📄 {contest_name} - Lần {lan_num}:")
            results = read_contest_results(files)
            print(f"      → Đọc được {len(results)} bản ghi")

            matched = match_students(master["all_students"], results)
            participated = sum(1 for v in matched.values() if v is not None)
            not_participated = sum(1 for v in matched.values() if v is None)
            print(f"      ✅ Đã thi: {participated} | ❌ Chưa thi: {not_participated}")

            all_matches[(contest_name, lan_num)] = matched

    # Step 5: Generate output
    print(f"\n📝 Tạo file thống kê...")
    generate_output(master, contests, all_matches, OUTPUT_FILE)

    # Summary
    print(f"\n{'=' * 60}")
    print(f"  HOÀN TẤT!")
    print(f"{'=' * 60}")
    print(f"  📁 Output: {OUTPUT_FILE}")
    print(f"  📊 Sheets:")
    print(f"     • Tổng hợp — Bảng tổng ✅/❌ tất cả cuộc thi")
    for contest_name in sorted(contests.keys()):
        print(f"     • Chi tiết {contest_name} — Điểm, vòng, thời gian")
    print(f"     • Nhắc nhở GV — Danh sách HS chưa thi gom theo lớp")


if __name__ == "__main__":
    main()
