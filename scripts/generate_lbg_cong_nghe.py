import sys
import os
import re
from datetime import date, timedelta
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.stdout.reconfigure(encoding='utf-8')

from cong_nghe_data import DATA_CONG_NGHE

# Directory settings
LBG_DIR_OUT = r'D:\UNIGO\KHBD_Cong_nghe\Lịch báo giảng'
os.makedirs(LBG_DIR_OUT, exist_ok=True)

TEMPLATE_LBG = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Lịch báo giảng\template_lbg_goc.docx'
if not os.path.exists(TEMPLATE_LBG):
    TEMPLATE_LBG = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Lịch báo giảng\Lịch báo giảng - Tuần 01.docx'

TUAN_01_START = date(2026, 8, 3)
THU_LABELS = {0: 'Hai', 1: 'Ba', 2: 'Tư', 3: 'Năm', 4: 'Sáu'}

# Map lessons for Công nghệ
PPCT_CN = {}
for g in ['3', '4', '5', '6', '7', '8']:
    PPCT_CN[g] = {}
    PPCT_CN[g][0] = 'Tiết 0: Định hướng môn học'
    lessons = DATA_CONG_NGHE[g]['lessons']
    for stt, name, st, yccd, is_topic in lessons:
        if not is_topic and stt is not None:
            PPCT_CN[g][stt] = name

# Import helper functions from generate_lbg
import importlib.util
spec = importlib.util.spec_from_file_location("generate_lbg", r"D:\UNIGO\scripts\generate_lbg.py")
gen_lbg = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gen_lbg)
PPCT_TIN = gen_lbg.PPCT_TIN

# Schedule mapping: Replace Robotics with Cong nghe for Grade 3-8, and STEM for Grade 1-2
MASTER_SCHEDULE_SANG = {
    (0, 3): ('4C1', 'Công nghệ', False),
    (0, 4): ('1A1', 'Tin học', False),
    (1, 1): ('1C1', 'Tin học', False),
    (1, 3): ('5C1', 'Tin - Công nghệ', True),
    (1, 4): ('5C1', 'Tin - Công nghệ', True),
    (2, 2): ('3A1', 'Công nghệ', False),
    (2, 4): ('3C1', 'Tin học', False),
    (3, 1): ('TT3', 'Tin học', False),
    (3, 2): ('1A1', 'Trải nghiệm STEM', False),
    (3, 4): ('2C1', 'Trải nghiệm STEM', False),
    (4, 1): ('2A1', 'Tin học', False),
    (4, 2): ('3C1', 'Công nghệ', False),
    (4, 4): ('6A1', 'Tin - Công nghệ', True),
    (4, 5): ('6A1', 'Tin - Công nghệ', True),
}

MASTER_SCHEDULE_CHIEU = {
    (0, 3): ('2C1', 'Tin học', False),
    (1, 3): ('7A1', 'Tin - Công nghệ', True),
    (1, 4): ('7A1', 'Tin - Công nghệ', True),
    (2, 3): ('4C1', 'Tin học', False),
    (2, 4): ('2A1', 'Trải nghiệm STEM', False),
    (3, 1): ('3A1', 'Tin học', False),
    (3, 2): ('TTH 1', 'Tin học', False),
    (3, 3): ('TTH2', 'Tin học', False),
    (3, 4): ('1C1', 'Trải nghiệm STEM', False),
    (4, 1): ('8A1', 'Tin - Công nghệ', True),
    (4, 2): ('8A1', 'Tin - Công nghệ', True),
}

def week_dates(tuan_so):
    delta = timedelta(weeks=tuan_so - 1)
    start = TUAN_01_START + delta
    end = start + timedelta(days=4)
    return start, end

def day_label(d):
    thu = THU_LABELS[d.weekday()]
    return f'{thu} ({d.day:02d}/{d.month:02d})'

def compute_rotation_ppct(tuan_so, is_even_week, period_index_in_day):
    if tuan_so <= 1:
        return 0
    if is_even_week:
        k = (tuan_so - 2) // 2 + 1
        return (k - 1) * 2 + period_index_in_day + 1
    else:
        k = (tuan_so - 3) // 2 + 1
        return (k - 1) * 2 + period_index_in_day + 1

def populate_lbg_table(tbl, schedule_map, tuan_so, start_date, is_sang=True):
    is_even = (tuan_so % 2 == 0)
    for day_idx in range(5):
        d = start_date + timedelta(days=day_idx)
        d_lbl = day_label(d)
        for tiet in range(1, 6):
            ri = day_idx * 5 + tiet
            if ri >= len(tbl.rows):
                continue
            row = tbl.rows[ri]
            gen_lbg.set_cell_text(row.cells[0], d_lbl)
            gen_lbg.set_cell_text(row.cells[1], str(tiet))
            key = (day_idx, tiet)
            if key in schedule_map:
                lop, mon_display, is_rot = schedule_map[key]
                grade_char = lop[0] if lop[0].isdigit() else '3'
                
                if is_rot:
                    # Tuần chẵn -> Tin học; Tuần lẻ -> Công nghệ
                    mon_lbl = 'Tin - Công nghệ'
                    p_idx = 0 if (day_idx, tiet - 1) not in schedule_map else 1
                    if is_even:
                        dd = 'Phòng Tin học'
                        ppct = compute_rotation_ppct(tuan_so, True, p_idx)
                        lesson = PPCT_TIN.get(grade_char, {}).get(ppct, f'Bài học Tiết {ppct}') if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                    else:
                        dd = 'Phòng thực hành Công nghệ'
                        ppct = compute_rotation_ppct(tuan_so, False, p_idx)
                        lesson = PPCT_CN.get(grade_char, {}).get(ppct, f'Bài học Tiết {ppct}') if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                else:
                    mon_lbl = mon_display
                    ppct = (tuan_so - 1) if tuan_so > 1 else 0
                    if mon_lbl == 'Công nghệ':
                        dd = 'Phòng thực hành Công nghệ'
                        lesson = PPCT_CN.get(grade_char, {}).get(ppct, f'Bài học Tiết {ppct}') if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                    elif mon_lbl == 'Tin học':
                        dd = 'Phòng Tin học'
                        lesson = PPCT_TIN.get(grade_char, {}).get(ppct, f'Bài học Tiết {ppct}') if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                    else: # Trải nghiệm STEM
                        dd = 'Bộ đồ dùng STEM'
                        lesson = f'Chủ đề STEM trải nghiệm sáng tạo {ppct}' if tuan_so > 1 else 'Tiết 0: Định hướng môn học'

                ppct_str = str(ppct) if tuan_so > 1 else '1'
                if not lesson and ppct == 0:
                    lesson = 'Tiết 0: Định hướng môn học'
                gen_lbg.set_cell_text(row.cells[2], ppct_str)
                gen_lbg.set_cell_text(row.cells[3], mon_lbl)
                gen_lbg.set_cell_text(row.cells[4], lop)
                gen_lbg.set_cell_text(row.cells[5], lesson)
                gen_lbg.set_cell_text(row.cells[6], dd)
            else:
                gen_lbg.set_cell_text(row.cells[2], '')
                gen_lbg.set_cell_text(row.cells[3], '')
                gen_lbg.set_cell_text(row.cells[4], '')
                gen_lbg.set_cell_text(row.cells[5], '')
                gen_lbg.set_cell_text(row.cells[6], '')

    # Tuần 5: Block merge Nghỉ lễ 02/09 & Tổng duyệt khai giảng
    if tuan_so == 5:
        gen_lbg.merge_and_format_block(tbl, start_row=1, end_row=15, start_col=2, end_col=6, text='Nghỉ lễ 02/09')
        if is_sang:
            gen_lbg.merge_and_format_block(tbl, start_row=16, end_row=25, start_col=2, end_col=6, text='Tổng duyệt khai giảng')

def populate_lbg_week(tuan_so):
    start_date, end_date = week_dates(tuan_so)
    doc = docx.Document(TEMPLATE_LBG)
    gen_lbg.update_headers(doc, tuan_so, start_date, end_date)
    
    lbg_tbls = [t for t in doc.tables
                if len(t.columns) == 7
                and 'Lớp' in [c.text.strip() for c in t.rows[0].cells]]
                
    if len(lbg_tbls) >= 1:
        populate_lbg_table(lbg_tbls[0], MASTER_SCHEDULE_SANG, tuan_so, start_date, is_sang=True)
    if len(lbg_tbls) >= 2:
        populate_lbg_table(lbg_tbls[1], MASTER_SCHEDULE_CHIEU, tuan_so, start_date, is_sang=False)
        
    gen_lbg.update_ky_ten(doc, start_date)
    
    out_file = os.path.join(LBG_DIR_OUT, f'Lịch báo giảng - Tuần {tuan_so:02d} (Công nghệ).docx')
    gen_lbg.save_safe(doc, out_file)
    return out_file

print("Starting generation of 35 weeks of LBG for Công nghệ...")
for w in range(1, 36):
    f = populate_lbg_week(w)
    print(f"✓ Tuần {w:02d}: {os.path.basename(f)}")

print("\n>>> ĐÃ HOÀN TẤT TOÀN BỘ 35 TUẦN LỊCH BÁO GIẢNG MÔN CÔNG NGHỆ! <<<")
