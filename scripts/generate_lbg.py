"""
generate_lbg.py — Tu dong tao Lich bao giang theo tuan
======================================================

QUY TAC ROTATION:
  - Lop 5, 6, 7, 8: Moi lop co 2 tiet / tuan
      Tuan CHAN (2, 4, 6...): ca 2 tiet -> Tin hoc
      Tuan LE  (3, 5, 7...): ca 2 tiet -> Robotics
  - Cac lop con lai: giu nguyen mon hoc goc

QUY TRINH (moi tuan):
  1. Doc file Tuan 01 lam template
  2. Cap nhat tieu de tuan + ngay
  3. Cap nhat nhan Thu/ngay trong bang
  4. Ap dung rotation cho lop 5,6,7,8
  5. Cap nhat PPCT (tuan 2 = tiet 1, tuan 3 = tiet 2, ...)
  6. Luu ban goc tong hop
  7. Tach 2 ban: TTH+TH va THCS
  8. Chen page break (sang/chieu/nhan xet = 3 trang)

CACH DUNG:
  python generate_lbg.py <so_tuan>
  Vi du: python generate_lbg.py 2
         python generate_lbg.py 3

OUTPUT (trong thu muc LBG_DIR):
  Lich bao giang - Tuan XX.docx          <- ban goc tong hop
  Lich bao giang - Tuan XX (TTH+TH).docx <- Tien TH + TH
  Lich bao giang - Tuan XX (THCS).docx   <- THCS
"""

import sys
import os
import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import date, timedelta

sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────────────────────────────────────
# CAU HINH
# ─────────────────────────────────────────────────────────────────────────────

LBG_DIR = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Lịch báo giảng'
TEMPLATE = os.path.join(LBG_DIR, 'Lịch báo giảng - Tuần 01.docx')

# Ngay dau Tuan 01 (Thu Hai 03/08/2026)
TUAN_01_START = date(2026, 8, 3)

# Lop co rotation Tin hoc / Robotics tuan chan-le (bat dau tu lop 5)
# Phat hien dong: neu lop bat dau bang 5, 6, 7, 8 thi rotation
ROTATION_PREFIXES = ('5', '6', '7', '8')

THU_LABELS = {0: 'Hai', 1: 'Ba', 2: 'Tư', 3: 'Năm', 4: 'Sáu'}

# ─────────────────────────────────────────────────────────────────────────────
# TIEN ICH
# ─────────────────────────────────────────────────────────────────────────────

def week_dates(tuan_so):
    delta = timedelta(weeks=tuan_so - 1)
    start = TUAN_01_START + delta
    end = start + timedelta(days=4)
    return start, end


def fmt_date(d):
    return f'{d.day:02d}/{d.month:02d}/{d.year}'


def day_label(d):
    thu = THU_LABELS[d.weekday()]
    return f'{thu} ({d.day:02d}/{d.month:02d})'


def classify_lop(lop):
    """'TTH_TH' hoac 'THCS' hoac None."""
    s = lop.strip().upper()
    if not s:
        return None
    if re.match(r'^[6789]', s):
        return 'THCS'
    return 'TTH_TH'


def is_rotation_lop(lop):
    """Tra ve True neu lop thuoc nhom 5,6,7,8 (co rotation)."""
    s = lop.strip().upper()
    return any(s.startswith(p) for p in ROTATION_PREFIXES)


def rotation_mon(tuan_so):
    """Mon hoc cho lop rotation theo tuan: chan=Tin hoc, le=Robotics."""
    return 'Tin học' if tuan_so % 2 == 0 else 'Robotics'


def rotation_do_dung(tuan_so):
    return 'Phòng Tin học' if tuan_so % 2 == 0 else 'Bộ Kit Robotics'


def set_cell_text(cell, text):
    """Ghi text vao cell, giu dinh dang run dau tien."""
    for p in cell.paragraphs:
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ''
            break
        else:
            p.text = text
            break


def save_safe(doc, path):
    """Luu file; neu bi khoa thi luu sang _v2."""
    try:
        doc.save(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        alt = base + '_v2' + ext
        doc.save(alt)
        print(f'   ⚠ Bi khoa, luu: {os.path.basename(alt)}')
        return alt


# ─────────────────────────────────────────────────────────────────────────────
# CAP NHAT NOI DUNG
# ─────────────────────────────────────────────────────────────────────────────

def update_headers(doc, tuan_so, start_date, end_date):
    """Cap nhat tieu de tuan va ngay trong paragraph."""
    tuan_str = f'TUẦN {tuan_so:02d}'
    fs = fmt_date(start_date)
    fe = fmt_date(end_date)
    ngay_soan = start_date - timedelta(days=3)  # Thu Sau tuan truoc

    for p in doc.paragraphs:
        txt = p.text.strip()

        if 'TUẦN' in txt.upper() and 'từ ngày' in txt:
            new = f'{tuan_str} (từ ngày {fs} đến ngày {fe})'
        elif re.match(r'^Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+', txt) and len(txt) < 45:
            new = (f'Ngày {ngay_soan.day:02d} tháng '
                   f'{ngay_soan.month:02d} năm {ngay_soan.year}')
        elif 'sáng' in txt.lower() and 'tuần' in txt.lower():
            new = (f'Buổi…sáng……..Tuần…{tuan_so:02d}…'
                   f'(Từ ngày…{fs} …đến ngày:… {fe}….)')
        elif 'chiều' in txt.lower() and 'tuần' in txt.lower():
            new = (f'Buổi…chiều…..Tuần…{tuan_so:02d}…'
                   f'(Từ ngày…{fs}…đến ngày:… {fe}….)')
        else:
            continue

        if p.runs:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ''
        else:
            p.text = new


def update_day_labels(doc, start_date):
    """Cap nhat nhan Thu/ngay trong Table[1] va Table[3]."""
    for ti in [1, 3]:
        tbl = doc.tables[ti]
        for ri in range(1, len(tbl.rows)):
            day_idx = (ri - 1) // 5   # 0=Hai..4=Sau
            d = start_date + timedelta(days=day_idx)
            label = day_label(d)
            cell0 = tbl.rows[ri].cells[0]
            if cell0.text.strip():
                set_cell_text(cell0, label)


def update_table_data(doc, tuan_so):
    """
    Cap nhat mon hoc, PPCT, do dung cho cac hang co du lieu.
    - Lop 5,6,7,8: ap dung rotation
    - Tat ca: cap nhat PPCT = tuan_so - 1 (tu tuan 2)
    """
    if tuan_so == 1:
        return

    ppct = str(tuan_so - 1)

    for ti in [1, 3]:
        tbl = doc.tables[ti]
        for ri in range(1, len(tbl.rows)):
            row = tbl.rows[ri]
            lop = row.cells[4].text.strip()
            if not lop:
                continue

            # Cap nhat PPCT
            set_cell_text(row.cells[2], ppct)

            # Rotation cho lop 5,6,7,8
            if is_rotation_lop(lop):
                mon = rotation_mon(tuan_so)
                dd = rotation_do_dung(tuan_so)
                set_cell_text(row.cells[3], mon)
                set_cell_text(row.cells[6], dd)


def update_ky_ten(doc, start_date):
    """Cap nhat ngay trong bang ky ten."""
    ngay_soan = start_date - timedelta(days=3)
    ngay_str = (f'Ngày {ngay_soan.day} tháng '
                f'{ngay_soan.month} năm {ngay_soan.year}')
    for ti in [2, 4]:
        if ti >= len(doc.tables):
            continue
        for row in doc.tables[ti].rows:
            for cell in row.cells:
                txt = cell.text
                if 'Ngày' in txt and 'tháng' in txt and 'năm' in txt:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if 'Ngày' in r.text and 'tháng' in r.text:
                                # Thay the doan "Ngay ... nam XXXX"
                                m = re.search(
                                    r'Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
                                    r.text)
                                if m:
                                    r.text = r.text[:m.start()] + ngay_str + r.text[m.end():]


# ─────────────────────────────────────────────────────────────────────────────
# PAGE BREAKS
# ─────────────────────────────────────────────────────────────────────────────

def add_page_break_after_table(tbl_element):
    """Them paragraph co page break ngay sau table."""
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    tbl_element.addnext(p)


def fix_page_breaks(doc):
    """
    Xoa empty paragraph thua giua cac section.
    Chen page break sau bang ky ten sang va chieu.
    Ket qua: 3 trang (sang | chieu | nhan xet BGH).
    """
    body = doc.element.body
    children = list(body)

    tbls = [c for c in children
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'tbl']

    sign_sang = tbls[2]   # Bang ky ten sang
    sign_chieu = tbls[4]  # Bang ky ten chieu

    def find_para(keyword):
        for c in list(body):
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'p':
                txt = ''.join(t.text or '' for t in c.findall('.//' + qn('w:t')))
                if keyword.lower() in txt.lower():
                    return c
        return None

    chieu_el = find_para('chiều')
    nhanxet_el = find_para('nhận xét của bgh')

    def remove_empty_between(el_a, el_b):
        ch = list(body)
        if el_a not in ch or el_b not in ch:
            return 0
        ia, ib = ch.index(el_a), ch.index(el_b)
        removed = 0
        for c in ch[ia + 1:ib]:
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'p':
                if not ''.join(t.text or '' for t in c.findall('.//' + qn('w:t'))).strip():
                    body.remove(c)
                    removed += 1
        return removed

    if chieu_el is not None:
        n = remove_empty_between(sign_sang, chieu_el)
        if n:
            print(f'   Xoa {n} para trong (sau ky ten sang)')
    if nhanxet_el is not None:
        n = remove_empty_between(sign_chieu, nhanxet_el)
        if n:
            print(f'   Xoa {n} para trong (sau ky ten chieu)')

    add_page_break_after_table(sign_sang)
    add_page_break_after_table(sign_chieu)
    print('   Page breaks: sang | chieu | nhan xet BGH')


# ─────────────────────────────────────────────────────────────────────────────
# TACH TTH+TH / THCS
# ─────────────────────────────────────────────────────────────────────────────

def clear_row_data(row):
    for ci in range(2, min(7, len(row.cells))):
        for p in row.cells[ci].paragraphs:
            for r in p.runs:
                r.text = ''
            if not p.runs and p.text.strip():
                p.text = ''


def filter_for_cap(doc, keep_cap):
    for ti in [1, 3]:
        for ri in range(1, len(doc.tables[ti].rows)):
            row = doc.tables[ti].rows[ri]
            lop = row.cells[4].text.strip()
            cap = classify_lop(lop)
            if cap is not None and cap != keep_cap:
                clear_row_data(row)


# ─────────────────────────────────────────────────────────────────────────────
# HAM CHINH
# ─────────────────────────────────────────────────────────────────────────────

def generate_lbg(tuan_so):
    start_date, end_date = week_dates(tuan_so)
    loai = 'CHẴN → Tin học' if tuan_so % 2 == 0 else 'LẺ → Robotics'

    print(f'\n📅 Tuan {tuan_so:02d}: {fmt_date(start_date)} -> {fmt_date(end_date)}')
    print(f'   Lop 5,6,7,8: {loai}')

    # B1: Load template
    doc = Document(TEMPLATE)

    # B2: Cap nhat tieu de
    update_headers(doc, tuan_so, start_date, end_date)
    print('   Cap nhat tieu de + ngay')

    # B3: Cap nhat nhan Thu/ngay
    update_day_labels(doc, start_date)
    print('   Cap nhat nhan Thu/ngay')

    # B4: Rotation + PPCT
    update_table_data(doc, tuan_so)
    mon_info = f'Tin hoc' if tuan_so % 2 == 0 else 'Robotics'
    print(f'   Rotation lop 5,6,7,8 -> {mon_info} | PPCT = {tuan_so - 1}')

    # B5: Ky ten
    update_ky_ten(doc, start_date)

    # B6: Luu ban goc
    main_name = f'Lịch báo giảng - Tuần {tuan_so:02d}.docx'
    main_path = os.path.join(LBG_DIR, main_name)
    saved_main = save_safe(doc, main_path)
    print(f'   Ban goc: {os.path.basename(saved_main)}')

    # B7+B8: Tach TTH+TH va THCS
    for keep_cap, suffix, label in [
        ('TTH_TH', 'TTH+TH', 'Tien TH + TH'),
        ('THCS', 'THCS', 'THCS'),
    ]:
        d2 = Document(saved_main)
        filter_for_cap(d2, keep_cap)
        fix_page_breaks(d2)

        split_name = f'Lịch báo giảng - Tuần {tuan_so:02d} ({suffix}).docx'
        split_path = os.path.join(LBG_DIR, split_name)
        saved = save_safe(d2, split_path)

        # Kiem tra
        dc = Document(saved)
        cs = sum(1 for ri in range(1, len(dc.tables[1].rows))
                 if dc.tables[1].rows[ri].cells[4].text.strip())
        cc = sum(1 for ri in range(1, len(dc.tables[3].rows))
                 if dc.tables[3].rows[ri].cells[4].text.strip())
        print(f'   {label}: sang={cs} tiet, chieu={cc} tiet -> {os.path.basename(saved)}')

    print(f'\n HOAN TAT Tuan {tuan_so:02d}! Thu muc: {LBG_DIR}')


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Cach dung: python generate_lbg.py <so_tuan>')
        print('Vi du:     python generate_lbg.py 2')
        sys.exit(1)
    try:
        tuan = int(sys.argv[1])
    except ValueError:
        print(f'Loi: "{sys.argv[1]}" khong phai so tuan hop le.')
        sys.exit(1)
    if not (1 <= tuan <= 52):
        print('Loi: So tuan phai tu 1 den 52.')
        sys.exit(1)
    generate_lbg(tuan)
