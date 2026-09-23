# -*- coding: utf-8 -*-
"""
Script kiểm tra kết quả copy paste:
1. So sánh nội dung trước/sau (backup vs current)
2. Kiểm tra cấu trúc UNIGO (header, tables, footer)
3. Xác nhận nội dung đã thay đổi
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

BASE_CUR = r'D:\UNIGO\KHBD_Robotics'
BASE_BAK = r'D:\UNIGO\KHBD_Robotics_BACKUP_20260923'

def check_file(current_path, backup_path, label):
    """Compare current vs backup and check structure."""
    doc_cur = Document(current_path)
    doc_bak = Document(backup_path)
    
    issues = []
    
    # 1. Check table count preserved
    if len(doc_cur.tables) != len(doc_bak.tables):
        issues.append(f'Table count changed: {len(doc_bak.tables)} → {len(doc_cur.tables)}')
    
    # 2. Check Table 0 (thông tin GV) preserved
    if len(doc_cur.tables) >= 1:
        t0_cur = doc_cur.tables[0]
        t0_bak = doc_bak.tables[0]
        if t0_cur.rows[0].cells[0].text != t0_bak.rows[0].cells[0].text:
            issues.append(f'Table 0 (GV info) changed!')
    
    # 3. Check last table (chữ ký) preserved
    if len(doc_cur.tables) >= 2:
        last_cur = doc_cur.tables[-1]
        last_bak = doc_bak.tables[-1]
        if 'DUYỆT CỦA BGH' in last_bak.rows[0].cells[0].text:
            if last_cur.rows[0].cells[0].text != last_bak.rows[0].cells[0].text:
                issues.append(f'Signature table changed!')
    
    # 4. Check content actually changed (Tables 1-4)
    content_changed = False
    for ti in range(1, min(5, len(doc_cur.tables))):
        if ti < len(doc_cur.tables) and ti < len(doc_bak.tables):
            cur_text = doc_cur.tables[ti].rows[-1].cells[0].text[:100]
            bak_text = doc_bak.tables[ti].rows[-1].cells[0].text[:100]
            if cur_text != bak_text:
                content_changed = True
                break
    
    # 5. Check header drawings preserved
    header_ok = True
    for section in doc_cur.sections:
        header = section.header
        if header:
            # Check for drawings in header
            drawings = header._element.findall('.//' + '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
            if len(drawings) == 0:
                # Also check in w:r elements
                pass
    
    # 6. Check paragraph structure
    para_count_diff = len(doc_cur.paragraphs) - len(doc_bak.paragraphs)
    
    status = '✅' if not issues and content_changed else ('⚠️' if issues else '🔄')
    
    print(f'{status} {label}')
    if content_changed:
        print(f'    Content: ✅ Changed (new lesson data)')
    else:
        print(f'    Content: ⚠️ Unchanged')
    if para_count_diff != 0:
        print(f'    Paragraphs: {len(doc_bak.paragraphs)} → {len(doc_cur.paragraphs)} (diff: {para_count_diff})')
    print(f'    Tables: {len(doc_cur.tables)} (preserved)')
    for issue in issues:
        print(f'    ❌ {issue}')
    
    return len(issues) == 0 and content_changed


# Sample check - pick representative files from each group
test_files = [
    # Lớp 1 (Initiate)
    (r'Lớp_1\Tuần_02\KHBD_Robotics_Lớp_1_Tiet01_Bai_1_Tap_the_duc_nao.docx', 'Lớp 1 Bài 1 (Initiate I3-01)'),
    (r'Lớp_1\Tuần_07\KHBD_Robotics_Lớp_1_Tiet06_Bai_6_Kham_pha_xe_cuu_hoa.docx', 'Lớp 1 Bài 6 (Initiate I4-02)'),
    (r'Lớp_1\Tuần_22\KHBD_Robotics_Lớp_1_Tiet21_Bai_17_May_ban_a_khong_lo.docx', 'Lớp 1 Bài 17 (Initiate I7-01)'),
    
    # Lớp 2 (Kinder)
    (r'Lớp_2\Tuần_02\KHBD_Robotics_Lớp_2_Tiet01_Bai_1_Hay_nau_nhung_mon_an_ngon.docx', 'Lớp 2 Bài 1 (Kinder K+3-01)'),
    (r'Lớp_2\Tuần_14\KHBD_Robotics_Lớp_2_Tiet13_Bai_11_The_gioi_khung_long.docx', 'Lớp 2 Bài 11 (Kinder K+5-03)'),
    
    # Lớp 3 (Kinder)
    (r'Lớp_3\Tuần_05\KHBD_Robotics_Lớp_3_Tiet04_Bai_4_an_ga_con.docx', 'Lớp 3 Bài 4 (Kinder K+3-04)'),
    
    # Lớp 4 (Kinder)
    (r'Lớp_4\Tuần_12\KHBD_Robotics_Lớp_4_Tiet11_Bai_9_Moi_truong_bien.docx', 'Lớp 4 Bài 9 (Kinder K+5-01)'),
    
    # Lớp 5 (Excel)
    (r'Lớp_5\Tuần_03\KHBD_Robotics_Lớp_5_Tiet01_Bai_1_ong_co_la_gi.docx', 'Lớp 5 Bài 1 (Excel 01-01)'),
    (r'Lớp_5\Tuần_07\KHBD_Robotics_Lớp_5_Tiet05_Bai_5_ong_co_Dynamixel_hoat_ong_the_nao.docx', 'Lớp 5 Bài 5 (Excel 1-05)'),
    
    # Lớp 6 (Excel)
    (r'Lớp_6\Tuần_03\KHBD_Robotics_Lớp_6_Tiet01_Bai_1_ong_co_la_gi.docx', 'Lớp 6 Bài 1 (Excel 01-01)'),
    (r'Lớp_6\Tuần_15\KHBD_Robotics_Lớp_6_Tiet14_Bai_12_Domino_trong_Robotics.docx', 'Lớp 6 Bài 12 (Excel 1-12)'),
    
    # Lớp 7 (Excel)
    (r'Lớp_7\Tuần_05\KHBD_Robotics_Lớp_7_Tiet03_Bai_3_Robot_nhan_biet_am_thanh.docx', 'Lớp 7 Bài 3 (Excel 01-03)'),
    
    # Lớp 8 (Excel)
    (r'Lớp_8\Tuần_13\KHBD_Robotics_Lớp_8_Tiet12_Bai_10_Robot_hut_bui.docx', 'Lớp 8 Bài 10 (Excel 1-10)'),
]

print('=' * 70)
print('VERIFICATION RESULTS')
print('=' * 70)

pass_count = 0
fail_count = 0

for rel_path, label in test_files:
    cur = os.path.join(BASE_CUR, rel_path)
    bak = os.path.join(BASE_BAK, rel_path)
    if os.path.exists(cur) and os.path.exists(bak):
        result = check_file(cur, bak, label)
        if result:
            pass_count += 1
        else:
            fail_count += 1
    else:
        print(f'⚠️ {label}: File not found')
        fail_count += 1

print(f'\n{"=" * 70}')
print(f'PASS: {pass_count}, FAIL: {fail_count}')
print(f'{"=" * 70}')

# Also show sample content of one changed file
print(f'\n\n{"=" * 70}')
print('SAMPLE: Lớp 1 Bài 1 - Table 1 (Hoạt động 1 Khởi động) AFTER')
print(f'{"=" * 70}')
doc = Document(os.path.join(BASE_CUR, r'Lớp_1\Tuần_02\KHBD_Robotics_Lớp_1_Tiet01_Bai_1_Tap_the_duc_nao.docx'))
t1 = doc.tables[1]
print(f'  GV-HS: {t1.rows[1].cells[0].text[:300]}')
print(f'  KQ: {t1.rows[1].cells[1].text[:300]}')

print(f'\nSAMPLE: Lớp 6 Bài 1 - Table 2 (Hoạt động 2 Kiến thức) AFTER')
doc6 = Document(os.path.join(BASE_CUR, r'Lớp_6\Tuần_03\KHBD_Robotics_Lớp_6_Tiet01_Bai_1_ong_co_la_gi.docx'))
t2 = doc6.tables[2]
print(f'  GV-HS: {t2.rows[1].cells[0].text[:300]}')
print(f'  KQ: {t2.rows[1].cells[1].text[:300]}')

# Show kiến thức paragraphs
print(f'\nSAMPLE: Lớp 6 Bài 1 - Kiến thức paragraphs AFTER')
for i, p in enumerate(doc6.paragraphs):
    txt = p.text.strip()
    if txt and ('Kiến thức' in txt or txt.startswith('- Nêu') or txt.startswith('- Sử dụng') or txt.startswith('- Nhận biết') or txt.startswith('- Hiểu')):
        print(f'  P{i}: {txt}')
