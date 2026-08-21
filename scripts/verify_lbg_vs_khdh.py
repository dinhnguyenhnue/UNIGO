import sys, os, glob, docx
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'D:/UNIGO/scripts')

from generate_lbg import PPCT_TIN, PPCT_ROB

def parse_khdh(docx_path):
    doc = docx.Document(docx_path)
    lessons = {}
    for tbl in doc.tables:
        if len(tbl.columns) >= 4 or len(tbl.rows) > 10:
            for row in tbl.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0].isdigit():
                    lessons[int(cells[0])] = cells[1].replace('\n', ' ')
    return lessons

grade_map = {
    '1': 'Lớp 1', '2': 'Lớp 2', '3': 'Lớp 3', '4': 'Lớp 4',
    '5': 'Lớp 5', '6': 'Lớp 6', '7': 'Lớp 7', '8': 'Lớp 8',
    'TT': 'Tiền tiểu học'
}

print('=================== SO SÁNH TIN HỌC ===================')
khdh_tin = {}
for f in sorted(glob.glob('D:/UNIGO/Hệ thống mẫu văn bản/Nguyên đã làm/Kế hoạch dạy học Tin học từng lớp/*.docx')):
    grade = os.path.basename(f).split(' - ')[1]
    khdh_tin[grade] = parse_khdh(f)

for code, g_name in grade_map.items():
    lbg_data = PPCT_TIN.get(code, {})
    kh_data = khdh_tin.get(g_name, {})
    diffs = []
    max_ppct = max(max(lbg_data.keys(), default=0), max(kh_data.keys(), default=0))
    for p in range(1, max_ppct + 1):
        lbg_title = lbg_data.get(p, '<THIẾU TRONG LBG>').strip()
        kh_title = kh_data.get(p, '<THIẾU TRONG KHDH>').strip()
        if lbg_title.lower() != kh_title.lower():
            diffs.append((p, lbg_title, kh_title))
    
    if not diffs:
        print(f'[V] {g_name} ({code}): Khớp hoàn toàn {len(kh_data)} tiết.')
    else:
        print(f'[X] {g_name} ({code}): Khác nhau {len(diffs)} tiết:')
        for p, l, k in diffs[:5]:
            print(f'   - Tiết {p}:')
            print(f'       LBG : \"{l}\"')
            print(f'       KHDH: \"{k}\"')
        if len(diffs) > 5:
            print(f'       ... và {len(diffs)-5} tiết khác')

print('\n================== SO SÁNH ROBOTICS ==================')
khdh_rob = {}
for f in sorted(glob.glob('D:/UNIGO/Hệ thống mẫu văn bản/Nguyên đã làm/Kế hoạch dạy học Robotics từng lớp/*.docx')):
    grade = os.path.basename(f).split(' - ')[1]
    khdh_rob[grade] = parse_khdh(f)

for code, g_name in grade_map.items():
    lbg_data = PPCT_ROB.get(code, {})
    kh_data = khdh_rob.get(g_name, {})
    diffs = []
    max_ppct = max(max(lbg_data.keys(), default=0), max(kh_data.keys(), default=0))
    for p in range(1, max_ppct + 1):
        lbg_title = lbg_data.get(p, '<THIẾU TRONG LBG>').strip()
        kh_title = kh_data.get(p, '<THIẾU TRONG KHDH>').strip()
        if lbg_title.lower() != kh_title.lower():
            diffs.append((p, lbg_title, kh_title))
    
    if not diffs:
        print(f'[V] {g_name} ({code}): Khớp hoàn toàn {len(kh_data)} tiết.')
    else:
        print(f'[X] {g_name} ({code}): Khác nhau {len(diffs)} tiết:')
        for p, l, k in diffs[:5]:
            print(f'   - Tiết {p}:')
            print(f'       LBG : \"{l}\"')
            print(f'       KHDH: \"{k}\"')
        if len(diffs) > 5:
            print(f'       ... và {len(diffs)-5} tiết khác')
