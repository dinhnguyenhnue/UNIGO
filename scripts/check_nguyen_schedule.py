import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r'D:\UNIGO\TKB toàn trường CHECK - 24.10.xlsx', data_only=True)
ws = wb['TKB_LOP_SC']

col_to_class = {}
blocks = [
    (1, 2, range(3, 11)),
    (11, 12, range(13, 21)),
    (21, 22, range(23, 33)),
    (33, 34, range(35, 47)),
]

for b in blocks:
    curr_c = ''
    for c in b[2]:
        val4 = ws.cell(4, c).value
        if val4:
            curr_c = str(val4).strip()
        sess = str(ws.cell(5, c).value or '').strip()
        col_to_class[c] = (curr_c, sess, b[0], b[1])

nguyen_slots = []
for r in range(6, 31):
    for c, (cname, sess, thu_c, tiet_c) in col_to_class.items():
        val = str(ws.cell(r, c).value or '').strip()
        if 'nguyên' in val.lower() or 'nguyen' in val.lower():
            thu_val = None
            for r_back in range(r, 5, -1):
                v_thu = ws.cell(r_back, thu_c).value
                if v_thu is not None and str(v_thu).strip() != '':
                    thu_val = v_thu
                    break
            tiet_val = ws.cell(r, tiet_c).value
            nguyen_slots.append({
                'r': r, 'c': c, 'thu': thu_val, 'session': sess, 'tiet': tiet_val,
                'class': cname, 'subject': val
            })

print(f'Total slots found: {len(nguyen_slots)}')
slot_counts = {}
for s in nguyen_slots:
    key = (s['thu'], s['session'], s['tiet'])
    slot_counts.setdefault(key, []).append(s)

for k, items in sorted(slot_counts.items(), key=lambda x: (str(x[0][0]), str(x[0][1]), str(x[0][2]))):
    status = 'OK' if len(items) == 1 else '!!! CONFLICT (TRÙNG) !!!'
    desc = ' | '.join([f"{it['class']} ({it['subject']}) [r{it['r']}c{it['c']}]" for it in items])
    print(f"Thứ {k[0]} {k[1]} Tiết {k[2]}: {status} -> {desc}")
