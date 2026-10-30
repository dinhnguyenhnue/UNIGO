import sys, openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook(r'D:\UNIGO\TKB toàn trường CHECK - 26.10.xlsx', data_only=True)
ws = wb['TKB_LOP_SC']

# Determine class name for each column
col_class = {}
col_session = {}
current_class = None
for c in range(1, ws.max_column + 1):
    val4 = ws.cell(4, c).value
    val5 = ws.cell(5, c).value
    if val4 is not None and str(val4).strip() not in ['THỨ', 'TIẾT']:
        current_class = str(val4).strip()
    elif val4 in ['THỨ', 'TIẾT']:
        current_class = None
    col_class[c] = current_class
    col_session[c] = str(val5).strip() if val5 else ''

# Determine Day and Period for each row
row_day = {}
row_period = {}
cur_day = None
for r in range(6, 31):
    val_a = ws.cell(r, 1).value
    if val_a is not None and str(val_a).strip().isdigit():
        cur_day = int(str(val_a).strip())
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
            slots.append({
                'row': r,
                'col': c,
                'col_letter': openpyxl.utils.get_column_letter(c),
                'day': f'Thứ {day}',
                'day_num': day,
                'period': period,
                'session': session,
                'class': cls,
                'val': str(v).strip()
            })

print(f'Total teaching slots found: {len(slots)}')
for s in sorted(slots, key=lambda x: (x['day_num'], 0 if x['session']=='Sáng' else 1, x['period'])):
    print(f"{s['day']} | {s['session']:<5} | Tiết {s['period']} | Lớp: {s['class']:<8} | Môn: {s['val']:<20} | Cell: {s['col_letter']}{s['row']}")

# Check for any overlapping periods (same day, same session, same period)
from collections import defaultdict
time_slots = defaultdict(list)
for s in slots:
    key = (s['day_num'], s['session'], s['period'])
    time_slots[key].append(s)

conflicts = {k: v for k, v in time_slots.items() if len(v) > 1}
if conflicts:
    print('\n!!! CÁC TIẾT TRÙNG LỊCH (CONFLICTS) !!!')
    for k, v in conflicts.items():
        print(f"Thứ {k[0]} - {k[1]} - Tiết {k[2]}:")
        for item in v:
            print(f"   -> Lớp {item['class']} ({item['val']}) [Cell {item['col_letter']}{item['row']}]")
else:
    print('\n>>> KHÔNG CÓ TIẾT TRÙNG LỊCH NÀO!')
