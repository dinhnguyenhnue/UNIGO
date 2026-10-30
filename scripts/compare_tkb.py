import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 1. Load Old TKB
wb_old = openpyxl.load_workbook(r'D:\UNIGO\Thời khóa biểu giáo viên\Thời khóa biểu - Đậu Đình Nguyên.xlsx', data_only=True)
ws_old = wb_old.active

days = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu']
old_slots = {}
for r in range(7, 18):
    if r == 12: continue # header
    sess = 'Sáng' if r <= 11 else 'Chiều'
    tiet_num = (r - 6) if r <= 11 else (r - 12)
    for c_idx, day in enumerate(days, 4):
        v = ws_old.cell(r, c_idx).value
        if v:
            clean_v = str(v).replace('\n', ' ').strip()
            # extract subject and class
            old_slots[(day, sess, tiet_num)] = clean_v

# 2. Load New TKB from TKB_LOP_SC
wb_new = openpyxl.load_workbook(r'D:\UNIGO\TKB toàn trường CHECK - 24.10.xlsx', data_only=True)
ws_sc = wb_new['TKB_LOP_SC']

blocks = [
    (1, 2, range(3, 11)),
    (11, 12, range(13, 21)),
    (21, 22, range(23, 33)),
    (33, 34, range(35, 47)),
]

col_to_class = {}
for b in blocks:
    curr_c = ''
    for c in b[2]:
        val4 = ws_sc.cell(4, c).value
        if val4: curr_c = str(val4).strip()
        sess = str(ws_sc.cell(5, c).value or '').strip()
        col_to_class[c] = (curr_c, sess, b[0], b[1])

day_map = {2: 'Thứ Hai', 3: 'Thứ Ba', 4: 'Thứ Tư', 5: 'Thứ Năm', 6: 'Thứ Sáu',
           '2': 'Thứ Hai', '3': 'Thứ Ba', '4': 'Thứ Tư', '5': 'Thứ Năm', '6': 'Thứ Sáu'}

new_slots = {}
for r in range(6, 31):
    for c, (cname, sess, thu_c, tiet_c) in col_to_class.items():
        val = str(ws_sc.cell(r, c).value or '').strip()
        if 'nguyên' in val.lower() or 'nguyen' in val.lower():
            thu_val = None
            for r_back in range(r, 5, -1):
                v_thu = ws_sc.cell(r_back, thu_c).value
                if v_thu is not None and str(v_thu).strip() != '':
                    thu_val = v_thu
                    break
            tiet_val = ws_sc.cell(r, tiet_c).value
            day_name = day_map.get(thu_val, f'Thứ {thu_val}')
            key = (day_name, sess, tiet_val)
            if key not in new_slots:
                new_slots[key] = []
            new_slots[key].append(f"{val} ({cname})")

print("=== SO SÁNH TKB CŨ VÀ TKB MỚI (TKB_LOP_SC) ===")
print(f"Tổng tiết TKB Cũ: {len(old_slots)} | Tổng tiết TKB Mới: {sum(len(v) for v in new_slots.values())}\n")

all_keys = sorted(set(list(old_slots.keys()) + list(new_slots.keys())), 
                  key=lambda x: (days.index(x[0]) if x[0] in days else 99, 0 if x[1]=='Sáng' else 1, x[2]))

print(f"{'Thời gian':<25} | {'TKB Cũ':<30} | {'TKB Mới (24.10)':<40} | Trạng thái")
print("-" * 115)

for k in all_keys:
    time_str = f"{k[0]} {k[1]} T{k[2]}"
    old_val = old_slots.get(k, "---")
    new_val_list = new_slots.get(k, ["---"])
    new_val = " + ".join(new_val_list)
    
    if len(new_val_list) > 1:
        status = "⚠️ TRÙNG LỊCH MỚI"
    elif old_val == "---":
        status = "➕ Tiết mới"
    elif new_val == "---":
        status = "➖ Bị bỏ"
    elif old_val == new_val or (cname in old_val for cname in new_val_list):
        status = "🔄 Giống / Khác format"
    else:
        status = "🔀 Đổi tiết"
    print(f"{time_str:<25} | {old_val:<30} | {new_val:<40} | {status}")
