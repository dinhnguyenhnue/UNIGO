import os
import glob
import sys
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

base = r'D:\UNIGO\KHBD_Robotics'
for grade in [1, 2, 3, 4, 5, 6, 7, 8]:
    g_dir = os.path.join(base, f'Lớp_{grade}')
    files = glob.glob(os.path.join(g_dir, '**', '*.docx'), recursive=True)
    weeks = sorted(os.listdir(g_dir))
    print(f'=== LỚP {grade}: {len(files)} files, {len(weeks)} week folders ===')
    assessments = [f for f in files if 'anh_gia_inh_ky' in f or 'Tong_ket' in f]
    for af in sorted(assessments):
        rel = os.path.relpath(af, g_dir)
        print(f'   -> {rel}')

sample = os.path.join(base, 'Lớp_6', 'Tuần_09', 'KHBD_Robotics_Lớp_6_Tiet08_anh_gia_inh_ky_1.docx')
doc = Document(sample)
print(f'\nKiểm tra viền bảng file mẫu: {os.path.basename(sample)}')
ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
for i, t in enumerate(doc.tables):
    tblPr = t._tbl.tblPr
    borders = tblPr.find(f'{{{ns}}}tblBorders')
    border_info = []
    if borders is not None:
        for c in borders:
            tag = c.tag.split('}')[-1]
            val = c.get(f'{{{ns}}}val')
            border_info.append(f'{tag}={val}')
    print(f'  Bảng {i} ({len(t.rows)}x{len(t.columns)}): {" ".join(border_info) if border_info else "Không có tblBorders"}')
