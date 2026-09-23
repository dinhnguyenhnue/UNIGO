# -*- coding: utf-8 -*-
"""
Đọc TOÀN BỘ nội dung KHBD Lớp 1 Tuần 02 để hiểu chính xác cấu trúc cần thay thế.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

# Full KHBD Lớp 1
doc = Document(r'D:\UNIGO\KHBD_Robotics\Lớp_1\Tuần_02\KHBD_Robotics_Lớp_1_Tiet01_Bai_1_Tap_the_duc_nao.docx')
print(f'Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}')

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    bold = any(r.bold for r in p.runs if r.bold) if p.runs else False
    if txt:
        print(f'P{i:3d} [{"B" if bold else " "}]: {txt[:250]}')

for ti, t in enumerate(doc.tables):
    print(f'\n  TABLE {ti}: {len(t.rows)}x{len(t.columns)}')
    for ri, row in enumerate(t.rows):
        cells = []
        for c in row.cells:
            ct = c.text[:120].replace('\n', ' | ')
            cells.append(ct)
        print(f'    R{ri}: {cells}')
