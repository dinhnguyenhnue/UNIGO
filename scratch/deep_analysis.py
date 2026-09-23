# -*- coding: utf-8 -*-
"""
Script đọc sâu cấu trúc giáo án gốc để hiểu các section cần extract.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

base = r'D:\UNIGO\Phân phối chương trình\Robotics'

# Read FULL content of one file from each program
for label, fpath in [
    ('INITIATE', os.path.join(base, 'Giáo án-Initiate', 'I 03 - 01 - đu xà.docx')),
    ('KINDER', os.path.join(base, 'Giáo án-Kinder', 'K+3 - 01 - chiếc nồi.docx')),
    ('EXCEL', os.path.join(base, 'Giáo án-Excel', 'robohub-exce.01-01.docx')),
]:
    print(f'\n{"="*80}')
    print(f'=== {label}: {os.path.basename(fpath)} ===')
    print(f'{"="*80}')
    doc = Document(fpath)
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if txt:
            bold = any(r.bold for r in p.runs if r.bold)
            print(f'P{i:3d} [{"B" if bold else " "}]: {txt[:250]}')
    
    # Also print tables
    for ti, t in enumerate(doc.tables):
        print(f'\n  TABLE {ti}: {len(t.rows)}x{len(t.columns)}')
        for ri, row in enumerate(t.rows):
            cells = [c.text[:80].replace('\n', ' | ') for c in row.cells]
            print(f'    R{ri}: {cells}')
