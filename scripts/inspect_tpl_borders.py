import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

tpl_path = r'd:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx'
doc = Document(tpl_path)
print("=== TEMPLATE THCS TABLE 0 ===")
tbl = doc.tables[0]
print("Table Style:", tbl.style.name if tbl.style else "None")
tblPr = tbl._tbl.find(qn('w:tblPr'))
tblBorders = tblPr.find(qn('w:tblBorders')) if tblPr is not None else None
if tblBorders is not None:
    for c in tblBorders:
        print("  Border:", c.tag.split('}')[-1], c.attrib)
else:
    print("  No tblBorders in tblPr (style default)")

# Inspect cells
for r in tbl._tbl.findall(qn('w:tr')):
    for c in r.findall(qn('w:tc')):
        tcPr = c.find(qn('w:tcPr'))
        if tcPr is not None:
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is not None:
                print("  Cell border:", [b.attrib for b in tcBorders])
