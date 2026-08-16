import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

p = r'd:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_03\KHBD_Tin_hoc_Lớp_7_Bai02_Bai_2_Phan_mem_may_tinh.docx'
doc = Document(p)
tbl = doc.tables[0]

for ri, r in enumerate(tbl._tbl.findall(qn('w:tr'))):
    for ci, c in enumerate(r.findall(qn('w:tc'))):
        tcPr = c.find(qn('w:tcPr'))
        print(f"\nRow {ri}, Cell {ci}:")
        if tcPr is not None:
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is not None:
                for b in tcBorders:
                    print("  tcBorder:", b.tag.split('}')[-1], b.attrib)
            else:
                print("  No tcBorders in cell")
        else:
            print("  No tcPr in cell")
