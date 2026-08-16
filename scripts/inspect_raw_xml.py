import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.oxml import xmlchemy
from xml.etree import ElementTree as ET

p = r'd:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_03\KHBD_Tin_hoc_Lớp_7_Bai02_Bai_2_Phan_mem_may_tinh.docx'
doc = Document(p)
tbl = doc.tables[0]

# Print raw XML of table 0
xml_str = tbl._tbl.xml
print("=== TABLE 0 RAW XML ===")
print(xml_str[:2500])
