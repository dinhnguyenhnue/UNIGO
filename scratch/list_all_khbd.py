# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

# Read ALL KHBD files to understand current naming pattern
base = r'D:\UNIGO\KHBD_Robotics'
for lop in sorted(os.listdir(base)):
    lop_dir = os.path.join(base, lop)
    if not os.path.isdir(lop_dir):
        continue
    print(f'\n=== {lop} ===')
    for tuan in sorted(os.listdir(lop_dir)):
        tuan_dir = os.path.join(lop_dir, tuan)
        if not os.path.isdir(tuan_dir):
            continue
        files = os.listdir(tuan_dir)
        if files:
            for f in sorted(files):
                print(f'  {tuan}/{f}')
