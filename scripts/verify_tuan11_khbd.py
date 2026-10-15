import os, sys, glob
import docx
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
base = r'D:\UNIGO\KHBD_Tin_học'
lessons = [
    ('Tiền_tiểu_học', 'KHBD_Tin_hoc_Tiền_tiểu_học_Tiet10_On_tap_Be_lam_chu_chuot.docx'),
    ('Lớp_1', 'KHBD_Tin_hoc_Lớp_1_Tiet10_On_tap_ky_nang_so.docx'),
    ('Lớp_2', 'KHBD_Tin_hoc_Lớp_2_Tiet10_On_tap_tep_thu_muc.docx'),
    ('Lớp_3', 'KHBD_Tin_hoc_Lớp_3_Tiet10_Bai_4_Lam_viec_voi_may_tinh_Tiet_2.docx'),
    ('Lớp_4', 'KHBD_Tin_hoc_Lớp_4_Tiet10_Bai_4_Tim_kiem_thong_tin_tren_Internet_Tiet_2.docx')
]

all_passed = True

for grade, fname in lessons:
    fpath = os.path.join(base, grade, 'Tuần_11', fname)
    print(f"\n==========================================")
    print(f"VERIFYING: {grade} -> {fname}")
    if not os.path.exists(fpath):
        print(f"  [FAIL] File does not exist!")
        all_passed = False
        continue
        
    doc = docx.Document(fpath)
    
    # Check Header
    drawings = 0
    for s in doc.sections:
        for p in s.header.paragraphs:
            for r in p.runs:
                if 'drawing' in r._r.xml:
                    drawings += 1
    print(f"  Header drawings: {drawings} (>= 1 required)")
    if drawings < 1:
        print(f"  [FAIL] Missing header drawing!")
        all_passed = False
        
    # Check tables count
    print(f"  Total tables: {len(doc.tables)} (3 required)")
    if len(doc.tables) != 3:
        print(f"  [FAIL] Table count is not 3!")
        all_passed = False
        continue
        
    t0 = doc.tables[0]
    t1 = doc.tables[1]
    t2 = doc.tables[2]
    
    # Check Table 0
    print(f"  Table 0: {len(t0.rows)}x{len(t0.columns)} (3x2 required)")
    if len(t0.rows) != 3 or len(t0.columns) != 2:
        print(f"  [FAIL] Table 0 dimensions incorrect!")
        all_passed = False
    print(f"    R0: {[c.text.strip() for c in t0.rows[0].cells]}")
    print(f"    R1: {[c.text.strip() for c in t0.rows[1].cells]}")
    print(f"    R2: {[c.text.strip() for c in t0.rows[2].cells]}")
    
    # Check Table 1
    print(f"  Table 1: {len(t1.rows)} rows, 2 columns required")
    if len(t1.columns) != 2:
        print(f"  [FAIL] Table 1 column count is not 2!")
        all_passed = False
        
    # Check Table 2
    print(f"  Table 2: {len(t2.rows)}x{len(t2.columns)} (3x3 required)")
    if len(t2.rows) != 3 or len(t2.columns) != 3:
        print(f"  [FAIL] Table 2 dimensions incorrect!")
        all_passed = False
    print(f"    R0: {[c.text.strip() for c in t2.rows[0].cells]}")
    print(f"    R2: {[c.text.strip() for c in t2.rows[2].cells]}")

print("\n" + "="*50)
if all_passed:
    print(">>> ALL 5 KHBD PASSED EMPIRICAL VERIFICATION! <<<")
else:
    print(">>> SOME CHECKS FAILED! <<<")
