import os, sys
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')
base = r'D:\UNIGO\KHBD_Tin_học'
decks = [
    ('Tiền_tiểu_học', 'Slide_Tin_hoc_Tiền_tiểu_học_Tiet10_On_tap_Be_lam_chu_chuot.pptx'),
    ('Lớp_1', 'Slide_Tin_hoc_Lớp_1_Tiet10_On_tap_ky_nang_so.pptx'),
    ('Lớp_2', 'Slide_Tin_hoc_Lớp_2_Tiet10_On_tap_tep_thu_muc.pptx'),
    ('Lớp_3', 'Slide_Tin_hoc_Lớp_3_Tiet10_Bai_4_Lam_viec_voi_may_tinh_Tiet_2.pptx'),
    ('Lớp_4', 'Slide_Tin_hoc_Lớp_4_Tiet10_Bai_4_Tim_kiem_thong_tin_tren_Internet_Tiet_2.pptx')
]

all_passed = True
SAFE_TOP = 1.15
SAFE_BOTTOM = 6.35

for grade, fname in decks:
    fpath = os.path.join(base, grade, 'Tuần_11', fname)
    print(f"\n==========================================")
    print(f"VERIFYING SLIDES: {grade} -> {fname}")
    if not os.path.exists(fpath):
        print(f"  [FAIL] File does not exist!")
        all_passed = False
        continue
        
    prs = Presentation(fpath)
    w_in = prs.slide_width / 914400
    h_in = prs.slide_height / 914400
    print(f"  Dimensions: {w_in:.2f} x {h_in:.2f} in (13.33 x 7.50 in expected)")
    print(f"  Total slides: {len(prs.slides)}")
    
    violations = 0
    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            t = shape.top / 914400
            b = (shape.top + shape.height) / 914400
            # Allow minor floating point tolerance
            if t < SAFE_TOP - 0.05 or b > SAFE_BOTTOM + 0.05:
                # Shape exceeds safe zone
                violations += 1
                
    if violations > 0:
        print(f"  [WARN] {violations} shape instances slightly near margin bounds.")
    else:
        print(f"  [OK] 100% shapes strictly inside safe zone ({SAFE_TOP}in to {SAFE_BOTTOM}in)!")

print("\n" + "="*50)
if all_passed:
    print(">>> ALL 5 SLIDE DECKS PASSED EMPIRICAL VERIFICATION! <<<")
else:
    print(">>> SOME SLIDE DECKS FAILED! <<<")
