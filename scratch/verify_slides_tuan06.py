# -*- coding: utf-8 -*-
import sys, os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
sys.stdout.reconfigure(encoding='utf-8')

SAFE_TOP = 1.15 - 0.01  # tolerance
SAFE_BOTTOM = 6.35 + 0.01

FILES = [
    r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lớp_5_Tiet05_Bai_3_Tim_kiem_thong_tin_Tiet_1.pptx",
    r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lớp_5_Tiet06_Bai_3_Tim_kiem_thong_tin_Tiet_2.pptx",
    r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lop_5_Bai03_Tuan06.pptx",
    r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lớp_7_Tiet05_Bai_3_Quan_ly_du_lieu_trong_may_tinh_Tiet_1.pptx",
    r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lớp_7_Tiet06_Bai_3_Quan_ly_du_lieu_trong_may_tinh_Tiet_2.pptx",
    r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lop_7_Bai03_Tuan06.pptx"
]

all_passed = True

for fp in FILES:
    print(f"\n=======================================================")
    print(f"VERIFYING: {os.path.basename(fp)}")
    print(f"=======================================================")
    if not os.path.exists(fp):
        print(f"❌ File not found: {fp}")
        all_passed = False
        continue

    prs = Presentation(fp)
    print(f"Slide count: {len(prs.slides)}")
    
    violations = 0
    pic_count = 0
    orange_box_count = 0
    
    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            # Check safe zone
            top = shape.top.inches
            height = shape.height.inches
            bottom = top + height
            
            # Allow background full rectangle (top=SAFE_TOP)
            if top < SAFE_TOP:
                print(f"  [!] Slide {s_idx+1}: Top violation: top={top:.2f} in < {SAFE_TOP:.2f} in (Shape: {shape.name})")
                violations += 1
            if bottom > SAFE_BOTTOM:
                print(f"  [!] Slide {s_idx+1}: Bottom violation: bottom={bottom:.2f} in > {SAFE_BOTTOM:.2f} in (Shape: {shape.name})")
                violations += 1
                
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pic_count += 1
                
            # Check orange conclusion
            if shape.has_text_frame and "GHI NHỚ TRỌNG TÂM" in shape.text_frame.text:
                orange_box_count += 1
                
    print(f"Total Pictures embedded: {pic_count}")
    print(f"Orange Conclusion Boxes: {orange_box_count}")
    print(f"Safe Zone Violations: {violations}")
    
    if violations == 0 and pic_count > 0:
        print(f"✅ PASSED 100% UNIGO STANDARDS: {os.path.basename(fp)}")
    else:
        print(f"❌ FAILED CHECKS: {os.path.basename(fp)}")
        all_passed = False

if all_passed:
    print("\n🎉 ALL 6 SLIDE DECKS FULLY VERIFIED AND PASSED 100%!")
else:
    print("\n⚠️ SOME CHECKS FAILED!")
