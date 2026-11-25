# -*- coding: utf-8 -*-
r"""
build_all_ck1_exams_and_reviews.py
==================================
Xuất trọn bộ 32 tệp Word (.docx) chuẩn UNIGO (Căn đều 2 bên Justify, bảng cân đối 100% vùng in, bỏ bảng chữ ký, tách riêng Đề & Đáp án):
1. De_kiem_tra_Cuoi_ki_1/          (8 tệp Đề thi chính thức phát cho học sinh Lớp 1 -> Lớp 8)
2. Dap_an_De_kiem_tra_Cuoi_ki_1/   (8 tệp Đáp án, Hướng dẫn chấm & Ma trận Đề thi Lớp 1 -> Lớp 8)
3. De_cuong_on_tap_Cuoi_ki_1/      (8 tệp Đề cương lý thuyết & 80 câu ôn tập phát cho học sinh Lớp 1 -> Lớp 8)
4. Dap_an_De_cuong_Cuoi_ki_1/      (8 tệp Đáp án & Hướng dẫn giải chi tiết 80 câu Đề cương Lớp 1 -> Lớp 8)
"""

import os
import sys
import importlib
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ck1_engine import (
    get_existing_or_generate_images,
    validate_grade_data,
    build_exam_and_answer_pair,
    build_review_and_answer_pair,
    OUT_ROOT,
)


def audit_single_docx(fp, expect_no_signature=True):
    doc = Document(fp)
    sec = doc.sections[0]
    hdr_draw = len(sec.header._element.findall(f'.//{qn("w:drawing")}'))
    ftr_paras = len(sec.footer.paragraphs)
    soft_br = sum(
        1 for p in doc.paragraphs
        if "\n" in p.text or any(br.get(qn("w:type")) != "page" for br in p._p.findall(f'.//{qn("w:br")}'))
    )
    just_cnt = sum(1 for p in doc.paragraphs if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY)
    full_w_tbls = sum(1 for t in doc.tables if 'w:w="5000"' in t._tbl.tblPr.xml and 'w:type="pct"' in t._tbl.tblPr.xml)

    has_sig = False
    for t in doc.tables:
        t_txt = " ".join(c.text for r in t.rows for c in r.cells).upper()
        if "DUYỆT CỦA BGH" in t_txt or "NGƯỜI SOẠN" in t_txt or "GIÁO VIÊN RA ĐỀ" in t_txt:
            has_sig = True

    errs = []
    if hdr_draw < 1:
        errs.append("Mất Logo Header UNIGO")
    if ftr_paras < 1:
        errs.append("Mất Footer UNIGO")
    if soft_br > 0:
        errs.append(f"Còn {soft_br} đoạn ngắt dòng mềm")
    if just_cnt == 0:
        errs.append("Chưa có đoạn văn căn đều 2 bên (Justify)")
    if full_w_tbls != len(doc.tables):
        errs.append(f"Bảng chưa đạt 100% độ rộng ({full_w_tbls}/{len(doc.tables)})")
    if expect_no_signature and has_sig:
        errs.append("Vẫn còn bảng chữ ký cuối bài")
    return errs, len(doc.paragraphs), len(doc.tables), just_cnt


def main():
    img_map = get_existing_or_generate_images()
    print(f"🎨 Đã nạp {len(img_map)} hình minh họa Tiếng Việt từ {os.path.join(OUT_ROOT, 'images')}")

    issues = []
    built_files = []

    for lop in range(1, 9):
        mod = importlib.import_module(f"ck1_data.lop_{lop}")
        gdata = mod.GRADE_DATA
        v_errs = validate_grade_data(gdata)
        if v_errs:
            issues.extend(v_errs)
            continue

        exam_p, exam_ans_p, _ = build_exam_and_answer_pair(gdata, img_map)
        rev_p, rev_ans_p = build_review_and_answer_pair(gdata, img_map)

        # Kiểm tra Đề thi không chứa chữ "BẢNG ĐÁP ÁN" hay "HƯỚNG DẪN CHẤM"
        d_ex = Document(exam_p)
        ex_all_txt = "\n".join(p.text for p in d_ex.paragraphs)
        if "BẢNG ĐÁP ÁN" in ex_all_txt or "HƯỚNG DẪN CHẤM" in ex_all_txt or "KHUNG MA TRẬN" in ex_all_txt:
            issues.append(f"[Lớp {lop}] File Đề thi vẫn còn dính Ma trận / Đáp án!")

        for tag, fp in [("Đề thi", exam_p), ("Đáp án Đề thi", exam_ans_p), ("Đề cương", rev_p), ("Đáp án Đề cương", rev_ans_p)]:
            f_errs, n_p, n_t, n_j = audit_single_docx(fp)
            for e in f_errs:
                issues.append(f"[Lớp {lop} - {tag}] {e}")
            built_files.append((lop, tag, os.path.basename(fp), n_p, n_t, n_j))

        print(f"  ✅ Lớp {lop}: Đã xuất tách rời 4 tệp (Đề thi | Đáp án Đề thi | Đề cương 80 câu | Đáp án Đề cương)!")

    print("\n=================================================================================")
    print(f"📊 TỔNG SỐ TỆP ĐÃ XUẤT & KIỂM ĐỊNH: {len(built_files)} / 32 tệp Word (.docx)")
    print(f"🔍 Số lỗi trình bày / bảng / chữ ký / trộn đáp án: {len(issues)}")
    for iss in issues:
        print("  ❌", iss)
    print("=================================================================================")
    if issues:
        sys.exit(1)


if __name__ == "__main__":
    main()
