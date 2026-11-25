# -*- coding: utf-8 -*-
r"""
ck1_engine.py
=============
Bộ máy chuẩn hóa xuất bản Word (.docx) môn Tin học UNIGO (Lớp 1 -> Lớp 8) cho Cuối học kì 1 (Đánh giá định kỳ 2):
- Căn đều 2 bên (Justify) toàn bộ văn bản nội dung, tách ngắt đoạn Enter thực sự cho từng ý nhỏ (0 lỗi giãn khoảng trắng).
- Toàn bộ bảng biểu rộng khít 100% vùng in của trang (tblW 100%, fixed layout, căn giữa dọc ô, lưới đáp án chia hết vuông vức 7 cột / 8 cột).
- Không chèn bảng chữ ký cuối bài trên các bản Đề thi, Đáp án và Đề cương.
- Tách bạch độc lập thành 4 bộ tệp:
  1. De_kiem_tra_Cuoi_ki_1/          (8 tệp Đề thi chính thức phát cho Học sinh)
  2. Dap_an_De_kiem_tra_Cuoi_ki_1/   (8 tệp Đáp án, Hướng dẫn chấm & Ma trận dành cho Giáo viên)
  3. De_cuong_on_tap_Cuoi_ki_1/      (8 tệp Đề cương lý thuyết + 80 câu hỏi ôn tập phát cho Học sinh)
  4. Dap_an_De_cuong_Cuoi_ki_1/      (8 tệp Đáp án 56 câu TN & Hướng dẫn giải 24 câu TL của Đề cương)
"""

import os
import re
import sys
import random
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_TH = r"D:\UNIGO\Hệ thống mẫu văn bản\Khung_KHBD_Tieu_hoc_UNIGO.docx"
TEMPLATE_THCS = r"D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx"

OUT_ROOT = r"D:\UNIGO\Kiemtra_danhgia\Cuoi_ki_1_Tin_hoc"
OUT_EXAM_DIR = os.path.join(OUT_ROOT, "De_kiem_tra_Cuoi_ki_1")
OUT_EXAM_ANS_DIR = os.path.join(OUT_ROOT, "Dap_an_De_kiem_tra_Cuoi_ki_1")
OUT_REV_DIR = os.path.join(OUT_ROOT, "De_cuong_on_tap_Cuoi_ki_1")
OUT_REV_ANS_DIR = os.path.join(OUT_ROOT, "Dap_an_De_cuong_Cuoi_ki_1")
OUT_IMG_DIR = os.path.join(OUT_ROOT, "images")
KHBD_ROOT = r"D:\UNIGO\KHBD_Tin_học"


# ═════════════════════════════════════════════════════════════════════════════
# 1. QUẢN LÝ HÌNH MINH HỌA TRỰC QUAN 100% TIẾNG VIỆT (16 HÌNH LỚP 1..8)
# ═════════════════════════════════════════════════════════════════════════════

def get_existing_or_generate_images():
    """Trả về dict 16 hình ảnh minh họa Tiếng Việt đã sinh sẵn trong OUT_IMG_DIR."""
    os.makedirs(OUT_IMG_DIR, exist_ok=True)
    keys = [
        "g1_bo_phan_may_tinh", "g1_me_cung_robot",
        "g2_tep_va_thu_muc", "g2_thuat_toan_chuoi_lenh",
        "g3_so_do_xu_li_thong_tin", "g3_khu_vuc_ban_phim",
        "g4_phan_cung_phan_mem", "g4_cay_thu_muc_powerpoint",
        "g5_cay_thu_muc_tim_kiem", "g5_word_wrap_va_paint",
        "g6_mang_may_tinh_internet", "g6_trinh_duyet_va_email",
        "g7_thiet_bi_va_phan_mem", "g7_bang_tinh_excel_mau",
        "g8_the_he_may_tinh_tin_cay", "g8_excel_tuyet_doi_bieu_do",
    ]
    paths = {}
    for k in keys:
        p = os.path.join(OUT_IMG_DIR, f"{k}.png")
        if os.path.exists(p):
            paths[k] = p
    return paths


# ═════════════════════════════════════════════════════════════════════════════
# 2. KIỂM ĐỊNH DỮ LIỆU & CÂN BẰNG ĐÁP ÁN CỐ ĐỊNH (BẢO TOÀN 100% NỘI DUNG)
# ═════════════════════════════════════════════════════════════════════════════

FORBIDDEN_OPT_SUBSTR = (
    "tất cả các đáp án",
    "tất cả các phương án",
    "tất cả các ý trên",
    "cả a và b",
    "cả b và c",
    "cả 3 phương án",
    "cả ba phương án",
    "phương án trên",
    "đáp án trên",
)


def validate_grade_data(data):
    lop = data["lop"]
    exp_exam_mcq = 14 if lop in (1, 2) else 28
    exp_exam_nb = 10 if lop in (1, 2) else 20
    exp_exam_th = 4 if lop in (1, 2) else 8

    errs = []
    emcq = data.get("exam_mcq", [])
    etl = data.get("exam_tl", [])
    if len(emcq) != exp_exam_mcq:
        errs.append(f"Lớp {lop}: exam_mcq có {len(emcq)} câu (cần {exp_exam_mcq})")
    nb_c = sum(1 for q in emcq if q.get("level") == "NB")
    th_c = sum(1 for q in emcq if q.get("level") == "TH")
    if nb_c != exp_exam_nb or th_c != exp_exam_th:
        errs.append(f"Lớp {lop}: exam_mcq NB={nb_c}, TH={th_c}")
    if len(etl) != 3:
        errs.append(f"Lớp {lop}: exam_tl có {len(etl)} câu (cần 3 câu)")

    r_nb = data.get("review_nb_mcq", [])
    r_th_m = data.get("review_th_mcq", [])
    r_th_t = data.get("review_th_tl", [])
    r_vd = data.get("review_vd_tl", [])
    r_vdc = data.get("review_vdc_tl", [])

    if (len(r_nb), len(r_th_m), len(r_th_t), len(r_vd), len(r_vdc)) != (40, 16, 8, 8, 8):
        errs.append(f"Lớp {lop}: Số câu đề cương lệch chuẩn 40-16-8-8-8")

    for sec_name, qlist in [("exam_mcq", emcq), ("review_nb_mcq", r_nb), ("review_th_mcq", r_th_m)]:
        seen_q = set()
        for idx, item in enumerate(qlist, 1):
            qt = item.get("q", "").strip()
            opts = [o.strip() for o in item.get("options", [])]
            ans = item.get("ans", -1)
            norm_q = re.sub(r'\s+', ' ', qt.lower())
            if not qt or norm_q in seen_q:
                errs.append(f"Lớp {lop} {sec_name}[{idx}]: Câu hỏi rỗng hoặc trùng")
            seen_q.add(norm_q)
            if len(opts) != 4 or len(set(o.lower() for o in opts)) != 4 or ans not in (0, 1, 2, 3):
                errs.append(f"Lớp {lop} {sec_name}[{idx}]: Lỗi phương án {opts}")
            for o in opts:
                if any(fb in o.lower() for fb in FORBIDDEN_OPT_SUBSTR) or re.match(r'^[A-D][\.\)]\s', o):
                    errs.append(f"Lớp {lop} {sec_name}[{idx}]: Lỗi định dạng đáp án '{o}'")
    return errs


def balance_mcq_options(mcq_list, seed=2026):
    """Giữ nguyên seed cố định để thứ tự phương án A, B, C, D không thay đổi so với bản trước."""
    rng = random.Random(seed)
    n = len(mcq_list)
    letters = ["A", "B", "C", "D"]
    targets = [i % 4 for i in range(n)]
    rng.shuffle(targets)

    balanced = []
    for item, tgt_idx in zip(mcq_list, targets):
        opts = list(item["options"])
        correct_text = opts[item["ans"]]
        wrong_texts = [opts[i] for i in range(4) if i != item["ans"]]
        rng.shuffle(wrong_texts)
        new_opts = []
        w_ptr = 0
        for pos in range(4):
            if pos == tgt_idx:
                new_opts.append(f"{letters[pos]}. {correct_text}")
            else:
                new_opts.append(f"{letters[pos]}. {wrong_texts[w_ptr]}")
                w_ptr += 1
        new_item = dict(item)
        new_item["formatted_options"] = new_opts
        new_item["ans_letter"] = letters[tgt_idx]
        balanced.append(new_item)
    return balanced


# ═════════════════════════════════════════════════════════════════════════════
# 3. ĐỊNH DẠNG VĂN BẢN CĂN ĐỀU 2 BÊN (JUSTIFY) & BẢNG CÂN ĐỐI 100% VÙNG IN
# ═════════════════════════════════════════════════════════════════════════════

def prepare_clean_doc(lop):
    tpl = TEMPLATE_TH if lop <= 5 else TEMPLATE_THCS
    doc = Document(tpl)
    body = doc.element.body
    for child in list(body):
        if not child.tag.endswith("sectPr"):
            body.remove(child)

    # Cờ chống kéo giãn khoảng trắng nếu có ngắt dòng mềm
    settings_el = doc.settings.element
    compat = settings_el.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings_el.append(compat)
    if compat.find(qn("w:doNotExpandShiftReturn")) is None:
        compat.append(OxmlElement("w:doNotExpandShiftReturn"))

    return doc


def get_printable_width_emu(doc):
    sec = doc.sections[0]
    return int(sec.page_width - sec.left_margin - sec.right_margin)


def set_run_font(run, size_pt=13, bold=False, italic=False, color_rgb=None):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    for old_rf in rPr.findall(qn("w:rFonts")):
        rPr.remove(old_rf)
    rPr.append(parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>'
    ))


def add_p(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False,
          size_pt=13, space_before=0, space_after=3, color_rgb=None, keep_with_next=False, left_indent_in=0.0):
    """Thêm đoạn văn bản; nếu `text` có chứa `\\n` thì tự động tách thành nhiều đoạn Enter độc lập để căn đều 2 bên không bị giãn chữ."""
    lines = str(text).split("\n") if text else [""]
    last_p = None
    for idx, ln in enumerate(lines):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(space_before if idx == 0 else 1)
        pf.space_after = Pt(space_after if idx == len(lines) - 1 else 2)
        pf.line_spacing = 1.18
        if left_indent_in:
            pf.left_indent = Inches(left_indent_in)
        if keep_with_next:
            pf.keep_with_next = True
        if ln:
            r = p.add_run(ln)
            set_run_font(r, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
        last_p = p
    return last_p


def add_prefix_body_p(doc, prefix_bold, body_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      space_before=4, space_after=2, left_indent_in=0.0, sub_indent_in=0.18, keep_with_next=False):
    """Tạo đoạn có phần đầu in đậm (`Câu X.` / `• Hướng dẫn Câu X:`) + phần thân in thường; tự tách `\\n` thành các đoạn Enter chuẩn."""
    lines = str(body_text).split("\n")
    p0 = doc.add_paragraph()
    p0.alignment = align
    pf0 = p0.paragraph_format
    pf0.space_before = Pt(space_before)
    pf0.space_after = Pt(space_after if len(lines) == 1 else 2)
    pf0.line_spacing = 1.18
    if left_indent_in:
        pf0.left_indent = Inches(left_indent_in)
    if keep_with_next:
        pf0.keep_with_next = True

    r1 = p0.add_run(prefix_bold)
    set_run_font(r1, size_pt=13, bold=True)
    if lines[0]:
        r2 = p0.add_run(lines[0])
        set_run_font(r2, size_pt=13, bold=False)

    for idx, extra_ln in enumerate(lines[1:]):
        is_last = (idx == len(lines) - 2)
        add_p(
            doc, extra_ln, align=align, bold=False, size_pt=13,
            space_before=1, space_after=(space_after if is_last else 2),
            left_indent_in=max(left_indent_in, sub_indent_in),
            keep_with_next=(keep_with_next and not is_last)
        )
    return p0


def add_dot_leader_lines(doc, n_lines):
    """Tạo `n_lines` hàng dòng chấm Tab Leader phẳng tắp từ lề trái đến sát lề phải."""
    pw_emu = max(Emu(1000000), int(get_printable_width_emu(doc) - Pt(2)))
    for _ in range(n_lines):
        pd = doc.add_paragraph()
        pd.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = pd.paragraph_format
        pf.left_indent = Pt(0)
        pf.right_indent = Pt(0)
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(2)
        pf.space_after = Pt(4)
        pf.line_spacing = 1.32
        pf.tab_stops.add_tab_stop(Emu(pw_emu), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        rd = pd.add_run("\t")
        set_run_font(rd, size_pt=13)


def create_balanced_table(doc, n_rows, col_ratios, bordered=True, header_rows=1):
    """
    Tạo bảng cân đối tuyệt đối 100% chiều rộng vùng in (`w:type="pct" w:w="5000"` + `w:tblGrid` chính xác theo dxa),
    không tự co giãn lệch cột (`fixed` layout), chống ngắt dòng giữa hàng (`cantSplit`).
    """
    n_cols = len(col_ratios)
    total_ratio = float(sum(col_ratios))
    norm_ratios = [r / total_ratio for r in col_ratios]

    pw_emu = get_printable_width_emu(doc)
    pw_dxa = int(round(pw_emu / 635.0))
    col_emus = [int(round(pw_emu * r)) for r in norm_ratios]
    col_dxas = [int(round(pw_dxa * r)) for r in norm_ratios]
    # Bù sai số làm tròn vào cột cuối để tổng bằng chính xác 100%
    col_emus[-1] = pw_emu - sum(col_emus[:-1])
    col_dxas[-1] = pw_dxa - sum(col_dxas[:-1])

    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    tblPr = tbl._tbl.tblPr
    # Đặt độ rộng bảng = 100% vùng in (5000 = 100% theo chuẩn OOXML pct)
    for old_w in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old_w)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)

    # Cố định bố cục cột để Word không tự bóp méo cột
    for old_lay in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old_lay)
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    # Đặt viền bảng
    for old_b in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old_b)
    if bordered:
        tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>'
        ))
    else:
        tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
            f'<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
        ))

    # Thiết lập tblGrid chính xác
    old_grid = tbl._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        tbl._tbl.remove(old_grid)
    new_grid = OxmlElement("w:tblGrid")
    for dxa in col_dxas:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(dxa))
        new_grid.append(gc)
    tblPr.addnext(new_grid)

    # Thiết lập thuộc tính từng hàng và từng ô
    for r_idx, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        if bordered and r_idx < header_rows and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))
        for c_idx, cell in enumerate(row.cells):
            cell.width = Emu(col_emus[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=85, bottom=85, left=95, right=95)
            tcPr = cell._tc.get_or_add_tcPr()
            for old_tcw in tcPr.findall(qn("w:tcW")):
                tcPr.remove(old_tcw)
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_dxas[c_idx]))
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)

    return tbl


def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    for old_s in tcPr.findall(qn("w:shd")):
        tcPr.remove(old_s)
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


def set_cell_margins(cell, top=85, bottom=85, left=95, right=95):
    tcPr = cell._tc.get_or_add_tcPr()
    for old_m in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old_m)
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def fmt_cell(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, italic=False, size_pt=13, color_rgb=None, v_align=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = v_align
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    for idx, line in enumerate(str(text).split("\n")):
        if idx > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
        r = p.add_run(line)
        set_run_font(r, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)


def insert_diagram(doc, img_path, caption_text=""):
    if not img_path or not os.path.exists(img_path):
        return
    pw_in = get_printable_width_emu(doc) / 914400.0
    img_w_in = min(5.90, pw_in * 0.94)
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(5)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    p_img.add_run().add_picture(img_path, width=Inches(img_w_in))
    if caption_text:
        add_p(doc, caption_text, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=13, space_after=4, color_rgb=(30, 58, 138))


def render_mcq_item(doc, q_num, q_obj, img_map, show_meta=False):
    """Hiển thị 1 câu trắc nghiệm: Câu dẫn căn đều 2 bên (Justify); đáp án chia cột tab cân đối theo độ rộng trang thực tế."""
    if q_obj.get("image_key") and q_obj["image_key"] in img_map:
        insert_diagram(doc, img_map[q_obj["image_key"]], caption_text=q_obj.get("image_caption", ""))

    prefix = f"Câu {q_num}. "
    if show_meta and q_obj.get("bai"):
        prefix = f"Câu {q_num} [{q_obj['bai']}]. "

    add_prefix_body_p(
        doc, prefix, q_obj["q"],
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        space_before=4, space_after=2, keep_with_next=True
    )

    opts = q_obj["formatted_options"]
    max_len = max(len(o) for o in opts)
    pw_in = get_printable_width_emu(doc) / 914400.0

    if max_len <= 14:
        po = doc.add_paragraph()
        po.alignment = WD_ALIGN_PARAGRAPH.LEFT
        po.paragraph_format.space_before = Pt(0)
        po.paragraph_format.space_after = Pt(3)
        po.paragraph_format.line_spacing = 1.18
        step = (pw_in - 0.20) / 4.0
        for i in range(4):
            po.paragraph_format.tab_stops.add_tab_stop(Inches(0.20 + i * step), WD_TAB_ALIGNMENT.LEFT)
        ro = po.add_run(f"\t{opts[0]}\t{opts[1]}\t{opts[2]}\t{opts[3]}")
        set_run_font(ro, size_pt=13)
    elif max_len <= 32 and max(len(opts[0]), len(opts[2])) <= 27:
        col2_pos = round(pw_in * 0.52, 2)
        for idx, (o1, o2) in enumerate([(opts[0], opts[1]), (opts[2], opts[3])]):
            po = doc.add_paragraph()
            po.alignment = WD_ALIGN_PARAGRAPH.LEFT
            po.paragraph_format.space_before = Pt(0)
            po.paragraph_format.space_after = Pt(1.5 if idx == 0 else 3)
            po.paragraph_format.line_spacing = 1.18
            if idx == 0:
                po.paragraph_format.keep_with_next = True
            po.paragraph_format.tab_stops.add_tab_stop(Inches(0.22), WD_TAB_ALIGNMENT.LEFT)
            po.paragraph_format.tab_stops.add_tab_stop(Inches(col2_pos), WD_TAB_ALIGNMENT.LEFT)
            ro = po.add_run(f"\t{o1}\t{o2}")
            set_run_font(ro, size_pt=13)
    else:
        for idx, opt in enumerate(opts):
            po = doc.add_paragraph()
            po.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            po.paragraph_format.space_before = Pt(0)
            po.paragraph_format.space_after = Pt(1.5 if idx < 3 else 3)
            po.paragraph_format.line_spacing = 1.18
            if idx < 3:
                po.paragraph_format.keep_with_next = True
            po.paragraph_format.left_indent = Inches(0.24)
            ro = po.add_run(opt)
            set_run_font(ro, size_pt=13)


def build_school_title_header(doc, lop, doc_kind, thoi_gian):
    """Bảng tiêu đề Trường Tiểu học & THCS UNIGO không viền, rộng 100% vùng in cân đối 2 cột."""
    is_thcs = lop >= 6
    to_cm = "TỔ CHUYÊN MÔN THCS" if is_thcs else "TỔ CHUYÊN MÔN TIỂU HỌC"
    tbl = create_balanced_table(doc, n_rows=1, col_ratios=[42, 58], bordered=False)
    c0, c1 = tbl.rows[0].cells
    set_cell_margins(c0, 40, 60, 40, 40)
    set_cell_margins(c1, 40, 60, 40, 40)
    fmt_cell(c0, f"TRƯỜNG TIỂU HỌC & THCS UNIGO\n{to_cm}\n───────────────", bold=True)

    titles = {
        "EXAM": (
            f"ĐỀ KIỂM TRA CUỐI HỌC KÌ 1 (ĐÁNH GIÁ ĐỊNH KỲ 2)\n"
            f"NĂM HỌC 2026 – 2027 | MÔN: TIN HỌC {lop}\n"
            f"Thời gian làm bài: {thoi_gian} (Không kể thời gian phát đề)"
        ),
        "EXAM_ANS": (
            f"ĐÁP ÁN, HƯỚNG DẪN CHẤM & MA TRẬN ĐỀ CUỐI HỌC KÌ 1\n"
            f"NĂM HỌC 2026 – 2027 | MÔN: TIN HỌC {lop}\n"
            f"(Kèm theo Đề kiểm tra Đánh giá định kỳ 2 – Thời gian: {thoi_gian})"
        ),
        "REVIEW": (
            f"ĐỀ CƯƠNG ÔN TẬP CUỐI HỌC KÌ 1 (80 CÂU HỎI)\n"
            f"NĂM HỌC 2026 – 2027 | MÔN: TIN HỌC {lop}\n"
            f"(Phạm vi kiến thức: Từ đầu năm học đến Đánh giá định kỳ 2)"
        ),
        "REVIEW_ANS": (
            f"ĐÁP ÁN & HƯỚNG DẪN GIẢI ĐỀ CƯƠNG ÔN TẬP CUỐI HỌC KÌ 1\n"
            f"NĂM HỌC 2026 – 2027 | MÔN: TIN HỌC {lop}\n"
            f"(Đáp án chi tiết cho bộ 80 câu hỏi ôn tập Học kì 1)"
        ),
    }
    fmt_cell(c1, titles[doc_kind], bold=True)
    add_p(doc, "", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=3)


def build_student_box(doc, lop_hien_thi):
    """Khung thông tin Học sinh + Điểm & Lời phê cân đối tuyệt đối 100% vùng in (để trống Lớp cho HS tự ghi, bỏ SBD/Số máy)."""
    tbl = create_balanced_table(doc, n_rows=3, col_ratios=[26, 74], bordered=True, header_rows=0)
    # Hàng 0: Gộp 2 cột cho Họ và tên, Lớp (trống), Ngày kiểm tra (đúng 2 dòng phẳng mép bằng Tab Leader)
    c_info = tbl.cell(0, 0).merge(tbl.cell(0, 1))
    set_cell_margins(c_info, top=95, bottom=95, left=120, right=120)
    c_info.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    pw_emu = get_printable_width_emu(doc)
    inner_w_emu = int(pw_emu - Inches(0.22))
    tab1_emu = int(inner_w_emu * 0.72)
    tab2_emu = int(inner_w_emu - Pt(2))

    p_line1 = c_info.paragraphs[0]
    p_line1.clear()
    p_line1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_line1.paragraph_format.space_before = Pt(2)
    p_line1.paragraph_format.space_after = Pt(4)
    p_line1.paragraph_format.line_spacing = 1.20
    p_line1.paragraph_format.tab_stops.add_tab_stop(Emu(tab1_emu), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.DOTS)
    p_line1.paragraph_format.tab_stops.add_tab_stop(Emu(tab2_emu), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    for txt_seg in ["Họ và tên học sinh: ", "\t", "   Lớp: ", "\t"]:
        r = p_line1.add_run(txt_seg)
        set_run_font(r, size_pt=13)

    p_line2 = c_info.add_paragraph()
    p_line2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_line2.paragraph_format.space_before = Pt(2)
    p_line2.paragraph_format.space_after = Pt(2)
    p_line2.paragraph_format.line_spacing = 1.20
    r_dt = p_line2.add_run("Ngày kiểm tra: ....... / ....... / 2026")
    set_run_font(r_dt, size_pt=13)

    # Hàng 1: Tiêu đề ĐIỂM SỐ | NHẬN XÉT CỦA GIÁO VIÊN
    set_cell_bg(tbl.cell(1, 0), "D9E8F6")
    set_cell_bg(tbl.cell(1, 1), "D9E8F6")
    fmt_cell(tbl.cell(1, 0), "ĐIỂM SỐ", bold=True)
    fmt_cell(tbl.cell(1, 1), "LỜI NHẬN XÉT CỦA GIÁO VIÊN", bold=True)

    # Hàng 2: Ô trống chấm điểm & Ô nhận xét
    fmt_cell(tbl.cell(2, 0), "\n\n\n")
    c_cmt = tbl.cell(2, 1)
    c_cmt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = c_cmt.paragraphs[0]
    p0.clear()
    pw_emu = get_printable_width_emu(doc)
    cmt_w_emu = int(pw_emu * 0.74 - Inches(0.28))
    for i in range(3):
        pc = p0 if i == 0 else c_cmt.add_paragraph()
        pc.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pc.paragraph_format.space_before = Pt(2)
        pc.paragraph_format.space_after = Pt(3)
        pc.paragraph_format.line_spacing = 1.25
        pc.paragraph_format.tab_stops.add_tab_stop(Emu(cmt_w_emu), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        rc = pc.add_run("\t")
        set_run_font(rc, size_pt=13)

    add_p(doc, "", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)


def build_matrix_exam_table(doc, rows_mt):
    """Bảng Ma trận Đề kiểm tra 7 cột rộng khít 100% vùng in, tỷ lệ cột cân đối."""
    headers = [
        "Chủ đề / Nội dung kiến thức",
        "Nhận biết\n(TN)",
        "Thông hiểu\n(TN)",
        "Thông hiểu\n(TL)",
        "Vận dụng\n(TL)",
        "Vận dụng\ncao (TL)",
        "Tổng điểm"
    ]
    tbl = create_balanced_table(doc, n_rows=len(rows_mt) + 1, col_ratios=[36, 10.5, 10.5, 10.5, 10.5, 11, 11], bordered=True, header_rows=1)
    for ci, h in enumerate(headers):
        c = tbl.cell(0, ci)
        set_cell_bg(c, "D9E8F6")
        fmt_cell(c, h, bold=True)
    for ri, rdata in enumerate(rows_mt, start=1):
        is_tot = (ri == len(rows_mt))
        for ci, val in enumerate(rdata):
            c = tbl.cell(ri, ci)
            if is_tot:
                set_cell_bg(c, "EEF2FF")
            fmt_cell(
                c, str(val),
                align=(WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER),
                bold=(is_tot or ci == 0)
            )
    add_p(doc, "", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)


def build_answer_grid(doc, balanced_mcqs):
    """
    Bảng đáp án trắc nghiệm vuông vức tuyệt đối 100% vùng in (không có ô thừa lẻ):
    - 14 câu -> 7 cột × 2 cặp hàng (4 hàng)
    - 28 câu -> 7 cột × 4 cặp hàng (8 hàng)
    - 56 câu -> 8 cột × 7 cặp hàng (14 hàng)
    """
    total = len(balanced_mcqs)
    per_row = 8 if total % 8 == 0 else 7
    chunks = [balanced_mcqs[i:i + per_row] for i in range(0, total, per_row)]
    tbl = create_balanced_table(doc, n_rows=len(chunks) * 2, col_ratios=[1] * per_row, bordered=True, header_rows=0)

    for c_idx, chunk in enumerate(chunks):
        r_hdr = c_idx * 2
        r_ans = c_idx * 2 + 1
        base_num = c_idx * per_row + 1
        for col_i in range(per_row):
            ch = tbl.cell(r_hdr, col_i)
            ca = tbl.cell(r_ans, col_i)
            set_cell_bg(ch, "D9E8F6")
            if col_i < len(chunk):
                fmt_cell(ch, f"Câu {base_num + col_i}", bold=True)
                fmt_cell(ca, chunk[col_i]["ans_letter"], bold=True, color_rgb=(180, 0, 0))
            else:
                fmt_cell(ch, "-")
                fmt_cell(ca, "-")
    add_p(doc, "", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)


def enforce_tnr_13pt_all(doc):
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(13)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(13)


def save_safe(doc, out_path):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path


# ═════════════════════════════════════════════════════════════════════════════
# 4. XUẤT 4 BỘ TỆP ĐỘC LẬP (ĐỀ THI / ĐÁP ÁN ĐỀ THI / ĐỀ CƯƠNG / ĐÁP ÁN ĐỀ CƯƠNG)
# ═════════════════════════════════════════════════════════════════════════════

def build_exam_and_answer_pair(data, img_map):
    """Xuất tách biệt: (1) File Đề thi học sinh & (2) File Đáp án + Hướng dẫn chấm + Ma trận."""
    lop = data["lop"]
    balanced_mcq = balance_mcq_options(data["exam_mcq"], seed=100 + lop)
    pts_per_mcq = "0,5 điểm" if lop in (1, 2) else "0,25 điểm"
    n_mcq = len(balanced_mcq)

    # ── FILE 1: ĐỀ KIỂM TRA CHÍNH THỨC (CHỈ PHẦN ĐỀ IN PHÁT HỌC SINH, BỎ CHỮ KÝ) ──
    doc_exam = prepare_clean_doc(lop)
    build_school_title_header(doc_exam, lop, "EXAM", data["thoi_gian"])
    build_student_box(doc_exam, data["lop_hien_thi"])

    add_p(
        doc_exam,
        f"PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (7,0 điểm – gồm {n_mcq} câu, mỗi câu đúng được {pts_per_mcq})",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=4, keep_with_next=True
    )
    add_p(
        doc_exam,
        "Em hãy khoanh tròn vào chữ cái (A, B, C hoặc D) đặt trước câu trả lời đúng nhất:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, italic=True, space_after=4
    )

    for idx, q_obj in enumerate(balanced_mcq, start=1):
        render_mcq_item(doc_exam, idx, q_obj, img_map, show_meta=False)

    add_p(
        doc_exam,
        "PHẦN II. TỰ LUẬN VÀ BÀI TẬP VẬN DỤNG (3,0 điểm – gồm 3 câu)",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=8, keep_with_next=True
    )
    for idx, tl in enumerate(data["exam_tl"], start=1):
        if tl.get("image_key") and tl["image_key"] in img_map:
            insert_diagram(doc_exam, img_map[tl["image_key"]], caption_text=tl.get("image_caption", ""))
        add_prefix_body_p(
            doc_exam,
            f"Câu {idx} ({tl['points']} – Mức {tl['level']}): ",
            tl["q"],
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            space_before=4, space_after=2, sub_indent_in=0.18, keep_with_next=True
        )
        add_dot_leader_lines(doc_exam, tl.get("lines", 5))

    add_p(doc_exam, "-------------------- HẾT --------------------", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, space_before=8)
    add_p(doc_exam, "(Học sinh không sử dụng tài liệu. Giám thị coi kiểm tra không giải thích gì thêm)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=2)
    enforce_tnr_13pt_all(doc_exam)
    exam_path = save_safe(doc_exam, os.path.join(OUT_EXAM_DIR, f"De_kiem_tra_Cuoi_ki_1_Tin_hoc_Lop_{lop}.docx"))

    # ── FILE 2: ĐÁP ÁN, HƯỚNG DẪN CHẤM & MA TRẬN ĐỀ KIỂM TRA (RIÊNG BIỆT, BỎ CHỮ KÝ) ──
    doc_ans = prepare_clean_doc(lop)
    build_school_title_header(doc_ans, lop, "EXAM_ANS", data["thoi_gian"])

    add_p(doc_ans, f"I. KHUNG MA TRẬN ĐỀ KIỂM TRA CUỐI HỌC KÌ 1 – MÔN TIN HỌC {lop}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=2, keep_with_next=True)
    add_p(doc_ans, f"• Hình thức kiểm tra: Kết hợp 70% Trắc nghiệm khách quan ({n_mcq} câu = 7,0 điểm) và 30% Tự luận / Vận dụng (3 câu = 3,0 điểm).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2)
    add_p(doc_ans, "• Tỷ lệ mức độ nhận thức (5 – 3 – 1 – 1): Nhận biết 50% (5,0 điểm) – Thông hiểu 30% (3,0 điểm) – Vận dụng 10% (1,0 điểm) – Vận dụng cao 10% (1,0 điểm).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=4)
    build_matrix_exam_table(doc_ans, data["ma_tran_kiem_tra"])

    add_p(doc_ans, f"II. BẢNG ĐÁP ÁN PHẦN TRẮC NGHIỆM KHÁCH QUAN (7,0 điểm – {n_mcq} câu × {pts_per_mcq}):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=6, space_after=3, keep_with_next=True)
    build_answer_grid(doc_ans, balanced_mcq)

    add_p(doc_ans, "III. HƯỚNG DẪN CHẤM CHI TIẾT PHẦN TỰ LUẬN (3,0 điểm):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=6, space_after=3, keep_with_next=True)
    tbl_rb = create_balanced_table(doc_ans, n_rows=len(data["exam_tl"]) + 1, col_ratios=[18, 66, 16], bordered=True, header_rows=1)
    rb_headers = ["Câu hỏi", "Nội dung đáp án & Yêu cầu cần đạt", "Biểu điểm"]
    for ci, h in enumerate(rb_headers):
        set_cell_bg(tbl_rb.cell(0, ci), "D9E8F6")
        fmt_cell(tbl_rb.cell(0, ci), h, bold=True)
    for ri, tl in enumerate(data["exam_tl"], start=1):
        fmt_cell(tbl_rb.cell(ri, 0), f"Câu {ri}\n({tl['level']})", bold=True)
        rubric_txt = "\n".join(f"- {rb}" for rb in tl["rubric"])
        fmt_cell(tbl_rb.cell(ri, 1), rubric_txt, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        fmt_cell(tbl_rb.cell(ri, 2), tl["points"], bold=True)

    enforce_tnr_13pt_all(doc_ans)
    exam_ans_path = save_safe(doc_ans, os.path.join(OUT_EXAM_ANS_DIR, f"Dap_an_Cuoi_ki_1_Tin_hoc_Lop_{lop}.docx"))
    return exam_path, exam_ans_path, balanced_mcq


def build_review_and_answer_pair(data, img_map):
    """Xuất tách biệt: (1) File Đề cương 80 câu ôn tập phát Học sinh & (2) File Đáp án 80 câu Đề cương (bỏ chữ ký)."""
    lop = data["lop"]
    all_56_mcq = balance_mcq_options(data["review_nb_mcq"] + data["review_th_mcq"], seed=500 + lop)
    nb_40 = all_56_mcq[:40]
    th_16 = all_56_mcq[40:]
    tl_sections = [
        ("1. Mức độ 2 – Thông hiểu (8 câu Tự luận: từ Câu 57 đến Câu 64):", 57, data["review_th_tl"]),
        ("2. Mức độ 3 – Vận dụng (8 câu Tự luận: từ Câu 65 đến Câu 72):", 65, data["review_vd_tl"]),
        ("3. Mức độ 4 – Vận dụng cao (8 câu Tự luận: từ Câu 73 đến Câu 80):", 73, data["review_vdc_tl"]),
    ]

    # ── FILE 3: ĐỀ CƯƠNG ÔN TẬP 80 CÂU (PHÁT CHO HỌC SINH ÔN LUYỆN, KHÔNG KÈM ĐÁP ÁN, BỎ CHỮ KÝ) ──
    doc_rev = prepare_clean_doc(lop)
    build_school_title_header(doc_rev, lop, "REVIEW", data["thoi_gian"])

    add_p(doc_rev, f"A. HỆ THỐNG HÓA KIẾN THỨC TRỌNG TÂM HỌC KÌ 1 (TIN HỌC {lop})", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=4, space_after=4)
    shown_imgs = set()
    for sec in data["tom_tat_ly_thuyet"]:
        add_p(doc_rev, f"❖ {sec['bai']}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=4, space_after=2, keep_with_next=True)
        for yc in sec["y_chinh"]:
            add_p(doc_rev, f"- {yc}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent_in=0.20, space_after=2)
        ik = sec.get("image_key")
        if ik and ik in img_map and ik not in shown_imgs:
            insert_diagram(doc_rev, img_map[ik], caption_text=f"Hình minh họa trực quan: {sec['bai']}")
            shown_imgs.add(ik)

    add_p(doc_rev, "B. MA TRẬN CẤU TRÚC 80 CÂU HỎI ÔN TẬP (TỶ LỆ 5 – 3 – 1 – 1)", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=8, space_after=3, keep_with_next=True)
    tbl_sum = create_balanced_table(doc_rev, n_rows=6, col_ratios=[26, 16, 21, 21, 16], bordered=True, header_rows=1)
    s_hdr = ["Mức độ nhận thức", "Tỷ lệ chuẩn", "Trắc nghiệm (70%)", "Tự luận (30%)", "Tổng số câu"]
    s_rows = [
        ["1. Nhận biết (Mức 1)", "5 phần (50%)", "40 câu (Câu 1 – 40)", "0 câu", "40 câu"],
        ["2. Thông hiểu (Mức 2)", "3 phần (30%)", "16 câu (Câu 41 – 56)", "8 câu (Câu 57 – 64)", "24 câu"],
        ["3. Vận dụng (Mức 3)", "1 phần (10%)", "0 câu", "8 câu (Câu 65 – 72)", "8 câu"],
        ["4. Vận dụng cao (Mức 4)", "1 phần (10%)", "0 câu", "8 câu (Câu 73 – 80)", "8 câu"],
        ["TỔNG CỘNG ĐỀ CƯƠNG", "10 phần (100%)", "56 câu TN (70%)", "24 câu TL (30%)", "80 câu (100%)"],
    ]
    for ci, h in enumerate(s_hdr):
        set_cell_bg(tbl_sum.cell(0, ci), "D9E8F6")
        fmt_cell(tbl_sum.cell(0, ci), h, bold=True)
    for ri, rd in enumerate(s_rows, start=1):
        is_tot = (ri == 5)
        for ci, val in enumerate(rd):
            c = tbl_sum.cell(ri, ci)
            if is_tot:
                set_cell_bg(c, "EEF2FF")
            fmt_cell(c, val, align=(WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER), bold=(is_tot or ci == 0))

    add_p(doc_rev, "C. NGÂN HÀNG 80 CÂU HỎI ÔN LUYỆN HỌC KÌ 1 (PHÂN THEO 4 MỨC ĐỘ)", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=10, space_after=4)
    add_p(doc_rev, "PHẦN I. CÂU HỎI TRẮC NGHIỆM KHÁCH QUAN (56 câu – chiếm 70%)", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=4, space_after=2)

    add_p(doc_rev, "1. Mức độ 1 – Nhận biết (40 câu: từ Câu 1 đến Câu 40):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, italic=True, space_before=3, space_after=3, keep_with_next=True)
    for idx, q_obj in enumerate(nb_40, start=1):
        render_mcq_item(doc_rev, idx, q_obj, img_map, show_meta=True)

    add_p(doc_rev, "2. Mức độ 2 – Thông hiểu (16 câu Trắc nghiệm: từ Câu 41 đến Câu 56):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, italic=True, space_before=6, space_after=3, keep_with_next=True)
    for idx, q_obj in enumerate(th_16, start=41):
        render_mcq_item(doc_rev, idx, q_obj, img_map, show_meta=True)

    add_p(doc_rev, "PHẦN II. CÂU HỎI TỰ LUẬN & BÀI TẬP TÌNH HUỐNG (24 câu – chiếm 30%)", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=8, space_after=3, keep_with_next=True)
    for sec_title, start_idx, qlist in tl_sections:
        add_p(doc_rev, sec_title, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, italic=True, space_before=5, space_after=3, keep_with_next=True)
        for offset, item in enumerate(qlist):
            qn_num = start_idx + offset
            tag = f" [{item['bai']}]" if item.get("bai") else ""
            add_prefix_body_p(
                doc_rev, f"Câu {qn_num}{tag}. ", item["q"],
                align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=3, space_after=2, sub_indent_in=0.18
            )

    add_p(doc_rev, "-------------------- HẾT ĐỀ CƯƠNG ÔN TẬP --------------------", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, space_before=8)
    enforce_tnr_13pt_all(doc_rev)
    rev_path = save_safe(doc_rev, os.path.join(OUT_REV_DIR, f"De_cuong_on_tap_Cuoi_ki_1_Tin_hoc_Lop_{lop}.docx"))

    # ── FILE 4: ĐÁP ÁN & HƯỚNG DẪN GIẢI 80 CÂU ĐỀ CƯƠNG ÔN TẬP (RIÊNG BIỆT, BỎ CHỮ KÝ) ──
    doc_rev_ans = prepare_clean_doc(lop)
    build_school_title_header(doc_rev_ans, lop, "REVIEW_ANS", data["thoi_gian"])

    add_p(doc_rev_ans, "I. BẢNG ĐÁP ÁN 56 CÂU TRẮC NGHIỆM KHÁCH QUAN (Câu 1 – Câu 56):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=4, space_after=3, keep_with_next=True)
    build_answer_grid(doc_rev_ans, all_56_mcq)

    add_p(doc_rev_ans, "II. HƯỚNG DẪN GIẢI CHI TIẾT 24 CÂU TỰ LUẬN (Câu 57 – Câu 80):", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, space_before=6, space_after=3, keep_with_next=True)
    for sec_title, start_idx, qlist in tl_sections:
        add_p(doc_rev_ans, sec_title, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True, italic=True, space_before=5, space_after=2, keep_with_next=True)
        for offset, item in enumerate(qlist):
            qn_num = start_idx + offset
            add_prefix_body_p(
                doc_rev_ans,
                f"• Hướng dẫn Câu {qn_num}: ",
                item["answer"],
                align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                space_before=3, space_after=2,
                left_indent_in=0.12, sub_indent_in=0.22
            )

    enforce_tnr_13pt_all(doc_rev_ans)
    rev_ans_path = save_safe(doc_rev_ans, os.path.join(OUT_REV_ANS_DIR, f"Dap_an_De_cuong_Cuoi_ki_1_Tin_hoc_Lop_{lop}.docx"))
    return rev_path, rev_ans_path
