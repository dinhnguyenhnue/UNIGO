# -*- coding: utf-8 -*-
"""
generate_khbd_cong_nghe_hki.py
==============================
Tự động tạo KHBD môn Công nghệ Học kì 1 (Tuần 1 -> Tuần 19) cho Lớp 3, 4, 5, 6, 7, 8
theo chuẩn UNIGO:
- Bảng 3x2 không viền (Table 0)
- Mục tiêu kiến thức dùng Danh từ / Cụm danh từ (không dùng động từ)
- Phân nhóm Năng lực (Năng lực công nghệ, Năng lực số TT02/2025, Năng lực chung) & Phẩm chất
- Tiến trình 4 Hoạt động (Bảng 2 cột có viền)
- Chèn ảnh minh họa từ SGK Công nghệ KNTT
- Rút kinh nghiệm & Bảng chữ ký 3x3 không viền
- Lưu theo cấu trúc Tuần đồng bộ Lịch báo giảng
"""

import os
import sys
import re
from datetime import date, timedelta
import docx
from docx import Document
from docx.shared import Pt, Cm, Emu, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

sys.stdout.reconfigure(encoding='utf-8')

from cong_nghe_data import DATA_CONG_NGHE

BASE_OUT = r"D:\UNIGO\KHBD_Cong_nghe\KHBD"
SGK_IMG_DIR = r"D:\UNIGO\KHBD_Cong_nghe\images_sgk"
TEMPLATE_DOCX = r"D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx"

TUAN_01_START = date(2026, 8, 3)

# EMU indents
INDENT_0 = 0
INDENT_1 = 180340
INDENT_2 = 360045
INDENT_BULLET = 540000

FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(13)
LINE_SPACING = 1.15
SPACE_AFTER = Emu(38100)

def set_font(run, bold=False, italic=False, size=FONT_SIZE):
    run.font.name = FONT_NAME
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

def add_p(doc, text="", bold=False, italic=False, first_indent=None, left_indent=None,
          align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=SPACE_AFTER, space_before=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.space_after = space_after
    pf.space_before = Emu(space_before) if isinstance(space_before, int) else space_before
    if first_indent is not None:
        pf.first_line_indent = Emu(first_indent)
    if left_indent is not None:
        pf.left_indent = Emu(left_indent)
    if text:
        run = p.add_run(text)
        set_font(run, bold=bold, italic=italic)
    return p

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        for child in list(tblPr[0]):
            if child.tag.endswith('tblBorders'):
                tblPr[0].remove(child)
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'</w:tblBorders>' % nsdecls("w")
        )
        tblPr[0].append(borders)

def set_table_no_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        for child in list(tblPr[0]):
            if child.tag.endswith('tblBorders'):
                tblPr[0].remove(child)
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="none"/>'
            r'  <w:left w:val="none"/>'
            r'  <w:bottom w:val="none"/>'
            r'  <w:right w:val="none"/>'
            r'  <w:insideH w:val="none"/>'
            r'  <w:insideV w:val="none"/>'
            r'</w:tblBorders>' % nsdecls("w")
        )
        tblPr[0].append(borders)

def set_cell_box(cell, width_cm):
    cell.width = Cm(width_cm)
    tcPr = cell._element.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def build_khbd_file(grade, week, ppct_num, lesson_title, yccd_text, ngay_day_str, ngay_soan_str, lop_str):
    doc = Document()
    is_thcs = int(grade) >= 6
    
    # 1. BẢNG THÔNG TIN ĐẦU BÀI (Table 0: 3 hàng x 2 cột, NO BORDER)
    t0 = doc.add_table(rows=3, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t0)
    set_cell_box(t0.rows[0].cells[0], 8.5)
    set_cell_box(t0.rows[0].cells[1], 8.0)
    set_cell_box(t0.rows[1].cells[0], 8.5)
    set_cell_box(t0.rows[1].cells[1], 8.0)
    set_cell_box(t0.rows[2].cells[0], 8.5)
    set_cell_box(t0.rows[2].cells[1], 8.0)
    
    # Row 0
    p = t0.rows[0].cells[0].paragraphs[0]
    set_font(p.add_run("Trường: Tiểu học và THCS UNIGO"), bold=True)
    p = t0.rows[0].cells[1].paragraphs[0]
    set_font(p.add_run(f"Ngày soạn: {ngay_soan_str}"), bold=True)
    
    # Row 1
    p = t0.rows[1].cells[0].paragraphs[0]
    set_font(p.add_run("GV: Đậu Đình Nguyên"))
    p = t0.rows[1].cells[1].paragraphs[0]
    set_font(p.add_run(f"Ngày dạy: {ngay_day_str}"))
    
    # Row 2
    to_chuyen_mon = "Tổ chuyên môn THCS" if is_thcs else "Tổ chuyên môn Tiểu học"
    p = t0.rows[2].cells[0].paragraphs[0]
    set_font(p.add_run(f"Tổ: {to_chuyen_mon}"))
    p = t0.rows[2].cells[1].paragraphs[0]
    set_font(p.add_run(f"Lớp: {lop_str}"))
    
    # 2. TIÊU ĐỀ BÀI DẠY
    if is_thcs:
        p1 = add_p(doc, f"TÊN BÀI DẠY: {lesson_title.upper()}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=2)
        p2 = add_p(doc, "Môn học: Công nghệ", bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p3 = add_p(doc, "Thời lượng: 1 tiết (45 phút)", bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p4 = add_p(doc, f"Tiết theo PPCT: {ppct_num}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    else:
        p1 = add_p(doc, "KẾ HOẠCH DẠY HỌC MÔN CÔNG NGHỆ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=4)
        p2 = add_p(doc, f"{lesson_title.upper()} (Tiết: {ppct_num} theo PPCT)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
        
    # 3. MỤC I. YÊU CẦU CẦN ĐẠT
    add_p(doc, "I. YÊU CẦU CẦN ĐẠT", bold=True, first_indent=INDENT_0, space_after=4)
    
    # 1. Kiến thức (Danh từ / cụm danh từ trực tiếp, không dùng động từ)
    add_p(doc, "1. Kiến thức", bold=True, first_indent=INDENT_1, space_after=2)
    # Tách các ý từ YCCD
    yccd_items = [item.strip() for item in yccd_text.split(';') if item.strip()]
    for item in yccd_items:
        # Chuẩn hóa thành cụm danh từ
        item_noun = item
        for bad_verb in ["Nêu được ", "Trình bày được ", "Nhận biết được ", "Phân biệt được ", "Hiểu được ", "Kể tên được ", "Mô tả được ", "Biết cách "]:
            if item_noun.startswith(bad_verb):
                item_noun = item_noun[len(bad_verb):]
                item_noun = item_noun[0].upper() + item_noun[1:]
        add_p(doc, f"- {item_noun}.", left_indent=INDENT_BULLET, space_after=2)
        
    # 2. Năng lực
    add_p(doc, "2. Năng lực", bold=True, first_indent=INDENT_1, space_after=2)
    add_p(doc, "2.1. Năng lực đặc thù (Công nghệ)", bold=True, first_indent=INDENT_2, space_after=2)
    add_p(doc, "- Năng lực nhận thức công nghệ: Nhận thức các khái niệm, quy trình và thuật ngữ kĩ thuật của bài học. (Đạt được thông qua Hoạt động 1, Hoạt động 2)", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Năng lực giao tiếp công nghệ: Sử dụng thuật ngữ, sơ đồ, bản vẽ và ngôn ngữ kĩ thuật để trao đổi ý tưởng. (Đạt được thông qua Hoạt động 2, Hoạt động 3)", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Năng lực sử dụng công nghệ: Thao tác an toàn, đúng quy trình kĩ thuật trên dụng cụ và vật liệu thực hành. (Đạt được thông qua Hoạt động 3, Hoạt động 4)", left_indent=INDENT_BULLET, space_after=2)
    
    add_p(doc, "2.2. Năng lực số (Thông tư 02/2025 - CV 3456)", bold=True, first_indent=INDENT_2, space_after=2)
    b_val = "Bậc 1" if int(grade) <= 4 else ("Bậc 2" if int(grade) <= 6 else "Bậc 3")
    add_p(doc, f"- Miền I. Khai thác dữ liệu và thông tin số: Khai thác hình ảnh, video minh họa công nghệ trực quan qua bài học số ({b_val}). (Đạt được thông qua Hoạt động 2)", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, f"- Miền III. Sáng tạo nội dung số: Phác thảo, ghi chép và chia sẻ ý tưởng thiết kế mô hình ({b_val}). (Đạt được thông qua Hoạt động 3)", left_indent=INDENT_BULLET, space_after=2)
    
    add_p(doc, "2.3. Năng lực chung", bold=True, first_indent=INDENT_2, space_after=2)
    add_p(doc, "- Năng lực tự chủ và tự học: Tự chuẩn bị học liệu, chủ động quan sát và hoàn thành nhiệm vụ được giao. (Đạt được thông qua Hoạt động 1, Hoạt động 2)", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Năng lực giao tiếp và hợp tác: Tương tác tích cực, làm việc nhóm hiệu quả khi thảo luận và thực hành kĩ thuật. (Đạt được thông qua Hoạt động 3)", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Năng lực giải quyết vấn đề và sáng tạo: Đề xuất các giải pháp kĩ thuật phù hợp, chế tạo sản phẩm sáng tạo. (Đạt được thông qua Hoạt động 4)", left_indent=INDENT_BULLET, space_after=2)
    
    # 3. Phẩm chất
    add_p(doc, "3. Phẩm chất", bold=True, first_indent=INDENT_1, space_after=2)
    add_p(doc, "- Chăm chỉ: Miệt mài thực hành, kiên trì khắc phục khó khăn để hoàn thiện sản phẩm kĩ thuật.", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Trách nhiệm: Giữ gìn an toàn lao động, sử dụng tiết kiệm vật liệu và dọn dẹp vệ sinh khu vực học tập.", left_indent=INDENT_BULLET, space_after=4)
    
    # 4. MỤC II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
    add_p(doc, "II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU", bold=True, first_indent=INDENT_0, space_after=4)
    add_p(doc, "- Giáo viên: Sách giáo khoa Công nghệ, bài giảng điện tử (slide), hình ảnh minh họa trích xuất từ SGK, mẫu vật hoặc mô hình kĩ thuật thực tế, phiếu học tập nhóm.", left_indent=INDENT_BULLET, space_after=2)
    add_p(doc, "- Học sinh: Sách giáo khoa Công nghệ, vở ghi, đồ dùng học tập, dụng cụ và vật liệu thực hành theo hướng dẫn của giáo viên.", left_indent=INDENT_BULLET, space_after=6)
    
    # 5. MỤC III. TIẾN TRÌNH DẠY HỌC (4 Hoạt động - Bảng 2 cột CÓ VIỀN)
    add_p(doc, "III. TIẾN TRÌNH DẠY HỌC", bold=True, first_indent=INDENT_0, space_after=4)
    
    # Kiểm tra xem có ảnh SGK phù hợp không
    img_path = None
    g_img_dir = os.path.join(SGK_IMG_DIR, f"Lop_{grade}")
    if os.path.exists(g_img_dir):
        # Chọn trang ảnh tương ứng với tiết/bài
        candidate_img = os.path.join(g_img_dir, f"page_{min(15, ppct_num + 3):02d}.png")
        if os.path.exists(candidate_img):
            img_path = candidate_img
            
    activities = [
        ("Hoạt động 1. Khởi động (5 phút)",
         "Tạo hứng thú, kích thích sự tò mò và dẫn dắt học sinh vào nội dung bài học mới.",
         "Quan sát tranh tình huống hoặc trả lời câu hỏi khởi động của giáo viên.",
         "Câu trả lời và sự tương tác ban đầu của học sinh về vấn đề bài học.",
         "GV trình chiếu hình ảnh/video tình huống dẫn dắt. HS lắng nghe, thảo luận cặp đôi và đưa ra suy đoán ban đầu. GV nhận xét và kết nối vào bài mới.",
         "Học sinh định hướng được vấn đề trọng tâm cần tìm hiểu trong bài học."),
         
        ("Hoạt động 2. Hình thành kiến thức mới (18 phút)",
         f"Nắm vững nội dung cốt lõi: {yccd_items[0] if yccd_items else lesson_title}.",
         "Khám phá kiến thức qua kênh hình, kênh chữ trong SGK và trao đổi nhóm.",
         "Kết quả thảo luận nhóm, ghi nhận kiến thức trọng tâm vào vở ghi.",
         f"Bước 1. Chuyển giao nhiệm vụ: GV yêu cầu HS quan sát kênh hình SGK, thảo luận nhóm 4 trả lời câu hỏi trọng tâm.\nBước 2. Thực hiện nhiệm vụ: HS quan sát hình ảnh minh họa, đối chiếu kiến thức và ghi phiếu học tập.\nBước 3. Báo cáo: Đại diện nhóm trình bày kết quả.\nBước 4. Kết luận: GV nhận xét, chuẩn hóa kiến thức.",
         f"Kiến thức trọng tâm:\n- {yccd_text}\n- Học sinh nắm chắc bản chất kĩ thuật và quy tắc an toàn."),
         
        ("Hoạt động 3. Luyện tập (12 phút)",
         "Củng cố kiến thức đã học, rèn luyện kĩ năng nhận diện và thực hành thao tác cơ bản.",
         "Thực hiện các câu hỏi trắc nghiệm, bài tập tình huống hoặc thao tác thực hành mẫu.",
         "Đáp án bài tập của học sinh hoặc sản phẩm thực hành giai đoạn 1.",
         "Bước 1. GV giao bài tập tình huống / phiếu bài tập.\nBước 2. HS làm việc cá nhân kết hợp thảo luận cặp đôi.\nBước 3. GV gọi học sinh lên bảng trình bày hoặc thực hiện thao tác mẫu.\nBước 4. Đánh giá, chữa bài chi tiết.",
         "Học sinh vận dụng đúng kiến thức để giải quyết trọn vẹn các bài tập luyện tập."),
         
        ("Hoạt động 4. Vận dụng (5 phút)",
         "Vận dụng kiến thức bài học vào thực tiễn gia đình và môi trường xung quanh.",
         "Liên hệ thực tế: Tìm hiểu sản phẩm công nghệ tương tự tại gia đình hoặc đề xuất ý tưởng cải tiến.",
         "Kế hoạch thực hành tại nhà hoặc bản phác thảo ý tưởng sáng tạo.",
         "Bước 1. GV hướng dẫn nhiệm vụ vận dụng thực tế tại nhà.\nBước 2. HS lắng nghe, ghi chép yêu cầu.\nBước 3. GV nhắc nhở an toàn và thời hạn nộp sản phẩm ở tiết học sau.",
         "Học sinh có ý thức giữ gìn, bảo quản và vận dụng sản phẩm công nghệ an toàn, tiết kiệm."),
    ]
    
    for act_idx, (act_title, act_mt, act_nd, act_sp, act_gv_hs, act_kq) in enumerate(activities):
        add_p(doc, act_title, bold=True, first_indent=INDENT_1, space_before=4, space_after=2)
        add_p(doc, f"a) Mục tiêu: {act_mt}", left_indent=INDENT_2, space_after=2)
        add_p(doc, f"b) Nội dung: {act_nd}", left_indent=INDENT_2, space_after=2)
        add_p(doc, f"c) Sản phẩm: {act_sp}", left_indent=INDENT_2, space_after=2)
        add_p(doc, "d) Tổ chức thực hiện:", left_indent=INDENT_2, space_after=3)
        
        # Bảng 2 cột CÓ VIỀN
        tbl_act = doc.add_table(rows=2, cols=2)
        tbl_act.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl_act)
        set_cell_box(tbl_act.rows[0].cells[0], 9.5)
        set_cell_box(tbl_act.rows[0].cells[1], 7.0)
        set_cell_box(tbl_act.rows[1].cells[0], 9.5)
        set_cell_box(tbl_act.rows[1].cells[1], 7.0)
        
        # Header bảng
        c0 = tbl_act.rows[0].cells[0]
        c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(c0.paragraphs[0].add_run("HOẠT ĐỘNG CỦA GV – HS"), bold=True)
        tcPr = c0._element.get_or_add_tcPr()
        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="F2F2F2"/>' % nsdecls('w'))
        tcPr.append(shd)
        
        c1 = tbl_act.rows[0].cells[1]
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(c1.paragraphs[0].add_run("KẾT QUẢ CẦN ĐẠT"), bold=True)
        tcPr = c1._element.get_or_add_tcPr()
        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="F2F2F2"/>' % nsdecls('w'))
        tcPr.append(shd)
        
        # Body bảng
        cell_left = tbl_act.rows[1].cells[0]
        cell_right = tbl_act.rows[1].cells[1]
        
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for line in act_gv_hs.split('\n'):
            if line.startswith("Bước"):
                parts = line.split(':', 1)
                r_step = p_left.add_run(parts[0] + ":")
                set_font(r_step, bold=True, italic=True)
                if len(parts) > 1:
                    r_desc = p_left.add_run(parts[1] + "\n")
                    set_font(r_desc)
            else:
                r_line = p_left.add_run(line + "\n")
                set_font(r_line)
                
        # Nếu là Hoạt động 2 và có ảnh SGK, chèn ảnh vào cột trái
        if act_idx == 1 and img_path and os.path.exists(img_path):
            p_img = cell_left.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(2)
            try:
                run_img = p_img.add_run()
                run_img.add_picture(img_path, width=Inches(3.2))
                p_cap = cell_left.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_after = Pt(4)
                r_cap = p_cap.add_run("Hình minh họa bài học (Nguồn: SGK Kết nối tri thức)")
                set_font(r_cap, italic=True, size=Pt(10.5))
            except Exception as e:
                pass
                
        p_right = cell_right.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for line in act_kq.split('\n'):
            r_kq = p_right.add_run(line + "\n")
            set_font(r_kq)
            
    # 6. RÚT KINH NGHIỆM SAU BÀI DẠY
    add_p(doc, "RÚT KINH NGHIỆM SAU BÀI DẠY:", bold=True, first_indent=INDENT_0, space_before=8, space_after=2)
    add_p(doc, ".......................................................................................................................................................................", first_indent=INDENT_0, space_after=2)
    add_p(doc, ".......................................................................................................................................................................", first_indent=INDENT_0, space_after=8)
    
    # 7. BẢNG CHỮ KÝ 3x3 NO BORDER
    t_sig = doc.add_table(rows=3, cols=3)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t_sig)
    for row in t_sig.rows:
        set_cell_box(row.cells[0], 5.5)
        set_cell_box(row.cells[1], 5.5)
        set_cell_box(row.cells[2], 5.5)
        
    # Row 0
    p = t_sig.rows[0].cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("DUYỆT CỦA BGH"), bold=True)
    p = t_sig.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("DUYỆT CỦA TỔ CM"), bold=True)
    p = t_sig.rows[0].cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("NGƯỜI SOẠN"), bold=True)
    
    # Row 1
    p = t_sig.rows[1].cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("(Ký, ghi rõ họ tên)"), italic=True)
    p = t_sig.rows[1].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("(Ký, ghi rõ họ tên)"), italic=True)
    p = t_sig.rows[1].cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("(Ký, ghi rõ họ tên)"), italic=True)
    
    # Row 2
    p = t_sig.rows[2].cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(36)
    set_font(p.add_run("Đậu Đình Nguyên"), bold=True)
    
    return doc

# ============================================================
# BATCH GENERATION FOR SEMESTER 1 (WEEKS 1 - 19)
# ============================================================

def get_day_for_grade(grade, week_start):
    # Lớp 4C1 dạy Thứ 2
    if grade == '4':
        d = week_start
        lop = '4C1'
    # Lớp 5C1, 7A1 dạy Thứ 3
    elif grade in ('5', '7'):
        d = week_start + timedelta(days=1)
        lop = '5C1' if grade == '5' else '7A1'
    # Lớp 3A1 dạy Thứ 4, 3C1 dạy Thứ 6 -> Lấy Thứ 4 (3A1, 3C1)
    elif grade == '3':
        d = week_start + timedelta(days=2)
        lop = '3A1, 3C1'
    # Lớp 6A1, 8A1 dạy Thứ 6
    elif grade in ('6', '8'):
        d = week_start + timedelta(days=4)
        lop = '6A1' if grade == '6' else '8A1'
    else:
        d = week_start
        lop = f'{grade}A1'
    ngay_soan = week_start - timedelta(days=2) # Thứ 7 tuần trước
    return d.strftime('%d/%m/%Y'), ngay_soan.strftime('%d/%m/%Y'), lop

print("Starting batch generation for Semester 1 (Weeks 1 - 19)...")

total_created = 0

for g in ['3', '4', '5', '6', '7', '8']:
    lessons = DATA_CONG_NGHE[g]['lessons']
    # Build lesson map: stt -> (name, yccd)
    l_map = {}
    l_map[0] = ("Tiết 0: Định hướng môn học", "Tìm hiểu mục tiêu, yêu cầu và phương pháp học tập môn Công nghệ.")
    for item in lessons:
        stt, name, st, yccd, is_topic = item
        if not is_topic and stt is not None:
            l_map[stt] = (name, yccd)
            
    is_rot = g in ('5', '6', '7', '8')
    
    # Weeks 1 to 19
    for w in range(1, 20):
        w_start = TUAN_01_START + timedelta(weeks=w-1)
        ngay_day, ngay_soan, lop_str = get_day_for_grade(g, w_start)
        w_dir = os.path.join(BASE_OUT, f"Lớp_{g}", f"Tuần_{w:02d}")
        os.makedirs(w_dir, exist_ok=True)
        
        if is_rot:
            # Tuần 1: Tiết 0
            if w == 1:
                ppct_list = [0]
            elif w % 2 == 1:
                # Tuần lẻ: 2 tiết liên tiếp
                k = (w - 3) // 2 + 1
                p1 = (k - 1) * 2 + 1
                p2 = p1 + 1
                ppct_list = [p1, p2]
            else:
                # Tuần chẵn là tuần dạy môn Tin học (không có tiết Công nghệ)
                ppct_list = []
        else:
            # Lớp 3, 4: mỗi tuần 1 tiết
            if w == 1:
                ppct_list = [0]
            else:
                ppct_list = [w - 1]
                
        for ppct_val in ppct_list:
            if ppct_val in l_map:
                l_title, l_yccd = l_map[ppct_val]
            else:
                l_title = f"Bài học Tiết {ppct_val}"
                l_yccd = f"Nắm vững kiến thức và kĩ năng thực hành môn Công nghệ Lớp {g}."
                
            doc = build_khbd_file(
                grade=g,
                week=w,
                ppct_num=ppct_val,
                lesson_title=l_title,
                yccd_text=l_yccd,
                ngay_day_str=ngay_day,
                ngay_soan_str=ngay_soan,
                lop_str=lop_str
            )
            # Tên file chuẩn UNIGO: KHBD_Cong_nghe_Lop_[X]_Tiet[YY]_[Ten_bai].docx
            clean_title = re.sub(r'[\\/*?:"<>|]', '_', l_title).replace(' ', '_')[:40]
            out_filename = f"KHBD_Cong_nghe_Lop_{g}_Tiet{ppct_val:02d}_{clean_title}.docx"
            out_filepath = os.path.join(w_dir, out_filename)
            doc.save(out_filepath)
            total_created += 1

print(f"\n>>> HOÀN TẤT! Đã tạo thành công {total_created} file KHBD môn Công nghệ Học kì 1! <<<")

