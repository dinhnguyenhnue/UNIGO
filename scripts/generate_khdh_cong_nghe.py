import os
import sys
import copy
import docx
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

sys.stdout.reconfigure(encoding='utf-8')

from cong_nghe_data import DATA_CONG_NGHE

DIR_MAU = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm'
DIR_PPCT = r'D:\UNIGO\KHBD_Cong_nghe\Phân phối chương trình'
os.makedirs(DIR_PPCT, exist_ok=True)

TEMPLATE_DOC = os.path.join(DIR_MAU, 'Kế hoạch dạy học môn Tin học 2026-2027.docx')

COL_WIDTHS_4 = [1.2, 5.0, 1.8, 8.5]  # Tổng 16.5 cm
COL_WIDTHS_5 = [3.5, 2.2, 2.2, 6.0, 2.6] # Tổng 16.5 cm

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        for child in list(tblPr[0]):
            if child.tag.endswith('tblBorders'):
                tblPr[0].remove(child)
        borders_xml = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr[0].append(borders_xml)

def set_font_all(doc):
    for p in doc.paragraphs:
        p.style.font.name = 'Times New Roman'
        p.style.font.size = Pt(13)
        for r in p.runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)
            
    for t in doc.tables:
        set_table_borders(t)
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.style.font.name = 'Times New Roman'
                    p.style.font.size = Pt(13)
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(13)

def set_cell(cell, width_cm, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False):
    cell.width = Cm(width_cm)
    tcPr = cell._element.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.bold = bold
        r.italic = italic

def build_4col_table(doc, lessons):
    """Xây dựng bảng PPCT 4 cột chuẩn UNIGO."""
    tbl = doc.add_table(rows=0, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)
    
    # Header
    hdr = tbl.add_row()
    headers = ['STT', 'Bài học (1)', 'Số tiết (2)', 'Yêu cầu cần đạt (3)']
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        set_cell(c, COL_WIDTHS_4[i], h, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
        # Nền xám nhẹ tiêu đề
        tcPr = c._element.get_or_add_tcPr()
        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="F2F2F2"/>' % nsdecls('w'))
        tcPr.append(shd)
        
    for item in lessons:
        stt, name, st, yccd, is_topic = item
        row = tbl.add_row()
        if is_topic:
            # Gộp cả 4 ô thành 1 hàng tiêu đề chủ đề
            c0 = row.cells[0]
            c1 = row.cells[1]
            c2 = row.cells[2]
            c3 = row.cells[3]
            c0.merge(c1).merge(c2).merge(c3)
            set_cell(c0, sum(COL_WIDTHS_4), name, WD_ALIGN_PARAGRAPH.LEFT, bold=True)
            tcPr = c0._element.get_or_add_tcPr()
            shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="E9EDF4"/>' % nsdecls('w'))
            tcPr.append(shd)
        else:
            set_cell(row.cells[0], COL_WIDTHS_4[0], str(stt), WD_ALIGN_PARAGRAPH.CENTER)
            is_exam = "Đánh giá định kỳ" in name or "Tổng kết" in name
            set_cell(row.cells[1], COL_WIDTHS_4[1], name, WD_ALIGN_PARAGRAPH.LEFT, bold=is_exam)
            set_cell(row.cells[2], COL_WIDTHS_4[2], str(st), WD_ALIGN_PARAGRAPH.CENTER)
            set_cell(row.cells[3], COL_WIDTHS_4[3], yccd, WD_ALIGN_PARAGRAPH.JUSTIFY)
    return tbl

def build_5col_table(doc, grade):
    """Xây dựng bảng kiểm tra đánh giá định kỳ 5 cột chuẩn UNIGO."""
    tbl = doc.add_table(rows=0, cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)
    
    # Header
    hdr = tbl.add_row()
    headers = ['Bài kiểm tra, đánh giá', 'Thời gian (1)', 'Thời điểm (2)', 'Yêu cầu cần đạt (3)', 'Hình thức (4)']
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        set_cell(c, COL_WIDTHS_5[i], h, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
        tcPr = c._element.get_or_add_tcPr()
        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="F2F2F2"/>' % nsdecls('w'))
        tcPr.append(shd)
        
    exams = [
        ("Đánh giá định kỳ 1", "35 phút" if int(grade) <= 5 else "45 phút", "Tuần 10", f"Đánh giá mức độ đạt yêu cầu cần đạt về kiến thức và kĩ năng môn Công nghệ Lớp {grade} từ Tuần 1 đến Tuần 9.", "Viết hoặc thực hành"),
        ("Đánh giá định kỳ 2", "35 phút" if int(grade) <= 5 else "45 phút", "Tuần 19", f"Đánh giá tổng hợp kiến thức, kĩ năng môn Công nghệ Lớp {grade} trong toàn bộ Học kì 1.", "Viết hoặc thực hành"),
        ("Đánh giá định kỳ 3", "35 phút" if int(grade) <= 5 else "45 phút", "Tuần 28", f"Đánh giá mức độ đạt yêu cầu cần đạt về kiến thức và kĩ năng môn Công nghệ Lớp {grade} từ Tuần 20 đến Tuần 27.", "Viết hoặc thực hành"),
        ("Đánh giá định kỳ 4", "35 phút" if int(grade) <= 5 else "45 phút", "Tuần 34", f"Đánh giá toàn diện kiến thức, kĩ năng thực hành môn Công nghệ Lớp {grade} trong toàn bộ năm học.", "Viết hoặc thực hành / Sản phẩm"),
    ]
    
    for ex in exams:
        row = tbl.add_row()
        set_cell(row.cells[0], COL_WIDTHS_5[0], ex[0], WD_ALIGN_PARAGRAPH.LEFT, bold=True)
        set_cell(row.cells[1], COL_WIDTHS_5[1], ex[1], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[2], COL_WIDTHS_5[2], ex[2], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[3], COL_WIDTHS_5[3], ex[3], WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell(row.cells[4], COL_WIDTHS_5[4], ex[4], WD_ALIGN_PARAGRAPH.CENTER)
    return tbl

def create_khdh_document(grades_list, output_path, subtitle=""):
    """Tạo văn bản KHDH dựa trên template chuẩn bảo tồn header logo UNIGO."""
    doc = docx.Document(TEMPLATE_DOC)
    
    # Giữ nguyên sectPr chứa header & logo
    # Xóa toàn bộ nội dung trong body trừ sectPr
    body = doc._element.body
    children_to_remove = []
    for child in list(body):
        if not child.tag.endswith('sectPr'):
            children_to_remove.append(child)
    for child in children_to_remove:
        body.remove(child)
        
    # 1. Bảng Quốc hiệu & Trường (Table 0)
    tbl0 = doc.add_table(rows=2, cols=2)
    tbl0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell(tbl0.rows[0].cells[0], 8.0, "TRƯỜNG: Tiểu học & THCS UNIGO\nTỔ: Tổ chuyên môn", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell(tbl0.rows[0].cells[1], 8.5, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell(tbl0.rows[1].cells[0], 8.0, "Số:      /KH-UNIGO", WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(tbl0.rows[1].cells[1], 8.5, "Hoài Đức, ngày 30 tháng 07 năm 2026", WD_ALIGN_PARAGRAPH.CENTER, italic=True)
    
    # 2. Tiêu đề
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(4)
    r = p_title.add_run("KẾ HOẠCH DẠY HỌC MÔN CÔNG NGHỆ\nNĂM HỌC 2026 - 2027")
    r.bold = True
    r.font.size = Pt(15)
    r.font.name = 'Times New Roman'
    
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(10)
        r_sub = p_sub.add_run(subtitle)
        r_sub.bold = True
        r_sub.font.size = Pt(13)
        r_sub.font.name = 'Times New Roman'
        
    # 3. Phần I: Tiêu chuẩn cơ sở vật chất
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(10)
    r_p1 = p_p1.add_run("I. Tiêu chuẩn cơ sở vật chất, thiết bị dạy học")
    r_p1.bold = True
    r_p1.font.size = Pt(13)
    
    p_tb = doc.add_paragraph()
    r_tb = p_tb.add_run("1. Thiết bị dạy học:")
    r_tb.bold = True
    
    tbl_tb = doc.add_table(rows=0, cols=4)
    tbl_tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tb)
    hdr_tb = tbl_tb.add_row()
    for idx, (txt, w) in enumerate([("STT", 1.2), ("Thiết bị dạy học", 6.5), ("Số lượng", 2.0), ("Các bài thực hành / Dự án", 6.8)]):
        set_cell(hdr_tb.cells[idx], w, txt, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    tb_rows = [
        ("1", "Dụng cụ trồng hoa, cây cảnh (chậu, xẻng, bình tưới)", "02 bộ / phòng", "Trồng, chăm sóc cây cảnh (Lớp 4)"),
        ("2", "Bộ lắp ghép mô hình kĩ thuật Tiểu học", "30 bộ", "Lắp ráp mô hình bập bênh, robot (Lớp 4, 5)"),
        ("3", "Mô hình năng lượng sạch (pin mặt trời, máy phát gió)", "15 bộ", "Mô hình năng lượng tái tạo (Lớp 5)"),
        ("4", "Bộ dụng cụ gia công cơ khí cầm tay, mạch điện cảm biến", "20 bộ", "Thực hành cơ khí & Kĩ thuật điện (Lớp 8)"),
    ]
    for r_data in tb_rows:
        row = tbl_tb.add_row()
        set_cell(row.cells[0], 1.2, r_data[0], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[1], 6.5, r_data[1], WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(row.cells[2], 2.0, r_data[2], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[3], 6.8, r_data[3], WD_ALIGN_PARAGRAPH.LEFT)
        
    p_phong = doc.add_paragraph()
    p_phong.paragraph_format.space_before = Pt(8)
    r_phong = p_phong.add_run("2. Phòng học bộ môn:")
    r_phong.bold = True
    
    tbl_phong = doc.add_table(rows=0, cols=4)
    tbl_phong.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_phong)
    hdr_ph = tbl_phong.add_row()
    for idx, (txt, w) in enumerate([("STT", 1.2), ("Tên phòng", 5.5), ("Số lượng", 2.0), ("Phạm vi và nội dung sử dụng", 7.8)]):
        set_cell(hdr_ph.cells[idx], w, txt, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    ph_rows = [
        ("1", "Phòng thực hành Công nghệ / STEM", "01 phòng", "Phục vụ các tiết thực hành lắp ráp mô hình, dự án sáng tạo kĩ thuật"),
        ("2", "Khuôn viên vườn trường / Sân thực hành sinh học", "01 khu vực", "Phục vụ thực hành gieo hạt, trồng và chăm sóc cây cảnh"),
    ]
    for r_data in ph_rows:
        row = tbl_phong.add_row()
        set_cell(row.cells[0], 1.2, r_data[0], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[1], 5.5, r_data[1], WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(row.cells[2], 2.0, r_data[2], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(row.cells[3], 7.8, r_data[3], WD_ALIGN_PARAGRAPH.LEFT)
        
    # 4. Phần II: Phân phối chương trình chi tiết từng khối lớp
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(14)
    r_p2 = p_p2.add_run("II. Kế hoạch dạy học chi tiết")
    r_p2.bold = True
    r_p2.font.size = Pt(13)
    
    for g in grades_list:
        data = DATA_CONG_NGHE[g]
        p_g = doc.add_paragraph()
        p_g.paragraph_format.space_before = Pt(10)
        p_g.paragraph_format.space_after = Pt(4)
        r_g = p_g.add_run(f"1. Phân phối chương trình môn Công nghệ - Lớp {g} (35 tiết/năm học)")
        r_g.bold = True
        r_g.font.size = Pt(13)
        
        # Bảng 4 cột bài học
        build_4col_table(doc, data['lessons'])
        
        # Bảng 5 cột kiểm tra đánh giá định kỳ
        p_kt = doc.add_paragraph()
        p_kt.paragraph_format.space_before = Pt(8)
        p_kt.paragraph_format.space_after = Pt(4)
        r_kt = p_kt.add_run(f"2. Kiểm tra, đánh giá định kỳ - Lớp {g}")
        r_kt.bold = True
        r_kt.italic = True
        
        build_5col_table(doc, g)
        
    # 5. Khung ký tên cuối trang
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(16)
    
    tbl_sig = doc.add_table(rows=1, cols=2)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Chữ ký không viền
    tblPr = tbl_sig._element.xpath('w:tblPr')
    if tblPr:
        borders_xml = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="none"/>'
            r'  <w:left w:val="none"/>'
            r'  <w:bottom w:val="none"/>'
            r'  <w:right w:val="none"/>'
            r'  <w:insideH w:val="none"/>'
            r'  <w:insideV w:val="none"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr[0].append(borders_xml)
        
    set_cell(tbl_sig.rows[0].cells[0], 8.0, "TỔ TRƯỞNG CHUYÊN MÔN\n(Ký, ghi rõ họ tên)\n\n\n\n", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell(tbl_sig.rows[0].cells[1], 8.5, "Hoài Đức, ngày 30 tháng 07 năm 2026\nHIỆU TRƯỞNG\n(Ký và ghi rõ họ tên)\n\n\n\n", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    # Định dạng font chuẩn Times New Roman 13pt
    set_font_all(doc)
    doc.save(output_path)
    print(f"✓ Đã tạo thành công: {output_path}")

# ============================================================
# TIẾN HÀNH TẠO BỘ TÀI LIỆU KHDH MÔN CÔNG NGHỆ ĐỒNG BỘ
# ============================================================

print("--- 1. Tạo file Tổng hợp toàn trường (Lớp 3 -> Lớp 8) ---")
file_all = os.path.join(DIR_MAU, 'Kế hoạch dạy học môn Công nghệ 2026-2027.docx')
create_khdh_document(['3', '4', '5', '6', '7', '8'], file_all, "(Áp dụng cho Khối 3, 4, 5, 6, 7, 8)")

print("--- 2. Tạo file Tiểu học (Lớp 3, 4, 5) ---")
file_th = os.path.join(DIR_MAU, 'Kế hoạch dạy học môn Công nghệ (TH) - 2026 - 2027.docx')
create_khdh_document(['3', '4', '5'], file_th, "CẤP TIỂU HỌC (LỚP 3, 4, 5)")

print("--- 3. Tạo file THCS (Lớp 6, 7, 8) ---")
file_thcs = os.path.join(DIR_MAU, 'Kế hoạch dạy học môn Công nghệ (THCS) - 2026 - 2027.docx')
create_khdh_document(['6', '7', '8'], file_thcs, "CẤP THCS (LỚP 6, 7, 8)")

print("--- 4. Tạo từng file riêng biệt cho từng khối lớp ---")
for g in ['3', '4', '5', '6', '7', '8']:
    file_lop = os.path.join(DIR_PPCT, f'Kế hoạch dạy học môn Công nghệ - Lớp {g} - 2026 - 2027.docx')
    create_khdh_document([g], file_lop, f"KHỐI LỚP {g}")

print("\n>>> TOÀN BỘ HỆ THỐNG KHDH MÔN CÔNG NGHỆ ĐÃ HOÀN TẤT ĐỒNG BỘ! <<<")
