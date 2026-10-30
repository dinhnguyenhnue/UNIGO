# -*- coding: utf-8 -*-
"""
Script soạn Biên bản bàn giao công việc dạy CLB Robotics UNIGO
- Người bàn giao: Thầy Đậu Đình Nguyên (GV Tin học - Robotics UNIGO)
- Người nhận bàn giao: Cô Đào Hồng Anh (Sinh năm 2001, Chuyên viên Robohub)
- Thiết bị bàn giao: 15 bộ Initiate, 5 bộ Kinder, 5 bộ Excel
- Bàn giao công việc giảng dạy, hồ sơ KHBD, lịch trình và cam kết trách nhiệm theo form UNIGO
- Bảo tồn Logo UNIGO ở Header và Footer từ template chuẩn.
"""

import sys, os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_PATH = r"D:\UNIGO\CLB_Robotics\Kế hoạch dạy học CLB Robotics.docx"
OUT_PATH_CLB = r"D:\UNIGO\CLB_Robotics\Biên bản bàn giao công việc dạy CLB Robotics.docx"
OUT_PATH_MAU = r"D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Biên bản bàn giao công việc dạy CLB Robotics.docx"

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generate_bien_ban():
    print("Reading template:", TEMPLATE_PATH)
    doc = docx.Document(TEMPLATE_PATH)

    # Dọn dẹp body cũ, giữ nguyên sectPr chứa Header w:drawing logo UNIGO và footer
    body = doc.element.body
    children_to_remove = []
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag != 'sectPr':
            children_to_remove.append(child)
    for c in children_to_remove:
        body.remove(c)

    # 1. Header Table (Quốc hiệu & Tên trường)
    tbl_hdr = doc.add_table(rows=2, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl_hdr)
    w_hdr = [Inches(3.3), Inches(3.7)]

    c00 = tbl_hdr.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p00.add_run("TRƯỜNG TIỂU HỌC & THCS UNIGO\nCÂU LẠC BỘ ROBOTICS UNIGO")
    r.bold = True
    r.font.size = Pt(12)

    c01 = tbl_hdr.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc")
    r.bold = True
    r.font.size = Pt(12)

    c10 = tbl_hdr.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p10.add_run("--------------------")
    r.font.size = Pt(11)

    c11 = tbl_hdr.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p11.add_run("--------------------")
    r.font.size = Pt(11)

    for r_item in tbl_hdr.rows:
        for idx, cell_item in enumerate(r_item.cells):
            cell_item.width = w_hdr[idx]
            set_cell_margins(cell_item, top=30, bottom=30, left=40, right=40)

    # 2. Ngày tháng và Tiêu đề Biên bản
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.paragraph_format.space_before = Pt(6)
    p_date.paragraph_format.space_after = Pt(8)
    r = p_date.add_run("Hà Nội, ngày 22 tháng 10 năm 2026")
    r.italic = True
    r.font.size = Pt(12)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r = p_title.add_run("BIÊN BẢN BÀN GIAO CÔNG VIỆC VÀ THIẾT BỊ GIẢNG DẠY")
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(0, 32, 96)

    p_subtitle = doc.add_paragraph()
    p_subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subtitle.paragraph_format.space_after = Pt(10)
    r = p_subtitle.add_run("CÂU LẠC BỘ ROBOTICS UNIGO – NĂM HỌC 2026 - 2027")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(192, 0, 0)

    # 3. Căn cứ thực hiện
    doc.add_paragraph("Căn cứ Kế hoạch giáo dục và Kế hoạch phát triển các Câu lạc bộ năng khiếu năm học 2026 - 2027 của Trường Tiểu học và THCS UNIGO;")
    doc.add_paragraph("Căn cứ Thỏa thuận hợp tác liên kết đào tạo chương trình giáo dục STEM & Robotics giữa Trường Tiểu học & THCS UNIGO và Trung tâm Robohub;")
    doc.add_paragraph("Căn cứ Kế hoạch dạy học Câu lạc bộ Robotics năm học 2026 - 2027 đã được Ban Giám hiệu nhà trường phê duyệt;")
    doc.add_paragraph("Căn cứ nhu cầu thực tế về công tác phân công giáo viên giảng dạy và quản lý cơ sở vật chất phòng bộ môn Robotics.")

    # 4. Thời gian, địa điểm, thành phần
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    r = p.add_run("Hôm nay, vào hồi 09 giờ 00 phút, ngày 22 tháng 10 năm 2026, tại Phòng chức năng Robotics – Trường Tiểu học và THCS UNIGO, chúng tôi gồm có:")

    # 1. BÊN BÀN GIAO
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run("1. BÊN BÀN GIAO:")
    r.bold = True
    doc.add_paragraph("   - Họ và tên: Đậu Đình Nguyên                           Năm sinh: 2004")
    doc.add_paragraph("   - Chức vụ: Giáo viên Tin học & Robotics – Chủ nhiệm CLB Robotics UNIGO")
    doc.add_paragraph("   - Đơn vị công tác: Trường Tiểu học và THCS UNIGO")
    doc.add_paragraph("   - Số điện thoại: 0932.340.752                          Email: nguyendd@unigo.edu.vn")

    # 2. BÊN NHẬN BÀN GIAO
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run("2. BÊN NHẬN BÀN GIAO:")
    r.bold = True
    doc.add_paragraph("   - Họ và tên: Đào Hồng Anh                              Năm sinh: 2001")
    doc.add_paragraph("   - Chức danh: Giáo viên / Chuyên viên giảng dạy Robotics")
    doc.add_paragraph("   - Đơn vị công tác: Trung tâm Robohub (Đơn vị đối tác liên kết đào tạo Robotics)")
    doc.add_paragraph("   - Số điện thoại: ............................................ Email: ...........................................................")

    # 3. ĐẠI DIỆN BAN GIÁM HIỆU
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run("3. ĐẠI DIỆN BAN GIÁM HIỆU TRƯỜNG UNIGO:")
    r.bold = True
    doc.add_paragraph("   - Đại diện Ban Giám hiệu: ................................................................................................................")
    doc.add_paragraph("   - Chức vụ: ..........................................................................................................................................")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("Các bên đã tiến hành kiểm kê thực tế, đối chiếu và thống nhất bàn giao toàn bộ thiết bị dạy học, hồ sơ chuyên môn và nhiệm vụ giảng dạy Câu lạc bộ Robotics với các nội dung chi tiết như sau:")

    # ---------------- PHẦN I: THIẾT BỊ ----------------
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("I. NỘI DUNG BÀN GIAO THIẾT BỊ, GIÁO CỤ VÀ CƠ SỞ VẬT CHẤT")
    r.bold = True
    r.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph("Bên A tiến hành bàn giao đầy đủ số lượng thiết bị, bộ linh kiện Robotics theo đúng kiểm kê thực tế cho Bên B quản lý và phục vụ giảng dạy:")

    tbl_tb = doc.add_table(rows=1, cols=6)
    tbl_tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tb)
    w_tb = [Inches(0.6), Inches(2.2), Inches(0.8), Inches(0.7), Inches(1.8), Inches(1.0)]

    hdrs_tb = ["STT", "Tên thiết bị / Giáo cụ", "Số lượng", "ĐVT", "Tình trạng kỹ thuật khi bàn giao", "Ghi chú"]
    for i, h in enumerate(hdrs_tb):
        cell = tbl_tb.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "D9E1F2")
        set_cell_margins(cell)
        cell.width = w_tb[i]

    DATA_TB = [
        ("01", "Bộ Kit Robotics OLLO Initiate", "15", "Bộ",
         "Linh kiện đầy đủ theo danh mục chuẩn, chốt ghép an toàn, hộp đựng và khay phân loại nguyên vẹn; hoạt động tốt 100%.",
         "Phục vụ Khối 1 - 2 (Học kỳ 1)"),
        ("02", "Bộ Kit Robotics OLLO Kinder", "05", "Bộ",
         "Linh kiện cơ cấu máy đầy đủ, động cơ điện Kinder và hộp pin hoạt động ổn định, bánh răng và trục xoay trơn tru.",
         "Phục vụ Khối 3 - 7 (HK1) & Khối 1 - 2 (HK2)"),
        ("03", "Bộ Kit Robotics OLLO Excel (Level 1)", "05", "Bộ",
         "Linh kiện công nghệ cao đầy đủ: Động cơ servo thông minh Dynamixel, cảm biến âm thanh, cảm biến hồng ngoại, bộ điều khiển từ xa RC-100; thiết bị mới, hoạt động hoàn hảo 100%.",
         "Phục vụ Khối 3 - 7 (Học kỳ 2)"),
        ("04", "Thảm sa bàn thử nghiệm Robot UNIGO", "02", "Chiếc",
         "Thảm bạt sa bàn tiêu chuẩn thi đấu sạch sẽ, phẳng phiu, vạch kẻ đường line và các ô nhiệm vụ rõ ràng.",
         "Để tại sàn thực hành phòng Robotics"),
        ("05", "Phụ kiện hỗ trợ: Pin sạc, bộ sạc, kẹp tháo chốt", "01", "Hộp",
         "Gồm 30 kẹp tháo chốt chuyên dụng OLLO, hệ thống dây sạc và pin dự phòng đã được kiểm tra an toàn điện.",
         "Bảo quản trong tủ đồ dùng phòng chức năng")
    ]

    for stt, ten, sl, dvt, tt, gc in DATA_TB:
        row = tbl_tb.add_row()
        for idx, val in enumerate([stt, ten, sl, dvt, tt, gc]):
            cell = row.cells[idx]
            cell.text = val
            set_cell_margins(cell)
            cell.width = w_tb[idx]
            if idx in [0, 2, 3]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(4)
    r = p_sum.add_run("=> Tổng cộng thiết bị chính: ")
    r.bold = True
    p_sum.add_run("25 bộ Kit Robotics (15 bộ OLLO Initiate + 05 bộ OLLO Kinder + 05 bộ OLLO Excel) cùng toàn bộ phụ kiện kèm theo đang trong trạng thái kỹ thuật tốt, sẵn sàng cho công tác giảng dạy.")

    # ---------------- PHẦN II: CÔNG VIỆC VÀ HỒ SƠ CHUYÊN MÔN ----------------
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("II. NỘI DUNG BÀN GIAO HỒ SƠ CHUYÊN MÔN VÀ CÔNG VIỆC GIẢNG DẠY")
    r.bold = True
    r.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph("Bên A bàn giao toàn bộ hồ sơ chuyên môn, lịch trình và nhiệm vụ giảng dạy cho Bên B:")
    
    p = doc.add_paragraph()
    r = p.add_run("1. Hồ sơ Kế hoạch dạy học CLB Robotics năm học 2026 - 2027:")
    r.bold = True
    p.add_run(" Khung phân phối chương trình chi tiết 35 buổi/năm học được thiết kế phân hóa theo lứa tuổi:")
    doc.add_paragraph("   - Phân hệ 1 (Khối 1 - 2): Học kỳ 1 học OLLO Initiate (17 buổi) -> Sau khi hoàn thành, Học kỳ 2 chuyển tiếp lên học OLLO Kinder (18 buổi).")
    doc.add_paragraph("   - Phân hệ 2 (Khối 3 - 7): Học kỳ 1 học OLLO Kinder (17 buổi) -> Sau khi hoàn thành, Học kỳ 2 bứt phá công nghệ cao cùng OLLO Excel Level 1 (18 buổi).")

    p = doc.add_paragraph()
    r = p.add_run("2. Hệ thống Kế hoạch bài dạy (KHBD) chuẩn UNIGO:")
    r.bold = True
    p.add_run(" Đầy đủ trọn bộ ")
    r = p.add_run("70 file giáo án chi tiết (.docx)")
    r.bold = True
    p.add_run(" (35 buổi cho Khối 1-2 và 35 buổi cho Khối 3-7) được số hóa và sắp xếp khoa học theo từng tuần tại thư mục ")
    r = p.add_run("D:\\UNIGO\\CLB_Robotics\\KHBD_CLB_Robotics")
    r.italic = True
    p.add_run(". Toàn bộ KHBD tuân thủ đúng Công văn 5512/BGDĐT, Năng lực số Thông tư 02/2025 – CV 3456/BGDĐT và bảng tiến trình 4 hoạt động chuẩn UNIGO.")

    p = doc.add_paragraph()
    r = p.add_run("3. Lịch trình và thời gian biểu sinh hoạt:")
    r.bold = True
    doc.add_paragraph("   - Thời gian sinh hoạt: Sáng Thứ 7 hàng tuần, từ 09h45 đến 11h15 (Thời lượng: 90 phút/buổi).")
    doc.add_paragraph("   - Địa điểm: Phòng chức năng Robotics, Trường Tiểu học và THCS UNIGO.")
    doc.add_paragraph("   - Thời gian thực hiện: Bắt đầu từ ngày 12/09/2026 đến ngày 29/05/2027.")

    p = doc.add_paragraph()
    r = p.add_run("4. Kế hoạch Kiểm tra, Đánh giá định kỳ và Báo cáo thành tích:")
    r.bold = True
    doc.add_paragraph("   - Buổi 9 (07/11/2026): Đánh giá định kỳ 1 (Thao tác lắp ráp nhanh và nhận diện linh kiện).")
    doc.add_paragraph("   - Buổi 17 (02/01/2027): Đánh giá định kỳ 2 & Triển lãm Robot sáng tạo Học kì 1.")
    doc.add_paragraph("   - Buổi 28 (03/04/2027): Đánh giá định kỳ 3 (Lắp ráp nâng cao Kinder và tích hợp cảm biến Excel).")
    doc.add_paragraph("   - Buổi 34 (22/05/2027): Đánh giá định kỳ 4 (Bảo vệ dự án mô hình sáng tạo cuối năm).")
    doc.add_paragraph("   - Buổi 35 (29/05/2027): Ngày hội Robotics Festival UNIGO & Bế mạc CLB toàn trường.")

    p = doc.add_paragraph()
    r = p.add_run("5. Danh sách học sinh và Sổ theo dõi nền nếp:")
    r.bold = True
    p.add_run(" Danh sách phân chia học sinh theo 2 nhóm lớp, sổ điểm danh chuyên cần và phiếu đánh giá năng lực học sinh từng buổi.")

    # ---------------- PHẦN III: CAM KẾT TRÁCH NHIỆM ----------------
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("III. CAM KẾT TRÁCH NHIỆM CỦA BÊN NHẬN BÀN GIAO (BÊN B - CÔ ĐÀO HỒNG ANH)")
    r.bold = True
    r.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph("Sau khi kiểm tra thực tế, tiếp nhận đầy đủ thiết bị, hồ sơ chuyên môn và nhiệm vụ giảng dạy, cô Đào Hồng Anh (đại diện giáo viên giảng dạy từ Robohub) xin cam kết thực hiện nghiêm túc các quy định sau:")

    p = doc.add_paragraph()
    r = p.add_run("1. Trách nhiệm chuyên môn và nền nếp giảng dạy:")
    r.bold = True
    doc.add_paragraph("- Thực hiện đúng, đủ Kế hoạch dạy học và hệ thống KHBD đã được Ban Giám hiệu Trường UNIGO phê duyệt; tuyệt đối không tự ý cắt xén chương trình hoặc thay đổi nội dung khi chưa có sự đồng ý của Ban Giám hiệu.")
    doc.add_paragraph("- Đảm bảo giờ giấc lên lớp: Có mặt tại phòng học trước giờ sinh hoạt 15 phút (09h30) để chuẩn bị giáo cụ, kiểm tra sạc pin và đón học sinh; kết thúc buổi học đúng 11h15.")
    doc.add_paragraph("- Tác phong sư phạm chuẩn mực, trang phục lịch sự, ngôn phong chuẩn mực, tận tụy, tôn trọng học sinh và tuân thủ đúng Bộ Quy tắc ứng xử văn hóa học đường của Trường UNIGO.")
    doc.add_paragraph("- Áp dụng phương pháp dạy học STEM tích cực, tạo không khí học tập hào hứng, khuyến khích học sinh khám phá sáng tạo và rèn luyện kỹ năng làm việc nhóm.")

    p = doc.add_paragraph()
    r = p.add_run("2. Trách nhiệm quản lý học sinh và đảm bảo an toàn tuyệt đối:")
    r.bold = True
    doc.add_paragraph("- Quản lý 100% sĩ số học sinh trong suốt 90 phút sinh hoạt; điểm danh chính xác đầu giờ và chỉ cho học sinh ra về khi có phụ huynh hoặc giáo viên phụ trách bàn giao.")
    doc.add_paragraph("- Đảm bảo an toàn tuyệt đối trong phòng học: Hướng dẫn học sinh sử dụng dụng cụ kẹp tháo chốt đúng cách, không để học sinh ngậm, nuốt linh kiện nhỏ hoặc đùa nghịch gây nguy hiểm.")
    doc.add_paragraph("- Kịp thời xử lý và báo cáo ngay với Ban Giám hiệu nếu có bất kỳ tình huống bất thường nào xảy ra liên quan đến sức khỏe, tâm lý hoặc an toàn của học sinh.")

    p = doc.add_paragraph()
    r = p.add_run("3. Trách nhiệm bảo quản cơ sở vật chất và thiết bị học tập:")
    r.bold = True
    doc.add_paragraph("- Chịu trách nhiệm trực tiếp về số lượng, tình trạng kỹ thuật của toàn bộ 25 bộ Kit (15 Initiate, 5 Kinder, 5 Excel) và trang thiết bị phòng thực hành trong suốt quá trình tiếp nhận.")
    doc.add_paragraph("- Rèn luyện cho học sinh thói quen phân loại, kiểm đếm linh kiện cẩn thận vào khay và dọn dẹp vệ sinh phòng học sạch sẽ, ngăn nắp trước khi rời phòng.")
    doc.add_paragraph("- Nếu phát hiện linh kiện bị nứt vỡ, hỏng hóc hoặc thiếu hụt trong quá trình học, phải ghi chép vào sổ nhật ký thiết bị và thông báo cho Bên A/Ban Giám hiệu để có phương án bổ sung kịp thời.")
    doc.add_paragraph("- Cam kết bồi thường hoặc chịu trách nhiệm theo thỏa thuận nếu để xảy ra thất thoát, mất mát thiết bị do lỗi chủ quan trong công tác quản lý.")

    p = doc.add_paragraph()
    r = p.add_run("4. Trách nhiệm phối hợp, báo cáo và truyền thông:")
    r.bold = True
    doc.add_paragraph("- Thường xuyên gửi hình ảnh, video hoạt động thực hành nổi bật của học sinh sau mỗi buổi học về Ban Giám hiệu/bộ phận truyền thông để nhà trường đăng tải bản tin kết nối phụ huynh.")
    doc.add_paragraph("- Báo cáo định kỳ tình hình hoạt động của CLB và tham gia họp giao ban chuyên môn với Ban Giám hiệu UNIGO khi có yêu cầu.")
    doc.add_paragraph("- Phối hợp chặt chẽ với Nhà trường và Thầy Đậu Đình Nguyên tổ chức thành công các kỳ Đánh giá định kỳ, Triển lãm sáng tạo và Ngày hội Robotics Festival UNIGO Cup cuối năm.")

    # ---------------- PHẦN IV: ĐIỀU KHOẢN THI HÀNH ----------------
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("IV. ĐIỀU KHOẢN THI HÀNH")
    r.bold = True
    r.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph("1. Biên bản này có hiệu lực kể từ ngày ký. Hai bên cam kết thực hiện đúng và đầy đủ các nội dung đã thống nhất.")
    doc.add_paragraph("2. Trong quá trình triển khai, nếu có khó khăn, vướng mắc phát sinh, hai bên sẽ chủ động trao đổi trên tinh thần hợp tác, trách nhiệm và báo cáo Ban Giám hiệu để cùng tháo gỡ.")
    doc.add_paragraph("3. Biên bản bàn giao được lập thành 03 (ba) bản có giá trị pháp lý như nhau: Bên bàn giao giữ 01 bản, Bên nhận bàn giao giữ 01 bản, Ban Giám hiệu Trường UNIGO lưu 01 bản để theo dõi và chỉ đạo thực hiện.")

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(6)
    p_end.paragraph_format.space_after = Pt(8)
    r = p_end.add_run("Biên bản được các bên nhất trí thông qua và cùng ký tên xác nhận dưới đây./.")
    r.italic = True

    # ---------------- BẢNG KÝ TÊN 3 BÊN (BGH & 2 GIÁO VIÊN GIAO NHẬN) ----------------
    tbl_sig = doc.add_table(rows=3, cols=3)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl_sig)
    w_sig = [Inches(2.3), Inches(2.3), Inches(2.3)]

    # Row 0: Chức danh
    headers_sig = [
        "ĐẠI DIỆN BAN GIÁM HIỆU\nTRƯỜNG UNIGO",
        "BÊN NHẬN BÀN GIAO\n(GIÁO VIÊN ROBOHUB)",
        "BÊN BÀN GIAO\n(CHỦ NHIỆM CLB)"
    ]
    for i, h in enumerate(headers_sig):
        c = tbl_sig.cell(0, i)
        c.text = h
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(12)
        c.width = w_sig[i]
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)

    # Row 1: Ghi chú ký
    for i in range(3):
        c = tbl_sig.cell(1, i)
        c.text = "(Ký, ghi rõ họ tên)"
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].italic = True
        c.paragraphs[0].runs[0].font.size = Pt(11)
        c.width = w_sig[i]
        set_cell_margins(c, top=20, bottom=20, left=40, right=40)

    # Row 2: Họ tên
    names_sig = [
        "\n\n\n\n...........................................",
        "\n\n\n\nĐào Hồng Anh",
        "\n\n\n\nĐậu Đình Nguyên"
    ]
    for i, n in enumerate(names_sig):
        c = tbl_sig.cell(2, i)
        c.text = n
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if "Đào Hồng Anh" in n or "Đậu Đình Nguyên" in n:
            c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(12)
        c.width = w_sig[i]
        set_cell_margins(c, top=20, bottom=40, left=40, right=40)

    # Định dạng toàn bộ font Times New Roman 13pt
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.15
        for r in p.runs:
            r.font.name = "Times New Roman"
            r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
            if not r.font.size:
                r.font.size = Pt(13)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
                        if not r.font.size:
                            r.font.size = Pt(13)

    # Lưu cả 2 nơi: CLB_Robotics và Hệ thống mẫu văn bản
    os.makedirs(os.path.dirname(OUT_PATH_MAU), exist_ok=True)
    doc.save(OUT_PATH_CLB)
    doc.save(OUT_PATH_MAU)
    print("Successfully generated and saved:")
    print("1.", OUT_PATH_CLB)
    print("2.", OUT_PATH_MAU)

if __name__ == "__main__":
    generate_bien_ban()
