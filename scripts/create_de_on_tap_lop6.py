# -*- coding: utf-8 -*-
import os, sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def set_no_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="none"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:bottom w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'  <w:insideH w:val="none"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def add_p(doc, text="", bold=False, italic=False, size_pt=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4, line_spacing=1.15, color_rgb=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    return p

def main():
    doc = docx.Document()

    # Set page margins
    for sec in doc.sections:
        sec.top_margin = Inches(0.7)
        sec.bottom_margin = Inches(0.7)
        sec.left_margin = Inches(0.75)
        sec.right_margin = Inches(0.75)

    # Header table (No borders)
    hdr_table = doc.add_table(rows=2, cols=2)
    hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_no_borders(hdr_table)

    # Row 0
    c00 = hdr_table.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r00 = p00.add_run("TRƯỜNG TIỂU HỌC & THCS UNIGO\nTỔ CHUYÊN MÔN THCS")
    r00.font.name = "Times New Roman"
    r00.font.size = Pt(12)
    r00.font.bold = True

    c01 = hdr_table.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r01 = p01.add_run("ĐỀ ÔN TẬP VÀ RÈN LUYỆN KIẾN THỨC\nNĂM HỌC 2026 – 2027")
    r01.font.name = "Times New Roman"
    r01.font.size = Pt(12)
    r01.font.bold = True

    # Row 1
    c10 = hdr_table.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r10 = p10.add_run("Môn: TIN HỌC – KHỐI 6")
    r10.font.name = "Times New Roman"
    r10.font.size = Pt(12)
    r10.font.italic = True

    c11 = hdr_table.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r11 = p11.add_run("Chủ đề 1: Máy tính và cộng đồng (Bài 1, 2, 3)\nThời gian làm bài: 45 phút")
    r11.font.name = "Times New Roman"
    r11.font.size = Pt(12)
    r11.font.italic = True

    add_p(doc, "", space_after=4)

    # Student info table
    info_table = doc.add_table(rows=2, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(info_table, color="666666", sz="4")
    
    info_table.cell(0, 0).paragraphs[0].add_run("Họ và tên học sinh: ................................................................").font.name = "Times New Roman"
    info_table.cell(0, 1).paragraphs[0].add_run("Lớp: 6A1").font.name = "Times New Roman"
    info_table.cell(0, 2).paragraphs[0].add_run("Điểm").font.name = "Times New Roman"
    info_table.cell(0, 3).paragraphs[0].add_run("Lời nhận xét của thầy cô giáo").font.name = "Times New Roman"

    info_table.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_table.cell(0, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_table.cell(0, 2).paragraphs[0].runs[0].font.bold = True
    info_table.cell(0, 3).paragraphs[0].runs[0].font.bold = True
    set_cell_background(info_table.cell(0, 2), "F2F4F8")
    set_cell_background(info_table.cell(0, 3), "F2F4F8")

    info_table.cell(1, 0).paragraphs[0].add_run("Ngày kiểm tra / rèn luyện: ......./......./2026").font.name = "Times New Roman"
    info_table.cell(1, 1).paragraphs[0].add_run("STT: .........").font.name = "Times New Roman"
    info_table.cell(1, 2).paragraphs[0].add_run("\n\n").font.name = "Times New Roman"
    info_table.cell(1, 3).paragraphs[0].add_run("\n....................................................................................\n....................................................................................").font.name = "Times New Roman"

    for r in info_table.rows:
        for c in r.cells:
            set_cell_margins(c, 80, 80, 100, 100)
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(12)

    add_p(doc, "", space_after=6)

    # MA TRẬN ĐỀ
    add_p(doc, "A. KHUNG MA TRẬN & BẢN ĐẶC TẢ KIẾN THỨC", bold=True, size_pt=13, color_rgb=RGBColor(27, 79, 155))
    add_p(doc, "• Bài 1: Thông tin và dữ liệu (Khái niệm thông tin, dữ liệu, vật mang tin, các dạng thông tin cơ bản, tầm quan trọng của thông tin).\n"
               "• Bài 2: Xử lí thông tin (Quy trình 4 bước: Thu nhận → Lưu trữ → Xử lý → Truyền; vai trò vượt trội của máy tính trong xử lý thông tin).\n"
               "• Bài 3: Thông tin trong máy tính (Khái niệm bit, biểu diễn số, văn bản, hình ảnh, âm thanh thành dãy bit; đơn vị đo dung lượng thông tin B, KB, MB, GB, TB; tính toán dung lượng thiết bị lưu trữ).",
          italic=True, size_pt=12, space_after=6)

    # Matrix Table
    mt_table = doc.add_table(rows=5, cols=6)
    mt_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(mt_table, color="000000", sz="4")
    headers = ["Nội dung / Chủ đề", "Nhận biết", "Thông hiểu", "Vận dụng", "Vận dụng cao", "Tổng điểm"]
    for i, h in enumerate(headers):
        cell = mt_table.cell(0, i)
        set_cell_background(cell, "EBF3FE")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.bold = True

    matrix_rows = [
        ["Bài 1. Thông tin và dữ liệu", "2 câu TN (0.5đ)", "2 câu TN (0.5đ)", "1 ý TNĐ/S (0.25đ)", "-", "1.25 điểm"],
        ["Bài 2. Xử lí thông tin", "2 câu TN (0.5đ)", "2 câu TN (0.5đ)", "1 câu TL (1.5đ)", "-", "2.5 điểm"],
        ["Bài 3. Thông tin trong máy tính", "4 câu TN (1.0đ)", "4 câu TN (1.0đ)", "1 câu TL (1.0đ)", "1 câu TL (1.5đ) + TNĐ/S (1.75đ)", "6.25 điểm"],
        ["Tổng cộng", "2.0 điểm (20%)", "2.0 điểm (20%)", "3.0 điểm (30%)", "3.0 điểm (30%)", "10.0 điểm (100%)"]
    ]
    for r_idx, row_data in enumerate(matrix_rows, start=1):
        for c_idx, val in enumerate(row_data):
            cell = mt_table.cell(r_idx, c_idx)
            if r_idx == 4:
                set_cell_background(cell, "F3EEFF")
            p = cell.paragraphs[0]
            if c_idx > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            run.font.name = "Times New Roman"
            run.font.size = Pt(11.5)
            if r_idx == 4 or c_idx == 0:
                run.font.bold = True

    for r in mt_table.rows:
        for c in r.cells:
            set_cell_margins(c, 70, 70, 90, 90)

    add_p(doc, "", space_after=6)

    # PHẦN I: TRẮC NGHIỆM
    add_p(doc, "B. NỘI DUNG ĐỀ LUYỆN TẬP", bold=True, size_pt=14, color_rgb=RGBColor(27, 79, 155))
    add_p(doc, "PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (4.0 điểm)", bold=True, size_pt=13)
    add_p(doc, "Khoanh tròn vào chữ cái (A, B, C hoặc D) đứng trước câu trả lời đúng nhất:", italic=True, size_pt=12, space_after=6)

    questions = [
        # Bài 1
        ("Câu 1. Khẳng định nào sau đây diễn đạt đúng nhất về khái niệm Thông tin?",
         "A. Thông tin là các con số, chữ viết và hình ảnh được lưu trữ trong sách vở.",
         "B. Thông tin là những hiểu biết của con người về thế giới xung quanh và về chính bản thân mình.",
         "C. Thông tin là các thiết bị phần cứng hiện đại như máy tính, điện thoại, tivi.",
         "D. Thông tin là các tín hiệu điện tử chạy bên trong mạng cáp quang."),

        ("Câu 2. Khi nhìn thấy biển báo giao thông vẽ hình người đi bộ gạch chéo màu đỏ, điều em tiếp nhận được gọi là gì?",
         "A. Vật mang tin.",
         "B. Dữ liệu văn bản.",
         "C. Thông tin (đoạn đường cấm người đi bộ).",
         "D. Tệp dữ liệu máy tính."),

        ("Câu 3. Phát biểu nào sau đây phân biệt CHÍNH XÁC giữa Dữ liệu và Vật mang tin?",
         "A. Dữ liệu là chiếc USB, còn vật mang tin là bài hát ghi trên USB.",
         "B. Dữ liệu là thông tin được ghi lên vật mang tin dưới dạng số, chữ, hình ảnh; còn vật mang tin là phương tiện chứa dữ liệu đó.",
         "C. Dữ liệu và vật mang tin là hai khái niệm hoàn toàn giống nhau.",
         "D. Vật mang tin luôn luôn là giấy viết hoặc sách in, không thể là thiết bị điện tử."),

        ("Câu 4. Trong các tình huống sau, trường hợp nào cho thấy tầm quan trọng của thông tin đối với việc ra quyết định?",
         "A. Xem dự báo thời tiết báo trời mưa to kèm sấm sét nên cả lớp quyết định dời buổi cắm trại sang tuần sau.",
         "B. Học sinh ngồi đọc một cuốn truyện tranh giải trí sau giờ học căng thẳng.",
         "C. Chiếc đồng hồ treo tường đang chạy pin bình thường trong phòng học.",
         "D. Bạn Nam dùng bút màu xanh để viết bài vào vở."),

        # Bài 2
        ("Câu 5. Quy trình xử lí thông tin của con người và máy tính bao gồm mấy bước cơ bản?",
         "A. 2 bước: Nhập và Xuất dữ liệu.",
         "B. 3 bước: Thu nhận → Xử lý → Xóa bỏ.",
         "C. 4 bước: Thu nhận → Lưu trữ → Xử lí → Truyền thông tin.",
         "D. 5 bước: Nhìn → Nghe → Nghĩ → Viết → Đọc."),

        ("Câu 6. Khi bạn Minh nghe tiếng chuông báo thức reo, não bạn nghĩ 'đã 6 giờ sáng rồi, phải dậy đi học thôi' và ngồi dậy tắt chuông. Hoạt động 'nghe tiếng chuông reo' thuộc bước nào trong quy trình xử lí thông tin?",
         "A. Thu nhận thông tin.",
         "B. Xử lí thông tin.",
         "C. Lưu trữ thông tin.",
         "D. Truyền thông tin."),

        ("Câu 7. Bộ phận nào trong máy tính đóng vai trò tương tự như 'bộ não' của con người trong việc trực tiếp xử lí thông tin?",
         "A. Màn hình máy tính (Monitor).",
         "B. Chuột và bàn phím (Mouse & Keyboard).",
         "C. Bộ vi xử lí trung tâm (CPU).",
         "D. Ổ đĩa cứng (Hard Disk)."),

        ("Câu 8. Vì sao máy tính được coi là công cụ hỗ trợ xử lí thông tin vô cùng hiệu quả và vượt trội so với con người?",
         "A. Máy tính có cảm xúc phong phú và biết sáng tạo nghệ thuật độc lập.",
         "B. Máy tính có tốc độ tính toán siêu nhanh, dung lượng lưu trữ khổng lồ, làm việc liên tục không mệt mỏi và độ chính xác rất cao.",
         "C. Máy tính tự động làm mọi việc mà không cần con người lập trình hay điều khiển.",
         "D. Máy tính chỉ có thể xử lí được văn bản tiếng Anh, không xử lí được hình ảnh."),

        # Bài 3
        ("Câu 9. Trong máy tính, đơn vị nhỏ nhất dùng để đo lượng thông tin và lưu trữ là gì?",
         "A. Byte (B).",
         "B. Bit (b).",
         "C. Kilobyte (KB).",
         "D. Megabyte (MB)."),

        ("Câu 10. Bit là viết tắt của cụm từ tiếng Anh nào sau đây và nhận những giá trị nào?",
         "A. Binary digIT, nhận một trong hai giá trị là 0 hoặc 1.",
         "B. Basic Information Tool, nhận các giá trị từ 1 đến 10.",
         "C. Byte Information Transfer, nhận các chữ cái từ A đến Z.",
         "D. Binary Technology, nhận các số chẵn 0, 2, 4, 6."),

        ("Câu 11. Vì sao máy tính điện tử lại sử dụng hệ nhị phân (chỉ gồm các kí hiệu 0 và 1) để biểu diễn và xử lí mọi thông tin?",
         "A. Vì con người chỉ thích số 0 và số 1.",
         "B. Vì các mạch điện tử trong máy tính rất dễ dàng và tin cậy khi nhận diện hai trạng thái vật lý (đóng mạch/ngắt mạch, có điện/không có điện).",
         "C. Vì máy tính không thể đếm được các số lớn hơn 1.",
         "D. Vì bảng chữ cái tiếng Việt không tương thích với bàn phím máy tính."),

        ("Câu 12. Trong biểu diễn hình ảnh kĩ thuật số dạng đen trắng (như chữ A hay hình trái tim trong SGK), mỗi điểm ảnh (pixel) được quy ước biểu diễn như thế nào?",
         "A. Mỗi pixel được biểu diễn bằng 10 bit.",
         "B. Mỗi pixel đen được quy ước là 1 (hoặc 0) và pixel trắng quy ước ngược lại bằng 0 (hoặc 1).",
         "C. Mỗi pixel được biểu diễn bằng một đoạn văn bản mô tả màu sắc.",
         "D. Hình ảnh không thể chuyển đổi thành dãy bit."),

        ("Câu 13. Mối quan hệ chuẩn giữa đơn vị Byte và đơn vị Bit là gì?",
         "A. 1 Byte = 2 Bit.",
         "B. 1 Byte = 4 Bit.",
         "C. 1 Byte = 8 Bit.",
         "D. 1 Byte = 10 Bit."),

        ("Câu 14. Đơn vị 1 Gigabyte (GB) tương đương với khoảng bao nhiêu Byte?",
         "A. Khoảng 1 nghìn Byte (10³ B).",
         "B. Khoảng 1 triệu Byte (10⁶ B).",
         "C. Khoảng 1 tỉ Byte (10⁹ B).",
         "D. Khoảng 1 nghìn tỉ Byte (10¹² B)."),

        ("Câu 15. Thiết bị nào sau đây thường có dung lượng lưu trữ lớn nhất trong số các thiết bị được liệt kê?",
         "A. Đĩa CD thông thường (~ 700 MB).",
         "B. Thẻ nhớ điện thoại MicroSD (32 GB).",
         "C. Đĩa DVD tiêu chuẩn (~ 4.7 GB).",
         "D. Ổ đĩa cứng HDD / SSD gắn trong máy tính (1 TB đến 2 TB)."),

        ("Câu 16. Thứ tự sắp xếp các đơn vị đo dung lượng thông tin từ NHỎ đến LỚN nào sau đây là ĐÚNG?",
         "A. Bit → Byte → MB → KB → GB → TB.",
         "B. Bit → Byte → KB → MB → GB → TB.",
         "C. Byte → Bit → KB → GB → MB → TB.",
         "D. TB → GB → MB → KB → Byte → Bit.")
    ]

    for q, a, b, c, d in questions:
        add_p(doc, q, bold=True, size_pt=12.5, space_after=2)
        add_p(doc, f"   {a}\n   {b}\n   {c}\n   {d}", size_pt=12, space_after=5)

    add_p(doc, "", space_after=6)

    # PHẦN II: ĐÚNG / SAI
    add_p(doc, "PHẦN II. CÂU HỎI ĐÚNG / SAI (2.0 điểm)", bold=True, size_pt=13)
    add_p(doc, "Em hãy ghi chữ Đ (Đúng) hoặc S (Sai) vào ô tương ứng trước mỗi khẳng định sau:", italic=True, size_pt=12, space_after=6)

    ds_table = doc.add_table(rows=9, cols=3)
    ds_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ds_table, color="000000", sz="4")

    ds_headers = ["STT", "Nội dung khẳng định kiến thức", "Chọn (Đ/S)"]
    for i, h in enumerate(ds_headers):
        cell = ds_table.cell(0, i)
        set_cell_background(cell, "EBF3FE")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.bold = True

    ds_data = [
        ("1", "Văn bản, hình ảnh và âm thanh khi đưa vào máy tính đều phải được chuyển đổi (mã hóa) thành dãy các bit 0 và 1.", "......"),
        ("2", "Dung lượng bộ nhớ RAM của máy tính cá nhân hiện nay thường chỉ lưu trữ được tối đa khoảng 512 Byte.", "......"),
        ("3", "Một đĩa quang CD (dung lượng khoảng 700 MB) có thể lưu trữ được nhiều dữ liệu hơn một thẻ nhớ 32 GB.", "......"),
        ("4", "Để chuyển đổi số từ 0 đến 7 thành dãy bit trong SGK, ta có thể dùng phương pháp chia đôi liên tiếp và thu được các dãy gồm 3 bit (ví dụ: số 3 mã hóa thành 011).", "......"),
        ("5", "Khi em gửi một bức ảnh qua Zalo cho bạn, thao tác này thuộc bước 'Truyền thông tin' trong quy trình xử lí thông tin.", "......"),
        ("6", "Trong cùng một kích thước ảnh, ảnh có số lượng điểm ảnh (pixel) càng nhiều thì dung lượng tệp càng nhỏ và ảnh hiển thị càng mờ.", "......"),
        ("7", "1 Kilobyte (KB) theo quy chuẩn tính toán chính xác của hệ nhị phân bằng đúng 1024 Byte, chứ không phải tròn 1000 Byte.", "......"),
        ("8", "Thông tin và dữ liệu hoàn toàn độc lập với nhau, dữ liệu không thể chuyển hóa thành thông tin cho con người hiểu được.", "......")
    ]

    for r_idx, (stt, content, ans) in enumerate(ds_data, start=1):
        c0 = ds_table.cell(r_idx, 0)
        c1 = ds_table.cell(r_idx, 1)
        c2 = ds_table.cell(r_idx, 2)
        c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        r0 = c0.paragraphs[0].add_run(stt)
        r1 = c1.paragraphs[0].add_run(content)
        r2 = c2.paragraphs[0].add_run(ans)
        for r in (r0, r1, r2):
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)

    for r in ds_table.rows:
        for c in r.cells:
            set_cell_margins(c, 60, 60, 80, 80)

    add_p(doc, "", space_after=6)

    # PHẦN III: TỰ LUẬN
    add_p(doc, "PHẦN III. TỰ LUẬN & BÀI TOÁN THỰC TẾ (4.0 điểm)", bold=True, size_pt=13)
    
    # Câu 1
    add_p(doc, "Câu 1 (1.5 điểm) – Xử lí thông tin trong đời sống:", bold=True, size_pt=12.5)
    add_p(doc, "Bạn Lan đang đi xe đạp trên đường đến trường, khi tới ngã tư, Lan nhìn thấy đèn tín hiệu giao thông chuyển sang màu đỏ. Ngay lập tức, Lan giảm tốc độ, bóp phanh và dừng xe lại trước vạch quy định.\n"
               "a) Em hãy phân tích quy trình 4 bước xử lí thông tin trong tình huống trên của bạn Lan (Thu nhận → Lưu trữ → Xử lý → Truyền/Hành động)?\n"
               "b) Hãy chỉ ra bộ phận nào trên cơ thể bạn Lan đóng vai trò thu nhận thông tin, bộ phận nào đóng vai trò xử lí thông tin?", size_pt=12, space_after=6)
    add_p(doc, "Bài làm:\n...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................", italic=True, size_pt=11, space_after=8)

    # Câu 2
    add_p(doc, "Câu 2 (1.0 điểm) – Biểu diễn thông tin trong máy tính:", bold=True, size_pt=12.5)
    add_p(doc, "Trong SGK Tin học 6, một phần bảng mã kí tự sang dãy bit được cho như sau:\n"
               "   • Kí tự 'A' được mã hoá thành: 01000001\n"
               "   • Kí tự 'B' được mã hoá thành: 01000010\n"
               "   • Kí tự 'C' được mã hoá thành: 01000011\n"
               "a) Em hãy cho biết từ 'CAB' sẽ được máy tính biểu diễn thành dãy bit như thế nào?\n"
               "b) Từ 'CAB' ở câu a) gồm bao nhiêu kí tự và chiếm bao nhiêu Byte, bao nhiêu Bit trong bộ nhớ máy tính?", size_pt=12, space_after=6)
    add_p(doc, "Bài làm:\n...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................", italic=True, size_pt=11, space_after=8)

    # Câu 3
    add_p(doc, "Câu 3 (1.5 điểm) – Bài toán tính toán dung lượng thiết bị nhớ thực tế:", bold=True, size_pt=12.5)
    add_p(doc, "Bạn An có một chiếc máy ảnh kĩ thuật số sử dụng thẻ nhớ MicroSD dung lượng 16 GB.\n"
               "a) Giả sử mỗi bức ảnh chất lượng cao chụp bằng máy ảnh có dung lượng trung bình khoảng 16 MB (Megabyte). Hỏi chiếc thẻ nhớ 16 GB của bạn An có thể lưu trữ tối đa được khoảng bao nhiêu bức ảnh như vậy? (Biết 1 GB = 1 024 MB).\n"
               "b) Bố của An có một chiếc USB dung lượng 8 GB còn trống hoàn toàn. Bạn An muốn sao chép một thư mục bài tập và video học tập nặng 9 500 MB vào chiếc USB đó. Hỏi chiếc USB có chứa hết thư mục đó không? Vì sao?", size_pt=12, space_after=6)
    add_p(doc, "Bài làm:\n...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................\n"
               "...........................................................................................................................................................................................", italic=True, size_pt=11, space_after=12)

    # ĐÁP ÁN & HƯỚNG DẪN CHẤM
    doc.add_page_break()
    add_p(doc, "ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM CHI TIẾT", bold=True, size_pt=14, align=WD_ALIGN_PARAGRAPH.CENTER, color_rgb=RGBColor(27, 79, 155))
    add_p(doc, "(Dành cho Giáo viên chấm điểm và Học sinh đối chiếu tự rèn luyện)", italic=True, size_pt=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    # Đáp án trắc nghiệm
    add_p(doc, "I. ĐÁP ÁN PHẦN I: TRẮC NGHIỆM KHÁCH QUAN (4.0 điểm – Mỗi câu đúng đạt 0.25 điểm)", bold=True, size_pt=12.5)
    
    ans_table = doc.add_table(rows=2, cols=9)
    ans_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ans_table, color="000000", sz="4")

    row1_cells = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8"]
    row1_ans   = ["Đ/A", "B", "C", "B", "A", "C", "A", "C", "B"]
    row2_cells = ["Câu", "9", "10", "11", "12", "13", "14", "15", "16"]
    row2_ans   = ["Đ/A", "B", "A", "B", "B", "C", "C", "D", "B"]

    # Fill table 1 (questions 1-8)
    for c_idx, (q_t, a_t) in enumerate(zip(row1_cells, row1_ans)):
        cell_q = ans_table.cell(0, c_idx)
        cell_a = ans_table.cell(1, c_idx)
        set_cell_background(cell_q, "EBF3FE")
        cell_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_q = cell_q.paragraphs[0].add_run(q_t)
        r_a = cell_a.paragraphs[0].add_run(a_t)
        r_q.font.name = "Times New Roman"
        r_a.font.name = "Times New Roman"
        r_q.font.bold = True
        r_a.font.bold = True
        if a_t != "Đ/A":
            r_a.font.color.rgb = RGBColor(194, 65, 12)

    for r in ans_table.rows:
        for c in r.cells:
            set_cell_margins(c, 50, 50, 60, 60)

    add_p(doc, "", space_after=4)

    # Second row table for 9-16
    ans_table2 = doc.add_table(rows=2, cols=9)
    ans_table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ans_table2, color="000000", sz="4")

    for c_idx, (q_t, a_t) in enumerate(zip(row2_cells, row2_ans)):
        cell_q = ans_table2.cell(0, c_idx)
        cell_a = ans_table2.cell(1, c_idx)
        set_cell_background(cell_q, "EBF3FE")
        cell_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_q = cell_q.paragraphs[0].add_run(q_t)
        r_a = cell_a.paragraphs[0].add_run(a_t)
        r_q.font.name = "Times New Roman"
        r_a.font.name = "Times New Roman"
        r_q.font.bold = True
        r_a.font.bold = True
        if a_t != "Đ/A":
            r_a.font.color.rgb = RGBColor(194, 65, 12)

    for r in ans_table2.rows:
        for c in r.cells:
            set_cell_margins(c, 50, 50, 60, 60)

    add_p(doc, "", space_after=6)

    # Đáp án Đúng / Sai
    add_p(doc, "II. ĐÁP ÁN PHẦN II: CÂU HỎI ĐÚNG / SAI (2.0 điểm – Mỗi ý đúng đạt 0.25 điểm)", bold=True, size_pt=12.5)
    ds_ans_text = (
        "• Ý 1: ĐÚNG (Mọi văn bản, hình ảnh, âm thanh đều được mã hóa thành dãy bit 0 và 1).\n"
        "• Ý 2: SAI (Bộ nhớ RAM của máy tính cá nhân hiện nay thường từ 4 GB đến 16 GB, chứ không phải 512 Byte).\n"
        "• Ý 3: SAI (Đĩa CD có dung lượng khoảng 700 MB, nhỏ hơn rất nhiều so với thẻ nhớ 32 GB = 32 768 MB).\n"
        "• Ý 4: ĐÚNG (Áp dụng đúng thuật toán chia đôi liên tiếp trong SGK Hình 1.3 để mã hóa số 3 thành 011).\n"
        "• Ý 5: ĐÚNG (Gửi ảnh qua Zalo là hành động đưa thông tin đến người khác, thuộc bước Truyền thông tin).\n"
        "• Ý 6: SAI (Càng nhiều điểm ảnh pixel thì hình ảnh càng sắc nét, chi tiết nhưng dung lượng tệp sẽ càng lớn).\n"
        "• Ý 7: ĐÚNG (Theo hệ nhị phân lũy thừa của 2, 1 KB = 2¹⁰ B = 1 024 Byte).\n"
        "• Ý 8: SAI (Dữ liệu được xử lí sẽ chuyển hóa thành thông tin có ích giúp con người hiểu biết và đưa ra quyết định)."
    )
    add_p(doc, ds_ans_text, size_pt=12, space_after=6)

    # Đáp án Tự luận
    add_p(doc, "III. ĐÁP ÁN VÀ THANG ĐIỂM PHẦN III: TỰ LUẬN (4.0 điểm)", bold=True, size_pt=12.5)
    
    tl_content = [
        ("Câu 1 (1.5 điểm):",
         "a) Phân tích 4 bước xử lí thông tin (1.0 điểm – mỗi bước đúng được 0.25đ):\n"
         "   - Bước 1 (Thu nhận thông tin): Mắt bạn Lan nhìn thấy tín hiệu đèn giao thông chuyển sang màu đỏ.\n"
         "   - Bước 2 (Lưu trữ thông tin): Hình ảnh đèn đỏ và luật giao thông được ghi nhớ trong não bộ bạn Lan.\n"
         "   - Bước 3 (Xử lí thông tin): Não bạn Lan phân tích 'đèn đỏ nghĩa là phải dừng lại, nếu đi tiếp sẽ vi phạm luật và nguy hiểm'.\n"
         "   - Bước 4 (Truyền thông tin / Hành động): Não truyền lệnh xuống tay bóp phanh, chân hạ xuống để xe dừng lại trước vạch.\n"
         "b) Bộ phận cơ thể (0.5 điểm):\n"
         "   - Mắt (thị giác) đóng vai trò thu nhận thông tin (0.25đ).\n"
         "   - Bộ não đóng vai trò xử lí thông tin (0.25đ)."),

        ("Câu 2 (1.0 điểm):",
         "a) Mã hóa từ 'CAB' (0.5 điểm):\n"
         "   - Kí tự 'C' = 01000011\n"
         "   - Kí tự 'A' = 01000001\n"
         "   - Kí tự 'B' = 01000010\n"
         "   → Dãy bit biểu diễn từ 'CAB' là: 01000011 01000001 01000010 (hoặc viết liền: 010000110100000101000010).\n"
         "b) Dung lượng lưu trữ (0.5 điểm):\n"
         "   - Từ 'CAB' gồm 3 kí tự (0.1đ).\n"
         "   - Mỗi kí tự được biểu diễn bằng 1 Byte (8 bit), do đó từ 'CAB' chiếm: 3 Byte (0.2đ).\n"
         "   - Đổi ra bit: 3 Byte × 8 = 24 Bit (0.2đ)."),

        ("Câu 3 (1.5 điểm):",
         "a) Tính số bức ảnh lưu trữ được (0.75 điểm):\n"
         "   - Đổi dung lượng thẻ nhớ từ GB sang MB: 16 GB = 16 × 1 024 MB = 16 384 MB (0.25đ).\n"
         "   - Số lượng bức ảnh thẻ nhớ có thể lưu trữ tối đa là:\n"
         "       16 384 MB : 16 MB/ảnh = 1 024 (bức ảnh) (0.5đ).\n"
         "   (Học sinh nếu ước lượng 16 GB ≈ 16 000 MB → 16 000 : 16 = 1 000 bức ảnh: cho 0.5đ/0.75đ).\n"
         "b) Kiểm tra dung lượng USB (0.75 điểm):\n"
         "   - Đổi dung lượng USB từ GB sang MB:\n"
         "       8 GB = 8 × 1 024 MB = 8 192 MB (0.25đ).\n"
         "   - So sánh: Vì 8 192 MB < 9 500 MB (dung lượng thư mục cần sao chép) (0.25đ).\n"
         "   - Kết luận: Chiếc USB KHÔNG chứa hết được thư mục đó vì dung lượng của USB nhỏ hơn dung lượng của thư mục (thiếu khoảng 1 308 MB) (0.25đ).")
    ]

    for title, content in tl_content:
        add_p(doc, title, bold=True, size_pt=12, color_rgb=RGBColor(27, 79, 155))
        add_p(doc, content, size_pt=12, space_after=6)

    # Save document
    out_dir = r"D:\UNIGO\KHBD_Tin_học\Lớp_6"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "De_on_tap_Bai_1_2_3_Tin_hoc_6.docx")
    
    try:
        doc.save(out_path)
        print(f"SUCCESS: Saved file to {out_path}")
    except PermissionError:
        out_path_alt = os.path.join(out_dir, "De_on_tap_Bai_1_2_3_Tin_hoc_6_v2.docx")
        doc.save(out_path_alt)
        print(f"SUCCESS: Saved file to {out_path_alt}")

if __name__ == "__main__":
    main()
