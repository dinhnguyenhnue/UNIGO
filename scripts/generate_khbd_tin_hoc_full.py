import os
import re
import sys
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')

TPL_PRIMARY = r'd:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx'
TPL_SECONDARY = r'd:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx'
OUT_BASE_DIR = r'd:\UNIGO\KHBD_Tin_học'

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def apply_font(run, font_name="Times New Roman", size_pt=13, bold=False, italic=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color

def format_paragraph(p, font_name="Times New Roman", size_pt=13, line_spacing=1.15, space_after=3, bold=False, italic=False):
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = Pt(space_after)
    for run in p.runs:
        apply_font(run, font_name=font_name, size_pt=size_pt, bold=bold, italic=italic)

def sanitize_filename(name):
    import unicodedata
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8')
    name = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_')
    return name

def populate_primary_doc(grade_str, lesson_title, lesson_idx, tiet_ppct, yccd):
    # Clone template preserving exact header logo drawings
    doc = docx.Document(TPL_PRIMARY)
    
    # Safely update header runs without destroying drawing element (Run 0)
    for s in doc.sections:
        for hp in s.header.paragraphs:
            if len(hp.runs) >= 6:
                hp.runs[2].text = "Đậu Đình Nguyên"
                hp.runs[5].text = f"{grade_str} "

    # Paragraph 0: Week & Dates
    tuan = (tiet_ppct - 1) // 2 + 1 if isinstance(tiet_ppct, int) else 1
    p0 = doc.paragraphs[0]
    p0.text = f"TUẦN: {tuan:02d}\t\t\t\tNgày soạn: 01/09/2026\n\t\t\t\t\tNgày dạy: 05/09/2026"
    format_paragraph(p0, size_pt=12, italic=True)

    p1 = doc.paragraphs[1]
    p1.text = ""

    p2 = doc.paragraphs[2]
    p2.text = f"KẾ HOẠCH DẠY HỌC MÔN TIN HỌC - {grade_str.upper()}"
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p2, size_pt=14, bold=True)

    p3 = doc.paragraphs[3]
    p3.text = f"CHỦ ĐỀ: THẾ GIỚI CÔNG NGHỆ & TƯ DUY SỐ"
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p3, size_pt=13, bold=True)

    p4 = doc.paragraphs[4]
    p4.text = f"BÀI: {lesson_title.upper()} (Thời lượng: 1 tiết | Tiết PPCT: {tiet_ppct})"
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p4, size_pt=13, bold=True)

    # Clear paragraphs 5 onwards
    for p in doc.paragraphs[5:]:
        p.text = ""

    # Build structured body
    doc.paragraphs[6].text = "I. YÊU CẦU CẦN ĐẠT:"
    format_paragraph(doc.paragraphs[6], bold=True)

    doc.paragraphs[7].text = f"- Sau bài học này, học sinh sẽ: {yccd}"
    format_paragraph(doc.paragraphs[7], italic=True)

    doc.paragraphs[8].text = "1. Phát triển phẩm chất:"
    format_paragraph(doc.paragraphs[8], bold=True)

    doc.paragraphs[9].text = (
        "- Chăm chỉ: Chủ động tham gia các hoạt động tìm hiểu bài học, hăng hái phát biểu và hoàn thành bài thực hành.\n"
        "- Trách nhiệm: Giữ gìn an toàn thiết bị máy tính phòng thực hành, có ý thức bảo vệ tài sản công cộng và vệ sinh lớp học.\n"
        "- Trung thực: Tự giác thực hiện các bài tập cá nhân/nhóm, tôn trọng kết quả học tập của bản thân và bạn bè."
    )
    format_paragraph(doc.paragraphs[9])

    doc.paragraphs[10].text = "2. Phát triển năng lực:"
    format_paragraph(doc.paragraphs[10], bold=True)

    doc.paragraphs[11].text = (
        f"- 2.1. Năng lực môn học (Tin học):\n"
        f"  + NLa (Nhận biết & Khám phá): Nhận diện các thiết bị công nghệ, phần mềm và ứng dụng thực tế theo nội dung bài {lesson_title}.\n"
        f"  + NLb (Sử dụng & Quản lý): Thao tác thành thạo bàn phím, chuột hoặc phần mềm theo đúng quy trình hướng dẫn.\n"
        f"- 2.2. Năng lực chung:\n"
        f"  + Tự chủ và tự học: Tự giác theo dõi thao tác mẫu của giáo viên, tự mình hoàn thành bài tập thực hành trên máy tính.\n"
        f"  + Giao tiếp và hợp tác: Lắng nghe ý kiến của bạn trong nhóm đôi/nhóm lớn, phối hợp cùng hoàn thành nhiệm vụ chung.\n"
        f"  + Giải quyết vấn đề và sáng tạo: Biết xử lý các tình huống lỗi đơn giản (kẹt chuột, nhầm trang web, gõ sai phím).\n"
        f"- 2.3. Năng lực số (BẮT BUỘC):\n"
        f"  + Khai thác thông tin số: Nhận biết và thao tác với thông tin hiển thị trên màn hình máy tính/thiết bị thông minh.\n"
        f"  + An toàn & Văn hóa số: Ý thức tuân thủ quy tắc an toàn về điện, bảo vệ mắt và giữ vệ sinh thiết bị công nghệ."
    )
    format_paragraph(doc.paragraphs[11])

    doc.paragraphs[12].text = "II. ĐỒ DÙNG DẠY HỌC:"
    format_paragraph(doc.paragraphs[12], bold=True)

    doc.paragraphs[13].text = (
        f"1. Giáo viên: Phòng máy tính có kết nối mạng, máy chiếu, slide bài trình chiếu minh họa {lesson_title}, phiếu học tập thực hành.\n"
        f"2. Học sinh: Vở ghi, SGK Tin học, thiết bị máy tính thực hành."
    )
    format_paragraph(doc.paragraphs[13])

    doc.paragraphs[14].text = "III. PHƯƠNG PHÁP VÀ KĨ THUẬT DẠY HỌC:"
    format_paragraph(doc.paragraphs[14], bold=True)

    doc.paragraphs[15].text = (
        "- Phương pháp dạy học: Trực quan minh họa, làm mẫu thao tác (Demonstration), thực hành cá nhân/nhóm, học qua trò chơi.\n"
        "- Kĩ thuật dạy học: Think-Pair-Share, đặt câu hỏi gợi mở, chia sẻ nhóm đôi."
    )
    format_paragraph(doc.paragraphs[15])

    doc.paragraphs[16].text = "IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU:"
    format_paragraph(doc.paragraphs[16], bold=True)

    # Populate Table 0
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        set_table_borders(t0)
        t0.alignment = WD_TABLE_ALIGNMENT.CENTER

        while len(t0.rows) < 5:
            t0.add_row()

        # Table header
        hdr = t0.rows[0].cells
        hdr[0].text = "HOẠT ĐỘNG CỦA GIÁO VIÊN"
        hdr[1].text = "HOẠT ĐỘNG CỦA HỌC SINH"
        for cell in hdr:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    apply_font(r, size_pt=13, bold=True)

        # Row 1: Khởi động
        r1 = t0.rows[1].cells
        r1[0].text = (
            f"1. Hoạt động MỞ ĐẦU (Khởi động, kết nối) (5 phút)\n"
            f"a) Chuyển giao: GV chiếu video/hình ảnh sinh động liên quan đến {lesson_title}. Nêu câu hỏi kết nối thực tế.\n"
            f"b) Thực hiện: GV quan sát, gợi mở trí tò mò của học sinh.\n"
            f"c) Báo cáo: Mời 2-3 HS đưa ra câu trả lời.\n"
            f"d) Kết luận: GV chốt kiến thức, dẫn dắt giới thiệu vào bài mới."
        )
        r1[1].text = (
            f"1. Hoạt động MỞ ĐẦU (Khởi động, kết nối) (5 phút)\n"
            f"a) Tiếp nhận: HS tập trung quan sát màn chiếu, lắng nghe tình huống.\n"
            f"b) Thực hiện: Suy nghĩ cá nhân hoặc trao đổi nhanh với bạn bên cạnh.\n"
            f"c) Báo cáo: Hăng hái giơ tay phát biểu ý kiến.\n"
            f"d) Kết luận: Lắng nghe GV và sẵn sàng bài học."
        )

        # Row 2: Khám phá
        r2 = t0.rows[2].cells
        r2[0].text = (
            f"2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI (Khám phá) (12 phút)\n"
            f"a) Chuyển giao: GV hướng dẫn HS quan sát màn hình, thao tác mẫu trên máy chiếu các bước trong bài {lesson_title}.\n"
            f"b) Thực hiện: GV di chuyển quan sát, giải đáp thắc mắc cho HS.\n"
            f"c) Báo cáo: Gọi đại diện HS lên máy tính giáo viên thực hiện lại thao tác mẫu.\n"
            f"d) Kết luận: GV chốt kiến thức chuẩn và lưu ý thao tác đúng."
        )
        r2[1].text = (
            f"2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI (Khám phá) (12 phút)\n"
            f"a) Tiếp nhận: HS quan sát kỹ các thao tác GV minh họa trên máy chiếu.\n"
            f"b) Thực hiện: Theo dõi SGK/phiếu hướng dẫn, ghi nhớ thứ tự các bước.\n"
            f"c) Báo cáo: HS đại diện lên thao tác mẫu cho cả lớp quan sát.\n"
            f"d) Kết luận: Ghi nhận các lưu ý quan trọng vào vở."
        )

        # Row 3: Thực hành
        r3 = t0.rows[3].cells
        r3[0].text = (
            f"3. HOẠT ĐỘNG LUYỆN TẬP - THỰC HÀNH (15 phút)\n"
            f"a) Chuyển giao: GV giao bài tập thực hành cá nhân trên máy tính theo yêu cầu bài {lesson_title}.\n"
            f"b) Thực hiện: GV tới từng vị trí máy tính hướng dẫn HS còn lúng túng.\n"
            f"c) Báo cáo: Hướng dẫn HS kiểm tra bài làm của bạn bên cạnh (Peer assessment).\n"
            f"d) Kết luận: GV nhận xét kết quả thực hành, tuyên dương các HS hoàn thành xuất sắc."
        )
        r3[1].text = (
            f"3. HOẠT ĐỘNG LUYỆN TẬP - THỰC HÀNH (15 phút)\n"
            f"a) Tiếp nhận: HS bật phần mềm/nhiệm vụ thực hành trên máy tính.\n"
            f"b) Thực hiện: Tự giác hoàn thành bài tập. Nhờ GV giúp đỡ nếu gặp lỗi.\n"
            f"c) Báo cáo: Quan sát bài làm của bạn bên cạnh và góp ý.\n"
            f"d) Kết luận: Hoàn thiện sản phẩm trên máy tính."
        )

        # Row 4: Vận dụng
        r4 = t0.rows[4].cells
        r4[0].text = (
            f"4. HOẠT ĐỘNG VẬN DỤNG, SÁNG TẠO (3 phút)\n"
            f"a) Chuyển giao: GV đưa ra câu hỏi/tình huống vận dụng thực tế cuộc sống.\n"
            f"b) Thực hiện: Hướng dẫn HS lưu tệp (nếu có), đóng phần mềm và tắt máy đúng quy trình.\n"
            f"c) Báo cáo: Mời 1-2 HS phát biểu ứng dụng bài học.\n"
            f"d) Kết luận: GV dặn dò bài sau và nhận xét tiết học."
        )
        r4[1].text = (
            f"4. HOẠT ĐỘNG VẬN DỤNG, SÁNG TẠO (3 phút)\n"
            f"a) Tiếp nhận: Lắng nghe tình huống vận dụng.\n"
            f"b) Thực hiện: Thực hiện thao tác Shutdown tắt máy an toàn, xếp bàn phím/chuột gọn gàng.\n"
            f"c) Báo cáo: Trả lời câu hỏi vận dụng.\n"
            f"d) Kết luận: Đẩy ghế gọn gàng trước khi rời phòng máy."
        )

        for row in t0.rows[1:]:
            for cell in row.cells:
                set_cell_margins(cell)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        apply_font(run, size_pt=12)

    doc.paragraphs[18].text = "V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY:"
    format_paragraph(doc.paragraphs[18], bold=True)

    doc.paragraphs[19].text = "(Giáo viên ghi nhận xét và điều chỉnh phương pháp sau khi giảng dạy thực tế)"
    format_paragraph(doc.paragraphs[19], italic=True)

    doc.paragraphs[20].text = "VI. PHỤ LỤC:"
    format_paragraph(doc.paragraphs[20], bold=True)

    doc.paragraphs[21].text = f"Phụ lục 1: Phiếu bài tập thực hành {lesson_title}.\nPhụ lục 2: Rubric đánh giá kỹ năng thực hành máy tính."
    format_paragraph(doc.paragraphs[21])

    return doc

def populate_secondary_doc(grade_str, lesson_title, lesson_idx, tiet_ppct, yccd):
    # Clone THCS template preserving exact logo drawings
    doc = docx.Document(TPL_SECONDARY)
    
    # Table 0: Info
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        set_table_borders(t0)
        t0.rows[0].cells[0].text = "Trường: TH & THCS UNIGO\nTổ: Khoa học Tự nhiên & Công nghệ"
        t0.rows[0].cells[1].text = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"
        t0.rows[1].cells[0].text = "Họ và tên GV: Đậu Đình Nguyên"
        t0.rows[1].cells[1].text = f"Ngày soạn: 01/09/2026 | Ngày dạy: 05/09/2026\nLớp: {grade_str}"
        for row in t0.rows:
            for cell in row.cells:
                set_cell_margins(cell)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        apply_font(r, size_pt=12, bold=True)

    # Title paragraph P2
    for p in doc.paragraphs:
        if "TÊN BÀI DẠY" in p.text:
            p.text = f"TÊN BÀI DẠY: {lesson_title.upper()}\n"
            r1 = p.add_run(f"MÔN HỌC: TIN HỌC - {grade_str.upper()}\n")
            apply_font(r1, size_pt=13, bold=True)
            r2 = p.add_run(f"Thời lượng: 1 tiết | Tiết theo PPCT: {tiet_ppct}")
            apply_font(r2, size_pt=12, italic=True)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            break

    # Replace body content
    for p in doc.paragraphs:
        if "1. Kiến thức:" in p.text:
            p.text = f"1. Kiến thức: {yccd}. Nắm vững các khái niệm, quy trình và kỹ năng sử dụng công nghệ trong bài {lesson_title}."
            format_paragraph(p)
        elif "2. Năng lực:" in p.text or "- Năng lực chung:" in p.text:
            p.text = (
                f"2. Năng lực:\n"
                f"  - Năng lực đặc thù (Tin học): Nhận biết, phân tích và thao tác thành thạo các phần mềm/công cụ trong bài {lesson_title}.\n"
                f"  - Năng lực số (BẮT BUỘC): Khai thác thông tin số an toàn, bảo mật dữ liệu cá nhân, tuân thủ pháp luật và văn hóa ứng xử trong môi trường số.\n"
                f"  - Năng lực chung: Tự chủ tự học nghiên cứu tài liệu; giao tiếp hợp tác làm việc nhóm; giải quyết vấn đề kỹ thuật số sáng tạo."
            )
            format_paragraph(p)
        elif "3. Phẩm chất:" in p.text:
            p.text = "3. Phẩm chất: Tác phong công nghiệp, trung thực trong học tập, có ý thức trách nhiệm bảo vệ thiết bị số và tài nguyên mạng."
            format_paragraph(p)

    # Table 1: Progress
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        set_table_borders(t1)
        while len(t1.rows) < 5:
            t1.add_row()

        hdr = t1.rows[0].cells
        hdr[0].text = "HOẠT ĐỘNG CỦA GV VÀ HỌC SINH"
        hdr[1].text = "KẾT QUẢ CẦN ĐẠT / SẢN PHẨM"
        for cell in hdr:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    apply_font(r, size_pt=13, bold=True)

        c1 = t1.rows[1].cells
        c1[0].text = (
            f"Hoạt động 1: Khởi động (5 phút)\n"
            f"Bước 1. GV giao nhiệm vụ: Chiếu hình ảnh/tình huống thực tế đặt vấn đề cho bài {lesson_title}.\n"
            f"Bước 2. HS tiếp nhận nhiệm vụ: Quan sát và suy nghĩ câu hỏi gợi mở.\n"
            f"Bước 3. Báo cáo thảo luận: Đại diện HS nêu ý kiến ban đầu.\n"
            f"Bước 4. GV chốt kiến thức và giới thiệu bài mới."
        )
        c1[1].text = "Sản phẩm: Câu trả lời của HS xác định đúng vấn đề trọng tâm cần giải quyết."

        c2 = t1.rows[2].cells
        c2[0].text = (
            f"Hoạt động 2: Hình thành kiến thức mới (12 phút)\n"
            f"Bước 1. GV giao nhiệm vụ: Hướng dẫn HS nghiên cứu SGK và quan sát thao tác minh họa trên phần mềm.\n"
            f"Bước 2. HS thực hiện nhiệm vụ: Đọc tài liệu, phân tích quy trình các bước thực hiện.\n"
            f"Bước 3. Báo cáo: Đại diện nhóm trình bày sơ đồ các bước thực hiện.\n"
            f"Bước 4. GV chuẩn hóa kiến thức lý thuyết."
        )
        c2[1].text = f"Sản phẩm: Kiến thức chuẩn về {lesson_title} được HS tiếp thu và ghi nhớ."

        c3 = t1.rows[3].cells
        c3[0].text = (
            f"Hoạt động 3: Luyện tập - Thực hành (23 phút)\n"
            f"Bước 1. GV giao nhiệm vụ: Yêu cầu HS mở phần mềm trên máy tính và tiến hành làm bài tập thực hành.\n"
            f"Bước 2. HS thực hiện nhiệm vụ: Thao tác thực hành cá nhân/nhóm trên máy tính.\n"
            f"Bước 3. Báo cáo: Trình diễn sản phẩm trên màn hình máy tính cho GV đánh giá.\n"
            f"Bước 4. GV nhận xét, chấm điểm thực hành."
        )
        c3[1].text = f"Sản phẩm: Tệp/kết quả thực hành {lesson_title} trên máy tính đạt yêu cầu."

        c4 = t1.rows[4].cells
        c4[0].text = (
            f"Hoạt động 4: Vận dụng & Nhiệm vụ về nhà (5 phút)\n"
            f"Bước 1. GV giao nhiệm vụ: Yêu cầu HS nêu ứng dụng thực tế hoặc giao bài tập vận dụng mở rộng.\n"
            f"Bước 2. HS tiếp nhận nhiệm vụ mở rộng.\n"
            f"Bước 3. Báo cáo ý tưởng vận dụng.\n"
            f"Bước 4. GV hướng dẫn HS lưu tệp và tắt máy an toàn."
        )
        c4[1].text = "Sản phẩm: Báo cáo ý tưởng vận dụng và máy tính được tắt an toàn theo quy định."

        for row in t1.rows[1:]:
            for cell in row.cells:
                set_cell_margins(cell)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        apply_font(run, size_pt=12)

    return doc

def get_grade_9_lessons():
    return [
        ("Bài 1. Thế giới kĩ thuật số", 1, "Sự nhận biết sự phát triển của máy tính và các thiết bị số thông minh làm thay đổi xã hội loài người. Phân tích tác động hai mặt của thế giới kỹ thuật số."),
        ("Bài 2. Thông tin trong giải quyết vấn đề", 2, "Giải thích được tầm quan trọng của thông tin trong việc đưa ra quyết định giải quyết vấn đề thực tế."),
        ("Bài 3. Mạng máy tính", 3, "Nêu được khái niệm mạng máy tính, thành phần mạng và lợi ích của việc kết nối mạng."),
        ("Bài 4. Internet", 4, "Trình bày được đặc điểm của Internet, các dịch vụ phổ biến trên Internet và cách truy cập thông tin."),
        ("Ôn tập Đánh giá định kỳ 1", 5, "Hệ thống hóa kiến thức chuẩn bị Đánh giá định kỳ 1."),
        ("Đánh giá định kỳ 1", 6, "Kiểm tra đánh giá kết quả học tập môn Tin học."),
        ("Bài 5. Văn hoá ứng xử trên mạng", 7, "Nhận biết các quy tắc văn hóa, đạo đức và pháp luật khi giao tiếp, chia sẻ thông tin trên không gian mạng."),
        ("Bài 6. Sử dụng phần mềm bảng tính nâng cao", 8, "Sử dụng các hàm địa chỉ tương đối/tuyệt đối, hàm điều kiện IF và xử lý dữ liệu phức tạp trong bảng tính."),
        ("Bài 7. Sắp xếp và lọc dữ liệu nâng cao", 9, "Thực hiện các thao tác sắp xếp dữ liệu theo nhiều tiêu chí và lọc dữ liệu tùy biến."),
        ("Bài 8. Trình bày dữ liệu bằng biểu đồ nâng cao", 10, "Lựa chọn và tạo các dạng biểu đồ thích hợp để trực quan hóa dữ liệu thống kê."),
        ("Ôn tập Đánh giá định kỳ 2", 11, "Hệ thống hóa kiến thức chuẩn bị Đánh giá định kỳ 2."),
        ("Đánh giá định kỳ 2", 12, "Kiểm tra đánh giá kết quả học tập môn Tin học."),
        ("Bài 9. Làm quen với phần mềm tạo trang web", 13, "Nhận biết cấu trúc trang web, tạo và chỉnh sửa nội dung trang web đơn giản."),
        ("Bài 10. Chèn hình ảnh và tạo liên kết cho trang web", 14, "Thực hiện thao tác chèn hình ảnh, tạo siêu liên kết (hyperlink) giữa các trang web."),
        ("Bài 11. Thuật toán và lập trình", 15, "Mô tả thuật toán bằng sơ đồ khối, viết chương trình giải quyết bài toán thực tế."),
        ("Bài 12. Câu lệnh rẽ nhánh và lặp trong lập trình", 16, "Sử dụng thành thạo cấu trúc rẽ nhánh và cấu trúc lặp trong ngôn ngữ lập trình."),
        ("Ôn tập Đánh giá định kỳ 3", 17, "Hệ thống hóa kiến thức chuẩn bị Đánh giá định kỳ 3."),
        ("Đánh giá định kỳ 3", 18, "Kiểm tra đánh giá kết quả học tập môn Tin học."),
        ("Bài 13. Dự án học tập: Tạo sản phẩm kĩ thuật số", 19, "Làm việc nhóm lập kế hoạch, thiết kế và hoàn thiện sản phẩm kỹ thuật số phục vụ học tập."),
        ("Bài 14. Tin học và định hướng nghề nghiệp", 20, "Tìm hiểu các ngành nghề trong lĩnh vực CNTT và ứng dụng tin học trong xã hội hiện đại."),
        ("Ôn tập Đánh giá định kỳ 4", 21, "Hệ thống hóa kiến thức chuẩn bị Đánh giá định kỳ 4."),
        ("Đánh giá định kỳ 4", 22, "Kiểm tra đánh giá kết quả học tập môn Tin học cả năm.")
    ]

def main():
    # 1. Clear existing KHBD_Tin_học directory
    print("--- 1. Xóa các Kế hoạch bài dạy Tin học cũ ---")
    if os.path.exists(OUT_BASE_DIR):
        for item in os.listdir(OUT_BASE_DIR):
            item_path = os.path.join(OUT_BASE_DIR, item)
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)
            else:
                os.remove(item_path)
        print("  [+] Đã xóa sạch thư mục KHBD_Tin_học cũ.")

    os.makedirs(OUT_BASE_DIR, exist_ok=True)

    # 2. Parse PPCT
    ppct_doc_path = r'd:\UNIGO\Phân phối chương trình\Tin học\Kế hoạch dạy học môn Tin học 2026-2027.docx'
    doc_ppct = docx.Document(ppct_doc_path)
    tables = doc_ppct.tables

    grade_configs = [
        ("Tiền_tiểu_học", "Tiền tiểu học", 3, True),
        ("Lớp_1", "Lớp 1", 4, True),
        ("Lớp_2", "Lớp 2", 5, True),
        ("Lớp_3", "Lớp 3", 6, True),
        ("Lớp_4", "Lớp 4", 7, True),
        ("Lớp_5", "Lớp 5", 8, True),
        ("Lớp_6", "Lớp 6", 9, False),
        ("Lớp_7", "Lớp 7", 10, False),
        ("Lớp_8", "Lớp 8", 11, False),
    ]

    total_created = 0

    for folder_prefix, grade_str, t_idx, is_primary in grade_configs:
        print(f"\n--- Đang xử lý {grade_str} (Table {t_idx}) ---")
        t = tables[t_idx]
        lesson_count = 0

        for row_idx, row in enumerate(t.rows[1:], start=1):
            cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
            title = cells[1]
            tiet_str = cells[2]
            ppct_str = cells[3]
            yccd = cells[4]

            if not title or title.startswith('Chủ đề'):
                continue

            try:
                tiet_ppct = int(ppct_str)
            except ValueError:
                tiet_ppct = row_idx

            lesson_count += 1
            safe_title = sanitize_filename(title)
            bai_folder = f"Bài_{lesson_count:02d}"
            
            out_dir = os.path.join(OUT_BASE_DIR, folder_prefix, bai_folder)
            os.makedirs(out_dir, exist_ok=True)

            filename = f"KHBD_Tin_hoc_{folder_prefix}_Bai{lesson_count:02d}_{safe_title}.docx"
            out_file = os.path.join(out_dir, filename)

            if is_primary:
                doc = populate_primary_doc(grade_str, title, lesson_count, tiet_ppct, yccd)
            else:
                doc = populate_secondary_doc(grade_str, title, lesson_count, tiet_ppct, yccd)

            try:
                doc.save(out_file)
                total_created += 1
                print(f"  [+] Đã tạo: {folder_prefix} -> {bai_folder} -> {filename}")
            except Exception as e:
                print(f"  [!] Lỗi khi lưu {out_file}: {e}")

    # Process Lớp 9
    print(f"\n--- Đang xử lý Lớp 9 (SGK Kết nối tri thức 9) ---")
    grade_9_lessons = get_grade_9_lessons()
    lesson_count = 0
    for title, ppct, yccd in grade_9_lessons:
        lesson_count += 1
        safe_title = sanitize_filename(title)
        bai_folder = f"Bài_{lesson_count:02d}"
        
        out_dir = os.path.join(OUT_BASE_DIR, "Lớp_9", bai_folder)
        os.makedirs(out_dir, exist_ok=True)

        filename = f"KHBD_Tin_hoc_Lop_9_Bai{lesson_count:02d}_{safe_title}.docx"
        out_file = os.path.join(out_dir, filename)

        doc = populate_secondary_doc("Lớp 9", title, lesson_count, ppct, yccd)
        try:
            doc.save(out_file)
            total_created += 1
            print(f"  [+] Đã tạo: Lớp_9 -> {bai_folder} -> {filename}")
        except Exception as e:
            print(f"  [!] Lỗi khi lưu {out_file}: {e}")

    print(f"\n==========================================")
    print(f" HOÀN THÀNH TẠO {total_created} FILE KHBD TIN HỌC")
    print(f"==========================================")

if __name__ == '__main__':
    main()
