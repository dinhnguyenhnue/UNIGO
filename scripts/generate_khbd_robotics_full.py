import os
import re
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')

# XML border helper
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

def set_doc_margins(doc):
    for section in doc.sections:
        section.top_margin = Inches(0.787)  # 2cm
        section.bottom_margin = Inches(0.787)  # 2cm
        section.left_margin = Inches(1.181)  # 3cm
        section.right_margin = Inches(0.787)  # 2cm

def sanitize_filename(name):
    # Remove accents/special chars for safe filenames
    import unicodedata
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8')
    name = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_')
    return name

def get_kit_name(grade):
    if grade in [1, 2]:
        return "OLLO Kinder"
    elif grade in [3, 4]:
        return "OLLO Initiate"
    else:
        return "OLLO Excel 1"

def build_khbd_doc_primary(grade, lesson_title, lesson_idx, tiet_ppct, yccd):
    doc = docx.Document()
    set_doc_margins(doc)
    kit = get_kit_name(grade)
    
    # Header table 1 row 2 cols
    header_p1 = doc.add_paragraph()
    header_p1.paragraph_format.line_spacing = 1.15
    header_p1.paragraph_format.space_after = Pt(2)
    r1 = header_p1.add_run(f"TUẦN: {(tiet_ppct - 1) // 2 + 1:02d}\t\t\t\tNgày soạn: 01/09/2026\n\t\t\t\t\tNgày dạy: 05/09/2026")
    apply_font(r1, size_pt=12, italic=True)
    
    # Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = title_p.add_run(f"KẾ HOẠCH DẠY HỌC MÔN ROBOTICS - LỚP {grade}\n")
    apply_font(r_t1, size_pt=14, bold=True)
    r_t2 = title_p.add_run(f"BỘ THIẾT BỊ: {kit.upper()}\n")
    apply_font(r_t2, size_pt=13, bold=True)
    r_t3 = title_p.add_run(f"{lesson_title.upper()}\n")
    apply_font(r_t3, size_pt=14, bold=True)
    r_t4 = title_p.add_run(f"(Thời lượng: 1 tiết | Tiết PPCT: {tiet_ppct})")
    apply_font(r_t4, size_pt=12, italic=True)
    
    # Section I: YÊU CẦU CẦN ĐẠT
    p = doc.add_paragraph()
    r = p.add_run("I. YÊU CẦU CẦN ĐẠT")
    apply_font(r, size_pt=13, bold=True)
    
    p = doc.add_paragraph()
    r = p.add_run(f"Mô tả yêu cầu trọng tâm: {yccd}")
    apply_font(r, size_pt=13, italic=True)
    
    p = doc.add_paragraph()
    r = p.add_run("1. Phát triển phẩm chất:")
    apply_font(r, size_pt=13, bold=True)
    
    p = doc.add_paragraph()
    p.add_run("- Chăm chỉ: Hăng hái tham gia hoạt động tìm hiểu linh kiện, kiên trì thực hiện lắp ráp mô hình robot theo hướng dẫn.\n"
              "- Trách nhiệm: Giữ gìn cẩn thận các linh kiện thiết bị trong bộ Kit Robotics, thu dọn gọn gàng và bảo quản thiết bị sau tiết học.\n"
              "- Trung thực: Tự giác làm việc nhóm, tôn trọng kết quả thử nghiệm mô hình robot của bản thân và các nhóm bạn.")
    format_paragraph(p)

    p = doc.add_paragraph()
    r = p.add_run("2. Phát triển năng lực:")
    apply_font(r, size_pt=13, bold=True)

    p = doc.add_paragraph()
    p.add_run(f"- 2.1. Năng lực môn học (Robotics):\n"
              f"  + Nhận biết linh kiện & nguyên lý cơ khí: Nhận biết tên gọi, chức năng các chi tiết khung, chốt nối, động cơ, cảm biến thuộc bộ kit {kit}.\n"
              f"  + Kĩ năng lắp ráp & vận hành: Thực hiện lắp ráp đúng quy trình mô hình {lesson_title}, vận hành chạy thử và tinh chỉnh mô hình.\n"
              f"- 2.2. Năng lực chung:\n"
              f"  + Tự chủ và tự học: Quan sát sơ đồ hướng dẫn lắp ráp, chủ động từng bước hoàn thành mô hình robot.\n"
              f"  + Giao tiếp và hợp tác: Biết phân công nhiệm vụ trong nhóm, phối hợp ăn ý khi thực hành lắp ráp và vận hành robot.\n"
              f"  + Giải quyết vấn đề và sáng tạo: Phát hiện các lỗi sai trong quá trình lắp ráp (kẹt khớp, ngược chốt) và tìm cách khắc phục.\n"
              f"- 2.3. Năng lực số (BẮT BUỘC):\n"
              f"  + Nhận biết & điều khiển thiết bị số: Nhận diện các linh kiện điện tử (mạch điều khiển, động cơ, cảm biến) và nguyên lý nhận/phát tín hiệu số.\n"
              f"  + An toàn thiết bị công nghệ: Biết thao tác an toàn khi gắn pin, tháo lắp chốt nhựa chuyên dụng và bảo vệ mạch điện tử.")
    format_paragraph(p)

    # Section II: ĐỒ DÙNG DẠY HỌC
    p = doc.add_paragraph()
    r = p.add_run("II. ĐỒ DÙNG DẠY HỌC")
    apply_font(r, size_pt=13, bold=True)
    
    p = doc.add_paragraph()
    p.add_run(f"1. Giáo viên: Bộ Kit Robotics {kit} mẫu, máy tính giáo viên, máy chiếu, bài trình chiếu slide hướng dẫn từng bước lắp ráp {lesson_title}, phiếu hướng dẫn thực hành.\n"
              f"2. Học sinh: Bộ Kit Robotics {kit} theo nhóm/cá nhân, dụng cụ tháo chốt, vở ghi bài.")
    format_paragraph(p)

    # Section III: PHƯƠNG PHÁP VÀ KĨ THUẬT DẠY HỌC
    p = doc.add_paragraph()
    r = p.add_run("III. PHƯƠNG PHÁP VÀ KĨ THUẬT DẠY HỌC")
    apply_font(r, size_pt=13, bold=True)
    
    p = doc.add_paragraph()
    p.add_run("- Phương pháp dạy học: Trực quan mô hình, hướng dẫn thực hành qua từng bước (Step-by-step), làm việc nhóm, giải quyết vấn đề (STEM/STEAM).\n"
              "- Kĩ thuật dạy học: Think-Pair-Share, động não, giao nhiệm vụ thực hành phân tầng.")
    format_paragraph(p)

    # Section IV: CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU
    p = doc.add_paragraph()
    r = p.add_run("IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU")
    apply_font(r, size_pt=13, bold=True)

    table = doc.add_table(rows=5, cols=2)
    set_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Table header
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "HOẠT ĐỘNG CỦA GIÁO VIÊN"
    hdr_cells[1].text = "HOẠT ĐỘNG CỦA HỌC SINH"
    for cell in hdr_cells:
        set_cell_margins(cell)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                apply_font(r, size_pt=13, bold=True)

    # Row 1: Khởi động
    r1_cells = table.rows[1].cells
    r1_cells[0].text = (
        f"1. Hoạt động MỞ ĐẦU (Khởi động, kết nối) (5 phút)\n"
        f"a) Chuyển giao nhiệm vụ: GV trình chiếu hình ảnh/video thực tế liên quan đến {lesson_title}. Đặt câu hỏi gợi mở kiến thức.\n"
        f"b) Thực hiện nhiệm vụ: GV quan sát, khơi gợi trí tò mò của HS.\n"
        f"c) Báo cáo: Mời 2-3 đại diện HS trả lời câu hỏi.\n"
        f"d) Kết luận: GV chốt nội dung, dẫn dắt giới thiệu bài học {lesson_title}."
    )
    r1_cells[1].text = (
        f"1. Hoạt động MỞ ĐẦU (Khởi động, kết nối) (5 phút)\n"
        f"a) Tiếp nhận nhiệm vụ: HS tập trung quan sát màn chiếu và lắng nghe câu hỏi.\n"
        f"b) Thực hiện nhiệm vụ: Suy nghĩ độc lập hoặc thảo luận nhanh với bạn bên cạnh.\n"
        f"c) Báo cáo: Hăng hái giơ tay phát biểu ý kiến.\n"
        f"d) Kết luận: Lắng nghe GV chốt kiến thức và mở bộ Kit Robotics."
    )

    # Row 2: Khám phá
    r2_cells = table.rows[2].cells
    r2_cells[0].text = (
        f"2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI (Khám phá linh kiện & cơ chế) (10 phút)\n"
        f"a) Chuyển giao: GV giới thiệu các linh kiện cần dùng trong bài và phân tích nguyên lý hoạt động của mô hình {lesson_title}.\n"
        f"b) Thực hiện: GV hướng dẫn HS cách quan sát sơ đồ lắp ráp 2D/3D.\n"
        f"c) Báo cáo: GV kiểm tra việc chuẩn bị linh kiện của các nhóm.\n"
        f"d) Kết luận: GV chuẩn hóa tên gọi linh kiện và quy trình các bước lắp ráp."
    )
    r2_cells[1].text = (
        f"2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI (Khám phá linh kiện & cơ chế) (10 phút)\n"
        f"a) Tiếp nhận: HS theo dõi bài giảng slide và quan sát mẫu linh kiện GV giơ lên.\n"
        f"b) Thực hiện: Chọn lọc và lấy đúng số lượng linh kiện (khung, chốt, động cơ) ra khay chứa.\n"
        f"c) Báo cáo: Đại diện nhóm giơ khay linh kiện đã chọn đúng để GV kiểm tra.\n"
        f"d) Kết luận: Ghi nhớ vị trí khớp nối và hướng lắp."
    )

    # Row 3: Thực hành
    r3_cells = table.rows[3].cells
    r3_cells[0].text = (
        f"3. HOẠT ĐỘNG LUYỆN TẬP - THỰC HÀNH (Lắp ráp & Vận hành robot) (17 phút)\n"
        f"a) Chuyển giao: GV giao nhiệm vụ cho từng nhóm tiến hành lắp ráp mô hình {lesson_title} theo sơ đồ từng bước.\n"
        f"b) Thực hiện: GV đi tới từng bàn hỗ trợ các nhóm gặp khó khăn, nhắc nhở an toàn.\n"
        f"c) Báo cáo: Cho các nhóm đóng công tắc/vận hành thử nghiệm robot.\n"
        f"d) Kết luận: GV đánh giá sản phẩm hoàn thiện của từng nhóm, tuyên dương nhóm làm đúng và đẹp."
    )
    r3_cells[1].text = (
        f"3. HOẠT ĐỘNG LUYỆN TẬP - THỰC HÀNH (Lắp ráp & Vận hành robot) (17 phút)\n"
        f"a) Tiếp nhận: Phân công công việc (1 bạn xem sơ đồ, 1 bạn lấy chốt, 1 bạn tiến hành gắp chốt lắp khung).\n"
        f"b) Thực hiện: Tiến hành lắp ráp cẩn thận từng bước. Kiểm tra chuyển động cơ khí.\n"
        f"c) Báo cáo: Đặt robot lên bàn thử nghiệm, bật nguồn và quan sát robot hoạt động.\n"
        f"d) Kết luận: Tự tinh chỉnh nếu robot bị kẹt hoặc chạy lệch."
    )

    # Row 4: Vận dụng
    r4_cells = table.rows[4].cells
    r4_cells[0].text = (
        f"4. HOẠT ĐỘNG VẬN DỤNG, SÁNG TẠO (3 phút)\n"
        f"a) Chuyển giao: GV đặt câu hỏi mở rộng: 'Em có thể cải tiến hoặc gắn thêm linh kiện gì để robot {lesson_title} hoạt động tốt hơn?'\n"
        f"b) Thực hiện: Hướng dẫn HS tháo rời linh kiện và phân loại về đúng ngăn khay chứa.\n"
        f"c) Báo cáo: Mời 1-2 HS chia sẻ ý tưởng cải tiến.\n"
        f"d) Kết luận: GV dặn dò bài học sau và nhận xét tiết học."
    )
    r4_cells[1].text = (
        f"4. HOẠT ĐỘNG VẬN DỤNG, SÁNG TẠO (3 phút)\n"
        f"a) Tiếp nhận: Suy nghĩ ý tưởng cải tiến robot.\n"
        f"b) Thực hiện: Nhanh chóng tháo dỡ robot, xếp linh kiện gọn gàng vào hộp bộ Kit.\n"
        f"c) Báo cáo: Phát biểu ý tưởng sáng tạo cá nhân.\n"
        f"d) Kết luận: Đóng hộp kit, đẩy ghế ngăn nắp."
    )

    for row in table.rows[1:]:
        for cell in row.cells:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    apply_font(run, size_pt=12)

    # Section V: ĐIỀU CHỈNH
    p = doc.add_paragraph()
    r = p.add_run("V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY")
    apply_font(r, size_pt=13, bold=True)
    p = doc.add_paragraph()
    p.add_run("(Giáo viên ghi nhận xét và điều chỉnh phương pháp sau khi giảng dạy thực tế)")
    format_paragraph(p, italic=True)

    # Section VI: PHỤ LỤC
    p = doc.add_paragraph()
    r = p.add_run("VI. PHỤ LỤC")
    apply_font(r, size_pt=13, bold=True)
    p = doc.add_paragraph()
    p.add_run(f"Phụ lục 1: Phiếu hướng dẫn quy trình lắp ráp mô hình {lesson_title}.\n"
              f"Phụ lục 2: Bảng Rubric đánh giá tiêu chí sản phẩm robot (Chức năng, Thẩm mỹ, Kĩ năng nhóm, Vệ sinh).")
    format_paragraph(p)

    return doc

def build_khbd_doc_secondary(grade, lesson_title, lesson_idx, tiet_ppct, yccd):
    doc = docx.Document()
    set_doc_margins(doc)
    kit = get_kit_name(grade)

    # Table 0: Info
    t0 = doc.add_table(rows=2, cols=2)
    set_table_borders(t0)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    t0.rows[0].cells[0].text = "Trường: TH & THCS UNIGO\nTổ: Robotics"
    t0.rows[0].cells[1].text = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"
    t0.rows[1].cells[0].text = "Họ và tên GV: Đậu Đình Nguyên"
    t0.rows[1].cells[1].text = f"Ngày soạn: 01/09/2026 | Ngày dạy: 05/09/2026\nLớp: {grade}"
    
    for row in t0.rows:
        for cell in row.cells:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    apply_font(run, size_pt=12, bold=True)

    doc.add_paragraph()

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_title.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}\n")
    apply_font(r1, size_pt=14, bold=True)
    r2 = p_title.add_run(f"MÔN HỌC: ROBOTICS {grade} (BỘ KIT {kit.upper()})\n")
    apply_font(r2, size_pt=13, bold=True)
    r3 = p_title.add_run(f"Thời lượng: 1 tiết | Tiết theo PPCT: {tiet_ppct}")
    apply_font(r3, size_pt=12, italic=True)

    # I. Mục tiêu
    p = doc.add_paragraph()
    r = p.add_run("I. Mục tiêu")
    apply_font(r, size_pt=13, bold=True)

    p = doc.add_paragraph()
    p.add_run(f"1. Kiến thức:\n  - Trọng tâm bài học: {yccd}\n"
              f"  - Hiểu rõ cấu tạo nguyên lý hoạt động cơ khí, thuật toán điều khiển và các linh kiện cảm biến trong mô hình {lesson_title}.\n"
              f"2. Năng lực:\n"
              f"  - Năng lực chung: Tự học và tự nghiên cứu tài liệu hướng dẫn kỹ thuật; giao tiếp và làm việc nhóm hiệu quả; giải quyết vấn đề sáng tạo kỹ thuật.\n"
              f"  - Năng lực đặc thù Robotics: Năng lực thiết kế mô phỏng, kĩ năng thao tác lắp ráp chuẩn xác, năng lực lập trình và hiệu chỉnh thiết bị tự động.\n"
              f"  - Năng lực số (BẮT BUỘC): Làm chủ thiết bị điều khiển thông minh, ứng dụng phần mềm nạp code lập trình cho robot, đảm bảo an toàn thiết bị số và nguồn điện.\n"
              f"3. Phẩm chất:\n"
              f"  - Trách nhiệm trong việc bảo vệ thiết bị công nghệ, trung thực trong báo cáo thử nghiệm sản phẩm, tác phong công nghiệp và tư duy khoa học.")
    format_paragraph(p)

    # II. Thiết bị dạy học và học liệu
    p = doc.add_paragraph()
    r = p.add_run("II. Thiết bị dạy học và học liệu")
    apply_font(r, size_pt=13, bold=True)
    p = doc.add_paragraph()
    p.add_run(f"1. Thiết bị: Bộ Kit Robotics {kit}, máy tính giáo viên và máy tính nhóm học sinh, phần mềm nạp lập trình, máy chiếu chiếu slide bài giảng 3D.\n"
              f"2. Học liệu: Phiếu học tập thực hành, sơ đồ nguyên lý mạch điện tử và thuật toán điều khiển {lesson_title}.")
    format_paragraph(p)

    # III. Tiến trình dạy học
    p = doc.add_paragraph()
    r = p.add_run("III. Tiến trình dạy học")
    apply_font(r, size_pt=13, bold=True)

    t_proc = doc.add_table(rows=5, cols=2)
    set_table_borders(t_proc)
    t_proc.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr = t_proc.rows[0].cells
    hdr[0].text = "HOẠT ĐỘNG CỦA GV VÀ HỌC SINH"
    hdr[1].text = "KẾT QUẢ CẦN ĐẠT / SẢN PHẨM"
    for cell in hdr:
        set_cell_margins(cell)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                apply_font(r, size_pt=13, bold=True)

    # Act 1
    c1 = t_proc.rows[1].cells
    c1[0].text = (
        f"Hoạt động 1: Khởi động (5 phút)\n"
        f"Bước 1. GV giao nhiệm vụ: Trình chiếu tình huống kỹ thuật thực tế liên quan đến {lesson_title}.\n"
        f"Bước 2. HS tiếp nhận nhiệm vụ: Quan sát và suy nghĩ câu hỏi phân tích cơ chế.\n"
        f"Bước 3. Báo cáo thảo luận: Đại diện HS nêu giả thuyết vận hành.\n"
        f"Bước 4. GV đánh giá, chốt kiến thức và vào bài mới."
    )
    c1[1].text = "Sản phẩm: Câu trả lời của HS xác định được vấn đề kỹ thuật cần giải quyết trong bài học."

    # Act 2
    c2 = t_proc.rows[2].cells
    c2[0].text = (
        f"Hoạt động 2: Hình thành kiến thức mới (10 phút)\n"
        f"Bước 1. GV giao nhiệm vụ: Yêu cầu HS nghiên cứu sơ đồ thiết kế và thuật toán cho mô hình {lesson_title}.\n"
        f"Bước 2. HS thực hiện nhiệm vụ: Phân tích cấu trúc phần cứng và luồng điều khiển phần mềm.\n"
        f"Bước 3. Báo cáo: Đại diện nhóm trình bày sơ đồ khối nguyên lý.\n"
        f"Bước 4. GV chuẩn hóa kiến thức kỹ thuật."
    )
    c2[1].text = "Sản phẩm: Sơ đồ khối nguyên lý cơ khí/lập trình được HS phân tích chính xác."

    # Act 3
    c3 = t_proc.rows[3].cells
    c3[0].text = (
        f"Hoạt động 3: Luyện tập - Thực hành (25 phút)\n"
        f"Bước 1. GV giao nhiệm vụ: Yêu cầu các nhóm tiến hành lắp ráp, kết nối cảm biến/động cơ và lập trình vận hành {lesson_title}.\n"
        f"Bước 2. HS thực hiện nhiệm vụ: Lắp ráp phần cứng, viết/nạp chương trình điều khiển, test mô hình.\n"
        f"Bước 3. Báo cáo: Trình diễn khả năng vận hành tự động của robot tại bàn thực hành.\n"
        f"Bước 4. GV nhận xét, chấm điểm sản phẩm."
    )
    c3[1].text = f"Sản phẩm: Mô hình robot {lesson_title} hoàn thiện, hoạt động chính xác theo yêu cầu bài học."

    # Act 4
    c4 = t_proc.rows[4].cells
    c4[0].text = (
        f"Hoạt động 4: Vận dụng & Sáng tạo (5 phút)\n"
        f"Bước 1. GV giao nhiệm vụ: Nêu yêu cầu cải tiến tối ưu thuật toán hoặc thêm tính năng mở rộng cho robot.\n"
        f"Bước 2. HS tiếp nhận & thảo luận ý tưởng mở rộng.\n"
        f"Bước 3. Báo cáo ý tưởng cải tiến.\n"
        f"Bước 4. Đánh giá tổng kết & dọn dẹp vệ sinh phòng thực hành."
    )
    c4[1].text = "Sản phẩm: Báo cáo ý tưởng cải tiến mô hình robot và bộ Kit Robotics được sắp xếp ngăn nắp."

    for row in t_proc.rows[1:]:
        for cell in row.cells:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    apply_font(run, size_pt=12)

    doc.add_paragraph()

    # Signatures
    t_sign = doc.add_table(rows=1, cols=3)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.rows[0].cells[0].text = "BAN GIÁM HIỆU\n\n\n\n__________________"
    t_sign.rows[0].cells[1].text = "TỔ CHUYÊN MÔN\n\n\n\n__________________"
    t_sign.rows[0].cells[2].text = "NGƯỜI SOẠN\n\n\n\nĐậu Đình Nguyên"
    for cell in t_sign.rows[0].cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                apply_font(r, size_pt=12, bold=True)

    return doc

def main():
    ppct_path = r'd:\UNIGO\Phân phối chương trình\Robotics\Kế hoạch dạy học môn Robotics 2026-2027.docx'
    doc_ppct = docx.Document(ppct_path)
    tables = doc_ppct.tables

    created_count = 0
    base_out_dir = r'd:\UNIGO\KHBD_Robotics'

    for grade, t_idx in enumerate(range(4, 12), start=1):
        t = tables[t_idx]
        print(f"\n--- Đang xử lý LỚP {grade} (Table {t_idx}) ---")
        
        lesson_count_in_grade = 0
        for row_idx, row in enumerate(t.rows[1:], start=1):
            cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
            stt = cells[0]
            title = cells[1]
            so_tiet = cells[2]
            tiet_ppct_str = cells[3]
            yccd = cells[4]

            if not title:
                continue

            try:
                tiet_ppct = int(tiet_ppct_str)
            except ValueError:
                tiet_ppct = row_idx

            lesson_count_in_grade += 1
            safe_title = sanitize_filename(title)
            folder_name = f"Bài_{lesson_count_in_grade:02d}"
            
            out_folder = os.path.join(base_out_dir, f"Lớp_{grade}", folder_name)
            os.makedirs(out_folder, exist_ok=True)

            filename = f"KHBD_Robotics_{grade}_Bai{lesson_count_in_grade:02d}_{safe_title}.docx"
            out_filepath = os.path.join(out_folder, filename)

            if grade <= 5:
                doc = build_khbd_doc_primary(grade, title, lesson_count_in_grade, tiet_ppct, yccd)
            else:
                doc = build_khbd_doc_secondary(grade, title, lesson_count_in_grade, tiet_ppct, yccd)

            try:
                doc.save(out_filepath)
                created_count += 1
                print(f"  [+] Đã tạo: Lớp {grade} -> {folder_name} -> {filename}")
            except Exception as e:
                print(f"  [!] Lỗi khi lưu {out_filepath}: {e}")

    print(f"\n==========================================")
    print(f" HOÀN THÀNH TẠO {created_count} FILE KHBD ROBOTICS")
    print(f"==========================================")

if __name__ == '__main__':
    main()
