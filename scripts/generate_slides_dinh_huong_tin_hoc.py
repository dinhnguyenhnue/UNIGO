import os
import re
import sys
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

sys.stdout.reconfigure(encoding='utf-8')

TPL_PATH = r'd:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide UNIGO.pptx'
BASE_OUT_DIR = r'd:\UNIGO\KHBD_Tin_học'

# Color Palette Tokens
C_NAVY = RGBColor(0x1E, 0x3A, 0x8A)      # Primary Banner / Accent
C_BLUE = RGBColor(0x02, 0x84, 0xC7)      # Secondary Blue
C_ORANGE = RGBColor(0xEA, 0x58, 0x0C)    # Badge / Accent Orange
C_AMBER = RGBColor(0xF5, 0x9E, 0x0B)     # Highlight Amber
C_GREEN = RGBColor(0x10, 0xB9, 0x81)     # Accent Green
C_BG = RGBColor(0xF8, 0xFA, 0xFC)        # Slide background pastel
C_CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)   # Card White
C_BORDER = RGBColor(0xE2, 0xE8, 0xF0)    # Soft Gray Border
C_TEXT_DARK = RGBColor(0x0F, 0x17, 0x2A) # Dark text
C_TEXT_MUTED = RGBColor(0x47, 0x55, 0x69)# Muted text
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)     # Pure White

def set_slide_bg(slide, color=C_BG):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header_banner(slide, grade, slide_title):
    # Top banner bar
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.1))
    banner.fill.solid()
    banner.fill.fore_color.rgb = C_NAVY
    banner.line.color.rgb = C_NAVY
    
    # Badge Pill
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.2), Inches(2.8), Inches(0.35))
    pill.fill.solid()
    pill.fill.fore_color.rgb = C_ORANGE
    pill.line.color.rgb = C_ORANGE
    tf_pill = pill.text_frame
    tf_pill.word_wrap = True
    p_p = tf_pill.paragraphs[0]
    p_p.alignment = PP_ALIGN.CENTER
    r_p = p_p.add_run()
    r_p.text = f"TIN HỌC {grade} • TIẾT 0"
    r_p.font.name = "Times New Roman"
    r_p.font.size = Pt(11)
    r_p.font.bold = True
    r_p.font.color.rgb = C_WHITE

    # Title text in Banner
    tb = slide.shapes.add_textbox(Inches(3.4), Inches(0.12), Inches(9.4), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = slide_title.upper()
    r.font.name = "Times New Roman"
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

def add_footer(slide):
    # Bottom UNIGO bar
    footer_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45))
    footer_bar.fill.solid()
    footer_bar.fill.fore_color.rgb = RGBColor(0x00, 0x70, 0xC0)
    footer_bar.line.color.rgb = RGBColor(0x00, 0x70, 0xC0)
    
    tf = footer_bar.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO — NĂM HỌC 2026 - 2027"
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

def add_card(slide, left, top, width, height, title="", items=None, accent_color=C_NAVY):
    # Main card rounded container
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD_BG
    card.line.color.rgb = C_BORDER
    
    # Left accent strip
    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + Inches(0.1), Inches(0.15), height - Inches(0.2))
    strip.fill.solid()
    strip.fill.fore_color.rgb = accent_color
    strip.line.color.rgb = accent_color

    # Card Content Textbox
    tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.15), width - Inches(0.45), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    
    if title:
        p_t = tf.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(16)
        r_t.font.bold = True
        r_t.font.color.rgb = accent_color
        p_t.space_after = Pt(6)

    if items:
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.space_after = Pt(4)
            p.line_spacing = 1.15
            r_num = p.add_run()
            r_num.text = f"• "
            r_num.font.name = "Times New Roman"
            r_num.font.size = Pt(13)
            r_num.font.bold = True
            r_num.font.color.rgb = accent_color
            
            r_txt = p.add_run()
            r_txt.text = item
            r_txt.font.name = "Times New Roman"
            r_txt.font.size = Pt(13)
            r_txt.font.color.rgb = C_TEXT_DARK

def create_title_slide(prs, grade, subtitle_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank
    set_slide_bg(slide, C_BG)

    # Hero Card Container
    hero = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.0), Inches(11.33), Inches(5.0))
    hero.fill.solid()
    hero.fill.fore_color.rgb = C_CARD_BG
    hero.line.color.rgb = C_BORDER

    # Left Hero Accent Bar
    hero_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.0), Inches(0.3), Inches(5.0))
    hero_strip.fill.solid()
    hero_strip.fill.fore_color.rgb = C_ORANGE
    hero_strip.line.color.rgb = C_ORANGE

    # Badge Pill
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.6), Inches(1.4), Inches(3.5), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_NAVY
    badge.line.color.rgb = C_NAVY
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = f"CHƯƠNG TRÌNH TIN HỌC {grade} — GDPT 2018"
    r_b.font.name = "Times New Roman"
    r_b.font.size = Pt(11)
    r_b.font.bold = True
    r_b.font.color.rgb = C_WHITE

    # Title Text
    tb = slide.shapes.add_textbox(Inches(1.6), Inches(2.1), Inches(10.2), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p1 = tf.paragraphs[0]
    r1 = p1.add_run()
    r1.text = f"ĐỊNH HƯỚNG MÔN HỌC TIN HỌC {grade}\n"
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(28)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY

    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = subtitle_text.upper()
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(20)
    r2.font.bold = True
    r2.font.color.rgb = C_ORANGE

    # Teacher / School Info Box
    tb_info = slide.shapes.add_textbox(Inches(1.6), Inches(4.5), Inches(10.2), Inches(1.2))
    tf_info = tb_info.text_frame
    p_i = tf_info.paragraphs[0]
    r_i = p_i.add_run()
    r_i.text = "Trường: TH & THCS UNIGO   |   Giáo viên: Đậu Đình Nguyên   |   Năm học: 2026 - 2027"
    r_i.font.name = "Times New Roman"
    r_i.font.size = Pt(14)
    r_i.font.italic = True
    r_i.font.color.rgb = C_TEXT_MUTED

    add_footer(slide)
    return slide

def build_orientation_deck(grade):
    prs = Presentation(TPL_PATH)

    subtitles = {
        1: "Nội quy & An toàn phòng máy tính",
        2: "Em trở thành nhà sáng tạo số",
        3: "Khám phá môn Tin học 3",
        4: "Khám phá môn Tin học 4",
        5: "Khám phá môn Tin học 5 & Kỹ năng số thế kỷ 21",
        6: "Phương pháp học tập & An toàn số",
        7: "Tổng quan chương trình & Kỹ năng số",
        8: "Định hướng học tập & Tư duy máy tính"
    }
    subtitle = subtitles.get(grade, f"Khám phá môn Tin học {grade}")

    # Slide 1: Trang bìa
    create_title_slide(prs, grade, subtitle)

    # Slide 2: Mục tiêu tiết học
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s2)
    add_header_banner(s2, grade, "I. MỤC TIÊU TIẾT HỌC (YÊU CẦU CẦN ĐẠT)")
    add_footer(s2)
    
    add_card(s2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.1), 
             title="1. Năng lực môn học & Kỹ năng số", 
             items=[
                 f"Nắm vững cấu trúc môn Tin học {grade} theo chuẩn GDPT 2018.",
                 "Hiểu rõ quy trình sử dụng thiết bị máy tính an toàn và đúng cách.",
                 "Hình thành thói quen thao tác đúng quy chuẩn phòng thực hành Tin học.",
                 "Ứng dụng Khung năng lực số (CV 3456) vào quá trình học tập."
             ], accent_color=C_NAVY)
             
    add_card(s2, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.1), 
             title="2. Phẩm chất & Năng lực chung", 
             items=[
                 "Chăm chỉ: Tích cực khám phá kiến thức và chủ động chuẩn bị bài.",
                 "Trách nhiệm: Bảo vệ tài sản công cộng, giữ gìn phòng máy sạch sẽ.",
                 "Trung thực: Tự giác làm việc nhóm, tôn trọng sản phẩm số của bạn.",
                 "Hợp tác: Phối hợp hiệu quả với bạn cùng bàn trong giờ thực hành."
             ], accent_color=C_ORANGE)

    # Slide 3: Chào mừng & Khởi động
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s3)
    add_header_banner(s3, grade, "II. HOẠT ĐỘNG 1: KHỞI ĐỘNG — CHÀO MỪNG & KẾT NỐI")
    add_footer(s3)

    add_card(s3, Inches(0.8), Inches(1.5), Inches(11.6), Inches(2.3),
             title="🎮 Trò chơi khởi động: 'Thế giới số xung quanh em'",
             items=[
                 "Quan sát hình ảnh các thiết bị công nghệ quen thuộc trong cuộc sống (Máy tính, Máy tính bảng, Robot...).",
                 "Thảo luận nhanh: Em đã từng dùng máy tính để làm những công việc gì?",
                 "Báo cáo: 3-4 bạn học sinh đại diện giơ tay chia sẻ trải nghiệm cá nhân."
             ], accent_color=C_BLUE)

    add_card(s3, Inches(0.8), Inches(4.1), Inches(11.6), Inches(2.5),
             title="💡 Gợi mở thông điệp kết nối",
             items=[
                 f"Môn Tin học {grade} sẽ giúp các em chuyển từ người 'sử dụng máy tính' thành 'nhà sáng tạo số' thông minh!",
                 "Chúng ta sẽ cùng nhau khám phá tri thức mới, rèn luyện kỹ năng thực hành và tạo ra các sản phẩm công nghệ ấn tượng."
             ], accent_color=C_GREEN)

    # Slide 4: Tổng quan chương trình
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s4)
    add_header_banner(s4, grade, f"III. HOẠT ĐỘNG 2: TỔNG QUAN CHƯƠNG TRÌNH TIN HỌC {grade}")
    add_footer(s4)

    if grade <= 5:
        topics = [
            ("Chủ đề A: Máy tính & Em", "Khám phá phần cứng, phần mềm, chuột và bàn phím máy tính."),
            ("Chủ đề B: Mạng máy tính & Internet", "Tìm hiểu thế giới Internet, an toàn thông tin và giao tiếp số."),
            ("Chủ đề C: Tổ chức lưu trữ dữ liệu", "Sắp xếp thư mục, quản lý tệp tin và tìm kiếm thông tin hiệu quả."),
            ("Chủ đề D: Đạo đức & Pháp luật số", "Văn hóa ứng xử trên mạng, tôn trọng bản quyền và thông tin cá nhân."),
            ("Chủ đề E: Ứng dụng tin học", "Soạn thảo văn bản, vẽ tranh, trình chiếu slide và giải trí số."),
            ("Chủ đề F: Giải quyết vấn đề & Lập trình", "Tư duy thuật toán, làm quen lập trình Scratch sinh động.")
        ]
    else:
        topics = [
            ("Chủ đề 1: Máy tính & Xã hội tri thức", "Lịch sử phát triển máy tính, hệ điều hành và thiết bị vào/ra."),
            ("Chủ đề 2: Mạng máy tính & Internet", "Mạng cục bộ (LAN), Internet, dịch vụ đám mây và tìm kiếm nâng cao."),
            ("Chủ đề 3: Đạo đức, pháp luật & văn hóa số", "An toàn số, bản quyền tác giả, phòng chống lừa đảo trên mạng."),
            ("Chủ đề 4: Ứng dụng tin học", "Soạn thảo nâng cao, Bảng tính Excel, Biên tập ảnh & Video số."),
            ("Chủ đề 5: Giải thuật & Lập trình", "Tư duy máy tính (Computational Thinking), lập trình ngôn ngữ bậc cao.")
        ]

    card_w = Inches(3.6)
    card_h = Inches(2.4)
    for idx, (t_title, t_desc) in enumerate(topics[:6]):
        col = idx % 3
        row = idx // 3
        left = Inches(0.8) + col * Inches(3.9)
        top = Inches(1.5) + row * Inches(2.6)
        add_card(s4, left, top, card_w, card_h, title=t_title, items=[t_desc], accent_color=C_NAVY if row==0 else C_ORANGE)

    # Slide 5: Phương pháp học tập
    s5 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s5)
    add_header_banner(s5, grade, "IV. PHƯƠNG PHÁP HỌC TẬP & TƯ DUY CÔNG NGHỆ")
    add_footer(s5)

    add_card(s5, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="🚀 4 Phương pháp học tập chủ đạo",
             items=[
                 "1. Học qua trải nghiệm (Hands-on Practice): Tăng thời lượng thực hành trực tiếp trên máy tính.",
                 "2. Học qua dự án (Project-based Learning): Thực hiện bài tập nhóm và tạo ra sản phẩm thực tế.",
                 "3. Học qua giải quyết vấn đề: Tìm lỗi sai (Debug) và tối ưu quy trình xử lý công việc.",
                 "4. Tự học & Đọc tài liệu: Khai thác học liệu số và sách giáo khoa hiệu quả."
             ], accent_color=C_BLUE)

    add_card(s5, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="🧠 Tư duy máy tính (Computational Thinking)",
             items=[
                 "• Phân rã bài toán (Decomposition): Chia nhỏ vấn đề lớn thành các phần dễ xử lý.",
                 "• Nhận diện mẫu (Pattern Recognition): Tìm điểm chung giữa các nhiệm vụ.",
                 "• Trừu tượng hóa (Abstraction): Tập trung vào chi tiết quan trọng nhất.",
                 "• Thuật toán (Algorithm): Thiết kế các bước thực hiện tuần tự chính xác."
             ], accent_color=C_GREEN)

    # Slide 6: Quy định phòng máy UNIGO
    s6 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s6)
    add_header_banner(s6, grade, "V. HOẠT ĐỘNG 3: NỘI QUY & AN TOÀN PHÒNG MÁY TÍNH UNIGO")
    add_footer(s6)

    rules = [
        "1. Xếp hàng trật tự trước khi vào phòng máy. Để giày dép đúng nơi quy định.",
        "2. TUYỆT ĐỐI KHÔNG mang đồ ăn, nước uống vào phòng thực hành máy tính.",
        "3. Thao tác bật/tắt máy đúng quy trình. Không tự ý tháo dỡ, cắm rút dây điện.",
        "4. Ngồi đúng vị trí máy được phân công. Giữ vệ sinh chung ngăn nắp.",
        "5. Khi gặp sự cố máy tính hoặc nguồn điện: BÁO NGAY CHO GIÁO VIÊN HỖ TRỢ."
    ]
    add_card(s6, Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.1),
             title="⚠️ 5 NGUYÊN TẮC VÀNG PHÒNG THỰC HÀNH TIN HỌC UNIGO",
             items=rules, accent_color=C_ORANGE)

    # Slide 7: An toàn số & Đạo đức số
    s7 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s7)
    add_header_banner(s7, grade, "VI. AN TOÀN SỐ & ĐẠO ĐỨC SỐ (DIGITAL SAFETY & ETHICS)")
    add_footer(s7)

    add_card(s7, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="🛡️ An toàn thông tin cá nhân",
             items=[
                 "Giữ bí mật mật khẩu tài khoản cá nhân và tài khoản trường cấp.",
                 "Không tự ý cung cấp họ tên, số điện thoại, địa chỉ cho người lạ trên mạng.",
                 "Cảnh giác với các đường link lạ, file tải về không rõ nguồn gốc.",
                 "Đăng xuất tài khoản sau khi hoàn thành giờ học trên máy tính dùng chung."
             ], accent_color=C_NAVY)

    add_card(s7, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="🌐 Văn hóa & Đạo đức số",
             items=[
                 "Tôn trọng bạn bè: Không sử dụng ngôn từ thiếu văn hóa trên môi trường số.",
                 "Tôn trọng bản quyền: Tải tài liệu, hình ảnh phải ghi rõ nguồn trích dẫn.",
                 "Sử dụng thời gian hợp lý: Cân bằng thời gian dùng máy tính và vận động.",
                 "Chia sẻ năng lượng tích cực: Lan tỏa các nội dung học tập hữu ích."
             ], accent_color=C_AMBER)

    # Slide 8: Đánh giá & Học liệu
    s8 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s8)
    add_header_banner(s8, grade, "VII. CÁCH THỨC ĐÁNH GIÁ & ĐỒ DÙNG HỌC TẬP")
    add_footer(s8)

    add_card(s8, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="📊 Hình thức đánh giá môn học",
             items=[
                 "• Đánh giá thường xuyên: Nhận xét thái độ học tập, bài thực hành cá nhân.",
                 "• Đánh giá sản phẩm nhóm: Tiêu chí sáng tạo, kỹ năng làm việc nhóm.",
                 "• Đánh giá định kỳ 1 & 2 (Học kỳ 1): Bài kiểm tra lý thuyết & thực hành.",
                 "• Đánh giá định kỳ 3 & 4 (Học kỳ 2): Bài kiểm tra & Sản phẩm dự án cuối năm."
             ], accent_color=C_BLUE)

    add_card(s8, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.1),
             title="📚 Đồ dùng & Học liệu cần chuẩn bị",
             items=[
                 f"1. Sách giáo khoa Tin học {grade} (Bộ sách GDPT 2018).",
                 "2. Vở ghi bài & Sổ tay ghi chép cá nhân.",
                 "3. Tài khoản học tập trực tuyến UNIGO (Office 365 / LMS).",
                 "4. Thẻ nhớ / USB cá nhân (dùng cho các khối lớp THCS)."
             ], accent_color=C_GREEN)

    # Slide 9: Nhiệm vụ mở rộng
    s9 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s9)
    add_header_banner(s9, grade, "VIII. HOẠT ĐỘNG 4: NHIỆM VỤ MỞ RỘNG & KẾ HOẠCH CÁ NHÂN")
    add_footer(s9)

    add_card(s9, Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.1),
             title="📝 Nhiệm vụ thiết lập mục tiêu cá nhân",
             items=[
                 "Thảo luận nhóm 2 bạn cùng bàn:",
                 "1. Viết ra 2 mục tiêu em mong muốn đạt được nhất trong môn Tin học năm học này.",
                 "2. Đề xuất 1 ý tưởng sản phẩm số em muốn tự tay tạo ra (Ví dụ: Tranh vẽ, Bài trình chiếu, Trò chơi Scratch...).",
                 "3. Cam kết thực hiện đúng 5 nguyên tắc vàng phòng máy tính UNIGO."
             ], accent_color=C_NAVY)

    # Slide 10: Tóm tắt & Dặn dò
    s10 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s10)
    add_header_banner(s10, grade, "IX. TỔNG KẾT & DẶN DÒ NHIỆM VỤ VỀ NHÀ")
    add_footer(s10)

    add_card(s10, Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.1),
             title="📌 3 Điểm ghi nhớ quan trọng",
             items=[
                 f"1. Môn Tin học {grade} mang lại kiến thức công nghệ hiện đại và kỹ năng thực hành thiết thực.",
                 "2. Luôn tuân thủ tuyệt đối quy định an toàn phòng máy và nguyên tắc ứng xử văn minh số.",
                 "3. Dặn dò nhiệm vụ về nhà: Xem trước Bài 1 trong SGK, chuẩn bị đầy đủ vở ghi và đồ dùng học tập cho tiết sau."
             ], accent_color=C_ORANGE)

    # Slide 11: Cảm ơn
    s11 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(s11, C_NAVY)
    
    tb_c = s11.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.33), Inches(3.5))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    
    p_c1 = tf_c.paragraphs[0]
    p_c1.alignment = PP_ALIGN.CENTER
    r_c1 = p_c1.add_run()
    r_c1.text = "CHÚC CÁC EM CÓ MỘT NĂM HỌC THÀNH CÔNG VÀ NHIỀU NIỀM VUI SÁNG TẠO SỐ!\n\n"
    r_c1.font.name = "Times New Roman"
    r_c1.font.size = Pt(26)
    r_c1.font.bold = True
    r_c1.font.color.rgb = C_AMBER

    p_c2 = tf_c.add_paragraph()
    p_c2.alignment = PP_ALIGN.CENTER
    r_c2 = p_c2.add_run()
    r_c2.text = "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO"
    r_c2.font.name = "Times New Roman"
    r_c2.font.size = Pt(22)
    r_c2.font.bold = True
    r_c2.font.color.rgb = C_WHITE

    add_footer(s11)

    return prs

def main():
    print("==================================================")
    print(" BẮT ĐẦU TẠO HỆ THỐNG SLIDE ĐỊNH HƯỚNG TIN HỌC (TIẾT 0)")
    print("==================================================")

    created_count = 0

    for grade in range(1, 9):
        print(f"\n---> Đang tạo Slide Tiết 0 môn Tin học Lớp {grade}...")
        prs = build_orientation_deck(grade)

        out_folder = os.path.join(BASE_OUT_DIR, f"Lớp_{grade}", "Tiết_00")
        os.makedirs(out_folder, exist_ok=True)

        filename = f"Slide_Tin_hoc_Lớp_{grade}_Tiet00_Dinh_huong_mon_hoc.pptx"
        out_filepath = os.path.join(out_folder, filename)

        try:
            prs.save(out_filepath)
            created_count += 1
            print(f"  [+] Đã tạo slide: {out_filepath}")
        except Exception as e:
            print(f"  [!] Lỗi khi lưu slide {out_filepath}: {e}")

    print(f"\n==================================================")
    print(f" HOÀN THÀNH TẠO {created_count} SLIDE ĐỊNH HƯỚNG MON TIN HỌC!")
    print(f"==================================================")

if __name__ == '__main__':
    main()
