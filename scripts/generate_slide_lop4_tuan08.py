# -*- coding: utf-8 -*-
"""
Tạo Slide Bài Giảng Chuẩn UNIGO: Lớp 4 - Tuần 8 - Tiết 7
Bài 4: Tìm kiếm thông tin trên Internet (Tiết 1)
Tuân thủ toàn bộ 13 quy tắc Anti-Bug và chuẩn mực Visual-First từ tao-slide-bai-giang/SKILL.md.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

TEMPLATE_PATH = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx'
OUTPUT_DIR    = r'D:\UNIGO\KHBD_Tin_học\Lớp_4\Tuần_08'
OUTPUT_FILE   = os.path.join(OUTPUT_DIR, 'Slide_Tin_hoc_Lớp_4_Tiet07_Bai_4_Tim_kiem_thong_tin_tren_Internet_Tiet_1.pptx')
IMG_DIR       = os.path.join(OUTPUT_DIR, 'images')

# Tọa độ Vùng An Toàn (Safe Zone)
SAFE_TOP    = 1.15
SAFE_BOTTOM = 6.35
SAFE_LEFT   = 0.35
SAFE_RIGHT  = 13.00
SLIDE_W     = 13.33
SLIDE_H     = 7.50

# Bảng màu Indigo - Deep Ocean cao cấp cho Lớp 4
C_PRIMARY       = "1E3A8A"  # Deep Blue / Indigo đậm
C_ACCENT        = "2563EB"  # Royal Blue tươi
C_ACCENT_LIGHT  = "DBEAFE"  # Xanh pastel nhạt
C_BG            = "F8FAFC"  # Nền trang nhã
C_CARD          = "FFFFFF"  # Thẻ card trắng
C_BORDER        = "CBD5E1"  # Viền card xám bạc
C_TEXT_DARK     = "0F172A"  # Chữ đen than
C_TEXT_MUTED    = "475569"  # Chữ phụ xám xanh
C_TEXT_WHITE    = "FFFFFF"  # Chữ trắng
C_ORANGE        = "EA580C"  # Cam Hộp Kết Luận
C_ORANGE_BG     = "FFF7ED"  # Nền cam nhạt
C_GREEN         = "16A34A"  # Xanh lá thành công

def hex_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def set_font(run, size_pt, bold=False, color_hex=C_TEXT_DARK, font_name="Arial"):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = hex_rgb(color_hex)
    run.font.name = font_name

def add_safe_shape(slide, shape_type, left, top, width, height, fill_hex,
                    border_hex=None, border_width_pt=1, send_to_back=False):
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)

    shape = slide.shapes.add_shape(
        shape_type,
        Inches(left), Inches(actual_top), Inches(width), Inches(actual_height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_rgb(border_hex)
        shape.line.width = Pt(border_width_pt)
    else:
        shape.line.fill.background()

    if send_to_back:
        sp = shape._element
        spTree = sp.getparent()
        spTree.remove(sp)
        spTree.insert(2, sp)
    return shape

def add_textbox(slide, left, top, width, height, text, size_pt=18, bold=False,
                color_hex=C_TEXT_DARK, alignment=PP_ALIGN.LEFT, font_name="Arial"):
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.2)

    txbox = slide.shapes.add_textbox(Inches(left), Inches(actual_top), Inches(width), Inches(actual_height))
    tf = txbox.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    set_font(run, size_pt, bold, color_hex, font_name)
    return txbox

def add_picture_safe(slide, img_path, left, top, width, height):
    if not os.path.isfile(img_path):
        return None
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.3)
    try:
        return slide.shapes.add_picture(img_path, Inches(left), Inches(actual_top),
                                         Inches(width), Inches(actual_height))
    except Exception as e:
        print(f"Error adding picture {img_path}: {e}")
        return None

def add_badge(slide, text, left, top=1.20, width=2.8, height=0.36, fill_hex=C_PRIMARY, text_hex=C_TEXT_WHITE):
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, fill_hex=fill_hex)
    add_textbox(slide, left, top, width, height, text, size_pt=12, bold=True, color_hex=text_hex, alignment=PP_ALIGN.CENTER)

def add_slide_header(slide, badge_text, title_text, subtitle_text=None):
    add_badge(slide, badge_text, left=0.45, top=1.20, width=len(badge_text)*0.13 + 0.6)
    # Title
    add_textbox(slide, 0.45, 1.62, 12.4, 0.45, title_text, size_pt=23, bold=True, color_hex=C_PRIMARY)
    if subtitle_text:
        add_textbox(slide, 0.45, 2.05, 12.4, 0.30, subtitle_text, size_pt=14, bold=False, color_hex=C_TEXT_MUTED)

def add_slide_transition(slide, transition_type="fade"):
    transitions = {
        'fade': '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>',
        'push': '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:push dir="l"/></p:transition>',
    }
    xml_str = transitions.get(transition_type, transitions['fade'])
    slide._element.append(etree.fromstring(xml_str))

def build_deck():
    prs = Presentation(TEMPLATE_PATH)
    blank_layout = prs.slide_layouts[6]

    # Xóa slide mẫu cũ nếu có
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # =========================================================================
    # SLIDE 1: TRANG BÌA (Cover Slide - Layout Hero Banner)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s1, "fade")
    # Background panel an toàn
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.25, 12.43, 5.00, fill_hex=C_BG, border_hex=C_BORDER, send_to_back=True)

    # Cột trái: Tiêu đề & Giới thiệu
    add_badge(s1, "TIN HỌC LỚP 4 • TUẦN 8 • TIẾT 7", left=0.85, top=1.60, width=3.4, height=0.40, fill_hex=C_PRIMARY)
    add_textbox(s1, 0.85, 2.25, 6.0, 1.60, "TÌM KIẾM THÔNG TIN\nTRÊN INTERNET", size_pt=28, bold=True, color_hex=C_PRIMARY)
    add_textbox(s1, 0.85, 3.90, 6.0, 0.60, "Bài 4 (Tiết 1) • Chủ đề 3: Tìm kiếm và trao đổi thông tin\nBộ sách: Kết nối tri thức với cuộc sống", size_pt=15, bold=False, color_hex=C_TEXT_MUTED)

    # 3 Bullet điểm nhấn nhỏ
    add_textbox(s1, 0.85, 4.60, 5.8, 1.20,
                "🔑 Hiểu khái niệm Từ khóa (Keyword) & Máy tìm kiếm\n"
                "🚀 Nắm chắc Quy trình 4 bước tìm kiếm bằng Google\n"
                "🛡️ Thực hành an toàn và tìm thông tin hiệu quả",
                size_pt=14, bold=False, color_hex=C_TEXT_DARK)

    # Cột phải: Hero Image sinh động
    cover_img = os.path.join(IMG_DIR, 'cover_hero.jpg')
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 7.10, 1.55, 5.45, 4.40, fill_hex=C_CARD, border_hex=C_ACCENT, border_width_pt=2)
    add_picture_safe(s1, cover_img, 7.18, 1.63, 5.29, 4.24)

    # =========================================================================
    # SLIDE 2: KHỞI ĐỘNG (Tình huống SGK tr.20 - Khoa & Minh)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s2, "push")
    add_slide_header(s2, "KHỞI ĐỘNG 🔍", "Tình huống: Giúp bạn Khoa và Minh tìm tài liệu Vua Hùng",
                     "Cùng quan sát đoạn hội thoại mở đầu trong SGK Tin học 4 (trang 20)")

    # Cột trái: Nội dung tình huống
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_textbox(s2, 0.65, 2.55, 5.40, 0.40, "💬 ĐỌC TÌNH HUỐNG DẪN DẮT:", size_pt=16, bold=True, color_hex=C_PRIMARY)
    add_textbox(s2, 0.65, 2.95, 5.40, 1.80,
                "• Bạn Khoa và Minh đang cần tìm thông tin về các đời Vua Hùng để hoàn thành bài tập nhóm môn Lịch sử & Địa lí.\n\n"
                "• Khoa gợi ý: Cần xác định chủ đề trước, rồi chọn từ khóa phù hợp để tìm kiếm trên Internet.",
                size_pt=15, bold=False, color_hex=C_TEXT_DARK)

    # Box câu hỏi tương tác màu cam
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 4.85, 5.40, 1.25, fill_hex=C_ORANGE_BG, border_hex=C_ORANGE)
    add_textbox(s2, 0.80, 4.95, 5.10, 1.05,
                "❓ CÂU HỎI THẢO LUẬN:\n"
                "Theo em, hai bạn đang cần tìm thông tin theo chủ đề gì?\n"
                "Em hãy gợi ý từ khóa ngắn gọn để hai bạn tìm kiếm nhanh nhất?",
                size_pt=14, bold=True, color_hex=C_ORANGE)

    # Cột phải: Tranh SGK trích xuất
    sgk_p20 = os.path.join(IMG_DIR, 'sgk_p20_khoa_minh_hoi_thoai.png')
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s2, sgk_p20, 6.55, 2.50, 6.23, 3.65)

    # =========================================================================
    # SLIDE 3: KHÁM PHÁ 1 - TỪ KHÓA LÀ GÌ? (Keyword)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s3, "fade")
    add_slide_header(s3, "KHÁM PHÁ KIẾN THỨC 🔑", "1. Từ khóa (Keyword) là gì?",
                     "Khái niệm cốt lõi giúp học sinh tìm kiếm chính xác điều mình mong muốn")

    # Cột trái: 3 Thẻ bài học
    # Card 1: Khái niệm
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 6.60, 1.70, fill_hex=C_CARD, border_hex=C_ACCENT, border_width_pt=2)
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 0.20, 1.70, fill_hex=C_ACCENT)
    add_textbox(s3, 0.85, 2.55, 6.00, 0.35, "📌 Khái niệm Từ khóa (Keyword):", size_pt=17, bold=True, color_hex=C_PRIMARY)
    add_textbox(s3, 0.85, 2.95, 6.00, 1.00,
                "Từ khóa là từ hoặc cụm từ ngắn gọn thể hiện nội dung thông tin chúng ta muốn tìm kiếm trên môi trường số.",
                size_pt=16, color_hex=C_TEXT_DARK)

    # Card 2: Ví dụ SGK
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 4.30, 6.60, 1.95, fill_hex=C_CARD, border_hex=C_BORDER)
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 4.30, 0.20, 1.95, fill_hex=C_ORANGE)
    add_textbox(s3, 0.85, 4.45, 6.00, 0.35, "💡 Ví dụ lựa chọn từ khóa:", size_pt=17, bold=True, color_hex=C_ORANGE)
    add_textbox(s3, 0.85, 4.85, 6.00, 1.30,
                "• Chủ đề: Tìm hiểu về các vị vua thời Hùng Vương\n"
                "• Từ khóa gợi ý: \"Các đời vua Hùng\" hoặc \"Vua Hùng\"\n"
                "👉 Chú ý: Từ khóa càng rõ ràng, kết quả tìm kiếm càng chính xác!",
                size_pt=15, color_hex=C_TEXT_DARK)

    # Cột phải: Ảnh 3D Chiếc chìa khóa tri thức
    key_img = os.path.join(IMG_DIR, 'keyword_treasure_key.jpg')
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 7.30, 2.40, 5.58, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s3, key_img, 7.40, 2.50, 5.38, 3.65)

    # =========================================================================
    # SLIDE 4: KHÁM PHÁ 2 - MÁY TÌM KIẾM LÀ GÌ? (Search Engine)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s4, "fade")
    add_slide_header(s4, "CÔNG CỤ TÌM KIẾM 🌐", "2. Máy tìm kiếm (Search Engine) là gì?",
                     "Những trang web chuyên dụng giúp em khám phá kho tàng tri thức trên Internet")

    # Cột trái: Định nghĩa + Giao diện tìm kiếm SGK tr.21
    add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 1.30, fill_hex=C_ACCENT_LIGHT, border_hex=C_ACCENT)
    add_textbox(s4, 0.65, 2.48, 5.40, 1.15,
                "🌟 ĐỊNH NGHĨA MÁY TÌM KIẾM:\n"
                "Máy tìm kiếm là trang web đặc biệt giúp người sử dụng tìm kiếm thông tin trên Internet nhanh chóng bằng từ khóa.",
                size_pt=15, bold=True, color_hex=C_PRIMARY)

    # Ảnh giao diện Google search box SGK
    sgk_p21 = os.path.join(IMG_DIR, 'sgk_p21_google_search_ui.png')
    add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 3.85, 5.80, 2.40, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s4, sgk_p21, 0.55, 3.95, 5.60, 2.20)

    # Cột phải: 3 Thẻ máy tìm kiếm phổ biến
    engines = [
        ("1. Google (google.com)", "Phổ biến & thông dụng nhất toàn cầu 🌍", C_PRIMARY),
        ("2. Bing (bing.com)", "Máy tìm kiếm thông minh của Microsoft 💻", C_ACCENT),
        ("3. Cốc Cốc (coccoc.com)", "Trình duyệt & máy tìm kiếm tối ưu cho người Việt 🇻🇳", C_GREEN),
    ]

    card_h = 1.15
    card_gap = 0.18
    for i, (name, desc, col) in enumerate(engines):
        cy = 2.40 + i * (card_h + card_gap)
        add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 6.55, cy, 6.33, card_h, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 6.55, cy, 0.20, card_h, fill_hex=col)
        add_textbox(s4, 6.95, cy + 0.15, 5.80, 0.35, name, size_pt=17, bold=True, color_hex=col)
        add_textbox(s4, 6.95, cy + 0.55, 5.80, 0.50, desc, size_pt=14, color_hex=C_TEXT_DARK)

    # =========================================================================
    # SLIDE 5: QUY TRÌNH 4 BƯỚC TÌM KIẾM BẰNG GOOGLE (Quy trình chuẩn)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s5, "push")
    add_slide_header(s5, "QUY TRÌNH THỰC HIỆN 🚀", "3. Quy trình 4 bước tìm kiếm thông tin bằng Google",
                     "Ghi nhớ các bước tuần tự để thao tác thành thạo trên phòng máy")

    # 4 Thẻ bước nằm ngang
    steps = [
        ("BƯỚC 1", "Xác định từ khóa", "Xác định rõ ràng từ khóa về nội dung em cần tìm kiếm.", C_PRIMARY),
        ("BƯỚC 2", "Mở máy tìm kiếm", "Mở trình duyệt web và gõ địa chỉ: google.com", C_ACCENT),
        ("BƯỚC 3", "Gõ từ khóa & Enter", "Gõ từ khóa vào ô tìm kiếm, sau đó nhấn phím Enter.", C_ORANGE),
        ("BƯỚC 4", "Chọn siêu liên kết", "Nháy chuột vào liên kết thích hợp trong danh sách kết quả.", C_GREEN),
    ]

    col_w = 2.85
    col_gap = 0.34
    for i, (b_name, b_title, b_desc, b_col) in enumerate(steps):
        bx = 0.45 + i * (col_w + col_gap)
        # Card
        add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, bx, 2.40, col_w, 2.45, fill_hex=C_CARD, border_hex=b_col, border_width_pt=2)
        # Top banner
        add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, bx, 2.40, col_w, 0.48, fill_hex=b_col)
        add_textbox(s5, bx, 2.46, col_w, 0.35, b_name, size_pt=15, bold=True, color_hex=C_TEXT_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s5, bx + 0.15, 3.00, col_w - 0.30, 0.50, b_title, size_pt=16, bold=True, color_hex=b_col, alignment=PP_ALIGN.CENTER)
        add_textbox(s5, bx + 0.15, 3.60, col_w - 0.30, 1.15, b_desc, size_pt=14, color_hex=C_TEXT_DARK, alignment=PP_ALIGN.CENTER)

    # Sơ đồ trực quan bên dưới
    steps_diag = os.path.join(IMG_DIR, 'search_steps_diagram_vn.png')
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 5.05, 12.43, 1.20, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s5, steps_diag, 0.65, 5.10, 12.03, 1.10)

    # =========================================================================
    # SLIDE 6: THỰC HÀNH MẪU - CÙNG TÌM HIỂU HỒ GƯƠM (SGK tr.22-23)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s6, "fade")
    add_slide_header(s6, "THỰC HÀNH MẪU 🌊", "4. Thực hành mẫu: Tìm kiếm thông tin về Hồ Gươm",
                     "Quan sát các thao tác thực tế và trang kết quả tìm kiếm Google (SGK tr.22 - 23)")

    # Cột trái: Hướng dẫn chi tiết
    add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.00, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_textbox(s6, 0.65, 2.55, 4.60, 0.35, "🎯 CÁC BƯỚC THỰC HÀNH:", size_pt=16, bold=True, color_hex=C_PRIMARY)
    add_textbox(s6, 0.65, 2.95, 4.60, 2.40,
                "1. Từ khóa: \"Hồ Gươm\"\n\n"
                "2. Truy cập: google.com\n\n"
                "3. Kết quả trả về: Hơn 5.000.000 trang web chỉ trong 0,47 giây!\n\n"
                "4. Xem chi tiết: Nháy chuột vào bài viết của vietnamtourism.gov.vn (trang chính thức của Tổng cục Du lịch).",
                size_pt=15, color_hex=C_TEXT_DARK)

    add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 5.40, 4.60, 0.70, fill_hex=C_ACCENT_LIGHT, border_hex=C_ACCENT)
    add_textbox(s6, 0.75, 5.45, 4.40, 0.60, "💡 Mẹo nhỏ: Ưu tiên chọn các trang có đuôi .gov.vn hoặc .edu.vn để có thông tin tin cậy!", size_pt=13, bold=True, color_hex=C_PRIMARY)

    # Cột phải: Hình 13 & 14 SGK
    sgk_p22 = os.path.join(IMG_DIR, 'sgk_p22_hinh13_ket_qua_ho_guom.png')
    add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 5.65, 2.40, 7.23, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s6, sgk_p22, 5.75, 2.50, 7.03, 3.65)

    # =========================================================================
    # SLIDE 7: LUYỆN TẬP 1 - CHỌN TỪ KHÓA (SGK tr.21)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s7, "push")
    add_slide_header(s7, "LUYỆN TẬP 🎯", "Bài tập 1: Thử tài chọn từ khóa chính xác nhất",
                     "Đọc kỹ yêu cầu và cùng thảo luận nhóm đôi để đưa ra phương án đúng nhất")

    # Khung câu hỏi trên
    add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 8.80, 0.95, fill_hex=C_ACCENT_LIGHT, border_hex=C_PRIMARY, border_width_pt=2)
    add_textbox(s7, 0.65, 2.48, 8.40, 0.75,
                "❓ CÂU HỎI (SGK Trang 21):\n"
                "Nếu muốn tìm kiếm thông tin về Bảo tàng Dân tộc học Việt Nam, từ khóa nào sau đây là phù hợp nhất?",
                size_pt=16, bold=True, color_hex=C_PRIMARY)

    # Ảnh Bảo tàng Dân tộc học bên phải câu hỏi
    museum_img = os.path.join(IMG_DIR, 'museum_icon.jpg')
    add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 9.45, 2.40, 3.43, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s7, museum_img, 9.55, 2.50, 3.23, 3.65)

    # 4 đáp án bên trái
    ans = [
        ("A. Bảo tàng", "Chưa phù hợp: Quá rộng, ra hàng nghìn bảo tàng khác nhau ❌", C_CARD, C_BORDER, C_TEXT_DARK),
        ("B. Bảo tàng Dân tộc học Việt Nam", "ĐÁP ÁN CHÍNH XÁC NHẤT: Đầy đủ, rõ ràng và đúng trọng tâm! 🏆 ✅", C_ORANGE_BG, C_ORANGE, C_ORANGE),
        ("C. Bảo tàng Việt Nam", "Chưa phù hợp: Không chỉ rõ bảo tàng dân tộc học nào ❌", C_CARD, C_BORDER, C_TEXT_DARK),
        ("D. Các dân tộc ở Việt Nam", "Chưa phù hợp: Sẽ tìm về con người, văn hóa thay vì bảo tàng ❌", C_CARD, C_BORDER, C_TEXT_DARK),
    ]

    for i, (opt_title, opt_desc, bg_col, bdr_col, txt_col) in enumerate(ans):
        ay = 3.50 + i * 0.70
        add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, ay, 8.80, 0.62, fill_hex=bg_col, border_hex=bdr_col, border_width_pt=(2 if i == 1 else 1))
        add_textbox(s7, 0.65, ay + 0.05, 8.40, 0.28, opt_title, size_pt=15, bold=True, color_hex=txt_col)
        add_textbox(s7, 0.65, ay + 0.32, 8.40, 0.25, opt_desc, size_pt=13, bold=(i == 1), color_hex=txt_col)

    # =========================================================================
    # SLIDE 8: LUYỆN TẬP 2 - GHÉP NỐI CHỦ ĐỀ & TỪ KHÓA (SGK tr.23)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s8, "fade")
    add_slide_header(s8, "TRÒ CHƠI GHÉP NỐI 🧩", "Bài tập 2: Ghép nối Chủ đề & Từ khóa tương ứng",
                     "Cùng nối chủ đề bài học ở Cột A với từ khóa tìm kiếm phù hợp nhất ở Cột B")

    matches = [
        ("Chủ đề 1: Tìm hiểu các vua thời Hùng Vương", "Từ khóa: \"Các đời vua Hùng\"", C_PRIMARY),
        ("Chủ đề 2: Tìm hiểu danh lam thắng cảnh Hà Nội", "Từ khóa: \"Hồ Gươm Hà Nội\"", C_ACCENT),
        ("Chủ đề 3: Tìm hiểu vai trò của không khí", "Từ khóa: \"Vai trò của không khí\"", C_GREEN),
    ]

    # Cột trái 3 cặp thẻ (chiếm 8.0 in)
    for i, (left_txt, right_txt, col) in enumerate(matches):
        my = 2.45 + i * 1.25
        # Cột A
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, my, 3.80, 1.05, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_textbox(s8, 0.60, my + 0.20, 3.50, 0.65, left_txt, size_pt=15, bold=True, color_hex=col)

        # Mũi tên kết nối
        add_textbox(s8, 4.30, my + 0.25, 0.80, 0.55, "──►", size_pt=20, bold=True, color_hex=col, alignment=PP_ALIGN.CENTER)

        # Cột B
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 5.15, my, 3.60, 1.05, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_textbox(s8, 5.30, my + 0.25, 3.30, 0.55, right_txt, size_pt=16, bold=True, color_hex=col)

    # Cột phải: Hình SGK minh họa trang web Hồ Gươm
    sgk_p23 = os.path.join(IMG_DIR, 'sgk_p23_hinh14_trang_web_ho_guom.png')
    add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 8.95, 2.40, 3.93, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s8, sgk_p23, 9.05, 2.50, 3.73, 3.65)

    # =========================================================================
    # SLIDE 9: THỬ THÁCH VẬN DỤNG & AN TOÀN SỐ
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s9, "push")
    add_slide_header(s9, "VẬN DỤNG & AN TOÀN 🛡️", "5. Thử thách: Em là nhà thám hiểm Internet an toàn",
                     "Vận dụng kiến thức để tìm kiếm địa danh yêu thích và tuân thủ nguyên tắc an toàn")

    # Cột trái: Nhiệm vụ & Quy tắc vàng
    add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 7.50, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_textbox(s9, 0.70, 2.55, 7.00, 0.35, "🎯 NHIỆM VỤ DÀNH CHO EM:", size_pt=17, bold=True, color_hex=C_PRIMARY)
    add_textbox(s9, 0.70, 2.95, 7.00, 1.10,
                "Em hãy chọn một danh lam thắng cảnh ở quê hương em (hoặc \"Văn Miếu - Quốc Tử Giám\"), hãy:\n"
                "  1. Xác định từ khóa tìm kiếm phù hợp.\n"
                "  2. Sử dụng Google để tra cứu và ghi lại 2 thông tin thú vị em tìm được!",
                size_pt=15, color_hex=C_TEXT_DARK)

    # Hộp nguyên tắc an toàn
    add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, 0.70, 4.25, 7.00, 1.80, fill_hex=C_ACCENT_LIGHT, border_hex=C_PRIMARY)
    add_textbox(s9, 0.90, 4.35, 6.60, 0.35, "🛡️ QUY TẮC VÀNG KHI TÌM KIẾM TRÊN MẠNG:", size_pt=16, bold=True, color_hex=C_PRIMARY)
    add_textbox(s9, 0.90, 4.75, 6.60, 1.20,
                "✅ Chỉ truy cập các trang web có nội dung lành mạnh, phục vụ học tập.\n"
                "❌ Tuyệt đối không bấm vào các đường link lạ, quảng cáo nhấp nháy.\n"
                "🤝 Báo ngay với thầy cô giáo hoặc bố mẹ khi gặp nội dung không an toàn.",
                size_pt=14, bold=False, color_hex=C_TEXT_DARK)

    # Cột phải: Mascot an toàn mạng
    shield_img = os.path.join(IMG_DIR, 'safety_shield_mascot.jpg')
    add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, 8.15, 2.40, 4.73, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s9, shield_img, 8.25, 2.50, 4.53, 3.65)

    # =========================================================================
    # SLIDE 10: TỔNG KẾT & GHI NHỚ (Hộp Kết luận Cam Chuẩn UNIGO)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s10, "fade")
    add_slide_header(s10, "GHI NHỚ TRỌNG TÂM 💡", "Tổng kết kiến thức em cần ghi nhớ sau bài học",
                      "Những nội dung cốt lõi của Tiết 1 - Bài 4 SGK Tin học 4")

    # Hộp KẾT LUẬN Cam Chuẩn UNIGO (Layout 14)
    add_safe_shape(s10, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 12.43, 3.85, fill_hex=C_ORANGE_BG, border_hex=C_ORANGE, border_width_pt=2)
    # Header box cam
    add_safe_shape(s10, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 12.43, 0.60, fill_hex=C_ORANGE)
    add_textbox(s10, 0.65, 2.48, 12.00, 0.40, "⭐ EM CẦN GHI NHỚ TRỌNG TÂM:", size_pt=18, bold=True, color_hex=C_TEXT_WHITE, alignment=PP_ALIGN.CENTER)

    # 4 Gạch đầu dòng ghi nhớ bên trái (chiếm 8.6 in)
    add_textbox(s10, 0.75, 3.15, 8.40, 2.95,
                "1. Máy tìm kiếm (Google, Bing, Cốc Cốc...) giúp chúng ta tìm kiếm thông tin nhanh chóng và thuận tiện trên Internet.\n\n"
                "2. Từ khóa (Keyword) là từ hoặc cụm từ ngắn gọn, chính xác thể hiện nội dung cần tìm — là chiếc chìa khóa quyết định kết quả.\n\n"
                "3. Quy trình 4 bước tìm kiếm: (1) Xác định từ khóa ➔ (2) Mở google.com ➔ (3) Gõ từ khóa & Enter ➔ (4) Chọn siêu liên kết phù hợp.\n\n"
                "4. An toàn số: Luôn cẩn thận khi lựa chọn nguồn tin và không truy cập vào các trang web không an toàn.",
                size_pt=15, bold=False, color_hex=C_TEXT_DARK)

    # Cột phải: Chú cú thông minh cầm bóng đèn
    owl_img = os.path.join(IMG_DIR, 'smart_owl.jpg')
    add_safe_shape(s10, MSO_SHAPE.ROUNDED_RECTANGLE, 9.35, 3.15, 3.35, 2.95, fill_hex=C_CARD, border_hex=C_ORANGE)
    add_picture_safe(s10, owl_img, 9.45, 3.25, 3.15, 2.75)

    # =========================================================================
    # SLIDE 11: LỜI CHÀO & DẶN DÒ (Thank you slide)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s11, "fade")
    add_slide_header(s11, "KẾT THÚC TIẾT HỌC 🌟", "Chúc mừng các em đã hoàn thành xuất sắc tiết học!",
                      "Cùng chuẩn bị cho Tiết 2: Thực hành tìm kiếm thông tin nâng cao")

    # Cột trái: Lời khen & Dặn dò về nhà
    add_safe_shape(s11, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 7.50, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_textbox(s11, 0.70, 2.60, 7.00, 0.45, "🎉 CÁC EM ĐÃ LÀ NHÀ THÁM HIỂM TÀI BA!", size_pt=20, bold=True, color_hex=C_PRIMARY)
    add_textbox(s11, 0.70, 3.15, 7.00, 0.90,
                "Hôm nay các em đã nắm vững cách chọn từ khóa chuẩn xác và làm chủ 4 bước tìm kiếm thông tin trên Internet!",
                size_pt=16, color_hex=C_TEXT_DARK)

    add_safe_shape(s11, MSO_SHAPE.ROUNDED_RECTANGLE, 0.70, 4.25, 7.00, 1.80, fill_hex=C_BG, border_hex=C_ACCENT)
    add_textbox(s11, 0.90, 4.35, 6.60, 0.35, "📝 DẶN DÒ VỀ NHÀ:", size_pt=16, bold=True, color_hex=C_PRIMARY)
    add_textbox(s11, 0.90, 4.75, 6.60, 1.20,
                "• Ôn lại 4 bước tìm kiếm thông tin bằng Google.\n"
                "• Cùng bố mẹ thực hành tìm kiếm về 1 danh lam thắng cảnh yêu thích.\n"
                "• Đọc trước nội dung Tiết 2 trong SGK Tin học 4.",
                size_pt=15, color_hex=C_TEXT_DARK)

    # Cột phải: Mascot ăn mừng cúp vàng
    trophy_img = os.path.join(IMG_DIR, 'mascot_celebrate_trophy.jpg')
    add_safe_shape(s11, MSO_SHAPE.ROUNDED_RECTANGLE, 8.15, 2.40, 4.73, 3.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_picture_safe(s11, trophy_img, 8.25, 2.50, 4.53, 3.65)

    # Lưu file
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    prs.save(OUTPUT_FILE)
    print(f"ĐÃ TẠO THÀNH CÔNG: {OUTPUT_FILE}")
    print(f"Tổng số slide: {len(prs.slides)}")

if __name__ == '__main__':
    build_deck()
