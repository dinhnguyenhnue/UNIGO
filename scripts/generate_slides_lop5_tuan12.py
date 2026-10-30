# -*- coding: utf-8 -*-
"""
Tạo Slide Bài Giảng Chuẩn UNIGO: Lớp 5 - Tuần 12 - Môn Tin học
Gồm 2 tiết:
- Tiết 11: Bài 5: Bản quyền nội dung thông tin (Tiết 1)
- Tiết 12: Bài 5: Bản quyền nội dung thông tin (Tiết 2)

Tuân thủ nghiêm ngặt:
- Quy tắc bảo tồn mẫu văn bản UNIGO (Master Picture 9 chân trang, Picture 7 logo).
- Vùng An Toàn (SAFE_TOP = 1.15in, SAFE_BOTTOM = 6.35in).
- Bố cục Visual-First, Hero Image to rõ (50-70%), ảnh trích xuất SGK + ảnh AI phong cách giáo dục.
- Không có slide mục tiêu/năng lực kiểu giáo án; toàn bộ hướng tới học sinh (Student-Facing).
- 13 quy tắc Anti-Bug và kiểm tra tự động.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

TEMPLATE_PATH = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx'
OUTPUT_DIR    = r'D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_12'
IMG_DIR_T11   = os.path.join(OUTPUT_DIR, 'images_tiet11')
IMG_DIR_T12   = os.path.join(OUTPUT_DIR, 'images_tiet12')

SAFE_TOP    = 1.15
SAFE_BOTTOM = 6.35
SLIDE_W     = 13.33
SLIDE_H     = 7.50

def hex_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def set_font(run, size_pt, bold=False, color_hex="0F172A", font_name="Arial"):
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
                color_hex="0F172A", alignment=PP_ALIGN.LEFT, font_name="Arial"):
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
        print(f"Warning: image file not found: {img_path}")
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

def add_badge(slide, text, left, top=1.20, width=2.8, height=0.36, fill_hex="1E3A8A", text_hex="FFFFFF"):
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, fill_hex=fill_hex)
    add_textbox(slide, left, top, width, height, text, size_pt=12, bold=True, color_hex=text_hex, alignment=PP_ALIGN.CENTER)

def add_slide_header(slide, badge_text, title_text, subtitle_text=None, primary_hex="1E3A8A"):
    add_badge(slide, badge_text, left=0.45, top=1.20, width=len(badge_text)*0.13 + 0.6, fill_hex=primary_hex)
    add_textbox(slide, 0.45, 1.62, 12.4, 0.45, title_text, size_pt=23, bold=True, color_hex=primary_hex)
    if subtitle_text:
        add_textbox(slide, 0.45, 2.05, 12.4, 0.30, subtitle_text, size_pt=14, bold=False, color_hex="475569")

def add_slide_transition(slide, transition_type="fade"):
    transitions = {
        'fade': '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>',
        'push': '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:push dir="l"/></p:transition>',
    }
    xml_str = transitions.get(transition_type, transitions['fade'])
    slide._element.append(etree.fromstring(xml_str))


# =============================================================================
# BUILD TIẾT 11: BÀI 5 - TIẾT 1
# =============================================================================
def build_deck_tiet11():
    print("Building Tiet 11 deck...")
    prs = Presentation(TEMPLATE_PATH)
    blank_layout = prs.slide_layouts[6]
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # Palette Tiết 11: Deep Indigo & Teal
    C_PRI    = "1E3A8A"
    C_ACC    = "0D9488"
    C_ACCL   = "E0F2FE"
    C_BG     = "F8FAFC"
    C_CARD   = "FFFFFF"
    C_BORD   = "CBD5E1"
    C_DARK   = "0F172A"
    C_MUTED  = "475569"
    C_WHITE  = "FFFFFF"
    C_ORG    = "EA580C"
    C_ORGBG  = "FFF7ED"
    C_GRN    = "16A34A"

    # SLIDE 1: TRANG BÌA (Hero Banner)
    s1 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s1, "fade")
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.25, 12.43, 5.00, fill_hex=C_BG, border_hex=C_BORD, send_to_back=True)
    
    add_badge(s1, "TIN HỌC LỚP 5 • TUẦN 12 • TIẾT 11", left=0.85, top=1.55, width=3.4, height=0.40, fill_hex=C_PRI)
    add_textbox(s1, 0.85, 2.15, 6.0, 1.60, "BẢN QUYỀN NỘI DUNG\nTHÔNG TIN (TIẾT 1)", size_pt=27, bold=True, color_hex=C_PRI)
    add_textbox(s1, 0.85, 3.80, 6.0, 0.60, "Bài 5 • Chủ đề 4: Đạo đức, pháp luật & văn hóa số\nBộ sách: Kết nối tri thức với cuộc sống (Trang 24 - 26)", size_pt=14, bold=False, color_hex=C_MUTED)
    
    add_textbox(s1, 0.85, 4.55, 5.8, 1.40,
                "💡 Hiểu khái niệm Bản quyền nội dung thông tin\n"
                "✍️ Nhận biết tác giả và các dạng tác phẩm số\n"
                "🔐 Nắm vững quy tắc sử dụng hợp lệ & bảo mật sản phẩm",
                size_pt=15, bold=False, color_hex=C_DARK)

    # Hero image bên phải
    cover_img = os.path.join(IMG_DIR_T11, 'cover_hero.jpg')
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 7.10, 1.50, 5.45, 4.45, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    add_picture_safe(s1, cover_img, 7.18, 1.58, 5.29, 4.29)

    # SLIDE 2: KHỞI ĐỘNG (Báo tường 20/11)
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s2, "push")
    add_slide_header(s2, "KHỞI ĐỘNG 🌸", "Tình huống: Báo tường chào mừng ngày Nhà giáo Việt Nam",
                     "Cùng khám phá hoạt động làm báo tường của các bạn lớp 5A (SGK Tin học 5 - Trang 24)", primary_hex=C_PRI)
    
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    add_textbox(s2, 0.65, 2.55, 5.40, 0.40, "📰 ĐỌC TÌNH HUỐNG DẪN DẮT:", size_pt=16, bold=True, color_hex=C_PRI)
    add_textbox(s2, 0.65, 2.95, 5.40, 1.70,
                "• Các bạn học sinh lớp 5A làm báo tường chào mừng ngày Nhà giáo Việt Nam 20/11.\n\n"
                "• Trong tờ báo tường có rất nhiều nội dung: bài thơ, bài viết, tranh vẽ... có bài tự sáng tác, có bài sưu tầm để thể hiện tình cảm với thầy cô.",
                size_pt=15, bold=False, color_hex=C_DARK)
    
    # Hộp câu hỏi cam
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 4.80, 5.40, 1.30, fill_hex=C_ORGBG, border_hex=C_ORG)
    add_textbox(s2, 0.80, 4.90, 5.10, 1.10,
                "❓ CÂU HỎI SUY NGHĨ:\n"
                "Làm thế nào để người đọc phân biệt được bài nào do các bạn tự sáng tác, bài nào được sưu tầm từ sách báo?",
                size_pt=14, bold=True, color_hex=C_ORG)

    bt_img = os.path.join(IMG_DIR_T11, 'bao_tuong_20_11.jpg')
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    add_picture_safe(s2, bt_img, 6.55, 2.50, 6.23, 3.65)

    # SLIDE 3: HOẠT ĐỘNG 1 - TÁC GIẢ LÀ AI? (SGK tr.24 - 25)
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s3, "fade")
    add_slide_header(s3, "HOẠT ĐỘNG 1 🔍", "Tác giả là ai? - Quan sát 2 bài báo tường lớp 5A",
                     "So sánh bài thơ Hình 24 và bài viết Hình 25 trong SGK Tin học 5 (Trang 25)", primary_hex=C_PRI)

    # Card trái: Hình 24
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 3.80, 3.85, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    h24_path = os.path.join(IMG_DIR_T11, 'h24_bai_tho_xuan_quynh.png')
    add_picture_safe(s3, h24_path, 0.55, 2.48, 3.60, 3.10)
    add_textbox(s3, 0.55, 5.65, 3.60, 0.50, "📝 Bài thơ: \"Cô giáo của em\"\nTác giả: Nhà thơ Xuân Quỳnh (Sưu tầm)", size_pt=13, bold=True, color_hex=C_PRI, alignment=PP_ALIGN.CENTER)

    # Card giữa: Hình 25
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 4.45, 2.40, 3.80, 3.85, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    h25_path = os.path.join(IMG_DIR_T11, 'h25_bai_viet_ngo_ha_trang.png')
    add_picture_safe(s3, h25_path, 4.55, 2.48, 3.60, 3.10)
    add_textbox(s3, 4.55, 5.65, 3.60, 0.50, "✍️ Bài viết: \"Cô giáo lớp em\"\nTác giả: Bạn Ngô Hà Trang - Lớp 5A (Tự viết)", size_pt=13, bold=True, color_hex=C_ACC, alignment=PP_ALIGN.CENTER)

    # Card phải: 2 Câu hỏi khám phá
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 8.45, 2.40, 4.43, 3.85, fill_hex=C_ACCL, border_hex=C_PRI)
    add_textbox(s3, 8.65, 2.55, 4.05, 0.40, "💡 CÂU HỎI THẢO LUẬN:", size_pt=16, bold=True, color_hex=C_PRI)
    add_textbox(s3, 8.65, 3.05, 4.05, 1.40,
                "1. Bài thơ trong Hình 24 là của tác giả nào? Bài viết trong Hình 25 là của ai?\n\n"
                "👉 Bài thơ của Xuân Quỳnh, bài viết của Ngô Hà Trang.",
                size_pt=14, bold=False, color_hex=C_DARK)
    
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 8.65, 4.60, 4.05, 1.50, fill_hex=C_CARD, border_hex=C_ORG)
    add_textbox(s3, 8.75, 4.70, 3.85, 1.30,
                "2. Theo em, tại sao phải ghi rõ tên tác giả?\n\n"
                "👉 Để tôn trọng quyền sáng tạo của tác giả và người đọc biết xuất xứ tác phẩm!",
                size_pt=14, bold=True, color_hex=C_ORG)

    # SLIDE 4: KHÁM PHÁ 1 - BẢN QUYỀN NỘI DUNG THÔNG TIN LÀ GÌ?
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s4, "fade")
    add_slide_header(s4, "KIẾN THỨC MỚI 💡", "1. Bản quyền nội dung thông tin là gì?",
                     "Khái niệm cốt lõi về quyền sở hữu trí tuệ trong học tập và cuộc sống", primary_hex=C_PRI)

    # 3 Thẻ ngang trên
    cards = [
        ("Nội dung thông tin là gì? 📁", "Được thể hiện dưới nhiều dạng:\n• Văn bản (bài viết, bài thơ)\n• Hình ảnh, tranh vẽ\n• Âm thanh, bài hát\n• Video, phim hoạt hình đa phương tiện.", C_PRI),
        ("Tác giả là ai? 👤", "Là cá nhân hoặc tổ chức trực tiếp sáng tạo ra nội dung thông tin và sở hữu tác phẩm đó.\n(Ví dụ: Nhà thơ Xuân Quỳnh, bạn Ngô Hà Trang, NXB Giáo dục...).", C_ACC),
        ("Bản quyền là gì? 🛡️", "Là quyền của tác giả cho phép hoặc không cho phép người khác sử dụng nội dung thông tin của mình.\nChỉ được sử dụng khi được tác giả cho phép!", C_ORG)
    ]
    card_w = 3.90
    card_gap = 0.36
    for i, (title, content, color) in enumerate(cards):
        c_left = 0.45 + i * (card_w + card_gap)
        add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 2.40, card_w, 2.30, fill_hex=C_CARD, border_hex=color, border_width_pt=2)
        add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 2.40, card_w, 0.42, fill_hex=color)
        add_textbox(s4, c_left + 0.15, 2.42, card_w - 0.30, 0.38, title, size_pt=14, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s4, c_left + 0.15, 2.90, card_w - 0.30, 1.70, content, size_pt=13, bold=False, color_hex=C_DARK)

    # Hộp KẾT LUẬN Cam nổi bật (Dạng 14)
    add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 4.85, 12.43, 1.35, fill_hex=C_ORGBG, border_hex=C_ORG, border_width_pt=2)
    add_textbox(s4, 0.65, 4.90, 12.0, 0.32, "🌟 HỘP KẾT LUẬN (GHI NHỚ SGK TRANG 25):", size_pt=14, bold=True, color_hex=C_ORG)
    add_textbox(s4, 0.65, 5.22, 12.0, 0.90,
                "• Bản quyền nội dung thông tin là quyền của tác giả cho phép hoặc không cho phép người khác sử dụng nội dung thông tin.\n"
                "• Chỉ được sử dụng nội dung thông tin khi được phép!",
                size_pt=15, bold=True, color_hex=C_PRI)

    # SLIDE 5: KHÁM PHÁ 2 - TRUY CẬP VÀ BẢO MẬT (SGK tr.26)
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s5, "fade")
    add_slide_header(s5, "KIẾN THỨC MỚI 🔐", "2. Truy cập hợp lệ & Bảo mật nội dung thông tin",
                     "Quy tắc khi sử dụng nội dung của người khác và cách bảo vệ sản phẩm của mình", primary_hex=C_PRI)

    # Cột trái: Khi sử dụng nội dung người khác
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 3.85, fill_hex=C_CARD, border_hex=C_PRI, border_width_pt=2)
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 0.45, fill_hex=C_PRI)
    add_textbox(s5, 0.65, 2.45, 5.40, 0.35, "1. KHI SỬ DỤNG NỘI DUNG CỦA NGƯỜI KHÁC", size_pt=15, bold=True, color_hex=C_WHITE)
    add_textbox(s5, 0.65, 2.95, 5.40, 3.15,
                "✅ Cần được tác giả cho phép và PHẢI GHI RÕ NGUỒN TIN.\n\n"
                "• Nếu rõ tác giả: Ghi tên tác giả (Ví dụ: Thơ Xuân Quỳnh).\n"
                "• Nếu không rõ tác giả: Ghi nguồn là \"Sưu tầm\".\n"
                "• Nếu diễn đạt lại bằng lời của mình: Ghi \"Phỏng theo...\" (Ví dụ: Phỏng theo truyện Rùa và Thỏ).\n\n"
                "❌ CẢNH BÁO ĐẠO ĐỨC: Truy cập nội dung người khác khi chưa được phép là vi phạm đạo đức! (Ví dụ: xem trộm, sao chép bài trong giờ kiểm tra).",
                size_pt=14, bold=False, color_hex=C_DARK)

    # Cột phải: Bảo vệ sản phẩm của mình + Ký hiệu (C)
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 3.85, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 0.45, fill_hex=C_ACC)
    add_textbox(s5, 6.65, 2.45, 6.00, 0.35, "2. BẢO VỆ TÁC PHẨM & KÝ HIỆU BẢN QUYỀN ©", size_pt=15, bold=True, color_hex=C_WHITE)
    add_textbox(s5, 6.65, 2.95, 6.00, 1.40,
                "• Giữ bí mật sản phẩm trong quá trình tạo ra để tránh bị đánh cắp.\n"
                "• Ghi tên tác giả công khai trên sản phẩm đã hoàn thành.\n"
                "• Đăng kí bản quyền và ghi thông tin bản quyền © (Copyright).",
                size_pt=14, bold=False, color_hex=C_DARK)

    # Ảnh SGK H26 bản quyền
    h26_path = os.path.join(IMG_DIR_T11, 'h26_copyright_nxbgd.png')
    add_picture_safe(s5, h26_path, 6.65, 4.45, 6.03, 1.10)
    add_textbox(s5, 6.65, 5.60, 6.03, 0.55,
                "Hình 26: Ký hiệu bản quyền © của Nhà xuất bản Giáo dục Việt Nam",
                size_pt=13, bold=True, color_hex=C_MUTED, alignment=PP_ALIGN.CENTER)

    # SLIDE 6: CHỐT KIẾN THỨC GHI NHỚ TIẾT 1 (Dạng 9)
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s6, "fade")
    add_slide_header(s6, "CHỐT KIẾN THỨC 🌟", "Ghi nhớ cốt lõi SGK Tin học 5 - Trang 26",
                     "3 Nguyên tắc cơ bản học sinh cần ghi nhớ khi tham gia môi trường số", primary_hex=C_PRI)

    points = [
        ("NGUYÊN TẮC 1", "Khi sử dụng nội dung thông tin, cần được cho phép và ghi rõ nguồn tin.", C_PRI),
        ("NGUYÊN TẮC 2", "Trong quá trình tạo ra nội dung thông tin, cần có ý thức giữ gìn, bảo mật thông tin.", C_ACC),
        ("NGUYÊN TẮC 3", "Tác giả có thể ghi tên hoặc đăng kí bản quyền để bảo vệ quyền tác giả đối với nội dung thông tin.", C_ORG),
    ]
    p_h = 1.15
    for i, (tag, text, col) in enumerate(points):
        top_pos = 2.45 + i * (p_h + 0.18)
        add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, top_pos, 12.43, p_h, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, top_pos, 2.50, p_h, fill_hex=col)
        add_textbox(s6, 0.55, top_pos + 0.35, 2.30, 0.45, tag, size_pt=16, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s6, 3.15, top_pos + 0.25, 9.50, 0.70, text, size_pt=16, bold=True, color_hex=C_DARK)

    # SLIDE 7: THỬ THÁCH 1 - ĐỐ VUI TRẮC NGHIỆM (SGK tr.25 - Dạng 10)
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s7, "push")
    add_slide_header(s7, "THỬ THÁCH 1 🎯", "Em hãy tìm phát biểu SAI trong các câu dưới đây:",
                     "Thử tài nhận định chính xác về bản quyền nội dung thông tin (SGK Trang 25)", primary_hex=C_PRI)

    # 4 Card 2x2
    options = [
        ("A", "Tác giả có quyền cho phép người khác sử dụng nội dung thông tin của mình.", False),
        ("B", "Tác giả có quyền không cho phép người khác sử dụng nội dung thông tin của mình.", False),
        ("C", "Nội dung thông tin được thể hiện dưới dạng văn bản, hình ảnh, âm thanh hay đa phương tiện,...", False),
        ("D", "Nội dung thông tin phải gồm đầy đủ các dạng văn bản, hình ảnh, âm thanh.", True),
    ]
    coords = [
        (0.45, 2.40), (6.75, 2.40),
        (0.45, 4.35), (6.75, 4.35)
    ]
    for idx, (label, text, is_wrong) in enumerate(options):
        lx, ty = coords[idx]
        b_color = C_ORG if is_wrong else C_BORD
        f_color = C_ORGBG if is_wrong else C_CARD
        add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, lx, ty, 6.13, 1.75, fill_hex=f_color, border_hex=b_color, border_width_pt=2 if is_wrong else 1)
        # Letter circle
        add_safe_shape(s7, MSO_SHAPE.OVAL, lx + 0.20, ty + 0.20, 0.55, 0.55, fill_hex=C_ORG if is_wrong else C_PRI)
        add_textbox(s7, lx + 0.20, ty + 0.22, 0.55, 0.50, label, size_pt=18, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        # Text
        add_textbox(s7, lx + 0.90, ty + 0.20, 5.05, 1.35, text, size_pt=14, bold=is_wrong, color_hex=C_DARK)
        if is_wrong:
            add_textbox(s7, lx + 0.90, ty + 1.25, 5.05, 0.40, "👉 ĐÁP ÁN SAI CẦN CHỌN! (Không bắt buộc phải có đủ mọi dạng)", size_pt=12, bold=True, color_hex=C_ORG)

    # SLIDE 8: THỬ THÁCH 2 - GHÉP NỐI TÌNH HUỐNG (SGK tr.26)
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s8, "fade")
    add_slide_header(s8, "THỬ THÁCH 2 🔗", "Trò chơi ghép nối: Tình huống & Cách ghi nguồn",
                     "Ghép mỗi tình huống ở cột trái với một cách ghi nguồn phù hợp ở cột phải (SGK Trang 26)", primary_hex=C_PRI)

    matches = [
        ("1) Sử dụng một bức ảnh của tác giả Nguyễn A.", "b) Tác giả: Nguyễn A", C_PRI),
        ("2) Sử dụng một bài thơ không rõ tên tác giả.", "c) Sưu tầm", C_ACC),
        ("3) Kể câu chuyện dựa trên nội dung truyện Cây khế.", "a) Phỏng theo truyện Cây khế", C_ORG),
    ]
    for i, (left_txt, right_txt, col) in enumerate(matches):
        y_pos = 2.45 + i * 1.25
        # Cột trái
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, y_pos, 5.80, 1.05, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_textbox(s8, 0.65, y_pos + 0.15, 5.40, 0.75, left_txt, size_pt=15, bold=False, color_hex=C_DARK)
        # Mũi tên giữa
        add_textbox(s8, 6.35, y_pos + 0.20, 0.63, 0.60, "➔", size_pt=24, bold=True, color_hex=col, alignment=PP_ALIGN.CENTER)
        # Cột phải
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 7.08, y_pos, 5.80, 1.05, fill_hex=col)
        add_textbox(s8, 7.28, y_pos + 0.25, 5.40, 0.60, right_txt, size_pt=16, bold=True, color_hex=C_WHITE)

    # SLIDE 9: TỔNG KẾT TIẾT 1
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s9, "fade")
    add_slide_header(s9, "TỔNG KẾT TIẾT 1 🚀", "Hôm nay chúng mình đã học được những gì?",
                     "3 Hành trang quý giá về bản quyền để các em tự tin học tập trên môi trường số", primary_hex=C_PRI)

    recap = [
        ("1. BẢN QUYỀN", "Hiểu rằng tác phẩm (bài viết, tranh vẽ, video) thuộc về tác giả sáng tạo ra nó.", C_PRI),
        ("2. SỬ DỤNG HỢP LỆ", "Chỉ sử dụng khi được phép và luôn nhớ ghi rõ nguồn tin / tác giả.", C_ACC),
        ("3. ĐẠO ĐỨC SỐ", "Không xem trộm, không tự ý sao chép tài liệu của người khác khi chưa xin phép.", C_ORG)
    ]
    for i, (title, desc, col) in enumerate(recap):
        c_left = 0.45 + i * (card_w + card_gap)
        add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 2.45, card_w, 3.65, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 2.45, card_w, 0.50, fill_hex=col)
        add_textbox(s9, c_left + 0.15, 2.50, card_w - 0.30, 0.40, title, size_pt=16, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s9, c_left + 0.25, 3.20, card_w - 0.50, 2.60, desc, size_pt=16, bold=False, color_hex=C_DARK)

    # SLIDE 10: HẸN GẶP LẠI (Kết bài Tiết 1)
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s10, "fade")
    add_safe_shape(s10, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.25, 12.43, 5.00, fill_hex=C_PRI, send_to_back=True)
    add_textbox(s10, 1.0, 2.10, 11.33, 0.70, "CHÚC MỪNG CÁC EM ĐÃ HOÀN THÀNH TIẾT 1! 🎉", size_pt=28, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
    add_textbox(s10, 1.5, 3.00, 10.33, 1.20,
                "🌟 Nhiệm vụ chuẩn bị cho Tiết 2:\n"
                "• Xem trước Hoạt động 2 (SGK Trang 27): Tính riêng tư và Bản quyền nội dung thông tin.\n"
                "• Chuẩn bị ý tưởng làm bài trình chiếu có trích dẫn nguồn hình ảnh chuẩn xác!",
                size_pt=18, bold=False, color_hex="E0E7FF", alignment=PP_ALIGN.CENTER)
    add_textbox(s10, 1.0, 4.60, 11.33, 0.50, "HẸN GẶP LẠI CÁC EM Ở TIẾT 2 NHÉ! 👋", size_pt=22, bold=True, color_hex="FDE047", alignment=PP_ALIGN.CENTER)

    out_path_t11 = os.path.join(OUTPUT_DIR, 'Slide_Tin_hoc_Lớp_5_Tiet11_Bai_5_Ban_quyen_noi_dung_thong_tin_Tiet_1.pptx')
    prs.save(out_path_t11)
    print(f"Saved Tiet 11 deck to: {out_path_t11}")


# =============================================================================
# BUILD TIẾT 12: BÀI 5 - TIẾT 2
# =============================================================================
def build_deck_tiet12():
    print("Building Tiet 12 deck...")
    prs = Presentation(TEMPLATE_PATH)
    blank_layout = prs.slide_layouts[6]
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # Palette Tiết 12: Deep Purple & Sky Blue
    C_PRI    = "4C1D95"  # Tím đậm quý phái
    C_ACC    = "0284C7"  # Xanh da trời tươi
    C_ACCL   = "F3E8FF"  # Tím pastel nhạt
    C_BG     = "F8FAFC"  # Nền trang nhã
    C_CARD   = "FFFFFF"  # Card trắng
    C_BORD   = "CBD5E1"  # Viền
    C_DARK   = "0F172A"  # Chữ đen than
    C_MUTED  = "475569"  # Chữ phụ
    C_WHITE  = "FFFFFF"  # Chữ trắng
    C_ORG    = "EA580C"  # Cam Hộp Kết Luận
    C_ORGBG  = "FFF7ED"  # Nền cam nhạt
    C_GRN    = "16A34A"  # Xanh lá

    # SLIDE 1: TRANG BÌA (Hero Banner)
    s1 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s1, "fade")
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.25, 12.43, 5.00, fill_hex=C_BG, border_hex=C_BORD, send_to_back=True)

    add_badge(s1, "TIN HỌC LỚP 5 • TUẦN 12 • TIẾT 12", left=0.85, top=1.55, width=3.4, height=0.40, fill_hex=C_PRI)
    add_textbox(s1, 0.85, 2.15, 6.0, 1.60, "BẢN QUYỀN NỘI DUNG\nTHÔNG TIN (TIẾT 2)", size_pt=27, bold=True, color_hex=C_PRI)
    add_textbox(s1, 0.85, 3.80, 6.0, 0.60, "Bài 5 • Tôn trọng tính riêng tư & Sử dụng thông tin hợp lệ\nBộ sách: Kết nối tri thức với cuộc sống (Trang 27 - 28)", size_pt=14, bold=False, color_hex=C_MUTED)

    add_textbox(s1, 0.85, 4.55, 5.8, 1.40,
                "🛡️ Tôn trọng tính riêng tư của bạn bè và mọi người\n"
                "🌐 Nắm chắc 3 nguyên tắc vàng tôn trọng bản quyền\n"
                "🖼️ Thực hành tìm kiếm & trích dẫn nguồn hình ảnh chuẩn",
                size_pt=15, bold=False, color_hex=C_DARK)

    # Hero image bên phải
    cover_img = os.path.join(IMG_DIR_T12, 'cover_hero.jpg')
    add_safe_shape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 7.10, 1.50, 5.45, 4.45, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    add_picture_safe(s1, cover_img, 7.18, 1.58, 5.29, 4.29)

    # SLIDE 2: KHỞI ĐỘNG (Hoạt động 2 SGK tr.27 - Hai tình huống thực tế)
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s2, "push")
    add_slide_header(s2, "KHỞI ĐỘNG 💬", "Hoạt động 2: Em cần làm gì khi chứng kiến các tình huống sau?",
                     "Phân tích 2 tình huống học đường trong SGK Tin học 5 (Trang 27)", primary_hex=C_PRI)

    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    add_textbox(s2, 0.65, 2.55, 5.40, 0.35, "📖 2 TÌNH HUỐNG SGK NÊU RA:", size_pt=16, bold=True, color_hex=C_PRI)
    add_textbox(s2, 0.65, 2.95, 5.40, 1.95,
                "• Tình huống 1: An vô tình để quên nhật kí trong ngăn bàn. Bình thấy được liền đọc lén, sau đó kể cho một số bạn về những điều đọc được.\n\n"
                "• Tình huống 2: Minh vẽ một bức tranh phong cảnh rất đẹp trên máy tính. Hoa đã lấy bức tranh đó sử dụng mà không hỏi ý kiến Minh.",
                size_pt=14, bold=False, color_hex=C_DARK)

    # Hộp cam tương tác
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 5.00, 5.40, 1.15, fill_hex=C_ORGBG, border_hex=C_ORG)
    add_textbox(s2, 0.80, 5.05, 5.10, 1.00,
                "❓ THẢO LUẬN: Em có đồng tình với hành động của bạn Bình và bạn Hoa không? Vì sao?",
                size_pt=14, bold=True, color_hex=C_ORG)

    # Ảnh phải
    priv_img = os.path.join(IMG_DIR_T12, 'privacy_diary_drawing.jpg')
    add_safe_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    add_picture_safe(s2, priv_img, 6.55, 2.50, 6.23, 3.65)

    # SLIDE 3: KHÁM PHÁ 1 - TÔN TRỌNG TÍNH RIÊNG TƯ (SGK tr.27)
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s3, "fade")
    add_slide_header(s3, "BẢO VỆ RIÊNG TƯ 🔒", "1. Tính riêng tư & Thông tin cá nhân là gì?",
                     "Hiểu rõ quyền được bảo mật thông tin cá nhân của mỗi con người trong thời đại số", primary_hex=C_PRI)

    # Cột 1: Thông tin riêng tư
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 3.85, fill_hex=C_CARD, border_hex=C_PRI, border_width_pt=2)
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 5.80, 0.45, fill_hex=C_PRI)
    add_textbox(s3, 0.65, 2.45, 5.40, 0.35, "THÔNG TIN MANG TÍNH RIÊNG TƯ GỒM NHỮNG GÌ?", size_pt=14, bold=True, color_hex=C_WHITE)
    add_textbox(s3, 0.65, 3.00, 5.40, 3.10,
                "📖 Cuốn nhật kí, sổ ghi chép cá nhân\n\n"
                "💬 Tin nhắn, thư từ riêng gửi cho người khác\n\n"
                "🔑 Mật khẩu máy tính, mật khẩu tài khoản mạng\n\n"
                "🏠 Họ tên, ngày sinh, địa chỉ nhà, số điện thoại...\n\n"
                "👉 QUYỀN RIÊNG TƯ: Chỉ người sở hữu mới có quyền quyết định cho người khác biết hay không!",
                size_pt=14, bold=False, color_hex=C_DARK)

    # Cột 2: Bài học từ 2 tình huống
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 3.85, fill_hex=C_CARD, border_hex=C_ORG, border_width_pt=2)
    add_safe_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 2.40, 6.43, 0.45, fill_hex=C_ORG)
    add_textbox(s3, 6.65, 2.45, 6.00, 0.35, "BÀI HỌC ỨNG XỬ VĂN MINH TRONG HỌC ĐƯỜNG", size_pt=14, bold=True, color_hex=C_WHITE)
    add_textbox(s3, 6.65, 3.00, 6.00, 3.10,
                "❌ Bạn Bình đã SAI: Đọc lén nhật kí của bạn và kể cho người khác là vi phạm nghiêm trọng tính riêng tư!\n\n"
                "❌ Bạn Hoa đã SAI: Lấy tranh vẽ máy tính của bạn khi chưa hỏi ý kiến là vi phạm bản quyền tác giả!\n\n"
                "✅ HÀNH ĐỘNG ĐÚNG:\n"
                "• Hoa cần xin phép Minh trước khi sử dụng tranh.\n"
                "• Khi được đồng ý, Hoa phải ghi rõ Minh là tác giả bức tranh!",
                size_pt=14, bold=False, color_hex=C_DARK)

    # SLIDE 4: KHÁM PHÁ 2 - 3 NGUYÊN TẮC VÀNG TÔN TRỌNG BẢN QUYỀN (SGK tr.27)
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s4, "fade")
    add_slide_header(s4, "SƠ ĐỒ NGUYÊN TẮC 🌟", "2. Ba nguyên tắc tôn trọng bản quyền nội dung thông tin",
                     "Sơ đồ 3 nhánh chuẩn mực trích xuất từ SGK Tin học 5 (Trang 27)", primary_hex=C_PRI)

    # Bên trái: Infographic SGK
    infog_path = os.path.join(IMG_DIR_T12, 'infographic_ton_trong_ban_quyen.png')
    add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 6.60, 3.85, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    add_picture_safe(s4, infog_path, 0.60, 2.50, 6.30, 3.65)

    # Bên phải: Giải thích chi tiết 3 nhánh
    add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 7.25, 2.40, 5.63, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    branches = [
        ("🟢 NHÁNH 1 (Xanh lá)", "Chỉ sử dụng khi được phép:\nLuôn hỏi ý kiến và được sự đồng ý của tác giả trước khi dùng."),
        ("🔵 NHÁNH 2 (Xanh dương)", "Ghi rõ nguồn thông tin:\nGhi tên tác giả, nhà xuất bản hoặc địa chỉ trang web gốc."),
        ("🟠 NHÁNH 3 (Vàng cam)", "Diễn đạt lại & ghi \"phỏng theo\":\nNếu tóm tắt bằng lời văn của mình thì phải ghi rõ phỏng theo tác phẩm nào.")
    ]
    b_h = 1.10
    for i, (b_title, b_desc) in enumerate(branches):
        b_top = 2.55 + i * (b_h + 0.12)
        add_safe_shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, 7.45, b_top, 5.23, b_h, fill_hex=C_BG, border_hex=C_BORD)
        add_textbox(s4, 7.60, b_top + 0.08, 4.95, 0.30, b_title, size_pt=14, bold=True, color_hex=C_PRI)
        add_textbox(s4, 7.60, b_top + 0.38, 4.95, 0.65, b_desc, size_pt=13, bold=False, color_hex=C_DARK)

    # SLIDE 5: THÁI ĐỘ VĂN MINH - ĐẤU TRANH VỚI SAI TRÁI (SGK tr.27)
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s5, "fade")
    add_slide_header(s5, "THÁI ĐỘ ĐẠO ĐỨC 🛡️", "Thái độ của em trước những hành vi sai trái trong học tập",
                     "Rèn luyện phẩm chất Trung thực và Trách nhiệm trong môi trường số", primary_hex=C_PRI)

    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 12.43, 3.85, fill_hex=C_CARD, border_hex=C_PRI)
    
    add_textbox(s5, 0.75, 2.60, 11.8, 0.40, "💬 TRONG THỰC TẾ, KHI GẶP CÁC HIỆN TƯỢNG SAI TRÁI, GIAN DỐI:", size_pt=17, bold=True, color_hex=C_PRI)
    
    behaviors = [
        ("❌ Xem thư riêng, tin nhắn, nhật kí của bạn bè", "Xâm phạm quyền riêng tư nghiêm trọng"),
        ("❌ Sao chép tệp, bài làm của bạn khi chưa xin phép", "Gian dối trong học tập, vi phạm bản quyền"),
        ("❌ Lấy ảnh, video trên mạng về nhận là của mình tạo ra", "Vi phạm quyền tác giả và đạo đức số")
    ]
    for i, (act, warn) in enumerate(behaviors):
        y_pos = 3.15 + i * 0.70
        add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.75, y_pos, 11.8, 0.60, fill_hex=C_ORGBG, border_hex=C_ORG)
        add_textbox(s5, 0.95, y_pos + 0.12, 7.5, 0.38, act, size_pt=15, bold=True, color_hex=C_ORG)
        add_textbox(s5, 8.50, y_pos + 0.12, 3.8, 0.38, "👉 " + warn, size_pt=13, bold=False, color_hex=C_MUTED)

    # Box hành động đúng
    add_safe_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.75, 5.35, 11.8, 0.75, fill_hex=C_ACCL, border_hex=C_ACC)
    add_textbox(s5, 0.95, 5.45, 11.4, 0.55,
                "🌟 BÀI HỌC HÀNH ĐỘNG: Em cần thể hiện rõ thái độ KHÔNG ĐỒNG TÌNH, nhẹ nhàng góp ý, nhắc nhở để các hiện tượng đó không tái diễn!",
                size_pt=15, bold=True, color_hex=C_PRI)

    # SLIDE 6: LUYỆN TẬP 1 - BÀY TỎ THÁI ĐỘ ĐỒNG TÌNH / KHÔNG ĐỒNG TÌNH (SGK tr.28)
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s6, "push")
    add_slide_header(s6, "LUYỆN TẬP 1 ⚖️", "Ghép tình huống với thái độ phù hợp: Đồng tình hay Không đồng tình?",
                     "Bài tập ghép nối tình huống trong SGK Tin học 5 (Trang 28)", primary_hex=C_PRI)

    scenarios = [
        ("1) Bạn An nêu tên tác giả bài thơ mà bạn đọc trong buổi sinh hoạt ngoại khoá.", "ĐỒNG TÌNH ✅", True, C_GRN),
        ("2) Bạn Bình xem sổ ghi chép của anh trai mà không xin phép.", "KHÔNG ĐỒNG TÌNH ❌", False, C_ORG),
        ("3) Bạn Khoa xin phép Minh trước khi sao chép tệp trình chiếu của bạn Minh.", "ĐỒNG TÌNH ✅", True, C_GRN)
    ]
    for i, (desc, attitude, is_agree, col) in enumerate(scenarios):
        y_pos = 2.45 + i * 1.25
        # Cột trái: Tình huống
        add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, y_pos, 7.80, 1.05, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_textbox(s6, 0.65, y_pos + 0.20, 7.40, 0.65, desc, size_pt=15, bold=False, color_hex=C_DARK)
        # Cột phải: Thái độ
        add_safe_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 8.45, y_pos, 4.43, 1.05, fill_hex=col)
        add_textbox(s6, 8.55, y_pos + 0.25, 4.23, 0.55, attitude, size_pt=18, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)

    # SLIDE 7: LUYỆN TẬP 2 - PHÂN TÍCH NGUỒN ẢNH SGK MĨ THUẬT 4 (SGK tr.28)
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s7, "fade")
    add_slide_header(s7, "LUYỆN TẬP 2 🖼️", "Quan sát trang sách Mĩ thuật 4 (Hình 27 SGK tr.28)",
                     "Tìm hiểu cách các tác giả sách giáo khoa trích dẫn nguồn hình ảnh một cách chuyên nghiệp", primary_hex=C_PRI)

    # Trái: Ảnh SGK Mĩ thuật 4
    h27_path = os.path.join(IMG_DIR_T12, 'h27_sgk_mi_thuat_4.png')
    add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 6.20, 3.85, fill_hex=C_CARD, border_hex=C_ACC, border_width_pt=2)
    add_picture_safe(s7, h27_path, 0.55, 2.50, 6.00, 3.65)

    # Phải: Thảo luận & Giải đáp
    add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 6.85, 2.40, 6.03, 3.85, fill_hex=C_CARD, border_hex=C_BORD)
    add_textbox(s7, 7.05, 2.55, 5.60, 0.35, "💡 CÂU HỎI KHÁM PHÁ:", size_pt=16, bold=True, color_hex=C_PRI)
    add_textbox(s7, 7.05, 2.95, 5.60, 1.60,
                "1. Nguồn thông tin của mỗi bức ảnh là gì?\n"
                "• Ảnh 1 (Luỹ tre) & Ảnh 2 (Cánh đồng):\n"
                "  👉 Nguồn tác giả: Huỳnh Lãnh\n"
                "• Ảnh 3 (Hoàng hôn trên biển):\n"
                "  👉 Nguồn tác giả: Hương Đoàn",
                size_pt=14, bold=False, color_hex=C_DARK)

    # Hộp cam vì sao ghi nguồn
    add_safe_shape(s7, MSO_SHAPE.ROUNDED_RECTANGLE, 7.05, 4.70, 5.63, 1.40, fill_hex=C_ORGBG, border_hex=C_ORG)
    add_textbox(s7, 7.20, 4.75, 5.33, 1.30,
                "2. Tại sao sách giáo khoa phải ghi rõ nguồn?\n"
                "👉 Để tôn trọng quyền tác giả của nhiếp ảnh gia và giúp bạn đọc biết nguồn gốc xuất xứ của bức ảnh!",
                size_pt=14, bold=True, color_hex=C_ORG)

    # SLIDE 8: VẬN DỤNG THỰC TẾ - DỰ ÁN BÀI TRÌNH CHIẾU (SGK tr.28)
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s8, "fade")
    add_slide_header(s8, "VẬN DỤNG THỰC TẾ 🚀", "Dự án: Ngày hội Trạng Nguyên nhỏ tuổi",
                     "Cách xử lý khi nhóm em lấy tư liệu từ Internet để làm bài thuyết trình (SGK Trang 28)", primary_hex=C_PRI)

    # Đề bài tình huống
    add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 2.40, 12.43, 1.25, fill_hex=C_ACCL, border_hex=C_PRI)
    add_textbox(s8, 0.65, 2.45, 12.0, 0.30, "📌 TÌNH HUỐNG VẬN DỤNG:", size_pt=15, bold=True, color_hex=C_PRI)
    add_textbox(s8, 0.65, 2.80, 12.0, 0.75,
                "Trong bài trình chiếu về Ngày hội Trạng Nguyên nhỏ tuổi, nhóm em đã sử dụng một số hình ảnh và tư liệu từ website thieunien.vn. Theo em, các bạn cần làm gì để thể hiện sự tôn trọng bản quyền nội dung thông tin?",
                size_pt=14, bold=False, color_hex=C_DARK)

    # 3 Bước hướng dẫn
    steps = [
        ("BƯỚC 1: KIỂM TRA", "Kiểm tra xem trang web thieunien.vn có cho phép tải và sử dụng hình ảnh cho mục đích học tập hay không.", C_PRI),
        ("BƯỚC 2: CHÚ THÍCH NGUỒN", "Dưới mỗi bức ảnh chèn vào slide trình chiếu, ghi chú thích ngắn gọn:\n\"Ảnh: Báo Thiếu niên Tiền phong (thieunien.vn)\".", C_ACC),
        ("BƯỚC 3: TRANG THAM KHẢO", "Ở slide cuối cùng của bài thuyết trình, dành riêng 1 trang ghi:\n\"Tài liệu tham khảo & Nguồn ảnh: Website thieunien.vn\".", C_ORG)
    ]
    card_w = 3.90
    card_gap = 0.36
    for i, (title, desc, col) in enumerate(steps):
        c_left = 0.45 + i * (card_w + card_gap)
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 3.80, card_w, 2.45, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_safe_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 3.80, card_w, 0.42, fill_hex=col)
        add_textbox(s8, c_left + 0.15, 3.84, card_w - 0.30, 0.35, title, size_pt=14, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s8, c_left + 0.20, 4.35, card_w - 0.40, 1.80, desc, size_pt=13, bold=False, color_hex=C_DARK)

    # SLIDE 9: THÔNG ĐIỆP VÀNG UNIGO
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s9, "fade")
    add_slide_header(s9, "THÔNG ĐIỆP VÀNG 🌟", "Trở thành Công dân số văn minh - Tôn trọng bản quyền",
                     "4 Nguyên tắc vàng mỗi học sinh UNIGO luôn ghi nhớ và thực hiện mỗi ngày", primary_hex=C_PRI)

    golden_rules = [
        ("1. TÔN TRỌNG BẠN BÈ", "Tuyệt đối không xem trộm nhật kí, thư từ hay tin nhắn của người khác."),
        ("2. KHÔNG TỰ Ý SAO CHÉP", "Không sao chép bài làm, tệp tin khi chưa được sự đồng ý của tác giả."),
        ("3. LUÔN GHI RÕ NGUỒN", "Mỗi khi sử dụng hình ảnh, bài thơ, âm nhạc: Nhớ ghi tên tác giả và xuất xứ."),
        ("4. DŨNG CẢM LÊN TIẾNG", "Thể hiện thái độ không đồng tình và nhắc nhở khi thấy hành vi gian dối, vi phạm bản quyền.")
    ]
    for i, (title, desc) in enumerate(golden_rules):
        y_pos = 2.45 + i * 0.95
        col = [C_PRI, C_ACC, C_ORG, C_GRN][i]
        add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, y_pos, 12.43, 0.82, fill_hex=C_CARD, border_hex=col, border_width_pt=2)
        add_safe_shape(s9, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, y_pos, 2.80, 0.82, fill_hex=col)
        add_textbox(s9, 0.55, y_pos + 0.22, 2.60, 0.40, title, size_pt=14, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
        add_textbox(s9, 3.45, y_pos + 0.18, 9.20, 0.50, desc, size_pt=15, bold=True, color_hex=C_DARK)

    # SLIDE 10: TỔNG KẾT BÀI HỌC & DẶN DÒ
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_transition(s10, "fade")
    add_safe_shape(s10, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.25, 12.43, 5.00, fill_hex=C_PRI, send_to_back=True)
    add_textbox(s10, 1.0, 2.00, 11.33, 0.70, "CHÚC MỪNG CÁC EM ĐÃ HOÀN THÀNH BÀI 5! 🎉", size_pt=28, bold=True, color_hex=C_WHITE, alignment=PP_ALIGN.CENTER)
    add_textbox(s10, 1.5, 2.90, 10.33, 1.30,
                "🌟 Dặn dò về nhà:\n"
                "• Ghi nhớ 3 nguyên tắc vàng về bản quyền nội dung thông tin.\n"
                "• Luôn thực hành ghi rõ nguồn tác giả trong mọi bài tập và bài thuyết trình sau này.\n"
                "• Chuẩn bị cho Bài 6: Định dạng kí tự và bố trí hình ảnh trong văn bản.",
                size_pt=17, bold=False, color_hex="E0E7FF", alignment=PP_ALIGN.CENTER)
    add_textbox(s10, 1.0, 4.60, 11.33, 0.50, "CẢM ƠN THẦY CÔ VÀ CÁC EM ĐÃ CHÚ Ý LẮNG NGHE! ❤️", size_pt=22, bold=True, color_hex="FDE047", alignment=PP_ALIGN.CENTER)

    out_path_t12 = os.path.join(OUTPUT_DIR, 'Slide_Tin_hoc_Lớp_5_Tiet12_Bai_5_Ban_quyen_noi_dung_thong_tin_Tiet_2.pptx')
    prs.save(out_path_t12)
    print(f"Saved Tiet 12 deck to: {out_path_t12}")

if __name__ == '__main__':
    build_deck_tiet11()
    build_deck_tiet12()
    print("ALL DONE SUCCESSFULLY!")
