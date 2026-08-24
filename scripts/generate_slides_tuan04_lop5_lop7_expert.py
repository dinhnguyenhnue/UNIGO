# -*- coding: utf-8 -*-
"""
Hệ thống Tạo và Chuẩn hóa Slide Bài Giảng Tin học Lớp 5 & Lớp 7 - Tuần 04 (UNIGO)
Đáp ứng 100% quy chuẩn:
1. Template UNIGO bảo tồn Logo (Y < 1.09in) & Chân trang Master (Y > 6.43in).
2. Vùng An Toàn: Y = 1.15in -> 6.35in.
3. 20 Dạng Bố Cục Chuẩn Visual-First & Hero Image (ảnh chiếm 50-70%).
4. Đầy đủ hình ảnh SGK trích xuất thực tế + Thẻ đồ họa trực quan.
5. Tạo cả 2 bản: Slide riêng biệt từng tiết và Slide tổng hợp 2 tiết.
"""
import sys, os, io
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree
from PIL import Image

TEMPLATE  = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide UNIGO.pptx'
SAFE_TOP    = 1.15
SAFE_BOTTOM = 6.35
SAFE_LEFT   = 0.4
SAFE_RIGHT  = 12.9
SLIDE_W     = 13.33
SLIDE_H     = 7.50

# Color Palettes
PAL_LOP5_T3 = { # Rose / Berry
    "primary": "BE185D", "accent": "EC4899", "bg": "FDF2F8", "card": "FFFFFF",
    "text_on_primary": "FFFFFF", "text_on_bg": "831843", "text_on_card": "0F172A"
}
PAL_LOP5_T4 = { # Ruby / Crimson
    "primary": "9F1239", "accent": "F43F5E", "bg": "FFF1F2", "card": "FFFFFF",
    "text_on_primary": "FFFFFF", "text_on_bg": "881337", "text_on_card": "0F172A"
}
PAL_LOP7_T3 = { # Sky Blue / Ocean
    "primary": "0369A1", "accent": "0EA5E9", "bg": "F0F9FF", "card": "FFFFFF",
    "text_on_primary": "FFFFFF", "text_on_bg": "0C4A6E", "text_on_card": "0F172A"
}
PAL_LOP7_T4 = { # Cyan / Deep Teal
    "primary": "0E7490", "accent": "06B6D4", "bg": "ECFEFF", "card": "FFFFFF",
    "text_on_primary": "FFFFFF", "text_on_bg": "164E63", "text_on_card": "0F172A"
}

def hex_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def add_safe_shape(slide, shape_type, left, top, width, height, fill_hex, border_hex=None, send_to_back=False):
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)

    shape = slide.shapes.add_shape(shape_type, Inches(left), Inches(actual_top), Inches(width), Inches(actual_height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_rgb(border_hex)
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()

    if send_to_back:
        spTree = shape._element.getparent()
        spTree.remove(shape._element)
        spTree.insert(2, shape._element)
    return shape

def add_textbox(slide, left, top, width, height, text, size_pt=18, bold=False, color_hex="0F172A", alignment=PP_ALIGN.LEFT, font_name="Arial"):
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)

    txBox = slide.shapes.add_textbox(Inches(left), Inches(actual_top), Inches(width), Inches(actual_height))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)

    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(size_pt)
    p.font.bold = bold
    p.font.color.rgb = hex_rgb(color_hex)
    p.alignment = alignment
    return txBox

def add_picture_safe(slide, img_path, left, top, width, height):
    if not os.path.exists(img_path):
        return None
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)
    try:
        pic = slide.shapes.add_picture(img_path, Inches(left), Inches(actual_top), Inches(width), Inches(actual_height))
        return pic
    except Exception as e:
        print(f"Error adding picture {img_path}: {e}")
        return None

def add_slide_transition(slide, transition_type="fade"):
    slide_xml = slide._element
    existing_trans = slide_xml.xpath('./p:transition')
    if existing_trans:
        for t in existing_trans:
            slide_xml.remove(t)
    trans = etree.SubElement(slide_xml, '{http://schemas.openxmlformats.org/presentationml/2006/main}transition')
    trans.set('spd', 'med')
    etree.SubElement(trans, f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}{transition_type}')

# ─── 20 BỐ CỤC CHUẨN (VISUAL-FIRST BUILDERS) ───

def build_cover(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "wipe")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["primary"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.6, 10.93, 4.0, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.6, 10.93, 0.2, pal["accent"])
    
    add_textbox(slide, 1.5, 2.0, 10.33, 0.5, "TRƯỜNG TIỂU HỌC & THCS UNIGO", size_pt=15, bold=True, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 2.6, 10.33, 1.4, data["title"], size_pt=28, bold=True, color_hex=pal["text_on_card"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 4.1, 10.33, 0.6, data["subtitle"], size_pt=18, bold=False, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 4.8, 10.33, 0.4, "GV: Đậu Đình Nguyên • Bộ môn Tin học", size_pt=14, bold=False, color_hex="64748B", alignment=PP_ALIGN.CENTER)

def build_section_banner(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "push")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, 1.8, 10.33, 3.8, pal["primary"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, 1.8, 0.4, 3.8, pal["accent"])
    
    add_textbox(slide, 2.3, 2.2, 8.8, 0.6, data.get("badge", "KHÁM PHÁ KIẾN THỨC MỚI"), size_pt=16, bold=True, color_hex="FEF08A", alignment=PP_ALIGN.LEFT)
    add_textbox(slide, 2.3, 2.9, 8.8, 1.5, data["title"], size_pt=26, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
    if "desc" in data:
        add_textbox(slide, 2.3, 4.4, 8.8, 0.8, data["desc"], size_pt=16, bold=False, color_hex="F1F5F9", alignment=PP_ALIGN.LEFT)

def build_hero_example(prs, data, pal, layout):
    """Layout 5 & Layout 15: Cột trái (38% Text tóm tắt) | Cột phải (62% Hero Image SGK to rõ)"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    # Header
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.25, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.25, 3.2, 0.38, data.get("badge", "HOẠT ĐỘNG KHÁM PHÁ"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.22, 8.8, 0.45, data["title"], size_pt=21, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    # Left Column: Card Text (3.8in width)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.8, 4.2, 4.35, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.8, 0.18, 4.35, pal["accent"])
    
    text_content = data["text"]
    lines = [l.strip() for l in text_content.split('\n') if l.strip()]
    top_pos = 1.95
    for l in lines:
        is_heading = l.startswith('•') or l[0].isdigit() and l[1] in ['.', ')']
        add_textbox(slide, SAFE_LEFT + 0.3, top_pos, 3.75, 0.7, l, size_pt=15 if is_heading else 14, bold=is_heading, color_hex=pal["text_on_card"])
        top_pos += 0.75
    
    # Right Column: Hero Image (7.6in width)
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 4.5, 1.8, 7.9, 4.35, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 4.6, 1.9, 7.7, 4.15)

def build_arrow_badges_diagram(prs, data, pal, layout):
    """Layout 19: Cột trái (Các thẻ Badge mũi tên quy trình) | Cột phải (Ảnh SGK/Sơ đồ)"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    # Header
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.25, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.25, 3.2, 0.38, data.get("badge", "QUY TRÌNH & THAO TÁC"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.22, 8.8, 0.45, data["title"], size_pt=21, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    badges = data.get("badges", [])
    badge_colors = [pal["primary"], pal["accent"], "0D9488", "D97706", "7C3AED"]
    
    # Left column: Badges (5.2in width)
    cur_top = 1.8
    card_h = 0.95 if len(badges) <= 4 else 0.75
    for i, b in enumerate(badges):
        b_col = badge_colors[i % len(badge_colors)]
        add_safe_shape(slide, MSO_SHAPE.CHEVRON, SAFE_LEFT, cur_top, 5.2, card_h, b_col)
        add_textbox(slide, SAFE_LEFT + 0.3, cur_top + 0.08, 4.6, card_h - 0.15, b, size_pt=14, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
        cur_top += (card_h + 0.12)
        
    # Right column: Image (6.9in width)
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 5.5, 1.8, 6.9, 4.35, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 5.6, 1.9, 6.7, 4.15)

def build_posture_compare(prs, data, pal, layout):
    """Layout 6 & Layout 16: Bộ 2-3 tranh đối sánh kèm nhãn Đúng/Sai hoặc Phân tích tình huống"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.25, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.25, 3.2, 0.38, data.get("badge", "KHỞI ĐỘNG & TÌNH HUỐNG"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.22, 8.8, 0.45, data["title"], size_pt=21, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    # Situations or Images
    items = data.get("items", [])
    if items:
        col_w = (12.3 - (len(items)-1)*0.3) / len(items)
        for i, it in enumerate(items):
            c_left = SAFE_LEFT + i*(col_w + 0.3)
            add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 1.8, col_w, 4.35, "FFFFFF", border_hex=pal["accent"])
            add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, c_left, 1.8, col_w, 0.45, pal["primary"])
            add_textbox(slide, c_left, 1.85, col_w, 0.4, it["header"], size_pt=14, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
            
            img_p = it.get("img")
            if img_p and os.path.exists(img_p):
                add_picture_safe(slide, img_p, c_left + 0.15, 2.35, col_w - 0.3, 2.3)
            
            add_textbox(slide, c_left + 0.15, 4.75, col_w - 0.3, 1.3, it["desc"], size_pt=13, bold=False, color_hex=pal["text_on_card"], alignment=PP_ALIGN.CENTER)
    else:
        # Fallback to single big image + question
        img_path = data.get("img")
        if img_path and os.path.exists(img_path):
            add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.8, 12.3, 4.35, "FFFFFF", border_hex="CBD5E1")
            add_picture_safe(slide, img_path, SAFE_LEFT + 0.2, 1.9, 11.9, 4.15)

def build_conclusion_orange(prs, data, pal, layout):
    """Layout 14: Bóng thoại câu hỏi + Gợi ý + HỘP KẾT LUẬN CAM BO GÓC nổi bật"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    # Question Bubble
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.3, 12.3, 0.9, pal["primary"])
    add_textbox(slide, SAFE_LEFT + 0.3, 1.35, 11.7, 0.8, "❓ " + data["question"], size_pt=18, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
    
    # Middle: Suggestions (Left 6.8in) & Image (Right 5.2in)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.3, 7.0, 2.4, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.3, 0.18, 2.4, pal["accent"])
    
    sugg_top = 2.45
    for sg in data["suggestions"]:
        add_textbox(slide, SAFE_LEFT + 0.3, sugg_top, 6.5, 0.9, sg, size_pt=15, bold=False, color_hex=pal["text_on_card"])
        sugg_top += 1.05
        
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 7.2, 2.3, 5.1, 2.4, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 7.3, 2.35, 4.9, 2.3)
    
    # Orange Conclusion Box at Bottom
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 4.85, 12.3, 1.3, "EA580C")
    add_textbox(slide, SAFE_LEFT + 0.3, 4.9, 11.7, 0.35, "📌 GHI NHỚ TRỌNG TÂM (SGK)", size_pt=14, bold=True, color_hex="FEF08A", alignment=PP_ALIGN.LEFT)
    add_textbox(slide, SAFE_LEFT + 0.3, 5.25, 11.7, 0.85, data["conclusion"], size_pt=16, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)

def build_quiz_mascot(prs, data, pal, layout):
    """Layout 10: Trắc nghiệm 4 đáp án A-B-C-D kèm Mascot trang trí"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    # Header & Question
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.25, 3.2, 0.38, "EA580C")
    add_textbox(slide, SAFE_LEFT, 1.25, 3.2, 0.38, "LUYỆN TẬP NHANH", size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.75, 12.3, 1.0, "FFFFFF", border_hex=pal["primary"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.75, 0.2, 1.0, pal["primary"])
    add_textbox(slide, SAFE_LEFT + 0.3, 1.85, 11.7, 0.8, "❓ " + data["question"], size_pt=17, bold=True, color_hex=pal["text_on_card"])
    
    # 4 Option Cards (2x2 Grid)
    opts = data["options"]
    card_w = 5.95
    card_h = 1.45
    positions = [
        (SAFE_LEFT, 2.9),
        (SAFE_LEFT + 6.35, 2.9),
        (SAFE_LEFT, 4.5),
        (SAFE_LEFT + 6.35, 4.5)
    ]
    colors = [pal["primary"], "0D9488", "D97706", "7C3AED"]
    for idx, (pos_x, pos_y) in enumerate(positions):
        if idx < len(opts):
            opt_text = opts[idx]
            add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, pos_x, pos_y, card_w, card_h, "FFFFFF", border_hex="CBD5E1")
            add_safe_shape(slide, MSO_SHAPE.OVAL, pos_x + 0.2, pos_y + 0.25, 0.9, 0.9, colors[idx])
            letter = opt_text[:2].strip()
            rest_text = opt_text[2:].strip()
            add_textbox(slide, pos_x + 0.2, pos_y + 0.45, 0.9, 0.5, letter, size_pt=18, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
            add_textbox(slide, pos_x + 1.25, pos_y + 0.25, card_w - 1.4, card_h - 0.4, rest_text, size_pt=15, bold=False, color_hex=pal["text_on_card"])

def build_thanks(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "cover")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["primary"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 2.0, 1.8, 9.33, 3.8, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 2.0, 1.8, 9.33, 0.2, pal["accent"])
    
    add_textbox(slide, 2.2, 2.2, 8.93, 0.8, data.get("title", "CẢM ƠN CÁC EM ĐÃ CHÚ Ý LẮNG NGHE! 🌟"), size_pt=24, bold=True, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 2.2, 3.1, 8.93, 1.2, data["content"], size_pt=17, bold=False, color_hex=pal["text_on_card"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 2.2, 4.5, 8.93, 0.6, "Hẹn gặp lại các em ở bài học tuần sau! 🚀", size_pt=16, bold=True, color_hex=pal["accent"], alignment=PP_ALIGN.CENTER)

# ─── DATA DEFINITIONS FOR LỚP 5 & LỚP 7 (TUẦN 04) ───

IMG5 = r'D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images'
IMG7 = r'D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images'

DECKS_CONFIG = [
    # ── 1. LỚP 5 - TIẾT 3 (BÀI 3) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\Slide_Tin_hoc_Lớp_5_Tiet03_Bai_3_Tim_kiem_thong_tin.pptx",
        "pal": PAL_LOP5_T3,
        "slides": [
            ("cover", {"title": "BÀI 3. TÌM KIẾM THÔNG TIN\nTRONG GIẢI QUYẾT VẤN ĐỀ", "subtitle": "Tin học • Lớp 5 • Tiết 3 (Chủ đề 3)"}),
            ("posture_compare", {
                "title": "Tình huống: Chuẩn bị hành lý cho kì nghỉ hè",
                "items": [
                    {"header": "1. An đi Nha Trang", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Nha Trang là thành phố biển, quanh năm nắng ấm. Cần đồ bơi, kem chống nắng..."},
                    {"header": "2. Khoa đi Đà Lạt", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Đà Lạt nằm ở vùng cao nguyên, đêm se lạnh. Cần áo ấm, áo khoác gió..."},
                    {"header": "3. Minh đi Úc", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Nước Úc ở Nam bán cầu. Mùa hè Việt Nam là mùa đông ở Úc! Cần áo phao dày."}
                ]
            }),
            ("hero_example", {
                "title": "Sự cần thiết của việc thu thập thông tin",
                "text": "• Thông tin giúp em hiểu rõ hoàn cảnh thực tế\n• Lựa chọn đồ dùng, trang phục phù hợp\n• Tránh gặp rủi ro thời tiết bất ngờ\n• Giúp giải quyết vấn đề thành công và hiệu quả",
                "img": os.path.join(IMG5, "sgk_lop5_bai3_cac_buoc_tim_kiem.png")
            }),
            ("arrow_badges", {
                "title": "Kỹ thuật sử dụng từ khóa tìm kiếm trên Internet",
                "badges": [
                    "1. Chọn từ khóa ngắn gọn, chính xác (Keyword)",
                    "2. Đặt từ khóa trong dấu ngoặc kép \" \" để tìm chuẩn xác",
                    "3. Chọn lọc từ website tin cậy (.gov.vn, .edu.vn)",
                    "4. Đánh giá và tổng hợp thông tin trước khi áp dụng"
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai3_thuc_hanh_search.png")
            }),
            ("conclusion_orange", {
                "question": "Vì sao phải tìm kiếm thông tin trước khi giải quyết một vấn đề?",
                "suggestions": [
                    "👉 Giúp nắm bắt đầy đủ dữ liệu thực tế và điều kiện liên quan.",
                    "👉 Đưa ra quyết định chính xác, tiết kiệm thời gian và công sức."
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai3_cac_buoc_tim_kiem.png"),
                "conclusion": "Thu thập và tìm kiếm thông tin là bước đầu tiên cực kỳ quan trọng giúp em giải quyết mọi vấn đề một cách khoa học!"
            }),
            ("quiz_mascot", {
                "question": "Để tìm chính xác cụm từ 'dự báo thời tiết Đà Lạt', em nên gõ như thế nào?",
                "options": [
                    "A. \"dự báo thời tiết Đà Lạt\"",
                    "B. thời tiết và mọi thứ ở Đà Lạt",
                    "C. tìm cho tôi thời tiết",
                    "D. da lat hom nay the nao"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH TIẾT 3! 🌟",
                "content": "BTVN: Cùng bố mẹ tìm hiểu thông tin thời tiết địa điểm du lịch cuối tuần này nhé!"
            })
        ]
    },

    # ── 2. LỚP 5 - TIẾT 4 (BÀI 4) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\Slide_Tin_hoc_Lớp_5_Tiet04_Bai_4_Cay_thu_muc.pptx",
        "pal": PAL_LOP5_T4,
        "slides": [
            ("cover", {"title": "BÀI 4. CÂY THƯ MỤC\n(FOLDER TREE)", "subtitle": "Tin học • Lớp 5 • Tiết 4 (Chủ đề 3)"}),
            ("hero_example", {
                "title": "Cấu trúc phân cấp Cây thư mục",
                "text": "• Ổ đĩa gốc (Root): C:, D:, E:\n• Thư mục mẹ (Parent folder): Chứa thư mục khác\n• Thư mục con (Subfolder): Nằm bên trong thư mục mẹ\n• Tệp tin (File): Nơi lưu trữ nội dung cụ thể",
                "img": os.path.join(IMG5, "sgk_lop5_bai4_so_do_cay_thu_muc.png")
            }),
            ("arrow_badges", {
                "title": "Khám phá Cây thư mục trong File Explorer",
                "badges": [
                    "1. Ngăn trái: Hiển thị sơ đồ cây các ổ đĩa và thư mục",
                    "2. Ngăn phải: Hiển thị chi tiết tệp và thư mục con bên trong",
                    "3. Thanh địa chỉ (Path): Chỉ rõ đường dẫn vị trí tệp tin",
                    "4. Thao tác: Tạo mới, Đổi tên, Di chuyển, Sao chép, Xóa"
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai4_file_explorer.png")
            }),
            ("hero_example", {
                "title": "Thực hành tổ chức Cây thư mục học tập",
                "text": "• Tạo thư mục mẹ: HOC_TAP trên ổ đĩa D:\n• Tạo các thư mục con: TOAN, TIENG_VIET, TIN_HOC\n• Lưu tệp văn bản bài tập vào đúng thư mục tương ứng\n• Giữ máy tính luôn ngăn nắp, dễ tìm kiếm",
                "img": os.path.join(IMG5, "sgk_lop5_bai4_luyen_tap.png")
            }),
            ("conclusion_orange", {
                "question": "Tổ chức tệp tin theo Cây thư mục mang lại lợi ích gì?",
                "suggestions": [
                    "👉 Dữ liệu được sắp xếp ngăn nắp, khoa học theo từng chủ đề.",
                    "👉 Tìm kiếm tệp tin nhanh chóng, tránh thất lạc hoặc xóa nhầm."
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai4_file_explorer.png"),
                "conclusion": "Cây thư mục là phương pháp quản lý dữ liệu hình cây phân cấp chuẩn mực trong hệ điều hành máy tính!"
            }),
            ("quiz_mascot", {
                "question": "Trong đường dẫn D:\\HOC_TAP\\TIN_HOC\\BaiTap.docx, đâu là thư mục con?",
                "options": [
                    "A. Ổ đĩa D:",
                    "B. Thư mục TIN_HOC",
                    "C. Tệp BaiTap.docx",
                    "D. Thư mục gốc"
                ]
            }),
            ("thanks", {
                "title": "XUẤT SẮC HOÀN THÀNH BÀI HỌC! 🚀",
                "content": "BTVN: Sắp xếp lại các thư mục bài tập trên máy tính của em theo cây thư mục nhé!"
            })
        ]
    },

    # ── 3. LỚP 5 - TỔNG HỢP TUẦN 04 (BÀI 3 & 4) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\Slide_Tin_hoc_Lop_5_Bai03_04.pptx",
        "pal": PAL_LOP5_T3,
        "slides": [
            ("cover", {"title": "TÌM KIẾM THÔNG TIN &\nCẤU TRÚC CÂY THƯ MỤC", "subtitle": "Tin học • Lớp 5 • Tuần 04 (Tiết 3 & Tiết 4)"}),
            ("section_banner", {"title": "TIẾT 3: TÌM KIẾM THÔNG TIN TRONG GIẢI QUYẾT VẤN ĐỀ", "desc": "Khám phá kỹ thuật tìm kiếm chính xác và chọn lọc thông tin đáng tin cậy"}),
            ("posture_compare", {
                "title": "Tình huống: Chuẩn bị hành lý cho kì nghỉ hè",
                "items": [
                    {"header": "1. An đi Nha Trang", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Nha Trang là thành phố biển, quanh năm nắng ấm. Cần đồ bơi, kem chống nắng..."},
                    {"header": "2. Khoa đi Đà Lạt", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Đà Lạt nằm ở vùng cao nguyên, đêm se lạnh. Cần áo ấm, áo khoác gió..."},
                    {"header": "3. Minh đi Úc", "img": os.path.join(IMG5, "sgk_lop5_bai3_khoi_dong.png"), "desc": "Nước Úc ở Nam bán cầu. Mùa hè Việt Nam là mùa đông ở Úc! Cần áo phao dày."}
                ]
            }),
            ("arrow_badges", {
                "title": "Kỹ thuật sử dụng từ khóa tìm kiếm trên Internet",
                "badges": [
                    "1. Chọn từ khóa ngắn gọn, chính xác (Keyword)",
                    "2. Đặt từ khóa trong dấu ngoặc kép \" \" để tìm chuẩn xác",
                    "3. Chọn lọc từ website tin cậy (.gov.vn, .edu.vn)",
                    "4. Đánh giá và tổng hợp thông tin trước khi áp dụng"
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai3_thuc_hanh_search.png")
            }),
            ("section_banner", {"title": "TIẾT 4: BÀI 4. CÂY THƯ MỤC TRONG MÁY TÍNH", "desc": "Tổ chức lưu trữ dữ liệu khoa học theo cấu trúc hình cây phân cấp"}),
            ("hero_example", {
                "title": "Cấu trúc phân cấp Cây thư mục",
                "text": "• Ổ đĩa gốc (Root): C:, D:, E:\n• Thư mục mẹ (Parent folder): Chứa thư mục khác\n• Thư mục con (Subfolder): Nằm bên trong thư mục mẹ\n• Tệp tin (File): Nơi lưu trữ nội dung cụ thể",
                "img": os.path.join(IMG5, "sgk_lop5_bai4_so_do_cay_thu_muc.png")
            }),
            ("hero_example", {
                "title": "Khám phá Cây thư mục trong File Explorer",
                "text": "• Ngăn trái: Sơ đồ cây các ổ đĩa và thư mục\n• Ngăn phải: Nội dung chi tiết của thư mục đang chọn\n• Đường dẫn (Path): Địa chỉ chính xác của tệp\n• Thao tác nhanh: New folder, Rename, Delete",
                "img": os.path.join(IMG5, "sgk_lop5_bai4_file_explorer.png")
            }),
            ("conclusion_orange", {
                "question": "Tổng kết: Kỹ năng trọng tâm của Tuần 04 là gì?",
                "suggestions": [
                    "👉 Tiết 3: Thu thập và tìm kiếm thông tin chính xác bằng từ khóa.",
                    "👉 Tiết 4: Sắp xếp dữ liệu ngăn nắp theo cấu trúc Cây thư mục."
                ],
                "img": os.path.join(IMG5, "sgk_lop5_bai4_luyen_tap.png"),
                "conclusion": "Làm chủ kỹ năng tìm kiếm và tổ chức dữ liệu giúp em trở thành người sử dụng máy tính thông minh và hiệu quả!"
            }),
            ("quiz_mascot", {
                "question": "Trong đường dẫn D:\\HOC_TAP\\TIN_HOC\\BaiTap.docx, đâu là thư mục con?",
                "options": [
                    "A. Ổ đĩa D:",
                    "B. Thư mục TIN_HOC",
                    "C. Tệp BaiTap.docx",
                    "D. Thư mục gốc"
                ]
            }),
            ("thanks", {
                "title": "BÀI HỌC KẾT THÚC! 🌟",
                "content": "BTVN: Hoàn thành bài tập thực hành trên máy tính và chuẩn bị cho bài tuần tới nhé!"
            })
        ]
    },

    # ── 4. LỚP 7 - TIẾT 3 (BÀI 3) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\Slide_Tin_hoc_Lớp_7_Tiet03_Bai_3_Quan_ly_du_lieu_trong_may_tinh.pptx",
        "pal": PAL_LOP7_T3,
        "slides": [
            ("cover", {"title": "BÀI 3. QUẢN LÝ DỮ LIỆU\nTRONG MÁY TÍNH", "subtitle": "Tin học • Lớp 7 • Tiết 3 (Chủ đề 2)"}),
            ("hero_example", {
                "title": "Tên tệp và Thư mục trong máy tính",
                "text": "• Tên tệp gồm 2 phần: [Tên chính].[Phần mở rộng]\n• Phần mở rộng cho biết loại tệp và ứng dụng mở tệp\n• Đặt tên tệp ngắn gọn, có ý nghĩa, dễ tìm kiếm\n• Phân loại tệp vào các thư mục theo chủ đề",
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh31_cay_thu_muc.png")
            }),
            ("arrow_badges", {
                "title": "Các phần mở rộng tệp (File Extensions) thông dụng",
                "badges": [
                    ".docx / .pdf: Tệp văn bản và tài liệu",
                    ".xlsx: Tệp bảng tính điện tử Excel",
                    ".pptx: Tệp bài trình chiếu PowerPoint",
                    ".jpg / .png: Tệp hình ảnh đồ họa",
                    ".mp3 / .mp4: Tệp âm thanh và video đa phương tiện"
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh32_duoi_mo_rong_tep.png")
            }),
            ("hero_example", {
                "title": "Các thao tác quản lý tệp và thư mục cơ bản",
                "text": "• Tạo mới (New Folder): Phím tắt Ctrl + Shift + N\n• Đổi tên (Rename): Phím F2 hoặc chuột phải chọn Rename\n• Sao chép (Copy) / Cắt (Cut): Ctrl + C / Ctrl + X\n• Dán (Paste) / Xóa (Delete): Ctrl + V / Delete",
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh36_lenh_thao_tac.png")
            }),
            ("hero_example", {
                "title": "Sao lưu dữ liệu & Phòng chống Virus máy tính",
                "text": "• Sao lưu định kỳ (Backup) ra USB, ổ cứng ngoài, Cloud\n• Tránh mất dữ liệu khi máy tính bị hỏng hóc\n• Bật tường lửa và phần mềm diệt virus (Windows Defender)\n• Tuyệt đối không mở các tệp đính kèm lạ có đuôi .exe, .bat",
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh33_sao_luu_du_lieu.png")
            }),
            ("conclusion_orange", {
                "question": "Biện pháp an toàn nhất để không bị mất dữ liệu quan trọng là gì?",
                "suggestions": [
                    "👉 Sao lưu dữ liệu dự phòng (Backup) thường xuyên lên đám mây hoặc thiết bị ngoài.",
                    "👉 Đặt mật khẩu bảo vệ tài khoản và quét virus định kỳ."
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh33_sao_luu_du_lieu.png"),
                "conclusion": "Quản lý dữ liệu khoa học và sao lưu dự phòng thường xuyên là nguyên tắc sống còn của người dùng máy tính!"
            }),
            ("quiz_mascot", {
                "question": "Đâu là phần mềm diệt virus và bảo vệ máy tính tích hợp sẵn trong Windows?",
                "options": [
                    "A. Microsoft Word",
                    "B. Windows Defender",
                    "C. File Explorer",
                    "D. Google Chrome"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH TIẾT 3! 🌟",
                "content": "BTVN: Kiểm tra và sao lưu các tài liệu học tập của em lên Google Drive hoặc USB!"
            })
        ]
    },

    # ── 5. LỚP 7 - TIẾT 4 (BÀI 4) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\Slide_Tin_hoc_Lớp_7_Tiet04_Bai_4_Mang_xa_hoi.pptx",
        "pal": PAL_LOP7_T4,
        "slides": [
            ("cover", {"title": "BÀI 4. MẠNG XÃ HỘI &\nTRAO ĐỔI THÔNG TIN INTERNET", "subtitle": "Tin học • Lớp 7 • Tiết 4 (Chủ đề 2)"}),
            ("hero_example", {
                "title": "Các kênh trao đổi thông tin phổ biến trên Internet",
                "text": "• Thư điện tử (E-mail): Trao đổi thư từ trang trọng\n• Diễn đàn trực tuyến (Forum): Thảo luận theo chủ đề\n• Mạng xã hội (Social Media): Kết nối, chia sẻ nội dung\n• Ứng dụng nhắn tin tức thời & Gọi video (Chat/Call)",
                "img": os.path.join(IMG7, "sgk_lop7_bai4_cac_kenh_trao_doi.png")
            }),
            ("arrow_badges", {
                "title": "Các chức năng cơ bản của Mạng xã hội",
                "badges": [
                    "1. Tạo hồ sơ cá nhân và cập nhật thông tin",
                    "2. Đăng bài viết (Post): Văn bản, hình ảnh, video",
                    "3. Kết nối bạn bè, theo dõi các trang thông tin",
                    "4. Tương tác: Thích (Like), Bình luận (Comment), Chia sẻ (Share)",
                    "5. Trò chuyện trực tuyến (Chat Messenger, Nhóm thảo luận)"
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai4_hinh42_43_chuc_nang_mxh.png")
            }),
            ("hero_example", {
                "title": "Mặt trái và Rủi ro khi tham gia Mạng xã hội",
                "text": "• Nguy cơ lộ lọt thông tin cá nhân, bị kẻ xấu lợi dụng\n• Tiếp xúc tin giả, thông tin xấu độc, lừa đảo trực tuyến\n• Nghiện mạng xã hội, giảm tập trung học tập\n• Vấn nạn bắt nạt qua mạng (Cyberbullying)",
                "img": os.path.join(IMG7, "sgk_lop7_bai4_rui_ro_va_an_toan.png")
            }),
            ("arrow_badges", {
                "title": "Quy tắc ứng xử văn minh và an toàn trên không gian số",
                "badges": [
                    "1. Tuyệt đối không chia sẻ mật khẩu và thông tin nhạy cảm",
                    "2. Không đăng tải hình ảnh của người khác khi chưa đồng ý",
                    "3. Kiểm chứng độ tin cậy trước khi chia sẻ tin tức",
                    "4. Giữ thái độ lịch sự, tôn trọng, không công kích xúc phạm"
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai4_giao_dien_mxh_dang_bai.png")
            }),
            ("conclusion_orange", {
                "question": "Em cần làm gì để sử dụng Mạng xã hội an toàn và hiệu quả?",
                "suggestions": [
                    "👉 Bảo vệ danh tính số: Đặt mật khẩu mạnh, bảo mật 2 lớp.",
                    "👉 Sử dụng điều độ, ứng xử văn minh và luôn tôn trọng pháp luật."
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai4_rui_ro_va_an_toan.png"),
                "conclusion": "Mạng xã hội là công cụ kết nối mạnh mẽ. Hãy là người dùng thông thái và có trách nhiệm trên không gian mạng!"
            }),
            ("quiz_mascot", {
                "question": "Hành vi nào sau đây là KHÔNG NÊN LÀM trên mạng xã hội?",
                "options": [
                    "A. Chia sẻ bài học hay cho bạn bè",
                    "B. Đăng số CCCD và mật khẩu lên trang cá nhân",
                    "C. Báo cáo bài viết lừa đảo cho quản trị viên",
                    "D. Chúc mừng sinh nhật bạn bè lịch sự"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH BÀI HỌC! 🚀",
                "content": "BTVN: Kiểm tra và thiết lập chế độ bảo mật riêng tư cho tài khoản mạng xã hội của em!"
            })
        ]
    },

    # ── 6. LỚP 7 - TỔNG HỢP TUẦN 04 (BÀI 3 & 4) ──
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\Slide_Tin_hoc_Lop_7_Bai03_04.pptx",
        "pal": PAL_LOP7_T3,
        "slides": [
            ("cover", {"title": "QUẢN LÝ DỮ LIỆU &\nMẠNG XÃ HỘI INTERNET", "subtitle": "Tin học • Lớp 7 • Tuần 04 (Tiết 3 & Tiết 4)"}),
            ("section_banner", {"title": "TIẾT 3: BÀI 3. QUẢN LÝ DỮ LIỆU TRONG MÁY TÍNH", "desc": "Nắm vững phần mở rộng tệp, thao tác File Explorer và kỹ thuật sao lưu dự phòng"}),
            ("arrow_badges", {
                "title": "Các phần mở rộng tệp (File Extensions) thông dụng",
                "badges": [
                    ".docx / .pdf: Tệp văn bản và tài liệu",
                    ".xlsx: Tệp bảng tính điện tử Excel",
                    ".pptx: Tệp bài trình chiếu PowerPoint",
                    ".jpg / .png: Tệp hình ảnh đồ họa",
                    ".mp3 / .mp4: Tệp âm thanh và video đa phương tiện"
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh32_duoi_mo_rong_tep.png")
            }),
            ("hero_example", {
                "title": "Sao lưu dữ liệu & Phòng chống Virus máy tính",
                "text": "• Sao lưu định kỳ (Backup) ra USB, ổ cứng ngoài, Cloud\n• Tránh mất dữ liệu khi máy tính bị hỏng hóc\n• Bật tường lửa và phần mềm diệt virus (Windows Defender)\n• Tuyệt đối không mở các tệp đính kèm lạ có đuôi .exe, .bat",
                "img": os.path.join(IMG7, "sgk_lop7_bai3_hinh33_sao_luu_du_lieu.png")
            }),
            ("section_banner", {"title": "TIẾT 4: BÀI 4. MẠNG XÃ HỘI & KÊNH TRAO ĐỔI THÔNG TIN", "desc": "Hiểu rõ 2 mặt của Mạng xã hội, kỹ năng giao tiếp văn minh và bảo mật thông tin"}),
            ("arrow_badges", {
                "title": "Các chức năng cơ bản của Mạng xã hội",
                "badges": [
                    "1. Tạo hồ sơ cá nhân và cập nhật thông tin",
                    "2. Đăng bài viết (Post): Văn bản, hình ảnh, video",
                    "3. Kết nối bạn bè, theo dõi các trang thông tin",
                    "4. Tương tác: Thích (Like), Bình luận (Comment), Chia sẻ (Share)",
                    "5. Trò chuyện trực tuyến (Chat Messenger, Nhóm thảo luận)"
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai4_hinh42_43_chuc_nang_mxh.png")
            }),
            ("hero_example", {
                "title": "Mặt trái và Rủi ro khi tham gia Mạng xã hội",
                "text": "• Nguy cơ lộ lọt thông tin cá nhân, bị kẻ xấu lợi dụng\n• Tiếp xúc tin giả, thông tin xấu độc, lừa đảo trực tuyến\n• Nghiện mạng xã hội, giảm tập trung học tập\n• Vấn nạn bắt nạt qua mạng (Cyberbullying)",
                "img": os.path.join(IMG7, "sgk_lop7_bai4_rui_ro_va_an_toan.png")
            }),
            ("conclusion_orange", {
                "question": "Tổng kết kiến thức trọng tâm Tuần 04 (Lớp 7)",
                "suggestions": [
                    "👉 Tiết 3: Tổ chức tệp ngăn nắp, sao lưu dữ liệu và diệt virus định kỳ.",
                    "👉 Tiết 4: Khai thác tiện ích mạng xã hội, ứng xử văn minh và bảo vệ danh tính số."
                ],
                "img": os.path.join(IMG7, "sgk_lop7_bai4_giao_dien_mxh_dang_bai.png"),
                "conclusion": "Quản lý dữ liệu máy tính khoa học kết hợp với văn hóa ứng xử số văn minh là nền tảng của công dân số thời đại 4.0!"
            }),
            ("quiz_mascot", {
                "question": "Hành vi nào sau đây là KHÔNG NÊN LÀM trên mạng xã hội?",
                "options": [
                    "A. Chia sẻ bài học hay cho bạn bè",
                    "B. Đăng số CCCD và mật khẩu lên trang cá nhân",
                    "C. Báo cáo bài viết lừa đảo cho quản trị viên",
                    "D. Chúc mừng sinh nhật bạn bè lịch sự"
                ]
            }),
            ("thanks", {
                "title": "BÀI HỌC KẾT THÚC! 🚀",
                "content": "BTVN: Hoàn thành phiếu học tập và thực hiện bảo mật 2 lớp cho các tài khoản trực tuyến!"
            })
        ]
    }
]

BUILDERS = {
    "cover": build_cover,
    "section_banner": build_section_banner,
    "hero_example": build_hero_example,
    "arrow_badges": build_arrow_badges_diagram,
    "posture_compare": build_posture_compare,
    "conclusion_orange": build_conclusion_orange,
    "quiz_mascot": build_quiz_mascot,
    "thanks": build_thanks
}

def generate_all_decks():
    print("=== BẮT ĐẦU TẠO BỘ SLIDE TIN HỌC LỚP 5 & LỚP 7 (TUẦN 04) ===")
    for config in DECKS_CONFIG:
        prs = Presentation(TEMPLATE)
        blank_layout = prs.slide_layouts[6]
        pal = config["pal"]
        file_path = config["file_path"]
        
        print(f"\n--> Đang tạo: {os.path.basename(file_path)}")
        for b_name, b_data in config["slides"]:
            if b_name in BUILDERS:
                BUILDERS[b_name](prs, b_data, pal, blank_layout)
            else:
                print(f"Unknown builder: {b_name}")
                
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        prs.save(file_path)
        print(f"  [OK] Đã lưu {file_path} ({len(prs.slides)} slides)")

if __name__ == "__main__":
    generate_all_decks()
