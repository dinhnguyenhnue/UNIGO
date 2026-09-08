# -*- coding: utf-8 -*-
r"""
Hệ thống Tạo Slide Bài Giảng UNIGO Chuẩn Visual-First
Tin học Lớp 5 & Lớp 7 - Tuần 6 (Bài 3: Tiết 1 & Tiết 2)
- Tuân thủ nghiêm ngặt Quy định Tạo Slide Bài Giảng (.pptx) chuẩn UNIGO (AGENTS.md Phần VII & tao-slide-bai-giang/SKILL.md).
- Bảo tồn 100% Vùng An Toàn (SAFE ZONE: Top 1.15 in -> Bottom 6.35 in; Left 0.4 in -> 12.9 in).
- Sử dụng Layout 6 (Blank) từ Template UNIGO 'Hệ thống mẫu văn bản\Mẫu slide UNIGO.pptx' để giữ nguyên Header & Footer.
- Tích hợp đầy đủ ảnh SGK và Infographic chất lượng cao (Semantic Alignment 100%).
- Hộp Chốt kiến thức màu cam (#EA580C) với tiêu đề vàng (#FEF08A) và nội dung trắng nổi bật.
- Trắc nghiệm tương tác 4 đáp án A-B-C-D với huy hiệu tròn màu sắc rõ nét.
"""
import sys, os, shutil
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

TEMPLATE    = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide UNIGO.pptx'
SAFE_TOP    = 1.15
SAFE_BOTTOM = 6.35
SAFE_LEFT   = 0.4
SLIDE_W     = 13.33
SLIDE_H     = 7.50

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
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)

    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(size_pt)
    p.font.bold = bold
    p.font.color.rgb = hex_rgb(color_hex)
    p.alignment = alignment
    return txBox

def add_picture_safe(slide, img_path, left, top, width, height):
    if not img_path or not os.path.exists(img_path):
        return None
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)
    try:
        pic = slide.shapes.add_picture(img_path, Inches(left), Inches(actual_top), Inches(width), Inches(actual_height))
        return pic
    except Exception as e:
        print(f"Lỗi chèn ảnh {img_path}: {e}")
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

# ─── CÁC LAYOUT BUILDER CHUẨN ───

def build_cover(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "wipe")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["primary"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.55, 10.93, 4.35, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.55, 10.93, 0.25, pal["accent"])
    
    add_textbox(slide, 1.5, 1.95, 10.33, 0.5, "TRƯỜNG TIỂU HỌC & THCS UNIGO", size_pt=16, bold=True, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 2.55, 10.33, 1.5, data["title"], size_pt=27, bold=True, color_hex=pal["text_on_card"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 4.15, 10.33, 0.6, data["subtitle"], size_pt=17, bold=False, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    add_textbox(slide, 1.5, 4.95, 10.33, 0.45, "GV: Đậu Đình Nguyên • Bộ môn Tin học", size_pt=14, bold=False, color_hex="64748B", alignment=PP_ALIGN.CENTER)

def build_section_banner(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "push")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.7, 10.93, 4.1, pal["primary"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 1.7, 0.35, 4.1, pal["accent"])
    
    add_textbox(slide, 1.8, 2.05, 9.8, 0.5, data.get("badge", "MỤC TIÊU BÀI HỌC"), size_pt=16, bold=True, color_hex="FEF08A", alignment=PP_ALIGN.LEFT)
    add_textbox(slide, 1.8, 2.65, 9.8, 1.1, data["title"], size_pt=24, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
    if "desc" in data:
        lines = data["desc"].split('\n')
        cur_y = 3.9
        for l in lines:
            if l.strip():
                add_textbox(slide, 1.8, cur_y, 9.8, 0.55, l.strip(), size_pt=16, bold=False, color_hex="F1F5F9", alignment=PP_ALIGN.LEFT)
                cur_y += 0.55

def build_hero_example(prs, data, pal, layout):
    """Layout 5 & 15: Cột trái (Card Text tóm tắt) | Cột phải (Hero Image SGK lớn 65%)"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.22, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.22, 3.2, 0.38, data.get("badge", "HOẠT ĐỘNG KHÁM PHÁ"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.20, 8.8, 0.45, data["title"], size_pt=20, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    # Left Column: Card Text (width 5.1 in)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.75, 5.1, 4.45, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.75, 0.18, 4.45, pal["accent"])
    
    txBox = slide.shapes.add_textbox(Inches(SAFE_LEFT + 0.3), Inches(1.85), Inches(4.6), Inches(4.25))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    
    text_content = data["text"]
    lines = [l.strip() for l in text_content.split('\n') if l.strip()]
    for idx, l in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = l
        p.font.name = "Arial"
        is_heading = l.startswith('•') or (l[0].isdigit() and l[1] in ['.', ')']) or l.startswith('👉') or l.startswith('❓')
        p.font.size = Pt(13 if len(lines) > 6 else 14)
        p.font.bold = is_heading
        p.font.color.rgb = hex_rgb(pal["text_on_card"])
        p.space_after = Pt(4)
    
    # Right Column: Hero Image (width 7.0 in)
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 5.4, 1.75, 7.1, 4.45, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 5.5, 1.85, 6.9, 4.25)

def build_full_hero_slide(prs, data, pal, layout):
    """Layout Hero Image lớn ở trên + Card câu hỏi tình huống ở dưới"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.22, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.22, 3.2, 0.38, data.get("badge", "KHỞI ĐỘNG & TÌNH HUỐNG"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.20, 8.8, 0.45, data["title"], size_pt=20, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.75, 12.3, 3.3, "FFFFFF", border_hex=pal["accent"])
        add_picture_safe(slide, img_path, SAFE_LEFT + 0.1, 1.80, 12.1, 3.2)
        
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 5.15, 12.3, 1.05, "FFFFFF", border_hex=pal["primary"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 5.15, 0.2, 1.05, pal["primary"])
    add_textbox(slide, SAFE_LEFT + 0.35, 5.25, 11.6, 0.85, data["question"], size_pt=15, bold=True, color_hex=pal["text_on_card"])

def build_process_diagram_slide(prs, data, pal, layout):
    """Layout 4: Cột trái (Các bước tóm tắt) | Cột phải (Sơ đồ Infographic các bước)"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.22, 3.2, 0.38, pal["primary"])
    add_textbox(slide, SAFE_LEFT, 1.22, 3.2, 0.38, data.get("badge", "QUY TRÌNH & THAO TÁC"), size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    add_textbox(slide, SAFE_LEFT + 3.4, 1.20, 8.8, 0.45, data["title"], size_pt=20, bold=True, color_hex=pal["text_on_bg"], alignment=PP_ALIGN.LEFT)
    
    badges = data.get("badges", [])
    badge_colors = [pal["primary"], pal["accent"], "0D9488", "D97706", "7C3AED"]
    
    # Left column: Badges (4.7 in)
    cur_top = 1.75
    card_h = 0.95 if len(badges) <= 4 else 0.75
    for i, b in enumerate(badges):
        b_col = badge_colors[i % len(badge_colors)]
        add_safe_shape(slide, MSO_SHAPE.CHEVRON, SAFE_LEFT, cur_top, 4.7, card_h, b_col)
        add_textbox(slide, SAFE_LEFT + 0.25, cur_top + 0.08, 4.1, card_h - 0.15, b, size_pt=14, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
        cur_top += (card_h + 0.12)
        
    # Right column: Infographic Flow (7.3 in)
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 5.0, 1.75, 7.3, 4.45, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 5.1, 1.85, 7.1, 4.25)

def build_conclusion_orange(prs, data, pal, layout):
    """Layout 14: Câu hỏi thảo luận + Gợi ý/Hình ảnh + Hộp GHI NHỚ TRỌNG TÂM Cam #EA580C"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    # Question Bubble (1.25 -> 2.10)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.25, 12.3, 0.85, pal["primary"])
    add_textbox(slide, SAFE_LEFT + 0.3, 1.30, 11.7, 0.75, "❓ " + data["question"], size_pt=17, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)
    
    # Suggestions & Image (2.20 -> 4.70)
    img_path = data.get("img")
    if img_path and os.path.exists(img_path):
        # Suggestions left (7.0 in)
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.20, 6.9, 2.45, "FFFFFF", border_hex=pal["accent"])
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.20, 0.18, 2.45, pal["accent"])
        sugg_top = 2.30
        for sg in data["suggestions"]:
            add_textbox(slide, SAFE_LEFT + 0.3, sugg_top, 6.4, 0.95, sg, size_pt=14, bold=False, color_hex=pal["text_on_card"])
            sugg_top += 1.05
        # Image right (5.2 in)
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT + 7.1, 2.20, 5.2, 2.45, "FFFFFF", border_hex="CBD5E1")
        add_picture_safe(slide, img_path, SAFE_LEFT + 7.2, 2.25, 5.0, 2.35)
    else:
        # Full width suggestions (12.3 in)
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.20, 12.3, 2.45, "FFFFFF", border_hex=pal["accent"])
        add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 2.20, 0.20, 2.45, pal["accent"])
        sugg_top = 2.30
        for sg in data["suggestions"]:
            add_textbox(slide, SAFE_LEFT + 0.4, sugg_top, 11.5, 0.70, sg, size_pt=15, bold=False, color_hex=pal["text_on_card"])
            sugg_top += 0.75
            
    # Orange Conclusion Box (4.75 -> 6.25)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 4.75, 12.3, 1.45, "EA580C")
    add_textbox(slide, SAFE_LEFT + 0.3, 4.80, 11.7, 0.35, "📌 GHI NHỚ TRỌNG TÂM (SGK)", size_pt=14, bold=True, color_hex="FEF08A", alignment=PP_ALIGN.LEFT)
    add_textbox(slide, SAFE_LEFT + 0.3, 5.15, 11.7, 0.95, data["conclusion"], size_pt=15, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.LEFT)

def build_quiz_mascot(prs, data, pal, layout):
    """Layout 10: Trắc nghiệm vui 4 đáp án A-B-C-D dạng thẻ 2x2"""
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "fade")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["bg"], send_to_back=True)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.22, 3.2, 0.38, "EA580C")
    add_textbox(slide, SAFE_LEFT, 1.22, 3.2, 0.38, "LUYỆN TẬP NHANH 🎯", size_pt=13, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
    
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.70, 12.3, 1.10, "FFFFFF", border_hex=pal["primary"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SAFE_LEFT, 1.70, 0.2, 1.10, pal["primary"])
    add_textbox(slide, SAFE_LEFT + 0.3, 1.80, 11.7, 0.90, "❓ " + data["question"], size_pt=16, bold=True, color_hex=pal["text_on_card"])
    
    opts = data["options"]
    card_w = 5.95
    card_h = 1.45
    positions = [(SAFE_LEFT, 2.95), (SAFE_LEFT + 6.35, 2.95), (SAFE_LEFT, 4.60), (SAFE_LEFT + 6.35, 4.60)]
    colors = [pal["primary"], "0D9488", "D97706", "7C3AED"]
    for idx, (pos_x, pos_y) in enumerate(positions):
        if idx < len(opts):
            opt_text = opts[idx]
            add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, pos_x, pos_y, card_w, card_h, "FFFFFF", border_hex="CBD5E1")
            add_safe_shape(slide, MSO_SHAPE.OVAL, pos_x + 0.2, pos_y + 0.25, 0.9, 0.9, colors[idx])
            letter = opt_text[:2].strip()
            rest_text = opt_text[2:].strip()
            add_textbox(slide, pos_x + 0.2, pos_y + 0.45, 0.9, 0.5, letter, size_pt=18, bold=True, color_hex="FFFFFF", alignment=PP_ALIGN.CENTER)
            add_textbox(slide, pos_x + 1.25, pos_y + 0.20, card_w - 1.4, card_h - 0.35, rest_text, size_pt=14, bold=False, color_hex=pal["text_on_card"])

def build_thanks(prs, data, pal, layout):
    slide = prs.slides.add_slide(layout)
    add_slide_transition(slide, "cover")
    add_safe_shape(slide, MSO_SHAPE.RECTANGLE, 0, SAFE_TOP, SLIDE_W, SAFE_BOTTOM - SAFE_TOP, pal["primary"], send_to_back=True)
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.8, 1.65, 9.73, 4.15, "FFFFFF", border_hex=pal["accent"])
    add_safe_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.8, 1.65, 9.73, 0.25, pal["accent"])
    
    add_textbox(slide, 2.0, 2.05, 9.33, 0.7, data.get("title", "CẢM ƠN CÁC EM ĐÃ CHÚ Ý LẮNG NGHE! 🌟"), size_pt=23, bold=True, color_hex=pal["primary"], alignment=PP_ALIGN.CENTER)
    
    content = data.get("content", "")
    lines = content.split('\n')
    cur_y = 2.85
    for l in lines:
        if l.strip():
            add_textbox(slide, 2.2, cur_y, 8.93, 0.60, l.strip(), size_pt=16, bold=False, color_hex=pal["text_on_card"], alignment=PP_ALIGN.CENTER)
            cur_y += 0.60
            
    add_textbox(slide, 2.0, 4.85, 9.33, 0.5, "Chúc các em có những giờ học Tin học thật hào hứng và bổ ích! 🚀", size_pt=15, bold=True, color_hex=pal["accent"], alignment=PP_ALIGN.CENTER)

# ─── BẢNG MÀU PHÙ HỢP TỪNG KHỐI LỚP ───
PAL_L5_T1 = {"primary": "BE185D", "accent": "EC4899", "bg": "FDF2F8", "card": "FFFFFF", "text_on_card": "0F172A", "text_on_bg": "831843"}
PAL_L5_T2 = {"primary": "9F1239", "accent": "F43F5E", "bg": "FFF1F2", "card": "FFFFFF", "text_on_card": "0F172A", "text_on_bg": "881337"}

PAL_L7_T1 = {"primary": "0369A1", "accent": "0EA5E9", "bg": "F0F9FF", "card": "FFFFFF", "text_on_card": "0F172A", "text_on_bg": "0C4A6E"}
PAL_L7_T2 = {"primary": "0E7490", "accent": "06B6D4", "bg": "ECFEFF", "card": "FFFFFF", "text_on_card": "0F172A", "text_on_bg": "164E63"}

CROP5_DIR = r"D:\UNIGO\SGK\Lớp_5\bai3_images\crops"
INFO5_DIR = r"D:\UNIGO\SGK\Lớp_5\bai3_images\infographics"

CROP7_DIR = r"D:\UNIGO\SGK\Lớp_7\bai3_images\crops"
INFO7_DIR = r"D:\UNIGO\SGK\Lớp_7\bai3_images\infographics"

# ─── CẤU HÌNH BỘ DECK CHO LỚP 5 VÀ LỚP 7 (TUẦN 6) ───

DECKS = [
    # ═══════════════════════════════════════════════════════════════════════
    # 1. LỚP 5 - TIẾT 1 (PPCT: 5)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lớp_5_Tiet05_Bai_3_Tim_kiem_thong_tin_Tiet_1.pptx",
        "pal": PAL_L5_T1,
        "slides": [
            ("cover", {
                "title": "BÀI 3. TÌM KIẾM THÔNG TIN\nTRONG GIẢI QUYẾT VẤN ĐỀ (TIẾT 1)",
                "subtitle": "Môn Tin học • Lớp 5 • Tuần 6 • Sách Kết nối tri thức với cuộc sống"
            }),
            ("section_banner", {
                "badge": "MỤC TIÊU TIẾT HỌC",
                "title": "Sau tiết 1 hôm nay, các em sẽ nắm vững:",
                "desc": "✓ Giải thích được sự cần thiết và tầm quan trọng của việc thu thập, tìm kiếm thông tin.\n✓ Hiểu được có đầy đủ thông tin sẽ giúp em ra quyết định giải quyết vấn đề tốt hơn.\n✓ Biết cách hợp tác nhóm, thảo luận và phân công công việc dựa vào thế mạnh của mỗi bạn."
            }),
            ("full_hero", {
                "badge": "HOẠT ĐỘNG 1: KHỞI ĐỘNG",
                "title": "Tình huống: Chuẩn bị hành lý kì nghỉ hè của An, Khoa, Minh",
                "img": os.path.join(CROP5_DIR, "lop5_khoi_dong.png"),
                "question": "👉 An đi biển Nha Trang, Khoa đi Đà Lạt mát lạnh, Minh đi Úc vào mùa đông tuyết rơi.\nCả 3 bạn đều xếp đồ bơi và quần áo mùa hè vào va li. Em đoán xem kì nghỉ của ai sẽ gặp khó khăn?"
            }),
            ("hero_example", {
                "badge": "KHÁM PHÁ 1",
                "title": "Sự cần thiết của việc thu thập thông tin trước khi hành động",
                "text": (
                    "• Quan sát Hình 15 & Hình 16 SGK:\n"
                    "  - An đi biển Nha Trang (nắng nóng quanh năm) ➡️ Đồ mùa hè rất phù hợp!\n"
                    "  - Minh đi Úc: Lúc này ở Úc là MÙA ĐÔNG tuyết rơi ➡️ Minh bị lạnh co ro vì không có quần áo ấm!\n"
                    "  - Khoa đi Đà Lạt: Ban đêm rất lạnh ➡️ Thiếu áo ấm để đi dạo phố.\n\n"
                    "👉 Bài học rút ra:\n"
                    "Thiếu thông tin sẽ gây rắc rối. Có đủ thông tin giúp giải quyết vấn đề thành công!"
                ),
                "img": os.path.join(CROP5_DIR, "lop5_hinh15_16_an_minh.png")
            }),
            ("conclusion_orange", {
                "question": "Thu thập và tìm kiếm thông tin có vai trò quan trọng như thế nào?",
                "suggestions": [
                    "• Để giải quyết vấn đề thực tế, em cần đưa ra quyết định đúng đắn.",
                    "• Muốn có quyết định chính xác, bắt buộc phải có thông tin đầy đủ và tin cậy."
                ],
                "img": os.path.join(CROP5_DIR, "lop5_hinh15_16_an_minh.png"),
                "conclusion": "• Thu thập và tìm kiếm thông tin là cần thiết và quan trọng trong giải quyết vấn đề.\n• Có đầy đủ thông tin sẽ giúp em giải quyết vấn đề tốt hơn."
            }),
            ("hero_example", {
                "badge": "KHÁM PHÁ 2",
                "title": "Hợp tác nhóm để cùng nhau giải quyết vấn đề 🤝",
                "text": (
                    "• Nhiệm vụ: Nhóm 3 bạn hoàn thành bài trình chiếu kì nghỉ hè.\n\n"
                    "• Phân công thông minh theo thế mạnh:\n"
                    "  - Mỗi bạn chuẩn bị nội dung và hình ảnh trang của mình.\n"
                    "  - An chỉnh sửa bài viết văn bản cho chuẩn xác.\n"
                    "  - Minh chọn và chỉnh sửa hình ảnh cho sinh động.\n"
                    "  - Khoa ghép các trang chiếu và hoàn thiện bài.\n\n"
                    "👉 Cả nhóm cùng vui vẻ, hoàn thành bài xuất sắc!"
                ),
                "img": os.path.join(CROP5_DIR, "lop5_hoatdong2_nhom.png")
            }),
            ("conclusion_orange", {
                "question": "Làm thế nào để làm việc nhóm hiệu quả và gắn kết?",
                "suggestions": [
                    "• Không để một bạn làm hộ cả nhóm, cũng không làm riêng lẻ rời rạc.",
                    "• Chia nhỏ công việc và trao đổi giúp đỡ nhau trong quá trình thực hiện."
                ],
                "img": os.path.join(CROP5_DIR, "lop5_hoatdong2_nhom.png"),
                "conclusion": "• Khi làm việc nhóm, cần thảo luận, phân công công việc dựa vào điểm mạnh của mỗi người.\n• Khi được phân công, cần có ý thức, tinh thần trách nhiệm; hợp tác và giúp đỡ nhau."
            }),
            ("quiz_mascot", {
                "question": "Để chuẩn bị cho chuyến đi tham quan cùng lớp, em cần những thông tin nào? (SGK tr.15)",
                "options": [
                    "A. Vị trí điểm tham quan",
                    "B. Dự báo thời tiết ngày đi",
                    "C. Chương trình và hoạt động",
                    "D. Cần tất cả các thông tin trên"
                ]
            }),
            ("quiz_mascot", {
                "question": "Trong hoạt động làm việc nhóm, phát biểu nào sau đây là ĐÚNG NHẤT? (SGK tr.16)",
                "options": [
                    "A. Bạn giỏi nhất nên làm hộ cả nhóm",
                    "B. Mỗi bạn cần có tinh thần trách nhiệm",
                    "C. Ai thích làm thì làm, không ép buộc",
                    "D. Chỉ cần nhóm trưởng báo cáo kết quả"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH XUẤT SẮC TIẾT 1! 🌟",
                "content": (
                    "1. Ghi nhớ 2 Hộp Chốt kiến thức quan trọng của Tiết 1.\n"
                    "2. Chuẩn bị cho Tiết 2: Thực hành tìm kiếm thắng cảnh Sa Pa trên Google!\n"
                    "3. Chúc các em luôn có tinh thần hợp tác nhóm tuyệt vời!"
                )
            })
        ]
    },

    # ═══════════════════════════════════════════════════════════════════════
    # 2. LỚP 5 - TIẾT 2 (PPCT: 6)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lớp_5_Tiet06_Bai_3_Tim_kiem_thong_tin_Tiet_2.pptx",
        "pal": PAL_L5_T2,
        "slides": [
            ("cover", {
                "title": "BÀI 3. TÌM KIẾM THÔNG TIN\nTRONG GIẢI QUYẾT VẤN ĐỀ (TIẾT 2)",
                "subtitle": "Môn Tin học • Lớp 5 • Tuần 6 • Sách Kết nối tri thức với cuộc sống"
            }),
            ("section_banner", {
                "badge": "MỤC TIÊU TIẾT HỌC",
                "title": "Sau tiết 2 hôm nay, các em sẽ nắm vững:",
                "desc": "✓ Nắm vững Quy trình 4 bước tìm kiếm thông tin trên Internet bằng máy tìm kiếm.\n✓ Sử dụng từ khoá chính xác đặt trong dấu ngoặc kép \" \" để thu hẹp kết quả.\n✓ Lựa chọn thông tin phù hợp theo hoàn cảnh thực tế (địa điểm, thời tiết, thành viên gia đình)."
            }),
            ("full_hero", {
                "badge": "KHỞI ĐỘNG TIẾT 2",
                "title": "Tình huống: Lập kế hoạch du lịch Sa Pa dịp Tết Dương lịch 🏔️",
                "img": os.path.join(CROP5_DIR, "lop5_hinh17_tim_kiem_sapa.png"),
                "question": "👉 Gia đình em dự định đi Sa Pa vào dịp Tết. Có rất nhiều thắng cảnh đẹp.\nLàm thế nào để tìm kiếm thông tin nhanh chóng và chọn được những địa điểm phù hợp nhất cho cả nhà?"
            }),
            ("process_diagram", {
                "badge": "HƯỚNG DẪN QUY TRÌNH",
                "title": "Quy trình 4 bước tìm kiếm thông tin chuẩn xác trên Internet",
                "badges": [
                    "Bước 1: Xác định thông tin",
                    "Bước 2: Chọn từ khoá \" \"",
                    "Bước 3: Chọn nguồn tin cậy",
                    "Bước 4: Tổng hợp & Áp dụng"
                ],
                "img": os.path.join(INFO5_DIR, "quy_trinh_4_buoc_tim_kiem.png")
            }),
            ("hero_example", {
                "badge": "THỰC HÀNH NHIỆM VỤ 1",
                "title": "Tìm kiếm các điểm tham quan Sa Pa trên Google",
                "text": (
                    "• Thao tác trên máy tính:\n"
                    "  - Mở trình duyệt web (Google Chrome hoặc Cốc Cốc).\n"
                    "  - Truy cập địa chỉ máy tìm kiếm: google.com\n"
                    "  - Gõ từ khoá: Các điểm tham quan ở Sa Pa (nhấn Enter).\n\n"
                    "• Quan sát kết quả (Hình 17 SGK):\n"
                    "  - Bản Cát Cát, Nhà thờ đá Sa Pa, Đèo Ô Quy Hồ, Đỉnh Fansipan, Bãi đá cổ, Núi Hàm Rồng...\n"
                    "  - Nháy chuột vào từng kết quả để đọc thông tin chi tiết!"
                ),
                "img": os.path.join(CROP5_DIR, "lop5_hinh17_tim_kiem_sapa.png")
            }),
            ("hero_example", {
                "badge": "THỰC HÀNH NHIỆM VỤ 2",
                "title": "Lựa chọn điểm đến thông minh theo đặc điểm chuyến đi",
                "text": (
                    "• Phân loại điểm tham quan theo thực tế:\n"
                    "  - Điểm gần trung tâm: Nhà thờ đá Sa Pa, hồ Sa Pa, chợ đêm...\n"
                    "  - Điểm phải đi xa/hiểm trở: Đèo Ô Quy Hồ, Cầu Mây, thác Bạc...\n"
                    "  - Điểm leo trèo, lội suối: Đỉnh Fansipan, thung lũng Mường Hoa...\n\n"
                    "👉 Lưu ý quan trọng:\n"
                    "Nếu đoàn có người già và trẻ em nhỏ, nên ưu tiên chọn điểm ở gần, đường bằng phẳng, tránh leo dốc trơn trượt!"
                ),
                "img": os.path.join(CROP5_DIR, "lop5_hinh17_tim_kiem_sapa.png")
            }),
            ("conclusion_orange", {
                "question": "Bài học lớn nhất khi tìm kiếm và lựa chọn thông tin là gì?",
                "suggestions": [
                    "• Không phải thông tin nào tìm thấy cũng phù hợp với nhu cầu của mình.",
                    "• Cần đọc kỹ, đối chiếu và lựa chọn giải pháp an toàn, phù hợp nhất với hoàn cảnh."
                ],
                "img": os.path.join(INFO5_DIR, "quy_trinh_4_buoc_tim_kiem.png"),
                "conclusion": "Tìm kiếm bằng từ khoá chính xác và chọn lọc thông tin phù hợp với điều kiện thực tế sẽ giúp em đưa ra quyết định sáng suốt và giải quyết vấn đề hiệu quả nhất!"
            }),
            ("quiz_mascot", {
                "question": "Trong các phát biểu sau, đâu là PHƯƠNG ÁN SAI về vai trò của thông tin? (SGK tr.18)",
                "options": [
                    "A. Để giải quyết vấn đề cần đưa ra quyết định",
                    "B. Để có quyết định đúng cần có thông tin",
                    "C. Thông tin không có vai trò gì quan trọng",
                    "D. Thu thập thông tin rất quan trọng trong đời sống"
                ]
            }),
            ("quiz_mascot", {
                "question": "Địa chỉ website chính thức của Sở Du lịch Hà Nội để tìm kiếm thông tin Thủ đô là gì?",
                "options": [
                    "A. sodulich.hanoi.gov.vn",
                    "B. facebook.com/dulichhanoi",
                    "C. dulichgiare.net",
                    "D. thongtintintuc24h.com"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH CHỦ ĐỀ 3 - BÀI 3! 🚀",
                "content": (
                    "1. Thực hành ở nhà: Giúp mẹ tìm hiểu thông tin một bộ truyện tranh hay để mua tặng sinh nhật.\n"
                    "2. Nhớ ghi lại tên truyện, giá tiền và địa chỉ bán uy tín vào tệp văn bản nhé.\n"
                    "3. Tắt máy tính an toàn, xếp gọn bàn ghế phòng học trước khi ra về!"
                )
            })
        ]
    },

    # ═══════════════════════════════════════════════════════════════════════
    # 3. LỚP 5 - BỘ SLIDE TỔNG HỢP TUẦN 6 (TIẾT 5 & 6)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_06\Slide_Tin_hoc_Lop_5_Bai03_Tuan06.pptx",
        "pal": PAL_L5_T1,
        "slides": [
            ("cover", {
                "title": "BÀI 3. TÌM KIẾM THÔNG TIN\nTRONG GIẢI QUYẾT VẤN ĐỀ",
                "subtitle": "Tin học • Lớp 5 • Tuần 6 (Trọn bộ Tiết 5 & Tiết 6)"
            }),
            ("section_banner", {
                "badge": "TIẾT 1: LÝ THUYẾT & TÌNH HUỐNG",
                "title": "SỰ CẦN THIẾT CỦA THÔNG TIN & HỢP TÁC NHÓM",
                "desc": "✓ Hiểu tầm quan trọng của việc tìm kiếm thông tin trước khi hành động.\n✓ Phân công công việc nhóm khoa học theo điểm mạnh của mỗi người."
            }),
            ("full_hero", {
                "badge": "TÌNH HUỐNG KHỞI ĐỘNG",
                "title": "Chuẩn bị hành lý: An đi biển Nha Trang, Minh đi Úc mùa đông tuyết",
                "img": os.path.join(CROP5_DIR, "lop5_hinh15_16_an_minh.png"),
                "question": "👉 Minh bị lạnh co ro vì thiếu thông tin thời tiết ở Úc lúc đó là mùa đông.\nCó đầy đủ thông tin sẽ giúp chúng ta giải quyết vấn đề tốt hơn rất nhiều!"
            }),
            ("conclusion_orange", {
                "question": "Ghi nhớ trọng tâm Tiết 1",
                "suggestions": [
                    "• Thu thập và tìm kiếm thông tin là cần thiết và quan trọng trong giải quyết vấn đề.",
                    "• Khi làm việc nhóm, phân công theo điểm mạnh và có tinh thần trách nhiệm giúp đỡ nhau."
                ],
                "img": os.path.join(CROP5_DIR, "lop5_hoatdong2_nhom.png"),
                "conclusion": "• Thu thập và tìm kiếm thông tin là cần thiết và quan trọng trong giải quyết vấn đề.\n• Có đầy đủ thông tin sẽ giúp em giải quyết vấn đề tốt hơn."
            }),
            ("section_banner", {
                "badge": "TIẾT 2: THỰC HÀNH TÌM KIẾM",
                "title": "QUY TRÌNH 4 BƯỚC & CHỌN LỌC THÔNG TIN SA PA",
                "desc": "✓ Thực hành sử dụng máy tìm kiếm Google với từ khoá chính xác.\n✓ Lựa chọn điểm đến phù hợp theo đặc điểm đối tượng trong đoàn."
            }),
            ("process_diagram", {
                "badge": "QUY TRÌNH 4 BƯỚC",
                "title": "Sơ đồ 4 bước tìm kiếm thông tin chuẩn xác trên Internet",
                "badges": [
                    "Bước 1: Xác định thông tin",
                    "Bước 2: Chọn từ khoá \" \"",
                    "Bước 3: Chọn nguồn tin cậy",
                    "Bước 4: Tổng hợp & Áp dụng"
                ],
                "img": os.path.join(INFO5_DIR, "quy_trinh_4_buoc_tim_kiem.png")
            }),
            ("hero_example", {
                "badge": "KẾT QUẢ THỰC HÀNH",
                "title": "Kết quả tìm kiếm và lựa chọn điểm tham quan Sa Pa",
                "text": (
                    "• Tìm kiếm từ khoá: Các điểm tham quan ở Sa Pa\n"
                    "• Kết quả: Nhà thờ đá, đèo Ô Quy Hồ, đỉnh Fansipan, Cát Cát...\n"
                    "• Đưa ra quyết định thông minh:\n"
                    "  - Chọn điểm gần, an toàn cho người già và trẻ nhỏ.\n"
                    "  - Lên lịch trình hợp lý theo dự báo thời tiết thực tế!"
                ),
                "img": os.path.join(CROP5_DIR, "lop5_hinh17_tim_kiem_sapa.png")
            }),
            ("quiz_mascot", {
                "question": "Để tìm kiếm thông tin chính xác cụm từ về Sa Pa, em nên đặt từ khoá trong dấu gì?",
                "options": [
                    "A. Dấu ngoặc đơn ( )",
                    "B. Dấu ngoặc kép \" \"",
                    "C. Dấu ngoặc nhọn { }",
                    "D. Dấu gạch chéo / /"
                ]
            }),
            ("thanks", {
                "title": "BÀI HỌC TUẦN 6 KẾT THÚC! 🌟",
                "content": "Chúc các em luôn là những học sinh UNIGO năng động, biết tìm kiếm và hợp tác thông minh!"
            })
        ]
    },

    # ═══════════════════════════════════════════════════════════════════════
    # 4. LỚP 7 - TIẾT 1 (PPCT: 5)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lớp_7_Tiet05_Bai_3_Quan_ly_du_lieu_trong_may_tinh_Tiet_1.pptx",
        "pal": PAL_L7_T1,
        "slides": [
            ("cover", {
                "title": "BÀI 3. QUẢN LÝ DỮ LIỆU\nTRONG MÁY TÍNH (TIẾT 1)",
                "subtitle": "Môn Tin học • Lớp 7 • Tuần 6 • Sách Kết nối tri thức với cuộc sống"
            }),
            ("section_banner", {
                "badge": "MỤC TIÊU TIẾT HỌC",
                "title": "Sau tiết 1 hôm nay, các em sẽ nắm vững:",
                "desc": "✓ Hiểu cấu trúc tên tệp gồm Phần tên và Phần mở rộng (.docx, .xlsx, .pptx, .exe, .bat...).\n✓ Nguyên tắc đặt tên tệp và thư mục khoa học, dễ nhớ, dễ quản lý.\n✓ Biết được tệp chương trình cũng là dữ liệu và hiểu sự cần thiết phải sao lưu dữ liệu (Backup)."
            }),
            ("full_hero", {
                "badge": "HOẠT ĐỘNG 1: KHỞI ĐỘNG",
                "title": "Tình huống: Quản lý hàng trăm bức ảnh du lịch cùng gia đình 📸",
                "img": os.path.join(CROP7_DIR, "lop7_hinh31_thumuc_tep.png"),
                "question": "👉 Trong chuyến du lịch, em chụp rất nhiều ảnh và ghi chép nhật ký.\nNếu để toàn bộ ra màn hình nền Desktop thì sẽ như thế nào? Cần tổ chức cây thư mục ra sao để dễ tìm kiếm?"
            }),
            ("hero_example", {
                "badge": "KHÁM PHÁ 1",
                "title": "Tên tệp và Thư mục trong hệ điều hành máy tính",
                "text": (
                    "• Quan sát Hình 3.1 SGK:\n"
                    "  - Đường dẫn: This PC > OS (C:) > DuLich > DiemDen\n"
                    "  - Thư mục con: BanNgay, BinhMinh, HoangHon, GiaDinh.\n"
                    "  - Tệp hình ảnh: PhaoHoa01.jpg, PhaoHoa02.jpg.\n\n"
                    "• Quy tắc vàng khi đặt tên:\n"
                    "  - Đặt tên ngắn gọn, dễ nhớ, phản ánh đúng nội dung bên trong.\n"
                    "  - Viết không dấu hoặc cách nhau bởi dấu gạch dưới để tránh lỗi hệ thống."
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh31_thumuc_tep.png")
            }),
            ("process_diagram", {
                "badge": "PHÂN LOẠI TỆP TIN",
                "title": "Nhận biết phần mở rộng của tệp dữ liệu & tệp chương trình",
                "badges": [
                    "Văn bản: .docx, .pdf",
                    "Đa phương tiện: .jpg, .mp4",
                    "Chương trình: .exe, .bat",
                    "Cảnh báo: Không xoá tệp lạ"
                ],
                "img": os.path.join(INFO7_DIR, "cau_truc_tep_va_phan_mo_rong.png")
            }),
            ("conclusion_orange", {
                "question": "Tệp chương trình máy tính có đặc điểm gì cần lưu ý?",
                "suggestions": [
                    "• Tệp chương trình máy tính được lưu trữ dưới dạng tệp như .exe, .com, .bat, .msi.",
                    "• Không nên xoá hay di chuyển những tệp này nếu không có lý do xác đáng."
                ],
                "img": os.path.join(CROP7_DIR, "lop7_hinh31_thumuc_tep.png"),
                "conclusion": "• Tên tệp và thư mục cần được đặt sao cho dễ nhớ, cho ta biết trong đó chứa những gì.\n• Chương trình máy tính được lưu trữ trên thiết bị nhớ giống như một tệp dữ liệu, thường có phần mở rộng .exe, .com, .bat, .msi."
            }),
            ("hero_example", {
                "badge": "KHÁM PHÁ 2",
                "title": "Các biện pháp bảo vệ dữ liệu & Thiết bị lưu trữ",
                "text": (
                    "• Vì sao phải Sao lưu dữ liệu (Backup)?\n"
                    "  - Máy tính có thể bị hỏng ổ cứng, bị mất cắp hoặc nhiễm mã độc.\n\n"
                    "• Hai hình thức sao lưu phổ biến:\n"
                    "  1. Sao lưu cục bộ: Lưu sang ổ D, USB, hoặc ổ cứng di động gắn ngoài.\n"
                    "  2. Sao lưu từ xa: Lưu lên đám mây (Google Drive, OneDrive).\n"
                    "  👉 Sao lưu từ xa giúp dữ liệu an toàn ngay cả khi máy tính bị hỏng hoàn toàn!"
                ),
                "img": os.path.join(CROP7_DIR, "lop7_bang31_thietbi_luutru.png")
            }),
            ("conclusion_orange", {
                "question": "Thiết bị lưu trữ nào tiện dụng và an toàn nhất hiện nay?",
                "suggestions": [
                    "• USB/Thẻ nhớ: Nhỏ gọn nhưng dễ bị rơi, dễ lây lan virus.",
                    "• Ổ cứng ngoài: Dung lượng lớn, phù hợp sao lưu trọn gói máy tính.",
                    "• Đám mây (Cloud): Truy cập mọi lúc mọi nơi có Internet, chống mất mát vật lý."
                ],
                "img": os.path.join(CROP7_DIR, "lop7_bang31_thietbi_luutru.png"),
                "conclusion": "Dữ liệu cần được sao lưu thường xuyên lên thiết bị lưu trữ ngoài máy tính chứa dữ liệu gốc để tránh bị mất hoặc bị hỏng dữ liệu!"
            }),
            ("quiz_mascot", {
                "question": "Tệp có phần mở rộng .exe thuộc loại tệp nào sau đây? (SGK tr.14)",
                "options": [
                    "A. Tệp dữ liệu Word",
                    "B. Tệp chương trình máy tính",
                    "C. Tệp bài hát âm thanh",
                    "D. Tệp hình ảnh đồ hoạ"
                ]
            }),
            ("quiz_mascot", {
                "question": "Thiết bị lưu trữ nào cho phép em truy cập dữ liệu ở bất kỳ máy tính nào có kết nối Internet? (SGK tr.15)",
                "options": [
                    "A. Đĩa quang CD-ROM",
                    "B. Ổ cứng gắn trong",
                    "C. Lưu trữ công nghệ đám mây",
                    "D. Thẻ nhớ điện thoại"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH TIẾT 1! 🌟",
                "content": (
                    "1. Ôn lại cấu trúc phần mở rộng của các loại tệp phổ biến (.docx, .pptx, .exe).\n"
                    "2. Chuẩn bị cho Tiết 2: Kỹ thuật đặt mật khẩu mạnh & 5 thao tác File Explorer thực chiến!\n"
                    "3. Chúc các em luôn có ý thức bảo vệ tài sản số của mình!"
                )
            })
        ]
    },

    # ═══════════════════════════════════════════════════════════════════════
    # 5. LỚP 7 - TIẾT 2 (PPCT: 6)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lớp_7_Tiet06_Bai_3_Quan_ly_du_lieu_trong_may_tinh_Tiet_2.pptx",
        "pal": PAL_L7_T2,
        "slides": [
            ("cover", {
                "title": "BÀI 3. QUẢN LÝ DỮ LIỆU\nTRONG MÁY TÍNH (TIẾT 2)",
                "subtitle": "Môn Tin học • Lớp 7 • Tuần 6 • Sách Kết nối tri thức với cuộc sống"
            }),
            ("section_banner", {
                "badge": "MỤC TIÊU TIẾT HỌC",
                "title": "Sau tiết 2 hôm nay, các em sẽ nắm vững:",
                "desc": "✓ Thiết lập mật khẩu mạnh và nhận diện các phần mềm diệt virus bảo vệ máy tính.\n✓ Thực hành thành thạo 5 thao tác quản lý dữ liệu trên File Explorer (Tạo mới, Đổi tên, Sao chép, Di chuyển, Xoá).\n✓ Xây dựng cây thư mục học tập khoa học và hình thành thói quen sao lưu định kỳ."
            }),
            ("hero_example", {
                "badge": "KHÁM PHÁ 1: AN TOÀN SỐ",
                "title": "Tài khoản người dùng, Mật khẩu mạnh & Phần mềm diệt virus",
                "text": (
                    "• Tiêu chuẩn của một MẬT KHẨU MẠNH:\n"
                    "  - Độ dài: Ít nhất 8 kí tự.\n"
                    "  - Kết hợp: Chữ in hoa, chữ thường, chữ số và kí hiệu đặc biệt (@, #, $, %).\n"
                    "  - Không dùng: Ngày sinh, tên riêng, số điện thoại, \"12345678\".\n"
                    "  👉 Ví dụ mạnh: 2n#M1nhKh0a\n\n"
                    "• Phần mềm diệt virus (Hình 3.2 SGK):\n"
                    "  - Windows Defender, Bkav, Kaspersky, Avast, Avira...\n"
                    "  - Luôn bật chế độ bảo vệ tự động và quét virus định kỳ!"
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh32_diet_virus.png")
            }),
            ("conclusion_orange", {
                "question": "Nguyên tắc vàng để bảo vệ dữ liệu khỏi mã độc và tin tặc là gì?",
                "suggestions": [
                    "• Đặt mật khẩu mạnh và không chia sẻ mật khẩu cho người khác.",
                    "• Không mở tệp đính kèm lạ, không cài phần mềm crack không rõ nguồn gốc."
                ],
                "img": os.path.join(CROP7_DIR, "lop7_hinh32_diet_virus.png"),
                "conclusion": "• Việc đặt mật khẩu cho tài khoản người dùng trên máy tính và Internet giúp bảo vệ dữ liệu khỏi sự truy cập trái phép.\n• Cần bảo vệ dữ liệu bằng cách luôn bật chế độ bảo vệ của phần mềm diệt virus và không dùng phần mềm lạ."
            }),
            ("process_diagram", {
                "badge": "QUY TRÌNH THAO TÁC",
                "title": "5 Thao tác quản lý dữ liệu cốt lõi trên File Explorer",
                "badges": [
                    "1. Tạo mới: New Folder",
                    "2. Đổi tên: Rename (F2)",
                    "3. Sao chép: Copy & Paste",
                    "4. Di chuyển: Cut & Paste",
                    "5. Xoá bỏ: Delete"
                ],
                "img": os.path.join(INFO7_DIR, "5_thao_tac_quan_ly_du_lieu.png")
            }),
            ("hero_example", {
                "badge": "THỰC HÀNH TẠO THƯ MỤC",
                "title": "Thực hành tạo và đổi tên thư mục (Hình 3.4 & 3.5 SGK)",
                "text": (
                    "• Các bước thực hiện:\n"
                    "  1. Nháy chuột phải vào vùng trống ➡️ Chọn New ➡️ Folder.\n"
                    "  2. Nhập tên thư mục mới (ví dụ: DuLich, HocTap).\n"
                    "  3. Nhấn phím Enter để hoàn tất.\n\n"
                    "• Thao tác đổi tên nhanh:\n"
                    "  - Nháy chọn thư mục ➡️ Nhấn phím F2.\n"
                    "  - Gõ tên mới và nhấn Enter."
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh34_35_tao_thumuc.png")
            }),
            ("hero_example", {
                "badge": "THỰC HÀNH NÂNG CAO",
                "title": "Thực hành Đổi tên, Di chuyển, Sao chép và Xoá (Hình 3.6 SGK)",
                "text": (
                    "• Menu ngữ cảnh chuột phải (Hình 3.6 SGK):\n"
                    "  - Cut (Ctrl + X): Chuyển thư mục vào bộ nhớ tạm để di chuyển.\n"
                    "  - Copy (Ctrl + C): Sao chép thư mục vào bộ nhớ tạm.\n"
                    "  - Paste (Ctrl + V): Dán thư mục từ bộ nhớ tạm ra vị trí mới.\n"
                    "  - Delete: Xoá thư mục chuyển vào Thùng rác (Recycle Bin).\n\n"
                    "👉 Nhiệm vụ: Di chuyển các tệp từ BinhMinh sang BanNgay, sau đó xoá thư mục rác!"
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh36_menu_thaotac.png")
            }),
            ("hero_example", {
                "badge": "SƠ ĐỒ CÂY THƯ MỤC",
                "title": "Sơ đồ Cây thư mục phân cấp hoàn chỉnh (Hình 3.3 SGK)",
                "text": (
                    "• Cấu trúc cây dữ liệu:\n"
                    "  - Gốc: Ổ đĩa C:\\\n"
                    "  - Nhánh lớn: DuLich, HocTap, GiaDinh.\n"
                    "  - Nhánh nhỏ trong DuLich: DiemDen (chứa BanNgay, HoangHon, PhaoHoa01.jpg, PhaoHoa02.jpg).\n\n"
                    "👉 Mô hình cây thư mục giúp sắp xếp dữ liệu ngăn nắp, tìm kiếm trong 1 giây!"
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh33_cay_thumuc.png")
            }),
            ("quiz_mascot", {
                "question": "Trong các mật khẩu sau đây, đâu là mật khẩu MẠNH NHẤT? (SGK tr.16)",
                "options": [
                    "A. 12345678",
                    "B. AnMinhKhoa",
                    "C. matkhau2026",
                    "D. 2n#M1nhKh0a"
                ]
            }),
            ("quiz_mascot", {
                "question": "Ứng dụng nào trong hệ điều hành Windows giúp em quản lý tệp và thư mục? (SGK tr.17)",
                "options": [
                    "A. Internet Explorer",
                    "B. Microsoft Word",
                    "C. File Explorer",
                    "D. Windows Media Player"
                ]
            }),
            ("thanks", {
                "title": "HOÀN THÀNH XUẤT SẮC BÀI 3! 🚀",
                "content": (
                    "1. Kiểm tra lại cây thư mục học tập trên máy tính của em và sắp xếp lại ngăn nắp.\n"
                    "2. Thiết lập mật khẩu mạnh bảo vệ tài khoản số cá nhân.\n"
                    "3. Tắt máy tính an toàn, vệ sinh chỗ ngồi trước khi rời phòng máy!"
                )
            })
        ]
    },

    # ═══════════════════════════════════════════════════════════════════════
    # 6. LỚP 7 - BỘ SLIDE TỔNG HỢP TUẦN 6 (TIẾT 5 & 6)
    # ═══════════════════════════════════════════════════════════════════════
    {
        "file_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_06\Slide_Tin_hoc_Lop_7_Bai03_Tuan06.pptx",
        "pal": PAL_L7_T1,
        "slides": [
            ("cover", {
                "title": "BÀI 3. QUẢN LÝ DỮ LIỆU\nTRONG MÁY TÍNH",
                "subtitle": "Tin học • Lớp 7 • Tuần 6 (Trọn bộ Tiết 5 & Tiết 6)"
            }),
            ("section_banner", {
                "badge": "TIẾT 1: CẤU TRÚC TỆP & SAO LƯU",
                "title": "PHẦN MỞ RỘNG TỆP & CÁC THIẾT BỊ LƯU TRỮ",
                "desc": "✓ Nắm vững phần mở rộng (.docx, .xlsx, .exe, .bat...).\n✓ Sao lưu dữ liệu dự phòng (Backup) lên USB và Đám mây (Cloud)."
            }),
            ("process_diagram", {
                "badge": "CẤU TRÚC PHẦN MỞ RỘNG",
                "title": "Nhận biết các loại tệp dữ liệu & tệp chương trình trong Windows",
                "badges": [
                    "Văn bản: .docx, .pdf",
                    "Đa phương tiện: .jpg, .mp4",
                    "Chương trình: .exe, .bat",
                    "Cảnh báo: Không xoá tệp lạ"
                ],
                "img": os.path.join(INFO7_DIR, "cau_truc_tep_va_phan_mo_rong.png")
            }),
            ("conclusion_orange", {
                "question": "Ghi nhớ trọng tâm Tiết 1",
                "suggestions": [
                    "• Đặt tên tệp, thư mục dễ nhớ và phản ánh đúng nội dung.",
                    "• Thường xuyên sao lưu dữ liệu sang thiết bị ngoài hoặc đám mây."
                ],
                "img": os.path.join(CROP7_DIR, "lop7_bang31_thietbi_luutru.png"),
                "conclusion": "• Tên tệp và thư mục cần được đặt sao cho dễ nhớ, cho ta biết trong đó chứa những gì.\n• Dữ liệu cần được sao lưu thường xuyên lên thiết bị lưu trữ ngoài hoặc đám mây để tránh bị mất mát."
            }),
            ("section_banner", {
                "badge": "TIẾT 2: BẢO MẬT & THỰC HÀNH",
                "title": "MẬT KHẨU MẠNH, DIỆT VIRUS & 5 THAO TÁC CỐT LÕI",
                "desc": "✓ Thiết lập mật khẩu mạnh và phần mềm diệt virus.\n✓ Thực hành thành thạo File Explorer: Tạo mới, Đổi tên, Sao chép, Di chuyển, Xoá."
            }),
            ("process_diagram", {
                "badge": "5 THAO TÁC CỐT LÕI",
                "title": "5 Thao tác quản lý tệp và thư mục trên File Explorer",
                "badges": [
                    "1. Tạo mới: New Folder",
                    "2. Đổi tên: Rename (F2)",
                    "3. Sao chép: Copy & Paste",
                    "4. Di chuyển: Cut & Paste",
                    "5. Xoá bỏ: Delete"
                ],
                "img": os.path.join(INFO7_DIR, "5_thao_tac_quan_ly_du_lieu.png")
            }),
            ("hero_example", {
                "badge": "SƠ ĐỒ CÂY DỮ LIỆU",
                "title": "Cây thư mục phân cấp hoàn chỉnh trong thực hành",
                "text": (
                    "• Ổ C:\\ ➡️ DuLich, HocTap, GiaDinh.\n"
                    "• Trong DuLich ➡️ DiemDen ➡️ BanNgay, HoangHon, PhaoHoa01.jpg...\n"
                    "• Thành thạo các phím tắt nhanh:\n"
                    "  - Ctrl + C (Copy), Ctrl + X (Cut), Ctrl + V (Paste)\n"
                    "  - F2 (Rename), Delete (Xoá)"
                ),
                "img": os.path.join(CROP7_DIR, "lop7_hinh33_cay_thumuc.png")
            }),
            ("quiz_mascot", {
                "question": "Trong các mật khẩu sau đây, đâu là mật khẩu MẠNH NHẤT? (SGK tr.16)",
                "options": [
                    "A. 12345678",
                    "B. AnMinhKhoa",
                    "C. matkhau2026",
                    "D. 2n#M1nhKh0a"
                ]
            }),
            ("thanks", {
                "title": "BÀI HỌC TUẦN 6 KẾT THÚC! 🌟",
                "content": "Chúc các em trở thành những công dân số thông thái, quản lý dữ liệu an toàn và hiệu quả!"
            })
        ]
    }
]

BUILDERS_MAP = {
    "cover": build_cover,
    "section_banner": build_section_banner,
    "hero_example": build_hero_example,
    "full_hero": build_full_hero_slide,
    "process_diagram": build_process_diagram_slide,
    "conclusion_orange": build_conclusion_orange,
    "quiz_mascot": build_quiz_mascot,
    "thanks": build_thanks
}

def generate_all_decks():
    print("=== TẠO BỘ SLIDE BÀI GIẢNG CHUẨN UNIGO - LỚP 5 & LỚP 7 (TUẦN 6) ===")
    created_files = []
    for config in DECKS:
        prs = Presentation(TEMPLATE)
        # Remove default template slides while keeping masters
        for i in range(len(prs.slides)-1, -1, -1):
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]
            
        blank_layout = prs.slide_layouts[6]
        pal = config["pal"]
        file_path = config["file_path"]
        
        print(f"\n--> Đang tạo: {os.path.basename(file_path)}")
        for b_name, b_data in config["slides"]:
            if b_name in BUILDERS_MAP:
                BUILDERS_MAP[b_name](prs, b_data, pal, blank_layout)
            else:
                print(f"Unknown builder: {b_name}")
                
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        prs.save(file_path)
        print(f"  [HOÀN TẤT] {file_path} ({len(prs.slides)} slides)")
        created_files.append(file_path)
    return created_files

if __name__ == "__main__":
    generate_all_decks()
