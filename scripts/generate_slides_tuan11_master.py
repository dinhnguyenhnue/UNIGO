# -*- coding: utf-8 -*-
"""
generate_slides_tuan11_master.py
Tạo 5 bộ Slide bài giảng (.pptx) chuẩn cho học sinh (Student-Facing) môn Tin học Tuần 11:
- Tiền tiểu học: Ôn tập - Bé làm chủ chuột
- Lớp 1: Ôn tập kỹ năng số
- Lớp 2: Ôn tập tệp – thư mục
- Lớp 3: Bài 4 - Làm việc với máy tính (Tiết 2)
- Lớp 4: Bài 4 - Tìm kiếm thông tin trên Internet (Tiết 2)

Tuân thủ nghiêm ngặt AGENTS.md, skill tao-slide-bai-giang:
- Bảo tồn Header/Logo và Footer chân trang UNIGO
- Vùng an toàn: Top 1.15in -> Bottom 6.35in, Left 0.3in -> Right 13.0in
- Visual-first, thẻ màu bo góc, font Arial/Segoe UI dễ đọc cho học sinh
- Chốt kiến thức hộp màu cam chuẩn mực
"""

import os, sys, copy, glob
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_PATH = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide UNIGO.pptx'
OUT_BASE_DIR = r'D:\UNIGO\KHBD_Tin_học'

# Safe zone constants
SAFE_TOP = 1.15
SAFE_BOTTOM = 6.35
SAFE_LEFT = 0.3
SAFE_RIGHT = 13.0
SAFE_WIDTH = SAFE_RIGHT - SAFE_LEFT
SAFE_HEIGHT = SAFE_BOTTOM - SAFE_TOP

def hex_to_rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

def set_font(run, size_pt, bold=False, italic=False, color_hex="333333", font_name="Arial"):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = hex_to_rgb(color_hex)
    run.font.name = font_name

def add_card(slide, left, top, width, height, fill_hex, border_hex=None, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE):
    actual_top = max(top, SAFE_TOP)
    actual_bottom = min(top + height, SAFE_BOTTOM)
    actual_height = max(actual_bottom - actual_top, 0.1)
    
    shape = slide.shapes.add_shape(
        shape_type, Inches(left), Inches(actual_top), Inches(width), Inches(actual_height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_to_rgb(border_hex)
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, left, top, width, height, text="", size_pt=18, bold=False, italic=False,
                color_hex="333333", align=PP_ALIGN.LEFT, font_name="Arial"):
    txbox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txbox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    if text:
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        set_font(run, size_pt, bold=bold, italic=italic, color_hex=color_hex, font_name=font_name)
    return txbox

def add_slide_header(slide, title_text, badge_text, palette):
    """Tiêu đề slide có badge nổi bật, nằm trong vùng an toàn Top 1.15in -> 1.85in"""
    # Header card
    h_card = add_card(slide, SAFE_LEFT, SAFE_TOP, SAFE_WIDTH, 0.65, palette["bg"], palette["primary"])
    
    # Badge
    badge_w = 2.4
    b_card = add_card(slide, SAFE_LEFT + 0.1, SAFE_TOP + 0.08, badge_w, 0.48, palette["primary"], None)
    tb_b = add_textbox(slide, SAFE_LEFT + 0.1, SAFE_TOP + 0.12, badge_w, 0.4, badge_text.upper(),
                       size_pt=13, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
    
    # Title text
    tb_t = add_textbox(slide, SAFE_LEFT + badge_w + 0.25, SAFE_TOP + 0.12, SAFE_WIDTH - badge_w - 0.4, 0.45,
                       title_text, size_pt=18, bold=True, color_hex=palette["text_on_bg"], align=PP_ALIGN.LEFT)

# ============================================================
# BUILDERS FOR DIFFERENT SLIDE TYPES
# ============================================================

def build_cover_slide(slide, data, palette):
    # Main Hero Card
    hero = add_card(slide, SAFE_LEFT + 0.5, SAFE_TOP + 0.2, SAFE_WIDTH - 1.0, SAFE_HEIGHT - 0.4, palette["bg"], palette["primary"])
    
    # Banner
    banner = add_card(slide, SAFE_LEFT + 1.0, SAFE_TOP + 0.5, SAFE_WIDTH - 2.0, 1.2, palette["primary"], None)
    tb_b = add_textbox(slide, SAFE_LEFT + 1.1, SAFE_TOP + 0.65, SAFE_WIDTH - 2.2, 0.8,
                       data["title"].upper(), size_pt=26, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
    
    # Subtitle
    tb_s = add_textbox(slide, SAFE_LEFT + 1.0, SAFE_TOP + 1.9, SAFE_WIDTH - 2.0, 0.6,
                       data["subtitle"], size_pt=18, bold=True, color_hex=palette["accent"], align=PP_ALIGN.CENTER)
    
    # Welcome message card
    msg_card = add_card(slide, SAFE_LEFT + 2.0, SAFE_TOP + 2.7, SAFE_WIDTH - 4.0, 1.8, "FFFFFF", palette["accent"])
    lines = [
        f"Chào mừng các em học sinh lớp {data['lop']}!",
        "Hôm nay cô trò mình cùng khám phá những điều kỳ thú.",
        "Sẵn sàng ngồi ngay ngắn và bước vào bài học nào!"
    ]
    tb_msg = add_textbox(slide, SAFE_LEFT + 2.2, SAFE_TOP + 2.9, SAFE_WIDTH - 4.4, 1.4, "")
    tf = tb_msg.text_frame
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = line
        set_font(r, 16, bold=(i == 0), color_hex=palette["text_on_card"])

def build_warmup_slide(slide, data, palette):
    add_slide_header(slide, data["title"], "1. Khởi động", palette)
    
    top_y = SAFE_TOP + 0.8
    # Left: Tình huống / Câu hỏi dẫn dắt
    left_w = 6.0
    c_left = add_card(slide, SAFE_LEFT, top_y, left_w, SAFE_BOTTOM - top_y, "FFFFFF", palette["primary"])
    
    tb_q_hdr = add_textbox(slide, SAFE_LEFT + 0.2, top_y + 0.15, left_w - 0.4, 0.4,
                           "CÂU HỎI ĐỐ VUI", size_pt=16, bold=True, color_hex=palette["primary"])
    
    tb_q = add_textbox(slide, SAFE_LEFT + 0.2, top_y + 0.6, left_w - 0.4, 3.4, "")
    tf_q = tb_q.text_frame
    for i, text in enumerate(data["questions"]):
        p = tf_q.paragraphs[0] if i == 0 else tf_q.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = text
        set_font(r, 15, bold=(i == 0), color_hex=palette["text_on_card"])
        
    # Right: Gợi ý & Hình minh họa
    right_x = SAFE_LEFT + left_w + 0.3
    right_w = SAFE_WIDTH - left_w - 0.3
    c_right = add_card(slide, right_x, top_y, right_w, SAFE_BOTTOM - top_y, palette["bg"], palette["accent"])
    
    tb_h_hdr = add_textbox(slide, right_x + 0.2, top_y + 0.15, right_w - 0.4, 0.4,
                           "GỢI Ý TỪ THẦY CÔ", size_pt=16, bold=True, color_hex=palette["accent"])
    
    tb_h = add_textbox(slide, right_x + 0.2, top_y + 0.6, right_w - 0.4, 3.4, "")
    tf_h = tb_h.text_frame
    for i, text in enumerate(data["hints"]):
        p = tf_h.paragraphs[0] if i == 0 else tf_h.add_paragraph()
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = text
        set_font(r, 14, color_hex=palette["text_on_bg"])

def build_steps_slide(slide, data, palette):
    add_slide_header(slide, data["title"], "2. Khám phá", palette)
    
    top_y = SAFE_TOP + 0.8
    content_h = SAFE_BOTTOM - top_y
    steps = data["steps"] # List of (step_title, step_desc)
    
    num_steps = len(steps)
    step_w = (SAFE_WIDTH - (num_steps - 1) * 0.25) / num_steps
    
    for idx, (st_title, st_desc) in enumerate(steps):
        cur_x = SAFE_LEFT + idx * (step_w + 0.25)
        # Step card
        c_step = add_card(slide, cur_x, top_y, step_w, content_h, "FFFFFF", palette["primary"])
        
        # Step number header
        c_num = add_card(slide, cur_x + 0.15, top_y + 0.15, step_w - 0.3, 0.6, palette["primary"], None)
        tb_num = add_textbox(slide, cur_x + 0.15, top_y + 0.2, step_w - 0.3, 0.5,
                             f"BƯỚC {idx+1}", size_pt=15, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
        
        # Step title
        tb_t = add_textbox(slide, cur_x + 0.15, top_y + 0.9, step_w - 0.3, 0.8,
                           st_title, size_pt=15, bold=True, color_hex=palette["primary"], align=PP_ALIGN.CENTER)
        
        # Step detail
        tb_d = add_textbox(slide, cur_x + 0.15, top_y + 1.8, step_w - 0.3, content_h - 2.0, "")
        tf_d = tb_d.text_frame
        for l_idx, line in enumerate(st_desc):
            p = tf_d.paragraphs[0] if l_idx == 0 else tf_d.add_paragraph()
            p.space_after = Pt(6)
            r = p.add_run()
            r.text = line
            set_font(r, 13, color_hex=palette["text_on_card"])

def build_visual_slide(slide, data, palette):
    add_slide_header(slide, data["title"], "2. Khám phá", palette)
    
    top_y = SAFE_TOP + 0.8
    content_h = SAFE_BOTTOM - top_y
    
    # Left: Text explanation (40% width)
    left_w = 5.2
    c_left = add_card(slide, SAFE_LEFT, top_y, left_w, content_h, "FFFFFF", palette["primary"])
    
    tb_l_hdr = add_textbox(slide, SAFE_LEFT + 0.2, top_y + 0.15, left_w - 0.4, 0.4,
                           data["left_header"], size_pt=16, bold=True, color_hex=palette["primary"])
    
    tb_l = add_textbox(slide, SAFE_LEFT + 0.2, top_y + 0.6, left_w - 0.4, content_h - 0.8, "")
    tf_l = tb_l.text_frame
    for i, line in enumerate(data["left_points"]):
        p = tf_l.paragraphs[0] if i == 0 else tf_l.add_paragraph()
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = line
        set_font(r, 14, color_hex=palette["text_on_card"])
        
    # Right: Image frame (60% width)
    right_x = SAFE_LEFT + left_w + 0.3
    right_w = SAFE_WIDTH - left_w - 0.3
    c_right = add_card(slide, right_x, top_y, right_w, content_h, palette["bg"], palette["accent"])
    
    img_path = data.get("img_path")
    if img_path and os.path.exists(img_path):
        try:
            # Fit image within the right frame
            img_pad = 0.15
            slide.shapes.add_picture(
                img_path,
                Inches(right_x + img_pad), Inches(top_y + img_pad),
                width=Inches(right_w - img_pad*2)
            )
        except Exception as e:
            print(f"Error adding picture {img_path}: {e}")
    else:
        # Fallback card
        tb_nofile = add_textbox(slide, right_x + 0.5, top_y + 1.5, right_w - 1.0, 1.5,
                                data.get("img_caption", "Hình ảnh minh họa"),
                                size_pt=16, bold=True, color_hex=palette["accent"], align=PP_ALIGN.CENTER)

def build_practice_slide(slide, data, palette):
    add_slide_header(slide, data["title"], "3. Luyện tập", palette)
    
    top_y = SAFE_TOP + 0.8
    content_h = SAFE_BOTTOM - top_y
    
    tasks = data["tasks"]
    task_w = (SAFE_WIDTH - (len(tasks)-1)*0.3) / len(tasks)
    
    for idx, task in enumerate(tasks):
        cur_x = SAFE_LEFT + idx * (task_w + 0.3)
        c_t = add_card(slide, cur_x, top_y, task_w, content_h, "FFFFFF", palette["accent"])
        
        # Task badge
        b_t = add_card(slide, cur_x + 0.15, top_y + 0.15, task_w - 0.3, 0.55, palette["accent"], None)
        tb_tb = add_textbox(slide, cur_x + 0.15, top_y + 0.2, task_w - 0.3, 0.45,
                            f"NHIỆM VỤ {idx+1}", size_pt=14, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
        
        # Task title
        tb_tt = add_textbox(slide, cur_x + 0.15, top_y + 0.85, task_w - 0.3, 0.7,
                            task["name"], size_pt=15, bold=True, color_hex=palette["primary"], align=PP_ALIGN.CENTER)
        
        # Task details
        tb_td = add_textbox(slide, cur_x + 0.15, top_y + 1.6, task_w - 0.3, content_h - 1.8, "")
        tf_td = tb_td.text_frame
        for l_idx, line in enumerate(task["steps"]):
            p = tf_td.paragraphs[0] if l_idx == 0 else tf_td.add_paragraph()
            p.space_after = Pt(6)
            r = p.add_run()
            r.text = line
            set_font(r, 13, color_hex=palette["text_on_card"])

def build_quiz_slide(slide, data, palette):
    add_slide_header(slide, "Thử thách trắc nghiệm nhanh 🎯", "4. Trò chơi", palette)
    
    top_y = SAFE_TOP + 0.8
    # Question card
    q_card = add_card(slide, SAFE_LEFT, top_y, SAFE_WIDTH, 1.1, palette["bg"], palette["primary"])
    tb_q = add_textbox(slide, SAFE_LEFT + 0.3, top_y + 0.15, SAFE_WIDTH - 0.6, 0.8,
                       f"CÂU HỎI: {data['question']}", size_pt=17, bold=True, color_hex=palette["primary"])
    
    # 4 options (2x2 grid)
    opt_y = top_y + 1.25
    opt_w = (SAFE_WIDTH - 0.3) / 2
    opt_h = 1.3
    
    opt_colors = [
        {"bg": "EBF3FE", "border": "2D7DD2", "code": "A"},
        {"bg": "F3EEFF", "border": "7C3AED", "code": "B"},
        {"bg": "ECFDF5", "border": "10B981", "code": "C"},
        {"bg": "FFF7ED", "border": "EA580C", "code": "D"}
    ]
    
    for i in range(4):
        col = i % 2
        row = i // 2
        ox = SAFE_LEFT + col * (opt_w + 0.3)
        oy = opt_y + row * (opt_h + 0.2)
        c_opt = add_card(slide, ox, oy, opt_w, opt_h, opt_colors[i]["bg"], opt_colors[i]["border"])
        
        # Badge A, B, C, D
        b_c = add_card(slide, ox + 0.15, oy + 0.15, 0.6, 0.6, opt_colors[i]["border"], None, shape_type=MSO_SHAPE.OVAL)
        tb_bc = add_textbox(slide, ox + 0.15, oy + 0.22, 0.6, 0.5, opt_colors[i]["code"],
                            size_pt=15, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
        
        # Option text
        opt_text = data["options"][i]
        tb_ot = add_textbox(slide, ox + 0.85, oy + 0.15, opt_w - 0.95, opt_h - 0.3,
                            opt_text, size_pt=14, bold=(i == data["correct_idx"]),
                            color_hex="134E4A" if i == data["correct_idx"] else "333333")

def build_summary_slide(slide, data, palette):
    add_slide_header(slide, "Chốt kiến thức bài học 📝", "Tổng kết", palette)
    
    top_y = SAFE_TOP + 0.8
    content_h = SAFE_BOTTOM - top_y
    
    # HỘP KẾT LUẬN CAM (Section VII.10 Layout 14)
    # Bright orange header box
    orange_box = add_card(slide, SAFE_LEFT + 0.5, top_y, SAFE_WIDTH - 1.0, 0.7, "EA580C", None)
    tb_ob = add_textbox(slide, SAFE_LEFT + 0.5, top_y + 0.12, SAFE_WIDTH - 1.0, 0.5,
                        "EM HÃY GHI NHỚ!", size_pt=18, bold=True, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
    
    # Body card
    body_card = add_card(slide, SAFE_LEFT + 0.5, top_y + 0.7, SAFE_WIDTH - 1.0, content_h - 0.7, "FFF7ED", "EA580C")
    
    tb_sum = add_textbox(slide, SAFE_LEFT + 0.8, top_y + 0.9, SAFE_WIDTH - 1.6, content_h - 1.1, "")
    tf_sum = tb_sum.text_frame
    for i, line in enumerate(data["summary_points"]):
        p = tf_sum.paragraphs[0] if i == 0 else tf_sum.add_paragraph()
        p.space_after = Pt(12)
        r = p.add_run()
        r.text = f"● {line}"
        set_font(r, 16, bold=True, color_hex="7C2D12")

def build_thanks_slide(slide, data, palette):
    card = add_card(slide, SAFE_LEFT + 1.0, SAFE_TOP + 0.4, SAFE_WIDTH - 2.0, SAFE_HEIGHT - 0.8, palette["bg"], palette["primary"])
    
    tb_h = add_textbox(slide, SAFE_LEFT + 1.5, SAFE_TOP + 0.8, SAFE_WIDTH - 3.0, 1.0,
                       "BÀI HỌC KẾT THÚC!", size_pt=28, bold=True, color_hex=palette["primary"], align=PP_ALIGN.CENTER)
    
    tb_c = add_textbox(slide, SAFE_LEFT + 1.5, SAFE_TOP + 1.9, SAFE_WIDTH - 3.0, 0.6,
                       "Chúc các em học sinh UNIGO luôn chăm ngoan, học giỏi! 🌟",
                       size_pt=17, bold=True, color_hex=palette["accent"], align=PP_ALIGN.CENTER)
    
    # BTVN / Dặn dò box
    d_card = add_card(slide, SAFE_LEFT + 2.0, SAFE_TOP + 2.7, SAFE_WIDTH - 4.0, 1.6, "FFFFFF", palette["accent"])
    tb_d_h = add_textbox(slide, SAFE_LEFT + 2.2, SAFE_TOP + 2.85, SAFE_WIDTH - 4.4, 0.4,
                         "DẶN DÒ VỀ NHÀ:", size_pt=15, bold=True, color_hex=palette["primary"])
    
    tb_d = add_textbox(slide, SAFE_LEFT + 2.2, SAFE_TOP + 3.25, SAFE_WIDTH - 4.4, 0.9, "")
    tf_d = tb_d.text_frame
    for l_idx, line in enumerate(data["hw"]):
        p = tf_d.paragraphs[0] if l_idx == 0 else tf_d.add_paragraph()
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = f"- {line}"
        set_font(r, 14, color_hex=palette["text_on_card"])

# ============================================================
# MASTER GENERATOR FOR 5 PRESENTATIONS
# ============================================================

SLIDE_DECKS = [
    # 1. TIỀN TIỂU HỌC
    {
        "folder": "Tiền_tiểu_học",
        "filename": "Slide_Tin_hoc_Tiền_tiểu_học_Tiet10_On_tap_Be_lam_chu_chuot.pptx",
        "palette": {
            "primary": "1B4F9B", "accent": "2D7DD2", "bg": "EBF3FE", "card": "FFFFFF",
            "text_on_bg": "1A2744", "text_on_card": "1A2744"
        },
        "slides": [
            {
                "type": "cover",
                "title": "BÉ LÀM CHỦ CHUỘT MÁY TÍNH",
                "subtitle": "Tin học • Tiền Tiểu học • Tiết 10 (PPCT)",
                "lop": "TT3"
            },
            {
                "type": "warmup",
                "title": "Bé hãy nhìn chú chuột máy tính nào! 🐭",
                "questions": [
                    "Đố bé: Chú chuột máy tính có mấy cái tai?",
                    "Ngón trỏ của bé đặt lên nút nào?",
                    "Ngón giữa của bé đặt lên nút nào?"
                ],
                "hints": [
                    "Chuột có 2 tai: Nút trái và Nút phải.",
                    "Ở giữa còn có một chiếc bánh lăn tròn xoe.",
                    "Tay phải bé cầm chuột thật ngoan, thả lỏng cổ tay nhé!"
                ]
            },
            {
                "type": "steps",
                "title": "4 Tuyệt chiêu của chú chuột thông minh",
                "steps": [
                    ("Di chuyển chuột", ["Cầm chuột nhẹ nhàng.", "Đẩy chuột trên mặt bàn.", "Mũi tên trên màn hình sẽ bay theo tay bé!"]),
                    ("Nháy chuột trái", ["Dùng ngón trỏ.", "Bấm nút trái 1 lần.", "Nghe tiếng 'tách' thật vui để chọn hình!"]),
                    ("Nháy đúp chuột", ["Dùng ngón trỏ.", "Bấm nút trái 2 lần thật nhanh.", "Mở trò chơi bé thích!"]),
                    ("Kéo thả chuột", ["Bấm giữ nút trái.", "Kéo hình đến nơi mới.", "Thả ngón tay ra để ghép hình!"])
                ]
            },
            {
                "type": "practice",
                "title": "Cùng bé luyện tập đôi tay khéo léo 🎮",
                "tasks": [
                    {"name": "Trò chơi Hái Quả", "steps": ["Di chuyển mũi tên đến quả táo.", "Nháy nút trái để hái quả bỏ vào giỏ.", "Hái đủ 5 quả táo chín nhé!"]},
                    {"name": "Trò chơi Ghép Hình", "steps": ["Nháy giữ mảnh ghép chú gấu.", "Kéo chú gấu về đúng khung hình.", "Thả tay ra để hoàn thành bức tranh!"]}
                ]
            },
            {
                "type": "quiz",
                "question": "Khi muốn mở một trò chơi trên màn hình, bé thực hiện thao tác nào?",
                "options": [
                    "A. Nhấc bổng chuột lên không trung",
                    "B. Nháy đúp chuột (bấm 2 lần nhanh)",
                    "C. Lắc mạnh chuột thật nhanh",
                    "D. Rút dây chuột ra ngoài"
                ],
                "correct_idx": 1
            },
            {
                "type": "summary",
                "summary_points": [
                    "Chuột máy tính có nút trái, nút phải và bánh lăn.",
                    "Bé nhớ cầm chuột bằng tay phải, thả lỏng cổ tay thoải mái.",
                    "4 thao tác làm chủ: Di chuyển, Nháy đơn, Nháy đúp và Kéo thả.",
                    "Học xong đặt chuột ngay ngắn, không giật dây chuột nhé!"
                ]
            },
            {
                "type": "thanks",
                "hw": [
                    "Bé hãy kể cho bố mẹ nghe về 4 thao tác của chú chuột máy tính.",
                    "Cùng bố mẹ chơi trò chơi rèn luyện ngón tay khéo léo."
                ]
            }
        ]
    },

    # 2. LỚP 1
    {
        "folder": "Lớp_1",
        "filename": "Slide_Tin_hoc_Lớp_1_Tiet10_On_tap_ky_nang_so.pptx",
        "palette": {
            "primary": "5B21B6", "accent": "7C3AED", "bg": "F3EEFF", "card": "FFFFFF",
            "text_on_bg": "2E1065", "text_on_card": "2E1065"
        },
        "slides": [
            {
                "type": "cover",
                "title": "ÔN TẬP KỸ NĂNG SỐ",
                "subtitle": "Tin học • Lớp 1 • Tiết 10 (PPCT)",
                "lop": "1A1"
            },
            {
                "type": "warmup",
                "title": "Em đã là một 'Chiến sĩ nhí công nghệ' chưa? 💻",
                "questions": [
                    "Em đã biết bật máy tính đúng cách chưa?",
                    "Làm thế nào để mở phần mềm vẽ Paint?",
                    "Vẽ xong một bức tranh đẹp, em phải làm gì để không bị mất tranh?"
                ],
                "hints": [
                    "Bật nút nguồn trên thân máy và màn hình.",
                    "Nháy đúp vào biểu tượng bảng màu Paint.",
                    "Nhớ lưu tệp bằng biểu tượng đĩa mềm (Save) màu tím!"
                ]
            },
            {
                "type": "steps",
                "title": "Quy trình 4 bước tạo sản phẩm số đầu đời",
                "steps": [
                    ("Khởi động Paint", ["Tìm biểu tượng Paint trên Desktop.", "Nháy đúp chuột trái để mở cửa sổ vẽ."]),
                    ("Vẽ nét & Tô màu", ["Chọn hình Shapes (hình tròn, vuông).", "Chọn thùng sơn Fill và chọn màu yêu thích."]),
                    ("Gõ tên của em", ["Chọn biểu tượng chữ A trên thanh công cụ.", "Gõ tên và lớp của mình (Ví dụ: AN 1A1)."]),
                    ("Lưu tệp bài vẽ", ["Nháy nút Save (biểu tượng đĩa mềm góc trên).", "Gõ tên tệp rồi nhấn Enter để lưu an toàn."])
                ]
            },
            {
                "type": "practice",
                "title": "Thử thách: Họa sĩ nhí 1A1 🎨",
                "tasks": [
                    {"name": "Nhiệm vụ 1: Vẽ ngôi nhà", "steps": ["Vẽ mái nhà hình tam giác.", "Vẽ thân nhà hình chữ nhật.", "Tô màu rực rỡ cho ngôi nhà ước mơ."]},
                    {"name": "Nhiệm vụ 2: Ký tên & Lưu bài", "steps": ["Dùng công cụ Text gõ tên mình.", "Nhấn nút Save lưu tệp với tên: Ten_1A1.", "Đóng cửa sổ Paint bằng nút dấu X đỏ."]}
                ]
            },
            {
                "type": "quiz",
                "question": "Biểu tượng nào dưới đây dùng để LƯU sản phẩm bài vẽ của em?",
                "options": [
                    "A. Nút dấu X màu đỏ ở góc trên",
                    "B. Biểu tượng đĩa mềm (Save)",
                    "C. Nút dấu trừ (-) thu nhỏ",
                    "D. Biểu tượng thùng rác"
                ],
                "correct_idx": 1
            },
            {
                "type": "summary",
                "summary_points": [
                    "Máy tính giúp em vẽ tranh, tô màu và gõ chữ rất đẹp.",
                    "Quy trình tạo tranh: Mở Paint -> Vẽ hình -> Tô màu -> Gõ tên -> Lưu bài.",
                    "Luôn nhấn nút Save để lưu giữ sản phẩm trước khi tắt máy.",
                    "Tắt máy tính an toàn bằng lệnh Shut down, không rút dây điện."
                ]
            },
            {
                "type": "thanks",
                "hw": [
                    "Về nhà em hãy chỉ cho bố mẹ nút Save trên phần mềm máy tính.",
                    "Khoe với ông bà bức tranh số đầu tiên em đã tự vẽ ở lớp."
                ]
            }
        ]
    },

    # 3. LỚP 2
    {
        "folder": "Lớp_2",
        "filename": "Slide_Tin_hoc_Lớp_2_Tiet10_On_tap_tep_thu_muc.pptx",
        "palette": {
            "primary": "0F766E", "accent": "14B8A6", "bg": "ECFDF5", "card": "FFFFFF",
            "text_on_bg": "134E4A", "text_on_card": "134E4A"
        },
        "slides": [
            {
                "type": "cover",
                "title": "ÔN TẬP TỆP VÀ THƯ MỤC",
                "subtitle": "Tin học • Lớp 2 • Tiết 10 (PPCT)",
                "lop": "2A1"
            },
            {
                "type": "warmup",
                "title": "Chiếc cặp sách và Thư mục máy tính 🎒📁",
                "questions": [
                    "Vì sao trong cặp sách của em lại có nhiều ngăn?",
                    "Trong máy tính, ngăn chứa tài liệu màu vàng được gọi là gì?",
                    "Nếu tất cả bài tập đều để lộn xộn ngoài màn hình thì sẽ ra sao?"
                ],
                "hints": [
                    "Chia ngăn giúp cặp sách gọn gàng và tìm sách vở nhanh hơn.",
                    "Thư mục (Folder) giống như ngăn cặp sách trong máy tính.",
                    "Tổ chức thư mục khoa học giúp em tìm bài tập trong chớp mắt!"
                ]
            },
            {
                "type": "steps",
                "title": "3 Bước vàng tạo và quản lý thư mục cá nhân",
                "steps": [
                    ("Mở ổ đĩa học tập", ["Nháy đúp mở biểu tượng This PC.", "Mở ổ đĩa D: hoặc thư mục Lop_2A1."]),
                    ("Tạo thư mục mới", ["Nháy nút 'New folder' trên thanh công cụ.", "Gõ tên em: Ten_STT (Ví dụ: Minh_12).", "Nhấn phím Enter để hoàn thành."]),
                    ("Lưu tệp vào thư mục", ["Khi vẽ xong trong Paint, nhấn Save.", "Chọn đúng thư mục mang tên mình.", "Gõ tên bức tranh và nhấn Save."])
                ]
            },
            {
                "type": "visual",
                "title": "Phân biệt Tệp (File) và Thư mục (Folder)",
                "left_header": "ĐẶC ĐIỂM CẦN NHỚ",
                "left_points": [
                    "📁 Thư mục (Folder):",
                    "● Biểu tượng màu vàng kẹp giấy.",
                    "● Dùng để chứa các thư mục con và các tệp.",
                    "📄 Tệp (File):",
                    "● Chứa dữ liệu cụ thể (tranh vẽ, bài viết, nhạc).",
                    "● Có tên tệp và phần mở rộng (.png, .docx).",
                    "● Không thể chứa thư mục khác bên trong."
                ],
                "img_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_2\Tuần_11\images\folder_tree_diagram_vn.png",
                "img_caption": "Sơ đồ cây thư mục trong máy tính"
            },
            {
                "type": "practice",
                "title": "Thử thách: Sắp xếp tài liệu gọn gàng 📂",
                "tasks": [
                    {"name": "Nhiệm vụ 1: Tạo thư mục", "steps": ["Mở ổ đĩa D: -> Lop_2A1.", "Tạo thư mục mang tên em.", "Đổi tên thư mục nếu viết chưa chuẩn."]},
                    {"name": "Nhiệm vụ 2: Lưu và kiểm tra", "steps": ["Vẽ một bức tranh trong Paint.", "Lưu vào đúng thư mục của mình.", "Mở lại thư mục kiểm tra xem tệp đã có chưa."]}
                ]
            },
            {
                "type": "quiz",
                "question": "Hành động nào sau đây là ĐÚNG quy tắc khi dùng chung máy tính ở trường?",
                "options": [
                    "A. Xóa thư mục của bạn để máy tính trống chỗ",
                    "B. Đổi tên thư mục của bạn thành tên mình",
                    "C. Chỉ lưu và quản lý dữ liệu trong thư mục riêng của mình",
                    "D. Lưu tất cả bài làm ra màn hình Desktop"
                ],
                "correct_idx": 2
            },
            {
                "type": "summary",
                "summary_points": [
                    "Thư mục (Folder) có màu vàng, dùng để lưu trữ và phân loại dữ liệu.",
                    "Tệp (File) chứa nội dung tranh ảnh, văn bản hoặc âm thanh.",
                    "Quy trình tạo thư mục: New folder -> Gõ tên -> Nhấn Enter.",
                    "Luôn tôn trọng tài sản số: Không sửa, không xóa dữ liệu của người khác."
                ]
            },
            {
                "type": "thanks",
                "hw": [
                    "Em hãy kiểm tra lại xem cặp sách và góc học tập ở nhà đã ngăn nắp chưa.",
                    "Tập thói quen đặt tên thư mục rõ ràng, dễ nhớ trên máy tính."
                ]
            }
        ]
    },

    # 4. LỚP 3
    {
        "folder": "Lớp_3",
        "filename": "Slide_Tin_hoc_Lớp_3_Tiet10_Bai_4_Lam_viec_voi_may_tinh_Tiet_2.pptx",
        "palette": {
            "primary": "C2410C", "accent": "EA580C", "bg": "FFF7ED", "card": "FFFFFF",
            "text_on_bg": "7C2D12", "text_on_card": "431407"
        },
        "slides": [
            {
                "type": "cover",
                "title": "BÀI 4. LÀM VIỆC VỚI MÁY TÍNH (TIẾT 2)",
                "subtitle": "Tin học • Lớp 3 • Tiết 10 (PPCT) • KNTT",
                "lop": "3A1"
            },
            {
                "type": "warmup",
                "title": "Màn hình nền Desktop có những gì? 🖥️",
                "questions": [
                    "Sau khi bật nút nguồn, màn hình đầu tiên xuất hiện gọi là gì?",
                    "Các hình ảnh nhỏ xinh xắn trên màn hình có tên là gì?",
                    "Thanh ngang ở đáy màn hình có chứa nút Start gọi là gì?"
                ],
                "hints": [
                    "Màn hình nền (Desktop) giống như mặt bàn làm việc của em.",
                    "Các hình vẽ nhỏ là Biểu tượng (Icon) của các phần mềm.",
                    "Thanh ngang ở dưới cùng là Thanh công việc (Taskbar)."
                ]
            },
            {
                "type": "visual",
                "title": "Khám phá màn hình nền Desktop và Biểu tượng",
                "left_header": "CÁC THÀNH PHẦN CHÍNH",
                "left_points": [
                    "1. Màn hình nền (Desktop):",
                    "● Toàn bộ khu vực làm việc chính.",
                    "2. Biểu tượng (Icon):",
                    "● Đại diện cho phần mềm hoặc tệp tin.",
                    "● Nháy đúp chuột để khởi động phần mềm.",
                    "3. Nút Start và Taskbar:",
                    "● Nút Start ở góc trái dưới cùng.",
                    "● Taskbar hiển thị các chương trình đang mở."
                ],
                "img_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_3\Tuần_11\images\desktop_icons.png",
                "img_caption": "Màn hình nền Desktop và các biểu tượng phần mềm"
            },
            {
                "type": "visual",
                "title": "Các nút điều khiển cửa sổ phần mềm",
                "left_header": "3 NÚT LỆNH QUAN TRỌNG",
                "left_points": [
                    "Ở góc trên bên phải mỗi cửa sổ phần mềm:",
                    "➖ Nút Thu nhỏ (Minimize):",
                    "● Thu cửa sổ tạm thời xuống thanh Taskbar.",
                    "🗖 Nút Phóng to / Thu lại (Maximize/Restore):",
                    "● Phóng to cửa sổ tràn màn hình hoặc thu lại.",
                    "❌ Nút Đóng (Close):",
                    "● Nháy vào dấu X đỏ để đóng hẳn phần mềm."
                ],
                "img_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_3\Tuần_11\images\window_controls.png",
                "img_caption": "3 nút điều khiển ở góc trên bên phải cửa sổ"
            },
            {
                "type": "steps",
                "title": "Quy trình tắt máy tính an toàn (Shut down)",
                "steps": [
                    ("Đóng các phần mềm", ["Lưu lại tất cả bài tập đang làm.", "Nháy nút X đỏ để đóng hết các cửa sổ."]),
                    ("Nháy nút Start", ["Di chuột xuống góc trái dưới cùng màn hình.", "Nháy nút chuột trái vào biểu tượng Start."]),
                    ("Chọn nút Power", ["Tìm biểu tượng nút nguồn Power.", "Nháy chuột trái vào nút Power."]),
                    ("Chọn lệnh Shut down", ["Nháy chọn dòng chữ 'Shut down'.", "Chờ đèn tín hiệu và màn hình tắt hẳn."])
                ]
            },
            {
                "type": "practice",
                "title": "Thực hành trên máy tính của em 💻",
                "tasks": [
                    {"name": "Nhiệm vụ 1: Điều khiển cửa sổ", "steps": ["Khởi động phần mềm Paint từ Desktop.", "Thử nháy nút Thu nhỏ (-) rồi mở lại từ Taskbar.", "Thử nháy nút Phóng to rồi đóng bằng nút X đỏ."]},
                    {"name": "Nhiệm vụ 2: Tắt máy chuẩn xác", "steps": ["Đóng tất cả các ứng dụng.", "Thực hiện trình tự: Start -> Power -> Shut down.", "Quan sát màn hình tắt hẳn trước khi rời vị trí."]}
                ]
            },
            {
                "type": "quiz",
                "question": "Vì sao chúng ta KHÔNG ĐƯỢC tắt máy tính bằng cách rút phích cắm đột ngột?",
                "options": [
                    "A. Vì sẽ làm máy tính chạy nhanh hơn",
                    "B. Vì có thể gây hỏng phần cứng và mất hết dữ liệu chưa lưu",
                    "C. Vì phích cắm sẽ phát ra tiếng nhạc",
                    "D. Vì máy tính sẽ tự động xóa sạch ổ đĩa"
                ],
                "correct_idx": 1
            },
            {
                "type": "summary",
                "summary_points": [
                    "Màn hình nền (Desktop) chứa các biểu tượng phần mềm (Icon).",
                    "Nháy đúp chuột vào biểu tượng để khởi động phần mềm.",
                    "Nháy vào nút lệnh Close (dấu X đỏ) để đóng cửa sổ phần mềm.",
                    "Luôn tắt máy tính an toàn bằng lệnh Shut down trong bảng chọn Start."
                ]
            },
            {
                "type": "thanks",
                "hw": [
                    "Đọc lại phần Ghi nhớ trang 22 trong SGK Tin học 3.",
                    "Quan sát và hướng dẫn người thân tắt máy tính đúng cách khi ở nhà."
                ]
            }
        ]
    },

    # 5. LỚP 4
    {
        "folder": "Lớp_4",
        "filename": "Slide_Tin_hoc_Lớp_4_Tiet10_Bai_4_Tim_kiem_thong_tin_tren_Internet_Tiet_2.pptx",
        "palette": {
            "primary": "3730A3", "accent": "4F46E5", "bg": "EEF2FF", "card": "FFFFFF",
            "text_on_bg": "1E1B4B", "text_on_card": "1E1B4B"
        },
        "slides": [
            {
                "type": "cover",
                "title": "BÀI 4. TÌM KIẾM THÔNG TIN TRÊN INTERNET (TIẾT 2)",
                "subtitle": "Tin học • Lớp 4 • Tiết 10 (PPCT) • KNTT",
                "lop": "4C1"
            },
            {
                "type": "warmup",
                "title": "Khám phá Văn Miếu - Quốc Tử Giám 🏛️",
                "questions": [
                    "Văn Miếu - Quốc Tử Giám được xây dựng từ năm nào?",
                    "Hình ảnh Khuê Văn Các xuất hiện trên biểu tượng của thành phố nào?",
                    "Em sẽ dùng từ khóa nào để tìm được nhanh nhất câu trả lời trên Google?"
                ],
                "hints": [
                    "Văn Miếu là trường đại học đầu tiên của Việt Nam.",
                    "Khuê Văn Các là biểu tượng thiêng liêng của Thủ đô Hà Nội.",
                    "Từ khóa càng chính xác và ngắn gọn thì tìm kiếm càng nhanh!"
                ]
            },
            {
                "type": "visual",
                "title": "Kỹ thuật chọn từ khóa và Nhập ô tìm kiếm",
                "left_header": "TỪ KHÓA TỐI ƯU",
                "left_points": [
                    "1. Xác định từ khóa (Keyword):",
                    "● Ngắn gọn, nêu đúng trọng tâm bài tìm.",
                    "● Ví dụ: 'Văn Miếu Quốc Tử Giám xây dựng năm nào'",
                    "● Hoặc: 'Hình ảnh Khuê Văn Các'",
                    "2. Nhập từ khóa:",
                    "● Gõ vào ô tìm kiếm trên google.com.",
                    "● Chọn cụm từ gợi ý thả xuống hoặc nhấn Enter."
                ],
                "img_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_4\Tuần_11\images\google_search_box.png",
                "img_caption": "Giao diện máy tìm kiếm Google và ô nhập từ khóa"
            },
            {
                "type": "visual",
                "title": "Nhận diện Siêu liên kết (Hyperlink) & Đọc bài",
                "left_header": "KHAI THÁC KẾT QUẢ",
                "left_points": [
                    "1. Danh sách kết quả tìm kiếm:",
                    "● Tiêu đề trang web màu xanh nước biển.",
                    "● Dòng tóm tắt nội dung ngắn gọn phía dưới.",
                    "2. Siêu liên kết (Hyperlink):",
                    "● Khi rê chuột vào dòng tiêu đề, con trỏ chuột chuyển thành hình bàn tay 👆.",
                    "● Nháy chuột để mở trang web xem bài chi tiết."
                ],
                "img_path": r"D:\UNIGO\KHBD_Tin_học\Lớp_4\Tuần_11\images\search_results_links.png",
                "img_caption": "Danh sách kết quả tìm kiếm với các siêu liên kết"
            },
            {
                "type": "steps",
                "title": "Các bước tìm kiếm và Lưu hình ảnh từ Internet",
                "steps": [
                    ("Chọn thẻ Hình ảnh", ["Nhập từ khóa cần tìm.", "Nháy chuột vào mục 'Hình ảnh' (Images) trên Google."]),
                    ("Chọn ảnh chất lượng", ["Quan sát các bức ảnh hiển thị.", "Nháy chuột vào bức ảnh đẹp và rõ nét nhất."]),
                    ("Mở bảng lệnh lưu", ["Nháy nút chuột phải vào bức ảnh phóng to.", "Chọn dòng lệnh 'Lưu hình ảnh thành...' (Save image as...)."]),
                    ("Lưu vào thư mục", ["Chọn đúng thư mục riêng của em trên ổ đĩa D:.", "Gõ tên ảnh rồi nhấn nút Save."])
                ]
            },
            {
                "type": "practice",
                "title": "Thử thách: Hướng dẫn viên du lịch nhí 🗺️",
                "tasks": [
                    {"name": "Nhiệm vụ 1: Tra cứu lịch sử", "steps": ["Truy cập google.com trên Google Chrome.", "Tìm năm xây dựng Văn Miếu - Quốc Tử Giám (Năm 1070).", "Ghi năm tìm được vào phiếu học tập."]},
                    {"name": "Nhiệm vụ 2: Tải bộ sưu tập ảnh", "steps": ["Chuyển sang mục Hình ảnh trên Google.", "Tìm và tải 1 ảnh Khuê Văn Các + 1 ảnh Vườn bia Tiến sĩ.", "Lưu cả 2 ảnh vào thư mục riêng của em trong ổ D:."]}
                ]
            },
            {
                "type": "quiz",
                "question": "Khi lướt web tìm tài liệu, em thấy xuất hiện cửa sổ lạ báo trúng thưởng điện thoại, em nên làm gì?",
                "options": [
                    "A. Nháy chuột ngay để nhận quà tặng",
                    "B. Điền thông tin cá nhân và số điện thoại của bố mẹ",
                    "C. Đóng ngay cửa sổ đó lại bằng nút X vì có thể là lừa đảo/virus",
                    "D. Rủ các bạn trong lớp cùng nháy vào nhận thưởng"
                ],
                "correct_idx": 2
            },
            {
                "type": "summary",
                "summary_points": [
                    "Máy tìm kiếm (Google) giúp tìm thông tin dạng văn bản và hình ảnh nhanh chóng.",
                    "Từ khóa tìm kiếm cần ngắn gọn, rõ nghĩa và bám sát nội dung cần tìm.",
                    "Nháy chuột vào siêu liên kết (Hyperlink) để mở trang web xem thông tin chi tiết.",
                    "An toàn trên mạng: Chỉ tải tài liệu học tập, không nhấn vào đường link lạ."
                ]
            },
            {
                "type": "thanks",
                "hw": [
                    "Cùng người thân tìm kiếm thông tin về một danh lam thắng cảnh ở quê hương.",
                    "Luyện tập kỹ năng lưu hình ảnh đẹp vào máy tính phục vụ học tập."
                ]
            }
        ]
    }
]

def generate_deck(deck_data):
    prs = Presentation(TEMPLATE_PATH)
    
    # Keep slide master & background footer layout
    # Remove existing demo slides except 1 to retain layouts
    while len(prs.slides) > 1:
        rId = prs.slides._sldIdLst[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[1]
        
    slide_layout = prs.slide_layouts[1] # Content layout with header/footer master
    
    palette = deck_data["palette"]
    slides_info = deck_data["slides"]
    
    for idx, s_info in enumerate(slides_info):
        if idx == 0:
            slide = prs.slides[0]
            # Clear demo placeholder shapes from slide 0
            for sh in list(slide.shapes):
                sp = sh._element
                sp.getparent().remove(sp)
        else:
            slide = prs.slides.add_slide(slide_layout)
            for sh in list(slide.shapes):
                sp = sh._element
                sp.getparent().remove(sp)
                
        s_type = s_info["type"]
        if s_type == "cover":
            build_cover_slide(slide, s_info, palette)
        elif s_type == "warmup":
            build_warmup_slide(slide, s_info, palette)
        elif s_type == "steps":
            build_steps_slide(slide, s_info, palette)
        elif s_type == "visual":
            build_visual_slide(slide, s_info, palette)
        elif s_type == "practice":
            build_practice_slide(slide, s_info, palette)
        elif s_type == "quiz":
            build_quiz_slide(slide, s_info, palette)
        elif s_type == "summary":
            build_summary_slide(slide, s_info, palette)
        elif s_type == "thanks":
            build_thanks_slide(slide, s_info, palette)
            
    out_dir = os.path.join(OUT_BASE_DIR, deck_data["folder"], "Tuần_11")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, deck_data["filename"])
    prs.save(out_path)
    print(f"[OK] Generated Slide Deck ({len(slides_info)} slides): {out_path}")
    return out_path

def main():
    print("=== START GENERATING 5 STUDENT-FACING SLIDE DECKS FOR TUẦN 11 ===")
    for deck in SLIDE_DECKS:
        generate_deck(deck)
    print("=== ALL 5 SLIDE DECKS COMPLETED SUCCESSFULLY! ===")

if __name__ == "__main__":
    main()
