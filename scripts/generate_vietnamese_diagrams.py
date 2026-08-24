# -*- coding: utf-8 -*-
"""
Tạo các Sơ đồ & Infographic Quy trình Tiếng Việt chuẩn mực, sắc nét cho Slide
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image, ImageDraw, ImageFont

IMG5_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images"
IMG7_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images"
os.makedirs(IMG5_DIR, exist_ok=True)
os.makedirs(IMG7_DIR, exist_ok=True)

def get_font(size, bold=False):
    # Try system fonts
    font_paths = [
        "C:\\Windows\\Fonts\\arialbd.ttf" if bold else "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\segoeuib.ttf" if bold else "C:\\Windows\\Fonts\\segoeui.ttf",
        "C:\\Windows\\Fonts\\tahoma.ttf"
    ]
    for p in font_paths:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()

# ── 1. LỚP 5 BÀI 4: SƠ ĐỒ 4 THAO TÁC QUẢN LÝ THƯ MỤC CÓ MŨI TÊN CHỈ ĐƯỜNG ──
def create_lop5_bai4_operations():
    w, h = 1200, 800
    img = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # Background soft gradient card
    draw.rounded_rectangle([(20, 20), (w-20, h-20)], radius=24, fill=(253, 242, 248), outline=(244, 63, 94), width=3)
    
    title_font = get_font(36, bold=True)
    step_font = get_font(26, bold=True)
    desc_font = get_font(20, bold=False)
    
    draw.text((w//2, 60), "QUY TRÌNH 4 THAO TÁC QUẢN LÝ THƯ MỤC", font=title_font, fill=(159, 18, 57), anchor="mm")
    
    steps = [
        {"num": "1", "title": "TẠO THƯ MỤC MỚI", "desc": "Nháy chuột phải → New → Folder\n(Phím tắt: Ctrl + Shift + N)", "col": (190, 24, 93), "icon": "📁➕"},
        {"num": "2", "title": "ĐỔI TÊN THƯ MỤC", "desc": "Nháy chuột phải → Rename\n(Hoặc nhấn phím F2)", "col": (13, 148, 136), "icon": "✏️"},
        {"num": "3", "title": "SAO CHÉP / DI CHUYỂN", "desc": "Copy (Ctrl+C) / Cut (Ctrl+X)\n→ Mở thư mục đích → Paste (Ctrl+V)", "col": (217, 119, 6), "icon": "📋"},
        {"num": "4", "title": "XÓA THƯ MỤC RÁC", "desc": "Chọn thư mục → Nhấn phím Delete\n(Xóa vĩnh viễn: Shift + Delete)", "col": (225, 29, 72), "icon": "🗑️"}
    ]
    
    # 2x2 Grid with Arrows
    box_w, box_h = 520, 260
    positions = [(60, 120), (620, 120), (60, 440), (620, 440)]
    
    for idx, (bx, by) in enumerate(positions):
        st = steps[idx]
        # Card
        draw.rounded_rectangle([(bx, by), (bx+box_w, by+box_h)], radius=18, fill=(255, 255, 255), outline=st["col"], width=3)
        # Header strip
        draw.rounded_rectangle([(bx, by), (bx+box_w, by+60)], radius=18, fill=st["col"])
        draw.rectangle([(bx, by+30), (bx+box_w, by+60)], fill=st["col"])
        
        # Step Number Badge
        draw.text((bx+30, by+30), f"BƯỚC {st['num']}", font=step_font, fill=(254, 240, 138), anchor="lm")
        draw.text((bx+150, by+30), st["title"], font=step_font, fill=(255, 255, 255), anchor="lm")
        
        # Description
        draw.text((bx+30, by+100), st["desc"], font=desc_font, fill=(15, 23, 42), spacing=12)
        
    out_p = os.path.join(IMG5_DIR, "lop5_bai4_4_operations_flow.png")
    img.save(out_p)
    print(f"Saved: {out_p}")

# ── 2. LỚP 5 BÀI 4: SƠ ĐỒ CÂY THƯ MỤC THỰC HÀNH HOC_TAP TIẾNG VIỆT ──
def create_lop5_bai4_tree_hoctap():
    w, h = 1200, 800
    img = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    draw.rounded_rectangle([(20, 20), (w-20, h-20)], radius=24, fill=(240, 253, 244), outline=(16, 185, 129), width=3)
    
    title_font = get_font(36, bold=True)
    node_font = get_font(24, bold=True)
    sub_font = get_font(20, bold=False)
    
    draw.text((w//2, 60), "SƠ ĐỒ CÂY THƯ MỤC HỌC TẬP TRÊN Ổ ĐĨA D:", font=title_font, fill=(6, 78, 59), anchor="mm")
    
    # Root Node D:
    draw.rounded_rectangle([(60, 360), (280, 440)], radius=16, fill=(15, 118, 110), outline=(13, 148, 136), width=2)
    draw.text((170, 400), "💾 Ổ ĐĨA (D:)\n(Thư mục gốc)", font=node_font, fill=(255, 255, 255), anchor="mm", align="center")
    
    # Line to Parent Folder
    draw.line([(280, 400), (380, 400)], fill=(15, 118, 110), width=5)
    
    # Parent Folder HOC_TAP
    draw.rounded_rectangle([(380, 350), (620, 450)], radius=16, fill=(190, 24, 93), outline=(159, 18, 57), width=3)
    draw.text((500, 400), "📁 HOC_TAP\n(Thư mục mẹ)", font=node_font, fill=(255, 255, 255), anchor="mm", align="center")
    
    # Branching lines to 3 subfolders
    draw.line([(620, 400), (700, 400)], fill=(190, 24, 93), width=5)
    draw.line([(700, 200), (700, 600)], fill=(190, 24, 93), width=5)
    draw.line([(700, 200), (760, 200)], fill=(190, 24, 93), width=5)
    draw.line([(700, 400), (760, 400)], fill=(190, 24, 93), width=5)
    draw.line([(700, 600), (760, 600)], fill=(190, 24, 93), width=5)
    
    # Subfolders & Files
    subs = [
        {"y": 200, "name": "📁 TOAN", "files": "📄 De_kiem_tra.docx\n📊 Bang_diem.xlsx", "col": (3, 105, 161)},
        {"y": 400, "name": "📁 TIENG_VIET", "files": "📝 Bai_tap_lam_van.docx\n📖 Ke_chuyen.docx", "col": (180, 83, 9)},
        {"y": 600, "name": "📁 TIN_HOC", "files": "💻 Bai4_Cay_thu_muc.pptx\n🎨 Tranh_ve.png", "col": (124, 58, 237)}
    ]
    
    for s in subs:
        sy = s["y"]
        # Subfolder Box
        draw.rounded_rectangle([(760, sy-45), (960, sy+45)], radius=14, fill=s["col"])
        draw.text((860, sy), f"{s['name']}\n(Thư mục con)", font=node_font, fill=(255, 255, 255), anchor="mm", align="center")
        
        # Line to Files
        draw.line([(960, sy), (1010, sy)], fill=s["col"], width=4)
        
        # Files Box
        draw.rounded_rectangle([(1010, sy-50), (1160, sy+50)], radius=12, fill=(255, 255, 255), outline=s["col"], width=2)
        draw.text((1085, sy), s["files"], font=sub_font, fill=(15, 23, 42), anchor="mm", align="center")
        
    out_p = os.path.join(IMG5_DIR, "lop5_bai4_cay_thuc_hanh_hoctap.png")
    img.save(out_p)
    print(f"Saved: {out_p}")

# ── 3. LỚP 5 BÀI 3: SƠ ĐỒ 4 BƯỚC TÌM KIẾM THÔNG TIN BẰNG TỪ KHÓA ──
def create_lop5_bai3_search_steps():
    w, h = 1200, 800
    img = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    draw.rounded_rectangle([(20, 20), (w-20, h-20)], radius=24, fill=(253, 242, 248), outline=(219, 39, 119), width=3)
    
    title_font = get_font(34, bold=True)
    step_font = get_font(24, bold=True)
    desc_font = get_font(20, bold=False)
    
    draw.text((w//2, 60), "QUY TRÌNH 4 BƯỚC TÌM KIẾM THÔNG TIN CHÍNH XÁC", font=title_font, fill=(131, 24, 67), anchor="mm")
    
    steps = [
        {"num": "BƯỚC 1", "title": "XÁC ĐỊNH THÔNG TIN", "desc": "• Em cần tìm thông tin gì?\n• Ví dụ: Thời tiết, địa điểm du lịch Đà Lạt", "col": (190, 24, 93)},
        {"num": "BƯỚC 2", "title": "CHỌN TỪ KHÓA CHUẨN", "desc": "• Đặt từ khóa trong ngoặc kép \" \"\n• Ví dụ: \"dự báo thời tiết Đà Lạt\"", "col": (13, 148, 136)},
        {"num": "BƯỚC 3", "title": "CHỌN NGUỒN UY TÍN", "desc": "• Ưu tiên website chính thống (.gov, .edu)\n• Kiểm tra tác giả và ngày cập nhật", "col": (217, 119, 6)},
        {"num": "BƯỚC 4", "title": "TỔNG HỢP & ÁP DỤNG", "desc": "• Ghi chép lại thông tin cần thiết\n• Chuẩn bị hành lý, trang phục phù hợp", "col": (124, 58, 237)}
    ]
    
    col_w = 260
    for idx, st in enumerate(steps):
        cx = 50 + idx * 280
        # Card
        draw.rounded_rectangle([(cx, 130), (cx+col_w, 730)], radius=18, fill=(255, 255, 255), outline=st["col"], width=3)
        draw.rounded_rectangle([(cx, 130), (cx+col_w, 220)], radius=18, fill=st["col"])
        draw.rectangle([(cx, 190), (cx+col_w, 220)], fill=st["col"])
        
        draw.text((cx + col_w//2, 160), st["num"], font=step_font, fill=(254, 240, 138), anchor="mm")
        draw.text((cx + col_w//2, 195), st["title"], font=get_font(18, bold=True), fill=(255, 255, 255), anchor="mm")
        
        draw.text((cx + 15, 260), st["desc"], font=desc_font, fill=(15, 23, 42), spacing=12)
        
        # Arrow connecting next step
        if idx < 3:
            ax = cx + col_w + 5
            draw.text((ax + 10, 420), "➔", font=get_font(36, bold=True), fill=st["col"], anchor="mm")
            
    out_p = os.path.join(IMG5_DIR, "lop5_bai3_quy_trinh_tim_kiem_4buoc.png")
    img.save(out_p)
    print(f"Saved: {out_p}")

# ── 4. LỚP 7 BÀI 3: BẢNG PHÍM TẮT & THAO TÁC TỆP TIẾNG VIỆT ──
def create_lop7_bai3_shortcuts():
    w, h = 1200, 800
    img = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    draw.rounded_rectangle([(20, 20), (w-20, h-20)], radius=24, fill=(240, 249, 255), outline=(14, 165, 233), width=3)
    
    title_font = get_font(34, bold=True)
    key_font = get_font(26, bold=True)
    desc_font = get_font(22, bold=False)
    
    draw.text((w//2, 60), "BẢNG TỔNG HỢP PHÍM TẮT QUẢN LÝ TỆP & THƯ MỤC", font=title_font, fill=(12, 74, 110), anchor="mm")
    
    shortcuts = [
        {"key": "F2", "action": "Đổi tên tệp hoặc thư mục (Rename)", "col": (3, 105, 161)},
        {"key": "Ctrl + Shift + N", "action": "Tạo nhanh thư mục mới (New Folder)", "col": (13, 148, 136)},
        {"key": "Ctrl + C", "action": "Sao chép tệp / thư mục vào bộ nhớ tạm (Copy)", "col": (217, 119, 6)},
        {"key": "Ctrl + X", "action": "Cắt (Di chuyển) tệp / thư mục (Cut)", "col": (190, 24, 93)},
        {"key": "Ctrl + V", "action": "Dán tệp / thư mục vào vị trí mới (Paste)", "col": (16, 185, 129)},
        {"key": "Delete", "action": "Xóa tệp / thư mục vào Thùng rác (Recycle Bin)", "col": (225, 29, 72)}
    ]
    
    cur_y = 130
    for sc in shortcuts:
        # Key button
        draw.rounded_rectangle([(60, cur_y), (360, cur_y + 85)], radius=14, fill=sc["col"])
        draw.text((210, cur_y + 42), sc["key"], font=key_font, fill=(255, 255, 255), anchor="mm")
        
        # Action box
        draw.rounded_rectangle([(380, cur_y), (1140, cur_y + 85)], radius=14, fill=(255, 255, 255), outline=sc["col"], width=2)
        draw.text((410, cur_y + 42), sc["action"], font=desc_font, fill=(15, 23, 42), anchor="lm")
        
        cur_y += 100
        
    out_p = os.path.join(IMG7_DIR, "lop7_bai3_phim_tat_thao_tac.png")
    img.save(out_p)
    print(f"Saved: {out_p}")

# ── 5. LỚP 7 BÀI 4: SƠ ĐỒ 5 NGUYÊN TẮC CÔNG DÂN SỐ THÔNG THÁI ──
def create_lop7_bai4_5k_rules():
    w, h = 1200, 800
    img = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    draw.rounded_rectangle([(20, 20), (w-20, h-20)], radius=24, fill=(236, 254, 255), outline=(6, 182, 212), width=3)
    
    title_font = get_font(34, bold=True)
    rule_font = get_font(24, bold=True)
    desc_font = get_font(20, bold=False)
    
    draw.text((w//2, 60), "5 NGUYÊN TẮC VÀNG AN TOÀN TRÊN MẠNG XÃ HỘI", font=title_font, fill=(22, 78, 99), anchor="mm")
    
    rules = [
        {"num": "1", "title": "BẢO VỆ MẬT KHẨU", "desc": "Không chia sẻ mật khẩu tài khoản và mã OTP cho bất kỳ ai.", "col": (14, 116, 144)},
        {"num": "2", "title": "BẢO VỆ QUYỀN RIÊNG TƯ", "desc": "Không đăng ảnh, địa chỉ nhà, số CCCD của bản thân và người khác.", "col": (13, 148, 136)},
        {"num": "3", "title": "KIỂM CHỨNG THÔNG TIN", "desc": "Không tin và không chia sẻ các tin đồn giật gân, tin giả (Fake News).", "col": (217, 119, 6)},
        {"num": "4", "title": "ỨNG XỬ VĂN MINH", "desc": "Tôn trọng bạn bè, bình luận lịch sự, không công kích hay bắt nạt qua mạng.", "col": (124, 58, 237)},
        {"num": "5", "title": "BÁO CÁO KỊP THỜI", "desc": "Báo ngay cho bố mẹ, thầy cô hoặc quản trị viên khi gặp nguy hiểm / đe dọa.", "col": (225, 29, 72)}
    ]
    
    cur_y = 120
    for r in rules:
        # Number badge circle
        draw.ellipse([(60, cur_y), (140, cur_y + 80)], fill=r["col"])
        draw.text((100, cur_y + 40), r["num"], font=rule_font, fill=(255, 255, 255), anchor="mm")
        
        # Content box
        draw.rounded_rectangle([(160, cur_y), (1140, cur_y + 80)], radius=14, fill=(255, 255, 255), outline=r["col"], width=2)
        draw.text((180, cur_y + 25), r["title"], font=get_font(20, bold=True), fill=r["col"], anchor="lm")
        draw.text((180, cur_y + 55), r["desc"], font=desc_font, fill=(15, 23, 42), anchor="lm")
        
        cur_y += 105
        
    out_p = os.path.join(IMG7_DIR, "lop7_bai4_quy_tac_5k_an_toan_so.png")
    img.save(out_p)
    print(f"Saved: {out_p}")

if __name__ == "__main__":
    create_lop5_bai4_operations()
    create_lop5_bai4_tree_hoctap()
    create_lop5_bai3_search_steps()
    create_lop7_bai3_shortcuts()
    create_lop7_bai4_5k_rules()
    print("=== TẤT CẢ SƠ ĐỒ TIẾNG VIỆT ĐÃ ĐƯỢC TẠO HOÀN HẢO! ===")
