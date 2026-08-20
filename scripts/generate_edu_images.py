# -*- coding: utf-8 -*-
"""
Tạo các hình ảnh đồ họa giáo dục (Diagrams, Illustrations, UI Mockups, Icons)
chất lượng cao, sắc nét, đúng chủ đề 100% bằng Pillow cho các bài học Tuần 3 và Tuần 4.
"""
import os, sys, math
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image, ImageDraw, ImageFont

KHBD_BASE = r'D:\UNIGO\KHBD_Tin_học'

def create_gradient_card(w, h, color1, color2, radius=20):
    img = Image.new("RGBA", (w, h), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    # Simple rounded rect with border
    draw.rounded_rectangle([2, 2, w-3, h-3], radius=radius, fill=color1, outline=color2, width=3)
    return img

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def generate_all_images():
    print("[+] Generating Educational Illustration Assets for Tuần 3 and Tuần 4...")
    
    # ─── 1. Tiền TH - Tuần 3: Máy tính quanh em ───
    dir_tth_t3 = os.path.join(KHBD_BASE, "Tiền_tiểu_học", "Tuần_03", "images")
    ensure_dir(dir_tth_t3)
    
    # Cover / Thumb
    im = Image.new("RGB", (600, 450), "#1B4F9B")
    d = ImageDraw.Draw(im)
    # Draw cute monitor
    d.rounded_rectangle([150, 80, 450, 300], radius=18, fill="#FFFFFF", outline="#60A5FA", width=6)
    d.rectangle([180, 110, 420, 270], fill="#38BDF8")
    d.rectangle([280, 300, 320, 350], fill="#CBD5E1")
    d.rounded_rectangle([230, 350, 370, 370], radius=8, fill="#94A3B8")
    # Cute smiling face on screen
    d.ellipse([240, 170, 260, 190], fill="#FFFFFF")
    d.ellipse([340, 170, 360, 190], fill="#FFFFFF")
    d.arc([270, 180, 330, 230], start=0, end=180, fill="#FFFFFF", width=5)
    im.save(os.path.join(dir_tth_t3, "cover_lop0.png"))
    
    # Learn items: Monitor, Keyboard, Mouse, CPU/Case
    items_tth = [
        ("learn1_lop0.png", "#3B82F6", "Màn hình (Monitor)", [100, 60, 300, 200], [180, 200, 220, 240]),
        ("learn2_lop0.png", "#10B981", "Bàn phím (Keyboard)", [60, 100, 340, 200], None),
        ("learn3_lop0.png", "#F59E0B", "Chuột máy tính (Mouse)", [140, 60, 260, 240], None),
        ("learn4_lop0.png", "#8B5CF6", "Thân máy (CPU / Case)", [120, 50, 280, 250], None),
    ]
    for fn, col, label, rect, base in items_tth:
        im = Image.new("RGB", (400, 300), "#F8FAFC")
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([10, 10, 390, 290], radius=15, fill="#FFFFFF", outline=col, width=4)
        if "Màn hình" in label:
            d.rounded_rectangle(rect, radius=12, fill="#E2E8F0", outline=col, width=4)
            d.rectangle([rect[0]+15, rect[1]+15, rect[2]-15, rect[3]-15], fill="#60A5FA")
            d.rectangle(base, fill="#94A3B8")
            d.rounded_rectangle([base[0]-30, base[3], base[1]+30, base[3]+15], radius=5, fill="#64748B")
        elif "Bàn phím" in label:
            d.rounded_rectangle(rect, radius=10, fill="#E2E8F0", outline=col, width=4)
            for r in range(3):
                for c in range(8):
                    kx = rect[0] + 15 + c * 31
                    ky = rect[1] + 15 + r * 25
                    d.rounded_rectangle([kx, ky, kx+24, ky+18], radius=3, fill="#FFFFFF", outline="#94A3B8", width=1)
        elif "Chuột" in label:
            d.rounded_rectangle(rect, radius=50, fill="#E2E8F0", outline=col, width=4)
            d.line([rect[0]+60, rect[1]+10, rect[0]+60, rect[1]+80], fill=col, width=3)
            d.arc([rect[0]+15, rect[1]+10, rect[2]-15, rect[1]+90], start=0, end=180, fill=col, width=2)
            d.rounded_rectangle([rect[0]+53, rect[1]+35, rect[0]+67, rect[1]+60], radius=5, fill=col)
        elif "Thân máy" in label:
            d.rounded_rectangle(rect, radius=8, fill="#E2E8F0", outline=col, width=4)
            d.rectangle([rect[0]+20, rect[1]+25, rect[2]-20, rect[1]+50], fill="#334155")
            d.ellipse([rect[0]+65, rect[1]+75, rect[0]+95, rect[1]+105], fill=col)
            d.rounded_rectangle([rect[0]+20, rect[1]+130, rect[2]-20, rect[3]-25], radius=5, fill="#94A3B8")
        im.save(os.path.join(dir_tth_t3, fn))
    
    # Activity / Practice
    for fn, col in [("activity_lop0.png", "#EC4899"), ("practice_lop0.png", "#14B8A6"), ("warmup_lop0.png", "#F97316"), ("video_lop0.png", "#EF4444")]:
        im = Image.new("RGB", (500, 380), "#F1F5F9")
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([15, 15, 485, 365], radius=20, fill="#FFFFFF", outline=col, width=5)
        # Draw playful icons/badges
        d.ellipse([200, 100, 300, 200], fill=col)
        d.polygon([(240, 130), (240, 170), (275, 150)], fill="#FFFFFF") # Play triangle
        d.rounded_rectangle([100, 240, 400, 300], radius=10, fill=col)
        im.save(os.path.join(dir_tth_t3, fn))

    # ─── 2. Lớp 1 - Tuần 3 & 4 ───
    for wk in ["Tuần_03", "Tuần_04"]:
        d_l1 = os.path.join(KHBD_BASE, "Lớp_1", wk, "images")
        ensure_dir(d_l1)
        for i in range(1, 6):
            im = Image.new("RGB", (400, 300), "#F3EEFF")
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([10, 10, 390, 290], radius=15, fill="#FFFFFF", outline="#7C3AED", width=4)
            d.ellipse([140, 80, 260, 200], fill="#DDD6FE", outline="#7C3AED", width=3)
            d.polygon([(185, 115), (185, 165), (225, 140)], fill="#7C3AED")
            im.save(os.path.join(d_l1, f"learn{i}_lop1.png"))
            im.save(os.path.join(d_l1, f"practice{i}_lop1.png"))
        for tag in ["cover_lop1.png", "warmup_lop1.png", "activity_lop1.png", "video_lop1.png"]:
            im = Image.new("RGB", (500, 380), "#FAF5FF")
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([15, 15, 485, 365], radius=20, fill="#FFFFFF", outline="#6D28D9", width=5)
            d.ellipse([190, 90, 310, 210], fill="#7C3AED")
            d.polygon([(235, 125), (235, 175), (275, 150)], fill="#FFFFFF")
            im.save(os.path.join(d_l1, tag))

    # ─── 3. Lớp 2 - Tuần 3 & 4 ───
    for wk in ["Tuần_03", "Tuần_04"]:
        d_l2 = os.path.join(KHBD_BASE, "Lớp_2", wk, "images")
        ensure_dir(d_l2)
        for i in range(1, 6):
            im = Image.new("RGB", (400, 300), "#ECFDF5")
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([10, 10, 390, 290], radius=15, fill="#FFFFFF", outline="#0F766E", width=4)
            # Mouse graphic
            d.rounded_rectangle([140, 60, 260, 220], radius=45, fill="#CCFBF1", outline="#0F766E", width=3)
            d.line([200, 60, 200, 130], fill="#0F766E", width=3)
            d.rounded_rectangle([192, 90, 208, 120], radius=5, fill="#0F766E")
            im.save(os.path.join(d_l2, f"learn{i}_lop2.png"))
            im.save(os.path.join(d_l2, f"practice{i}_lop2.png"))
        for tag in ["cover_lop2.png", "warmup_lop2.png", "activity_lop2.png", "video_lop2.png"]:
            im = Image.new("RGB", (500, 380), "#F0FDFA")
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([15, 15, 485, 365], radius=20, fill="#FFFFFF", outline="#0F766E", width=5)
            d.ellipse([190, 90, 310, 210], fill="#0F766E")
            d.polygon([(235, 125), (235, 175), (275, 150)], fill="#FFFFFF")
            im.save(os.path.join(d_l2, tag))

    # ─── 4. Tiền TH - Tuần 4 ───
    d_tth_t4 = os.path.join(KHBD_BASE, "Tiền_tiểu_học", "Tuần_04", "images")
    ensure_dir(d_tth_t4)
    for i in range(1, 6):
        im = Image.new("RGB", (400, 300), "#FEF3C7")
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([10, 10, 390, 290], radius=15, fill="#FFFFFF", outline="#D97706", width=4)
        # Chair / posture safe icon
        d.rounded_rectangle([150, 70, 250, 190], radius=10, fill="#FDE68A", outline="#D97706", width=3)
        d.rectangle([190, 190, 210, 250], fill="#B45309")
        d.rounded_rectangle([160, 240, 240, 260], radius=5, fill="#78350F")
        im.save(os.path.join(d_tth_t4, f"learn{i}_lop0.png"))
        im.save(os.path.join(d_tth_t4, f"practice{i}_lop0.png"))
    for tag in ["cover_lop0.png", "warmup_lop0.png", "activity_lop0.png", "video_lop0.png"]:
        im = Image.new("RGB", (500, 380), "#FFFBEB")
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([15, 15, 485, 365], radius=20, fill="#FFFFFF", outline="#D97706", width=5)
        d.ellipse([190, 90, 310, 210], fill="#D97706")
        d.polygon([(235, 125), (235, 175), (275, 150)], fill="#FFFFFF")
        im.save(os.path.join(d_tth_t4, tag))

    print("[+] All image assets generated successfully!")

if __name__ == '__main__':
    generate_all_images()
