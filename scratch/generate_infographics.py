# -*- coding: utf-8 -*-
import sys, os
from PIL import Image, ImageDraw, ImageFont
sys.stdout.reconfigure(encoding='utf-8')

info5_dir = r'd:\UNIGO\SGK\Lớp_5\bai3_images\infographics'
info7_dir = r'd:\UNIGO\SGK\Lớp_7\bai3_images\infographics'
os.makedirs(info5_dir, exist_ok=True)
os.makedirs(info7_dir, exist_ok=True)

try:
    font_bold = ImageFont.truetype('arialbd.ttf', 24)
    font_reg = ImageFont.truetype('arial.ttf', 18)
    font_title = ImageFont.truetype('arialbd.ttf', 30)
    font_small = ImageFont.truetype('arial.ttf', 15)
except Exception:
    font_bold = font_reg = font_title = font_small = ImageFont.load_default()

# 1. Lớp 5: Sơ đồ 4 bước tìm kiếm thông tin hiệu quả
img = Image.new('RGB', (1000, 500), '#FFFFFF')
draw = ImageDraw.Draw(img)
draw.rounded_rectangle([(10, 10), (990, 490)], radius=20, outline='#BE185D', width=3)
draw.rounded_rectangle([(30, 25), (970, 75)], radius=12, fill='#BE185D')
draw.text((500, 50), 'QUY TRÌNH 4 BƯỚC TÌM KIẾM THÔNG TIN CHUẨN XÁC', fill='#FFFFFF', font=font_title, anchor='mm')

steps = [
    ('BƯỚC 1', 'Xác định mục tiêu', 'Cần tìm thông tin gì?\nĐịa điểm, thời tiết, chi phí...'),
    ('BƯỚC 2', 'Lựa chọn từ khoá', 'Dùng từ khoá ngắn gọn\nĐặt trong ngoặc kép " "'),
    ('BƯỚC 3', 'Chọn nguồn tin cậy', 'Trang web chính thức\nBáo chí, đánh giá uy tín'),
    ('BƯỚC 4', 'Tổng hợp & Áp dụng', 'So sánh, chọn lọc\nĐưa ra quyết định tối ưu')
]
colors = ['#DB2777', '#E11D48', '#D97706', '#059669']
for i, (b_name, b_title, b_desc) in enumerate(steps):
    x = 40 + i * 235
    draw.rounded_rectangle([(x, 110), (x + 215, 460)], radius=15, fill='#FDF2F8', outline=colors[i], width=2)
    draw.rounded_rectangle([(x + 10, 125), (x + 205, 175)], radius=10, fill=colors[i])
    draw.text((x + 107, 150), b_name, fill='#FFFFFF', font=font_bold, anchor='mm')
    draw.text((x + 107, 210), b_title, fill='#0F172A', font=font_bold, anchor='mm')
    draw.line([(x + 25, 240), (x + 190, 240)], fill=colors[i], width=2)
    draw.multiline_text((x + 107, 310), b_desc, fill='#334155', font=font_reg, anchor='mm', align='center', spacing=8)
    if i < 3:
        ax = x + 218
        draw.polygon([(ax + 2, 280), (ax + 14, 285), (ax + 2, 290)], fill='#BE185D')

out_p = os.path.join(info5_dir, 'quy_trinh_4_buoc_tim_kiem.png')
img.save(out_p)
print('Created:', out_p)

# 2. Lớp 7: Sơ đồ các loại tệp và phần mở rộng
img7 = Image.new('RGB', (1000, 500), '#FFFFFF')
draw7 = ImageDraw.Draw(img7)
draw7.rounded_rectangle([(10, 10), (990, 490)], radius=20, outline='#0369A1', width=3)
draw7.rounded_rectangle([(30, 25), (970, 75)], radius=12, fill='#0369A1')
draw7.text((500, 50), 'CẤU TRÚC PHẦN MỞ RỘNG TỆP TRONG MÁY TÍNH', fill='#FFFFFF', font=font_title, anchor='mm')

file_types = [
    ('VĂN BẢN & SỐ', '.docx, .xlsx, .pptx\n.pdf, .txt', 'Mở bằng phần mềm\nWord, Excel, PPT', '#0284C7'),
    ('ĐA PHƯƠNG TIỆN', '.jpg, .png, .gif\n.mp3, .mp4, .avi', 'Hình ảnh, âm thanh\nVideo giải trí, bài học', '#0D9488'),
    ('CHƯƠNG TRÌNH', '.exe, .com, .bat\n.msi, .dll', 'Tệp thực thi hệ thống\nCài đặt phần mềm', '#EA580C'),
    ('CẢNH BÁO AN TOÀN', '⚠️ Cảnh báo Virus', 'Không xoá tệp .exe lạ\nKhông mở tệp đáng ngờ\nBật Windows Defender', '#DC2626')
]
for i, (f_name, f_ext, f_desc, f_col) in enumerate(file_types):
    x = 40 + i * 235
    draw7.rounded_rectangle([(x, 110), (x + 215, 460)], radius=15, fill='#F0F9FF', outline=f_col, width=2)
    draw7.rounded_rectangle([(x + 10, 125), (x + 205, 175)], radius=10, fill=f_col)
    draw7.text((x + 107, 150), f_name, fill='#FFFFFF', font=font_bold if len(f_name) < 14 else font_small, anchor='mm')
    draw7.multiline_text((x + 107, 230), f_ext, fill=f_col, font=font_bold, anchor='mm', align='center', spacing=6)
    draw7.line([(x + 25, 280), (x + 190, 280)], fill=f_col, width=2)
    draw7.multiline_text((x + 107, 350), f_desc, fill='#1E293B', font=font_reg, anchor='mm', align='center', spacing=8)

out_p7 = os.path.join(info7_dir, 'cau_truc_tep_va_phan_mo_rong.png')
img7.save(out_p7)
print('Created:', out_p7)

# 3. Lớp 7: Sơ đồ 5 thao tác cốt lõi trên tệp và thư mục
img_ops = Image.new('RGB', (1000, 500), '#FFFFFF')
draw_ops = ImageDraw.Draw(img_ops)
draw_ops.rounded_rectangle([(10, 10), (990, 490)], radius=20, outline='#0E7490', width=3)
draw_ops.rounded_rectangle([(30, 25), (970, 75)], radius=12, fill='#0E7490')
draw_ops.text((500, 50), '5 THAO TÁC QUẢN LÝ DỮ LIỆU CỐT LÕI (FILE EXPLORER)', fill='#FFFFFF', font=font_title, anchor='mm')

ops = [
    ('TẠO MỚI', 'New Folder', 'Chuột phải > New\n(Ctrl + Shift + N)', '#0284C7'),
    ('ĐỔI TÊN', 'Rename', 'Chuột phải > Rename\n(Phím tắt F2)', '#0D9488'),
    ('SAO CHÉP', 'Copy & Paste', 'Nhân đôi tệp/thư mục\n(Ctrl + C & Ctrl + V)', '#D97706'),
    ('DI CHUYỂN', 'Cut & Paste', 'Chuyển sang chỗ mới\n(Ctrl + X & Ctrl + V)', '#7C3AED'),
    ('XOÁ BỎ', 'Delete', 'Xoá vào Thùng rác\n(Phím Delete)', '#E11D48')
]
for i, (op_name, op_en, op_desc, op_col) in enumerate(ops):
    x = 30 + i * 190
    draw_ops.rounded_rectangle([(x, 110), (x + 175, 460)], radius=15, fill='#ECFEFF', outline=op_col, width=2)
    draw_ops.rounded_rectangle([(x + 8, 125), (x + 167, 175)], radius=10, fill=op_col)
    draw_ops.text((x + 87, 150), op_name, fill='#FFFFFF', font=font_bold, anchor='mm')
    draw_ops.text((x + 87, 210), op_en, fill=op_col, font=font_bold, anchor='mm')
    draw_ops.line([(x + 20, 240), (x + 155, 240)], fill=op_col, width=2)
    draw_ops.multiline_text((x + 87, 320), op_desc, fill='#1E293B', font=font_reg, anchor='mm', align='center', spacing=8)

out_pops = os.path.join(info7_dir, '5_thao_tac_quan_ly_du_lieu.png')
img_ops.save(out_pops)
print('Created:', out_pops)
