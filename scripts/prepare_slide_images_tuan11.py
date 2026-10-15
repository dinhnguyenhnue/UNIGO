# -*- coding: utf-8 -*-
"""
prepare_slide_images_tuan11.py
Chuẩn bị hình ảnh cho Slide Tuần 11 môn Tin học:
- Lớp 3: Crop ảnh từ bai4_pages (Màn hình Desktop, Biểu tượng, Các nút điều khiển cửa sổ, Bảng chọn Start -> Shut down)
- Lớp 4: Crop ảnh từ bai4_pages (Trang tìm kiếm Google, Ô nhập từ khóa, Danh sách kết quả & siêu liên kết, Chi tiết Hồ Gươm/Văn Miếu)
- Tiền TH, Lớp 1, Lớp 2: Copy hoặc chuẩn bị ảnh minh họa chuột, bàn phím, lưu bài, thư mục
"""

import os, sys
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

def crop_box(img_path, box, out_path):
    if not os.path.exists(img_path):
        print(f"Warning: {img_path} not found")
        return None
    img = Image.open(img_path)
    w, h = img.size
    crop_area = (int(box[0]*w), int(box[1]*h), int(box[2]*w), int(box[3]*h))
    cropped = img.crop(crop_area)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    cropped.save(out_path)
    print(f"Saved: {out_path} ({cropped.size})")
    return out_path

# 1. Lớp 3 Images
img3_dir = r'D:\UNIGO\KHBD_Tin_học\Lớp_3\Tuần_11\images'
os.makedirs(img3_dir, exist_ok=True)
p20_3 = r'D:\UNIGO\SGK\Lớp_3\bai4_pages\p20.png'
p21_3 = r'D:\UNIGO\SGK\Lớp_3\bai4_pages\p21.png'
p22_3 = r'D:\UNIGO\SGK\Lớp_3\bai4_pages\p22.png'

crop_box(p20_3, (0.05, 0.40, 0.95, 0.90), os.path.join(img3_dir, 'desktop_icons.png'))
crop_box(p21_3, (0.05, 0.15, 0.95, 0.55), os.path.join(img3_dir, 'window_controls.png'))
crop_box(p22_3, (0.05, 0.10, 0.95, 0.60), os.path.join(img3_dir, 'shutdown_steps.png'))

# 2. Lớp 4 Images
img4_dir = r'D:\UNIGO\KHBD_Tin_học\Lớp_4\Tuần_11\images'
os.makedirs(img4_dir, exist_ok=True)
p20_4 = r'D:\UNIGO\SGK\Lớp_4\bai4_pages\p20.png'
p21_4 = r'D:\UNIGO\SGK\Lớp_4\bai4_pages\p21.png'
p23_4 = r'D:\UNIGO\SGK\Lớp_4\bai4_pages\p23.png'

crop_box(p20_4, (0.05, 0.35, 0.95, 0.85), os.path.join(img4_dir, 'google_search_box.png'))
crop_box(p21_4, (0.05, 0.15, 0.95, 0.65), os.path.join(img4_dir, 'search_results_links.png'))
crop_box(p23_4, (0.05, 0.30, 0.95, 0.85), os.path.join(img4_dir, 'web_detail_article.png'))

# 3. Tiền TH Images
img_tth_dir = r'D:\UNIGO\KHBD_Tin_học\Tiền_tiểu_học\Tuần_11\images'
os.makedirs(img_tth_dir, exist_ok=True)

# 4. Lớp 1 Images
img1_dir = r'D:\UNIGO\KHBD_Tin_học\Lớp_1\Tuần_11\images'
os.makedirs(img1_dir, exist_ok=True)

# 5. Lớp 2 Images
img2_dir = r'D:\UNIGO\KHBD_Tin_học\Lớp_2\Tuần_11\images'
os.makedirs(img2_dir, exist_ok=True)

print("Images preparation finished!")
