import os, sys
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image

def crop_box(img_path, box, out_path):
    img = Image.open(img_path)
    w, h = img.size
    # box is given in relative coordinates (left, top, right, bottom) in 0.0 - 1.0
    crop_area = (int(box[0]*w), int(box[1]*h), int(box[2]*w), int(box[3]*h))
    cropped = img.crop(crop_area)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    cropped.save(out_path)
    print(f"Saved crop: {out_path} ({cropped.size})")

# ─── Lớp 5 Crops ───
IMG5_DIR = r'D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images'
os.makedirs(IMG5_DIR, exist_ok=True)

# Page 14: Bài 3 Khởi động 3 bạn đi du lịch (Nha Trang, Đà Lạt, Úc)
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_14.png', (0.08, 0.40, 0.92, 0.70), os.path.join(IMG5_DIR, 'sgk_lop5_bai3_khoi_dong.png'))

# Page 15: Bài 3 Các bước thu thập thông tin chuẩn bị chuyến đi
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_15.png', (0.08, 0.15, 0.92, 0.85), os.path.join(IMG5_DIR, 'sgk_lop5_bai3_cac_buoc_tim_kiem.png'))

# Page 16: Bài 3 Thực hành tìm kiếm trên máy tính / Google search
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_16.png', (0.08, 0.15, 0.92, 0.85), os.path.join(IMG5_DIR, 'sgk_lop5_bai3_thuc_hanh_search.png'))

# Page 18: Bài 4 Khởi động cây thư mục & Sơ đồ phân cấp
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_18.png', (0.08, 0.40, 0.92, 0.90), os.path.join(IMG5_DIR, 'sgk_lop5_bai4_so_do_cay_thu_muc.png'))

# Page 19: Bài 4 Cấu trúc File Explorer cây thư mục trong máy tính
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_19.png', (0.08, 0.15, 0.92, 0.85), os.path.join(IMG5_DIR, 'sgk_lop5_bai4_file_explorer.png'))

# Page 20: Bài 4 Luyện tập tạo cây thư mục
crop_box(r'D:\UNIGO\SGK\Lớp_5\pages_10_30\page_20.png', (0.08, 0.10, 0.92, 0.70), os.path.join(IMG5_DIR, 'sgk_lop5_bai4_luyen_tap.png'))

# ─── Lớp 7 Crops ───
IMG7_DIR = r'D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images'
os.makedirs(IMG7_DIR, exist_ok=True)

# Page 14: Bài 3 Hình 3.1 Thư mục và tệp trong File Explorer (DuLich, DiemDen, BanNgay...)
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_14.png', (0.15, 0.70, 0.85, 0.93), os.path.join(IMG7_DIR, 'sgk_lop7_bai3_hinh31_cay_thu_muc.png'))

# Page 15: Bài 3 Bảng các loại phần mở rộng tệp (.docx, .pptx, .xlsx, .jpg, .mp3, .exe...)
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_15.png', (0.08, 0.20, 0.92, 0.85), os.path.join(IMG7_DIR, 'sgk_lop7_bai3_hinh32_duoi_mo_rong_tep.png'))

# Page 16: Bài 3 Hình 3.3 Sao lưu dữ liệu ra ổ cứng ngoài / USB
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_16.png', (0.08, 0.15, 0.92, 0.60), os.path.join(IMG7_DIR, 'sgk_lop7_bai3_hinh33_sao_luu_du_lieu.png'))

# Page 18: Bài 3 Hình 3.6 Các lệnh thao tác tệp (Cut, Copy, Paste, Delete, Rename)
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_18.png', (0.15, 0.40, 0.85, 0.62), os.path.join(IMG7_DIR, 'sgk_lop7_bai3_hinh36_lenh_thao_tac.png'))

# Page 19: Bài 4 Hình các kênh trao đổi thông tin (Email, MXH, Diễn đàn)
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_19.png', (0.08, 0.45, 0.92, 0.90), os.path.join(IMG7_DIR, 'sgk_lop7_bai4_cac_kenh_trao_doi.png'))

# Page 21: Bài 4 Giao diện mạng xã hội & tạo bài viết
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_21.png', (0.08, 0.10, 0.92, 0.90), os.path.join(IMG7_DIR, 'sgk_lop7_bai4_giao_dien_mxh_dang_bai.png'))

# Page 22: Bài 4 Hình 4.2, 4.3, 4.4 Các chức năng kết bạn, nhắn tin, đăng bài
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_22.png', (0.10, 0.08, 0.90, 0.90), os.path.join(IMG7_DIR, 'sgk_lop7_bai4_hinh42_43_chuc_nang_mxh.png'))

# Page 23: Bài 4 Mặt trái của MXH & Quy tắc an toàn thông tin
crop_box(r'D:\UNIGO\SGK\Lớp_7\pages_10_30\page_23.png', (0.08, 0.10, 0.92, 0.70), os.path.join(IMG7_DIR, 'sgk_lop7_bai4_rui_ro_va_an_toan.png'))
