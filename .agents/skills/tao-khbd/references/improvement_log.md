# Nhật ký cải tiến KHBD (Kế hoạch bài dạy)

## 2026-07-28 - Bài 1 Tin học (Lớp 3, 4, 5, 6, 7, 8)

- **Sản phẩm**: Bộ KHBD .docx + Slide .pptx + Ảnh AI minh họa cho 6 khối lớp
- **Cấu trúc lưu trữ**: 
  `D:\UNIGO\KHBD\<Ten_Mon>_Lop<X>_Bai<YY>_<Ten_bai>\`
  - `KHBD_<Mon>_<Lop>_Bai<YY>_<Ten_bai>.docx`
  - `Slide_<Mon>_<Lop>_Bai<YY>_<Ten_bai>.pptx`
  - `images/` (chứa các file ảnh AI .png)
- **Quy chuẩn căn lề mới (User quy định)**:
  - Font: **Times New Roman**, cỡ **13pt**
  - Lề: **Trái 3.0cm**, **Phải 2.0cm**, **Trên 2.0cm**, **Dưới 2.0cm**
  - Line spacing: 1.15
- **Quy chuẩn Slide**:
  - Font: **Arial**, tiêu đề ≥ 28pt, nội dung ≥ 18pt
  - Chân trang: Thanh xanh "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO"
  - Hiệu ứng: Chuyển trang (Transition XML) + Xuất hiện nội dung (Animation XML)
  - Hình ảnh: Tạo bằng AI (`generate_image`) lưu trong folder `images/`
- **Kết quả**: Hoàn thành 100% đúng tiến độ và tiêu chuẩn Unigo.

## 2026-08-15 - Chuẩn hóa Ngày soạn/Ngày dạy & Tên lớp (Lớp 5-8)

- **Vấn đề**: KHBD tạo ra bị thừa 2 dòng "Lớp" trống trong Table[0] (THCS); Ngày soạn/dạy để trống; Tên lớp ghi sai (ghi "6" thay vì "6A1").
- **Giải pháp**:
  1. Xóa dòng `Lớp` thừa trong Table[0] Row[1] Cell[1] — chỉ giữ 2 paragraphs: `Ngày soạn:... Ngày dạy:...` và `Lớp: XA1`.
  2. Thêm hàm `compute_dates()` tính tự động: Ngày dạy = theo LBG, Ngày soạn = Thứ 7 tuần trước.
  3. Đổi tên lớp THCS: `6→6A1`, `7→7A1`, `8→8A1`; TH: `5→5C1`.
  4. Sửa toàn bộ file cũ (12 file KHBD Lớp 5-8, Tuần 1-3) bằng script `fix_khbd_dates.py`.
- **File đã sửa**:
  - `generate_khbd_all.py`: Thêm `LOP_SCHEDULE`, `compute_dates()`, sửa `create_khbd_thcs()`, sửa `create_khbd_th()`, cập nhật tất cả call sites.
  - `SKILL.md`: Thêm Bước 1.5 quy tắc Ngày/Lớp.
  - Tạo mới `references/KHBD_NGAY_LOP.md`: File tham chiếu luật ngày/lớp.
- **Kết quả**: 12/12 file KHBD đã sửa thành công. Không có file thừa.
