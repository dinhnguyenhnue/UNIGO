# Nhật ký cải tiến - Harness Giáo viên UNIGO

## 2026-07-28 - Nâng cấp thiết kế Slide & Tổ chức lưu trữ bộ tài liệu Tin học 3-8

- **Cập nhật quy chuẩn thiết kế Slide**:
  - Chuyển sang **Modern Card Grid System** với thiết kế giao diện đa tầng: Nền Slate nhạt (`#F8FAFC`), Card trắng bo góc có viền (`#E2E8F0`), Lề nhấn màu sắc bên trái (Left Accent Strips), Badge Pill thông tin môn học, Huy hiệu số thứ tự (`01`, `02`, `03`).
  - Lồng ảnh AI sinh động vào các Card container chuyên nghiệp.
  - Tích hợp XML Slide Transitions và Click Animations xuất hiện từng phần.

- **Cấu trúc lưu trữ từng bài**:
  ```
  D:\UNIGO\KHBD\Lớp_<X>\Bài <Y>\
  ├── KHBD_Tin_hoc_<X>_Bai01_<Ten_bai>.docx
  ├── Slide_Tin_hoc_<X>_Bai01_<Ten_bai>.pptx
  └── images\
      ├── lop<X>_<ten_anh_1>.png
      └── ...
  ```

- **Quy chuẩn KHBD Word**:
  - Font Times New Roman 13pt, căn lề Trái 3cm, Phải 2cm, Trên 2cm, Dưới 2cm.
  - Tích hợp mục **Năng lực số ⭐** bắt buộc.
- **Trạng thái**: Đã áp dụng lại thành công trọn bộ slide cho tất cả 6 khối lớp (3, 4, 5, 6, 7, 8).

## 2026-09-22 - Tích hợp Vòng điều phối Thống kê Cuộc thi (IOE Lần 3)
- **Hành động**:
  - Xử lý dữ liệu thô IOE Lần 3 (34 bản ghi) so khớp với danh sách 128 học sinh toàn trường UNIGO (11 lớp).
  - Nhận diện và xử lý thành công học sinh có nhiều tài khoản thi (Nguyễn Hương Mộc Lan 4C1 lấy bản ghi cao nhất: Vòng 5, 1740 điểm).
  - So khớp chuẩn xác học sinh trùng tên khác lớp (Nguyễn Đức Minh 1A1 đạt Vòng 6, 2060 điểm và Nguyễn Đức Minh 2A1 đạt Vòng 1, 340 điểm).
  - Tự động xuất file thống kê đợt thi chuyên biệt 4 sheets: `Check_các_cuộc_thi/IOE/Lần 3/Thống kê IOE - Lần 3.xlsx`.
  - Cập nhật Master File: `Check_các_cuộc_thi/Thống kê cuộc thi.xlsx` (đủ 3 lần thi IOE).
- **Lưu vào Bộ não (Knowledge / Quy chuẩn hệ thống)**:
  - Bổ sung **Mục IX** vào `D:\UNIGO\.agents\AGENTS.md` thiết lập quy chuẩn vòng điều phối cuộc thi học sinh khép kín.
  - Cập nhật Pipeline của `unigo-giao-vien` và `check-cuoc-thi` để AI Agent tự động điều phối khi có dữ liệu cuộc thi mới.
- **Kết quả IOE Lần 3**: 33/128 học sinh đã thi (25.78%), tăng +7 HS so với Lần 2 (26 HS) và gần gấp đôi Lần 1 (17 HS). Thủ khoa toàn trường: Trịnh Thảo Tiên (Lớp 2C1) - Vòng 7, 2440 điểm.

