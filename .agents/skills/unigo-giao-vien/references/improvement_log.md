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
