# Improvement Log — UNIGO Giáo viên Harness

Ghi nhận cải tiến sau mỗi phiên làm việc.

---

## 2026-07-26 - Bài 1 Lớp 3 Tin học (Thông tin và quyết định)
- **Sản phẩm**: KHBD .docx + Slide .pptx (12 slides)
- **Kết quả**: ✅ Thành công
- **Chi tiết**:
  - KHBD: Theo mẫu Unigo 2026-2027, đầy đủ 5 mục (I-V), có Năng lực số
  - Slide: 12 slides, template Unigo có chân trang, màu xanh dương + cam
  - SGK: PDF dạng image, dùng PyMuPDF render pages
- **Lưu ý kỹ thuật**:
  - SGK Tin học 3 KNTT là image-based PDF, không trích xuất text được
  - Cần render pages ra PNG rồi đọc visual
  - Font Arial cho slide, Times New Roman cho KHBD
  - python-pptx cần set cả `a:latin` attribute để tránh lỗi font
- **Feedback**: Đã yêu cầu tích hợp PaddleOCR vào pipeline để bóc tách SGK dạng ảnh scan.
- **Cải tiến áp dụng**: 
  - Đã tích hợp PaddleOCR module vào `d:\UNIGO\scripts\sgk_ocr.py`.
  - Cập nhật cả 3 Skill (`unigo-giao-vien`, `tao-khbd`, `tao-slide-bai-giang`) để tự động kích hoạt OCR khi đọc SGK scan/image.

