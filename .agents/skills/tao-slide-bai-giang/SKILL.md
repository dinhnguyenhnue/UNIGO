---
name: tao-slide-bai-giang
description: >
  Tạo Slide bài giảng (.pptx) cho các bài học. Sử dụng khi user yêu cầu tạo slide,
  bài trình chiếu, presentation cho bài giảng. Skill này đọc SGK, trích xuất hình ảnh,
  và tạo slide theo template Unigo với chân trang, font tiếng Việt, màu sắc linh hoạt.
---

# Skill Tạo Slide Bài Giảng

## Tổng quan
Tạo slide bài giảng .pptx dựa trên template Unigo có chân trang, với thiết kế chuyên nghiệp,
font tiếng Việt chuẩn, màu sắc thay đổi linh hoạt theo nội dung bài học.

## Quy trình bắt buộc

### Bước 1: Xác định thông tin
- Môn, Lớp, Bài, Chủ đề
- Nội dung chính từ SGK

### Bước 2: Đọc SGK & Trích xuất hình ảnh (PaddleOCR)
- Mở SGK PDF từ `D:\UNIGO\SGK\Lớp_{X}\`
- Dùng `PaddleOCR / PP-Structure` hoặc script `d:\UNIGO\scripts\sgk_ocr.py` để bóc tách thông tin bài học và tự động crop hình ảnh minh họa từ trang sách.
- Trích xuất hình ảnh minh họa để nhúng trực tiếp vào slide.
- Nếu cần sinh thêm hình ảnh minh họa mới, dùng tool `generate_image`.

### Bước 3: Load template
- **Luôn** bắt đầu từ template: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`
- Kích thước: 20×11.2 inches (widescreen)
- Giữ nguyên chân trang UNIGO

### Bước 4: Thiết kế slide

#### Quy tắc font:
- Font chính: **Arial** (hỗ trợ Unicode/tiếng Việt tốt nhất trên PowerPoint)
- **KHÔNG dùng** Times New Roman cho slide (chỉ dùng cho KHBD .docx)
- Font size: Tiêu đề ≥ 28pt, Nội dung ≥ 18pt, Chú thích ≥ 14pt
- Luôn set cả `a:latin` attribute để đảm bảo không lỗi font

#### Quy tắc màu sắc (LINH HOẠT theo bài):
| Chủ đề | Màu chủ đạo | Màu nhấn |
|--------|-------------|----------|
| Thông tin, Internet, Mạng | Xanh dương | Cam |
| Lập trình, Thuật toán | Tím | Xanh lá |
| Đạo đức số, An toàn | Đỏ cam | Vàng |
| Ứng dụng, Phần mềm | Xanh lá | Xanh dương |
| Máy tính, Phần cứng | Xám xanh | Cam |
| Sáng tạo, Mĩ thuật | Tím hồng | Vàng |

#### Quy tắc tương phản:
- Chữ tối trên nền sáng, chữ trắng trên nền tối
- Nền slide: luôn dùng màu rất nhạt (pastel), KHÔNG dùng trắng thuần
- Tiêu đề section: rounded rectangle với nền đậm + chữ trắng
- Hộp kiến thức: nền pastel nhạt + chữ tối

#### Quy tắc khoảng cách:
- Margin slide: tối thiểu 1cm từ mép
- Khoảng cách giữa các phần tử: ≥ 0.5cm
- Chân trang: 1.2cm từ đáy slide

### Bước 5: Cấu trúc slide deck (12-16 slides)

| # | Slide | Nội dung | Thiết kế |
|---|-------|---------|----------|
| 1 | Trang bìa | Tên bài + Chủ đề + Lớp + GV | Nền màu chủ đạo, chữ trắng lớn |
| 2 | Mục tiêu | Yêu cầu cần đạt | Icon ✅ + bullet points |
| 3 | Khởi động | Câu hỏi/Tình huống/Game | Sinh động, có emoji |
| 4-7 | Kiến thức mới | Hoạt động 1, 2, 3... | Hình ảnh + sơ đồ + text ngắn |
| 8-9 | Luyện tập | Bài tập + Câu hỏi | Màu nhấn, hộp gợi ý |
| 10 | Vận dụng | Liên hệ thực tế | Ví dụ thực tế |
| 11 | Tổng kết | Key takeaways + BTVN | Tóm tắt rõ ràng |
| 12 | Cảm ơn | Lời chào kết | Nền chủ đạo + chữ trắng |

### Bước 6: Chân trang (footer)
Mỗi slide đều có:
- Thanh ngang xanh dương (#0070C0) ở đáy
- Text: "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO" (trắng, bold, 14pt)

### Bước 7: Lưu file
```
D:\UNIGO\Slide_bài_giảng\Lớp_{X}\Slide_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.pptx
```

## Thư viện sử dụng
- `python-pptx`: Tạo .pptx
- `PyMuPDF (fitz)`: Đọc SGK PDF, trích xuất hình
- `Pillow`: Xử lý ảnh nếu cần crop

## Cải tiến liên tục
Sau mỗi lần tạo slide, ghi nhận feedback vào:
`D:\UNIGO\.agents\skills\tao-slide-bai-giang\references\improvement_log.md`
