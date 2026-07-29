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

#### Quy tắc thiết kế giao diện & Bố cục nâng cao (High-Aesthetic Design System):
1. **Top Header Banner (Thanh tiêu đề hàng đầu):**
   - Mọi slide nội dung đều có thanh Banner màu chủ đạo phía trên (chiều cao ~3.8cm).
   - Có **Badge Pill màu nhấn** hiển thị tên môn & số bài (ví dụ `[ TIN HỌC 3 • BÀI 1 ]`).
   - Tiêu đề slide chữ trắng in đậm (22pt) đặt vừa vặn trong Banner.
2. **Hệ thống Container dạng Thẻ (Modern Card Grid System):**
   - Nội dung không thả trôi tự do mà được đóng gói trong các **Card bo góc** (`ROUNDED_RECTANGLE`) màu trắng (`#FFFFFF`) hoặc pastel nhạt (`#F8FAFC`, `#EEF2FF`).
   - **Vạch màu nhấn bên trái Card (Left Accent Strip):** Mỗi card kiến thức/mục tiêu có thanh vạch màu chuyên biệt ở lề trái (rộng ~0.6cm) tạo điểm nhấn thị giác cực đẹp.
   - **Huy hiệu số thứ tự (Number Badges):** Mục tiêu hay danh sách được đánh số bằng hình tròn/Pill màu nổi bật (`01`, `02`, `03`).
3. **Khung hình ảnh AI (Image Framing):**
   - Hình ảnh AI được lồng bên trong một Card container màu trắng có viền bo tròn (`ROUNDED_RECTANGLE`) giúp bức ảnh chìm nổi hài hòa với slide, không bị cụt hay thô.
4. **Trang bìa Hero (Hero Title Slide):**
   - Thiết kế dạng Hero Card sang trọng chính giữa slide với thanh accent màu da cam lề trái, Badge Pill môn học, tiêu đề chữ tối nổi bật và thẻ thông tin GV/Trường bên phải.
5. **Quy tắc màu sắc (LINH HOẠT theo bài):**
| Chủ đề | Màu chủ đạo | Màu nhấn |
|--------|-------------|----------|
| Thông tin, Internet, Mạng | Xanh dương (`#1E3A8A`) | Cam (`#EA580C`) / Xanh lá (`#10B981`) |
| Lập trình, Thuật toán | Tím (`#581C87`) | Xanh dương (`#2563EB`) / Xanh lá |
| Đạo đức số, An toàn | Đỏ cam (`#DC2626`) | Vàng (`#F59E0B`) |
| Ứng dụng, Phần mềm | Xanh lá (`#065F46`) | Xanh biển (`#0284C7`) / Cam |
| Máy tính, Phần cứng | Xám xanh (`#0F172A`) | Xanh điện (`#0284C7`) / Hổ phách |
| Sáng tạo, Lịch sử | Tím hoàng gia (`#581C87`) | Xanh ngọc (`#0D9488`) |

#### Quy tắc tương phản & khoảng cách:
- Nền slide: luôn dùng màu rất nhạt (pastel/slate `#F8FAFC`), KHÔNG dùng trắng thuần hay tối thui.
- Card: màu trắng tinh (`#FFFFFF`) có viền nhẹ (`#E2E8F0`) + vạch lề màu nổi bật.
- Margin slide: tối thiểu 1cm từ mép; Chân trang: 1.2cm từ đáy slide.

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

### Bước 7: Hiệu ứng chuyển cảnh & Animation
- **Slide transition (Chuyển trang):** Can thiệp XML `p:transition` với các hiệu ứng `fade`, `push`, `wipe`, `cover`, `split`.
- **Animation (Xuất hiện nội dung):** Can thiệp XML `p:timing` với hiệu ứng `appear` hiển thị các ô văn bản và hình ảnh theo lượt click.

### Bước 8: Lưu file
**Vị trí lưu tập trung theo bài học:**
```
D:\UNIGO\KHBD\Lớp_{X}\Bài {Y}\
├── Slide_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.pptx
├── KHBD_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.docx
└── images\
```
Ví dụ: `D:\UNIGO\KHBD\Lớp_3\Bài 1\Slide_Tin_hoc_3_Bai01_Thong_tin_va_quyet_dinh.pptx`

## Thư viện sử dụng
- `python-pptx`: Tạo .pptx
- `PyMuPDF (fitz)`: Đọc SGK PDF, trích xuất hình
- `Pillow`: Xử lý ảnh nếu cần crop

## Cải tiến liên tục
Sau mỗi lần tạo slide, ghi nhận feedback vào:
`D:\UNIGO\.agents\skills\tao-slide-bai-giang\references\improvement_log.md`
