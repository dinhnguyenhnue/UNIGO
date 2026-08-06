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

### Bước 2: Đọc SGK & Trích xuất hình ảnh (BẮT BUỘC)

> **QUY TẮC: MỌI slide hoạt động PHẢI có hình ảnh minh họa. TUYỆT ĐỐI KHÔNG để slide trống chỉ có emoji.**

- Mở SGK PDF từ `D:\UNIGO\SGK\Lớp_{X}\`
- Chạy script `d:\UNIGO\scripts\extract_sgk_all_images.py` để trích xuất:
  - **Full page renders** (200 DPI) → `SGK/Lớp_{X}/bai1_images/full_pages/`
  - **Individual embedded images** → `SGK/Lớp_{X}/bai1_images/`
- **Chuỗi fallback ảnh (3 tầng)**:
  1. ✅ **Ảnh riêng lẻ từ SGK** (`bai1_images/*.jpeg`) — ưu tiên cao nhất
  2. ✅ **Full page SGK** (`full_pages/*.png`) — luôn có cho Lớp 3-8
  3. ✅ **AI-generated** (`KHBD_Tin_học/{folder}/Bài_{XX}/images/`) — dùng `generate_image` tool
- **Với các lớp không có SGK** (Tiền TH, Lớp 1, Lớp 2): BẮT BUỘC tạo ảnh AI gồm: `cover.png`, `activity.png`, `practice.png`, `summary.png`
- **Dùng `slide.shapes.add_picture()`** để chèn ảnh thật vào slide, KHÔNG dùng shape rỗng + emoji thay thế

### Bước 3: Load template
- **Luôn** bắt đầu từ template: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`
- Kích thước: 20×11.2 inches (widescreen)
- **Giữ nguyên chân trang UNIGO** trên MỌI slide (thanh xanh "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO").
- **Bảo tồn slide master/layout:** KHÔNG xóa shapes ở vị trí chân trang. Khi thêm slide mới, luôn dùng layout từ template có sẵn chân trang.
- **Kiểm tra sau xuất:** Xác nhận mỗi slide có shape chân trang UNIGO.

#### Quy tắc thiết kế giao diện & Bố cục nâng cao (High-Aesthetic Design System):
1. **Bảo tồn Slide Master & Vùng An Toàn:**
   - GIỮ NGUYÊN Logo UNIGO (`Picture 7` tại L=0.17in, T=0.15in) và Chân trang (`Picture 9` tại L=0.00in, T=6.43in) từ slide master.
   - VÙNG AN TOÀN NỘI DUNG: Chiều dọc từ **Y = 1.15in đến Y = 6.30in** (chiều cao 5.15in). Tuyệt đối không vẽ shape hoặc đè nền màu vượt quá Y = 6.30in hoặc chờm lên Y < 1.15in làm che logo/chân trang.

2. **Quy chuẩn Font chữ & Cỡ chữ:**
   - Tiêu đề slide chính / Giới thiệu: **24pt - 28pt** (Bold).
   - Nội dung thường (Bullet text): **18pt - 20pt**.
   - Ký tự đầu dòng (●): **14pt**, Giãn dòng (Line spacing): **28pt**, Khoảng cách sau đoạn (Space after): **8pt**.
   - Giới hạn nội dung mỗi slide ngắn gọn (tối đa 3 - 4 dòng bullet) để đảm bảo chữ to rõ, thoáng, không dính hoặc chồng chữ.

3. **Cấu trúc Group Card & Accent Bar:**
   - **Thanh Accent Bar khít tuyệt đối:** Thanh accent màu bên mép card BẮT BUỘC phải khít hoàn toàn chiều cao ô trắng (`bar.top = card.top`, `bar.height = card.height`), không tạo margin hở trên/dưới.
   - **Bắt buộc Grouping:** Card ô trắng (`ROUNDED_RECTANGLE`) và thanh accent bar BẮT BUỘC phải được nhóm lại thành một khối duy nhất (`group_shapes` qua `<p:grpSp>`) để giáo viên dễ dàng di chuyển và căn chỉnh trong PowerPoint.

4. **Nội dung Song ngữ:**
   - Có thể mix thuật ngữ tiếng Anh nhưng BẮT BUỘC có phần mở ngoặc `()` giải thích tiếng Việt rõ ràng.

5. **Slide Tổng kết & Màu sắc:**
   - Slide tổng kết dùng nền nhạt `bg` và chèn panel màu chỉ nằm gọn trong Vùng An Toàn (Y 1.15in → 6.30in). Tuyệt đối không gọi `set_slide_bg(primary)` phủ toàn slide làm che mất logo master.
   - Áp dụng hệ thống xoay vòng 8 bộ màu (Color Palette rotation) linh hoạt, hiện đại cho từng bài/lớp.

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
