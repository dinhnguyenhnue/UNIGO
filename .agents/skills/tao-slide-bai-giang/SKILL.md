---
name: tao-slide-bai-giang
description: >
  Tạo Slide bài giảng (.pptx) DÀNH CHO HỌC SINH nhìn và làm theo trong lớp.
  Slide KHÔNG phải bản mô tả lại giáo án — mà là công cụ trực quan giúp HS
  quan sát, thực hành, tương tác. Agent tự phân tích bài học và sinh nội dung
  phù hợp lứa tuổi. Giữ nguyên template Unigo (chân trang, đầu trang, vùng an toàn).
---

# Skill Tạo Slide Bài Giảng (Student-Facing)

## Triết lý cốt lõi

> **Slide bài giảng là thứ HỌC SINH nhìn lên màn hình và làm theo.**
> Đây KHÔNG phải bản mô tả lại Kế hoạch bài dạy cho giáo viên.

### Nguyên tắc vàng:
1. **Visual-first**: Ưu tiên hình ảnh, sơ đồ, minh họa trực quan. Chữ trên slide phải ngắn gọn, dễ hiểu.
2. **Ngôn ngữ hướng tới HS**: Dùng "Em hãy...", "Bước 1: ...", "Quan sát hình...", "Bạn nào biết...?" — KHÔNG dùng ngôn ngữ giáo viên như "HS nhận biết được...", "Mục tiêu: ...", "Năng lực cần đạt: ...".
3. **Tương tác**: Mỗi slide phải có yếu tố kích thích HS suy nghĩ hoặc hành động (câu hỏi, bài tập, thử thách).
4. **Phù hợp lứa tuổi**: Tiền TH/Lớp 1-2 dùng ngôn ngữ cực kỳ đơn giản + nhiều hình. Lớp 3-5 tăng dần. Lớp 6-8 có thể dùng thuật ngữ chuyên môn kèm giải thích.

### ❌ SAI (Slide kiểu mô tả giáo án):
```
Tiêu đề: "Mục tiêu bài học"
Nội dung: "- HS nhận biết được các bộ phận của máy tính
           - HS phân biệt được phần cứng và phần mềm
           - Phát triển năng lực số theo TT 02/2025"
```

### ✅ ĐÚNG (Slide cho HS nhìn và làm theo):
```
Tiêu đề: "Máy tính có những bộ phận nào? 🖥️"
Nội dung: [Hình ảnh máy tính với các mũi tên chỉ vào từng bộ phận]
           "Em hãy quan sát và gọi tên từng bộ phận nhé!"
```

---

## Quy trình bắt buộc

### Bước 1: Xác định thông tin bài học
- Môn, Lớp, Bài, Chủ đề
- Đọc file KHBD `.docx` tương ứng trong `KHBD_Tin_học/Lớp_{X}/Tuần_{YY}/` để hiểu mục tiêu và cấu trúc hoạt động
- **QUAN TRỌNG:** KHBD chỉ là nguồn tham khảo mục tiêu — KHÔNG copy nội dung KHBD vào slide

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

### Bước 3: Sinh nội dung slide bằng AI (THAY ĐỔI QUAN TRỌNG)

> **KHÔNG copy/paste từ KHBD vào slide. Agent phải TỰ PHÂN TÍCH bài học và SINH nội dung phù hợp cho học sinh.**

Quy trình sinh nội dung:
1. **Đọc KHBD** để hiểu: mục tiêu bài, các hoạt động, kiến thức trọng tâm
2. **Đọc SGK** để hiểu: nội dung chính xác HS cần học, hình ảnh, ví dụ trong sách
3. **Chuyển đổi** sang ngôn ngữ Student-Facing:
   - Mục tiêu GV → Câu hỏi dẫn dắt cho HS
   - Nội dung lý thuyết → Hình ảnh + chú thích ngắn
   - Hoạt động GV hướng dẫn → Bước thực hành cho HS (có đánh số)
   - Bài tập trong SGK → Slide luyện tập tương tác
4. **Điều chỉnh ngôn ngữ theo lứa tuổi:**
   - **Tiền TH / Lớp 1-2**: "Các con ơi, hãy nhìn xem...", "Con hãy chỉ vào...", dùng icon lớn, chữ to ≥24pt
   - **Lớp 3-5**: "Em hãy quan sát...", "Bước 1: Em mở...", "Em thử đoán xem..."
   - **Lớp 6-8**: "Hãy quan sát sơ đồ...", "Thảo luận nhóm: ...", thuật ngữ kèm giải thích

### Bước 4: Load template
- **Luôn** bắt đầu từ template: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`
- Kích thước: 20×11.2 inches (widescreen)
- **Giữ nguyên chân trang UNIGO** trên MỌI slide (thanh xanh "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO").
- **Bảo tồn slide master/layout:** KHÔNG xóa shapes ở vị trí chân trang. Khi thêm slide mới, luôn dùng layout từ template có sẵn chân trang.
- **Kiểm tra sau xuất:** Xác nhận mỗi slide có shape chân trang UNIGO.

#### Quy tắc thiết kế giao diện & Bố cục nâng cao (High-Aesthetic Design System):

1. **Bảo tồn Slide Master & Vùng An Toàn (VERIFIED từ template):**
   - **Logo UNIGO** = `Picture 7` tại L=0.17in, T=0.15in, W=0.95in, H=0.94in → kết thúc tại **Y=1.09in**
   - **Chân trang UNIGO** = `Picture 9` tại L=0.00in, T=6.43in, W=13.40in, H=1.23in → bắt đầu từ **Y=6.43in**
   - **VÙNG AN TOÀN NỘI DUNG:** Y = **1.15in → 6.35in** (chiều cao 5.20in)
   - **TUYỆT ĐỐI CẤM:**
     - Vẽ shape/rectangle/background có `top < 1.15in` (che logo)
     - Vẽ shape/rectangle/background có `top + height > 6.35in` (che chân trang)
     - Dùng `add_shape(RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)` phủ toàn bộ slide
   - **Kỹ thuật clamp:** Mọi shape phải được clamp: `actual_top = max(top, 1.15)`, `actual_bottom = min(top+height, 6.35)`

2. **KHÔNG thêm footer tự tạo:**
   - Template master ĐÃ CÓ chân trang `Picture 9` (thanh xanh + thông tin trường). KHÔNG thêm shape footer mới.
   - Chân trang tự động hiển thị trên mọi slide nhờ slide master.

3. **Z-Order & Tương phản (QUAN TRỌNG):**
   - **Background shapes phải `send_to_back`:** Khi tạo nền màu, gọi `sp.getparent().insert(0, sp)` để đẩy xuống dưới cùng.
   - **Text luôn ở trên:** Textbox phải được thêm SAU background shape để tự động nằm trên.
   - **Tương phản cao bắt buộc:**
     - Chữ trắng (`FFFFFF`) trên nền tối (primary/accent đậm)
     - Chữ tối (`1A2744`, `2E1065`...) trên nền nhạt (bg/card trắng)
     - KHÔNG dùng chữ nhạt trên nền nhạt hoặc chữ tối trên nền tối
   - **Palette phải có 3 loại text color:** `text_on_primary`, `text_on_bg`, `text_on_card`

4. **Quy chuẩn Font chữ & Cỡ chữ:**
   - Tiêu đề slide chính / Giới thiệu: **24pt - 28pt** (Bold).
   - Nội dung thường (Bullet text): **18pt - 20pt**.
   - Ký tự đầu dòng (●): **14pt**, Giãn dòng (Line spacing): **28pt**, Khoảng cách sau đoạn (Space after): **8pt**.
   - Giới hạn nội dung mỗi slide ngắn gọn (tối đa 3 - 4 dòng bullet) để đảm bảo chữ to rõ, thoáng, không dính hoặc chồng chữ.
   - **Lớp Tiền TH / 1-2**: Tăng cỡ chữ nội dung lên **22pt - 24pt**, giảm lượng chữ xuống **2 dòng/slide**.

5. **Cấu trúc Group Card & Accent Bar:**
   - **Thanh Accent Bar khít tuyệt đối:** Thanh accent màu bên mép card BẮT BUỘC phải khít hoàn toàn chiều cao ô trắng (`bar.top = card.top`, `bar.height = card.height`), không tạo margin hở trên/dưới.
   - **Bắt buộc Grouping:** Card ô trắng (`ROUNDED_RECTANGLE`) và thanh accent bar BẮT BUỘC phải được nhóm lại thành một khối duy nhất (`group_shapes` qua `<p:grpSp>`) để giáo viên dễ dàng di chuyển và căn chỉnh trong PowerPoint.

6. **Nội dung Song ngữ:**
   - Có thể mix thuật ngữ tiếng Anh nhưng BẮT BUỘC có phần mở ngoặc `()` giải thích tiếng Việt rõ ràng.
   - Ví dụ: "Hardware (Phần cứng)", "Input Device (Thiết bị vào)"

7. **Slide Tổng kết & Màu sắc:**
   - Slide tổng kết dùng nền nhạt `bg` và chèn panel màu chỉ nằm gọn trong Vùng An Toàn (Y 1.15in → 6.35in). Tuyệt đối không gọi `set_slide_bg(primary)` phủ toàn slide làm che mất logo master.
   - Áp dụng hệ thống xoay vòng 8+ bộ màu (Color Palette rotation) linh hoạt, hiện đại cho từng bài/lớp.

### Bước 5: Cấu trúc slide deck (10-14 slides) — STUDENT-FACING

| # | Slide | Mục đích (cho HS) | Nội dung mẫu | Thiết kế |
|---|-------|--------------------|--------------|----------|
| 1 | **Trang bìa** | Gây hứng thú, giới thiệu chủ đề | Hình ảnh lớn bắt mắt + Tên bài ngắn gọn | Nền màu chủ đạo, chữ trắng lớn |
| 2 | **Khởi động** | Kích thích tò mò, dẫn dắt vào bài | "Em hãy đoán xem đây là gì?" + hình ảnh bí ẩn / câu hỏi vui | Hình ảnh to + câu hỏi nổi bật |
| 3-6 | **Nội dung bài học** | Hướng dẫn HS quan sát, khám phá, thực hành từng bước | Hình ảnh SGK + chỉ dẫn ngắn: "Bước 1: Mở...", "Bước 2: Nhấp vào..." | Hình ảnh chiếm 50-60% slide + text ngắn |
| 7-8 | **Luyện tập** | HS tự làm bài tập | "Em hãy thực hành theo các bước sau...", bài tập có gợi ý | Card bài tập + gợi ý màu nhạt |
| 9 | **Vận dụng / Thử thách** | Áp dụng vào thực tế, mini game | Câu hỏi nối, đúng/sai, tình huống thực tế, "Em hãy chia sẻ..." | Thiết kế game/quiz tương tác |
| 10 | **Tổng kết** | Ghi nhớ điểm chính | Infographic tóm tắt 3-4 điểm chính + icon | Nền nhạt, panel trong Vùng An Toàn |
| 11 | **Cảm ơn** | Kết bài + BTVN | "Các em giỏi lắm! 🌟" + BTVN ngắn gọn | Nền chủ đạo + chữ trắng |

> **LƯU Ý QUAN TRỌNG:**
> - KHÔNG có slide "Mục tiêu bài học" liệt kê YCCD theo kiểu giáo án
> - KHÔNG có slide liệt kê "Năng lực", "Phẩm chất" 
> - Mục tiêu được lồng ghép tự nhiên vào nội dung các slide hoạt động
> - Mỗi slide tối đa 3-4 dòng text ngắn + 1 hình ảnh lớn

### Bước 6: Chân trang (footer) — TỰ ĐỘNG TỪ MASTER
Chân trang được cung cấp bởi slide master (`Picture 9`):
- Thanh xanh dương với text "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO" + thông tin liên hệ, địa chỉ
- **KHÔNG cần thêm shape footer mới.** Chỉ cần đảm bảo KHÔNG có shape nào che lên vùng Y > 6.35in.

### Bước 7: Hiệu ứng chuyển cảnh & Animation
- **Slide transition (Chuyển trang):** Can thiệp XML `p:transition` với các hiệu ứng `fade`, `push`, `wipe`, `cover`, `split`.
- **Animation (Xuất hiện nội dung):** Can thiệp XML `p:timing` với hiệu ứng `appear` hiển thị các ô văn bản và hình ảnh theo lượt click.
- **Đặc biệt cho slide Luyện tập/Thử thách:** Dùng animation reveal để HS suy nghĩ trước khi GV click hiện đáp án.

### Bước 8: Lưu file
**Vị trí lưu tập trung theo bài học:**
```
D:\UNIGO\KHBD_Tin_học\Lớp_{X}\Tuần_{YY}\
├── Slide_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.pptx
├── KHBD_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.docx
└── images\
```

Ví dụ: `D:\UNIGO\KHBD_Tin_học\Lớp_3\Tuần_02\Slide_Tin_hoc_Lop_3_Bai01_Thong_tin_va_quyet_dinh.pptx`

---

## Checklist kiểm tra trước khi giao slide

- [ ] Mỗi slide có hình ảnh minh họa (không chỉ có text/emoji)
- [ ] Ngôn ngữ hướng tới HS (không có "HS nhận biết được...", "Mục tiêu:...")  
- [ ] Chân trang UNIGO có trên mọi slide
- [ ] Logo UNIGO không bị che
- [ ] Nội dung nằm trong Vùng An Toàn (Y 1.15in → 6.30in)
- [ ] Tối đa 3-4 dòng text/slide
- [ ] Font chữ đúng chuẩn (cỡ phù hợp lứa tuổi)
- [ ] Có hiệu ứng transition và animation
- [ ] Không có slide liệt kê mục tiêu/năng lực kiểu giáo án

## Thư viện sử dụng
- `python-pptx`: Tạo .pptx
- `PyMuPDF (fitz)`: Đọc SGK PDF, trích xuất hình
- `Pillow`: Xử lý ảnh nếu cần crop

## Cải tiến liên tục
Sau mỗi lần tạo slide, ghi nhận feedback vào:
`D:\UNIGO\.agents\skills\tao-slide-bai-giang\references\improvement_log.md`
