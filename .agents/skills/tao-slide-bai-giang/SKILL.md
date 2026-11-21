---
name: tao-slide-bai-giang
description: >
  Tạo Slide bài giảng (.pptx) DÀNH CHO HỌC SINH nhìn và làm theo trong lớp.
  Slide KHÔNG phải bản mô tả lại giáo án — mà là công cụ trực quan giúp HS
  quan sát, thực hành, tương tác. Agent tự phân tích bài học và sinh nội dung
  phù hợp lứa tuổi. Giữ nguyên template Unigo (chân trang, đầu trang, vùng an toàn).
---

# Skill Tạo Slide Bài Giảng (Student-Facing)

> [!IMPORTANT]
> **QUY TẮC BẮT BUỘC VỀ PHẠM VI MÔN HỌC:**
> 1. **CHỈ TẠO SLIDE BÀI GIẢNG CHO MÔN TIN HỌC.**
> 2. **TUYỆT ĐỐI KHÔNG TẠO SLIDE BÀI GIẢNG CHO MÔN ROBOTICS** vì môn Robotics đã có sẵn hệ thống tài liệu / slide riêng.
> 3. Hình ảnh minh họa phải trung thực, lấy từ SGK trích xuất hoặc ảnh giáo dục đúng nội dung bài học.
> 4. Bắt buộc có phần **Chốt kiến thức / Ghi nhớ** bám sát SGK và tiến trình dạy học trong KHBD.

## Triết lý cốt lõi

> **Slide bài giảng là thứ HỌC SINH nhìn lên màn hình và làm theo.**
> Đây KHÔNG phải bản mô tả lại Kế hoạch bài dạy cho giáo viên.

### Nguyên tắc vàng:
1. **Visual-first & Tự do sáng tạo**: Ưu tiên hình ảnh minh họa 3D/Cartoon sống động (dùng `generate_image`), sơ đồ trực quan, bố cục infographic hiện đại (như Canva/Figma). TUYỆT ĐỐI KHÔNG cắt dán thô các trang sách scan gây khô khan, nhàm chán.
2. **Khớp nội dung chính xác từng slide (Semantic Alignment)**:
   - **Slide Quy trình / Các bước thao tác:** Hình ảnh PHẢI là sơ đồ các bước trực quan có mũi tên chỉ đường tuần tự (Bước 1 $\rightarrow$ Bước 2 $\rightarrow$ Bước 3 $\rightarrow$ Bước 4).
   - **Slide Ví dụ thực tế:** Hình ảnh PHẢI nêu đúng đặc điểm của ví dụ đó (ví dụ cây thư mục học tập `HOC_TAP` có `TOAN`, `TIENG_VIET`, `TIN_HOC` thì ảnh phải thể hiện chính xác các tên thư mục này).
   - **TUYỆT ĐỐI KHÔNG dùng 1 ảnh chung chung lặp lại cho nhiều slide khác nhau.** Mỗi slide nội dung phải có hình ảnh tương thích 100% với nội dung của slide đó.
3. **Chữ viết trong ảnh BẮT BUỘC PHẢI LÀ TIẾNG VIỆT**:
   - Mọi văn bản, nhãn, tên thư mục, chú thích bên trong hình ảnh tạo ra PHẢI LÀ TIẾNG VIỆT (ví dụ: `Ổ đĩa C:`, `Thư mục mẹ`, `Thư mục con`, `Tạo thư mục mới`, `Đổi tên`, `Sao chép`, `Xóa`).
   - TUYỆT ĐỐI KHÔNG để chữ tiếng Anh ngẫu nhiên không đúng ngữ cảnh như "KNOWLEDGE TREE", "ESSAYS", "PROJECTS".
4. **Linh hoạt & Đặc trưng từng bài**: Mỗi bài học có phong cách bố cục, màu sắc và hình ảnh tương ứng với chủ đề (không rập khuôn máy móc). Tự do sáng tạo thẻ bài tập, sơ đồ tư duy, bảng đối sánh sinh động dựa trên mục tiêu KHBD.
5. **Ngôn ngữ hướng tới HS**: Dùng "Em hãy...", "Bước 1: ...", "Quan sát hình...", "Bạn nào biết...?" — KHÔNG dùng ngôn ngữ giáo viên như "HS nhận biết được...", "Mục tiêu: ...", "Năng lực cần đạt: ...".
6. **Tương tác**: Mỗi slide phải có yếu tố kích thích HS suy nghĩ hoặc hành động (câu hỏi, bài tập, thử thách, mascot đố vui).
7. **Phù hợp lứa tuổi**: Tiền TH/Lớp 1-2 dùng ngôn ngữ cực kỳ đơn giản + hình cartoon 3D dễ thương. Lớp 3-5 hình minh họa tươi sáng, sinh động. Lớp 6-8 infographic công nghệ hiện đại, chuyên nghiệp.
8. **Chốt kiến thức chuẩn mực**: Mỗi bài học phải có slide Tổng kết / Chốt kiến thức đọng lại các ý cốt lõi từ SGK (Hộp Kết Luận Cam nổi bật).

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

### Bước 2.5: Tạo ảnh riêng cho từng gạch đầu dòng (Per-Bullet Images) — BẮT BUỘC

> **QUY TẮC MỚI: Mỗi bullet point trong slide nội dung PHẢI có ảnh minh họa riêng.**
> Không dùng 1 ảnh chung cho toàn slide. Mỗi ý → 1 ảnh → giúp HS ghi nhớ trực quan.

#### Quy trình tạo ảnh per-bullet:

1. **Tách bullet:** Phân tích nội dung slide, tách ra từng bullet point riêng biệt.
2. **Tạo prompt AI từ nội dung bullet:**
   - Lấy chính text bullet làm prompt gốc
   - Thêm context phù hợp lứa tuổi:
     - **Tiền TH / Lớp 1-2**: `"cute kawaii cartoon illustration for young children aged 5-7, bright vivid colors, simple shapes, friendly characters, no text in image"`
     - **Lớp 3-5**: `"friendly colorful educational cartoon illustration for children aged 8-10, clear details, educational theme, no text in image"`
     - **Lớp 6-8**: `"clean modern educational infographic illustration for middle school students aged 11-14, semi-realistic style, no text in image"`
   - Ví dụ:
     - Bullet `"Màn hình — để con nhìn"` → Prompt: `"Cute kawaii cartoon: a friendly computer monitor screen with colorful display, for young children, bright colors, no text"`
     - Bullet `"Phần cứng = sờ được, nhìn thấy"` → Prompt: `"Educational infographic: hands touching computer hardware components keyboard mouse monitor, for middle school, clean modern style, no text"`
3. **Gọi `generate_image` tool** cho mỗi bullet → lưu tại `images/slide_{N}_bullet_{M}.png`
4. **Chuỗi fallback ảnh per-bullet (3 tầng):**
   - Tầng 1: ✅ Ảnh SGK trích xuất (nếu match nội dung bullet)
   - Tầng 2: ✅ AI-generated từ prompt bullet (dùng `generate_image`)
   - Tầng 3: ✅ Ảnh đại diện chung cho loại nội dung (cover/activity/practice)

#### Layout Per-Bullet Image — Phân cấp theo lứa tuổi:

**Layout A — Grid Flashcard (Tiền TH → Lớp 5):**
```
┌──────────────────────────────────────────────────┐
│ [Badge Banner]                                    │
│ Tiêu đề slide                                    │
│                                                    │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐           │
│  │  ẢNH 1  │  │  ẢNH 2  │  │  ẢNH 3  │           │
│  │(2in×2in)│  │(2in×2in)│  │(2in×2in)│           │
│  ├─────────┤  ├─────────┤  ├─────────┤           │
│  │ Text 1  │  │ Text 2  │  │ Text 3  │           │
│  └─────────┘  └─────────┘  └─────────┘           │
└──────────────────────────────────────────────────┘
```
- 2-3 bullets/slide → 1 hàng × 2-3 cột
- 4 bullets/slide → 2 hàng × 2 cột
- Nếu > 4 bullets → CHIA thành 2 slides
- Ảnh kích thước: **2in × 2in** (đủ lớn HS nhìn rõ)
- Text bên dưới ảnh: **18pt**, căn giữa

#### Hệ thống 20 Dạng Bố cục Slide Chuẩn UNIGO (Visual-First & Hero Image):

> [!IMPORTANT]
> **QUY TẮC BẮT BUỘC VỀ KÍCH THƯỚC ẢNH & BỐ CỤC:**
> 1. **ƯU TIÊN HÌNH ẢNH TO, RÕ RÀNG (HERO IMAGE):** Ảnh minh họa phải chiếm **50% – 70%** diện tích slide hoặc tách slide riêng cho từng ý lớn.
> 2. **TEXT NGẮN GỌN — ĐÚNG Ý CHÍNH SGK:** Text trên slide chỉ nêu ý cốt lõi, từ ngữ thân thiện, dễ hiểu cho học sinh.
> 3. **20 DẠNG BỐ CỤC MẪU CHUẨN:**
>    - **Dạng 1 (Mục lục bài học):** Cửa sổ tìm kiếm / Search bar + Các thẻ mục La Mã (I, II, III, IV) màu sắc nổi bật (Đỏ, Xanh lá, Xanh dương, Vàng cam).
>    - **Dạng 2 (Slide chuyển phần / Section Banner):** Banner nền màu chủ đạo cỡ lớn chiếm trọn trung tâm, tiêu đề phần chữ in hoa màu trắng cực to rõ.
>    - **Dạng 3 (Hoạt động nhóm / Phiếu bài tập):** Khối trái (Câu hỏi dẫn dắt có icon chat) + Khối phải (Bảng phiếu bài tập 2 cột có tiêu đề màu).
>    - **Dạng 4 (Sơ đồ quy trình / Tương quan bước):** Khối trái (Hình ảnh nhân vật/khái niệm lớn chiếm 40%) + Khối phải (Các thẻ bước quy trình bo góc có mũi tên nối tuần tự).
>    - **Dạng 5 (Ví dụ thực tế - Hero Image):** Khối trái (35% Thẻ chữ bo góc mây tóm tắt) + Khối phải (65% Tranh minh họa / truyện tranh các bước to rõ, sinh động).
>    - **Dạng 6 (Phân tích tình huống 2 vế kèm Avatar tròn):** Đề bài trên + 2 Card đối sánh lớn (Vế A ↔ Vế B) kèm **Avatar tròn minh họa kẹp ở góc card**.
>    - **Dạng 7 (Thảo luận nhóm & Thẻ đáp án kèm sticker/vật thể đè mép):** Khối tình huống bên trái + Các thẻ đáp án bên phải có **sticker minh họa 3D đè lên góc mép card**.
>    - **Dạng 8 (Truyện tranh 2 khung tranh lớn nối mũi tên):** 2 Khung tranh minh họa to chiếm 90% diện tích nối bằng mũi tên đỏ lớn `Tranh A ──> Tranh B`.
>    - **Dạng 9 (Slide Kết luận / Chốt kiến thức tranh đôi):** Khối trái (30% Kết luận ngắn gọn + sơ đồ nút) + Khối phải (70% Bộ đôi tranh liên hoàn to rõ).
>    - **Dạng 10 (Trắc nghiệm 4 đáp án A-B-C-D kèm Mascot):** Câu hỏi trên + Lưới 4 thẻ đáp án (2×2) bo góc màu ấm, có Mascot dễ thương ở góc card.
>    - **Dạng 11 (Cửa sổ trình duyệt & Câu hỏi mở đầu):** Khung cửa sổ Browser Window (nút điều khiển, search bar) + Câu hỏi nhóm in nghiêng đậm + Hình ảnh thiết bị trung tâm.
>    - **Dạng 12 (Sơ đồ thiết bị trung tâm & 4 Nhãn Callout chỉ hướng):** Thiết bị lớn ở giữa + 4 Nhãn màu (Xanh lá, Cam, Tím, Xanh dương) có mũi tên chỉ thẳng vào từng bộ phận kèm chức năng.
>    - **Dạng 13 (Đề mục + Badge Hoạt động + Ảnh SGK đánh số 1-2-3-4):** Tiêu đề mục đỏ đậm + Mascot chú ong kèm Badge Hoạt động + Ảnh SGK trích xuất có đánh số kèm các thẻ nhãn 2 bên.
>    - **Dạng 14 (Bóng thoại câu hỏi + Gợi ý + Hộp KẾT LUẬN cam nổi bật):** Bóng thoại câu hỏi trên + Khối icon tròn Gợi ý ở giữa + **Hộp KẾT LUẬN màu cam bo góc** chốt kiến thức trọng tâm ở dưới.
>    - **Dạng 15 (Chia cột Hướng dẫn & Hình SGK máy tính xách tay):** Tiêu đề mục + Badge hoạt động + Cột trái (Đoạn hướng dẫn có từ khóa màu nổi bật) \| Cột phải (Hình SGK xách tay có đánh số).
>    - **Dạng 16 (Hoạt động nhóm quan sát & Bộ 3 tranh a-b-c):** Khung bo viền ribbon + Icon nhóm học tập + Bộ 3 tranh tư thế a, b, c to rõ ở giữa.
>    - **Dạng 17 (Chữa bài với Badge Đúng/Sai ❌ ✅ & Text giải thích màu):** Bộ tranh có gắn dấu tích ❌ ❌ ✅ lớn + Khối text giải thích từng hình có tiêu đề màu nổi bật.
>    - **Dạng 18 (Câu hỏi sức khỏe + Minh họa cột sống + Lời khuyên đỏ gạch):** Icon suy nghĩ hỏi chấm + 4 hình cột sống đối sánh + Khối chữ cảnh báo màu đỏ gạch nổi bật.
>    - **Dạng 19 (Sơ đồ tư thế đúng & 5 Thẻ Arrow Badge màu sắc):** Cột trái (5 Thẻ mũi tên màu sắc nổi bật: Cam, Xanh dương, Xanh lá, Vàng, Hồng) + Cột phải (Hình tư thế đúng lớn kèm đường dóng).
>    - **Dạng 20 (Giấy nhớ Sticky Note vàng & Các thẻ câu hỏi mũi tên):** Cột trái (Giấy ghi chú Sticky Note màu vàng cuộn góc) + Cột phải (Các thẻ câu hỏi có đầu mũi tên màu sắc).

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
   - **Background shapes phải `send_to_back`:** Khi tạo nền màu, gọi `spTree = sp.getparent(); spTree.remove(sp); spTree.insert(2, sp)` để đẩy xuống dưới cùng (sau `nvGrpSpPr` và `grpSpPr`). TUYỆT ĐỐI KHÔNG dùng `insert(0, sp)` vì sẽ làm vỡ XML schema (`spTree`).
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
| 3-6 | **Nội dung bài học** | Hướng dẫn HS quan sát, khám phá | **Mỗi bullet có ảnh riêng** — Layout A (Grid) cho TH, Layout B (Row) cho THCS | Per-Bullet Image Grid/Row |
| 7-8 | **Luyện tập** | HS tự làm bài tập, trả lời câu hỏi | Items xuất hiện **tuần tự theo click** + ảnh minh họa mỗi item | Animation appear per-item + ảnh |
| 9 | **Trò chơi / Thử thách** | Mini game tương tác | Ghép nối, Đúng/Sai, Sắp xếp — **BẮT BUỘC có ảnh minh họa** | Game layout + ảnh ~2in + animation |
| 10 | **Tổng kết** | Ghi nhớ điểm chính | Infographic tóm tắt 3-4 điểm chính + icon | Nền nhạt, panel trong Vùng An Toàn |
| 11 | **Cảm ơn** | Kết bài + BTVN | "Các em giỏi lắm! 🌟" + BTVN ngắn gọn | Nền chủ đạo + chữ trắng |

> **LƯU Ý QUAN TRỌNG:**
> - KHÔNG có slide "Mục tiêu bài học" liệt kê YCCD theo kiểu giáo án
> - KHÔNG có slide liệt kê "Năng lực", "Phẩm chất" 
> - Mục tiêu được lồng ghép tự nhiên vào nội dung các slide hoạt động
> - Mỗi slide tối đa 3-4 dòng text ngắn
> - **MỖI BULLET POINT phải có ẢNH RIÊNG** (per-bullet image)
> - **Slide câu hỏi/ghép nối: animation tuần tự + ảnh từng item**
> - **Slide trò chơi: BẮT BUỘC có ảnh minh họa đủ lớn (~2in) cho HS nhìn**

### Bước 6: Chân trang (footer) — TỰ ĐỘNG TỪ MASTER
Chân trang được cung cấp bởi slide master (`Picture 9`):
- Thanh xanh dương với text "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO" + thông tin liên hệ, địa chỉ
- **KHÔNG cần thêm shape footer mới.** Chỉ cần đảm bảo KHÔNG có shape nào che lên vùng Y > 6.35in.

### Bước 7: Hiệu ứng chuyển cảnh & Animation nâng cao

#### 7.1 Slide transition (Chuyển trang):
- Can thiệp XML `p:transition` với các hiệu ứng `fade`, `push`, `wipe`, `cover`, `split`.

#### 7.2 Animation tuần tự cho slide Câu hỏi / Ghép nối (BẮT BUỘC):

> **Slide practice/activity có `items` PHẢI animation từng item theo click.**
> HS cần thời gian suy nghĩ trước khi GV bấm hiện item/đáp án tiếp theo.

**Quy tắc animation per-item:**
1. Mỗi item (card text + ảnh minh họa) là 1 nhóm animation.
2. Thứ tự: Item 1 → click → Item 2 → click → Item 3...
3. Đáp án (nếu có) xuất hiện SAU câu hỏi, bằng 1 click riêng.
4. Ảnh minh họa của item xuất hiện CÙNG LÚC với text item (cùng 1 click).

**Quy tắc animation cho slide Ghép nối (items chứa `↔` hoặc `→`):**
1. Cột trái (vế A) + ảnh vế A xuất hiện trước → click → Cột phải (vế B) xuất hiện.
2. Mỗi cặp ghép nối có ảnh minh họa cho vế A (tạo bằng AI), kích thước ~2in×2in.
3. Đường nối/mũi tên giữa 2 vế xuất hiện sau khi cả 2 vế đã hiện.

**Hàm helper bắt buộc trong script:**
- `add_appear_animation(slide, shape, click_index)` — gán XML `p:timing` animation `appear` cho shape theo thứ tự click.
- `add_group_animation(slide, shapes_list, click_index)` — gán animation appear cho nhóm shapes (text + ảnh) cùng 1 click.

**XML Animation template (per-shape appear on click):**
```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <!-- Repeat for each click_index: -->
                <p:par>
                  <p:cTn id="N" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="N+1" presetID="1" presetClass="entr" presetSubtype="0" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="N+2" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="SHAPE_ID"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>
```

#### 7.3 Slide Trò chơi — Bắt buộc có hình ảnh minh họa:

> **MỌI slide trò chơi / hoạt động (`activity`) PHẢI có ảnh minh họa đủ lớn (~2in×2in) để HS nhìn rõ khi chơi.**

**Bảng loại trò chơi và ảnh cần tạo:**

| Loại trò chơi | Ảnh cần tạo | Prompt AI mẫu |
|---|---|---|
| **Đúng/Sai** | Ảnh minh họa từng câu hỏi + biểu tượng ✅❌ | `"[Nội dung câu hỏi], educational cartoon, clear yes/no visual"` |
| **Ghép nối** | Ảnh từng đối tượng vế A (VD: ảnh màn hình, bàn phím...) | `"[Đối tượng cần ghép], isolated on white background, cartoon style, clear"` |
| **Sắp xếp thứ tự** | Ảnh timeline / dòng thời gian minh họa | `"Timeline infographic showing [nội dung], colorful, educational"` |
| **Vẽ/Sáng tạo** | Ảnh mẫu hoặc ảnh HS đang thực hiện | `"[Hoạt động], children in classroom, bright illustration"` |
| **Thảo luận nhóm** | Ảnh nhóm HS đang thảo luận + chủ đề | `"Group of students discussing [topic] in classroom, cartoon"` |
| **Chỉ/Nhận diện** | Ảnh đối tượng cần nhận diện với mũi tên/label | `"[Đối tượng] with labeled arrows pointing to parts, educational diagram"` |

**Layout slide trò chơi:**
```
┌──────────────────────────────────────────────────┐
│ [Accent Badge: LUYỆN TẬP / TRÒ CHƠI]             │
│ Tiêu đề: "Cùng chơi nào!" (26pt Bold)            │
│ Hướng dẫn ngắn (17pt)                             │
│                                                    │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐        │
│  │  ẢNH 1   │  │  ẢNH 2   │  │  ẢNH 3   │        │
│  │ (2in×2in)│  │ (2in×2in)│  │ (2in×2in)│        │
│  ├──────────┤  ├──────────┤  ├──────────┤        │
│  │[Card chọn│  │[Card chọn│  │[Card chọn│        │
│  │ Item 1]  │  │ Item 2]  │  │ Item 3]  │        │
│  └──────────┘  └──────────┘  └──────────┘        │
│              ← Animation: hiện tuần tự →          │
└──────────────────────────────────────────────────┘
```

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

## Anti-Bug Checklist & Kiểm tra tự động (BẮT BUỘC)

> **Chạy kiểm tra SAU MỖI LẦN tạo slide. Nếu có FAIL → sửa và chạy lại.**

### Bảng quy tắc kỹ thuật cứng (Hard Rules):

| # | Quy tắc | Kiểm tra tự động | Hậu quả nếu vi phạm |
|---|---------|-------------------|----------------------|
| 1 | KHÔNG vẽ shape có `top < 1.15in` | `assert shape.top >= Inches(1.15)` | Logo bị che |
| 2 | KHÔNG vẽ shape có `top + height > 6.35in` | `assert shape.top + shape.height <= Inches(6.35)` | Chân trang bị che |
| 3 | KHÔNG dùng `add_shape(RECT, 0, 0, SLIDE_W, SLIDE_H)` | grep script | File corrupt / Repair |
| 4 | KHÔNG thêm shape footer mới | Đếm shapes ở Y > 6.35 = chỉ master | Footer chồng |
| 5 | Z-order: `insert(2, sp)` KHÔNG BAO GIỜ `insert(0, sp)` | grep `insert(0` | XML schema vỡ |
| 6 | 3 loại text color trong palette | Check keys | Chữ bị mất tương phản |
| 7 | Font tiêu đề 24-28pt, nội dung 18-20pt | Verify font sizes | Chữ quá to/nhỏ |
| 8 | Tối đa 3-4 bullets/slide | Count lines ≤ 4 | Chữ tràn/chồng |
| 9 | Mỗi slide nội dung có ≥ 1 hình ảnh | Count pictures ≥ 1 | Slide trống |
| 10 | Mỗi bullet có ảnh riêng (per-bullet) | Count pictures ≥ bullets | Thiếu minh họa |
| 11 | Slide practice/activity có animation | Check `p:timing` XML | Items hiện cùng lúc |
| 12 | Slide trò chơi có ảnh minh họa ≥ 1 | Count pictures on game slides | HS không có gì nhìn |
| 13 | Card + accent bar phải group | Check `<p:grpSp>` | Bố cục rời rạc |

### Script kiểm tra `_verify_slide_v2.py`:
- Chạy tự động sau mỗi lần tạo slide
- Input: file `.pptx` vừa tạo
- Output: Báo cáo PASS/FAIL cho từng slide + từng quy tắc
- Nếu bất kỳ quy tắc 1-5 FAIL → BẮT BUỘC sửa và tạo lại

---

## Checklist kiểm tra trước khi giao slide

### Kỹ thuật:
- [ ] Mọi shape nằm trong Vùng An Toàn (Y 1.15in → 6.35in)
- [ ] Logo UNIGO không bị che (không shape nào ở Y < 1.15in)
- [ ] Chân trang UNIGO có trên mọi slide (master Picture 9)
- [ ] Z-order đúng: background `insert(2, sp)`, text thêm sau
- [ ] Không có shape footer tự tạo
- [ ] Font chữ đúng chuẩn (cỡ phù hợp lứa tuổi)

### Nội dung:
- [ ] Ngôn ngữ hướng tới HS (không có "HS nhận biết được...", "Mục tiêu:...")
- [ ] Không có slide liệt kê mục tiêu/năng lực kiểu giáo án
- [ ] Tối đa 3-4 bullets/slide

### Hình ảnh (MỚI):
- [ ] **Mỗi bullet point có ảnh riêng** (per-bullet image)
- [ ] Ảnh per-bullet đủ lớn: ≥ 2in×2in (TH Grid) hoặc 2in×1.5in (THCS Row)
- [ ] Slide trò chơi có ảnh minh họa đủ lớn cho HS nhìn
- [ ] Layout ảnh đúng cấp: Grid Flashcard (TH) / Horizontal Row (THCS)

### Animation (MỚI):
- [ ] Có hiệu ứng transition giữa các slide
- [ ] Slide câu hỏi/ghép nối: items animation tuần tự theo click
- [ ] Slide ghép nối: vế A hiện trước → click → vế B hiện
- [ ] Ảnh và text cùng item hiện ĐỒNG THỜI (cùng 1 click)

## Thư viện sử dụng
- `python-pptx`: Tạo .pptx
- `PyMuPDF (fitz)`: Đọc SGK PDF, trích xuất hình
- `Pillow`: Xử lý ảnh nếu cần crop

## Cải tiến liên tục
Sau mỗi lần tạo slide, ghi nhận feedback vào:
`D:\UNIGO\.agents\skills\tao-slide-bai-giang\references\improvement_log.md`


---

## CẬP NHẬT CHUẨN ĐỒNG BỘ KHBD – SLIDE TRỰC QUAN UNIGO (GIAI ĐOẠN 5 & 8)

1. **Cấu trúc 12 Slide chuẩn đồng bộ 1-1 với KHBD mới:**
   - **Slide 1:** Tiêu đề bài học + Chủ đề + Trang SGK + Hình minh họa SGK mở đầu (Hero Image).
   - **Slide 2:** Mục tiêu bài học dành cho học sinh (4 thẻ màu trực quan, ngôn ngữ thân thiện với HS Tiểu học).
   - **Slide 3:** Hoạt động 1 – Tình huống khởi động (Tranh tình huống SGK + Câu chuyện của An, Minh, Khoa).
   - **Slide 4:** Hoạt động 1 – Câu hỏi gợi mở & Thảo luận nhóm đôi (Bóng thoại câu hỏi + Gợi ý dẫn dắt).
   - **Slide 5:** Hoạt động 2 – Khám phá kiến thức mới 1 (Hình ảnh SGK chiếm 60% diện tích + Nhiệm vụ khám phá).
   - **Slide 6:** Hoạt động 2 – Sơ đồ quy trình / Các bước thao tác chuẩn (Infographic 100% nhãn Tiếng Việt có mũi tên tuần tự Bước 1 $\rightarrow$ Bước 4).
   - **Slide 7:** Hoạt động 2 – Hộp Kiến thức mới / Chốt ghi nhớ SGK (Nền vàng cam nổi bật + Biểu tượng bóng đèn).
   - **Slide 8:** Hoạt động 3 – Trò chơi Trắc nghiệm nhanh 4 đáp án A, B, C, D.
   - **Slide 9:** Hoạt động 3 – Nhiệm vụ Thực hành phân hóa 2 mức độ (Cột trái: Mức Cơ bản | Cột phải: Thử thách Nâng cao).
   - **Slide 10:** Hoạt động 4 – Tình huống Vận dụng thực tiễn & Góc Công dân số / Trí tuệ nhân tạo (AI) an toàn.
   - **Slide 11:** Bảng Tự đánh giá của học sinh sau bài học (3 mức độ: 🌟 Hoàn thành tốt | 👍 Hoàn thành | 💪 Cần cố gắng).
   - **Slide 12:** Tổng kết bài học & Dặn dò chuẩn bị bài tiếp theo.
2. **Bảo tồn tuyệt đối mẫu UNIGO & Vùng an toàn:**
   - Luôn nhân bản từ `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`, giữ nguyên Logo góc trái trên (`Y <= 1.12in`) và Thanh chân trang UNIGO (`Y >= 6.40in`). Toàn bộ nội dung đặt trong `1.15in <= Y <= 6.35in`.


---

## POWERPOINT QUALITY ASSURANCE – UNIGO (QUY TRÌNH KIỂM ĐỊNH 9 BƯỚC BẮT BUỘC)

> [!CAUTION]
> **NGUYÊN TẮC BẮT BUỘC:** Tuyệt đối không được xem việc chạy script xuất file `.pptx` thành công là đồng nghĩa với việc slide đã hoàn thành. Mọi file `.pptx` sau khi tạo BẮT BUỘC phải đi qua đầy đủ 9 bước kiểm định dưới đây trước khi bàn giao cho người dùng.

### Bước 1: Kiểm tra và làm sạch Template UNIGO trước khi thiết kế
- Mở trực tiếp từ file chuẩn `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx` (`13.333 in × 7.50 in`, tỷ lệ 16:9).
- **Bảo toàn tuyệt đối Slide Master:** Giữ nguyên `Picture 7` (Logo UNIGO tại `left=0.17in, top=0.15in, right=1.12in, bottom=1.09in`) và `Picture 9` (Thanh chân trang UNIGO tại `left=0.00in, top=6.43in, right=13.40in, bottom=7.67in`).
- **Xóa sạch 4 slide mẫu mặc định trong `prs.slides._sldIdLst` trước khi thêm slide mới:** File mẫu gốc có sẵn 4 slide (`Slide 1: TRƯỜNG TIỂU HỌC VÀ THCS UNIGO`, `Slide 2: Môn Tin học lớp`, `Slide 3: Môn Tin học lớp`, `Slide 4: Thanks for listening`). Bắt buộc xóa cả 4 slide này khỏi `_sldIdLst` trước khi tạo bài mới để:
  - Không bao giờ để sót slide chỉ có tiêu đề (`Môn Tin học lớp`) không có nội dung ở Slide 2, 3.
  - Không bao giờ để slide kết thúc (`Thanks for listening`) nằm sai vị trí ở giữa bài (Slide 4).
  - Không bao giờ sao chép nhầm `TextBox 5` (`TRƯỜNG TIỂU HỌC VÀ THCS UNIGO`) đè lên tiêu đề của các slide nội dung.

### Bước 2: Thiết lập Vùng hiển thị an toàn (Safe Zone) nghiêm ngặt
- **Vùng cấm Logo UNIGO:** `X < 1.15 in` và `Y < 1.09 in` — Không đặt bất kỳ đối tượng nào đè vào vùng này.
- **Vùng Thanh tiêu đề Slide (Header Bar):** Đặt tại `X = 1.30 in .. 12.85 in`, `Y = 0.28 in .. 0.92 in` (nằm hoàn toàn bên phải Logo UNIGO).
- **Vùng Nội dung chính (Body Safe Zone):** `X = 0.50 in .. 12.85 in`, `Y = 1.15 in .. 6.30 in`.
- **Vùng cấm Chân trang UNIGO:** `Y >= 6.40 in` — Tuyệt đối không để khung chữ hoặc hình ảnh tràn xuống dưới `Y = 6.38 in`.

### Bước 3: Kiểm soát kích thước, vị trí đối tượng, Font chữ và Tỷ lệ ảnh
- **Gộp khung nền và văn bản làm một đối tượng duy nhất:** Khi tạo thẻ màu có chữ, ghi trực tiếp văn bản vào `text_frame` của `MSO_SHAPE.ROUNDED_RECTANGLE` thay vì đặt thêm `TextBox` rời đè lên trên (tránh lỗi lệch vị trí và chồng lấn bounding box).
- **Tự động co giãn cỡ chữ theo chiều cao khung (`Auto-fit Height Calculation`):** Trước khi ghi văn bản vào khung có kích thước `(width_in, height_in)`, phải tính số dòng sau khi ngắt dòng (`total_wrapped_lines`) và chiều cao cần thiết `req_h = total_wrapped_lines * (font_pt / 72.0 * 1.32) + 0.12`. Nếu `req_h > height_in`, tự động giảm `font_pt` (trong ngưỡng `14pt – 20pt`) hoặc chia tách slide, tuyệt đối không để chữ tràn khỏi khung.
- **Cấm ký tự Emoji lỗi font (`□`) và ký tự `\x0b`:** Loại bỏ toàn bộ ký tự Emoji ngoài vùng hỗ trợ của font `Arial` / `Times New Roman` (dải `U+10000..U+10FFFF`) và ký tự ngắt dòng dọc `\x0b` để không xuất hiện các ô vuông `□` trên slide.
- **Bảo toàn tỷ lệ khung hình ảnh (Aspect Ratio):** Luôn tính `ratio = w_px / h_px` từ ảnh gốc và căn giữa trong vùng chứa `(box_l, box_t, box_w, box_h)`, tuyệt đối không kéo giãn làm méo ảnh. Mỗi slide trực quan trong bài phải sử dụng một hình ảnh/sơ đồ riêng biệt, không lặp lại cùng 1 ảnh cho nhiều slide.

### Bước 4: Kiểm soát Nội dung dành cho Học sinh & Trình tự 11 Slide chuẩn
- **Chuyển hóa 100% ngôn ngữ dành cho Học sinh:** Tuyệt đối không copy nguyên văn kịch bản giáo án của giáo viên (`- GV giới thiệu...`, `- GV y/c HS...`, `a. Mục tiêu...`, `b. Cách thức tiến hành...`) lên slide.
- **Trình tự 11 Slide chuẩn sư phạm UNIGO:**
  1. `Slide 01 – Trang bìa:` Tên trường, môn/khối lớp, chủ đề, tên bài học, trang SGK, giáo viên và hình ảnh đại diện bài học.
  2. `Slide 02 – Khởi động:` Tình huống thực tế mở đầu và tranh minh họa tình huống SGK.
  3. `Slide 03 – Câu hỏi gợi mở:` Câu hỏi thảo luận nhóm đôi và gợi ý kết nối vào bài học mới.
  4. `Slide 04 – Mục tiêu bài học:` 3 dải băng mục tiêu rõ ràng, thân thiện với học sinh.
  5. `Slide 05 – Hình thành kiến thức:` Hình minh họa SGK / Sơ đồ tư duy khám phá kiến thức mới.
  6. `Slide 06 – Minh họa và thực hành mẫu:` Sơ đồ quy trình 4 bước thao tác tuần tự Tiếng Việt 100%.
  7. `Slide 07 – Chốt kiến thức:` Hộp kiến thức Ghi nhớ trọng tâm SGK nổi bật.
  8. `Slide 08 – Luyện tập:` Câu hỏi trắc nghiệm củng cố và Nhiệm vụ thực hành phân hóa 2 mức độ (Cơ bản & Nâng cao).
  9. `Slide 09 – Vận dụng:` Bài tập vận dụng thực tiễn & Cẩm nang Công dân số / AI an toàn.
  10. `Slide 10 – Tổng kết & Tự đánh giá:` Bảng tự đánh giá 3 mức độ (`Hoàn thành tốt` | `Hoàn thành` | `Cần cố gắng`) và Dặn dò về nhà.
  11. `Slide 11 – Slide Cảm ơn / Kết thúc:` Đặt đúng ở vị trí cuối cùng của bài giảng.

### Bước 5: Render PPTX sang Hình ảnh sau khi xuất
- Chạy công cụ render trực quan (`render_pptx_to_folder`) để xuất toàn bộ từng slide thành ảnh PNG độ phân giải `1333 × 750 px` (`slide_01.png` $\rightarrow$ `slide_11.png`) và ghép thành **Ảnh tổng quan toàn bộ slide (`overview_all_slides.png`)**.

### Bước 6: Kiểm tra Cấu trúc tự động & Kiểm tra Trực quan bằng mắt (Visual Inspection)
- Chạy script kiểm toán cấu trúc (`audit_single_pptx`) kiểm tra 7 nhóm lỗi: (1) Slide rỗng/chỉ có tiêu đề, (2) Slide kết thúc sai vị trí, (3) Tràn chữ (`ReqH > BoxH`), (4) Chồng lấn hình–chữ, (5) Chồng lấn chữ–chữ, (6) Vi phạm vùng Logo/Chân trang UNIGO, (7) Cỡ chữ nhỏ (`< 13.5pt`) hoặc lỗi ký tự `□`.
- Dùng công cụ `read_file` mở trực tiếp file ảnh `overview_all_slides.png` và các slide có nhiều thành phần đồ họa để kiểm tra trực quan 100%.

### Bước 7: Tự động sửa lỗi (Auto-Remediation)
- Nếu phát hiện bất kỳ lỗi nào ở Bước 6: tự động xóa slide thừa, dịch chuyển tọa độ thoát khỏi vùng Logo/Chân trang, mở rộng chiều cao khung hoặc giảm cỡ chữ để hết tràn chữ, tách khoảng cách các đối tượng giao nhau và khôi phục nội dung thiếu.

### Bước 8: Render lại và Kiểm tra lần hai (Second-Pass Verification)
- Sau khi sửa lỗi ở Bước 7, bắt buộc chạy lại trình kiểm toán và render lại ảnh `overview_all_slides.png` lần thứ hai để xác nhận `issue_count == 0` trên toàn bộ các slide.

### Bước 9: Chỉ bàn giao phiên bản đạt chuẩn 100%
- Chỉ bàn giao cho người dùng khi toàn bộ file `.pptx` đạt **0 lỗi cấu trúc & hiển thị**, kèm theo ảnh tổng quan `overview_all_slides.png` đã được kiểm chứng trực quan.


---

## SKILL B – UNIGO STUDENT-FACING SLIDE DESIGN (CHUẨN THIẾT KẾ SLIDE LẤY HỌC SINH LÀM TRUNG TÂM & QUẢN LÝ DỰ ÁN)

> [!IMPORTANT]
> **Nguyên tắc tối thượng của `SKILL B`:**
> **KHBD (`SKILL A`) mô tả cách giáo viên dạy; Slide (`SKILL B`) thể hiện những gì học sinh cần quan sát, suy nghĩ, khám phá và ghi nhớ.**
> Slide **KHÔNG** phải bản sao KHBD, **KHÔNG** phải tài liệu hướng dẫn giáo viên và **TUYỆT ĐỐI KHÔNG** đưa kịch bản điều hành lớp học (`"GV tổ chức..."`, `"HS nhận ra..."`, `"Bước 1..."`) hoặc bảng mã năng lực chuyên môn (`NLa–NLe`, `CV 3456`) lên màn hình.

### 1. Quy trình 4 Bước Phân tích Trước khi Thiết kế Slide
1. **Bước 1 – Đọc SGK (`D:\UNIGO\SGK\`):** Xác định tên bài, chủ đề, yêu cầu cần đạt, các đơn vị kiến thức, thuật ngữ chuẩn, hình ảnh minh họa gốc, sơ đồ/bảng biểu, câu hỏi luyện tập và nhiệm vụ vận dụng.
2. **Bước 2 – Đọc KHBD tương ứng (`D:\UNIGO\KHBD_Tin_học_Chuẩn_Mới\` hoặc `D:\UNIGO\KHBD_Tin_học\`):** Nắm trình tự hoạt động, thời điểm xuất hiện từng nội dung và các điểm chốt kiến thức (không sao chép nguyên văn lời dẫn GV).
3. **Bước 3 – Lập Bản đồ nội dung Slide (`Slide Content Map`):** Lưu tại `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\03_LESSON_CONTENT_MAP\` với bảng 4 cột: `Thứ tự | Nội dung kiến thức (Ngôn ngữ HS) | Mục đích slide | Hình ảnh cần có`.
4. **Bước 4 – Kiểm tra tính thống nhất:** Đối chiếu Bản đồ slide với SGK và KHBD, đảm bảo không thiếu kiến thức trọng tâm, không sai thuật ngữ SGK và không lẫn ngôn ngữ giáo viên.

### 2. Cấu trúc & Các Dạng Bố cục Slide Linh hoạt (`Layouts A – G`)
Không giới hạn cứng số lượng slide; thiết kế linh hoạt theo độ phức tạp của bài học (thông thường `12 – 15 slides/bài`):
- **A. Slide giới thiệu:** Tên chủ đề, tên bài học, hình ảnh chủ đề SGK và câu dẫn ngắn tạo hứng thú cho học sinh.
- **B. Slide khởi động:** Câu hỏi tình huống, hình ảnh quan sát cắt chuẩn từ SGK, câu đố hoặc thử thách mở đầu.
- **C. Slide mục tiêu:** 3 yêu cầu cần đạt viết bằng ngôn ngữ học sinh (*"Sau bài học này em sẽ..."*). Không đưa bảng phân tích mã năng lực của giáo viên lên slide.
- **D. Slide hình thành kiến thức (Nhiều slide tách theo từng đơn vị kiến thức):** Sử dụng linh hoạt hình ảnh SGK có chú thích (Callout), sơ đồ cấu tạo (đánh số `1-2-3-4`), sơ đồ quy trình có mũi tên nối, bảng so sánh, tình huống Đúng/Sai (`Nên / Không nên`).
- **E. Slide ghi nhớ & chốt kiến thức:** Sử dụng **Sơ đồ tư duy (Mindmap)**, bảng tổng hợp hoặc từ khóa cốt lõi (không viết thành đoạn văn dài).
- **F. Slide luyện tập (BẮT BUỘC TÁCH RIÊNG CÂU HỎI VÀ ĐÁP ÁN):**
  - *Slide Câu hỏi:* Hiển thị đề bài (trắc nghiệm, ghép cặp, đúng/sai, nhận diện hình ảnh SGK) để học sinh suy nghĩ và trả lời trước.
  - *Slide Đáp án & Giải thích (Slide kế tiếp):* Hiển thị đáp án đúng nổi bật kèm lời giải thích vì sao đúng.
- **G. Slide vận dụng:** Đưa kiến thức vào giải quyết vấn đề thực tiễn đời sống và giáo dục an toàn số / an toàn điện.

### 3. Quy trình Xử lý Hình ảnh & Kiểm định 15 Bước
1. **Ưu tiên hình ảnh SGK gốc:** Cắt chuẩn (`PIL.Image.crop`) từ trang SGK độ phân giải cao (`1700 × 2200 px`), giữ nguyên tỷ lệ khung hình (`aspect ratio`), không làm biến dạng hay cắt mất chữ/thông tin. Lưu vào `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\06_IMAGES\`.
2. **Tạo hình ảnh bổ sung & Sơ đồ tư duy:** Khi cần tổng hợp kiến thức ở phần Ghi nhớ hoặc mô phỏng quy trình, tạo sơ đồ trực quan Tiếng Việt 100%, màu sắc hài hòa, đúng lứa tuổi.
3. **Kiểm định chất lượng 15 bước bắt buộc:**
   - Sau khi xuất PPTX từ mẫu `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`, bắt buộc chạy `render_pptx_visual_qa.py` và `deep_audit_all_pptx.py` để render toàn bộ slide ra ảnh PNG (`1333 × 750 px`) và ghép `overview_all_slides.png` tại `07_QUALITY_CONTROL\`.
   - Kiểm tra trực quan đủ 15 mục: (1) Số lượng slide, (2) Thứ tự nội dung, (3) Nội dung từng slide 100% ngôn ngữ HS, (4) Render PNG, (5) Ảnh tổng quan, (6) Duyệt từng slide, (7) Không có slide trống, (8) Không tràn chữ, (9) Không chồng lấn, (10) Không lỗi font/ô vuông `□`, (11) Nguyên vẹn Logo & Chân trang UNIGO, (12) Ảnh sắc nét đúng tỷ lệ, (13) Khớp Bản đồ nội dung, (14) Chuẩn kiến thức SGK, (15) Khớp trình tự KHBD.

### 4. Cơ chế Quản lý Phiên làm việc & Checkpoint (`D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\`)
- Khi bắt đầu phiên mới, luôn đọc `PROJECT_MASTER.md`, `SESSION_CHECKPOINT.md`, `PROGRESS_TRACKER.md` và `CHANGELOG.md`.
- Sau mỗi công đoạn (lập bản đồ slide, cắt ảnh, xuất PPTX, kiểm định 15 bước), cập nhật ngay vào `PROGRESS_TRACKER.md` (`.xlsx`), `SESSION_CHECKPOINT.md` và `CHANGELOG.md`.


---

## MODULE BỔ SUNG: HIGH-QUALITY VISUAL STORYTELLING – UNIGO (bắt buộc cho mọi slide Tin học)

**Tư duy:** hình ảnh trước, chữ sau. Hình giúp HS nhìn thấy kiến thức, hiểu khái niệm, hình dung quy trình, khám phá tình huống, ghi nhớ lâu.

### 1. Mật độ hình ảnh (định hướng)
1. Mỗi slide kiến thức có ≥ 1 hình minh họa có nghĩa; khái niệm phức tạp dùng 2–4 hình/sơ đồ.
2. Slide giới thiệu, khởi động, tình huống: hình lớn chiếm ~40–65% diện tích.
3. Slide thao tác: ảnh từng bước hoặc ảnh chụp giao diện phần mềm.
4. Slide tổng kết: sơ đồ tư duy / infographic.
5. Không tạo slide chỉ có chữ nếu có thể diễn đạt bằng hình/sơ đồ mà vẫn chính xác.

### 2. Năm loại hình phải phối hợp
A. Ảnh SGK cắt đúng vùng (kiến thức cần chính xác) · B. Hình tạo mới (nhân vật, bối cảnh, tình huống, thiết bị) · C. Sơ đồ (cấu tạo, quy trình, phân loại, bảng so sánh, infographic) · D. Ảnh chụp giao diện thực tế (phần mềm, thao tác) · E. Hình tương tác (đố hình, đúng/sai, tìm lỗi, ghép cặp, thử thách quan sát).

### 3. Nhân vật & phong cách
- Mỗi bài/khối có nhân vật dẫn chuyện nhất quán (mẫu: robot **Bít**, hàm `mascot()` trong `scratch/visual_storytelling_bai3.py`), biểu cảm theo ngữ cảnh (happy/think/alert), hành động gắn với kiến thức. Không dùng nhân vật có bản quyền, không để nhân vật chỉ tạo dáng.
- Phong cách: 2D vector / modern colorful educational / isometric infographic / friendly technology; giữ nhất quán trong cả bài.
- Hình AI/ảnh sinh mới (khi dùng công cụ tạo ảnh): phải kiểm tra cấu tạo thiết bị, chữ tiếng Việt, ngoại hình nhân vật trước khi dùng.

### 4. Quy trình
1. **Lập Danh sách hình ảnh** trước khi tạo slide, lưu tại `03_LESSON_CONTENT_MAP/Danh_sach_hinh_anh_<Lớp>_<Bài>.md` gồm: Mã hình, Nội dung, Slide dùng, Mục tiêu sư phạm, Nguồn, Tỷ lệ, Phong cách.
2. Thiếu nguồn SGK thì chủ động tạo hình mới; không bỏ qua hình vì tốn thời gian.
3. Không lặp một hình trên nhiều slide nếu không có chủ đích (slide đáp án dùng hình đánh dấu đáp án riêng, ví dụ `VS06b`).
4. Không đặt tiêu đề trùng lặp bên trong ảnh và trên banner slide.

### 5. Không chấp nhận
Slide toàn chữ kéo dài · icon nhỏ thay hình minh họa · hình không liên quan/mờ/méo · thiết bị vẽ sai cấu tạo · chữ sai chính tả trong hình · nhân vật đổi ngoại hình · hình trang trí vô giá trị học tập.

### 6. Kiểm định hình ảnh (chấm riêng, ngoài bố cục & văn bản)
Phong phú · sinh động · phù hợp HS · liên quan kiến thức · chính xác (thiết bị, giao diện, thao tác Tin học) · nhất quán phong cách · giúp hiểu bài. Slide còn trống nhiều hoặc thiếu hình phải thiết kế lại; không đánh giá đạt chỉ vì bố cục đẹp và chữ đúng.
