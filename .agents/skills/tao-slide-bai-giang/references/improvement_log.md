# Nhật ký cải tiến Slide Bài giảng UNIGO

## 2026-08-13 (v3) — Nâng cấp: Per-Bullet Images, Animation tuần tự, Game Images, Anti-Bug Checklist

- **Yêu cầu từ user (4 nhóm):**
  1. Mỗi bullet point phải có ảnh AI riêng (tạo prompt từ nội dung bullet)
  2. Slide câu hỏi/ghép nối cần animation tuần tự (hiện dần theo click) + ảnh mỗi item
  3. Slide trò chơi bắt buộc có hình ảnh minh họa đủ lớn cho HS chơi
  4. Anti-bug checklist kỹ thuật để tránh mọi lỗi đã gặp (logo/footer/z-order/font)

- **Thay đổi đã áp dụng:**
  1. **Per-Bullet Images (Bước 2.5 mới trong SKILL.md):**
     - Mỗi bullet → 1 ảnh AI (prompt = nội dung bullet + context lứa tuổi)
     - Layout phân cấp: Grid Flashcard (TH, ảnh 2in×2in) vs Horizontal Row (THCS, ảnh 2in×1.5in)
     - Chuỗi fallback 3 tầng: SGK → AI → ảnh chung
     - Tối đa 3-4 bullets/slide, > 4 thì chia 2 slides
  2. **Animation tuần tự (Bước 7.2 mới):**
     - XML `p:timing` animation `appear` per-shape on click
     - Items practice/activity hiện lần lượt (HS suy nghĩ trước khi GV click)
     - Ghép nối: vế A + ảnh hiện trước → click → vế B hiện
     - Helper: `add_appear_animation()`, `add_group_animation()`
  3. **Game Images (Bước 7.3 mới):**
     - Bảng 6 loại trò chơi + prompt AI mẫu riêng
     - Layout: tiêu đề (40%) + grid ảnh 2in×2in + card (60%)
  4. **Anti-Bug Checklist (Section mới):**
     - 13 quy tắc kỹ thuật cứng kiểm tra tự động
     - Script `_verify_slide_v2.py` chạy sau mỗi lần tạo slide
     - Quy tắc 1-5 FAIL → bắt buộc sửa lại

- **Files đã cập nhật:**
  - `SKILL.md` — Thêm Bước 2.5, cập nhật Bước 5 (bảng deck), Bước 7 (animation nâng cao), Anti-Bug Checklist, Checklist cuối
  - `AGENTS.md` Section VII — Thêm quy tắc 7, 8, 9
  - `improvement_log.md` — Entry này

## 2026-08-10 (v2) - Fix: Logo/Chân trang bị che, Tương phản kém

- **Vấn đề nhận diện (từ user feedback + screenshot)**:
  1. File lỗi cần Repair khi mở trong PowerPoint (slide removal method gây corrupt)
  2. Nền shape `RECTANGLE(0, 0, SLIDE_W, SLIDE_H)` phủ toàn slide → che logo master + chân trang master
  3. Custom footer bar chồng lên footer master có sẵn (Picture 9)
  4. Banner tại Y=0.8 đè lên logo (Logo kết thúc tại Y=1.09in)
  5. Chữ nhạt trên nền nhạt → khó đọc, thiếu tương phản
- **Fix đã áp dụng**:
  1. **Clamp mọi shape vào SAFE ZONE** Y=1.15in→6.35in: `actual_top = max(top, 1.15)`
  2. **Xóa custom footer** — template master ĐÃ CÓ chân trang (Picture 9 tại Y=6.43)
  3. **Z-Order đúng**: Background → `send_to_back`; Text → thêm sau (tự nằm trên)
  4. **High-contrast palettes**: 3 loại `text_on_primary`, `text_on_bg`, `text_on_card`
  5. **Verified template**: Logo=Y0.15→1.09, Footer=Y6.43→7.66, Slide=13.33×7.50in
- **Quy tắc đã cập nhật**: SKILL.md, AGENTS.md Section VII

## 2026-08-10 - Chuyển đổi triết lý: Student-Facing Slides

- **Vấn đề nhận diện**: Slide Tuần 02 được tạo dưới dạng mô tả lại KHBD (liệt kê mục tiêu, năng lực, phẩm chất) — không phù hợp để chiếu cho học sinh nhìn và làm theo trong lớp.
- **Thay đổi cốt lõi**:
  1. **Triết lý Student-Facing**: Slide phải là thứ HS nhìn lên màn hình và làm theo, KHÔNG phải bản mô tả giáo án.
  2. **Ngôn ngữ hướng tới HS**: "Em hãy...", "Bước 1:...", "Quan sát hình..." thay vì "HS nhận biết được...", "Mục tiêu:...".
  3. **AI Content Generation**: Agent tự phân tích SGK + KHBD và sinh nội dung phù hợp, không copy/paste từ KHBD.
  4. **Bỏ slide Mục tiêu/Năng lực**: Mục tiêu được lồng ghép tự nhiên vào nội dung hoạt động.
  5. **Điều chỉnh theo lứa tuổi**: Tiền TH/Lớp 1-2 cực đơn giản, Lớp 3-5 tăng dần, Lớp 6-8 có thuật ngữ.
- **Hành động**: Xóa 9 file slide Tuần 02 cũ. Slide mới sẽ được tạo lại theo skill đã cập nhật.
- **Giữ nguyên**: Toàn bộ quy tắc kỹ thuật (chân trang, đầu trang, vùng an toàn, font, template, Card Grid, Accent Bar, Color Palette).

## 2026-07-28 - Nâng cấp toàn diện Thiết kế Giao diện Slide (Modern High-Aesthetic Design System)

- **Vấn đề nhận diện**: Slide tạo ban đầu bị đánh giá là chưa đủ thẩm mỹ (thiếu điểm nhấn thị giác, hộp văn bản thô, thiếu cấu trúc container).
- **Cải tiến thiết kế đã áp dụng**:
  1. **Top Header Banner**: Mọi slide nội dung đều có thanh Banner màu chủ đạo phía trên kèm Badge Pill nổi bật (`[ TIN HỌC X • BÀI Y ]`).
  2. **Modern Card Grid System**: Đóng gói nội dung trong các Card bo tròn (`ROUNDED_RECTANGLE`) màu trắng tinh (`#FFFFFF`) trên nền Slate nhạt (`#F8FAFC`).
  3. **Left Accent Strips**: Mỗi card có dải vạch màu nhấn nổi bật bên lề trái (rộng 0.6cm) giúp phân chia và thu hút ánh nhìn.
  4. **Number Badges (`01`, `02`, `03`)**: Sử dụng huy hiệu số thứ tự dạng hình tròn/Pill màu rực rỡ cho mục tiêu và các nội dung cốt lõi.
  5. **Image Framing**: Ảnh AI được lồng bên trong Card container bo góc chuyên nghiệp, không bị thả trôi tự do.
  6. **Hiệu ứng XML**: Tích hợp XML Transitions (chuyển slide mượt) và XML Animations (xuất hiện từng bước).
- **Kết quả**: Áp dụng thành công cho toàn bộ bộ slide Lớp 3, 4, 5, 6, 7, 8. Đã được ghi nhận vào quy chuẩn thiết kế của Agent.
