---
name: website-nang-luc
description: >
  Quy trình, tiêu chuẩn kỹ thuật, căn cứ pháp quy và cấu trúc giao diện
  khi phát triển, bảo trì và mở rộng Website Ôn tập Năng lực & Phẩm chất UNIGO.
  Sử dụng khi cần cập nhật câu hỏi, mở rộng chức năng tra cứu, chỉnh sửa giao diện
  hoặc đồng bộ kiến thức từ các văn bản pháp quy giáo dục mới.
---

# Skill: Quản trị & Phát triển Website Ôn tập Năng lực & Phẩm chất UNIGO

Tài liệu này quy chuẩn toàn bộ kiến trúc, cơ sở dữ liệu pháp quy, hệ thống câu hỏi kiểm tra, và quy tắc thiết kế giao diện cho Website Ôn tập Năng lực & Phẩm chất dành cho Giáo viên trường Tiểu học & THCS UNIGO.

---

## I. Căn cứ Pháp quy & Cơ sở Dữ liệu chuẩn

> [!IMPORTANT]
> **QUY TẮC BẤT BIẾN:** Tuyệt đối **KHÔNG ĐƯỢC TỰ Ý THÊM, BỚT, SỬA ĐỔI HOẶC XÓA BỎ** nội dung mô tả, định nghĩa, tên gọi hay mã chuẩn bậc đã được ban hành chính thức trong các văn bản pháp quy của Bộ Giáo dục và Đào tạo.

### 1. Khung Năng lực Số (NLS) — Thông tư 02/2025/TT-BGDĐT & Công văn 3456/BGDĐT-CNTT
- **Cấu trúc:** 6 Miền năng lực & 24 Thành tố cốt lõi.
  - **Miền I: Khai thác dữ liệu và thông tin** (Thành tố 1.1, 1.2, 1.3)
  - **Miền II: Giao tiếp và hợp tác trong môi trường số** (Thành tố 2.1, 2.2, 2.3, 2.4, 2.5, 2.6)
  - **Miền III: Sáng tạo nội dung số** (Thành tố 3.1, 3.2, 3.3, 3.4)
  - **Miền IV: An toàn** (Thành tố 4.1, 4.2, 4.3, 4.4)
  - **Miền V: Giải quyết vấn đề** (Thành tố 5.1, 5.2, 5.3, 5.4)
  - **Miền VI: Ứng dụng trí tuệ nhân tạo (AI)** (Thành tố 6.1, 6.2, 6.3)
- **5 Bậc chuẩn & Ký hiệu CB (Chuẩn Bậc):**
  - **Bậc 1 (`CB1`):** Tiền Tiểu học & Lớp 1, 2, 3 (Có sự hướng dẫn của giáo viên)
  - **Bậc 2 (`CB2`):** Lớp 4, 5 (Tự chủ nhiệm vụ quen thuộc + hỗ trợ khi cần)
  - **Bậc 3 (`CB3`):** Lớp 6, 7 (Tự chủ hoàn toàn nhiệm vụ thông thường)
  - **Bậc 4 (`CB4`):** Lớp 8, 9 (Độc lập, chủ động lựa chọn công cụ theo nhu cầu)
  - **Bậc 5 (`CB5`):** Lớp 10, 11, 12 (Nâng cao, làm chủ và hướng dẫn người khác)
- **Lưu ý chống nhầm lẫn:**
  - Lớp 6–7 là **Bậc 3 (`CB3`)**, KHÔNG PHẢI Bậc 2.
  - Lớp 8–9 là **Bậc 4 (`CB4`)**, KHÔNG PHẢI Bậc 3.

### 2. Khung Năng lực Trí tuệ Nhân tạo (AI) — Quyết định 3439/QĐ-BGDĐT
- **4 Chủ đề trụ cột (A – D):**
  - **Chủ đề A:** Tư duy lấy con người làm trung tâm (A1. Tính chủ động của con người; A2. AI vì sự tiến bộ của con người).
    * *Nguyên lý then chốt:* AI không có cảm xúc thật mà chỉ mô phỏng; con người luôn là chủ thể ra quyết định và chịu trách nhiệm tối thượng.
  - **Chủ đề B:** Đạo đức AI (B1. Khía cạnh đạo đức; B2. Tác động xã hội & thiên vị dữ liệu; B3. Nguyên tắc đạo đức và trách nhiệm xã hội).
  - **Chủ đề C:** Kỹ thuật và ứng dụng AI (C1. Đặc điểm chính; C2. Ứng dụng trong học tập/đời sống; C3. Công nghệ AI - ML, CV, NLP; C4. Dữ liệu huấn luyện; C5. Thuật toán dự đoán).
  - **Chủ đề D:** Thiết kế hệ thống AI (D1. Nhận diện & hình thành giải pháp; D2. Cấu trúc, tương tác & cải tiến hệ thống AI).

### 3. 5 Phẩm chất cốt lõi — Chương trình GDPT 2018 (Thông tư 32/2018/TT-BGDĐT)
Bắt buộc có đủ 5 phẩm chất (KHÔNG được để thiếu phẩm chất nào):
1. **Yêu nước:** Tình yêu quê hương, bảo vệ thiên nhiên, tự hào di sản văn hóa, quảng bá cảnh đẹp dân tộc qua bài trình chiếu số.
2. **Nhân ái:** Tôn trọng sự khác biệt, lịch sự và văn minh trên không gian mạng, tích cực hỗ trợ bạn bè cùng thực hành.
3. **Chăm chỉ:** Kiên trì thực hành trên máy tính, chủ động tìm lỗi và sửa lỗi lập trình (debugging), hoàn thành bài tập đúng hạn.
4. **Trung thực:** Tự giác làm bài thực hành, tôn trọng bản quyền phần mềm, trích dẫn đầy đủ nguồn gốc thông tin khi tìm kiếm trên Internet.
5. **Trách nhiệm:** Sử dụng thiết bị đúng quy trình, giữ gìn tài sản chung, sắp xếp gọn gàng phòng máy tính/bộ Kit Robotics sau giờ học; tuân thủ an toàn mạng.

### 4. 3 Năng lực Chung — Công văn 5512/BGDĐT & CT GDPT 2018
1. **Tự chủ và tự học (TC&TH):** Chủ động đọc SGK, tự khám phá tính năng phần mềm, tự lập kế hoạch học tập.
2. **Giao tiếp và hợp tác (GT&HT):** Thảo luận nhóm, trình bày ý tưởng, phối hợp phân công nhiệm vụ trong nhóm.
3. **Giải quyết vấn đề và sáng tạo (GQVĐ&ST):** Phát hiện lỗi thuật toán, đề xuất giải pháp tối ưu code, thiết kế mô hình robot mới.

### 5. Năng lực Đặc thù môn học
- **Tin học (CT GDPT 2018):** NLa (Sử dụng CNTT), NLb (Ứng xử số), NLc (Khám phá & GQVĐ với CNTT), NLd (Học & tự học với CNTT), NLe (Hợp tác môi trường số).
- **Robotics (UNIGO):** NL1 (Nhận thức công nghệ), NL2 (Lắp ráp cơ chế), NL3 (Lập trình điều khiển), NL4 (Thử nghiệm & Troubleshooting), NL5 (Hợp tác & Sáng tạo kỹ thuật).

---

## II. Kiến trúc Hệ thống Tệp tin Website

Toàn bộ website được tổ chức độc lập, chạy trực tiếp trên trình duyệt (không phụ thuộc vào server hay npm build):

```
d:\UNIGO\website-nang-luc\
├── index.html       # Cấu trúc HTML ngữ nghĩa (Semantic HTML5)
├── style.css        # Hệ thống Design System (Light mode, typography to rõ, hiệu ứng mượt)
├── data.js          # Cơ sở dữ liệu pháp quy chuẩn (CV 3456, QĐ 3439, GDPT 2018, Ngân hàng câu hỏi)
└── app.js           # Logic giao tác, NLS inspector, modal xem chi tiết, và Quiz Engine phân tầng
```

---

## III. Quy chuẩn Thiết kế UI/UX & Typography

1. **Light Mode chuẩn mực:**
   - Nền trang web (`body`): `#f8fafc` (Slate-50 nhẹ nhàng).
   - Nền thẻ Card / Box: `#ffffff` trắng tinh khiết với viền mỏng `#e2e8f0` và đổ bóng mềm mại `0 4px 14px rgba(0,0,0,0.06)`.
   - Phần màu xen kẽ (`.section-alt`): `#f1f5f9` (Slate-100) tạo nhịp điệu thị giác rõ ràng.
2. **Typography kích thước lớn & Tương phản cao:**
   - Base `html`: `font-size: 17px` (đảm bảo giáo viên đọc rõ ràng không bị mỏi mắt).
   - Màu chữ chính (`--text-primary`): `#0f172a` (Slate-900 đậm đà).
   - Màu chữ phụ (`--text-secondary`): `#334155` (Slate-700).
   - Heading cấp 1: `2.4rem - 3.4rem`, Heading cấp 2: `1.9rem - 2.5rem`, Card title: `1.25rem`.
   - Font Code/Mã số: `'JetBrains Mono', Consolas, monospace` cỡ `0.95rem` rõ nét.
3. **Hiệu ứng chuyển cảnh (Micro-animations):**
   - Chuyển tab / lọc bài kiểm tra mượt mà với `cubic-bezier(0.16, 1, 0.3, 1)`.
   - Thẻ hover: Nâng nhẹ `transform: translateY(-4px)` cùng shadow êm ái.
   - Hạt nền (particles): Tông xanh nhạt chuyển động tinh tế, không gây rối mắt.

---

## IV. Quy chuẩn Hệ thống Câu hỏi Kiểm tra (Quiz System)

1. **Cấu trúc phân tầng 2 cấp:**
   - **Cấp 1 — Lựa chọn Văn bản pháp quy (Document Tabs):**
     - Tab 1: Công văn 3456 & TT 02/2025 (Năng lực số)
     - Tab 2: Quyết định 3439 (Năng lực AI)
     - Tab 3: Công văn 5512 & CT GDPT 2018 (NL Chung & 5 Phẩm chất)
     - Tab 4: Năng lực Đặc thù (Tin học & Robotics)
     - Tab 5: 🏆 Thi thử Toàn diện (20 câu tổng hợp ngẫu nhiên)
   - **Cấp 2 — Lựa chọn Phần nhỏ ghi nhớ (Sub-section Pills):**
     - Trong mỗi văn bản chia thành 2-4 phần nhỏ tương ứng với từng chủ đề kiến thức.
2. **Quy chuẩn nội dung câu hỏi:**
   - Mỗi câu hỏi có 4 phương án (A, B, C, D) với 1 đáp án chính xác duy nhất.
   - **Bắt buộc có phần giải thích (`explanation`):** Nêu rõ căn cứ điều khoản/chương mục từ văn bản pháp quy gốc để giáo viên hiểu bản chất.
   - Khi nộp bài kiểm tra: Hiển thị bảng đối chiếu chi tiết từng câu (câu đúng/sai, đáp án giáo viên chọn, đáp án chuẩn, và trích dẫn văn bản đối chiếu).

---

## V. Quy trình Bảo trì & Mở rộng

Khi cần bổ sung câu hỏi hoặc cập nhật văn bản mới:
1. Mở `d:\UNIGO\website-nang-luc\data.js`:
   - Nếu có văn bản mới: Khai báo đối tượng dữ liệu tương ứng trong `data.js`.
   - Nếu thêm câu hỏi: Thêm vào đúng `docKey` và `subKey` trong `QUIZ_BANK`.
2. Kiểm tra tính toàn vẹn cú pháp bằng Node/Python hoặc mở trực tiếp `index.html` trong trình duyệt.
3. Chạy kiểm tra hiển thị đảm bảo không vỡ layout trên cả Desktop lẫn Mobile.
