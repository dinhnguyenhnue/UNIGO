# Báo Cáo Phase 1: Chấm Điểm Hiện Trạng (Baseline) & Lộ Trình Cải Tiến Website Ôn Tập Tin Học UNIGO

- **Ngày thực hiện:** 28/11/2026
- **Trạng thái:** **Hoàn tất Phase 1 (Chưa chỉnh sửa mã nguồn 2 trang website HTML — giữ nguyên để triển khai từ Phase 2)**
- **Phạm vi đánh giá:**
  1. [index.html](file:///d:/UNIGO/Website_on_tap/GK1/index.html) *(và bản đồng bộ [trang_n_t_p_tin_h_c_unigo.html](file:///d:/UNIGO/Website_on_tap/GK1/trang_n_t_p_tin_h_c_unigo.html))* — Trang chính Hệ thống Ôn tập & Kiểm tra Tin học UNIGO (Lớp 3–8 + Trung tâm Quản trị GV).
  2. [on-tap-tin-hoc-hk1.html](file:///d:/UNIGO/Website_on_tap/GK1/on-tap-tin-hoc-hk1.html) — Chuyên trang Hành trình Ôn tập Cuối học kì 1 cùng Robot ROBO (Lớp 1–8, 80 câu/khối, đồng bộ Trang tính 2 `OnTap_HK1`).

---

## I. Các Skills Đã Áp Dụng Từ Thư Mục `D:\AI\Skills`

| STT | Skill nguồn (`D:\AI\Skills`) | Vai trò trong đánh giá & cải tiến |
|---|---|---|
| 1 | `antigravity-skills/skills/ui-ux-pro-max` | Chấm điểm ưu tiên Accessibility (tương phản 4.5:1, `:focus-visible`, ARIA), Touch Target ≥ 44px, `touch-action`, `tabular-nums`, Responsive. |
| 2 | `antigravity-skills/skills/web-design-guidelines` | Rà soát Semantic HTML, landmarks (`<main>`), `<button>` chuẩn thay thẻ `<div>` bấm được, trạng thái tương tác và nhãn trợ năng. |
| 3 | `taste-skill/skills/redesign-skill` & `taste-skill` | Phát hiện và chấm điểm các mẫu AI rập khuôn ("AI-purple gradient" `#3a0ca3` / `#7209b7` / `#ede9fe`), `100vh`, `window.alert()`, `z-index: 9999`, inline styles dư thừa. |
| 4 | `antigravity-skills/skills/frontend-design` | Định hướng lại bản sắc màu **UNIGO Ocean / Emerald** đồng bộ giữa 2 trang và typography tiếng Việt (`Be Vietnam Pro` + `JetBrains Mono`). |
| 5 | `antigravity-skills/skills/evaluation` | Xây dựng rubric 6 miền (thang 100) và công cụ đo tự động [audit_website.py](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/scripts/audit_website.py). |
| 6 | `antigravity-skills/skills/webapp-testing` | Chuẩn hóa kịch bản kiểm thử Playwright trên Chrome Desktop (`1440×900`) & Mobile (`390×844`) cho Phase 3. |
| 7 | `antigravity-skills/skills/skill-creator` | Chuẩn hóa cấu trúc đóng gói Skill tái sử dụng tại [website-on-tap-tin-hoc](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/rubric-cham-diem.md). |

---

## II. Bảng Điểm Hiện Trạng — Vòng 1 (Baseline Score)

Kết quả trích xuất trực tiếp từ công cụ [audit_website.py](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/scripts/audit_website.py) (dữ liệu JSON lưu tại [phase1_baseline_score.json](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/phase1_baseline_score.json)):

| Miền đánh giá | Điểm tối đa | `index.html` (Vòng 1) | `on-tap-tin-hoc-hk1.html` (Vòng 1) | Nhận xét chính |
|---|---:|---:|---:|---|
| **D1. Accessibility & Semantic HTML** | 20 | **2 / 20** | **13 / 20** | `index.html` còn 30 thẻ `<div>` giả nút bấm, thiếu `<main>`, thiếu `role="dialog"` / `aria-modal`, thiếu skip-link, favicon, meta SEO, `prefers-reduced-motion`. |
| **D2. Visual Identity & Anti-AI-Slop** | 20 | **0 / 20** | **7 / 20** | `index.html` dùng nặng tông tím–hồng AI mặc định (18 vị trí), `z-index` nhảy cóc (`9999`, `5000`); `hk1.html` chưa đồng bộ token màu UNIGO mới và thiếu nút chuyển Sáng/Tối. |
| **D3. Student UX & Pedagogical Feedback** | 20 | **2 / 20** | **3 / 20** | `index.html` còn **17 `alert()`** + **4 `confirm()`**, lệch tuần thi CK1 (`Tuần 11 – Tuần 19` thay vì `Tuần 20`), thiếu bộ lọc `Khối 1, Khối 2` trong Quản trị GV; `hk1.html` còn **2 `alert()`** (thiếu tên & kết thúc Tự luận), thiếu đồng hồ làm bài và phím `Enter` qua câu. |
| **D4. Responsive & Touch Ergonomics** | 15 | **0 / 15** | **6 / 15** | Còn dùng `min-height: 100vh` (nhảy thanh địa chỉ Mobile), thiếu `touch-action: manipulation`, chưa stack lưới `.dash-grid-2` trên màn hình nhỏ. |
| **D5. Code Quality & Performance** | 15 | **8 / 15** | **15 / 15** | Cú pháp JS cả 2 trang hợp lệ (`node --check` OK), nhưng `index.html` có **247 inline `style="..."`**. |
| **D6. Data Integrity & Sync Preservation** | 10 | **10 / 10** | **10 / 10** | Đầy đủ ngân hàng câu hỏi, roster các lớp `1A1–8A1`, Webhook Google Sheets & logic đồng bộ `Trang tính1` + `OnTap_HK1`. |
| **TỔNG ĐIỂM** | **100** | **22 / 100** | **54 / 100** | **Điểm trung bình toàn hệ thống: 38.0 / 100** |

---

## III. Tổng Hợp Chi Tiết Các Điểm Cần Sửa Phát Hiện Ở Phase 1

### A. Trên `index.html` (Trang chính & Quản trị GV)
1. **Tiếp cận & Ngữ nghĩa HTML (D1):**
   - Thiếu `<meta name="description">`, `<meta name="theme-color">`, SVG `<link rel="icon">`, liên kết `.skip-link`, thẻ `<main>`.
   - Còn **30 phần tử `<div>` giả nút bấm** (`#btn-top-home`, `#btn-top-student`, 5 `.step-node`, 5 `.period-card`, 4 `.activity-card` và các phần tử liên quan) chưa dùng `<button type="button">`.
   - 3 modal (`#custom-modal`, `#teacher-pin-modal`, `#dashboard-modal`) chưa có `role="dialog" aria-modal="true"`; chưa có hỗ trợ `@media (prefers-reduced-motion: reduce)`.
2. **Thẩm mỹ & Đồng bộ hệ thống (D2):**
   - Lặp 18 mã màu tím/hồng AI cũ (`#3a0ca3`, `#7209b7`, `#ede9fe`, `#fce7f3`, `#f5f3ff`), font hệ thống mặc định (`Segoe UI`) thay vì `Be Vietnam Pro` + `tabular-nums` cho đồng hồ/điểm/bảng số liệu.
   - `z-index` đặt tùy tiện (`5000`, `9999`).
3. **UX & Đồng bộ nghiệp vụ UNIGO (D3):**
   - Còn **17 `window.alert()`** và **4 `window.confirm()`** (tại dòng 3941, 4296, 4355, 4416) làm đứt mạch trải nghiệm -> cần thay bằng Toast + Confirm Modal (`Escape` để đóng).
   - Nhãn tuần Cuối kì 1 vẫn ghi `Tuần 11 – Tuần 19` (cả ở HTML dòng 1380/1384 và `DEFAULT_PERIOD_CONFIG` dòng 2062) trong khi Lịch báo giảng đã dời ĐGĐK 2 sang **Tuần 20** (`Tuần 11 – Tuần 20` và `Tuần 21 – Tuần 28`).
   - Bộ lọc `#filter-score-grade` trong Quản trị GV mới có Khối 3–8, thiếu lựa chọn `Khối 1` và `Khối 2` cho dữ liệu đồng bộ từ chuyên trang HK1.
   - Chưa có `@media print` sạch cho nút `🖨️ In Bảng Điểm A4`.
4. **Responsive & Mã nguồn (D4, D5):**
   - Dùng `min-height: 100vh`, thiếu `touch-action: manipulation`, chưa có class `.dash-grid-2` tự xếp chồng 1 cột trên Mobile, còn `247` thuộc tính `style="..."` nội tuyến.

### B. Trên `on-tap-tin-hoc-hk1.html` (Chuyên trang Hành trình HK1 — Lớp 1–8)
1. **Tiếp cận & Giao diện (D1, D2):**
   - Thiếu `<meta name="description">`, SVG favicon dùng chung với trang chính, `.skip-link`.
   - Bảng màu chưa khớp hẳn token `#0f3d7a` / `#1d4ed8` của trang chính; thiếu nút chuyển **Giao diện Sáng / Tối (`data-theme`)** trực tiếp trên thanh điều hướng (dù CSS đã có biến dark mode).
2. **UX Học sinh & Đếm giờ (D3, D4):**
   - Còn **2 `window.alert()`**: khi học sinh quên chọn tên (dòng 199) -> cần đổi sang viền đỏ `.id-error` + thông báo lỗi nội tuyến + toast; và khi hoàn thành 6 câu Tự luận (dòng 513) -> cần thay bằng màn hình tổng kết `endEssay(...)` có điểm thang 10, số lần làm và nút làm tiếp.
   - Chưa có đồng hồ đếm thời gian làm bài trực tiếp (`#hk1-timer`) trên thanh trạng thái khi làm bài.
   - Trong `showQ()`, đã hỗ trợ phím `A–D` / `1–4` nhưng chưa hỗ trợ nhấn `Enter` để chuyển câu sau khi đã khóa đáp án.
   - Thiếu `min-height: 100dvh` và `touch-action: manipulation`.

---

## IV. Cam Kết Bảo Toàn Dữ Liệu Khi Triển Khai Tiếp (Bất Biến)

Khi thực hiện Phase 2–4 trong các lượt làm việc tiếp theo, tuyệt đối **KHÔNG** thay đổi:
1. Toàn bộ ngân hàng câu hỏi (`QUESTION_BANK`, `OFFICIAL_EXAM_BANK_BY_GRADE`, `SGK_ALIGNMENT_ROWS` trong `index.html`; dòng `const D={...}` đủ 8 khối Lớp 1–8 trong `on-tap-tin-hoc-hk1.html`).
2. Toàn bộ danh sách học sinh (`DEFAULT_STUDENT_ROSTERS` / `DEFAULT_ROSTERS`) và các khóa `localStorage` (`UNIGO_EXAM_RESULTS_V1`, `UNIGO_HK1_ATTEMPTS_V1`, `ontap.tin.hk1.v1`, `UNIGO_GSHEET_WEBHOOK_URL`, mã PIN GV `2026`).
3. Webhook Google Sheets `https://script.google.com/macros/s/AKfycbxdFxPYPnGL4vaW5gi7rYRh1Mz6UDrOUhlhKGInWrHhET8g90jONtXmXfS9wQ1dwe6j/exec` và cấu trúc cột của `Trang tính1` + `OnTap_HK1`.

---

## V. Lộ Trình Lưu Sẵn Để Làm Tiếp Từ Phase 2

1. **Phase 2 — Cải tiến mã nguồn Website:**
   - Nâng cấp `index.html` (và đồng bộ `trang_n_t_p_tin_h_c_unigo.html`) theo 4 nhóm mục tiêu ở Mục III.A.
   - Nâng cấp `on-tap-tin-hoc-hk1.html` (chỉ sửa vùng CSS/head và logic JS từ dòng 78 trở đi, giữ nguyên dòng 77 `const D`) theo Mục III.B.
2. **Phase 3 — Kiểm thử Chrome thực tế & Chấm điểm Vòng 2:**
   - Kiểm tra cú pháp JS (`node --check`), đối chiếu đồng bộ bản sao `trang_n_t_p_tin_h_c_unigo.html`.
   - Chạy `audit_website.py` chấm điểm Vòng 2 (mục tiêu ≥ **92/100** trên cả 2 trang, D6 = **10/10**).
   - Chạy kiểm thử trình duyệt Chrome (Desktop `1440×900` & Mobile `390×844`) toàn bộ luồng học sinh & Quản trị GV; sửa tiếp nếu phát hiện lỗi hiển thị/tràn ngang.
3. **Phase 4 — Báo cáo hoàn thiện & Đóng gói Skill chính thức:**
   - Viết báo cáo tổng kết Trước/Sau (`BAO_CAO_CAI_TIEN_VA_CHAM_DIEM_WEBSITE.md`).
   - Hoàn thiện file `d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\SKILL.md` tích hợp sẵn rubric và công cụ `audit_website.py`.
