# Rubric Chấm Điểm & Kiểm Định Website Ôn Tập Tin Học UNIGO (Thang 100)

Bộ tiêu chí đa chiều (6 miền, 100 điểm) được tổng hợp từ thư viện skill chuẩn trong `D:\AI\Skills`:
- **UI/UX & Tiếp cận:** `ui-ux-pro-max`, `web-design-guidelines`
- **Thẩm mỹ & Anti-AI-Slop:** `redesign-skill`, `taste-skill`, `frontend-design`
- **Đo kiểm & Kiểm thử thực tế:** `evaluation`, `webapp-testing`

---

## Bảng Phân Bổ 6 Miền Đánh Giá (Tổng 100 Điểm)

| Mã | Miền đánh giá | Trọng số | Tiêu chí kiểm tra định lượng |
|----|---------------|----------|------------------------------|
| **D1** | **Accessibility & Semantic HTML** | **20** | `<meta name="description">` (2đ), SVG `<link rel="icon">` (2đ), `.skip-link` (3đ), Landmark `<main>` (3đ), dùng `<button>` ngữ nghĩa thay `<div class="badge-pill\|step-node\|period-card\|activity-card">` & ARIA `role="dialog" aria-modal="true"` / `aria-live="polite"` (6đ), `:focus-visible` (2đ), `@media (prefers-reduced-motion)` (2đ) |
| **D2** | **Visual Identity & Anti-AI-Slop** | **20** | Loại bỏ bảng màu tím AI mặc định (`#3a0ca3`, `#7209b7`, `#ede9fe`, `#fce7f3`, `#f5f3ff`) và đồng bộ hệ token **UNIGO Ocean / Emerald** (`#0f3d7a`, `#1d4ed8`, `#047857`) (5đ), Typography tiếng Việt **Be Vietnam Pro** (4đ), `font-variant-numeric: tabular-nums` cho điểm/đồng hồ/bảng số liệu (3đ), thang `z-index` chuẩn (không dùng `5000`/`9999`) (4đ), hệ thống design tokens (ở `index.html`) / nút chuyển Sáng–Tối `data-theme` (ở `on-tap-tin-hoc-hk1.html`) (4đ) |
| **D3** | **Student UX, Gamification & Pedagogical Feedback** | **20** | **0** lệnh `window.alert()` gây chặn luồng (5đ), **0** lệnh `window.confirm()` thô cứng (3đ), hệ thống thông báo Toast/Inline phi chặn (`showToast` + `toast-container`) (4đ). Trang chính: đồng bộ lịch thi CK1 `Tuần 11 – Tuần 20` & `Tuần 21 – Tuần 28` (3đ), đủ bộ lọc `Khối 1` & `Khối 2` trong Quản trị GV (2đ), tối ưu in A4 `@media print` (3đ). Chuyên trang HK1: đồng hồ đếm thời gian làm bài trực tiếp `hk1-timer` (3đ), màn hình tổng kết Tự luận `endEssay()` thay `alert()` (2đ), phím tắt `A–D`/`1–4` & `Enter` sang câu (3đ) |
| **D4** | **Responsive & Touch Ergonomics** | **15** | Dùng `min-height: 100dvh` thay `100vh` tránh nhảy viewport trên iOS/mobile (5đ), `touch-action: manipulation` loại bỏ trễ chạm 300ms (4đ), vùng chạm tối thiểu ≥ `44px` (3đ), breakpoint mobile xếp chồng grid Quản trị GV `.dash-grid-2` / responsive `@media` (3đ) |
| **D5** | **Code Quality & Performance** | **15** | Toàn bộ khối `<script>` vượt qua kiểm tra cú pháp `node --check` không lỗi (8đ), giảm thiểu `style="..."` nội tuyến chuyển sang CSS class ngữ nghĩa (`≤ 75` ở `index.html`, `≤ 10` ở `on-tap-tin-hoc-hk1.html`) (7đ) |
| **D6** | **Data Integrity & Sync Preservation** | **10** | Bảo toàn nguyên vẹn Webhook Google Sheets `AKfycbxdFxPYPnGL4vaW5gi7rYRh1Mz6UDrOUhlhKGInWrHhET8g90jONtXmXfS9wQ1dwe6j` (3đ), giữ nguyên `DEFAULT_ROSTERS`/`DEFAULT_STUDENT_ROSTERS` & khóa đồng bộ `UNIGO_EXAM_RESULTS_V1` / `OnTap_HK1` (3đ), bảo toàn 100% ngân hàng câu hỏi (`QUESTION_BANK` + `SGK_ALIGNMENT_ROWS` + PIN `2026` ở trang chính; đủ 8 khối Lớp 1–8 trong `const D` ở trang HK1) (4đ) |

---

## Ngưỡng Nghiệm Thu (Quality Gate)

- **Chưa đạt (< 70 điểm):** Còn nhiều lỗi UX gây chặn (`alert/confirm`), màu sắc rập khuôn hoặc thiếu tiếp cận/responsive.
- **Khá (70 – 84 điểm):** Hoạt động ổn định nhưng chưa đồng bộ thiết kế UNIGO hoặc còn nợ một số tiêu chí WCAG/mobile.
- **Đạt chuẩn phát hành (≥ 90 điểm toàn hệ thống, D6 bắt buộc = 10/10):** Đảm bảo vừa đẹp mắt, đồng bộ thương hiệu UNIGO, thân thiện học sinh Tiểu học–THCS trên cả máy phòng Tin lẫn điện thoại, và bảo toàn tuyệt đối dữ liệu ôn tập + đồng bộ Google Sheets.

---

## Cách Chạy Công Cụ Chấm Điểm Tự Động

```powershell
python "d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\scripts\audit_website.py" --json-out "d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\references\score_latest.json"
```
