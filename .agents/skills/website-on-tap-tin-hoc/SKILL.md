---
name: website-on-tap-tin-hoc
description: Quy trình chuẩn chấm điểm, bảo trì, tối ưu UI/UX, Accessibility, hiệu năng và đồng bộ dữ liệu cho Hệ thống Website Ôn tập Tin học UNIGO (`Website_on_tap/GK1/index.html`, `trang_n_t_p_tin_h_c_unigo.html`, `on-tap-tin-hoc-hk1.html`). Sử dụng BẤT CỨ KHI NÀO cần kiểm tra chất lượng (audit), chấm điểm website theo thang 100, nâng cấp giao diện, sửa lỗi JS/DOM, cập nhật mốc tuần kiểm tra định kỳ hoặc đồng bộ bản mirror mà không làm hỏng ngân hàng câu hỏi và Google Sheets Webhook.
---

# Skill Quản trị, Chấm điểm & Cải tiến Website Ôn tập Tin học UNIGO

Skill này chuẩn hóa toàn bộ quy trình kiểm định, chấm điểm theo thang **100 điểm (D1–D6)** và nâng cấp kỹ thuật cho hệ thống Website Ôn tập Tin học (Tiểu học & THCS UNIGO) tại `d:\UNIGO\Website_on_tap\GK1\`.

---

## 1. Cấu trúc hệ thống & Tài nguyên đi kèm

```text
Website_on_tap/GK1/
├── index.html                        # Cổng chính: Ôn tập, Thi GK/CK, Hướng dẫn SGK, Dashboard GV
├── trang_n_t_p_tin_h_c_unigo.html    # Bản mirror BẮT BUỘC khớp byte-for-byte với index.html
└── on-tap-tin-hoc-hk1.html           # Module Ôn tập Học kì 1 chuyên sâu (Lớp 1–8)

.agents/skills/website-on-tap-tin-hoc/
├── SKILL.md                          # Tài liệu hướng dẫn thực thi (file này)
├── references/
│   └── rubric-cham-diem.md           # Bảng chấm điểm chi tiết 100 điểm (D1–D6) & Ngưỡng đạt
└── scripts/
    └── audit_website.py              # Công cụ chấm điểm tự động, kiểm tra node --check & mirror
```

---

## 2. Nguyên tắc bất biến (Guardrails) — CẤM VI PHẠM

Trước khi sửa bất kỳ dòng HTML/CSS/JS nào, PHẢI khóa cứng và bảo toàn 100% các thành phần nghiệp vụ sau:

1. **Google Sheets Webhook & Schema đồng bộ:**
   - Giữ nguyên URL Web App `https://script.google.com/macros/s/AKfycbxdFxPYPnGL4vaW5gi7rYRh1Mz6UDrOUhlhKGInWrHhET8g90jONtXmXfS9wQ1dwe6j/exec`.
   - Giữ nguyên khóa đồng bộ và cấu trúc cột của `Trang tính1` (`index.html`) và `OnTap_HK1` (`on-tap-tin-hoc-hk1.html`).
2. **Ngân hàng câu hỏi & Dữ liệu học sinh:**
   - Trong `index.html`: Bảo toàn `QUESTION_BANK`, `OFFICIAL_EXAM_BANK_BY_GRADE`, `SGK_ALIGNMENT_ROWS`, `DEFAULT_STUDENT_ROSTERS` và mã PIN giáo viên `"2026"`.
   - Trong `on-tap-tin-hoc-hk1.html`: **KHÔNG ĐƯỢC CHỈNH SỬA TRỰC TIẾP** dòng 77 (`const D={...}` chứa 8 khối Lớp 1–8: Khối 1–2 có `54 TN + 24 TL`; Khối 3–8 có `56 TN + 24 TL`) và dòng 78 (`DEFAULT_ROSTERS`) bằng regex toàn file. Mọi chỉnh sửa HTML/CSS/JS phải thao tác cục bộ bên ngoài hai dòng dữ liệu này.
3. **Đồng bộ Mirror:**
   - Sau mỗi lần sửa `index.html`, BẮT BUỘC copy ghi đè sang `trang_n_t_p_tin_h_c_unigo.html` để giữ đồng bộ 100%.

---

## 3. Hệ thống Design Tokens chuẩn UNIGO

Mọi nâng cấp giao diện phải dùng chung bảng màu **UNIGO Ocean / Emerald / Amber** và typography tiếng Việt (tuyệt đối không dùng màu tím/hồng mặc định kiểu AI như `#3a0ca3`, `#7209b7`, `#ede9fe`, `#fce7f3`):

- **Bảng màu cốt lõi:**
  - Ocean Dark: `#0f3d7a` (`--primary-dark`) | Ocean Primary: `#1d4ed8` (`--primary`) | Sky Secondary: `#0284c7` (`--secondary`)
  - Accent Teal: `#0d9488` | Success Emerald: `#047857` | Warning Amber: `#d97706` | Danger Red: `#dc2626`
- **Typography:**
  - Chính: `'Be Vietnam Pro', system-ui, sans-serif` (Google Fonts với `display=swap`).
  - Mã lệnh / Điểm số: `'JetBrains Mono', monospace` kết hợp `font-variant-numeric: tabular-nums`.
- **Thang `z-index` kỷ luật:** `1` (nền/card) → `20` (sticky bar) → `100` (dropdown) → `200` (modal/dialog) → `300` (toast thông báo). Cấm dùng `9999` hay `5000`.
- **Mobile & Khả năng tiếp cận:**
  - Dùng `min-height: 100dvh` (cấm dùng `100vh` trên container toàn màn hình).
  - Bắt buộc có `touch-action: manipulation`, vùng chạm tối thiểu `>= 44px`, `:focus-visible`, `@media (prefers-reduced-motion: reduce)` và `.skip-link` + `<main>`.

---

## 4. Quy trình 6 bước chuẩn khi bảo trì / cải tiến Website

### Bước 1: Chấm điểm Baseline trước khi sửa
Chạy script kiểm định để có số liệu gốc D1–D6 và kiểm tra cú pháp JS hiện tại:

```powershell
python "d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\scripts\audit_website.py" --json-out "d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\references\baseline_score.json"
```

Đọc thêm [rubric-cham-diem.md](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/rubric-cham-diem.md) nếu cần đối chiếu trọng số từng tiêu chí.

### Bước 2: Chuẩn hóa Metadata, Ngữ nghĩa & Accessibility (D1)
- Kiểm tra đủ `<meta name="description">`, `<meta name="theme-color">`, SVG Favicon UNIGO, `.skip-link`, thẻ `<main>`.
- Đảm bảo **không có `<div>` có `onclick` hoặc giả nút**; thay bằng `<button type="button">` kèm `aria-label` nếu là nút icon.
- Tất cả modal phải gắn `role="dialog"` và `aria-modal="true"`; thanh thông báo trạng thái dùng `aria-live="polite"`.

### Bước 3: Nâng cấp UI/UX, Giao diện & Phản hồi phi chặn (D2–D3)
- Loại bỏ hoàn toàn `window.alert()`, `window.confirm()`, `window.prompt()` chặn luồng; dùng hệ thống `showToast()`, `showConfirmModal()`, `showPromptModal()` hoặc lỗi inline có trạng thái rõ ràng (như `.idbar.id-error`).
- Đảm bảo các nút chức năng học sinh hỗ trợ phím tắt `A–D` / `1–4` chọn đáp án, `Enter` xác nhận/qua câu (khi đã có giải thích), `Escape` đóng modal, kèm gợi ý `.shortcut-hint`.
- Duy trì đồng hồ đếm thời gian làm bài trực tiếp (`tabular-nums`) và nút chuyển giao diện Sáng/Tối ở module HK1.

### Bước 4: Tối ưu Responsive & Sạch mã nguồn (D4–D5)
- Chuyển các `style="..."` lặp lại sang class CSS chuyên biệt để giữ số lượng inline style `<= 80` ở `index.html` và `<= 10` ở `on-tap-tin-hoc-hk1.html`.
- Kiểm tra stack 1 cột cho lưới Dashboard/Card ở breakpoint mobile (`<= 820px` / `<= 640px`) và quy tắc `@media print` cho bảng điểm A4.

### Bước 5: Cập nhật nghiệp vụ, Đồng bộ Mirror & Chạy Audit nghiệm thu (D6)
- Khi thay đổi cấu hình tuần/kì thi (ví dụ `CK1`, `GK2`), kiểm tra đồng bộ cả HTML hiển thị lẫn đối tượng cấu hình JS (`DEFAULT_PERIOD_CONFIG`) và bộ lọc khối lớp (`#filter-score-grade`).
- Copy `index.html` sang bản mirror:

```powershell
Copy-Item "d:\UNIGO\Website_on_tap\GK1\index.html" "d:\UNIGO\Website_on_tap\GK1\trang_n_t_p_tin_h_c_unigo.html" -Force
```

- Chạy lại `audit_website.py` yêu cầu:
  - Điểm cả 2 file `>= 95 / 100` (mục tiêu `100 / 100`), **D6 bắt buộc `10 / 10`**.
  - `Mirror sync: True` và `node --check` không báo lỗi.

### Bước 6: Kiểm tra DOM toàn vẹn & Nghiệm thu trình duyệt
- Dùng Python `BeautifulSoup(..., "lxml")` xác nhận không trùng ID (`dup_ids == {}`), đủ số khối và số câu hỏi `const D` (54+24 với Khối 1–2; 56+24 với Khối 3–8).
- Mở trực tiếp trên trình duyệt Chrome (Desktop + Mobile viewport) để kiểm tra chế độ Sáng/Tối, làm thử Trắc nghiệm/Tự luận, bấm Dashboard giáo viên (PIN `2026`), sau đó mới thực hiện `git add` / `git commit` / `git push` khi người dùng xác nhận nghiệm thu.
