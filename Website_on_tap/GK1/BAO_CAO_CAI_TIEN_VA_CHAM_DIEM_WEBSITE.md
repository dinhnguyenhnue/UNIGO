# Báo Cáo Chấm Điểm & Cải Tiến Website Ôn Tập Tin Học UNIGO

**Phạm vi:** Hệ thống Website Ôn tập Tin học UNIGO (`d:\UNIGO\Website_on_tap\GK1\`)  
**Đối tượng áp dụng:**
- [index.html](file:///d:/UNIGO/Website_on_tap/GK1/index.html) (và bản đồng bộ [trang_n_t_p_tin_h_c_unigo.html](file:///d:/UNIGO/Website_on_tap/GK1/trang_n_t_p_tin_h_c_unigo.html))
- [on-tap-tin-hoc-hk1.html](file:///d:/UNIGO/Website_on_tap/GK1/on-tap-tin-hoc-hk1.html)

**Bộ tiêu chí chấm điểm:** Thang 100 điểm (6 chiều D1–D6) quy chuẩn tại [rubric-cham-diem.md](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/rubric-cham-diem.md), thực thi tự động bằng [audit_website.py](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/scripts/audit_website.py).

---

## 1. Tổng hợp kết quả chấm điểm qua các vòng

| Trang web | Vòng 1 (Trước cải tiến) | Vòng 2 (Sau nâng cấp UI/UX) | Vòng 3 / Cuối cùng (Tối ưu cấu trúc) | Mức tăng |
| :--- | :---: | :---: | :---: | :---: |
| `index.html` | **22 / 100** | **93 / 100** | **100 / 100** | **+78 điểm** |
| `on-tap-tin-hoc-hk1.html` | **54 / 100** | **100 / 100** | **100 / 100** | **+46 điểm** |
| **Trung bình toàn hệ thống** | **38.0 / 100** | **96.5 / 100** | **100.0 / 100** | **+62.0 điểm** |
| Đồng bộ Mirror (`index.html == trang_n_t_p_tin_h_c_unigo.html`) | — | — | **True (Khớp tuyệt đối)** | Đạt |

> [!NOTE]
> Kết quả chi tiết từng kiểm tra được lưu dạng JSON tại [phase1_baseline_score.json](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/phase1_baseline_score.json), [phase3_round2_score.json](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/phase3_round2_score.json) và [phase3_final_score.json](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/references/phase3_final_score.json).

---

## 2. Chi tiết điểm theo 6 chiều (D1–D6)

### 2.1. Trang cổng ôn tập & thi chính (`index.html`)

| Chiều đánh giá | Vòng 1 | Cuối cùng | Thang tối đa | Thay đổi nổi bật |
| :--- | :---: | :---: | :---: | :--- |
| **D1. Tiếp cận & Ngữ nghĩa (A11y / SEO)** | 2 | 20 | 20 | Bổ sung `<meta name="description">`, `<meta name="theme-color">`, SVG Favicon, `.skip-link`, `<main id="main-content">`, `role="dialog" aria-modal="true"` (4 modal), `aria-label`, `:focus-visible`, `prefers-reduced-motion`; chuyển 30 `<div>` giả nút sang `<button type="button">` / `<span>`. |
| **D2. Hệ thống thiết kế & Thẩm mỹ** | 0 | 20 | 20 | Thay thế 18 mã màu tím/hồng AI cũ sang bảng màu UNIGO Ocean/Emerald/Amber (`#0f3d7a`, `#1d4ed8`, `#0284c7`, `#0d9488`, `#047857`, `#d97706`); font `Be Vietnam Pro` + `JetBrains Mono`; `tabular-nums`; thang `z-index` chuẩn (`1/20/100/200/300`); bỏ `9999`/`5000`. |
| **D3. Luồng tương tác & Phản hồi** | 2 | 20 | 20 | Thay thế toàn bộ `17 alert()`, `4 confirm()` và `1 prompt()` bằng `showToast()`, `showConfirmModal()`, `showPromptModal()` phi chặn; bổ sung phím tắt `Escape`, `Enter`, `A–D` / `1–4` + `.shortcut-hint`. |
| **D4. Responsive & Chạm di động** | 0 | 15 | 15 | Thay `100vh` bằng `min-height: 100dvh`; thêm `touch-action: manipulation`; vùng chạm ≥ 44px; xếp chồng `.dash-grid-2` trên màn hình nhỏ (`max-width: 820px`); thêm `@media print` cho bảng điểm A4. |
| **D5. Sạch mã nguồn & Chất lượng kỹ thuật** | 8 | 15 | 15 | Giữ cú pháp JS hợp lệ (`node --check`); rút gọn `style="..."` nội tuyến từ **247 xuống 72** bằng các lớp tiện ích (`.hk1-banner`, `.dash-grid-2`, `.dash-panel-card`, `.btn-sm`, `.btn-emerald`, `.btn-teal`, `.btn-amber`, `.btn-danger`, `.btn-ocean`, `.modal-emoji`, `.modal-box-sm`). |
| **D6. Toàn vẹn dữ liệu & Nghiệp vụ UNIGO** | 10 | 10 | 10 | Bảo toàn 100% Webhook Google Sheets, `DEFAULT_STUDENT_ROSTERS`, khóa đồng bộ `OnTap_HK1`, `QUESTION_BANK`, `OFFICIAL_EXAM_BANK_BY_GRADE`, `SGK_ALIGNMENT_ROWS`, PIN `"2026"`; đồng bộ mốc **CK1 (Tuần 11 – Tuần 20)** & **GK2 (Tuần 21 – Tuần 28)** và thêm lọc **Khối 1, Khối 2** ở `#filter-score-grade`. |

### 2.2. Trang Ôn tập Học kì 1 (`on-tap-tin-hoc-hk1.html`)

| Chiều đánh giá | Vòng 1 | Cuối cùng | Thang tối đa | Thay đổi nổi bật |
| :--- | :---: | :---: | :---: | :--- |
| **D1. Tiếp cận & Ngữ nghĩa (A11y / SEO)** | 13 | 20 | 20 | Bổ sung `<meta name="description">`, `<meta name="theme-color">`, UNIGO SVG Favicon, `.skip-link`. |
| **D2. Hệ thống thiết kế & Thẩm mỹ** | 7 | 20 | 20 | Đồng bộ token UNIGO Ocean/Emerald (`#0f3d7a`, `#1d4ed8`, `#047857`), font `Be Vietnam Pro` + `JetBrains Mono`, nút chuyển giao diện Sáng/Tối `tbtn()` lưu trong `ST.theme` (`document.documentElement.dataset.theme`). |
| **D3. Luồng tương tác & Phản hồi** | 3 | 20 | 20 | Loại bỏ `2 alert()` khi thiếu tên hoặc nộp tự luận; thêm `.idbar.id-error` + `#hk1-id-error` + `showToast()`, đồng hồ đếm thời gian làm bài trực tiếp `#hk1-timer` (`startQuizTimer`/`stopQuizTimer`), phím tắt `A–D` / `1–4` & `Enter` qua câu, màn hình tổng kết Tự luận `endEssay()`, huy hiệu `.rank-pill` và nút thu gọn/mở rộng lịch sử `buildStudentHistoryCard()`. |
| **D4. Responsive & Chạm di động** | 6 | 15 | 15 | Bổ sung `min-height: 100dvh` và `touch-action: manipulation` kết hợp breakpoints chuẩn cho di động. |
| **D5. Sạch mã nguồn & Chất lượng kỹ thuật** | 15 | 15 | 15 | Cú pháp JS sạch (`node --check`), `0` inline style thừa. |
| **D6. Toàn vẹn dữ liệu & Nghiệp vụ UNIGO** | 10 | 10 | 10 | Giữ nguyên vẹn dòng 77 (`const D={...}` với 8 khối Lớp 1–8) và dòng 78 (`DEFAULT_ROSTERS`) cùng Webhook/schema `OnTap_HK1`. |

---

## 3. Kết quả kiểm định & Giới hạn kiểm thử

### 3.1. Các bước kiểm định tự động đã chạy sau chỉnh sửa cuối

1. **`audit_website.py`** (chạy lại sau khi giảm inline style ở vòng cuối):
   - `index.html`: **100/100** (`alerts=0`, `confirms=0`, `inline_styles=72`, `100vh=0`, `100dvh=1`, `dialog=4`, `aria-modal=4`, `js_syntax_ok=True`).
   - `on-tap-tin-hoc-hk1.html`: **100/100** (`alerts=0`, `confirms=0`, `inline_styles=0`, `100vh=0`, `100dvh=1`, `js_syntax_ok=True`).
   - Đồng bộ mirror (`index.html == trang_n_t_p_tin_h_c_unigo.html`): **True**.
2. **Kiểm tra cấu trúc DOM & Ngân hàng dữ liệu bằng `BeautifulSoup + lxml` & `node --check`:**
   - `index.html`: `229` ID duy nhất (`dup_ids={}`), `79` thẻ `<button>`, `1` `<main>`, `4` `role="dialog"`.
   - `on-tap-tin-hoc-hk1.html`: `3` ID duy nhất (`dup_ids={}`), `1` `<main>`, ngân hàng `const D` toàn vẹn ở cả 8 khối (**Khối 1–2**: `54` Trắc nghiệm + `24` Tự luận = `78` câu/khối; **Khối 3–8**: `56` Trắc nghiệm + `24` Tự luận = `80` câu/khối).

### 3.2. Giới hạn kiểm thử trình duyệt trực tiếp (Cần kiểm tra thủ công khi nghiệm thu)

- Công cụ trình duyệt tự động (Chrome DevTools MCP) không thể kết nối do tiến trình Chrome trên máy chưa bật cổng Remote Debugging (`DevToolsActivePort ENOENT`), và chạy `chrome.exe --headless` trong môi trường sandbox bị hệ điều hành từ chối thực thi.
- Vì vậy, hiển thị thực tế và thao tác bấm chuột/cảm ứng trên trình duyệt chưa được chạy trực tiếp trong phiên làm việc này — giáo viên vui lòng mở [index.html](file:///d:/UNIGO/Website_on_tap/GK1/index.html) và [on-tap-tin-hoc-hk1.html](file:///d:/UNIGO/Website_on_tap/GK1/on-tap-tin-hoc-hk1.html) trên Chrome để nghiệm thu trực quan.

---

## 4. Quy trình chuẩn hóa thành Skill `website-on-tap-tin-hoc`

Quy trình 6 bước chuẩn đã được đóng gói tại [SKILL.md](file:///d:/UNIGO/.agents/skills/website-on-tap-tin-hoc/SKILL.md):

1. **Bước 1 — Sao lưu & Khóa vùng dữ liệu bất biến:** Xác định các vùng không được làm hỏng (`QUESTION_BANK`, `OFFICIAL_EXAM_BANK_BY_GRADE`, `SGK_ALIGNMENT_ROWS`, dòng 77 `const D` & dòng 78 `DEFAULT_ROSTERS`, `DEFAULT_WEBHOOK_URL`, `PIN "2026"`).
2. **Bước 2 — Chạy Audit Baseline:** Chạy `python d:\UNIGO\.agents\skills\website-on-tap-tin-hoc\scripts\audit_website.py --json-out ...` để đo điểm D1–D6 trước khi sửa.
3. **Bước 3 — Chuẩn hóa UI/UX & Accessibility (D1–D4):** Áp dụng token Ocean/Emerald UNIGO, `Be Vietnam Pro`, `tabular-nums`, thay `<div>` giả nút thành `<button type="button">`, chuyển `alert`/`confirm`/`prompt` sang Toast & Modal `role="dialog"`, tối ưu `100dvh` và `touch-action: manipulation`.
4. **Bước 4 — Tối ưu mã nguồn & Đồng bộ nghiệp vụ (D5–D6):** Tách inline style sang lớp CSS tiện ích, cập nhật mốc tuần PPCT/thi, đảm bảo `node --check` không lỗi.
5. **Bước 5 — Đồng bộ Mirror & Kiểm định DOM/Rubric:** Copy `index.html` sang `trang_n_t_p_tin_h_c_unigo.html`, chạy lại `audit_website.py` đạt ≥ 95/100 (mục tiêu 100/100) và kiểm tra `BeautifulSoup` không trùng ID / giữ đủ số câu hỏi.
6. **Bước 6 — Nghiệm thu trực quan & Lưu vết báo cáo:** Kiểm tra hiển thị thực tế trên Chrome (Desktop + Mobile) và ghi biên bản/báo cáo trước khi commit/push.
