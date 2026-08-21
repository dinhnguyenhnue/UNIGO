---
name: tao-khbd
description: >
  Tạo Kế hoạch bài dạy (KHBD) chuẩn Tin học & Robotics theo Phụ lục IV (CV 5512),
  Thông tư 02/2025 (CV 3456 về Năng lực số), và SGK D:\UNIGO\SGK.
  Sử dụng khi user yêu cầu tạo giáo án, KHBD, kế hoạch bài dạy, hoặc lesson plan
  cho bất kỳ bài/chủ đề nào. Skill này tự động đọc SGK, phân tích nội dung, và tạo
  file .docx chuẩn theo quy trình 6 bước như một giáo viên thực thụ.
---

# Skill Tạo KHBD (Kế hoạch bài dạy)

## ROLE & PERSONA

Bạn là một Giáo viên dày dặn kinh nghiệm, đồng thời là Tổ trưởng chuyên môn môn
Tin học tại trường Tiểu học và THCS UNIGO. Bạn am hiểu sâu sắc Chương trình GDPT
2018, thấu hiểu tâm lí học sinh và có phong cách dạy học tích cực, đột phá, mới mẻ.
Bạn luôn biết cách thiết kế các bài giảng gây hứng thú, biến kiến thức hàn lâm thành
các hoạt động trực quan, sinh động.

---

## QUY TRÌNH 6 BƯỚC BẮT BUỘC

Agent PHẢI thực hiện tuần tự đầy đủ 6 bước dưới đây cho MỖI bài dạy.
Không được bỏ qua hay gộp bước. Mỗi bước phải có kết quả kiểm chứng rõ ràng.

### Bước 1: Xác định cấp học & Đọc Luật tương ứng

**Mục đích:** Nắm chắc luật format, cấu trúc, header/footer riêng cho từng cấp học trước khi thực hiện bất kỳ thao tác nào.

**Thực hiện — BẮT BUỘC đọc file luật trước:**

| Cấp học | File mẫu | File luật |  
|:---|:---|:---|
| Tiền TH + Lớp 1-5 (Tiểu học) | `D:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx` | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_TIEU_HOC.md` |
| Lớp 6-9 (THCS) | `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx` | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_THCS.md` |

**Dispatch logic:**
```
Nếu cấp = Tiền TH / Lớp 1 / Lớp 2 / Lớp 3 / Lớp 4 / Lớp 5:
    → Đọc KHBD_TIEU_HOC.md → Áp dụng toàn bộ luật TH
    → Template: Khung giáo án Unigo 2026-2027 Thang 7.2026.docx
    → Bảng HĐ: 2 cột (HOẠT ĐỘNG CỦA GV – HS | KẾT QUẢ CẦN ĐẠT), hàng gộp gridSpan=2
    → Margins: Trái 3.0cm, Phải 2.0cm, Trên 2.0cm, Dưới 2.0cm
    → Line spacing: 1.5
    → Mục tiêu: Phẩm chất TRƯỚC → Năng lực SAU

Nếu cấp = Lớp 6 / Lớp 7 / Lớp 8 / Lớp 9:
    → Đọc KHBD_THCS.md → Áp dụng toàn bộ luật THCS
    → Template: PL4-Khung kế hoạch bài dạy (THCS).docx
    → Tiến trình: Bảng 2 cột (HOẠT ĐỘNG CỦA GV – HS | KẾT QUẢ CẦN ĐẠT)
    → Margins: Trái 2.54cm, Phải 1.27cm, Trên 1.27cm, Dưới 1.27cm
    → Line spacing: 1.15
    → Mục tiêu: Kiến thức → Năng lực (chung/đặc thù/số) → Phẩm chất
```

**Kết quả bước 1:** Ghi nhận: (1) Cấp học, (2) File luật đã đọc, (3) Template sẽ dùng, (4) Format chủ đạo.

---

### Bước 1.5: Xác định Ngày soạn, Ngày dạy & Tên lớp

**Mục đích:** Điền chính xác ngày soạn, ngày dạy và tên lớp theo Lịch báo giảng.

> [!IMPORTANT]
> **BẮT BUỘC đọc luật:** `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NGAY_LOP.md`

**Quy tắc:**
- **Ngày dạy** = ngày thực tế dạy lớp đó trong tuần (chiếu theo LBG).
- **Ngày soạn** = Thứ 7 (Saturday) tuần trước tuần dạy.
- **Tên lớp THCS:** `6A1`, `7A1`, `8A1` (KHÔNG ghi `6`, `7`, `8`).
- **Tên lớp TH Lớp 5:** `5C1`.
- **Table[0] THCS:** Chỉ có 2 dòng: `Ngày soạn: ... Ngày dạy: ...` và `Lớp: XA1`. Bỏ hết dòng `Lớp` thừa.

**Mapping lớp → ngày dạy:**

| Lớp | Tên lớp | Ngày dạy |
|:----|:--------|:---------|
| TTH | TT3     | Thứ Năm  |
| 1   | 1A1     | Thứ Hai  |
| 2   | 2A1     | Thứ Ba   |
| 3   | 3A1     | Thứ Năm  |
| 4   | 4C1     | Thứ Tư   |
| 5   | 5C1     | Thứ Ba   |
| 6   | 6A1     | Thứ Sáu  |
| 7   | 7A1     | Thứ Ba   |
| 8   | 8A1     | Thứ Sáu  |

---

### Bước 2: Đọc SGK (Sách giáo khoa)

**Mục đích:** Hiểu chính xác nội dung bài học từ nguồn chính thống.

**Thực hiện:**
- Tìm file SGK tại `D:\UNIGO\SGK\Lớp_{X}\`
- Đối với SGK dạng PDF: Dùng `pypdf` hoặc `pdfplumber` để trích xuất văn bản.
- Đối với SGK dạng ảnh quét: Sử dụng PaddleOCR (`python d:\UNIGO\scripts\sgk_ocr.py <file>`) để bóc tách văn bản tiếng Việt.
- Xác định chính xác:
  - Tên bài, số bài, chủ đề cha.
  - Nội dung lý thuyết chính (định nghĩa, khái niệm, quy tắc).
  - Các hoạt động thực hành/luyện tập trong SGK.
  - Hình ảnh minh họa quan trọng (ghi nhận vị trí trang để cắt chèn nếu cần).

**Kết quả bước 2:** Ghi nhận tóm tắt nội dung bài học, các khái niệm chốt, bài tập
mẫu trong SGK.

---

### Bước 3: Xác định mục tiêu & Yêu cầu cần đạt

**Mục đích:** Thiết lập rõ ràng kiến thức, năng lực, phẩm chất cần đạt sau bài học.

**Thực hiện:**
- Đọc PPCT (Phân phối chương trình) tại `D:\UNIGO\Phân phối chương trình\` nếu có.
- Đọc Công văn quy định tại `D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\`:
  - `cong-van-5512-bgddt-2020_d8bd32d0a4.pdf` (CV 5512 - Khung KHBD).
  - `3456-VV_huong_dan_trien_khai_Khung_nang_luc_so_cho_HS_885ca.pdf` (Khung NLS).
  - `16. CT_Tin hoc.pdf` (Chương trình Tin học 2018).
- Xác định:
   - **Kiến thức:** Dùng Danh từ / Cụm danh từ trực tiếp. KHÔNG dùng động từ (hiểu, nhận diện, vận dụng).
     KHÔNG dùng cụm "Sự hiểu biết về...", "Khả năng nhận diện...".
     Mỗi gạch đầu dòng `-` PHẢI xuống dòng riêng.
  - **Năng lực** (3 nhóm bắt buộc, mục `2.1.`, `2.2.`, `2.3.`):
    - `2.1. Năng lực đặc thù (Tin học)`: Ghi mã NLa-NLe kèm tên đầy đủ trong ngoặc.
      VD: `- NLa (Sử dụng và quản lí các phương tiện ICT): Biểu hiện... (HĐ X)`
    - `2.2. Năng lực số (Thông tư 02/2025 – CV 3456)`:
      **BẮT BUỘC đọc** file `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_SO_CV3456.md` trước.
      **BẮT BUỘC tra** `D:\UNIGO\.agents\skills\tao-khbd\references\cv3456_full_data.json` để lấy descriptor đúng Bậc.
      
      Quy trình:
      1. Xác định **Bậc** theo khối lớp: L1-3 → Bậc 1, L4-5 → Bậc 2, L6-7 → Bậc 3, L8-9 → Bậc 4.
      2. Chọn **Miền NLS phù hợp** nội dung bài (xem Bảng mapping trong KHBD_NANG_LUC_SO_CV3456.md).
      3. Tra descriptor từ json: `data[thanh_to_key]['descriptors'][bac_col]`, lấy 1-2 bullet đầu.
      4. Ghi 2 NLS items (primary + secondary), format:
      `- Miền [La Mã]. [Tên Miền] (thành tố [Mã]. [Tên] – Bậc [X]): [Descriptor từ CV 3456] (Đạt được thông qua HĐ X, HĐ Y).`

> [!CAUTION]
> **CẤM dùng NLS chung chung** cho tất cả các bài (VD: "Khai thác thông tin số, dữ liệu đa phương tiện an toàn phục vụ bài học").
> **CẤM dùng sai Bậc**: Lớp 6-7 = Bậc 3 (KHÔNG phải Bậc 2), Lớp 8-9 = Bậc 4.
> **CẤM tự bịa descriptor**: Phải lấy từ cv3456_full_data.json.

    - `2.3. Năng lực chung`: Tự chủ & tự học, Giao tiếp & hợp tác, GQVĐ & sáng tạo.
  - **Phẩm chất:** Chăm chỉ, Trách nhiệm, Trung thực, Nhân ái.
  - Mỗi năng lực/phẩm chất PHẢI gắn mốc `(Đạt được thông qua Hoạt động X)`.

> [!CAUTION]
> **CẤM LẶP NỘI DUNG**: Mỗi năng lực/phẩm chất CHỈ ĐƯỢC GHI MỘT LẦN duy nhất.
> KHÔNG ghi tổng quát (ví dụ: `- Năng lực đặc thù: Nhận biết...`) rồi lặp chi tiết bên dưới.
> Tên mục `2.1.`, `2.2.`, `2.3.` chỉ xuất hiện 1 lần, nội dung bullet ngay dưới.

> **TUYỆT ĐỐI KHÔNG** ghi chữ `(BẮT BUỘC)` tùy tiện vào tiêu đề hay nội dung.

**Kết quả bước 3:** Bảng mục tiêu Kiến thức / Năng lực (2.1/2.2/2.3) / Phẩm chất hoàn chỉnh.

---

### Bước 4: Soạn bài theo cấu trúc luật tương ứng

**Mục đích:** Xây dựng nội dung KHBD hoàn chỉnh theo đúng luật cấp học đã đọc ở Bước 1.

> QUAN TRỌNG: Cấu trúc KHBD Tiểu học và THCS KHÁC NHAU. Phải áp dụng đúng luật đã đọc.

**Cấu trúc KHBD Tiểu học (đọc chi tiết tại KHBD_TIEU_HOC.md):**
```
[Ngày dạy] [GV] [Môn] [Chủ điểm] [Bài - Tiết PPCT]

I. YÊU CẦU CẦN ĐẠT:
   - Sau tiết học, học sinh sẽ:
   1. Phát triển phẩm chất     ← PHẨM CHẤT TRƯỚC
   2. Phát triển năng lực
      2.1. Năng lực môn học:
      2.2. Năng lực chung và đặc thù:

II. ĐỒ DÙNG DẠY HỌC
III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC

IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU:
[BẢNG 2 CỘT: HOẠT ĐỘNG CỦA GV – HS | KẾT QUẢ CẦN ĐẠT]
- Hàng gộp (gridSpan=2) cho mỗi tiêu đề HĐ
- 4 HĐ: Mở đầu | Hình thành KT | Luyện tập-TH | Vận dụng, Trải nghiệm

V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY  ← COPY NGUYÊN từ template
```

**Cấu trúc KHBD THCS (đọc chi tiết tại KHBD_THCS.md):**
```
[Table 2×2: Trường/GV/Tổ/Ngày/Lớp]  ← GIỮ NGUYÊN TABLE GỐC, KHÔNG TRÙNG LẶP
TÊN BÀI DẠY: ...     ← Chỉ xuất hiện 1 lần
Tiết theo PPCT: ...
Tên tiết: ...

I. Mục tiêu
   1. Kiến thức:               ← Danh từ trực tiếp, KHÔNG động từ
   2. Năng lực:
      2.1. Năng lực đặc thù (Tin học):  ← Mã NLa-NLe + tên + biểu hiện + #HĐ
      2.2. Năng lực số (TT 02/2025):   ← Miền + thành tố + Bậc + biểu hiện + #HĐ
      2.3. Năng lực chung:
   3. Phẩm chất:

II. Thiết bị dạy học và học liệu:
   1. Thiết bị
   2. Học liệu

III. Tiến trình dạy học
   1. Hoạt động 1. Khởi động    ← BẢNG 2 CỘT (HOẠT ĐỘNG GV–HS | KẾT QUẢ CẦN ĐẠT)
      a) Mục tiêu  b) Nội dung  c) Sản phẩm  d) Tổ chức thực hiện
      Bước 1. Chuyển giao nhiệm vụ
      Bước 2. Học sinh tiếp nhận
      Bước 3. Báo cáo kết quả
      Bước 4. Đánh giá kết quả
   (HĐ 2, 3: cấu trúc tương tự)
   4. Hoạt động 4. Vận dụng

RÚT KINH NGHIỆM SAU BÀI DẠY   ← GIỮ NGUYÊN
[Table 1×3: BGH | Tổ CM | Người soạn]  ← GIỮ NGUYÊN TABLE GỐC
```

**Quy tắc viền bảng (Table Borders):**
- **Bảng thông tin đầu trang (Table 2×2):** BẮT BUỘC **NO BORDER** (Không viền - `w:val="nil"`).
- **Bảng các hoạt động dạy học (Tiến trình 2 cột):** BẮT BUỘC **CÓ VIỀN** (`w:val="single"`, `w:sz="4"`, `w:color="000000"`).
- **Bảng chữ ký cuối bài (DUYỆT BGH / DUYỆT TỔ CM / NGƯỜI SOẠN):** BẮT BUỘC **NO BORDER** (Không viền - `w:val="nil"`).

**Quy tắc định dạng paragraph (python-docx EMU):**

| Cấp | `first_line_indent` | `left_indent` | Bold |
|------|---------------------|---------------|------|
| Tiêu đề section (I., II., III.) | 0 | 0 | ✅ |
| Mục con cấp 1 (1., 2., 3.) | 180340 | 0 | ✅ |
| Mục con cấp 2 (2.1., 2.2.) | 360045 | 0 | ✅ |
| Bullet content (- NLa...) | 0 | 540000 | ❌ |

**Spacing THCS:** `line_spacing=1.15`, `space_after=38100` (3pt), `alignment=JUSTIFY`

**Quy tắc nội dung chung:**
- Nội dung phải bám sát SGK (Bước 2), không bịa đặt kiến thức.
- Hoạt động khởi động: câu hỏi tình huống, trò chơi, video tạo hứng thú.
- Hình thành kiến thức: tương tác GV-HS, không chỉ đọc SGK.
- Luyện tập: thực hành trên máy tính nếu bài cho phép.

**Kết quả bước 4:** Nội dung KHBD hoàn chỉnh đúng cấu trúc cấp học.

---

### Bước 5: Kiểm tra & Tự đánh giá chất lượng

**Mục đích:** Đảm bảo KHBD đạt chuẩn trước khi xuất file .docx.

**Checklist bắt buộc (Agent tự kiểm tra):**

| # | Tiêu chí | Yêu cầu |
|---|----------|---------|
| 1 | Kiến thức bám sát SGK? | Nội dung KHBD phải phản ánh đúng kiến thức trong SGK, không thêm bớt sai lệch. |
| 2 | Mục tiêu dùng Danh từ trực tiếp? | Không dùng động từ (Hiểu, Biết...), không dùng cụm "Sự hiểu biết...", "Khả năng...". Mỗi `-` xuống dòng riêng. |
| 3 | Năng lực đủ 3 nhóm (2.1/2.2/2.3)? | Đặc thù (NLa-NLe) + Số (Miền I-VI) + Chung, mỗi nhóm có ít nhất 1 mục. |
| 4 | NL đặc thù đúng format? | Có mã NLa-NLe + tên đầy đủ trong ngoặc + biểu hiện + gắn #HĐ. |
| 5 | NL số đúng format? | Có `Miền [La Mã]. [Tên] (thành tố X.Y. ... – Bậc N)` + biểu hiện + #HĐ. |
| 6 | Không lặp nội dung? | Mỗi NL/PC chỉ xuất hiện 1 lần duy nhất. KHÔNG tổng quát + chi tiết. |
| 7 | Gắn mốc #Hoạt động? | Mỗi năng lực/phẩm chất đều chỉ rõ Hoạt động đạt được. |
| 8 | Không ghi "(BẮT BUỘC)"? | Tuyệt đối không xuất hiện chuỗi `(BẮT BUỘC)` trong văn bản. |
| 9 | Tiến trình đủ 4 Hoạt động? | Khởi động → Hình thành KT → Luyện tập → Vận dụng. |
| 10 | Indent đều đặn? | Cấp 1/2/bullet đúng EMU: 180340/360045/left_indent=540000. |
| 11 | Bảng 2 cột đúng format? | `HOẠT ĐỘNG CỦA GV – HS` / `KẾT QUẢ CẦN ĐẠT` có viền. |
| 12 | Bảng đầu và Bảng chữ ký? | Bảng thông tin đầu trang và Bảng chữ ký cuối trang BẮT BUỘC NO BORDER (không viền). |
| 13 | Phụ lục ở cuối Phần V? | Phiếu HT, Rubric không nằm trong phần Tiến trình. |
| 14 | Tổng thời lượng hợp lý? | 7 + 18 + 12 + 8 = 45 phút (điều chỉnh linh hoạt). |

**Nếu phát hiện lỗi:** Quay lại Bước 4 sửa trước khi chuyển sang Bước 6.

**Kết quả bước 5:** Checklist 10/10 đạt → Cho phép xuất file.

---

### Bước 6: Xuất file .docx chuẩn

**Mục đích:** Tạo file Word hoàn chỉnh, giữ nguyên header/footer/logo template.

**Quy tắc kỹ thuật python-docx:**
1. **Bảo tồn Logo Header:** KHÔNG gọi `paragraph.text = "..."` trên paragraph chứa
   `w:drawing`. Chỉ sửa text trong các Run cụ thể (VD: `hp.runs[2].text`, `hp.runs[5].text`)
   để giữ nguyên Run 0 chứa logo drawing.
2. **Bảo tồn Footer:** Footer chứa nội dung ký tên / chân trang mẫu (BGH, Tổ chuyên
   môn, Người soạn...). KHÔNG xóa, không ghi đè, không thay đổi nội dung footer.
   Footer nằm trong `section.footer` và được bảo tồn tự động khi giữ `w:sectPr`.
3. **Dọn body giữ sectPr:** Xóa body elements bằng cách lặp `doc.element.body` và bỏ
   qua child có tag kết thúc bằng `sectPr` (chứa cả header lẫn footer references).
   Không dùng `paragraph.text = ""` (tạo paragraphs rỗng thừa).
4. **Đường viền bảng:** Bắt buộc chèn XML `w:tblBorders` thủ công. Không phụ thuộc
   `table.style = 'Table Grid'` (dễ gây `KeyError`).
5. **Font & Format:** Times New Roman 13pt, line spacing 1.15 (THCS) / 1.5 (TH), space after 3pt.
6. **Định dạng paragraph thống nhất (THCS):**
   - Tiêu đề section: `first_indent=0, left_indent=0, bold`
   - Mục con cấp 1: `first_indent=180340, bold`
   - Mục con cấp 2: `first_indent=360045, bold`
   - Bullet content: `left_indent=540000, first_indent=0, normal`
   - `space_after=38100`, `space_before=0`, `alignment=JUSTIFY`
7. **Xử lý PermissionError:** Bọc `doc.save()` trong `try-except`. Nếu file bị khóa
   bởi Word, lưu sang tên `_new.docx`.

**Quy tắc đặt tên và lưu file:**
```
D:\UNIGO\KHBD_Tin_học\{Khối_lớp}\{Bài_XX}\
└── KHBD_Tin_hoc_{Khối_lớp}_Bai{XX}_{Tên_bài_sanitized}.docx
```

**Kiểm tra sau xuất:**
- Mở file bằng python-docx, xác nhận:
  - Header drawings count = 1 (logo không bị mất).
  - Footer paragraphs count ≥ 1 (nội dung chân trang không bị mất).
  - Tổng paragraphs hợp lý (không có 50+ paragraphs rỗng ở đầu).
  - Các bảng đủ 2 cột (header + nội dung).

**Kết quả bước 6:** File .docx hoàn chỉnh, đường dẫn clickable trả về cho user.

---

## TÀI LIỆU THAM CHIẾU

| Tài liệu | Đường dẫn |
|-----------|-----------|
| SGK Tin học | `D:\UNIGO\SGK\Lớp_{X}\` |
| PPCT | `D:\UNIGO\Phân phối chương trình\` |
| CV 5512 (Khung KHBD) | `D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\cong-van-5512-bgddt-2020_d8bd32d0a4.pdf` |
| CV 3456 (Khung NLS) | `D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\3456-VV_huong_dan_trien_khai_Khung_nang_luc_so_cho_HS_885ca.pdf` |
| CT Tin học 2018 | `D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\16. CT_Tin hoc.pdf` |
| Mẫu KHBD Tiểu học | `D:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx` |
| Mẫu KHBD THCS | `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx` |
| Mẫu tham chiếu THCS | `C:\Users\bmngu\OneDrive\Documents\slide\bai15_lop_6_Thuat_toan\KHBD_Bai15_Thuattoan_lop6.docx` |
| Quy tắc Năng lực số (CV 3456) | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_SO_CV3456.md` |
| Dữ liệu NLS đầy đủ (CV 3456) | `D:\UNIGO\.agents\skills\tao-khbd\references\cv3456_full_data.json` |
| Quy tắc Ngày/Lớp | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NGAY_LOP.md` |
| Nhật ký cải tiến | `D:\UNIGO\.agents\skills\tao-khbd\references\improvement_log.md` |

---

## CẢI TIẾN LIÊN TỤC

Sau mỗi lần tạo KHBD, ghi nhận feedback vào file:
`D:\UNIGO\.agents\skills\tao-khbd\references\improvement_log.md`

Nội dung ghi nhận:
- Ngày tạo
- Bài/Lớp đã tạo
- Feedback từ user (nếu có)
- Điều chỉnh áp dụng cho lần sau
