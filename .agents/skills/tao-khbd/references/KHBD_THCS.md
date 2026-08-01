# Luật KHBD THCS (Lớp 6–9)

> Tài liệu này được tạo dựa trên phân tích thực tế file mẫu:
> `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx`
>
> Phương án được chọn: **Phương án B** — Dùng **bảng 3 cột** cho phần Tiến trình dạy học
> (`Bước | Hoạt động của GV | Hoạt động của HS`) — dễ đọc, rõ ràng, thẩm mỹ cao hơn.

---

## A. Thông số trang và định dạng

| Thuộc tính | Giá trị |
|:---|:---|
| **Lề trái** | 2.54 cm (914400 EMU) |
| **Lề phải** | 1.27 cm (457200 EMU) |
| **Lề trên** | 1.27 cm (457200 EMU) |
| **Lề dưới** | 1.27 cm (457200 EMU) |
| **Font chữ** | Times New Roman |
| **Cỡ chữ** | 13pt |
| **Line spacing chủ đạo** | 1.15 |
| **First-line indent mục con** | 180340 EMU (~0.5cm) |
| **Alignment nội dung** | LEFT |
| **Alignment tiêu đề bài** | CENTER |

---

## B. Header & Footer — Quy tắc bảo tồn tuyệt đối

### Header (3 header references: `even`, `default`, `first`)
- **Cấu trúc**: 2 paragraphs trong header:
  - `Paragraph[0]`: chứa `w:drawing` (logo UNIGO) + text thông tin trường
  - `Paragraph[1]`: rỗng
- **TUYỆT ĐỐI không gọi `paragraph.text =` trên Paragraph[0]**
- Template có 3 loại header — tất cả được bảo tồn tự động qua sectPr

### Footer
- Footer rỗng — **KHÔNG truy cập hay chỉnh sửa footer**
- Tự động bảo tồn khi giữ `w:sectPr`

---

## C. Cấu trúc KHBD THCS — 3 bảng quan trọng

Template THCS có **3 bảng**:

### Bảng Table[0] — Thông tin trường/GV (2×2) — GIỮ NGUYÊN STRUCTURE, chỉ sửa text ngày soạn/dạy

| Cell | Nội dung gốc | Chỉnh sửa |
|:---|:---|:---|
| Row[0], Col[0] | `Trường Tiểu học và THCS UNIGO` | Giữ nguyên |
| Row[0], Col[1] | `Họ tên giáo viên: Đậu Đình Nguyên` | Giữ nguyên |
| Row[1], Col[0] | `Tổ Tin học` | Giữ nguyên |
| Row[1], Col[1] | `Ngày soạn: / /2026\nNgày dạy: / /2026` | Điền ngày soạn + ngày dạy |

### Bảng Table[2] — Ký tên BGH/Tổ CM/Người soạn (1×3) — GIỮ NGUYÊN HOÀN TOÀN

```
| BAN GIÁM HIỆU | TỔ CHUYÊN MÔN | NGƯỜI SOẠN |
```

---

## D. Cấu trúc KHBD THCS — Nội dung xây dựng

### Phần đầu

| Element | Nội dung |
|:---|:---|
| Table[0] | 2×2 thông tin trường — sửa ngày soạn/dạy |
| P Tên bài | `TÊN BÀI DẠY: [TÊN IN HOA]` (CENTER, bold) |
| P Môn/Lớp | `Môn học: ... Lớp: ... Thời lượng: ...` (CENTER, bold) |
| P Tiết PPCT | `Tiết theo PPCT: [số]` (CENTER, bold) |
| P Tên tiết | `Tên tiết: [tên bài]` (CENTER, bold) |

### Phần I — MỤC TIÊU (thứ tự bắt buộc THCS)

```
I. Mục tiêu
   1. Kiến thức: [Dùng Danh từ/Cụm danh từ: "Sự hiểu biết về...", "Khả năng nhận diện..."]
   2. Năng lực:
      - Năng lực chung: [...] (Đạt được thông qua Hoạt động X)
      - Năng lực đặc thù: [...]
      - Năng lực số (CV 3456): [...]
   3. Phẩm chất: [mô tả hành vi biểu hiện cụ thể]
```

> QUAN TRỌNG: Thứ tự bắt buộc THCS: **Kiến thức → Năng lực → Phẩm chất** (ngược với Tiểu học)

### Phần II — THIẾT BỊ DẠY HỌC

```
II. Thiết bị dạy học và học liệu:
   1. Thiết bị: [máy chiếu, máy tính GV, ...]
   2. Học liệu: [SGK, phiếu học tập, ...]
```

### Phần III — TIẾN TRÌNH DẠY HỌC (PHƯƠNG ÁN B: BẢNG 3 CỘT)

> [!IMPORTANT]
> **Phương án B**: Mỗi hoạt động gồm 4 phần (a, b, c, d) + **bảng 3 cột** cho các bước tổ chức thực hiện.

**Cấu trúc 4 Hoạt động:**

```
III. Tiến trình dạy học

1. Hoạt động 1. Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)
   a) Mục tiêu: [italic] Mục tiêu của hoạt động
   b) Nội dung: [italic] Nội dung yêu cầu/nhiệm vụ
   c) Sản phẩm: [italic] Yêu cầu sản phẩm
   d) Tổ chức thực hiện: [italic]

   +-------+--------------------+--------------------+
   | Bước  | Hoạt động của GV   | Hoạt động của HS   |
   +-------+--------------------+--------------------+
   | Bước 1: Chuyển giao nhiệm vụ học tập  | ...GV... | ...HS... |
   | Bước 2: Học sinh tiếp nhận nhiệm vụ   | ...GV... | ...HS... |
   | Bước 3: Báo cáo kết quả hoạt động     | ...GV... | ...HS... |
   | Bước 4: Đánh giá kết quả thực hiện    | ...GV... | ...HS... |
   +-------+--------------------+--------------------+

2. Hoạt động 2. Hình thành kiến thức mới/giải quyết vấn đề
   (a-d + bảng 3 cột tương tự)

3. Hoạt động 3. Luyện tập
   (a-d + bảng 3 cột, Bước 4 mặc định)

4. Hoạt động mở rộng (Nhiệm vụ về nhà)
   (a-d + bảng 3 cột, Bước 4 = "Giáo viên nhắc nhở nhiệm vụ về nhà")
```

**Quy tắc formatting bảng 3 cột:**

| Cột | Nội dung | Tỉ lệ chiều rộng |
|:---|:---|:---|
| Cột 1: `Bước` | Tên bước (italic, bold) | ~20% |
| Cột 2: `Hoạt động của GV` | Mô tả hoạt động GV | ~40% |
| Cột 3: `Hoạt động của HS` | Mô tả hoạt động HS | ~40% |

- Header row: **bold**, center, có viền đầy đủ
- Nội dung: TNR 13pt, line spacing 1.15, căn trái
- Đường viền: `single`, sz=4, color=000000

**Tên 4 Bước (Bước 1-3 giống nhau cho tất cả hoạt động, Bước 4 thay đổi ở HĐ cuối):**
- `Bước 1: Chuyển giao nhiệm vụ học tập`
- `Bước 2: Học sinh tiếp nhận nhiệm vụ học tập`
- `Bước 3: Báo cáo kết quả hoạt động`
- HĐ 1-3: `Bước 4: Đánh giá kết quả thực hiện nhiệm vụ`
- HĐ cuối: `Bước 4: Giáo viên nhắc nhở nhiệm vụ về nhà`

**Tên 4 Hoạt động (đúng theo mẫu THCS):**
- `1. Hoạt động 1. Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)`
- `2. Hoạt động 2. Hình thành kiến thức mới/giải quyết vấn đề`
- `3. Hoạt động 3. Luyện tập`
- `4. Hoạt động mở rộng (Nhiệm vụ về nhà)`

**Thời lượng chuẩn THCS (tiết 45 phút):**
- Khởi động: 7 phút
- Hình thành kiến thức: 18 phút
- Luyện tập: 12 phút
- Mở rộng/Nhiệm vụ về nhà: 8 phút

---

## E. Kỹ thuật python-docx cho THCS (Phương án B)

### 1. Luồng code chính

```python
doc = Document(TPL_SECONDARY)  # Load template THCS
clean_body(doc)  # Xóa body, giữ sectPr

# Re-create: Table[0] thông tin, paragraphs I/II/III, bảng 3 cột mỗi HĐ, Table ký tên
```

### 2. Tạo bảng 3 cột cho mỗi hoạt động

```python
def add_activity_table_b(doc, buoc_rows, is_last=False):
    """
    Tạo bảng 3 cột cho 1 hoạt động (Phương án B)
    buoc_rows = list of (gv_text, hs_text) cho 4 bước
    """
    buoc_labels = [
        'Bước 1:\nChuyển giao\nnhiệm vụ học tập',
        'Bước 2:\nHọc sinh tiếp nhận\nnhiệm vụ học tập',
        'Bước 3:\nBáo cáo kết quả\nhoạt động',
        'Bước 4:\nĐánh giá kết quả\nthực hiện nhiệm vụ',
    ]
    if is_last:
        buoc_labels[3] = 'Bước 4:\nGiáo viên nhắc nhở\nnhiệm vụ về nhà'

    table = doc.add_table(rows=1, cols=3)
    set_table_borders(table)

    # Header row
    hdr = table.rows[0].cells
    fill_cell(hdr[0], 'Bước', bold=True, align=CENTER)
    fill_cell(hdr[1], 'Hoạt động của GV', bold=True, align=CENTER)
    fill_cell(hdr[2], 'Hoạt động của HS', bold=True, align=CENTER)

    # 4 data rows
    for i, (gv, hs) in enumerate(buoc_rows):
        row = table.add_row().cells
        fill_cell(row[0], buoc_labels[i], bold=True, italic=True, align=CENTER)
        fill_cell(row[1], gv)
        fill_cell(row[2], hs)
```

### 3. Cấu trúc dữ liệu hoat_dong_list (Phương án B)

```python
hoat_dong_list = [
    {
        'stt': 1,
        'ten': 'Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)',
        'muc_tieu': 'Kích hoạt hiểu biết nền...',
        'noi_dung': 'GV đặt câu hỏi...',
        'san_pham': 'HS trả lời được...',
        'to_chuc': 'Hoạt động cá nhân nhanh...',
        # 4 bước: mỗi bước có gv_text và hs_text
        'buoc1_gv': 'GV yêu cầu HS...',
        'buoc1_hs': 'HS quan sát và...',
        'buoc2_gv': 'GV theo dõi...',
        'buoc2_hs': 'HS thực hiện...',
        'buoc3_gv': 'GV ghi lên bảng...',
        'buoc3_hs': 'HS xung phong...',
        'buoc4_gv': 'GV đặt câu hỏi dẫn vào bài...',
        'buoc4_hs': 'HS lắng nghe và chuẩn bị...',
    },
    ...
]
```

---

## F. Đặc điểm nội dung riêng cấp THCS

1. **Kiến thức**: Dùng Danh từ/Cụm danh từ — "Sự hiểu biết về...", "Khả năng nhận diện..."
2. **Năng lực số**: Tách riêng thành mục `- Năng lực số:` (không tích hợp vào NL môn học như TH)
3. **Thời lượng**: Tiết 45 phút — Khởi động 7', Hình thành KT 18', Luyện tập 12', Mở rộng 8'
4. **HĐ 3 (Luyện tập)**: Phải có thực hành trên máy tính
5. **HĐ 4**: Tên là "Hoạt động mở rộng (Nhiệm vụ về nhà)" — KHÔNG phải "Vận dụng"
6. **Bảng ký tên cuối**: Luôn có Table[2] 3 cột (BGH / Tổ CM / Người soạn)

---

## G. Checklist kiểm tra file THCS sau xuất (Phương án B)

| # | Tiêu chí | Yêu cầu |
|:---|:---|:---|
| 1 | Margins đúng? | Trái 2.54cm, Phải 1.27cm |
| 2 | Header drawing intact? | Paragraphs[0] trong header có w:drawing |
| 3 | Table[0] tồn tại? | 2×2, có ngày soạn/dạy |
| 4 | Thứ tự mục tiêu? | Kiến thức → Năng lực (chung/đặc thù/số) → Phẩm chất |
| 5 | Tiến trình dùng bảng 3 cột? | Mỗi HĐ có 1 bảng (Bước \| GV \| HS) |
| 6 | Bước 1-4 đúng thứ tự? | Chuyển giao → Tiếp nhận → Báo cáo → Đánh giá/Nhắc nhở |
| 7 | HĐ 4 đúng tên? | "Hoạt động mở rộng (Nhiệm vụ về nhà)" |
| 8 | Bảng ký tên cuối? | Table cuối 1×3 tồn tại |
| 9 | Font 13pt? | Toàn bộ runs |
| 10 | Line spacing 1.15? | Toàn bộ paragraphs |
| 11 | Tổng số bảng? | ≥ 6 (1 TT + 4 HĐ + 1 ký tên) |
