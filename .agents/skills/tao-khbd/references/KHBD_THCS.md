# Luật KHBD THCS (Lớp 6–9)

> Tài liệu này được tạo dựa trên phân tích thực tế file mẫu:
> `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx`
>
> Phương án được chọn: **Phương án A** — Giữ đúng format template gốc THCS
> (Dùng paragraphs Bước 1-4, KHÔNG dùng bảng cho Tiến trình hoạt động)

---

## A. Thông số trang và định dạng

| Thuộc tính | Giá trị |
|:---|:---|
| **Lề trái** | 2.54 cm (914400 EMU) |
| **Lề phải** | 1.27 cm (457200 EMU) |
| **Lề trên** | 1.27 cm (457200 EMU) |
| **Lề dưới** | 1.27 cm (457200 EMU) |
| **Font chữ** | Times New Roman |
| **Cỡ chữ** | 13pt (kế thừa từ style, font.size=None) |
| **Line spacing chủ đạo** | 1.15 |
| **First-line indent mục con** | 180340 EMU (~0.5cm) |
| **First-line indent bước** | 360045 EMU (~1.0cm) |
| **Alignment nội dung** | None (kế thừa, thực tế là LEFT) |
| **Alignment tiêu đề bài** | CENTER (1) |

---

## B. Header & Footer — Quy tắc bảo tồn tuyệt đối

### Header (3 header references: `even`, `default`, `first`)
- **Cấu trúc**: 2 paragraphs trong header:
  - `Paragraph[0]`: chứa `w:drawing` (logo UNIGO) + text thông tin trường
  - `Paragraph[1]`: rỗng
- **TUYỆT ĐỐI không gọi `paragraph.text =` trên Paragraph[0]**
- Template có 3 loại header (even, default, first) — tất cả đều được bảo tồn tự động qua sectPr

### Footer (3 footer references: `even`, `default`, `first`)
- Footer rỗng (chỉ có paragraph trống)
- **KHÔNG truy cập hay chỉnh sửa footer**
- Tự động bảo tồn khi giữ `w:sectPr`

### Giữ `w:sectPr`:
```python
for child in list(doc.element.body):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag != 'sectPr':
        doc.element.body.remove(child)
```

---

## C. Cấu trúc KHBD THCS — 3 bảng quan trọng

Template THCS có **3 bảng** — mỗi bảng có vai trò khác nhau:

### Bảng Table[0] — Thông tin trường/GV (2×2) — GIỮ NGUYÊN STRUCTURE, chỉ sửa text

| Cell | Nội dung gốc | Chỉnh sửa |
|:---|:---|:---|
| Row[0], Col[0] | `Trường Tiểu học và THCS UNIGO` | Giữ nguyên |
| Row[0], Col[1] | `Họ tên giáo viên: Đậu Đình Nguyên` | Sửa tên GV nếu cần |
| Row[1], Col[0] | `Tổ Tin học` | Giữ nguyên |
| Row[1], Col[1] | `Ngày soạn: / /2026\nNgày dạy: / /2026` | Điền ngày soạn + ngày dạy |

**Kỹ thuật sửa text trong cell (không xóa table)**:
```python
table0 = doc.tables[0]
# Sửa ngày soạn/dạy
cell_ngay = table0.cell(1, 1)
for para in cell_ngay.paragraphs:
    for run in para.runs:
        if 'Ngày soạn' in run.text:
            run.text = f'Ngày soạn:  {ngay_soan}   Ngày dạy:  {ngay_day}'
```

### Bảng Table[1] — HĐ của GV và HS / Kết quả cần đạt (2×2) — TÙY CHỌN dùng cho HĐ Khởi động

Template gốc có bảng 2 cột này nhưng chỉ dùng cho Hoạt động 1 (Khởi động).
Các hoạt động còn lại dùng paragraphs Bước 1-4.

### Bảng Table[2] — Ký tên BGH/Tổ CM/Người soạn (1×3) — GIỮ NGUYÊN HOÀN TOÀN

```
| BAN GIÁM HIỆU | TỔ CHUYÊN MÔN | NGƯỜI SOẠN |
```

---

## D. Cấu trúc KHBD THCS — Phần nào GIỮ NGUYÊN / CHỈNH SỬA

### Phần ĐẦU

| Element | Nội dung | Chỉnh sửa |
|:---|:---|:---|
| Table[0] | 2×2 thông tin trường | Sửa ngày soạn/dạy trong cell |
| P[0]: Tên bài | `TÊN BÀI DẠY:` (CENTER, bold) | Điền tên bài IN HOA |
| P[1]: Môn/Lớp | `Môn học/Hoạt động giáo dục: ... Lớp:` | Điền môn + lớp |
| P[2]: Thời lượng | (nếu có) | Điền số tiết |
| P[3]: Tiết PPCT | `Tiết theo PPCT:` (CENTER, bold) | Điền số tiết PPCT |
| P[4]: Tên tiết | `Tên tiết:` (CENTER, bold) | Điền tên tiết/bài |

### Phần I — MỤC TIÊU (theo thứ tự mẫu THCS)

```
I. Mục tiêu
   1. Kiến thức: Nêu cụ thể yêu cầu cần đạt về kiến thức...
   2. Năng lực:
      - Năng lực chung:
      - Năng lực đặc thù
      - Năng lực số
   3. Phẩm chất: ...
```

> QUAN TRỌNG - Thứ tự bắt buộc THCS: Kiến thức → Năng lực (chung, đặc thù, số) → Phẩm chất
> (ngược với Tiểu học)

**Quy tắc viết mục tiêu THCS:**
- `1. Kiến thức`: Dùng Danh từ/Cụm danh từ: "Sự hiểu biết về...", "Khả năng nhận diện..."
- `2. Năng lực`: Mô tả biểu hiện cụ thể, gắn mốc `(Đạt được thông qua Hoạt động X)`
- `3. Phẩm chất`: Mô tả hành vi biểu hiện cụ thể

### Phần II — THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

```
II. Thiết bị dạy học và học liệu:
   1. Thiết bị: [máy chiếu, máy tính GV, ...]
   2. Học liệu: [SGK, phiếu học tập, ...]
```

### Phần III — TIẾN TRÌNH DẠY HỌC (PARAGRAPHS - KHÔNG DÙNG BẢNG)

Đây là điểm đặc trưng quan trọng nhất của THCS theo Phương án A:

```
III. Tiến trình dạy học

1. Hoạt động 1. Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)
   a) Mục tiêu: [italic] Nêu mục tiêu giúp HS xác định vấn đề cần giải quyết
   b) Nội dung: [italic] Nêu rõ nội dung yêu cầu/nhiệm vụ cụ thể
   c) Sản phẩm: [italic] Trình bày yêu cầu về sản phẩm
   d) Tổ chức thực hiện: [italic] Các bước tổ chức hoạt động
      Bước 1. Chuyển giao nhiệm vụ học tập [italic]
      Bước 2. Học sinh tiếp nhận nhiệm vụ học tập [italic]
      Bước 3. Báo cáo kết quả hoạt động [italic]
      Bước 4. Đánh giá kết quả thực hiện nhiệm vụ [italic]

2. Hoạt động 2. Hình thành kiến thức mới/giải quyết vấn đề
   (cấu trúc a-d + Bước 1-4 tương tự)

3. Hoạt động 3. Luyện tập
   (cấu trúc a-d + Bước 1-4 tương tự)

4. Hoạt động mở rộng (Nhiệm vụ về nhà)
   a) Mục tiêu
   b) Nội dung
   c) Sản phẩm
   d) Tổ chức thực hiện
      Bước 1. Chuyển giao nhiệm vụ học tập
      Bước 2. Học sinh tiếp nhận nhiệm vụ học tập
      Bước 3. Báo cáo kết quả hoạt động
      Bước 4. Giáo viên nhắc nhở nhiệm vụ về nhà
```

**Quy tắc formatting:**
- Tiêu đề HĐ (`1. Hoạt động 1. Khởi động...`): **bold**, first-line indent 180340 EMU
- `a) b) c) d)`: italic, first-line indent 180340 EMU
- `Bước 1. Bước 2...`: italic, first-line indent 360045 EMU
- Sub-labels trong `a)` chú thích: normal (không italic)
- Line spacing: 1.15 xuyên suốt

### Phần RÚT KINH NGHIỆM SAU BÀI DẠY (GIỮ NGUYÊN)

```
RÚT KINH NGHIỆM SAU BÀI DẠY    [bold]
….……………………………………………    [bold]
Lưu ý: Sau 1 tuần mới để phần kí    [bold, CENTER]
```

### Bảng Table[2] — Ký tên (GIỮ NGUYÊN HOÀN TOÀN)

KHÔNG xóa, KHÔNG chỉnh sửa Table[2] (1×3: BGH/Tổ CM/Người soạn).
Kỹ thuật: Copy nguyên bảng Table[2] từ template gốc vào cuối body.

---

## E. Kỹ thuật python-docx cho THCS

### 1. Luồng code chính

```python
doc = Document(TPL_SECONDARY)  # Load template THCS

# Giữ Table[0] (thông tin trường) và Table[2] (ký tên)
# Xóa các element khác, giữ sectPr
body_children = list(doc.element.body)
for child in body_children:
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag not in ('sectPr',):
        doc.element.body.remove(child)

# Re-create table[0] (thông tin trường) — hoặc copy từ template gốc
# Thêm paragraphs: Tên bài, Tiết PPCT, Tên tiết
# Thêm I. Mục tiêu (Kiến thức → Năng lực → Phẩm chất)
# Thêm II. Thiết bị dạy học
# Thêm III. Tiến trình dạy học (4 HĐ dạng paragraphs)
# Thêm RÚT KINH NGHIỆM
# Copy Table[2] từ template gốc
# doc.save(output_path)
```

### 2. Cách copy Table từ template gốc

```python
import copy
doc_tpl = Document(TPL_SECONDARY)
# Lấy Table[0] (thông tin) và Table[2] (ký tên) từ template
tbl_info_xml = copy.deepcopy(doc_tpl.tables[0]._tbl)
tbl_sign_xml = copy.deepcopy(doc_tpl.tables[2]._tbl)
# Insert vào body trước sectPr
sect_pr = doc.element.body.find(qn('w:sectPr'))
doc.element.body.insert(0, tbl_info_xml)  # Table thông tin ở đầu
doc.element.body.insert(-1, tbl_sign_xml)  # Table ký tên ở cuối
```

### 3. Thêm paragraph với format đúng

```python
def add_para_thcs(doc, text, bold=False, italic=False,
                   indent_first=None, line_spacing=1.15, align=None):
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    afont(run)
    # Line spacing 1.15
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '276')  # 1.15 * 240 = 276
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    # Indent
    if indent_first:
        para.paragraph_format.first_line_indent = Emu(indent_first)
    # Align
    if align:
        para.alignment = align
    return para
```

### 4. Đường viền bảng (tương tự TH)

```python
def set_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        b.set(qn('w:space'), '0')
        tblBorders.append(b)
    tblPr.append(tblBorders)
```

---

## F. Đặc điểm nội dung riêng cấp THCS

1. **Kiến thức**: Dùng Danh từ/Cụm danh từ — "Sự hiểu biết về...", "Khả năng nhận diện..."
2. **Năng lực số**: Tách riêng thành mục `- Năng lực số:` (không tích hợp vào NL môn học như TH)
3. **Thời lượng**: Tiết 45 phút — Khởi động 7', Hình thành KT 18', Luyện tập 12', Mở rộng 8'
4. **HĐ 3 (Luyện tập)**: Phải có thực hành trên máy tính
5. **HĐ 4**: Tên là "Hoạt động mở rộng (Nhiệm vụ về nhà)" — KHÔNG phải "Vận dụng"
6. **Phụ lục**: Phiếu học tập, Rubric đánh giá — được ghi vào nội dung HĐ tương ứng (không tách ra Section riêng như SKILL.md cũ)
7. **Bảng ký tên cuối**: Luôn có Table[2] 3 cột (BGH / Tổ CM / Người soạn)

---

## G. Checklist kiểm tra file THCS sau xuất

| # | Tiêu chí | Yêu cầu |
|:---|:---|:---|
| 1 | Margins đúng? | Trái 2.54cm, Phải 1.27cm |
| 2 | Header drawing intact? | Paragraphs[0] trong header có w:drawing |
| 3 | Table[0] tồn tại? | 2×2, có ngày soạn/dạy |
| 4 | Thứ tự mục tiêu? | Kiến thức → Năng lực (chung/đặc thù/số) → Phẩm chất |
| 5 | Tiến trình dùng paragraphs? | Không có bảng 3 cột trong phần III |
| 6 | Bước 1-4 đúng thứ tự? | Chuyển giao → Tiếp nhận → Báo cáo → Đánh giá |
| 7 | HĐ 4 đúng tên? | "Hoạt động mở rộng (Nhiệm vụ về nhà)" |
| 8 | Bảng ký tên cuối? | Table[2] 1×3 tồn tại |
| 9 | Font 13pt? | Toàn bộ runs |
| 10 | Line spacing 1.15? | Toàn bộ paragraphs |
