# Luật KHBD Tiểu học (Tiền tiểu học + Lớp 1–5)

> Tài liệu này được tạo dựa trên phân tích thực tế file mẫu:
> `D:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx`

---

## A. Thông số trang và định dạng

| Thuộc tính | Giá trị |
|:---|:---|
| **Lề trái** | 3.0 cm (1080000 EMU) |
| **Lề phải** | 2.0 cm (720000 EMU) |
| **Lề trên** | 2.0 cm (720000 EMU) |
| **Lề dưới** | 2.0 cm (720000 EMU) |
| **Font chữ** | Times New Roman |
| **Cỡ chữ** | 13pt (165100 EMU) |
| **Line spacing chủ đạo** | 1.5 |
| **Line spacing phần năng lực** | 1.33 |
| **First-line indent cấp 1** | 457200 EMU (~1.27cm) |
| **First-line indent cấp 2** | 450215 EMU (~1.25cm) |
| **Alignment nội dung** | JUSTIFY (3) |
| **Alignment tiêu đề bài** | CENTER (1) |

---

## B. Header & Footer — Quy tắc bảo tồn tuyệt đối

### Header (1 section, 1 header reference: `default`)
- **Cấu trúc**: 1 paragraph chứa nhiều Runs:
  - `Run[0]`: chứa `w:drawing` (logo UNIGO) — **TUYỆT ĐỐI không chỉnh sửa, không gọi `paragraph.text =` trên paragraph này**
  - `Run[1+]`: chứa text tên trường, thông tin giáo viên
- **Cách sửa thông tin GV trong header**:
  ```python
  hp = doc.sections[0].header.paragraphs[0]
  # Chỉ sửa các run text cụ thể, KHÔNG bao giờ gọi hp.text = ...
  for r in hp.runs:
      if 'Đậu Đình Nguyên' in r.text:
          r.text = r.text.replace('Đậu Đình Nguyên', ten_gv)
  ```
- **Kiểm tra**: đếm số runs có `w:drawing` >= 1

### Footer (1 footer reference: `default`)
- Chứa số trang ("1 /3")
- **KHÔNG bao giờ truy cập hay chỉnh sửa footer**
- Tự động bảo tồn khi giữ `w:sectPr` cuối body

### Giữ `w:sectPr`:
```python
# Khi dọn body, bỏ qua child có tag kết thúc bằng 'sectPr'
for child in list(doc.element.body):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag != 'sectPr':
        doc.element.body.remove(child)
```

---

## C. Cấu trúc KHBD Tiểu học — Phần nào GIỮ NGUYÊN / CHỈNH SỬA

### Phần ĐẦU (chỉnh sửa nội dung text, GIỮ định dạng paragraph)

| Paragraph | Nội dung mẫu gốc | Chỉnh sửa |
|:---|:---|:---|
| P[0] | `Thứ…ngày …. tháng ….năm 202..` | Điền ngày dạy thực tế |
| P[1] | `Họ và tên Giáo viên: Đậu Đình Nguyên` | Giữ tên GV |
| P[2] | `KẾ HOẠCH DẠY HỌC MÔN TIN HỌC` | Đổi môn học nếu cần |
| P[3] | `CHỦ ĐIỂM: ………` | Điền tên Chủ đề/Chủ điểm |
| P[4] | `BÀI:  …………….. (Tiết: (theo PPCT) )` | Điền tên bài + số tiết PPCT |
| P[5] | *(trống)* | Giữ nguyên |

### Phần I — YÊU CẦU CẦN ĐẠT

> QUAN TRỌNG - Thứ tự bắt buộc TH: **Phẩm chất TRƯỚC → Năng lực SAU** (ngược với THCS)

> [!CAUTION]
> **CẤM LẶP NỘI DUNG**: Mỗi năng lực/phẩm chất CHỈ ĐƯỢC GHI MỘT LẦN duy nhất.
> Không ghi tổng quát rồi lặp chi tiết bên dưới.

**Format paragraph chính xác (python-docx EMU):**

| Cấp | `first_line_indent` | `left_indent` | Bold | Ví dụ |
|-----|---------------------|---------------|------|-------|
| Tiêu đề section | 0 | 0 | ✅ | `I. YÊU CẦU CẦN ĐẠT:` |
| Giới thiệu | 457200 | 0 | ❌ | `- Sau tiết học, học sinh sẽ:` |
| Mục con cấp 1 | 457200 | 0 | ✅ | `1. Phát triển phẩm chất` |
| Mục con cấp 2 | 450215 | 0 | ✅ | `2.1. Năng lực đặc thù (Tin học):` |
| Bullet content | 0 | 540000 | ❌ | `- NLa (...): Biểu hiện...` |

**Spacing:** `line_spacing=1.5` (phần NL: 1.33), `space_after=38100` (3pt), `alignment=JUSTIFY`

**Cấu trúc mẫu chuẩn (verbatim):**

```
I. YÊU CẦU CẦN ĐẠT:                                 [bold, first_indent=0]
  - Sau tiết học, học sinh sẽ:

  1. Phát triển phẩm chất                             [bold, first_indent=457200]
    - Chăm chỉ: Biểu hiện cụ thể. (HĐ X)            [normal, left_indent=540000]
    - Trách nhiệm: Biểu hiện cụ thể. (HĐ X)         [normal, left_indent=540000]

  2. Phát triển năng lực                              [bold, first_indent=457200]
  2.1. Năng lực đặc thù (Tin học):                    [bold, first_indent=450215]
    - NLa (Sử dụng và quản lí các phương tiện ICT):   [normal, left_indent=540000]
      Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
    - NLd (Ứng dụng ICT trong học và tự học):          [normal, left_indent=540000]
      Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
  2.2. Năng lực số (Thông tư 02/2025 – CV 3456):      [bold, first_indent=450215]
    - Miền I. Khai thác dữ liệu và thông tin           [normal, left_indent=540000]
      (thành tố 1.1. Duyệt, tìm kiếm... – Bậc 1):
      Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
  2.3. Năng lực chung:                                [bold, first_indent=450215]
    - Tự chủ và tự học: Biểu hiện. (HĐ X, Y)          [normal, left_indent=540000]
    - Giao tiếp và hợp tác: Biểu hiện. (HĐ X, Y)     [normal, left_indent=540000]
```

#### Bảng tra cứu: 5 Năng lực đặc thù Tin học (CT GDPT 2018)

| Mã | Tên thành phần | Biểu hiện ở Tiểu học |
|----|---------------|----------------------|
| NLa | Sử dụng và quản lí các phương tiện ICT | Nhận diện, phân biệt hình dạng và chức năng thiết bị; thao tác cơ bản với phần mềm |
| NLb | Ứng xử phù hợp trong môi trường số | Biết bảo vệ thông tin cá nhân; biết quyền sở hữu trí tuệ mức đơn giản |
| NLc | Giải quyết vấn đề với sự hỗ trợ của ICT | Tìm thông tin từ nguồn dữ liệu số; diễn đạt các bước giải quyết vấn đề |
| NLd | Ứng dụng ICT trong học và tự học | Sử dụng phần mềm trò chơi, học tập; tạo sản phẩm số đơn giản |
| NLe | Hợp tác trong môi trường số | Sử dụng công cụ kĩ thuật số để chia sẻ, trao đổi thông tin |

#### Bảng tra cứu: 6 Miền Năng lực số (Thông tư 02/2025/TT-BGDĐT + CV 3456)

| Miền | Tên đầy đủ | Các thành tố |
|------|-----------|---------------|
| I | Khai thác dữ liệu và thông tin | 1.1 Duyệt, tìm kiếm và lọc dữ liệu; 1.2 Đánh giá dữ liệu; 1.3 Quản lý dữ liệu; 1.4 Xử lý dữ liệu |
| II | Giao tiếp và hợp tác trong môi trường số | 2.1 Tương tác qua công nghệ số; 2.2 Chia sẻ qua công nghệ số; 2.3 Tham gia cộng đồng số; 2.4 Hợp tác qua công nghệ số |
| III | Sáng tạo nội dung số | 3.1 Phát triển nội dung số; 3.2 Tích hợp và tái thiết kế; 3.3 Bản quyền và giấy phép; 3.4 Lập trình |
| IV | An toàn | 4.1 Bảo vệ thiết bị; 4.2 Bảo vệ dữ liệu cá nhân; 4.3 Bảo vệ sức khỏe; 4.4 Bảo vệ môi trường |
| V | Giải quyết vấn đề | 5.1 Giải quyết vấn đề kỹ thuật; 5.2 Xác định nhu cầu công nghệ; 5.3 Sử dụng sáng tạo công nghệ; 5.4 Nhận diện khoảng trống NLS |
| VI | Ứng dụng trí tuệ nhân tạo (AI) | 6.1 Nhận biết AI; 6.2 Sử dụng công cụ AI; 6.3 Đánh giá kết quả AI; 6.4 Đạo đức AI |

**Format bắt buộc khi ghi Năng lực số:**
```
- Miền [số La Mã]. [Tên miền đầy đủ] (thành tố [X.Y]. [Tên thành tố] – Bậc [1-8]): 
  [Biểu hiện cụ thể gắn với bài học]. (Đạt được thông qua HĐ X, Y)
```


### Phần II — ĐỒ DÙNG DẠY HỌC

```
II. ĐỒ DÙNG DẠY HỌC :
   1. Giáo viên: [máy chiếu, máy tính, phiếu HT, ...]
   2. Học sinh: [máy tính, SGK, ...]
```

### Phần III — PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC *(chỉ có ở TH)*

```
III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC
   - Các phương pháp: vấn đáp, thực hành, hoạt động nhóm, giải quyết vấn đề
   - Kĩ thuật: đặt câu hỏi, trình bày 1 phút, động não, chia sẻ nhóm đôi
```

### Phần IV — CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU *(BẢNG 2 CỘT)*

Tất cả hoạt động được đặt trong **1 bảng duy nhất** với **2 cột**: `Hoạt động của GV` | `Hoạt động của HS`

**Cấu trúc bảng (10 hàng mẫu gốc):**

| Loại hàng | Cột 1 | Cột 2 |
|:---|:---|:---|
| Tiêu đề HĐ1 (gridSpan=2) | `1. Hoạt động MỞ ĐẦU (... phút)\n*Mục tiêu: ...` | *(merge)* |
| Nội dung HĐ1 | GV: câu hỏi, quan sát | HS: trả lời, tương tác |
| Tiêu đề HĐ2 (gridSpan=2) | `2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI\n2.1 [Tên HĐ (...phút)]:\n*Mục tiêu:` | *(merge)* |
| Nội dung HĐ2.1 | GV hướng dẫn... | HS thực hiện... |
| Tiêu đề HĐ2.2 (gridSpan=2) | `2.2. [Tên HĐ (...phút)]:\n*Mục tiêu:` | *(merge)* |
| Nội dung HĐ2.2 | GV hướng dẫn... | HS thực hiện... |
| Tiêu đề HĐ3 (gridSpan=2) | `3. HĐ LUYỆN TẬP-THỰC HÀNH:\n3.1 [Tên HĐ (...phút)]:\n*Mục tiêu:` | *(merge)* |
| Nội dung HĐ3.1 | GV... | HS... |
| Tiêu đề HĐ3.2 (gridSpan=2) | `3.2 [Tên HĐ] (...phút):\n*Mục tiêu:` | *(merge)* |
| Nội dung HĐ4 | `4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (...phút)` | `...............` |

**Tên 4 loại Hoạt động (đúng theo mẫu TH):**
- `1. Hoạt động MỞ ĐẦU`
- `2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI`
- `3. HĐ LUYỆN TẬP-THỰC HÀNH`
- `4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM`

**Thời lượng cho TH (tiết 35 phút):**
- Mở đầu: 5 phút
- Hình thành KT: 15 phút
- Luyện tập: 10 phút
- Vận dụng: 5 phút

### Phần V — ĐIỀU CHỈNH - BỔ SUNG

CÁC PARAGRAPH SAU ĐÂY PHẢI được sao chép nguyên xi từ template (từ P[33] đến P[35]):
- `V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY :` (bold, center)
- `(GV ghi những nội dung mà mình đã bổ sung ngoài KHBD đã lên...)` (bold, center)
- Hàng dấu chấm `........`

> [!IMPORTANT]
> **LOẠI BỎ**: Phần `* Lưu ý khi soạn kế hoạch bài dạy:` và các dòng hướng dẫn bên dưới (P[36]-P[42]) chỉ mang tính chỉ dẫn trong mẫu, **BẮT BUỘC BỎ** khi tạo file KHBD chính thức.

Kỹ thuật: Mở template, copy XML chỉ lấy P[33] đến P[35]:
```python
doc_tpl = Document(TPL_PRIMARY)
tail_paras = doc_tpl.paragraphs[33:36]  # Chỉ lấy P[33], P[34], P[35], BỎ P[36] trở đi
for p in tail_paras:
    new_p = copy.deepcopy(p._p)
    doc.element.body._body.append(new_p)  # Thêm trước sectPr
```

---

## D. Kỹ thuật python-docx cho Tiểu học

### 1. Tạo bảng 2 cột với hàng gộp (gridSpan)

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_merged_row(table, text, bold=True):
    """Thêm hàng tiêu đề gộp 2 cột"""
    row = table.add_row()
    tc0 = row.cells[0]._tc
    tc1 = row.cells[1]._tc
    # gridSpan=2 cho cell đầu
    tcPr = tc0.get_or_add_tcPr()
    gs = OxmlElement('w:gridSpan')
    gs.set(qn('w:val'), '2')
    tcPr.append(gs)
    row._tr.remove(tc1)
    # Fill text
    para = tc0.paragraphs[0]
    run = para.add_run(text)
    run.bold = bold
    afont(run)

def add_content_row(table, gv_text, hs_text):
    """Thêm hàng nội dung GV | HS"""
    row = table.add_row()
    fill_cell(row.cells[0], gv_text)
    fill_cell(row.cells[1], hs_text)
```

### 2. Đường viền bảng

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

### 3. Line spacing 1.5

```python
def set_line_spacing(para, ratio=1.5):
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(int(240 * ratio)))
    spacing.set(qn('w:lineRule'), 'auto')
```

### 4. Luồng code chính

```python
doc = Document(TPL_PRIMARY)
# 1. Dọn body (giữ sectPr)
for child in list(doc.element.body):
    tag = child.tag.split('}')[-1]
    if tag != 'sectPr':
        doc.element.body.remove(child)
# 2. Thêm phần đầu (P[0]-P[5])
# 3. Thêm I. YÊU CẦU CẦN ĐẠT
# 4. Thêm II. ĐỒ DÙNG DẠY HỌC
# 5. Thêm III. PHƯƠNG PHÁP
# 6. Thêm IV. CÁC HOẠT ĐỘNG (bảng 2 cột)
# 7. Copy phần V từ template gốc (P[33]-P[43])
# 8. Save
```

---

## E. Đặc điểm nội dung riêng cấp Tiểu học

1. **Kiến thức**: Dùng "Sau tiết học, học sinh sẽ: [động từ trực tiếp]" — không dùng Danh từ như THCS
2. **Hoạt động Mở đầu**: Trò chơi khởi động, câu hỏi tình huống đơn giản (5 phút)
3. **Hình thành KT**: Quan sát tranh/video → thảo luận → kết luận (15 phút)
4. **Luyện tập**: Thực hành trên máy có hướng dẫn từng bước (10 phút)
5. **Vận dụng**: Kết hợp bài học với cuộc sống, nhiệm vụ về nhà (5 phút)
6. **Không có**: Phụ lục, Rubric riêng — phần này chỉ có ở THCS
