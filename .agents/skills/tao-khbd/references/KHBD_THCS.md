# Luật KHBD THCS (Lớp 6–9)

> Tài liệu này được tạo dựa trên phân tích thực tế file mẫu:
> `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx`
>
> Format: Dùng **bảng 2 cột** cho phần Tiến trình dạy học
> (`HOẠT ĐỘNG CỦA GV – HS | KẾT QUẢ CẦN ĐẠT`) — theo mẫu giáo án mới UNIGO.

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

### Bảng Table[0] — Thông tin trường/GV (3×2) — GIỮ NGUYÊN STRUCTURE, chỉ sửa text ngày soạn/dạy/lớp

> [!IMPORTANT]
> **KHÔNG trùng lặp thông tin**: Mỗi thông tin (lớp, tên bài, môn học) chỉ xuất hiện MỘT LẦN duy nhất.

| Cell | Nội dung | Ghi chú |
|:---|:---|:---|
| Row[0], Col[0] | `Trường: Tiểu học và THCS UNIGO` | **Tên trường chuẩn** (in đậm) |
| Row[0], Col[1] | `Ngày soạn: DD/MM/YYYY` | In đậm |
| Row[1], Col[0] | `GV: Đậu Đình Nguyên` | |
| Row[1], Col[1] | `Ngày dạy: DD/MM/YYYY` | |
| Row[2], Col[0] | `Tổ: Tổ chuyên môn THCS` | |
| Row[2], Col[1] | `Lớp: [XA1]` | Chỉ ghi 1 lần, KHÔNG lặp |

### Bảng Table[2] — Ký tên BGH/Tổ CM/Người soạn (1×3) — GIỮ NGUYÊN HOÀN TOÀN

```
| BAN GIÁM HIỆU | TỔ CHUYÊN MÔN | NGƯỜI SOẠN |
```

---

## D. Cấu trúc KHBD THCS — Nội dung xây dựng

### Phần đầu (KHÔNG TRÙNG LẶP)

> [!CAUTION]
> **CẤM trùng lặp**: Tên bài chỉ xuất hiện 1 lần (`TÊN BÀI DẠY:`), KHÔNG thêm `Tên tiết:` nữa.
> Lớp chỉ xuất hiện 1 lần trong Table[0], KHÔNG thêm dòng `Lớp:` bên dưới.

| Element | Nội dung |
|:---|:---|
| Table[0] | 3×2 thông tin trường/GV/Tổ/Ngày/Lớp — sửa ngày soạn/dạy/lớp |
| P Tên bài | `TÊN BÀI DẠY: BÀI [SỐ]. [TÊN IN HOA]` (CENTER, bold) |
| P Môn học | `Môn học: Tin học` / `Môn học: Robotics` (CENTER, bold, italic) |
| P Thời lượng | `Thời lượng: 1 tiết (45 phút)` (CENTER, bold, italic) |
| P Tiết PPCT | `Tiết theo PPCT: [số]` (CENTER, bold) |

### Phần I — MỤC TIÊU (thứ tự bắt buộc THCS)

> QUAN TRỌNG: Thứ tự bắt buộc THCS: **Kiến thức → Năng lực → Phẩm chất** (ngược với Tiểu học)

> [!CAUTION]
> **CẤM LẶP NỘI DUNG**: Mỗi năng lực/phẩm chất CHỈ ĐƯỢC GHI MỘT LẦN duy nhất.
> Không ghi tổng quát (ví dụ: `- Năng lực đặc thù: Nhận biết...`) rồi lặp chi tiết bên dưới.
> Tên mục `2.1.`, `2.2.`, `2.3.` CHỈ xuất hiện 1 lần, nội dung bullet ngay phía dưới.

**Format paragraph chính xác (python-docx EMU):**

| Cấp | `first_line_indent` | `left_indent` | Bold | Ví dụ |
|-----|---------------------|---------------|------|-------|
| Tiêu đề section | 0 | 0 | ✅ | `I. Mục tiêu` |
| Mục con cấp 1 | 180340 | 0 | ✅ | `1. Kiến thức:` |
| Bullet kiến thức | 0 | 540000 | ❌ | `- Biểu diễn thông tin với kí hiệu 0 và 1.` |
| Mục con cấp 2 | 360045 | 0 | ✅ | `2.1. Năng lực đặc thù (Tin học):` |
| Bullet content | 0 | 540000 | ❌ | `- NLa (...): Biểu hiện...` |

**Spacing:** `line_spacing=1.15`, `space_after=38100` (3pt), `space_before=0`, `alignment=JUSTIFY`

**Cấu trúc mẫu chuẩn (verbatim):**

```
I. Mục tiêu                                         [bold, first_indent=0]

  1. Kiến thức:                                      [bold, first_indent=180340]
  - [Danh từ/cụm danh từ trực tiếp từ SGK].          [normal, left_indent=540000]
  - [Mỗi gạch đầu dòng xuống dòng riêng].            [normal, left_indent=540000]

  2. Năng lực:                                       [bold, first_indent=180340]
  2.1. Năng lực đặc thù (Tin học):                   [bold, first_indent=360045]
    - NLa (Sử dụng và quản lí các phương tiện ICT):  [normal, left_indent=540000]
      Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
    - NLc (Giải quyết vấn đề với sự hỗ trợ của ICT): [normal, left_indent=540000]
      Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
  2.2. Năng lực số (Thông tư 02/2025 – CV 3456):     [bold, first_indent=360045]
    - Miền I. Khai thác dữ liệu và thông tin         [normal, left_indent=540000]
      (thành tố 1.1. Duyệt, tìm kiếm và lọc dữ liệu
      – Bậc 1): Biểu hiện cụ thể. (Đạt được thông qua HĐ X)
    - Miền IV. An toàn (thành tố 4.2. Bảo vệ dữ liệu [normal, left_indent=540000]
      cá nhân – Bậc 1): Biểu hiện cụ thể.
      (Đạt được thông qua HĐ X)
  2.3. Năng lực chung:                               [bold, first_indent=360045]
    - Tự chủ và tự học: Biểu hiện cụ thể.            [normal, left_indent=540000]
      (Đạt được thông qua HĐ X, Y)
    - Giao tiếp và hợp tác: Biểu hiện cụ thể.        [normal, left_indent=540000]
      (Đạt được thông qua HĐ X, Y)
    - Giải quyết vấn đề và sáng tạo: Biểu hiện.      [normal, left_indent=540000]
      (Đạt được thông qua HĐ X)

  3. Phẩm chất:                                      [bold, first_indent=180340]
    - Chăm chỉ: Biểu hiện hành vi cụ thể.            [normal, left_indent=540000]
      (Thông qua HĐ X, Y)
    - Trung thực: Biểu hiện hành vi cụ thể.           [normal, left_indent=540000]
      (Thông qua HĐ X)
    - Trách nhiệm: Biểu hiện hành vi cụ thể.          [normal, left_indent=540000]
      (Thông qua HĐ X)
```

#### Bảng tra cứu: 5 Năng lực đặc thù Tin học (CT GDPT 2018)

| Mã | Tên thành phần | Biểu hiện ở THCS |
|----|---------------|-------------------|
| NLa | Sử dụng và quản lí các phương tiện ICT | Sử dụng đúng cách thiết bị, phần mềm, mạng; biết tổ chức lưu trữ dữ liệu; tạo sản phẩm số |
| NLb | Ứng xử phù hợp trong môi trường số | Biết quy định bản quyền, tôn trọng quyền sở hữu, ứng xử văn hóa trên mạng |
| NLc | Giải quyết vấn đề với sự hỗ trợ của ICT | Sử dụng tài nguyên thông tin, kĩ thuật ICT để giải quyết vấn đề; tư duy thuật toán |
| NLd | Ứng dụng ICT trong học và tự học | Sử dụng phần mềm để tạo sản phẩm số phục vụ học tập; khai thác phần mềm ứng dụng |
| NLe | Hợp tác trong môi trường số | Sử dụng công cụ kĩ thuật số để chia sẻ, trao đổi, hợp tác nhóm |

#### Bảng tra cứu: 6 Miền Năng lực số (Thông tư 02/2025/TT-BGDĐT + CV 3456)

| Miền | Tên đầy đủ | Các thành tố |
|------|-----------|---------------|
| I | Khai thác dữ liệu và thông tin | 1.1 Duyệt, tìm kiếm và lọc dữ liệu, thông tin và nội dung số; 1.2 Đánh giá dữ liệu, thông tin và nội dung số; 1.3 Quản lý dữ liệu, thông tin và nội dung số; 1.4 Xử lý dữ liệu |
| II | Giao tiếp và hợp tác trong môi trường số | 2.1 Tương tác qua công nghệ số; 2.2 Chia sẻ qua công nghệ số; 2.3 Tham gia với tư cách công dân số; 2.4 Hợp tác qua công nghệ số |
| III | Sáng tạo nội dung số | 3.1 Phát triển nội dung số; 3.2 Tích hợp và tái thiết kế nội dung số; 3.3 Bản quyền và giấy phép; 3.4 Lập trình |
| IV | An toàn | 4.1 Bảo vệ thiết bị và nội dung số; 4.2 Bảo vệ dữ liệu cá nhân và quyền riêng tư; 4.3 Bảo vệ sức khỏe và hạnh phúc; 4.4 Bảo vệ môi trường |
| V | Giải quyết vấn đề | 5.1 Giải quyết vấn đề kỹ thuật; 5.2 Xác định nhu cầu và giải pháp công nghệ; 5.3 Sử dụng sáng tạo công nghệ số; 5.4 Nhận diện khoảng trống năng lực số |
| VI | Ứng dụng trí tuệ nhân tạo (AI) | 6.1 Nhận biết AI trong cuộc sống; 6.2 Sử dụng công cụ AI; 6.3 Đánh giá kết quả từ AI; 6.4 Đạo đức và trách nhiệm khi sử dụng AI |

**Format bắt buộc khi ghi Năng lực số:**
```
- Miền [số La Mã]. [Tên miền đầy đủ] (thành tố [X.Y]. [Tên thành tố] – Bậc [1-8]): 
  [Biểu hiện cụ thể gắn với bài học]. (Đạt được thông qua HĐ X, Y)
```

### Phần II — THIẾT BỊ DẠY HỌC

```
II. Thiết bị dạy học và học liệu:
   1. Thiết bị: [máy chiếu, máy tính GV, ...]
   2. Học liệu: [SGK, phiếu học tập, ...]
```

### Phần III — TIẾN TRÌNH DẠY HỌC (BẢNG 2 CỘT: GV–HS | KẾT QUẢ CẦN ĐẠT)

> [!IMPORTANT]
> **Format mới**: Mỗi hoạt động gồm 4 phần (a, b, c, d) + **bảng 2 cột** (`HOẠT ĐỘNG CỦA GV – HS` | `KẾT QUẢ CẦN ĐẠT`).
> Cột 1 gộp hoạt động GV và HS theo từng bước. Cột 2 ghi kiến thức/kết quả cần đạt sau hoạt động.

**Cấu trúc 4 Hoạt động:**

```
III. Tiến trình dạy học

1. Hoạt động 1. Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)
   a) Mục tiêu: [italic] Mục tiêu của hoạt động
   b) Nội dung: [italic] Nội dung yêu cầu/nhiệm vụ
   c) Sản phẩm: [italic] Yêu cầu sản phẩm
   d) Tổ chức thực hiện: [italic]

   +----------------------------------+----------------------------------+
   | HOẠT ĐỘNG CỦA GV – HS           | KẾT QUẢ CẦN ĐẠT                |
   +----------------------------------+----------------------------------+
   | *Bước 1: Chuyển giao nhiệm vụ:* |                                  |
   | - GV yêu cầu HS...              | I. [Tiêu đề kiến thức]          |
   | - GV yêu cầu các nhóm...        | 1. [Nội dung kiến thức]         |
   | *Bước 2: HS tiếp nhận:*         | - [Chi tiết kiến thức]          |
   | - HS thực hiện...                | - [Chi tiết kiến thức]          |
   | *Bước 3: Báo cáo kết quả:*      |                                  |
   | - HS trình bày...               |                                  |
   | *Bước 4: Đánh giá kết quả:*     |                                  |
   | - GV nhận xét, chốt kiến thức   |                                  |
   +----------------------------------+----------------------------------+

2. Hoạt động 2. Hình thành kiến thức mới/giải quyết vấn đề
   (a-d + bảng 2 cột tương tự)

3. Hoạt động 3. Luyện tập
   (a-d + bảng 2 cột tương tự)

4. Hoạt động 4. Vận dụng
   (a-d + bảng 2 cột, Bước 4 = "Giáo viên nhắc nhở nhiệm vụ về nhà")
```

**Quy tắc formatting bảng 2 cột:**

| Cột | Nội dung | Tỉ lệ chiều rộng |
|:---|:---|:---|
| Cột 1: `HOẠT ĐỘNG CỦA GV – HS` | Gộp hoạt động GV & HS theo Bước 1-4. Tên bước in đậm nghiêng. | ~55% |
| Cột 2: `KẾT QUẢ CẦN ĐẠT` | Kiến thức/kết quả HS cần đạt sau hoạt động | ~45% |

- Header row: **bold**, center, có viền đầy đủ
- Nội dung: TNR 13pt, line spacing 1.15, căn trái
- Đường viền: `single`, sz=4, color=000000
- Trong cột 1: Tên bước viết in đậm nghiêng (VD: `*Bước 1: Chuyển giao nhiệm vụ:*`)
- Hoạt động GV & HS viết xen kẽ, dấu `-` đầu mỗi dòng

**Tên 4 Bước trong cột HOẠT ĐỘNG CỦA GV – HS:**
- `Bước 1: Chuyển giao nhiệm vụ học tập`
- `Bước 2: Học sinh tiếp nhận nhiệm vụ học tập`
- `Bước 3: Báo cáo kết quả hoạt động`
- HĐ 1-3: `Bước 4: Đánh giá kết quả thực hiện nhiệm vụ`
- HĐ 4: `Bước 4: Giáo viên nhắc nhở nhiệm vụ về nhà`

**Tên 4 Hoạt động (đúng theo mẫu THCS):**
- `1. Hoạt động 1. Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)`
- `2. Hoạt động 2. Hình thành kiến thức mới/giải quyết vấn đề`
- `3. Hoạt động 3. Luyện tập`
- `4. Hoạt động 4. Vận dụng`

**Thời lượng chuẩn THCS (tiết 45 phút):**
- Khởi động: 7 phút
- Hình thành kiến thức: 18 phút
- Luyện tập: 12 phút
- Vận dụng: 8 phút

---

## E. Kỹ thuật python-docx cho THCS

### 1. Luồng code chính

```python
doc = Document(TPL_SECONDARY)  # Load template THCS
clean_body(doc)  # Xóa body, giữ sectPr

# Re-create: Table[0] thông tin, paragraphs I/II/III, bảng 2 cột mỗi HĐ, Table ký tên
```

### 2. Tạo bảng 2 cột cho mỗi hoạt động

```python
def add_activity_table_2col(doc, gv_hs_text, ket_qua_text):
    """
    Tạo bảng 2 cột cho 1 hoạt động
    gv_hs_text = text gộp hoạt động GV & HS (theo Bước 1-4)
    ket_qua_text = text kết quả cần đạt
    """
    table = doc.add_table(rows=1, cols=2)
    set_table_borders(table)

    # Header row
    hdr = table.rows[0].cells
    fill_cell(hdr[0], 'HOẠT ĐỘNG CỦA GV – HS', bold=True, align=CENTER)
    fill_cell(hdr[1], 'KẾT QUẢ CẦN ĐẠT', bold=True, align=CENTER)

    # 1 data row (nội dung chi tiết)
    row = table.add_row().cells
    fill_cell_multiline(row[0], gv_hs_text)  # Bước 1-4 + hoạt động GV/HS
    fill_cell_multiline(row[1], ket_qua_text)  # Kiến thức/kết quả cần đạt
```

### 3. Cấu trúc dữ liệu hoat_dong_list

```python
hoat_dong_list = [
    {
        'stt': 1,
        'ten': 'Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)',
        'muc_tieu': 'Kích hoạt hiểu biết nền...',
        'noi_dung': 'GV đặt câu hỏi...',
        'san_pham': 'HS trả lời được...',
        'to_chuc': 'Hoạt động cá nhân nhanh...',
        # Cột 1: Gộp hoạt động GV & HS theo 4 bước
        'gv_hs': [
            ('Bước 1: Chuyển giao nhiệm vụ:', [
                '- GV yêu cầu HS quan sát hình ảnh/video...',
                '- GV đặt câu hỏi gợi mở...',
            ]),
            ('Bước 2: HS tiếp nhận:', [
                '- HS quan sát, suy nghĩ cá nhân...',
            ]),
            ('Bước 3: Báo cáo kết quả:', [
                '- HS xung phong trả lời...',
                '- GV ghi nhận các câu trả lời...',
            ]),
            ('Bước 4: Đánh giá kết quả:', [
                '- GV nhận xét, dẫn dắt vào bài mới...',
            ]),
        ],
        # Cột 2: Kết quả cần đạt
        'ket_qua': 'HS nhận diện được vấn đề bài học...',
    },
    ...
]
```

---

## F. Đặc điểm nội dung riêng cấp THCS

1. **Kiến thức**: Dùng Danh từ/Cụm danh từ trực tiếp — KHÔNG dùng động từ (hiểu, nhận diện, vận dụng). KHÔNG dùng cụm "Sự hiểu biết...", "Khả năng nhận diện...". Mỗi `-` xuống dòng riêng.
2. **Năng lực đặc thù**: Tách mục `2.1.` riêng, ghi mã NLa-NLe kèm tên đầy đủ trong ngoặc
3. **Năng lực số**: Tách mục `2.2.` riêng, ghi Miền + thành tố + Bậc theo Thông tư 02/2025
4. **Năng lực chung**: Tách mục `2.3.` riêng
5. **CẤM lặp nội dung**: Mỗi NL/PC chỉ ghi 1 lần duy nhất, KHÔNG tổng quát + chi tiết
6. **Thời lượng**: Tiết 45 phút — Khởi động 7', Hình thành KT 18', Luyện tập 12', Vận dụng 8'
7. **HĐ 3 (Luyện tập)**: Phải có thực hành trên máy tính
8. **HĐ 4**: Tên là `Hoạt động 4. Vận dụng`
9. **Bảng ký tên cuối**: Luôn có Table[2] 3 cột (BGH / Tổ CM / Người soạn)

---

## G. Checklist kiểm tra file THCS sau xuất

| # | Tiêu chí | Yêu cầu |
|:---|:---|:---|
| 1 | Margins đúng? | Trái 2.54cm, Phải 1.27cm |
| 2 | Header drawing intact? | Paragraphs[0] trong header có w:drawing |
| 3 | Tên trường đúng? | `Trường Tiểu học và THCS UNIGO` |
| 4 | Table[0] tồn tại? | 3×2, có ngày soạn/dạy/lớp — KHÔNG trùng lặp thông tin |
| 5 | Phần đầu không trùng? | Tên bài chỉ 1 lần, Lớp chỉ 1 lần |
| 6 | Thứ tự mục tiêu? | Kiến thức → Năng lực (đặc thù 2.1/số 2.2/chung 2.3) → Phẩm chất |
| 7 | Kiến thức đúng format? | Danh từ trực tiếp, KHÔNG có động từ (hiểu, nhận diện, vận dụng), mỗi `-` xuống dòng riêng |
| 8 | NL đặc thù đúng format? | Có mã NLa-NLe + tên trong ngoặc + biểu hiện + #HĐ |
| 9 | NL số đúng format? | Có `Miền [La Mã]. [Tên] (thành tố X.Y. ... – Bậc N)` |
| 10 | Không lặp nội dung? | Mỗi NL/PC chỉ xuất hiện 1 lần duy nhất |
| 11 | Indent đều đặn? | Cấp 1/2/bullet đúng EMU: 180340/360045/left_indent=540000 |
| 12 | Tiến trình dùng bảng 2 cột? | Mỗi HĐ có 1 bảng (`HOẠT ĐỘNG CỦA GV – HS` \| `KẾT QUẢ CẦN ĐẠT`) |
| 13 | Bước 1-4 đúng thứ tự? | Chuyển giao → Tiếp nhận → Báo cáo → Đánh giá/Nhắc nhở |
| 14 | HĐ 4 đúng tên? | `Hoạt động 4. Vận dụng` |
| 15 | Bảng ký tên cuối? | Table cuối 1×3 tồn tại |
| 16 | Font 13pt? | Toàn bộ runs |
| 17 | Line spacing 1.15? | Toàn bộ paragraphs |
| 18 | Spacing đúng? | space_after=38100, space_before=0, alignment=JUSTIFY |
| 19 | Tổng số bảng? | ≥ 6 (1 TT + 4 HĐ + 1 ký tên) |
| 20 | Viền bảng đúng chuẩn? | Table[0] (Thông tin) và Table ký tên cuối: **NO BORDER** (`w:val="nil"`). |
| 21 | Viền bảng hoạt động? | 4 bảng tiến trình dạy học (HĐ 1-4): **CÓ VIỀN** (`w:val="single"`, `w:sz="4"`). |
