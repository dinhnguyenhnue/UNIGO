---
name: tao-slide-bai-giang
description: >
  Tạo Slide bài giảng (.pptx) DÀNH CHO HỌC SINH nhìn và làm theo trong lớp.
  Slide KHÔNG phải bản mô tả lại giáo án — mà là công cụ trực quan giúp HS
  quan sát, thực hành, tương tác. Agent tự phân tích bài học và sinh nội dung
  phù hợp lứa tuổi. Giữ nguyên template Unigo (chân trang, đầu trang, vùng an toàn).
---

# Skill Tạo Slide Bài Giảng (Student-Facing)

## Triết lý cốt lõi

> **Slide bài giảng là thứ HỌC SINH nhìn lên màn hình và làm theo.**
> Đây KHÔNG phải bản mô tả lại Kế hoạch bài dạy cho giáo viên.

### Nguyên tắc vàng:
1. **Visual-first**: Ưu tiên hình ảnh, sơ đồ, minh họa trực quan. Chữ trên slide phải ngắn gọn, dễ hiểu.
2. **Ngôn ngữ hướng tới HS**: Dùng "Em hãy...", "Bước 1: ...", "Quan sát hình...", "Bạn nào biết...?" — KHÔNG dùng ngôn ngữ giáo viên như "HS nhận biết được...", "Mục tiêu: ...", "Năng lực cần đạt: ...".
3. **Tương tác**: Mỗi slide phải có yếu tố kích thích HS suy nghĩ hoặc hành động (câu hỏi, bài tập, thử thách).
4. **Phù hợp lứa tuổi**: Tiền TH/Lớp 1-2 dùng ngôn ngữ cực kỳ đơn giản + nhiều hình. Lớp 3-5 tăng dần. Lớp 6-8 có thể dùng thuật ngữ chuyên môn kèm giải thích.

### ❌ SAI (Slide kiểu mô tả giáo án):
```
Tiêu đề: "Mục tiêu bài học"
Nội dung: "- HS nhận biết được các bộ phận của máy tính
           - HS phân biệt được phần cứng và phần mềm
           - Phát triển năng lực số theo TT 02/2025"
```

### ✅ ĐÚNG (Slide cho HS nhìn và làm theo):
```
Tiêu đề: "Máy tính có những bộ phận nào? 🖥️"
Nội dung: [Hình ảnh máy tính với các mũi tên chỉ vào từng bộ phận]
           "Em hãy quan sát và gọi tên từng bộ phận nhé!"
```

---

## Quy trình bắt buộc

### Bước 1: Xác định thông tin bài học
- Môn, Lớp, Bài, Chủ đề
- Đọc file KHBD `.docx` tương ứng trong `KHBD_Tin_học/Lớp_{X}/Tuần_{YY}/` để hiểu mục tiêu và cấu trúc hoạt động
- **QUAN TRỌNG:** KHBD chỉ là nguồn tham khảo mục tiêu — KHÔNG copy nội dung KHBD vào slide

### Bước 2: Đọc SGK & Trích xuất hình ảnh (BẮT BUỘC)

> **QUY TẮC: MỌI slide hoạt động PHẢI có hình ảnh minh họa. TUYỆT ĐỐI KHÔNG để slide trống chỉ có emoji.**

- Mở SGK PDF từ `D:\UNIGO\SGK\Lớp_{X}\`
- Chạy script `d:\UNIGO\scripts\extract_sgk_all_images.py` để trích xuất:
  - **Full page renders** (200 DPI) → `SGK/Lớp_{X}/bai1_images/full_pages/`
  - **Individual embedded images** → `SGK/Lớp_{X}/bai1_images/`
- **Chuỗi fallback ảnh (3 tầng)**:
  1. ✅ **Ảnh riêng lẻ từ SGK** (`bai1_images/*.jpeg`) — ưu tiên cao nhất
  2. ✅ **Full page SGK** (`full_pages/*.png`) — luôn có cho Lớp 3-8
  3. ✅ **AI-generated** (`KHBD_Tin_học/{folder}/Bài_{XX}/images/`) — dùng `generate_image` tool
- **Với các lớp không có SGK** (Tiền TH, Lớp 1, Lớp 2): BẮT BUỘC tạo ảnh AI gồm: `cover.png`, `activity.png`, `practice.png`, `summary.png`
- **Dùng `slide.shapes.add_picture()`** để chèn ảnh thật vào slide, KHÔNG dùng shape rỗng + emoji thay thế

### Bước 2.5: Tạo ảnh riêng cho từng gạch đầu dòng (Per-Bullet Images) — BẮT BUỘC

> **QUY TẮC MỚI: Mỗi bullet point trong slide nội dung PHẢI có ảnh minh họa riêng.**
> Không dùng 1 ảnh chung cho toàn slide. Mỗi ý → 1 ảnh → giúp HS ghi nhớ trực quan.

#### Quy trình tạo ảnh per-bullet:

1. **Tách bullet:** Phân tích nội dung slide, tách ra từng bullet point riêng biệt.
2. **Tạo prompt AI từ nội dung bullet:**
   - Lấy chính text bullet làm prompt gốc
   - Thêm context phù hợp lứa tuổi:
     - **Tiền TH / Lớp 1-2**: `"cute kawaii cartoon illustration for young children aged 5-7, bright vivid colors, simple shapes, friendly characters, no text in image"`
     - **Lớp 3-5**: `"friendly colorful educational cartoon illustration for children aged 8-10, clear details, educational theme, no text in image"`
     - **Lớp 6-8**: `"clean modern educational infographic illustration for middle school students aged 11-14, semi-realistic style, no text in image"`
   - Ví dụ:
     - Bullet `"Màn hình — để con nhìn"` → Prompt: `"Cute kawaii cartoon: a friendly computer monitor screen with colorful display, for young children, bright colors, no text"`
     - Bullet `"Phần cứng = sờ được, nhìn thấy"` → Prompt: `"Educational infographic: hands touching computer hardware components keyboard mouse monitor, for middle school, clean modern style, no text"`
3. **Gọi `generate_image` tool** cho mỗi bullet → lưu tại `images/slide_{N}_bullet_{M}.png`
4. **Chuỗi fallback ảnh per-bullet (3 tầng):**
   - Tầng 1: ✅ Ảnh SGK trích xuất (nếu match nội dung bullet)
   - Tầng 2: ✅ AI-generated từ prompt bullet (dùng `generate_image`)
   - Tầng 3: ✅ Ảnh đại diện chung cho loại nội dung (cover/activity/practice)

#### Layout Per-Bullet Image — Phân cấp theo lứa tuổi:

**Layout A — Grid Flashcard (Tiền TH → Lớp 5):**
```
┌──────────────────────────────────────────────────┐
│ [Badge Banner]                                    │
│ Tiêu đề slide                                    │
│                                                    │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐           │
│  │  ẢNH 1  │  │  ẢNH 2  │  │  ẢNH 3  │           │
│  │(2in×2in)│  │(2in×2in)│  │(2in×2in)│           │
│  ├─────────┤  ├─────────┤  ├─────────┤           │
│  │ Text 1  │  │ Text 2  │  │ Text 3  │           │
│  └─────────┘  └─────────┘  └─────────┘           │
└──────────────────────────────────────────────────┘
```
- 2-3 bullets/slide → 1 hàng × 2-3 cột
- 4 bullets/slide → 2 hàng × 2 cột
- Nếu > 4 bullets → CHIA thành 2 slides
- Ảnh kích thước: **2in × 2in** (đủ lớn HS nhìn rõ)
- Text bên dưới ảnh: **18pt**, căn giữa

**Layout B — Horizontal Row (Lớp 6 → Lớp 8):**
```
┌──────────────────────────────────────────────────┐
│ [Badge Banner]                                    │
│ Tiêu đề slide                                    │
│                                                    │
│  ┌──────┐  ┌────────────────────────────────┐    │
│  │ẢNH 1 │  │ Text bullet 1                   │    │
│  │2in×   │  │ Mô tả chi tiết...               │    │
│  │1.5in  │  └────────────────────────────────┘    │
│  └──────┘                                         │
│  ┌──────┐  ┌────────────────────────────────┐    │
│  │ẢNH 2 │  │ Text bullet 2                   │    │
│  │2in×   │  │ Mô tả chi tiết...               │    │
│  │1.5in  │  └────────────────────────────────┘    │
│  └──────┘                                         │
└──────────────────────────────────────────────────┘
```
- Ảnh bên trái: **2in × 1.5in** + text bên phải
- Tối đa 3 rows/slide
- Nếu > 3 bullets → CHIA thành 2 slides

### Bước 3: Sinh nội dung slide bằng AI (THAY ĐỔI QUAN TRỌNG)

> **KHÔNG copy/paste từ KHBD vào slide. Agent phải TỰ PHÂN TÍCH bài học và SINH nội dung phù hợp cho học sinh.**

Quy trình sinh nội dung:
1. **Đọc KHBD** để hiểu: mục tiêu bài, các hoạt động, kiến thức trọng tâm
2. **Đọc SGK** để hiểu: nội dung chính xác HS cần học, hình ảnh, ví dụ trong sách
3. **Chuyển đổi** sang ngôn ngữ Student-Facing:
   - Mục tiêu GV → Câu hỏi dẫn dắt cho HS
   - Nội dung lý thuyết → Hình ảnh + chú thích ngắn
   - Hoạt động GV hướng dẫn → Bước thực hành cho HS (có đánh số)
   - Bài tập trong SGK → Slide luyện tập tương tác
4. **Điều chỉnh ngôn ngữ theo lứa tuổi:**
   - **Tiền TH / Lớp 1-2**: "Các con ơi, hãy nhìn xem...", "Con hãy chỉ vào...", dùng icon lớn, chữ to ≥24pt
   - **Lớp 3-5**: "Em hãy quan sát...", "Bước 1: Em mở...", "Em thử đoán xem..."
   - **Lớp 6-8**: "Hãy quan sát sơ đồ...", "Thảo luận nhóm: ...", thuật ngữ kèm giải thích

### Bước 4: Load template
- **Luôn** bắt đầu từ template: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`
- Kích thước: 20×11.2 inches (widescreen)
- **Giữ nguyên chân trang UNIGO** trên MỌI slide (thanh xanh "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO").
- **Bảo tồn slide master/layout:** KHÔNG xóa shapes ở vị trí chân trang. Khi thêm slide mới, luôn dùng layout từ template có sẵn chân trang.
- **Kiểm tra sau xuất:** Xác nhận mỗi slide có shape chân trang UNIGO.

#### Quy tắc thiết kế giao diện & Bố cục nâng cao (High-Aesthetic Design System):

1. **Bảo tồn Slide Master & Vùng An Toàn (VERIFIED từ template):**
   - **Logo UNIGO** = `Picture 7` tại L=0.17in, T=0.15in, W=0.95in, H=0.94in → kết thúc tại **Y=1.09in**
   - **Chân trang UNIGO** = `Picture 9` tại L=0.00in, T=6.43in, W=13.40in, H=1.23in → bắt đầu từ **Y=6.43in**
   - **VÙNG AN TOÀN NỘI DUNG:** Y = **1.15in → 6.35in** (chiều cao 5.20in)
   - **TUYỆT ĐỐI CẤM:**
     - Vẽ shape/rectangle/background có `top < 1.15in` (che logo)
     - Vẽ shape/rectangle/background có `top + height > 6.35in` (che chân trang)
     - Dùng `add_shape(RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)` phủ toàn bộ slide
   - **Kỹ thuật clamp:** Mọi shape phải được clamp: `actual_top = max(top, 1.15)`, `actual_bottom = min(top+height, 6.35)`

2. **KHÔNG thêm footer tự tạo:**
   - Template master ĐÃ CÓ chân trang `Picture 9` (thanh xanh + thông tin trường). KHÔNG thêm shape footer mới.
   - Chân trang tự động hiển thị trên mọi slide nhờ slide master.

3. **Z-Order & Tương phản (QUAN TRỌNG):**
   - **Background shapes phải `send_to_back`:** Khi tạo nền màu, gọi `spTree = sp.getparent(); spTree.remove(sp); spTree.insert(2, sp)` để đẩy xuống dưới cùng (sau `nvGrpSpPr` và `grpSpPr`). TUYỆT ĐỐI KHÔNG dùng `insert(0, sp)` vì sẽ làm vỡ XML schema (`spTree`).
   - **Text luôn ở trên:** Textbox phải được thêm SAU background shape để tự động nằm trên.
   - **Tương phản cao bắt buộc:**
     - Chữ trắng (`FFFFFF`) trên nền tối (primary/accent đậm)
     - Chữ tối (`1A2744`, `2E1065`...) trên nền nhạt (bg/card trắng)
     - KHÔNG dùng chữ nhạt trên nền nhạt hoặc chữ tối trên nền tối
   - **Palette phải có 3 loại text color:** `text_on_primary`, `text_on_bg`, `text_on_card`

4. **Quy chuẩn Font chữ & Cỡ chữ:**
   - Tiêu đề slide chính / Giới thiệu: **24pt - 28pt** (Bold).
   - Nội dung thường (Bullet text): **18pt - 20pt**.
   - Ký tự đầu dòng (●): **14pt**, Giãn dòng (Line spacing): **28pt**, Khoảng cách sau đoạn (Space after): **8pt**.
   - Giới hạn nội dung mỗi slide ngắn gọn (tối đa 3 - 4 dòng bullet) để đảm bảo chữ to rõ, thoáng, không dính hoặc chồng chữ.
   - **Lớp Tiền TH / 1-2**: Tăng cỡ chữ nội dung lên **22pt - 24pt**, giảm lượng chữ xuống **2 dòng/slide**.

5. **Cấu trúc Group Card & Accent Bar:**
   - **Thanh Accent Bar khít tuyệt đối:** Thanh accent màu bên mép card BẮT BUỘC phải khít hoàn toàn chiều cao ô trắng (`bar.top = card.top`, `bar.height = card.height`), không tạo margin hở trên/dưới.
   - **Bắt buộc Grouping:** Card ô trắng (`ROUNDED_RECTANGLE`) và thanh accent bar BẮT BUỘC phải được nhóm lại thành một khối duy nhất (`group_shapes` qua `<p:grpSp>`) để giáo viên dễ dàng di chuyển và căn chỉnh trong PowerPoint.

6. **Nội dung Song ngữ:**
   - Có thể mix thuật ngữ tiếng Anh nhưng BẮT BUỘC có phần mở ngoặc `()` giải thích tiếng Việt rõ ràng.
   - Ví dụ: "Hardware (Phần cứng)", "Input Device (Thiết bị vào)"

7. **Slide Tổng kết & Màu sắc:**
   - Slide tổng kết dùng nền nhạt `bg` và chèn panel màu chỉ nằm gọn trong Vùng An Toàn (Y 1.15in → 6.35in). Tuyệt đối không gọi `set_slide_bg(primary)` phủ toàn slide làm che mất logo master.
   - Áp dụng hệ thống xoay vòng 8+ bộ màu (Color Palette rotation) linh hoạt, hiện đại cho từng bài/lớp.

### Bước 5: Cấu trúc slide deck (10-14 slides) — STUDENT-FACING

| # | Slide | Mục đích (cho HS) | Nội dung mẫu | Thiết kế |
|---|-------|--------------------|--------------|----------|
| 1 | **Trang bìa** | Gây hứng thú, giới thiệu chủ đề | Hình ảnh lớn bắt mắt + Tên bài ngắn gọn | Nền màu chủ đạo, chữ trắng lớn |
| 2 | **Khởi động** | Kích thích tò mò, dẫn dắt vào bài | "Em hãy đoán xem đây là gì?" + hình ảnh bí ẩn / câu hỏi vui | Hình ảnh to + câu hỏi nổi bật |
| 3-6 | **Nội dung bài học** | Hướng dẫn HS quan sát, khám phá | **Mỗi bullet có ảnh riêng** — Layout A (Grid) cho TH, Layout B (Row) cho THCS | Per-Bullet Image Grid/Row |
| 7-8 | **Luyện tập** | HS tự làm bài tập, trả lời câu hỏi | Items xuất hiện **tuần tự theo click** + ảnh minh họa mỗi item | Animation appear per-item + ảnh |
| 9 | **Trò chơi / Thử thách** | Mini game tương tác | Ghép nối, Đúng/Sai, Sắp xếp — **BẮT BUỘC có ảnh minh họa** | Game layout + ảnh ~2in + animation |
| 10 | **Tổng kết** | Ghi nhớ điểm chính | Infographic tóm tắt 3-4 điểm chính + icon | Nền nhạt, panel trong Vùng An Toàn |
| 11 | **Cảm ơn** | Kết bài + BTVN | "Các em giỏi lắm! 🌟" + BTVN ngắn gọn | Nền chủ đạo + chữ trắng |

> **LƯU Ý QUAN TRỌNG:**
> - KHÔNG có slide "Mục tiêu bài học" liệt kê YCCD theo kiểu giáo án
> - KHÔNG có slide liệt kê "Năng lực", "Phẩm chất" 
> - Mục tiêu được lồng ghép tự nhiên vào nội dung các slide hoạt động
> - Mỗi slide tối đa 3-4 dòng text ngắn
> - **MỖI BULLET POINT phải có ẢNH RIÊNG** (per-bullet image)
> - **Slide câu hỏi/ghép nối: animation tuần tự + ảnh từng item**
> - **Slide trò chơi: BẮT BUỘC có ảnh minh họa đủ lớn (~2in) cho HS nhìn**

### Bước 6: Chân trang (footer) — TỰ ĐỘNG TỪ MASTER
Chân trang được cung cấp bởi slide master (`Picture 9`):
- Thanh xanh dương với text "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO" + thông tin liên hệ, địa chỉ
- **KHÔNG cần thêm shape footer mới.** Chỉ cần đảm bảo KHÔNG có shape nào che lên vùng Y > 6.35in.

### Bước 7: Hiệu ứng chuyển cảnh & Animation nâng cao

#### 7.1 Slide transition (Chuyển trang):
- Can thiệp XML `p:transition` với các hiệu ứng `fade`, `push`, `wipe`, `cover`, `split`.

#### 7.2 Animation tuần tự cho slide Câu hỏi / Ghép nối (BẮT BUỘC):

> **Slide practice/activity có `items` PHẢI animation từng item theo click.**
> HS cần thời gian suy nghĩ trước khi GV bấm hiện item/đáp án tiếp theo.

**Quy tắc animation per-item:**
1. Mỗi item (card text + ảnh minh họa) là 1 nhóm animation.
2. Thứ tự: Item 1 → click → Item 2 → click → Item 3...
3. Đáp án (nếu có) xuất hiện SAU câu hỏi, bằng 1 click riêng.
4. Ảnh minh họa của item xuất hiện CÙNG LÚC với text item (cùng 1 click).

**Quy tắc animation cho slide Ghép nối (items chứa `↔` hoặc `→`):**
1. Cột trái (vế A) + ảnh vế A xuất hiện trước → click → Cột phải (vế B) xuất hiện.
2. Mỗi cặp ghép nối có ảnh minh họa cho vế A (tạo bằng AI), kích thước ~2in×2in.
3. Đường nối/mũi tên giữa 2 vế xuất hiện sau khi cả 2 vế đã hiện.

**Hàm helper bắt buộc trong script:**
- `add_appear_animation(slide, shape, click_index)` — gán XML `p:timing` animation `appear` cho shape theo thứ tự click.
- `add_group_animation(slide, shapes_list, click_index)` — gán animation appear cho nhóm shapes (text + ảnh) cùng 1 click.

**XML Animation template (per-shape appear on click):**
```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <!-- Repeat for each click_index: -->
                <p:par>
                  <p:cTn id="N" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="N+1" presetID="1" presetClass="entr" presetSubtype="0" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="N+2" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="SHAPE_ID"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>
```

#### 7.3 Slide Trò chơi — Bắt buộc có hình ảnh minh họa:

> **MỌI slide trò chơi / hoạt động (`activity`) PHẢI có ảnh minh họa đủ lớn (~2in×2in) để HS nhìn rõ khi chơi.**

**Bảng loại trò chơi và ảnh cần tạo:**

| Loại trò chơi | Ảnh cần tạo | Prompt AI mẫu |
|---|---|---|
| **Đúng/Sai** | Ảnh minh họa từng câu hỏi + biểu tượng ✅❌ | `"[Nội dung câu hỏi], educational cartoon, clear yes/no visual"` |
| **Ghép nối** | Ảnh từng đối tượng vế A (VD: ảnh màn hình, bàn phím...) | `"[Đối tượng cần ghép], isolated on white background, cartoon style, clear"` |
| **Sắp xếp thứ tự** | Ảnh timeline / dòng thời gian minh họa | `"Timeline infographic showing [nội dung], colorful, educational"` |
| **Vẽ/Sáng tạo** | Ảnh mẫu hoặc ảnh HS đang thực hiện | `"[Hoạt động], children in classroom, bright illustration"` |
| **Thảo luận nhóm** | Ảnh nhóm HS đang thảo luận + chủ đề | `"Group of students discussing [topic] in classroom, cartoon"` |
| **Chỉ/Nhận diện** | Ảnh đối tượng cần nhận diện với mũi tên/label | `"[Đối tượng] with labeled arrows pointing to parts, educational diagram"` |

**Layout slide trò chơi:**
```
┌──────────────────────────────────────────────────┐
│ [Accent Badge: LUYỆN TẬP / TRÒ CHƠI]             │
│ Tiêu đề: "Cùng chơi nào!" (26pt Bold)            │
│ Hướng dẫn ngắn (17pt)                             │
│                                                    │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐        │
│  │  ẢNH 1   │  │  ẢNH 2   │  │  ẢNH 3   │        │
│  │ (2in×2in)│  │ (2in×2in)│  │ (2in×2in)│        │
│  ├──────────┤  ├──────────┤  ├──────────┤        │
│  │[Card chọn│  │[Card chọn│  │[Card chọn│        │
│  │ Item 1]  │  │ Item 2]  │  │ Item 3]  │        │
│  └──────────┘  └──────────┘  └──────────┘        │
│              ← Animation: hiện tuần tự →          │
└──────────────────────────────────────────────────┘
```

### Bước 8: Lưu file
**Vị trí lưu tập trung theo bài học:**
```
D:\UNIGO\KHBD_Tin_học\Lớp_{X}\Tuần_{YY}\
├── Slide_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.pptx
├── KHBD_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.docx
└── images\
```

Ví dụ: `D:\UNIGO\KHBD_Tin_học\Lớp_3\Tuần_02\Slide_Tin_hoc_Lop_3_Bai01_Thong_tin_va_quyet_dinh.pptx`

---

## Anti-Bug Checklist & Kiểm tra tự động (BẮT BUỘC)

> **Chạy kiểm tra SAU MỖI LẦN tạo slide. Nếu có FAIL → sửa và chạy lại.**

### Bảng quy tắc kỹ thuật cứng (Hard Rules):

| # | Quy tắc | Kiểm tra tự động | Hậu quả nếu vi phạm |
|---|---------|-------------------|----------------------|
| 1 | KHÔNG vẽ shape có `top < 1.15in` | `assert shape.top >= Inches(1.15)` | Logo bị che |
| 2 | KHÔNG vẽ shape có `top + height > 6.35in` | `assert shape.top + shape.height <= Inches(6.35)` | Chân trang bị che |
| 3 | KHÔNG dùng `add_shape(RECT, 0, 0, SLIDE_W, SLIDE_H)` | grep script | File corrupt / Repair |
| 4 | KHÔNG thêm shape footer mới | Đếm shapes ở Y > 6.35 = chỉ master | Footer chồng |
| 5 | Z-order: `insert(2, sp)` KHÔNG BAO GIỜ `insert(0, sp)` | grep `insert(0` | XML schema vỡ |
| 6 | 3 loại text color trong palette | Check keys | Chữ bị mất tương phản |
| 7 | Font tiêu đề 24-28pt, nội dung 18-20pt | Verify font sizes | Chữ quá to/nhỏ |
| 8 | Tối đa 3-4 bullets/slide | Count lines ≤ 4 | Chữ tràn/chồng |
| 9 | Mỗi slide nội dung có ≥ 1 hình ảnh | Count pictures ≥ 1 | Slide trống |
| 10 | Mỗi bullet có ảnh riêng (per-bullet) | Count pictures ≥ bullets | Thiếu minh họa |
| 11 | Slide practice/activity có animation | Check `p:timing` XML | Items hiện cùng lúc |
| 12 | Slide trò chơi có ảnh minh họa ≥ 1 | Count pictures on game slides | HS không có gì nhìn |
| 13 | Card + accent bar phải group | Check `<p:grpSp>` | Bố cục rời rạc |

### Script kiểm tra `_verify_slide_v2.py`:
- Chạy tự động sau mỗi lần tạo slide
- Input: file `.pptx` vừa tạo
- Output: Báo cáo PASS/FAIL cho từng slide + từng quy tắc
- Nếu bất kỳ quy tắc 1-5 FAIL → BẮT BUỘC sửa và tạo lại

---

## Checklist kiểm tra trước khi giao slide

### Kỹ thuật:
- [ ] Mọi shape nằm trong Vùng An Toàn (Y 1.15in → 6.35in)
- [ ] Logo UNIGO không bị che (không shape nào ở Y < 1.15in)
- [ ] Chân trang UNIGO có trên mọi slide (master Picture 9)
- [ ] Z-order đúng: background `insert(2, sp)`, text thêm sau
- [ ] Không có shape footer tự tạo
- [ ] Font chữ đúng chuẩn (cỡ phù hợp lứa tuổi)

### Nội dung:
- [ ] Ngôn ngữ hướng tới HS (không có "HS nhận biết được...", "Mục tiêu:...")
- [ ] Không có slide liệt kê mục tiêu/năng lực kiểu giáo án
- [ ] Tối đa 3-4 bullets/slide

### Hình ảnh (MỚI):
- [ ] **Mỗi bullet point có ảnh riêng** (per-bullet image)
- [ ] Ảnh per-bullet đủ lớn: ≥ 2in×2in (TH Grid) hoặc 2in×1.5in (THCS Row)
- [ ] Slide trò chơi có ảnh minh họa đủ lớn cho HS nhìn
- [ ] Layout ảnh đúng cấp: Grid Flashcard (TH) / Horizontal Row (THCS)

### Animation (MỚI):
- [ ] Có hiệu ứng transition giữa các slide
- [ ] Slide câu hỏi/ghép nối: items animation tuần tự theo click
- [ ] Slide ghép nối: vế A hiện trước → click → vế B hiện
- [ ] Ảnh và text cùng item hiện ĐỒNG THỜI (cùng 1 click)

## Thư viện sử dụng
- `python-pptx`: Tạo .pptx
- `PyMuPDF (fitz)`: Đọc SGK PDF, trích xuất hình
- `Pillow`: Xử lý ảnh nếu cần crop

## Cải tiến liên tục
Sau mỗi lần tạo slide, ghi nhận feedback vào:
`D:\UNIGO\.agents\skills\tao-slide-bai-giang\references\improvement_log.md`
