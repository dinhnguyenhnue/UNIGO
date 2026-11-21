---
name: unigo-giao-vien
description: >
  Harness điều phối tổng hợp cho giáo viên UNIGO. Sử dụng khi user yêu cầu tạo 
  tài liệu giảng dạy (KHBD, slide, phiếu bài tập) cho bất kỳ bài/chủ đề nào.
  Skill này tự động dispatch sang tao-khbd và tao-slide-bai-giang.
---

# Harness Giáo viên UNIGO

## Tổng quan
Đây là skill điều phối chính, kết nối tất cả các skill phụ để tạo bộ tài liệu 
giảng dạy hoàn chỉnh cho giáo viên trường UNIGO.

## Luồng xử lý (Pipeline)

> [!IMPORTANT]
> **QUY TẮC MÔN HỌC & ĐIỀU PHỐI:**
> - **KHBD:** Tạo cho cả 2 môn **Tin học** và **Robotics**.
> - **Slide bài giảng (.pptx):** **CHỈ TẠO CHO MÔN TIN HỌC**. **TUYỆT ĐỐI KHÔNG TẠO SLIDE CHO MÔN ROBOTICS**.
> - **Thống kê Cuộc thi:** Tự động điều phối qua `check-cuoc-thi` khi có dữ liệu IOE, Violympic, ITE...

```
User Request
    ↓
    ├── [Tài liệu bài dạy]: Parse (Môn, Lớp, Bài) → [Đọc SGK OCR] → [Đọc PPCT]
    │       ├── tao-khbd → KHBD .docx (Tin học + Robotics) → D:\UNIGO\KHBD_[Môn]\
    │       ├── tao-slide-bai-giang → Slide .pptx (CHỈ TIN HỌC) → D:\UNIGO\KHBD_Tin_học\
    │       └── (tùy chọn) Phiếu bài tập, Đề kiểm tra
    │
    └── [Thống kê cuộc thi]: Check folder Check_các_cuộc_thi/ → So khớp danh sách 128 HS
            └── check-cuoc-thi → Xuất file Excel đợt thi (4 sheets) & Master file tổng thể
```

## Cách sử dụng

### Lệnh tạo bài dạy:
```
Tạo KHBD và slide cho Bài 1 Lớp 3 Tin học
Tạo KHBD và slide cho tất cả bài Lớp 3 Tin học
Tạo KHBD Bài 5 Lớp 6 Tin học
Tạo slide Bài 3 Lớp 4 Tin học
```

### Lệnh điều phối thống kê cuộc thi:
```
Thống kê cuộc thi IOE lần 3
Check tình hình học sinh thi IOE/Violympic
Xuất danh sách nhắc nhở GVCN học sinh chưa thi
```

### Lệnh chỉ slide:
```
Tạo slide Bài 3 Lớp 4 Tin học
```

## Parsing yêu cầu

Từ câu lệnh của user, trích xuất:
| Thông tin | Cách xác định | Mặc định |
|-----------|--------------|----------|
| Môn | Tìm trong câu | Tin học |
| Lớp | Số sau "Lớp" | Hỏi user |
| Bài | Số sau "Bài" | Bài 1 |
| Sản phẩm | "KHBD", "slide", hoặc cả hai | Cả hai |

## Cấu trúc thư mục workspace

```
D:\UNIGO\
├── SGK\Lớp_{X}\           # Input: SGK PDF
├── Công_văn_quy_định\     # Ref: Công văn, quy định
├── Hệ thống mẫu văn bản\ # Ref: Templates (KHBD, slide, v.v.)
├── Phân phối chương trình\ # Ref: PPCT
├── KHBD\Lớp_{X}\Bài {Y}\  # Output: Folder lưu KHBD .docx, Slide .pptx và ảnh AI
├── Tài liệu thêm\         # Tài liệu bổ sung
└── .agents\skills\         # Skills định nghĩa
    ├── unigo-giao-vien\    # Harness chính (file này)
    ├── tao-khbd\           # Skill tạo KHBD
    └── tao-slide-bai-giang\ # Skill tạo slide
```

## Quy tắc chung

### Font & Format
| Sản phẩm | Font | Cỡ chữ | Lề |
|----------|------|---------|-----|
| KHBD .docx | Times New Roman | 13 | Trái 3cm, Phải 2cm, Trên 2cm, Dưới 2cm |
| Slide .pptx | Arial | ≥18 (nội dung), ≥28 (tiêu đề) | Margin ≥1cm |

### Năng lực số
**BẮT BUỘC** trong mọi KHBD — phải có mục "Năng lực số" riêng biệt.

### Màu sắc slide & Hiệu ứng
- Thay đổi theo chủ đề bài học, KHÔNG cố định một màu.
- Đảm bảo tương phản tốt (chữ dễ đọc trên nền).
- Có hiệu ứng chuyển slide (Transition) và xuất hiện nội dung (Animation).

### Chân trang slide
Mọi slide đều có thanh chân trang xanh với "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO".

### Lưu file đúng vị trí
Tất cả tài liệu của từng bài học lưu gom trong 1 folder bài học nằm trong folder khối lớp:
- `D:\UNIGO\KHBD\Lớp_{X}\Bài {Y}\` (chứa file Word, PPTX và thư mục `images/`)

### Bảo tồn Header & Footer mẫu (QUY TẮC TOÀN CỤC)
Áp dụng cho TẤT CẢ sản phẩm .docx và .pptx:
- **Header .docx:** KHÔNG xóa/ghi đè paragraph chứa logo `w:drawing`. Chỉ sửa text trong Run cụ thể.
- **Footer .docx:** KHÔNG thao tác footer. Giữ `w:sectPr` khi dọn body để bảo tồn cả header lẫn footer.
- **Chân trang .pptx:** Giữ nguyên slide master/layout chứa thanh chân trang UNIGO. KHÔNG xóa shapes chân trang.
- **Phần đầu & cuối văn bản:** Tiêu đề, ngày tháng, khung ký tên, rút kinh nghiệm... luôn giữ đúng form mẫu gốc.

## Cải tiến liên tục (Learning Loop)

Sau mỗi phiên làm việc:
1. **Ghi nhận kết quả** vào `improvement_log.md`
2. **Nhận feedback** từ user (nếu có)
3. **Cập nhật skill** nếu phát hiện pattern mới hoặc lỗi cần sửa

### Format improvement log:
```markdown
## [Ngày] - [Bài/Lớp]
- **Sản phẩm**: KHBD / Slide / Cả hai
- **Kết quả**: Thành công / Cần sửa
- **Feedback**: [Ghi feedback của user]
- **Cải tiến áp dụng**: [Mô tả thay đổi cho lần sau]
```

### Nguyên tắc cải tiến:
- Nếu user chỉnh sửa format → cập nhật vào SKILL.md tương ứng
- Nếu user thêm yêu cầu mới → thêm vào references
- Nếu phát hiện lỗi font → cập nhật quy tắc font
- Nếu user yêu cầu thêm section → cập nhật cấu trúc template

## Tham chiếu nhanh

| Cần gì | Đọc ở đâu |
|--------|-----------|
| Mẫu KHBD Tiểu học | `Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx` |
| Mẫu KHBD THCS | `Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx` |
| Mẫu slide | `Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx` |
| Công văn 5512 | `Công_văn_quy_định\cong-van-5512-bgddt-2020_d8bd32d0a4.docx` |
| CT Tin học 2018 | `Công_văn_quy_định\16. CT_Tin hoc.docx` |
| **Khung CT Robotics UNIGO** | `Phân phối chương trình\Robotics\KHUNG CHƯƠNG TRÌNH ROBOTICS TIỂU HỌC & THCS UNIGO.docx` (.pdf) |
| **Giáo trình Robotics Khối 1-2** | `Phân phối chương trình\Robotics\Giáo trình - OLLO Initiate.docx` (.pdf) |
| **Giáo trình Robotics Khối 3-4** | `Phân phối chương trình\Robotics\Giáo trình - OLLO Kinder.docx` (.pdf) |
| **Giáo trình Robotics Khối 5-8** | `Phân phối chương trình\Robotics\Giáo trình - OLLO Excel 1.docx` (.pdf) |
| **Giáo trình Robotics Nâng cao** | `Phân phối chương trình\Robotics\Giáo trình - OLLO Spark.docx` (.pdf) |
| **Năng lực đặc thù Robotics** | `.agents\skills\tao-khbd\references\KHBD_NANG_LUC_DAC_THU_ROBOTICS.md` |
| **Năng lực đặc thù Tin học** | `.agents\skills\tao-khbd\references\KHBD_NANG_LUC_DAC_THU.md` |


---

## CẬP NHẬT ĐIỀU PHỐI QUY TRÌNH 10 BƯỚC & KIỂM ĐỊNH 100 ĐIỂM (GIAI ĐOẠN 8)

Khi nhận bất kỳ câu lệnh ngắn nào từ giáo viên (ví dụ: *"Soạn KHBD Tin học lớp X bài Y"*, *"Tạo giáo án và slide bài Y lớp X"*, *"Chuẩn hóa KHBD bài Y"*, *"Thiết kế lại bài dạy Tin học theo chuẩn UNIGO"*), `unigo-giao-vien` tự động điều phối:
1. Trích xuất nội dung & hình ảnh SGK từ `D:\UNIGO\SGK` và kho tham chiếu `D:\UNIGO\Phân phối chương trình\Tin học\KHBD Tin học 345 KNTT` (qua `sgk-ocr`).
2. Xây dựng Ma trận năng lực bài học (Năng lực chung, Năng lực đặc thù Tin học `NLa–NLe`, Năng lực số `CV 3456`, Định hướng Năng lực AI `QĐ 3439`) qua `nang-luc-so`.
3. Thiết kế KHBD `.docx` theo mẫu chuẩn `D:\UNIGO\Hệ thống mẫu văn bản\Khung_KHBD_Tieu_hoc_UNIGO.docx` có đầy đủ: Ma trận 4 cột (Mục tiêu – Hoạt động – Sản phẩm – Đánh giá), Phương pháp dạy học chủ đạo, 4 Hoạt động dạy học có Phân hóa 2 tầng (Cơ bản / Nâng cao), Hình ảnh SGK nhúng trực tiếp, và Bảng Rubric đánh giá 3 mức độ ở Mục V (qua `tao-khbd`).
4. Thiết kế bộ Slide bài giảng `.pptx` 12 slide trực quan dành cho học sinh (qua `tao-slide-bai-giang`).
5. Chạy kiểm định tự động thang 100 điểm (14 tiêu chí, đạt $\ge 90/100$ điểm, bảo tồn Header Logo & Footer UNIGO) trước khi xuất báo cáo.


---

## CƠ CHẾ ĐIỀU PHỐI LỆNH NGẮN (SHORT-COMMAND DISPATCH) & QUẢN LÝ CHECKPOINT (`D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\`)

### 1. Quy tắc Tự động Vận hành bằng Lệnh Ngắn (Mục X)
Khi người dùng đưa ra các lệnh ngắn như:
- *"Tạo slide Tin học lớp 3 bài 1"*
- *"Tạo KHBD và slide Tin học lớp 4 bài 5"*
- *"Chỉ tạo slide từ KHBD có sẵn"*
- *"Thiết kế lại slide bài học này"*
- *"Kiểm tra chất lượng slide vừa tạo"*
- *"Tiếp tục dự án từ checkpoint"*

Agent **TỰ ĐỘNG** xác định (không hỏi lại các quy chuẩn đã lưu):
1. **Tài liệu SGK cần sử dụng:** Tra tại `D:\UNIGO\SGK\Lớp_[X]\`.
2. **KHBD tương ứng:** Tra tại `D:\UNIGO\KHBD_Tin_học_Chuẩn_Mới\Lớp_[X]\` và `D:\UNIGO\KHBD_Tin_học\Lớp_[X]\`.
3. **Template UNIGO:**
   - Slide: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`.
   - KHBD: `D:\UNIGO\Hệ thống mẫu văn bản\Khung_KHBD_Tieu_hoc_UNIGO.docx`.
4. **Phân tách `SKILL A` & `SKILL B`:**
   - Khi tạo KHBD $\rightarrow$ Kích hoạt **`SKILL A – UNIGO KHBD DESIGN`** (`tao-khbd`).
   - Khi tạo Slide $\rightarrow$ Kích hoạt **`SKILL B – UNIGO STUDENT-FACING SLIDE DESIGN`** (`tao-slide-bai-giang`): Lập Bản đồ nội dung Slide (`03_LESSON_CONTENT_MAP\`), cắt ảnh SGK (`06_IMAGES\`), tạo Slide lấy học sinh làm trung tâm (`Layouts A–G`, tách riêng Slide Câu hỏi và Slide Đáp án & Giải thích), render PNG và kiểm định 15 bước (`07_QUALITY_CONTROL\`).

### 2. Quy tắc Tiếp tục Phiên làm việc (Session Resume Protocol – Mục VIII.4)
Khi bắt đầu phiên làm việc mới hoặc nhận lệnh *"Tiếp tục dự án từ checkpoint"*, bắt buộc thực hiện theo thứ tự:
1. Đọc `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\PROJECT_MASTER.md`.
2. Đọc `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\SESSION_CHECKPOINT.md`.
3. Đọc `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\PROGRESS_TRACKER.md` (`.xlsx`).
4. Đọc `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\CHANGELOG.md`.
5. Kiểm tra các file đầu ra thực tế trong `05_SLIDES_FINAL\` và `07_QUALITY_CONTROL\`, xác định trạng thái hiện tại và tiếp tục đúng công việc còn dang dở (không làm lại từ đầu, không ghi đè phiên bản đã hoàn thiện).
