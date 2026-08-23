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
> **QUY TẮC MÔN HỌC:**
> - **KHBD:** Tạo cho cả 2 môn **Tin học** và **Robotics**.
> - **Slide bài giảng (.pptx):** **CHỈ TẠO CHO MÔN TIN HỌC**. **TUYỆT ĐỐI KHÔNG TẠO SLIDE CHO MÔN ROBOTICS**.

```
User Request → Parse (Môn, Lớp, Bài) → [Đọc SGK + PaddleOCR Engine] → [Đọc PPCT] 
    ↓
    ├── tao-khbd → KHBD .docx (Tin học + Robotics) → D:\UNIGO\KHBD_Tin_học / KHBD_Robotics\
    ├── tao-slide-bai-giang → Slide .pptx (CHỈ TIN HỌC) → D:\UNIGO\KHBD_Tin_học\Lớp_{X}\Tuần_{YY}\
    └── (tùy chọn) Phiếu bài tập, Đề kiểm tra
```

## Cách sử dụng

### Lệnh đơn giản:
```
Tạo KHBD và slide cho Bài 1 Lớp 3 Tin học
```

### Lệnh batch:
```
Tạo KHBD và slide cho tất cả bài Lớp 3 Tin học
```

### Lệnh chỉ KHBD:
```
Tạo KHBD Bài 5 Lớp 6 Tin học
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
