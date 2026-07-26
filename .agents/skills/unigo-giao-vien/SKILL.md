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

```
User Request → Parse (Môn, Lớp, Bài) → [Đọc SGK + PaddleOCR Engine] → [Đọc PPCT] 
    ↓
    ├── tao-khbd → KHBD .docx → D:\UNIGO\KHBD\Lớp_{X}\
    ├── tao-slide-bai-giang → Slide .pptx → D:\UNIGO\Slide_bài_giảng\Lớp_{X}\
    └── (tùy chọn) Phiếu bài tập, Đề kiểm tra
    ↓
    Review → Feedback → Cải tiến → Lưu improvement_log.md
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
├── KHBD\Lớp_{X}\          # Output: Kế hoạch bài dạy
├── Slide_bài_giảng\Lớp_{X}\ # Output: Slide bài giảng
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
| KHBD .docx | Times New Roman | 13 | Trên/dưới 2cm, trái 2.5cm, phải 1.5cm |
| Slide .pptx | Arial | ≥18 (nội dung), ≥28 (tiêu đề) | Margin ≥1cm |

### Năng lực số
**BẮT BUỘC** trong mọi KHBD — phải có mục "Năng lực số" riêng biệt.

### Màu sắc slide
Thay đổi theo chủ đề bài học, KHÔNG cố định một màu.
Đảm bảo tương phản tốt (chữ dễ đọc trên nền).

### Chân trang slide
Mọi slide đều có thanh chân trang xanh với "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO".

### Lưu file đúng vị trí
- KHBD → `D:\UNIGO\KHBD\Lớp_{X}\`
- Slide → `D:\UNIGO\Slide_bài_giảng\Lớp_{X}\`

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
| Công văn 5512 | `Công_văn_quy_định\Khung Kế hoạch bài dạy_ Theo Phụ lục IV, Công văn 5512_BGDĐT-GDTrH.docx` |
| CT Tin học | `Công_văn_quy_định\16. CT_Tin hoc.pdf` |
| PPCT ICT | `Phân phối chương trình\Môn ICT\` |
