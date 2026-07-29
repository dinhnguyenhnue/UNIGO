---
name: tao-khbd
description: >
  Tạo Kế hoạch bài dạy (KHBD) theo mẫu Unigo 2026-2027 (Tiểu học) hoặc PL4-THCS (THCS).
  Sử dụng khi user yêu cầu tạo giáo án, KHBD, kế hoạch bài dạy, hoặc lesson plan cho bất kỳ bài/chủ đề nào.
  Skill này tự động đọc SGK, phân tích nội dung, và tạo file .docx chuẩn.
---

# Skill Tạo KHBD (Kế hoạch bài dạy)

## Tổng quan
Skill này tạo KHBD theo đúng mẫu quy định của trường UNIGO và Bộ GD&ĐT.
KHBD được tạo dạng file .docx với đầy đủ format, font, margin theo quy định.

## Quy trình bắt buộc

### Bước 1: Xác định thông tin bài học
Từ yêu cầu của user, xác định:
- **Môn học**: Tin học, Toán, Tiếng Việt, v.v.
- **Lớp**: 1-9 (hoặc Pre-primary)
- **Bài số**: Bài 1, Bài 2, v.v.
- **Số tiết**: Theo PPCT

### Bước 2: Chọn mẫu KHBD
| Cấp | Mẫu | File tham chiếu |
|-----|------|-----------------|
| Tiểu học (Lớp 1-5) | Khung giáo án Unigo 2026-2027 | `D:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx` |
| THCS (Lớp 6-9) | PL4-Khung KHBD (CV5512) | `D:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx` |

### Bước 3: Đọc SGK & Xử lý OCR (PaddleOCR)
- Tìm file SGK trong `D:\UNIGO\SGK\Lớp_{X}\`
- Đối với SGK PDF/Ảnh quét: Sử dụng module PaddleOCR (`python d:\UNIGO\scripts\sgk_ocr.py <file_sgk.pdf>`) để tự động bóc tách văn bản Tiếng Việt, tiêu đề bài học và cấu trúc bảng biểu.
- Trích xuất nội dung bài học (mục tiêu, hoạt động, luyện tập, vận dụng).

### Bước 4: Đọc PPCT (nếu có)
- Tìm trong `D:\UNIGO\Phân phối chương trình\`
- Xác định tuần, tiết, chủ đề theo PPCT

### Bước 5: Tạo KHBD .docx
Sử dụng `python-docx` để tạo file .docx với:

#### Format bắt buộc:
- Font: **Times New Roman**, cỡ **13**
- Lề: trái 3cm, phải 2cm, trên 2cm, dưới 2cm (3, 2, 2, 2)
- Line spacing: 1.15

#### Cấu trúc KHBD Tiểu học (Mẫu Unigo):
```
TUẦN: ...          Ngày soạn: .../.../ 202...
                   Ngày dạy: .../.../ 202...
KẾ HOẠCH DẠY HỌC MÔN ...
CHỦ ĐỀ: ...
BÀI: ... (Tiết: (theo PPCT))

I. YÊU CẦU CẦN ĐẠT:
- Sau tiết học, học sinh sẽ: [liệt kê cụ thể]
1. Phát triển phẩm chất
2. Phát triển năng lực
   2.1. Năng lực môn học
   2.2. Năng lực chung và đặc thù
   2.3. Năng lực số (BẮT BUỘC)

II. ĐỒ DÙNG DẠY HỌC:
1. Giáo viên: ...
2. Học sinh: ...

III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC:
- Phương pháp: vấn đáp, động não, quan sát, thực hành, nhóm...
- Kĩ thuật: Think-Pair-Share, động não, tia chớp...

IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU:
[Bảng 2 cột: Hoạt động của GV | Hoạt động của HS]
1. Hoạt động MỞ ĐẦU (5 phút)
   1.1 Khởi động
   1.2 Kết nối bài học
2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI
   2.1 [Tên hoạt động] (... phút)
   2.2 [Tên hoạt động] (... phút)
3. HĐ LUYỆN TẬP - THỰC HÀNH
4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM

V. ĐIỀU CHỈNH - BỔ SUNG SAU TIẾT DẠY:
(GV ghi sau khi dạy xong)
```

#### Cấu trúc KHBD THCS (PL4-CV5512):
```
Trường: TRƯỜNG TIỂU HỌC VÀ THCS UNIGO
Họ tên giáo viên: ...
Tổ: ...
Ngày soạn: ... | Ngày dạy: ...

TÊN BÀI DẠY: ...
Tổng số tiết: ... | Tiết theo PPCT: ...

I. Mục tiêu
1. Kiến thức: [cụ thể]
2. Năng lực:
   - Năng lực chung: ...
   - Năng lực đặc thù: ...
   - Năng lực số: ... (BẮT BUỘC)
3. Phẩm chất: [cụ thể]

II. Thiết bị dạy học và học liệu
1. Thiết bị: ...
2. Học liệu: ...

III. Tiến trình dạy học
1. Hoạt động 1. Khởi động
   a) Mục tiêu  b) Nội dung  c) Sản phẩm  d) Tổ chức thực hiện
   Bước 1-4 cho mỗi hoạt động
2. Hoạt động 2. Hình thành kiến thức mới
3. Hoạt động 3. Luyện tập
4. Hoạt động 4. Mở rộng (Nhiệm vụ về nhà)

RÚT KINH NGHIỆM SAU BÀI DẠY
BGH | TỔ CHUYÊN MÔN | NGƯỜI SOẠN
```

### Bước 6: Lưu file
**Quy tắc đặt tên và cấu trúc lưu file:**
```
D:\UNIGO\KHBD\Lớp_{X}\Bài {Y}\
├── KHBD_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.docx
├── Slide_{Môn}_{Lớp}_Bai{XX}_{Ten_bai}.pptx
└── images\
```
Ví dụ: `D:\UNIGO\KHBD\Lớp_3\Bài 1\KHBD_Tin_hoc_3_Bai01_Thong_tin_va_quyet_dinh.docx`

## Yêu cầu nội dung Năng lực số (BẮT BUỘC)
Mục Năng lực số phải bao gồm các biểu hiện cụ thể:
- Nhận biết và sử dụng các công cụ số cơ bản
- Ý thức an toàn thông tin trong môi trường số
- Tư duy số: xử lý, phân tích thông tin bằng công cụ CNTT
- Giao tiếp trong môi trường số một cách có trách nhiệm

## Cải tiến liên tục
Sau mỗi lần tạo KHBD, ghi nhận feedback vào file:
`D:\UNIGO\.agents\skills\tao-khbd\references\improvement_log.md`

Nội dung ghi nhận:
- Ngày tạo
- Bài/Lớp đã tạo
- Feedback từ user (nếu có)
- Điều chỉnh áp dụng cho lần sau
