---
name: check-cuoc-thi
description: >
  Thống kê cuộc thi học sinh UNIGO. Tự động quét thư mục Check_các_cuộc_thi,
  đọc kết quả thi từ IOE/Violympic/ITE..., so khớp với danh sách HS toàn trường,
  xuất file Excel thống kê (tổng hợp, chi tiết, nhắc nhở GV).
  Sử dụng khi user yêu cầu thống kê, check cuộc thi, hoặc tạo báo cáo HS tham gia thi.
---

# Skill: Thống kê Cuộc thi Học sinh UNIGO

## 1. Tổng quan

Hệ thống thống kê tự động giúp GV Tin học theo dõi tình hình tham gia các cuộc thi online (IOE, Violympic, ITE...) của toàn bộ học sinh trường UNIGO.

## 2. Cấu trúc thư mục

```
D:\UNIGO\Check_các_cuộc_thi\
├── DANH SÁCH HỌC SINH NĂM HỌC 2026 - 2027.xlsx   ← Danh sách gốc (11 sheet = 11 lớp)
├── IOE\                                              ← Tên cuộc thi
│   ├── Lần 1\                                        ← Lần thi (round)
│   │   └── *.xlsx                                    ← File kết quả tải từ hệ thống
│   ├── Lần 2\
│   │   └── *.xlsx
│   └── ...
├── Violympic\                                        ← Cuộc thi khác
│   ├── Lần 1\
│   └── ...
├── ITE\
│   └── ...
└── Thống kê cuộc thi.xlsx                            ← OUTPUT (tự động tạo bởi script)
```

## 3. Cách thêm cuộc thi mới

1. **Tạo thư mục** với tên cuộc thi trong `Check_các_cuộc_thi/` (VD: `Violympic/`)
2. **Tạo thư mục lần thi** bên trong: `Lần 1/`, `Lần 2/`, ...
3. **Bỏ file Excel** kết quả thi (tải từ hệ thống thi) vào thư mục lần tương ứng
4. **Chạy script**: `python D:\UNIGO\scripts\check_cuoc_thi.py`

> Script tự động nhận diện cột "Họ và Tên" / "Họ tên" và "Lớp" / "Khối" từ header của file Excel.
> Không cần cấu hình gì thêm.

## 4. Chạy script

```bash
python D:\UNIGO\scripts\check_cuoc_thi.py
```

## 5. Output: Thống kê cuộc thi.xlsx

File output có **3 loại sheet**:

### Sheet "Tổng hợp"
- Mỗi hàng = 1 học sinh, gom theo lớp
- Các cột = cuộc thi × lần thi → `✅ Đã thi` / `❌ Chưa thi`
- Cột tổng: Số cuộc thi đã tham gia
- Conditional formatting: Xanh = đã thi, Đỏ = chưa thi

### Sheet "Chi tiết [Tên cuộc thi]"
- Chi tiết kết quả: Vòng, Điểm, Thời gian tự luyện
- Phân tách theo từng lần thi

### Sheet "Nhắc nhở GV"
- Gom theo lớp + cuộc thi
- Chỉ liệt kê HS **chưa thi** → copy gửi GV chủ nhiệm
- Có thống kê tỷ lệ: `(đã thi/tổng, chưa thi/tổng)`

## 6. Danh sách HS gốc

File: `DANH SÁCH HỌC SINH NĂM HỌC 2026 - 2027.xlsx`

Cấu trúc:
- Mỗi sheet = 1 lớp (tên sheet = tên lớp: `1a1`, `1C1`, `2A1`...)
- Cột A = STT, Cột C = Họ tên hs
- Dòng 1 = tiêu đề, Dòng 2 = header, Dòng 3+ = dữ liệu HS

**Tổng: 128 học sinh, 11 lớp** (năm học 2026-2027)

## 7. Logic so khớp tên HS

1. Normalize: lowercase, collapse spaces, strip
2. Match exact normalized name
3. Ưu tiên match cùng lớp, fallback match khác lớp (cùng tên)
4. Tên lớp normalize uppercase (VD: `4c1` → `4C1`)

## 8. Xử lý lỗi

- File Excel không đọc được → in warning, bỏ qua
- Không tìm thấy cột tên/lớp → in warning cho sheet đó
- File output đang mở trong Excel → lưu bản `(new)` thay thế
- Console Windows không hỗ trợ Unicode → đã reconfigure UTF-8

## 9. Quy trình Agent khi user yêu cầu

1. Kiểm tra có file mới trong các thư mục cuộc thi không
2. Chạy `python D:\UNIGO\scripts\check_cuoc_thi.py`
3. Mở file output và báo cáo kết quả cho user
4. Nếu user muốn gửi nhắc nhở → trích sheet "Nhắc nhở GV" theo lớp cần thiết
