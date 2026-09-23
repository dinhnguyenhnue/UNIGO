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

## 5. Output 2 cấp độ chuyên nghiệp

Hệ thống tự động sinh 2 cấp file Excel thống kê:

### Cấp 1: File thống kê đợt thi (`Thống kê [Cuộc thi] - Lần [X].xlsx`)
Lưu trực tiếp trong thư mục của lần thi đó (VD: `IOE/Lần 3/Thống kê IOE - Lần 3.xlsx`), gồm **4 sheets chuyên biệt**:
1. **Sheet "Tổng hợp toàn trường"**: Danh sách 128 học sinh gom theo 11 lớp, trạng thái `✅ Đã thi` (Xanh) / `❌ Chưa thi` (Đỏ), Vòng, Điểm, Thời gian (s).
2. **Sheet "Bảng xếp hạng (Đã thi)"**: Xếp hạng học sinh hoàn thành (Vòng desc, Điểm desc, Thời gian asc) kèm danh hiệu `🥇 Thủ khoa toàn trường` (Top 1), `🥈 Top 2-3`, `🥉 Top 4-5`.
3. **Sheet "Nhắc nhở GVCN (Chưa thi)"**: Gom theo lớp, chỉ hiển thị HS chưa thi, ghi rõ tỷ lệ `[Đã thi]/[Sĩ số] ([%]) | [Chưa thi]/[Sĩ số]` để GVCN copy gửi phụ huynh đôn đốc.
4. **Sheet "Thống kê theo lớp"**: 11 lớp, Sĩ số, Đã thi, Chưa thi, Tỷ lệ %, Điểm cao nhất kèm vòng, Thủ khoa của lớp.

### Cấp 2: Master File (`Thống kê cuộc thi.xlsx`)
Lưu tại thư mục gốc `Check_các_cuộc_thi/`, gồm **3 sheets tổng hợp**:
1. **Sheet "Tổng hợp"**: Ma trận theo dõi toàn bộ HS $\times$ tất cả các đợt thi (IOE Lần 1, 2, 3...), cột tổng số cuộc thi đã tham gia, hàng tỷ lệ % toàn trường.
2. **Sheet "Chi tiết [Tên cuộc thi]"**: Chi tiết điểm, vòng, thời gian theo từng lần.
3. **Sheet "Nhắc nhở GV"**: Tổng hợp danh sách nhắc nhở theo từng lớp cho toàn bộ các cuộc thi.

## 6. Danh sách HS gốc

File: `DANH SÁCH HỌC SINH NĂM HỌC 2026 - 2027.xlsx`

Cấu trúc:
- Mỗi sheet = 1 lớp (tên sheet = tên lớp: `1a1`, `1C1`, `2A1`...)
- Cột A = STT, Cột C = Họ tên hs
- Dòng 1 = tiêu đề, Dòng 2 = header, Dòng 3+ = dữ liệu HS

**Tổng: 128 học sinh, 11 lớp** (năm học 2026-2027)

## 7. Logic so khớp học sinh thông minh

1. **Chuẩn hóa chuỗi (Normalization):** Lowercase, strip, gộp nhiều dấu cách thành một. Tên lớp in hoa (`4c1` → `4C1`).
2. **Ưu tiên cùng lớp (Class-first):** So khớp chính xác cả họ tên lẫn lớp đăng ký.
3. **Fallback tên duy nhất (Unique Name Fallback):** Chỉ tự động gán lớp nếu tên học sinh là **duy nhất** trên quy mô toàn trường (tránh gán nhầm giữa các học sinh trùng tên ở các khối lớp khác nhau, ví dụ Nguyễn Đức Minh 1A1 và Nguyễn Đức Minh 2A1).
4. **Xử lý tài khoản trùng lặp (Multi-account / Multiple attempts):** Nếu học sinh có nhiều tài khoản hoặc thi lại (như Nguyễn Hương Mộc Lan 4C1), hệ thống tự động ưu tiên lấy bản ghi có **Vòng thi cao nhất** và **Điểm thi cao nhất**.

## 8. Xử lý lỗi

- File Excel không đọc được → in warning, bỏ qua
- Không tìm thấy cột tên/lớp → in warning cho sheet đó
- File output đang mở trong Excel → lưu bản `(new)` thay thế
- Console Windows không hỗ trợ Unicode → đã reconfigure UTF-8

## 9. Quy trình Agent khi nhận lệnh (Vòng điều phối khép kín)

1. Kiểm tra có file raw mới trong các thư mục lần thi (`Check_các_cuộc_thi/[Cuộc thi]/Lần [X]/`)
2. Chạy script: `python D:\UNIGO\scripts\check_cuoc_thi.py`
3. Xác nhận đã tạo file thống kê đợt thi (4 sheets) và cập nhật Master File tổng thể
4. Trình bày báo cáo tổng quan: Số lượng tham gia, so sánh tiến độ với các lần trước, vinh danh Top 5, thống kê theo lớp
5. Cung cấp đường dẫn clickable `file:///...` và sẵn sàng trích xuất danh sách đôn đốc gửi GVCN theo yêu cầu

