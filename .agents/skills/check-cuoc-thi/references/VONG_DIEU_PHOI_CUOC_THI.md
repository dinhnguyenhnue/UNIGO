# Vòng điều phối Thống kê Cuộc thi Học sinh UNIGO (IOE, Violympic, ITE...)

## 1. Bản chất Vòng điều phối (Coordination Loop)
Vòng điều phối là quy trình khép kín tự động giữa AI Agent và Giáo viên Tin học / BGH / GVCN:

```
                      ┌───────────────────────────────────────┐
                      │ 1. Giáo viên tải file kết quả thô      │
                      │    vào Check_các_cuộc_thi/<Thi>/Lần <X>│
                      └──────────────────┬────────────────────┘
                                         │
                                         ▼
                      ┌───────────────────────────────────────┐
                      │ 2. Agent kích hoạt so khớp thông minh │
                      │    (Khớp lớp + Tên duy nhất + Multi-acc)
                      └──────────────────┬────────────────────┘
                                         │
                                         ▼
                      ┌───────────────────────────────────────┐
                      │ 3. Xuất file đợt thi (4 Sheets)       │
                      │    - Tổng hợp toàn trường              │
                      │    - Bảng xếp hạng (Vinh danh Top 1-5)│
                      │    - Nhắc nhở GVCN (Chưa thi theo lớp)│
                      │    - Thống kê tỷ lệ theo từng lớp      │
                      └──────────────────┬────────────────────┘
                                         │
                                         ▼
                      ┌───────────────────────────────────────┐
                      │ 4. Cập nhật Master Thống kê cuộc thi  │
                      │    (Ma trận toàn trường qua các lần)  │
                      └──────────────────┬────────────────────┘
                                         │
                                         ▼
                      ┌───────────────────────────────────────┐
                      │ 5. Agent báo cáo tóm tắt & trích xuất │
                      │    danh sách đôn đốc gửi GVCN         │
                      └───────────────────────────────────────┘
```

## 2. Quy tắc dữ liệu & Thuật toán
- **Danh sách học sinh gốc:** `DANH SÁCH HỌC SINH NĂM HỌC 2026 - 2027.xlsx` (128 HS, 11 lớp).
- **Trùng tên khác lớp:** Phải đối soát chính xác theo tên lớp đăng ký. Ví dụ: Nguyễn Đức Minh (1A1) và Nguyễn Đức Minh (2A1).
- **Học sinh thi nhiều tài khoản / làm lại:** Tự động lấy bản ghi có `Vòng` cao nhất, `Điểm` cao nhất. Ví dụ: Nguyễn Hương Mộc Lan (4C1) có 2 tài khoản, hệ thống chọn tài khoản Vòng 5, 1740 điểm.

## 3. Lịch sử các lần thi IOE (Năm học 2026 - 2027)
- **Lần 1:** 17/128 HS đã thi (13.28%)
- **Lần 2:** 26/128 HS đã thi (20.31%)
- **Lần 3:** 33/128 HS đã thi (25.78%)
  - Thủ khoa toàn trường: Trịnh Thảo Tiên (2C1) - Vòng 7, 2440 điểm
  - Lớp có tỷ lệ cao nhất: Lớp 2C1 (62%), Lớp 4C1 (58%), Lớp 3A1 (50%)
