# Quy trình & Quy chuẩn tạo Kế hoạch dạy học (KHDH / Khung chương trình) UNIGO

Tài liệu này quy định quy trình, quy chuẩn kỹ thuật và các yêu cầu đầu ra bắt buộc cho Agent AI khi tạo, chỉnh sửa hoặc chuẩn hóa các văn bản Kế hoạch dạy học (KHDH), Khung chương trình phân phối môn học (Phụ lục IV - Bộ GD&ĐT) cho các môn học tại UNIGO.

---

## I. Yêu cầu đầu ra & Cấu trúc Bảng bài học (Phụ lục IV)

1. **Chuẩn hóa Bảng bài học 4 CỘT:**
   Mọi bảng phân phối bài học phải có chính xác **4 CỘT**:
   - **Cột 1 (STT):** Số thứ tự bài học / tiết học (Căn giữa).
   - **Cột 2 (Bài học (1)):** Tên bài học / Tiêu đề chủ đề (Căn trái. Các hàng tiêu đề Chủ đề gộp cell cả 4 cột, in đậm).
   - **Cột 3 (Số tiết (2)):** Số tiết hoặc dải tiết học tương ứng (ví dụ: `1`, `2`, `2, 3`, `4, 5`, `14 - 16`) (Căn giữa).
   - **Cột 4 (Yêu cầu cần đạt (3)):** Nội dung mô tả Yêu cầu cần đạt (YCCD) chi tiết của bài học (Căn trái).

   > [!CAUTION]
   > **CẤM TUYỆT ĐỐI:** Không tạo bảng 5 cột bị lặp tiêu đề `Yêu cầu cần đạt`, không chèn cột số lượng tiết lặp lại làm đẩy nội dung văn bản sang cột thứ 5.

2. **Phân bổ 35 Tiết/năm & Tuần Đánh giá định kỳ chuẩn:**
   - Mỗi khối lớp (Lớp 1 đến Lớp 8) luôn đảm bảo đúng **35 tiết/năm học**.
   - Bắt buộc bố trí các tiết **Ôn tập & Đánh giá định kỳ** vào đúng **4 mốc tuần chuẩn**:
     - **Tuần 10 (Tiết 9, 10):** Tiết 9: `Ôn tập Đánh giá định kỳ 1` | Tiết 10: `Đánh giá định kỳ 1`
     - **Tuần 19 (Tiết 18, 19):** Tiết 18: `Ôn tập Đánh giá định kỳ 2` | Tiết 19: `Đánh giá định kỳ 2`
     - **Tuần 28 (Tiết 27, 28):** Tiết 27: `Ôn tập Đánh giá định kỳ 3` | Tiết 28: `Đánh giá định kỳ 3`
     - **Tuần 34 (Tiết 33, 34):** Tiết 33: `Ôn tập Đánh giá định kỳ 4` | Tiết 34: `Đánh giá định kỳ 4`
     - **Tiết 35:** Bài học cuối / `Tổng kết năm học`.
   - **Lọc bỏ hàng cũ:** Phải lọc bỏ hoàn toàn các hàng "Đánh giá định kỳ" cũ từ file mẫu gốc trước khi chèn chuỗi tiết học mới để tránh bị lặp lại hàng đánh giá.

3. **Phần II: Kiểm tra, đánh giá định kỳ:**
   - Tên gọi bắt buộc đổi từ *Kiểm tra Giữa kì 1, Cuối kì 1, Giữa kì 2, Cuối kì 2* thành **Đánh giá định kỳ 1, Đánh giá định kỳ 2, Đánh giá định kỳ 3, Đánh giá định kỳ 4**.
   - Bảng 5 cột: `Bài kiểm tra, đánh giá` | `Thời gian (1)` | `Thời điểm (2)` | `Yêu cầu cần đạt (3)` | `Hình thức (4)`.
   - Cập nhật đúng thời điểm `Tuần 10`, `Tuần 19`, `Tuần 28`, `Tuần 34`.

4. **Định dạng văn bản & Font chữ:**
   - **Font chữ:** **Times New Roman**, cỡ chữ **13pt** áp dụng cho toàn bộ văn bản (đoạn văn, tiêu đề, tiêu đề bảng và nội dung các ô trong bảng).
   - **Vietnamese Casing:** Tuân thủ quy tắc viết hoa tiếng Việt (chỉ viết hoa chữ cái đầu câu và danh từ riêng, không viết hoa Title Case tùy tiện cho danh từ chung).

---

## II. Quy trình xử lý của Agent khi nhận yêu cầu

1. **Khảo sát & Đọc file dữ liệu đầu vào:**
   - Nếu có PDF Khung chương trình (ví dụ: Robotics, Tin học...), dùng script Python extraction (`pypdf` / `pdfplumber`) đọc trích xuất chính xác danh sách bài học, số bài, trọng tâm và yêu cầu cần đạt.

2. **Xử lý bảng và định dạng bằng Script Python (`python-docx`):**
   - Đảm bảo can thiệp cấp XML (`w:gridSpan`, `w:tcPr`) khi xử lý dòng gộp (Chủ đề) hoặc dòng phân tách (Bài học).
   - Áp dụng hàm `afont()` / `set_font_all()` quét toàn bộ paragraph và run để áp font Times New Roman 13pt.
   - Bọc hàm lưu file `save_doc()` bằng try-except PermissionError để tránh lỗi khóa file khi người dùng đang mở trong Word.

3. **Đồng bộ hệ thống file đầu ra:**
   Sau khi xử lý, Agent phải tạo/cập nhật đầy đủ bộ file đồng bộ:
   - `Kế hoạch dạy học môn [Môn học] 2026-2027.docx` (File tổng hợp Lớp 1 - 8)
   - `Kế hoạch dạy học môn [Môn học] (TH) - 2026 - 2027.docx` (Tiểu học Lớp 1 - 5)
   - `Kế hoạch dạy học môn [Môn học] (THCS) - 2026 - 2027.docx` (THCS Lớp 6 - 8)
   - Thư mục riêng từng lớp: `Kế hoạch dạy học [Môn học] từng lớp/Kế hoạch dạy học môn [Môn học] - Lớp [X] - 2026 - 2027.docx`
   - Copy sang thư mục phân phối chương trình môn học tương ứng.

4. **Báo cáo & Nghiệm thu:**
   - Kiểm tra empirically thông qua script Python xác nhận 100% các bảng đạt 4 cột, đủ 35 tiết/năm và không bị lỗi lặp hàng.
   - Trả lời ngắn gọn, rõ ràng kèm đường dẫn gạch chân `file:///...` clickable cho người dùng.
