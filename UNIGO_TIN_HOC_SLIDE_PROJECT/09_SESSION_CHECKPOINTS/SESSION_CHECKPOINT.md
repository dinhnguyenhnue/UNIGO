# SESSION CHECKPOINT – TRẠNG THÁI PHIÊN LÀM VIỆC HIỆN TẠI

**Thời điểm lưu Checkpoint:** 2026-11-02 14:21  
**Phiên làm việc:** Chuẩn hóa & Nâng cấp Hệ thống Tạo Slide Bài giảng Tin học UNIGO (Giai đoạn 1 $\rightarrow$ 10)

---

## 1. Công việc đã hoàn thành trong phiên này
1. **Giai đoạn 1 & 2 (Kiểm tra quyền truy cập & Kiểm kê Skills):**
   - Đã xác nhận quyền đọc/ghi đầy đủ tại `D:\UNIGO\SGK`, `D:\UNIGO\KHBD_Tin_học_Chuẩn_Mới`, `D:\UNIGO\KHBD_Tin_học`, `D:\UNIGO\Hệ thống mẫu văn bản`, và `D:\UNIGO\.agents\skills\`.
   - Đã kiểm kê đầy đủ 7 Skills nội bộ của UNIGO (`tao-khbd`, `tao-slide-bai-giang`, `nang-luc-so`, `unigo-giao-vien`, `sgk-ocr`, `check-cuoc-thi`, `website-nang-luc`).
2. **Giai đoạn 3, 4 & 5 (Phân tích quy trình cũ, Khởi tạo cấu trúc dự án & Phân tách KHBD – Slide):**
   - Đã khởi tạo thư mục quản lý dự án `D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\` với đầy đủ 11 thư mục con (`00_PROJECT_PLAN` đến `10_CHANGELOG`).
   - Đã thiết lập bảng phân biệt rõ ràng giữa **KHBD (Tài liệu Giáo viên – `SKILL A`)** và **Slide (Học liệu Học sinh – `SKILL B`)** tại `02_ANALYSIS/PHAN_TICH_QUY_TRINH_CU_VA_PHAN_BIET_KHBD_SLIDE.md`.
   - Đã lập Bản đồ nội dung Slide (`Slide Content Map`) cho 2 bài mẫu chuẩn tại `03_LESSON_CONTENT_MAP/`.
3. **Giai đoạn 6, 7 & 8 (Tạo Bài mẫu hoàn chỉnh chuẩn Học sinh, Kiểm định 15 bước & Tinh chỉnh toàn hệ thống):**
   - Đã cắt chuẩn 13 hình ảnh gốc từ các trang SGK độ phân giải cao (`1700 × 2200 px`) và vẽ mới 2 sơ đồ tư duy Ghi nhớ Tiếng Việt 100% lưu tại `06_IMAGES/Lop3_Bai1/` và `06_IMAGES/Lop3_Bai3/`.
   - Đã xây dựng 2 bộ Slide bài mẫu hoàn chỉnh (14 slides/bộ, đa dạng bố cục `Layouts A–G`, tách riêng Slide Câu hỏi và Slide Đáp án & Giải thích, 100% ngôn ngữ học sinh) cho:
     - **Lớp 3 – Bài 1: Thông tin và quyết định** (`05_SLIDES_FINAL/Slide_Tin_hoc_Lop_3_Bai_1_Thong_tin_va_quyet_dinh_CHUAN_HS.pptx`)
     - **Lớp 3 – Bài 3: Máy tính và em** (`05_SLIDES_FINAL/Slide_Tin_hoc_Lop_3_Bai_3_May_tinh_va_em_CHUAN_HS.pptx`)
   - Đã đồng bộ 2 bộ bài mẫu này sang `KHBD_Tin_học_Chuẩn_Mới\Lớp_3\Slide_Bai_Giang\` và `KHBD_Tin_học\Lớp_3\`, đồng thời quét làm sạch `0/150` file PPTX toàn trường để loại bỏ triệt để mọi câu dẫn góc nhìn giáo viên (`"HS nhận ra..."`, `"GV nêu yêu cầu..."`).
   - Đã render toàn bộ slide bài mẫu ra ảnh PNG (`1333 × 750 px`) và ảnh tổng quan `overview_all_slides.png` tại `07_QUALITY_CONTROL/`, đạt tuyệt đối **15/15 tiêu chí kiểm định**.
4. **Giai đoạn 9 & 10 (Cập nhật Skills & Lưu Checkpoint):**
   - Đã cập nhật trực tiếp và xác nhận lưu thành công trên ổ đĩa 3 Skills cốt lõi:
     - `D:\UNIGO\.agents\skills\tao-khbd\SKILL.md` (**`SKILL A – UNIGO KHBD DESIGN`**)
     - `D:\UNIGO\.agents\skills\tao-slide-bai-giang\SKILL.md` (**`SKILL B – UNIGO STUDENT-FACING SLIDE DESIGN`**)
     - `D:\UNIGO\.agents\skills\unigo-giao-vien\SKILL.md` (Cơ chế liên kết `SKILL A` & `SKILL B`, lệnh ngắn & quy tắc đọc Checkpoint khi bắt đầu phiên mới).

## 2. Các quyết định đã thống nhất (Không được tự ý thay đổi ở phiên sau)
- **KHBD** là tài liệu cho giáo viên; **Slide** là tài liệu trình chiếu cho học sinh. Tuyệt đối không bê kịch bản điều hành lớp học hoặc bảng mã năng lực chuyên môn của giáo viên lên slide.
- Mọi câu hỏi Luyện tập trên slide phải có **Slide Câu hỏi riêng** để học sinh suy nghĩ trước, sau đó mới đến **Slide Đáp án & Giải thích** ở slide kế tiếp.
- Luôn sử dụng template `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx`, xóa 4 slide mặc định khi khởi tạo và giữ mọi đối tượng trong vùng an toàn `Y = 1.18 .. 6.25 in`.

## 3. Những vấn đề chưa xử lý / Công việc tiếp theo cho phiên sau
- Toàn bộ kiến trúc, 2 bài mẫu chuẩn Vàng (`Bài 1` & `Bài 3` Lớp 3) và 147 file PPTX hiện có đã đạt chuẩn.
- Khi người dùng yêu cầu thiết kế sâu từng bài cụ thể tiếp theo bằng lệnh ngắn (ví dụ: *"Tạo slide Tin học lớp 3 bài 4"*, *"Tạo KHBD và slide Tin học lớp 4 bài 5"*), Agent sẽ đọc `PROJECT_MASTER.md` và `SESSION_CHECKPOINT.md` này, lập `Slide Content Map` tại `03_LESSON_CONTENT_MAP/`, cắt ảnh SGK tương ứng vào `06_IMAGES/`, tạo slide theo chuẩn `SKILL B` (14 slides, tách Câu hỏi/Đáp án), render kiểm định 15 bước tại `07_QUALITY_CONTROL/` và cập nhật `PROGRESS_TRACKER`.

## 4. Các lỗi cần tránh lặp lại
1. Không gọi `import fitz` (PyMuPDF bị chặn bởi Windows Application Control Policy) $\rightarrow$ Sử dụng kho trang PNG đã trích xuất sẵn và `PIL.Image.crop()`.
2. Không dùng Emoji màu hệ thống (`📖`, `🎬`) hoặc thẻ `<a:br/>` (`\x0b`) trong PowerPoint $\rightarrow$ Dùng ký hiệu chuẩn Arial (`★`, `►`, `•`, `✓`, `➔`).
3. Không giữ lại 4 slide mặc định của file `Mẫu slide có chân trang.pptx` khi tạo bài mới.
4. Không đặt đáp án ngay dưới câu hỏi trắc nghiệm trên cùng 1 slide Luyện tập.
