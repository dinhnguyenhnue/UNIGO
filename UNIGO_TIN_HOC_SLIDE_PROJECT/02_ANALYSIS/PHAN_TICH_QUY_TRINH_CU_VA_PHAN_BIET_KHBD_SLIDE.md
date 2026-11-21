# BÁO CÁO PHÂN TÍCH QUY TRÌNH CŨ & CHUẨN HÓA RANH GIỚI KHBD (SKILL A) – SLIDE (SKILL B)

**Ngày cập nhật:** 2026-11-02 14:21  
**Dự án:** Chuẩn hóa và Nâng cấp Hệ thống Tạo Slide Bài giảng Tin học UNIGO

---

## 1. Phân tích quy trình cũ & Nguyên nhân gốc rễ (Giai đoạn 3)

Qua kiểm toán thực chứng 216 file KHBD (`.docx`) và 147 bộ Slide (`.pptx`) cùng kiểm tra ảnh render trực quan (`overview_all_slides.png`), quy trình tạo slide cũ có **5 hạn chế cốt lõi về mặt sư phạm và thiết kế**:

1. **Rò rỉ ngôn ngữ kịch bản giáo viên lên màn hình học sinh:**
   - Một số đoạn trích từ KHBD gốc vẫn còn câu mô tả của giáo viên như `"HS nhận ra được vai trò của thông tin..."` hoặc `"GV nêu yêu cầu bài tập: Ngày mai, ngoài giờ đi học..."`.
2. **Hiển thị đáp án ngay bên dưới câu hỏi trên cùng một slide:**
   - Ở phần Luyện tập trước đây, câu hỏi trắc nghiệm và dòng `"Đáp án đúng: ..."` xuất hiện cùng lúc trên 1 slide, làm mất cơ hội suy nghĩ, thảo luận và trả lời của học sinh.
3. **Bố cục cố định (Fixed Card Template) chưa biến thiên linh hoạt theo dạng bài:**
   - Mọi bài học đều bị ép cứng vào 11 slide với cùng một kiểu chia khung, chưa tách riêng slide **Sơ đồ tư duy Ghi nhớ (Mindmap)**, slide **Bảng so sánh**, slide **Nhận diện hình ảnh đánh số (1-2-3-4)** và cặp slide **Câu hỏi $\rightarrow$ Đáp án & Giải thích**.
4. **Chưa khai thác triệt để ảnh cắt chuẩn từ trang SGK gốc (`1700 × 2200 px`):**
   - Trước đây chỉ lấy ảnh nhúng sẵn trong file Word tham chiếu (đôi khi bị thiếu hoặc độ phân giải thấp) thay vì cắt đúng vùng (crop) trực tiếp từ trang SGK chuẩn (`Hình 1`, `Hình 2`, `Hình 3`, `Hình 8`, `Hình 14`, `Hình 17`...).

---

## 2. Bảng Phân biệt Tuyệt đối giữa KHBD (Tài liệu Giáo viên) và Slide (Học liệu Học sinh) (Giai đoạn 5)

| Tiêu chí đối sánh | KẾ HOẠCH BÀI DẠY – KHBD (`SKILL A`) | SLIDE BÀI GIẢNG – PPTX (`SKILL B`) |
|:---|:---|:---|
| **Đối tượng sử dụng** | **Giáo viên** (Tài liệu thiết kế và điều hành tiết học). | **Học sinh** (Tài liệu trình chiếu trực tiếp trên bảng tương tác/máy chiếu). |
| **Mục tiêu & Yêu cầu cần đạt** | Ghi đầy đủ mã năng lực đặc thù (`NLa–NLe`), mã Năng lực số (`CV 3456` – `1.1.CB1a`), Năng lực chung, Phẩm chất, Định hướng AI (`QĐ 3439`) và Bảng Ma trận 4 cột. | Chỉ trình bày **3 mục tiêu ngắn gọn, dễ hiểu bằng ngôn ngữ học sinh** (*"Sau bài học này em sẽ..."*). **CẤM** đưa bảng mã năng lực chuyên môn của GV lên slide. |
| **Cách diễn đạt** | Mô tả quy trình sư phạm: *"GV tổ chức...", "Bước 1: Chuyển giao nhiệm vụ...", "HS thảo luận nhóm đôi...", "Phương án dự kiến..."*. | Đặt câu hỏi gợi mở, tình huống trực quan, từ khóa cốt lõi nói trực tiếp với học sinh: *"Em hãy quan sát Hình 8 và cho biết...", "Vì sao bạn Minh quyết định...?"*. |
| **Hình thành kiến thức** | Kịch bản dẫn dắt, câu hỏi gợi mở, cách GV hỗ trợ HS yếu (Mức Cơ bản) và mở rộng cho HS khá giỏi (Mức Nâng cao). | Hình ảnh SGK lớn có nhãn chú thích (Callout), sơ đồ cấu tạo, sơ đồ quy trình, bảng so sánh Đúng/Sai (Nên/Không nên). |
| **Ghi nhớ & Chốt kiến thức** | Đoạn văn tổng kết kiến thức chuẩn trong hộp Ghi nhớ để GV chốt lời giảng. | **Sơ đồ tư duy (Mindmap)**, bảng từ khóa trung tâm, quy tắc ngắn gọn dễ nhớ. |
| **Luyện tập & Thực hành** | Đề bài + Kịch bản tổ chức trò chơi/thảo luận + Đáp án chi tiết + Bảng tiêu chí đánh giá Rubric 3 mức (`T / H / C`). | **Tách riêng 2 slide:** **Slide Câu hỏi** (để HS suy nghĩ trước) $\rightarrow$ **Slide Đáp án & Giải thích** (hiển thị sau khi HS trả lời). |
