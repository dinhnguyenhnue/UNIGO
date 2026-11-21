# PROJECT MASTER – HỆ THỐNG TẠO SLIDE BÀI GIẢNG TIN HỌC UNIGO

**Cập nhật lần cuối:** 2026-11-02 14:21  
**Trạng thái tổng thể:** ĐÃ CHUẨN HÓA KIẾN TRÚC, BÀI MẪU HOÀN CHỈNH, KIỂM ĐỊNH 15 BƯỚC & ĐỒNG BỘ SKILLS (`SKILL A` & `SKILL B`)

---

## 1. Mục tiêu dự án
Xây dựng và vận hành hệ thống thiết kế Kế hoạch bài dạy (KHBD) và Slide bài giảng môn Tin học UNIGO chuyên nghiệp, trực quan, chuẩn kiến thức Chương trình GDPT 2018 (Bộ sách *Kết nối tri thức với cuộc sống*), tích hợp Năng lực số (`TT 02/2025 – CV 3456`) và định hướng Năng lực AI (`QĐ 3439`), đảm bảo **phân tách tuyệt đối giữa tài liệu dành cho Giáo viên (KHBD) và học liệu trình chiếu dành cho Học sinh (Slide)**.

## 2. Phạm vi & Danh sách tài liệu nguồn
- **Sách giáo khoa (SGK) nguồn chuẩn:**
  - `D:\UNIGO\SGK\Lớp_3\` (`SGK Tin học 3 KNTT.pdf` + 25 trang trích xuất độ phân giải cao `1700 × 2200 px` tại `extracted_pages\`).
  - `D:\UNIGO\SGK\Lớp_4\` (`SGK Tin học 4 KNTT.pdf` + kho ảnh trích xuất `all_extracted_images\`).
  - `D:\UNIGO\SGK\Lớp_5\` (`SGK Tin học 5 KNTT.pdf` + kho ảnh trích xuất `all_extracted_images\`).
  - Các khối lớp Tiền Tiểu học, Lớp 1, 2, 6, 7, 8, 9 theo KHDH UNIGO.
- **Kho KHBD chuẩn UNIGO:**
  - `D:\UNIGO\KHBD_Tin_học_Chuẩn_Mới\` (`105 file .docx` Lớp 3, 4, 5 đạt chuẩn 99,86/100 điểm + `330 ảnh SGK` kèm `image_manifest.json`).
  - `D:\UNIGO\KHBD_Tin_học\` (`111 file .docx` Lớp 3, 4, 5 đồng bộ + toàn bộ các khối lớp).
- **Template thương hiệu UNIGO bắt buộc:**
  - Slide: `D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx` (`13.333 × 7.50 in`, bảo toàn `Picture 7` Logo UNIGO và `Picture 9` Chân trang UNIGO trên `slide_master`).
  - KHBD: `D:\UNIGO\Hệ thống mẫu văn bản\Khung_KHBD_Tieu_hoc_UNIGO.docx`.

## 3. Quy tắc thiết kế cốt lõi & Quyết định đã thống nhất
1. **Nguyên tắc phân tách KHBD (`SKILL A`) và Slide (`SKILL B`):**
   - **KHBD (`SKILL A`):** Dành cho giáo viên; mô tả kịch bản tổ chức dạy học, phương pháp, phân hóa 2 tầng (Cơ bản/Nâng cao), mã năng lực (`NLa–NLe`, `CV 3456`, `QĐ 3439`), dự kiến sản phẩm và bảng Rubric đánh giá.
   - **Slide (`SKILL B`):** Dành cho học sinh quan sát, suy nghĩ, khám phá và ghi nhớ; **100% ngôn ngữ hướng tới học sinh**; **CẤM** sao chép lời dẫn giáo viên (`"GV yêu cầu..."`, `"HS nhận ra..."`, `"Bước 1..."`) hoặc bảng mã năng lực chuyên môn lên slide.
2. **Quy trình 4 bước trước khi thiết kế Slide:**
   - `Bước 1: Đọc SGK` $\rightarrow$ `Bước 2: Đọc KHBD tương ứng` $\rightarrow$ `Bước 3: Lập Bản đồ nội dung Slide (Slide Content Map)` $\rightarrow$ `Bước 4: Kiểm tra tính thống nhất`.
3. **Đa dạng hóa bố cục theo mục đích sư phạm (Layouts A–G):**
   - `A. Slide giới thiệu` | `B. Slide khởi động` | `C. Slide mục tiêu (ngôn ngữ HS)` | `D. Slide hình thành kiến thức (chia cột, nhãn chú thích Callout, sơ đồ quy trình, bảng so sánh Đúng/Sai)` | `E. Slide ghi nhớ (Sơ đồ tư duy Mindmap / Từ khóa)` | `F. Slide luyện tập (Tách riêng Slide Câu hỏi và Slide Đáp án & Giải thích kế tiếp)` | `G. Slide vận dụng`.
4. **Quy trình xử lý hình ảnh:**
   - Ưu tiên cắt chuẩn (crop) hình ảnh gốc từ trang SGK đúng tỷ lệ, không méo, không cắt mất chữ/thông tin; bổ sung sơ đồ tư duy / infographic Tiếng Việt 100% khi cần hệ thống hóa kiến thức.
5. **Kiểm định chất lượng 15 bước bắt buộc:**
   - Mọi file PPTX sau khi xuất đều phải render sang PNG (`1333 × 750 px`), ghép ảnh tổng quan `overview_all_slides.png` và kiểm tra trực quan 15 tiêu chí trước khi bàn giao.

## 4. Cấu trúc thư mục Quản lý Dự án (`D:\UNIGO\UNIGO_TIN_HOC_SLIDE_PROJECT\`)
- `00_PROJECT_PLAN/`: Lưu `PROJECT_MASTER.md`, `PROGRESS_TRACKER.md`, `PROGRESS_TRACKER.xlsx`.
- `01_SOURCE_DOCUMENTS/`: Liên kết và danh mục tài liệu nguồn SGK, KHBD, Template UNIGO.
- `02_ANALYSIS/`: Báo cáo phân tích quy trình cũ & bảng phân biệt KHBD – Slide.
- `03_LESSON_CONTENT_MAP/`: Bản đồ nội dung slide (`Slide Content Map`) trước khi thiết kế.
- `04_SLIDES_WORKING/`: Các bản nháp slide đang xử lý.
- `05_SLIDES_FINAL/`: Các bộ slide chuẩn hoàn thiện đã qua kiểm định 15 bước.
- `06_IMAGES/`: Kho ảnh cắt chuẩn từ SGK và sơ đồ tư duy / infographic theo từng bài.
- `07_QUALITY_CONTROL/`: Ảnh render từng slide, `overview_all_slides.png` và báo cáo kiểm định JSON/MD.
- `08_SKILLS/`: Bản lưu đối chiếu của `SKILL A (tao-khbd)`, `SKILL B (tao-slide-bai-giang)`, `unigo-giao-vien`.
- `09_SESSION_CHECKPOINTS/`: Lưu `SESSION_CHECKPOINT.md` phục hồi ngữ cảnh tức thì.
- `10_CHANGELOG/`: Lưu `CHANGELOG.md` ghi nhật ký mọi thay đổi.

## 5. Tiến độ tổng thể
- **Giai đoạn 1–5 (Kiểm tra quyền truy cập, Kiểm kê Skills, Phân tích quy trình cũ, Lập kế hoạch, Phân biệt KHBD & Slide):** Đã hoàn thành 100%.
- **Giai đoạn 6–8 (Xây dựng Bài mẫu hoàn chỉnh chuẩn Học sinh + Kiểm định 15 bước + Chuẩn hóa toàn bộ 150 file PPTX):** Đã hoàn thành 100%.
- **Giai đoạn 9–10 (Nâng cấp `SKILL A`, `SKILL B`, `unigo-giao-vien` & Hoàn thiện hệ thống Checkpoint):** Đã hoàn thành 100%.
