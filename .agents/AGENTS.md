# Quy trình & Quy chuẩn tạo văn bản UNIGO

Tài liệu này quy định quy trình, quy chuẩn kỹ thuật và các yêu cầu đầu ra bắt buộc cho Agent AI khi tạo, chỉnh sửa hoặc chuẩn hóa các văn bản tại UNIGO (KHDH, KHBD, Kế hoạch tổ chuyên môn, Lịch báo giảng, Báo cáo, và tất cả các loại văn bản hành chính khác).

---

## 0. Quy tắc bảo tồn mẫu văn bản (Áp dụng cho TẤT CẢ các loại tài liệu)

> [!IMPORTANT]
> Quy tắc này áp dụng cho MỌI task xử lý file .docx/.pptx tại UNIGO: KHBD, Kế hoạch dạy học, Kế hoạch tổ chuyên môn, Lịch báo giảng, Báo cáo, và tất cả các văn bản hành chính khác.

1. **TUYỆT ĐỐI giữ nguyên phần ĐẦU và CUỐI của file mẫu gốc:**
   - **Header:** Logo UNIGO, tên trường, thông tin giáo viên nằm trong header. KHÔNG được xóa, ghi đè hoặc làm mất bất kỳ phần tử nào (đặc biệt là `w:drawing` chứa logo).
   - **Footer:** Chân trang mẫu (ký tên BGH, Tổ chuyên môn, Người soạn, hoặc thanh chân trang slide). KHÔNG được xóa, ghi đè hoặc thay đổi nội dung footer.
   - **Phần đầu trang văn bản:** Tiêu đề văn bản, thông tin trường, ngày tháng... phải giữ đúng format mẫu.
   - **Phần cuối trang văn bản:** Khung ký tên, Rút kinh nghiệm, Điều chỉnh bổ sung... phải giữ nguyên.

2. **Kỹ thuật python-docx bảo tồn Header/Footer:**
   - **Giữ `w:sectPr`:** Khi dọn body, lặp `doc.element.body` và bỏ qua child có tag kết thúc bằng `sectPr` (chứa references tới cả header lẫn footer).
   - **Sửa text trong Header:** KHÔNG gọi `paragraph.text = "..."` trên paragraph chứa `w:drawing`. Chỉ sửa text trong Run cụ thể (VD: `hp.runs[2].text`) để giữ nguyên Run 0 chứa logo drawing.
   - **Không thao tác Footer:** Footer nằm trong `section.footer` và được bảo tồn tự động khi giữ `w:sectPr`. KHÔNG truy cập hoặc chỉnh sửa footer trừ khi user yêu cầu rõ ràng.

3. **Kỹ thuật python-pptx bảo tồn chân trang slide:**
   - Giữ nguyên slide master/layout chứa thanh chân trang.
   - Khi thêm slide mới, luôn dùng layout từ template có sẵn chân trang.
   - KHÔNG xóa shapes ở vị trí chân trang.

4. **Kiểm tra sau xuất file:**
   - `.docx`: Xác nhận `header drawings count ≥ 1` và `footer paragraphs count ≥ 1`.
   - `.pptx`: Xác nhận mỗi slide có shape chân trang UNIGO.
   - Không có paragraphs rỗng thừa (50+) ở đầu file.

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
   - **Đường viền bảng (Table Borders):** Bắt buộc chèn XML `w:tblBorders` (`top`, `left`, `bottom`, `right`, `insideH`, `insideV` với `w:val="single"`, `w:sz="4"`, `w:color="000000"`) khi tạo hoặc thay thế bảng bằng `doc.add_table()`. Tuyệt đối không để bảng thiếu viền (borderless) và không phụ thuộc vào `table.style = 'Table Grid'` (tránh lỗi `KeyError: no style with name 'Table Grid'`).
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

---

## III. Quy định bắt buộc cho Kế hoạch Tổ chuyên môn (THCS)

1. **TUYỆT ĐỐI GIỮ NGUYÊN FORM MẪU FILE GỐC:**
   - Khi xử lý file `30.07.26. Kế hoạch tổ chuyên môn (THCS).docx` (hoặc các file Kế hoạch tổ chuyên môn tương tự), **CẤM TUYỆT ĐỐI** xóa bỏ form, thay đổi cấu trúc bảng mẫu hoặc tự ý nạp/thay thế các bảng 4 cột/bảng ngoài vào làm vỡ form của file gốc.
   - Giữ intact toàn bộ các Phần I, II, III, IV, V, VII, VIII của file gốc.

2. **CẤU TRÚC PHẦN VI. KẾ HOẠCH GIẢNG DẠY:**
   Nội dung phần **VI. KẾ HOẠCH GIẢNG DẠY** cho các môn học tổ THCS (Lớp 6, 7, 8) bắt buộc phải điền dữ liệu đúng theo khung form chuẩn sẵn có của file mẫu:
   - **Mỗi môn học gồm 3 mục:**
     - `X.1. Kế hoạch dạy học chính khoá`:
       - Đoạn thông tin số tiết (Cả năm, Học kì 1, Học kì 2, số điểm kiểm tra thường xuyên/định kỳ).
       - **Bảng 7 CỘT Phân phối bài học:** `TT` | `Bài/chủ đề` | `Tổng số tiết` | `Tuần` | `Tiết theo PPCT` | `Nội dung` | `Mục tiêu bài học`. Có dòng gộp ô `Học kì 1` và `Học kì 2`.
     - `X.2. Kế hoạch dạy học tăng cường`: Mô tả hoặc liệt kê nội dung dạy học tăng cường.
     - `X.3. Kế hoạch kiểm tra đánh giá`:
       - **Bảng 5 CỘT Kiểm tra đánh giá:** `TT` | `Lớp` | `Bài kiểm tra` | `Nội dung` | `Hình thức`. Cập nhật đủ 4 mốc (Đánh giá định kỳ 1, 2, 3, 4).

3. **Thứ tự 14 môn học cố định:**
   Sắp xếp lần lượt đúng 14 môn:
   `1. Ngữ văn` -> `2. Toán` -> `3. Tiếng Anh` -> `4. Tiếng Trung` -> `5. Khoa học tự nhiên` -> `6. Lịch sử và Địa lý` -> `7. Tin học` -> `8. Robotics` -> `9. Giáo dục công dân` -> `10. Giáo dục địa phương` -> `11. Hoạt động trải nghiệm, hướng nghiệp` -> `12. Âm nhạc` -> `13. Mỹ thuật` -> `14. Giáo dục thể chất`.
   - Nếu môn nào chưa có đủ chi tiết PPCT (ví dụ: Tiếng Trung, LS&ĐL, GDCD, Âm nhạc), giữ nguyên khung form chuẩn của môn đó và ghi rõ ghi chú bổ sung sau.

---

## IV. Quy định soạn Kế hoạch bài dạy (KHBD) chuẩn UNIGO (Tin học & Robotics)

1. **Mục tiêu Kiến thức & Mục tiêu Hoạt động**: Sử dụng Danh từ / Cụm danh từ trực tiếp. KHÔNG dùng động từ (hiểu, nhận diện, vận dụng, nêu, biết...). KHÔNG dùng cụm "Sự hiểu biết về...", "Khả năng nhận diện...". Mỗi gạch đầu dòng `-` PHẢI xuống dòng riêng. TUYỆT ĐỐI KHÔNG ghi chữ `(BẮT BUỘC)` tùy tiện.
2. **Phân nhóm Năng lực & Phẩm chất**: Phải chia làm 3 nhóm: Năng lực đặc thù (Tin học / Robotics với mã NLa-NLe chuẩn), Năng lực số (Thông tư 02/2025 - CV 3456 gồm 6 Miền I-VI, thành tố và Bậc chuẩn), Năng lực chung. Chỉ rõ mỗi năng lực/phẩm chất được phát triển qua `#Hoạt động` nào trong bài (ví dụ: `(Đạt được thông qua Hoạt động 2, Hoạt động 3)`). CẤM LẶP nội dung năng lực giữa các bài.
3. **Cấu trúc 4 Hoạt động & Bảng 2 Cột**:
   - 4 Hoạt động: `Hoạt động 1. Khởi động` -> `Hoạt động 2. Hình thành kiến thức mới` -> `Hoạt động 3. Luyện tập` -> `Hoạt động 4. Vận dụng`.
   - Mỗi hoạt động có 4 mục: `a) Mục tiêu`; `b) Nội dung`; `c) Sản phẩm`; `d) Tổ chức thực hiện`.
   - Bảng Tổ chức thực hiện dạng **2 Cột** (`HOẠT ĐỘNG CỦA GV – HS` | `KẾT QUẢ CẦN ĐẠT`). Cột 1 gộp hoạt động GV & HS theo Bước 1-4 (tên bước in đậm nghiêng). Cột 2 ghi kiến thức/kết quả cần đạt.
4. **Quy chuẩn Bảng thông tin đầu bài (Table 0 - NO BORDER) — Áp dụng cho cả Tin học & Robotics:**
   - **KHÔNG TRÙNG LẶP thông tin**: Mỗi thông tin (lớp, tên bài, môn học) chỉ xuất hiện MỘT LẦN.
   - **Bảng 3x2 không viền (`w:val="none"` / `w:val="nil"`):**
     - **Row 0:**
       - Cell 0: `Trường: Tiểu học và THCS UNIGO` (In đậm, TNR 13pt)
       - Cell 1: `Ngày soạn: DD/MM/YYYY` (In đậm, TNR 13pt - Thứ 7 tuần trước tuần dạy)
     - **Row 1:**
       - Cell 0: `GV: Đậu Đình Nguyên` (TNR 13pt)
       - Cell 1: `Ngày dạy: DD/MM/YYYY` (Đúng thứ theo LBG) (TNR 13pt)
     - **Row 2:**
       - Cell 0: `Tổ: Tổ chuyên môn Tiểu học` (Lớp 1-5) hoặc `Tổ chuyên môn THCS` (Lớp 6-8) (TNR 13pt)
       - Cell 1: `Lớp: [1A1, 2A1, 3A1, 4C1, 5C1, 6A1, 7A1, 8A1]` (TNR 13pt)
   - Tuyệt đối không để dòng rác (`Ngày dạy:    /     /2026`, dòng `Lớp` rỗng thừa).
   - Tên bài dạy chỉ xuất hiện 1 lần (`TÊN BÀI DẠY: ...`), KHÔNG thêm `Tên tiết:` trùng lặp.
5. **Quy chuẩn Rút kinh nghiệm & Bảng chữ ký cuối bài (Table 3x3 - NO BORDER) — Áp dụng cho cả Tin học & Robotics:**
   - **Phần Rút kinh nghiệm:**
     - `RÚT KINH NGHIỆM SAU BÀI DẠY:` (In đậm, TNR 13pt)
     - 2 dòng chấm: `...........................................................................................................................`
   - **Bảng chữ ký 3 hàng × 3 cột (NO BORDER):**
     - **Row 0:** `DUYỆT CỦA BGH` | `DUYỆT CỦA TỔ CM` | `NGƯỜI SOẠN` (In đậm, Căn giữa, TNR 13pt)
     - **Row 1:** `(Ký, ghi rõ họ tên)` | `(Ký, ghi rõ họ tên)` | `(Ký, ghi rõ họ tên)` (In nghiêng, Căn giữa, TNR 13pt)
     - **Row 2:** `\n\n\n` | `\n\n\n` | `\n\n\nĐậu Đình Nguyên` (Căn giữa, họ tên in đậm, TNR 13pt)
6. **Quy chuẩn Viền bảng (Table Borders) trong KHBD:**
   - **Bảng thông tin đầu bài (Table 0):** BẮT BUỘC **NO BORDER** (Không viền - `w:val="nil"`).
   - **Bảng các hoạt động dạy học (Tiến trình dạy học - Bảng 2 cột):** BẮT BUỘC **CÓ VIỀN** (`w:val="single"`, `w:sz="4"`, `w:color="000000"`).
   - **Bảng chữ ký cuối bài (Table 3x3):** BẮT BUỘC **NO BORDER** (Không viền - `w:val="nil"`).
7. **Định dạng thụt lề & Khoảng cách (EMU Rules):**
   - Tiêu đề mục La Mã (I., II., III.): `first_line_indent = 0` (In đậm).
   - Mục con cấp 1 (1. Kiến thức, 2. Năng lực, 3. Phẩm chất): `first_line_indent = 180340` (In đậm).
   - Mục con cấp 2 (2.1., 2.2., 2.3.): `first_line_indent = 360045` (In đậm).
   - Nội dung gạch đầu dòng bullet (- NLa..., - Miền I...): `left_indent = 540000`, `first_line_indent = 0`.
8. **Cấu trúc lưu trữ theo TUẦN (Đồng bộ Lịch báo giảng):**
   - Mọi file KHBD lưu theo đường dẫn: `KHBD_[Môn]/[Khối_lớp]/Tuần_[XX]/KHBD_[Môn]_[Khối_lớp]_Tiet[YY]_[Tên_bài].docx`.
   - Tuần được tính chính xác theo Lịch báo giảng và quy tắc xoay vòng Rotation chẵn/lẻ.

---

## V. Quy định Quản lý Git & Push dữ liệu (Git Workflow)

1. **Quy tắc Push toàn bộ (Full Push):**
   Khi người dùng yêu cầu "push", "đẩy lên git", hoặc sau khi hoàn thành task được yêu cầu đồng bộ:
   - **Bước 1:** Kiểm tra toàn bộ trạng thái repository bằng `git status`.
   - **Bước 2:** Bắt buộc dùng `git add -A` (hoặc `git add .`) để gom TOÀN BỘ file đã chỉnh sửa (Modified), file xóa (Deleted) và file mới tạo (Untracked). KHÔNG gom riêng lẻ từng file trừ khi user chỉ định cụ thể.
   - **Bước 3:** Commit với thông điệp rõ ràng, tóm tắt các công việc/file đã thay đổi: `git commit -m "..."`.
   - **Bước 4:** Đẩy lên remote repository: `git push origin [branch]`.
   - **Bước 5:** Chạy lại `git status` để xác nhận working tree hoàn toàn sạch (`nothing to commit, working tree clean`).

---

## VI. Quy định Lịch báo giảng (LBG) — PPCT & Rotation

1. **Quy tắc PPCT (Không áp dụng offset ngày trong tuần):**
   - **Tuần 1:** Tất cả các lớp (tất cả các Thứ) = PPCT 0 (`Tiết 0: Định hướng môn học`).
   - **Tuần 2 trở đi:** Tất cả các Thứ (từ Thứ 2 đến Thứ 6) đều đồng bộ `PPCT = tuan_so - 1` (Tuần 2 = PPCT 1, Tuần 3 = PPCT 2,...). Không áp dụng trễ tuần cho Thứ 2, Thứ 3.


2. **Quy tắc Rotation Tuần chẵn/lẻ cho lớp 5, 6, 7, 8:**
   - **Tuần CHẴN (2, 4, 6...):** Cả 2 tiết → **Tin học** (Đồ dùng: Phòng Tin học).
   - **Tuần LẺ (3, 5, 7...):** Cả 2 tiết → **Robotics** (Đồ dùng: Bộ Kit Robotics).
   - Khi có 2 tiết liên tiếp cùng môn (sau rotation), PPCT tính liên tiếp (VD: tiết 1 = PPCT N, tiết 2 = PPCT N+1) và tên bài cũng liên tiếp theo PPCT tương ứng.

3. **Tên bài tự động điền:**
   - PPCT = 0 → `"Tiết 0: Định hướng môn học"`.
   - PPCT ≥ 1 → Lấy tên bài từ `PPCT_TIN` / `PPCT_ROB` trong `generate_lbg.py` (dữ liệu trích xuất từ KHDH).

4. **Script & Output:**
   - Script: `scripts/generate_lbg.py <so_tuan>` (VD: `python generate_lbg.py 2`).
   - Output: 3 file trong `Hệ thống mẫu văn bản/Nguyên đã làm/Lịch báo giảng/`:
     - `Lịch báo giảng - Tuần XX.docx` (bản gốc tổng hợp)
     - `Lịch báo giảng - Tuần XX (TTH+TH).docx` (Tiền TH + TH)
     - `Lịch báo giảng - Tuần XX (THCS).docx` (THCS)

---

## VII. Quy định Tạo Slide Bài giảng (.pptx) chuẩn UNIGO

1. **Bảo tồn Slide Master & Vùng An Toàn (VERIFIED từ template):**
   - **Logo UNIGO** = `Picture 7` tại L=0.17in, T=0.15in, W=0.95in, H=0.94in → kết thúc tại **Y=1.09in**
   - **Chân trang UNIGO** = `Picture 9` tại L=0.00in, T=6.43in, W=13.40in, H=1.23in → bắt đầu từ **Y=6.43in**
   - **VÙNG AN TOÀN NỘI DUNG:** Y = **1.15in → 6.35in** (chiều cao 5.20in). Slide size: 13.33×7.50 inches.
   - **TUYỆT ĐỐI CẤM:**
     - Vẽ shape/rectangle/background có `top < 1.15in` (che logo)
     - Vẽ shape/rectangle/background có `top + height > 6.35in` (che chân trang)
     - Dùng `add_shape(RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)` phủ toàn bộ slide
   - **Kỹ thuật clamp bắt buộc:** `actual_top = max(top, 1.15)`, `actual_bottom = min(top+height, 6.35)`

2. **KHÔNG thêm footer tự tạo:**
   - Template master ĐÃ CÓ chân trang `Picture 9` (thanh xanh + thông tin trường + SĐT + địa chỉ).
   - KHÔNG thêm shape footer mới — chân trang tự động hiển thị trên mọi slide nhờ slide master.

3. **Z-Order & Tương phản (QUAN TRỌNG):**
   - **Background shapes phải `send_to_back`:** Gọi `spTree = sp.getparent(); spTree.remove(sp); spTree.insert(2, sp)` để đẩy xuống dưới cùng (sau `nvGrpSpPr` và `grpSpPr`). TUYỆT ĐỐI KHÔNG gọi `insert(0, sp)` vì sẽ làm hỏng cấu trúc XML làm PowerPoint báo lỗi repair.
   - **Text luôn ở trên:** Textbox phải được thêm SAU background shape.
   - **Tương phản cao bắt buộc:**
     - Chữ trắng (`FFFFFF`) trên nền tối (primary/accent đậm)
     - Chữ tối (`1A2744`, `2E1065`...) trên nền nhạt (bg/card trắng)
     - KHÔNG dùng chữ nhạt trên nền nhạt hoặc chữ tối trên nền tối
   - **Palette bắt buộc có 3 loại text color:** `text_on_primary`, `text_on_bg`, `text_on_card`

4. **Quy chuẩn Font chữ & Cỡ chữ:**
   - Tiêu đề slide chính / Giới thiệu: **24pt - 28pt** (Bold).
   - Nội dung thường (Bullet text): **18pt - 20pt**.
   - Ký tự đầu dòng (●): **14pt**, Giãn dòng (Line spacing): **28pt**, Khoảng cách sau đoạn (Space after): **8pt**.
   - Giới hạn nội dung mỗi slide ngắn gọn (tối đa 3 - 4 dòng bullet).

5. **Cấu trúc Group Card & Accent Bar:**
   - **Thanh Accent Bar khít tuyệt đối:** `bar.top = card.top`, `bar.height = card.height`.
   - **Bắt buộc Grouping:** Card + accent bar nhóm thành một khối (`group_shapes` qua `<p:grpSp>`).

6. **Slide Tổng kết & Màu sắc:**
   - Panel màu chỉ nằm gọn trong Vùng An Toàn (Y 1.15in → 6.35in).
   - KHÔNG gọi `set_slide_bg(primary)` phủ toàn slide.
   - Áp dụng hệ thống xoay vòng 8+ bộ màu (Color Palette rotation).

7. **Per-Bullet Images (BẮT BUỘC từ v3):**
   - **Mỗi bullet point trong slide nội dung (`learn`, `warmup`) PHẢI có ảnh minh họa riêng.**
   - Tạo prompt AI từ chính nội dung text bullet + context phù hợp lứa tuổi.
   - Layout phân cấp:
     - **Tiền TH → Lớp 5:** Layout A — Grid Flashcard (ảnh 2in×2in trên + text dưới)
     - **Lớp 6 → Lớp 8:** Layout B — Horizontal Row (ảnh 2in×1.5in trái + text phải)
   - Tối đa 3-4 bullets/slide. Nếu > 4 bullets → chia thành 2 slides.
   - Chuỗi fallback: SGK → AI-generated → Ảnh chung.

8. **Animation tuần tự cho slide Câu hỏi / Ghép nối (BẮT BUỘC từ v3):**
   - Slide `practice`/`activity` có `items` PHẢI animation từng item theo click.
   - Mỗi item (card text + ảnh) = 1 nhóm animation. HS suy nghĩ trước khi GV click hiện tiếp.
   - Slide ghép nối (chứa `↔`/`→`): vế A + ảnh hiện trước → click → vế B hiện.
   - Ảnh minh họa mỗi item xuất hiện CÙNG LÚC với text (cùng 1 click).
   - Hàm bắt buộc: `add_appear_animation(slide, shape, click_index)`, `add_group_animation(slide, shapes_list, click_index)`.

9. **Slide Trò chơi BẮT BUỘC có ảnh minh họa (từ v3):**
   - MỌI slide trò chơi / hoạt động (`activity`) PHẢI có ảnh minh họa đủ lớn (~2in×2in) để HS nhìn rõ.
   - Bảng loại trò chơi: Đúng/Sai, Ghép nối, Sắp xếp, Vẽ, Thảo luận, Nhận diện → mỗi loại có prompt AI mẫu riêng.
   - Layout: Phần trên (40%) = tiêu đề + hướng dẫn; Phần dưới (60%) = grid ảnh + card tương tác.
   - Animation tuần tự: ảnh + card item hiện lần lượt theo click.
