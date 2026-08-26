---
name: sgk-ocr
description: Quy trình và bộ công cụ OCR Sách Giáo Khoa (SGK) tự động chuyển đổi PDF sang DOCX chuẩn theo từng bài, từng khối lớp cho UNIGO.
---

# SGK OCR Pipeline — Hệ Thống OCR & Chuyển Đổi SGK sang DOCX

Hệ thống tự động hóa quá trình số hóa Sách giáo khoa (PDF) thành các file tài liệu Word (.docx) được phân tách chính xác theo từng bài học, khối lớp, phục vụ trực tiếp cho việc biên soạn Slide bài giảng và Kế hoạch bài dạy (KHBD).

---

## 1. Kiến Trúc 4 Giai Đoạn (4-Stage Pipeline)

```
[SGK PDF]
    │
    ▼ (Stage 1: sgk_renderer.py)
[Trang PNG 300 DPI + Ảnh Minh Họa Gốc]
    │
    ▼ (Stage 2: sgk_analyzer.py - Gemini Vision 2.5 Flash + 4 Key Rotation)
[JSON Layout, OCR Tiếng Việt, Phân Tích Bố Cục & Ranh Giới Bài]
    │
    ▼ (Stage 3: sgk_splitter.py)
[Lesson Manifest: Danh sách bài học & dải trang tương ứng]
    │
    ▼ (Stage 4: sgk_docx_builder.py)
[File DOCX Chuẩn Cho Từng Bài Học (Text + Bảng + Box Ghi Nhớ + Ảnh + Phụ Lục)]
```

---

## 2. Rào Cản Kỹ Thuật & Giải Pháp Thiết Kế

1. **Chặn DLL C-Extension (Windows Application Control Policy):**
   - Không thể chạy `PaddleOCR` do hệ điều hành chặn load DLL của `chardet`, `pandas`.
   - **Giải pháp:** Sử dụng **Gemini 2.5 Flash Vision** làm engine phân tích hình ảnh và nhận diện chữ Việt chất lượng cao.

2. **Cơ chế Xoay Vòng API Key (4-Key Rotation):**
   - Đọc tự động các key từ `D:\AI\local-ai-agent\.env`.
   - Tự động bắt mã lỗi `429 (ResourceExhausted)` hoặc `503 (Unavailable)` để chuyển sang key tiếp theo và tự động retry tối đa 3 lần.

3. **Cơ chế Caching Thông Minh (Resumable Execution):**
   - Kết quả phân tích mỗi trang được lưu thành `page_XXX.json` trong thư mục `ocr_output/analysis/`.
   - Khi chạy lại, hệ thống tự động bỏ qua các trang đã có kết quả cache, không tốn thêm token hay thời gian.

---

## 3. Danh Sách Script Thực Thi (`scripts/`)

| Script | Chức năng | Lệnh mẫu |
|---|---|---|
| [`sgk_pipeline.py`](file:///d:/UNIGO/scripts/sgk_pipeline.py) | Điều phối CLI toàn bộ 4 stage | `python scripts/sgk_pipeline.py --grade 3` |
| [`sgk_renderer.py`](file:///d:/UNIGO/scripts/sgk_renderer.py) | Stage 1: Render 300 DPI & trích xuất ảnh | `python scripts/sgk_pipeline.py --grade 3 --stage render` |
| [`sgk_analyzer.py`](file:///d:/UNIGO/scripts/sgk_analyzer.py) | Stage 2: Gemini Vision Layout & OCR | `python scripts/sgk_pipeline.py --grade 3 --stage analyze` |
| [`sgk_splitter.py`](file:///d:/UNIGO/scripts/sgk_splitter.py) | Stage 3: Nhóm bài học & trang | `python scripts/sgk_pipeline.py --grade 3 --stage split` |
| [`sgk_docx_builder.py`](file:///d:/UNIGO/scripts/sgk_docx_builder.py) | Stage 4: Tạo file DOCX chuẩn từng bài | `python scripts/sgk_pipeline.py --grade 3 --stage build-docx` |
| [`sgk_batch_analyze.py`](file:///d:/UNIGO/scripts/sgk_batch_analyze.py) | Chạy tự động liên tục nhiều khối lớp (Stage 2→3→4) | `python scripts/sgk_batch_analyze.py --grades 3 4 5 6 7 8` |

---

## 4. Hướng Dẫn Chạy & Tiếp Tục Phiên Làm Việc

### Tiếp tục chạy cho tất cả các lớp:
```bash
# Chạy nối tiếp từ Lớp 3 đến Lớp 8
python scripts/sgk_batch_analyze.py --grades 3 4 5 6 7 8

# Hoặc chạy riêng từng lớp
python scripts/sgk_pipeline.py --grade 4
python scripts/sgk_pipeline.py --grade 5
python scripts/sgk_pipeline.py --grade 6
python scripts/sgk_pipeline.py --grade 7
python scripts/sgk_pipeline.py --grade 8
```

### Cấu trúc thư mục đầu ra:
```
D:\UNIGO\SGK\Lớp_[X]\ocr_output\
├── pages/                  # Toàn bộ trang đã render 300 DPI (PNG)
├── images/                 # Các hình ảnh minh họa vector/raster trích xuất từ PDF
├── analysis/               # File JSON phân tích layout từng trang + analysis_manifest.json
├── lesson_manifest.json    # Bảng phân phối bài học và dải trang SGK
└── docx/                   # File .docx từng bài học hoàn chỉnh kèm ảnh & phụ lục
```
