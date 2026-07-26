"""
SGK OCR & Document Parsing Module for UNIGO Pipeline
Sử dụng PaddleOCR / PP-Structure để trích xuất văn bản, bảng biểu và hình ảnh từ SGK PDF/Scan.
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import fitz  # PyMuPDF

def extract_page_images(pdf_path, output_dir, dpi=200):
    """Render PDF pages as images for OCR processing."""
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    image_paths = []
    print(f"[SGK OCR] Opening {pdf_path} ({doc.page_count} pages)")
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        img_path = os.path.join(output_dir, f"page_{i+1:03d}.png")
        pix.save(img_path)
        image_paths.append(img_path)
    doc.close()
    return image_paths

def process_sgk_with_paddleocr(pdf_path, output_dir):
    """
    Xử lý SGK PDF với PaddleOCR:
    1. Render các trang PDF ra ảnh
    2. Sử dụng PaddleOCR / PP-Structure để OCR và bóc tách bố cục
    3. Lưu nội dung bài viết ra Markdown và các hình ảnh minh họa đã crop.
    """
    os.makedirs(output_dir, exist_ok=True)
    img_dir = os.path.join(output_dir, "pages")
    res_dir = os.path.join(output_dir, "extracted_assets")
    os.makedirs(res_dir, exist_ok=True)
    
    image_paths = extract_page_images(pdf_path, img_dir)
    
    # Try importing paddleocr
    try:
        from paddleocr import PaddleOCR, PPStructure
        print("[SGK OCR] Initializing PaddleOCR (lang='vi')...")
        ocr_engine = PaddleOCR(use_angle_cls=True, lang='vi', show_log=False)
        table_engine = PPStructure(show_log=False, lang='vi')
        
        all_results = []
        for idx, img_path in enumerate(image_paths):
            print(f"[SGK OCR] Processing page {idx+1}/{len(image_paths)}...")
            # Text recognition
            ocr_res = ocr_engine.ocr(img_path, cls=True)
            page_text_lines = []
            if ocr_res and ocr_res[0]:
                for line in ocr_res[0]:
                    text = line[1][0]
                    score = line[1][1]
                    page_text_lines.append(text)
            
            # Layout & table analysis
            table_res = table_engine(img_path)
            
            all_results.append({
                "page": idx + 1,
                "text_lines": page_text_lines,
                "full_text": "\n".join(page_text_lines),
                "structure": str(table_res)
            })
            
        # Write extracted OCR output to json
        out_json = os.path.join(output_dir, "sgk_ocr_result.json")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(all_results, f, ensure_ascii=False, indent=2)
            
        print(f"[SGK OCR] Complete! Result saved to {out_json}")
        return out_json

    except ImportError as e:
        print(f"[SGK OCR Warning] PaddleOCR not fully available yet: {e}")
        print("[SGK OCR] Fallback to PyMuPDF image rendering mode.")
        return None

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf = sys.argv[1]
        out = sys.argv[2] if len(sys.argv) > 2 else "d:\\UNIGO\\SGK\\ocr_output"
        process_sgk_with_paddleocr(pdf, out)
    else:
        print("Usage: python sgk_ocr.py <path_to_sgk.pdf> [output_dir]")
