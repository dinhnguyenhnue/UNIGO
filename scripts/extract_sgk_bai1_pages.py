"""Extract Bài 1 full-page images from SGK Tin học for all grades (3-8)."""
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')

import fitz  # PyMuPDF

SGK_DIR = r'd:\UNIGO\SGK'

for grade in [3, 4, 5, 6, 7, 8]:
    folder = os.path.join(SGK_DIR, f'Lớp_{grade}')
    pdfs = glob.glob(os.path.join(folder, '*.pdf'))
    if not pdfs:
        print(f'=== LỚP {grade}: Không tìm thấy PDF ===\n')
        continue

    pdf_path = pdfs[0]
    print(f'=== LỚP {grade}: {os.path.basename(pdf_path)} ===')
    doc = fitz.open(pdf_path)
    print(f'    Tổng: {len(doc)} trang')

    # Create output directory
    out_dir = os.path.join(folder, 'bai1_images', 'full_pages')
    os.makedirs(out_dir, exist_ok=True)

    # Find Bài 1 pages by searching text
    bai1_pages = []
    found_bai1 = False
    
    for i in range(min(30, len(doc))):
        text = doc[i].get_text()
        
        # Check if this page has Bài 1 or is part of Bài 1 content
        if 'Bài 1' in text or 'BÀI 1' in text:
            if not found_bai1:
                # Check if this is the actual Bài 1 title page (not TOC)
                if any(kw in text for kw in ['Mục lục', 'MỤC LỤC']):
                    continue
                found_bai1 = True
                print(f'  Bài 1 bắt đầu: trang {i+1}')
            bai1_pages.append(i)
        elif found_bai1:
            # Check if Bài 2 starts
            if 'Bài 2' in text or 'BÀI 2' in text:
                # Still include this page if it has both bài 1 content ending
                break
            # Continue collecting pages that might still be Bài 1
            bai1_pages.append(i)
            # Stop after 6 pages max
            if len(bai1_pages) >= 8:
                break

    if not bai1_pages:
        # Fallback: just extract pages 5-12 (common Bài 1 location)
        print(f'  Không tìm thấy "Bài 1" trong text, sẽ dùng OCR hoặc cố định trang')
        # Try OCR approach or manual pages
        bai1_pages = list(range(4, min(12, len(doc))))

    print(f'  Trang trích xuất: {[p+1 for p in bai1_pages]}')
    
    for page_idx in bai1_pages:
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=200)
        out_path = os.path.join(out_dir, f'bai1_page{page_idx+1}.png')
        pix.save(out_path)
        print(f'  Đã lưu: bai1_page{page_idx+1}.png ({pix.width}x{pix.height})')

    doc.close()
    print()

print("HOÀN THÀNH!")
