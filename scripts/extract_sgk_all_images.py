"""
Extract ALL images from SGK Tin học Bài 1 for grades 3-8.
- Full page renders (200 DPI)
- Individual embedded images (filtered: skip icons < 100x100)
"""
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')

import fitz  # PyMuPDF

SGK_DIR = r'd:\UNIGO\SGK'

# Bài 1 page ranges per grade (0-indexed, from inspection)
# These are approximate — the script also auto-detects
BAI1_MANUAL_RANGES = {
    3: (5, 8),   # pages 6-9 in the PDF (0-indexed: 5-8)
    4: (5, 12),
    5: (5, 12),
    6: (5, 12),
    7: (5, 12),
    8: (5, 12),
}


def find_bai1_pages(doc, grade):
    """Auto-detect Bài 1 page range by searching text."""
    bai1_pages = []
    found_bai1 = False

    for i in range(min(30, len(doc))):
        text = doc[i].get_text()

        if 'Bài 1' in text or 'BÀI 1' in text:
            if not found_bai1:
                if any(kw in text for kw in ['Mục lục', 'MỤC LỤC']):
                    continue
                found_bai1 = True
            bai1_pages.append(i)
        elif found_bai1:
            if 'Bài 2' in text or 'BÀI 2' in text:
                break
            bai1_pages.append(i)
            if len(bai1_pages) >= 8:
                break

    if not bai1_pages:
        # Fallback to manual range
        start, end = BAI1_MANUAL_RANGES.get(grade, (4, 12))
        bai1_pages = list(range(start, min(end + 1, len(doc))))

    return bai1_pages


def extract_full_pages(doc, pages, out_dir):
    """Render full pages at 200 DPI."""
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    for page_idx in pages:
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=200)
        fname = f'bai1_page{page_idx + 1}.png'
        out_path = os.path.join(out_dir, fname)
        pix.save(out_path)
        saved.append(out_path)
        print(f'    Full page: {fname} ({pix.width}x{pix.height})')
    return saved


def extract_individual_images(doc, pages, out_dir):
    """Extract individual embedded images from PDF pages."""
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    
    for page_idx in pages:
        page = doc[page_idx]
        image_list = page.get_images(full=True)

        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue

            img_bytes = base_image["image"]
            img_ext = base_image["ext"]
            width = base_image.get("width", 0)
            height = base_image.get("height", 0)

            # Skip tiny images (icons, bullets, decorations)
            if width < 100 or height < 100:
                continue

            # Skip very narrow images (likely borders/lines)
            if width < 50 or height < 50:
                continue
            
            # Skip images that are too small to be useful illustrations
            if width * height < 15000:
                continue

            fname = f'bai1_p{page_idx + 1}_img{img_idx + 1}.{img_ext}'
            out_path = os.path.join(out_dir, fname)
            with open(out_path, 'wb') as f:
                f.write(img_bytes)
            saved.append(out_path)
            print(f'    Image: {fname} ({width}x{height})')

    return saved


def main():
    print('=' * 60)
    print('  EXTRACT SGK IMAGES — BÀI 1 TIN HỌC — LỚP 3-8')
    print('=' * 60)

    for grade in [3, 4, 5, 6, 7, 8]:
        folder = os.path.join(SGK_DIR, f'Lớp_{grade}')
        pdfs = glob.glob(os.path.join(folder, '*.pdf'))
        if not pdfs:
            print(f'\n=== LỚP {grade}: Không tìm thấy PDF ===')
            continue

        pdf_path = pdfs[0]
        print(f'\n{"=" * 50}')
        print(f'  LỚP {grade}: {os.path.basename(pdf_path)}')
        print(f'{"=" * 50}')

        doc = fitz.open(pdf_path)
        print(f'  Tổng trang PDF: {len(doc)}')

        # Find Bài 1 pages
        bai1_pages = find_bai1_pages(doc, grade)
        print(f'  Bài 1 pages: {[p + 1 for p in bai1_pages]}')

        base_dir = os.path.join(folder, 'bai1_images')

        # 1. Full page renders
        print(f'\n  --- Full Pages ---')
        full_dir = os.path.join(base_dir, 'full_pages')
        full_saved = extract_full_pages(doc, bai1_pages, full_dir)

        # 2. Individual embedded images
        print(f'\n  --- Individual Images ---')
        img_saved = extract_individual_images(doc, bai1_pages, base_dir)

        doc.close()

        print(f'\n  Tổng kết Lớp {grade}: {len(full_saved)} full pages, {len(img_saved)} images')

    print(f'\n{"=" * 60}')
    print('  HOÀN THÀNH!')
    print(f'{"=" * 60}')


if __name__ == '__main__':
    main()
