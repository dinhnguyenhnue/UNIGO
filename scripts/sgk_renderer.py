"""
Stage 1: SGK PDF Renderer — PyMuPDF
Render trang PDF ra PNG (300 DPI) và extract hình ảnh minh hoạ nhúng.

Usage:
    python scripts/sgk_renderer.py --grade 3
    python scripts/sgk_renderer.py --grade 3 --pages 6-12
    python scripts/sgk_renderer.py --all
"""
import sys, os, io, json, glob, argparse
if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import fitz  # PyMuPDF

SGK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'SGK')


def find_pdf(grade):
    """Tìm file PDF SGK cho khối lớp chỉ định."""
    folder = os.path.join(SGK_DIR, f'Lớp_{grade}')
    pdfs = glob.glob(os.path.join(folder, '*.pdf'))
    if not pdfs:
        print(f'  [!] Không tìm thấy PDF trong {folder}')
        return None
    return pdfs[0]


def render_pages(pdf_path, output_dir, page_range=None, dpi=300):
    """
    Render các trang PDF ra PNG.
    
    Args:
        pdf_path: Đường dẫn file PDF
        output_dir: Thư mục lưu ảnh
        page_range: Tuple (start, end) 1-indexed inclusive, hoặc None = tất cả
        dpi: Độ phân giải render (default 300)
    
    Returns:
        List[dict] — metadata mỗi trang đã render
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    
    if page_range:
        start_idx = max(0, page_range[0] - 1)  # convert to 0-indexed
        end_idx = min(total_pages, page_range[1])
        pages_to_render = range(start_idx, end_idx)
    else:
        pages_to_render = range(total_pages)
    
    results = []
    print(f'  [Render] {len(list(pages_to_render))} trang @ {dpi} DPI')
    
    # Re-create range since we consumed it
    if page_range:
        start_idx = max(0, page_range[0] - 1)
        end_idx = min(total_pages, page_range[1])
        pages_to_render = range(start_idx, end_idx)
    else:
        pages_to_render = range(total_pages)
    
    for page_idx in pages_to_render:
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=dpi)
        fname = f'page_{page_idx + 1:03d}.png'
        out_path = os.path.join(output_dir, fname)
        pix.save(out_path)
        
        # Extract raw text (PyMuPDF text layer — may be partial for scanned PDFs)
        raw_text = page.get_text('text').strip()
        
        page_meta = {
            'page_number': page_idx + 1,
            'filename': fname,
            'filepath': out_path,
            'width': pix.width,
            'height': pix.height,
            'has_text_layer': len(raw_text) > 20,
            'raw_text_preview': raw_text[:200] if raw_text else '',
            'file_size': os.path.getsize(out_path)
        }
        results.append(page_meta)
        
        if (page_idx + 1) % 10 == 0 or page_idx == list(pages_to_render)[-1]:
            print(f'    Rendered: page {page_idx + 1} ({pix.width}x{pix.height}, {page_meta["file_size"]//1024}KB)')
    
    doc.close()
    return results


def extract_embedded_images(pdf_path, output_dir, page_range=None, min_size=100, min_area=15000):
    """
    Extract hình ảnh nhúng trong PDF (illustrations, diagrams, screenshots).
    
    Args:
        pdf_path: Đường dẫn file PDF
        output_dir: Thư mục lưu ảnh
        page_range: Tuple (start, end) 1-indexed inclusive
        min_size: Kích thước tối thiểu width/height (pixels)
        min_area: Diện tích tối thiểu (pixels²)
    
    Returns:
        List[dict] — metadata mỗi ảnh extracted
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    
    if page_range:
        start_idx = max(0, page_range[0] - 1)
        end_idx = min(total_pages, page_range[1])
        pages_to_process = range(start_idx, end_idx)
    else:
        pages_to_process = range(total_pages)
    
    results = []
    seen_xrefs = set()  # Tránh extract trùng ảnh
    
    for page_idx in pages_to_process:
        page = doc[page_idx]
        image_list = page.get_images(full=True)
        
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            
            # Skip đã extract
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            
            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue
            
            img_bytes = base_image['image']
            img_ext = base_image['ext']
            width = base_image.get('width', 0)
            height = base_image.get('height', 0)
            
            # Filter: skip icons, borders, decorations
            if width < min_size or height < min_size:
                continue
            if width * height < min_area:
                continue
            
            fname = f'p{page_idx + 1:03d}_img{img_idx + 1:02d}_{width}x{height}.{img_ext}'
            out_path = os.path.join(output_dir, fname)
            
            with open(out_path, 'wb') as f:
                f.write(img_bytes)
            
            img_meta = {
                'page_number': page_idx + 1,
                'image_index': img_idx + 1,
                'filename': fname,
                'filepath': out_path,
                'width': width,
                'height': height,
                'format': img_ext,
                'xref': xref,
                'file_size': len(img_bytes)
            }
            results.append(img_meta)
    
    doc.close()
    print(f'  [Images] Extracted {len(results)} hình ảnh minh hoạ')
    return results


def process_grade(grade, page_range=None, dpi=300):
    """
    Xử lý toàn bộ Stage 1 cho 1 khối lớp.
    
    Returns:
        dict — manifest chứa metadata pages + images
    """
    pdf_path = find_pdf(grade)
    if not pdf_path:
        return None
    
    grade_dir = os.path.join(SGK_DIR, f'Lớp_{grade}', 'ocr_output')
    pages_dir = os.path.join(grade_dir, 'pages')
    images_dir = os.path.join(grade_dir, 'images')
    
    print(f'\n{"=" * 60}')
    print(f'  STAGE 1 — LỚP {grade}: {os.path.basename(pdf_path)}')
    print(f'{"=" * 60}')
    
    # Get total page count
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    doc.close()
    print(f'  Tổng trang PDF: {total_pages}')
    
    # Render pages
    page_results = render_pages(pdf_path, pages_dir, page_range, dpi)
    
    # Extract embedded images
    image_results = extract_embedded_images(pdf_path, images_dir, page_range)
    
    # Build manifest
    manifest = {
        'grade': grade,
        'pdf_path': pdf_path,
        'pdf_filename': os.path.basename(pdf_path),
        'total_pages': total_pages,
        'rendered_pages': len(page_results),
        'extracted_images': len(image_results),
        'dpi': dpi,
        'page_range': list(page_range) if page_range else None,
        'pages': page_results,
        'images': image_results
    }
    
    # Save manifest
    manifest_path = os.path.join(grade_dir, 'page_manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    
    print(f'\n  ✓ Manifest: {manifest_path}')
    print(f'  ✓ {len(page_results)} pages rendered, {len(image_results)} images extracted')
    
    return manifest


def parse_page_range(s):
    """Parse '6-12' → (6, 12), '5' → (5, 5)"""
    if '-' in s:
        parts = s.split('-')
        return (int(parts[0]), int(parts[1]))
    else:
        n = int(s)
        return (n, n)


def main():
    parser = argparse.ArgumentParser(description='SGK PDF Renderer — Stage 1')
    parser.add_argument('--grade', type=int, choices=[3, 4, 5, 6, 7, 8],
                        help='Khối lớp cần xử lý')
    parser.add_argument('--all', action='store_true',
                        help='Xử lý tất cả các lớp')
    parser.add_argument('--pages', type=str, default=None,
                        help='Phạm vi trang (VD: 6-12 hoặc 5)')
    parser.add_argument('--dpi', type=int, default=300,
                        help='Độ phân giải render (default: 300)')
    
    args = parser.parse_args()
    
    page_range = parse_page_range(args.pages) if args.pages else None
    
    if args.all:
        grades = [3, 4, 5, 6, 7, 8]
    elif args.grade:
        grades = [args.grade]
    else:
        parser.print_help()
        return
    
    for grade in grades:
        manifest = process_grade(grade, page_range, args.dpi)
        if manifest:
            print(f'\n  LỚP {grade}: HOÀN THÀNH')
    
    print(f'\n{"=" * 60}')
    print('  STAGE 1 COMPLETE!')
    print(f'{"=" * 60}')


if __name__ == '__main__':
    main()
