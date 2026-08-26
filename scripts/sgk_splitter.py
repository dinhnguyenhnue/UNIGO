"""
Stage 3: SGK Lesson Splitter
Phát hiện ranh giới bài học từ kết quả phân tích Gemini Vision (Stage 2),
nhóm các trang theo từng bài.

Usage:
    python scripts/sgk_splitter.py --grade 3
    python scripts/sgk_splitter.py --all
"""
import sys, os, io, json, glob, argparse, re
if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SGK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'SGK')


def load_analysis_files(grade):
    """Load tất cả JSON phân tích cho 1 khối lớp."""
    analysis_dir = os.path.join(SGK_DIR, f'Lớp_{grade}', 'ocr_output', 'analysis')
    
    if not os.path.exists(analysis_dir):
        print(f'  [!] Thư mục analysis không tồn tại: {analysis_dir}')
        return []
    
    files = sorted(glob.glob(os.path.join(analysis_dir, 'page_*.json')))
    
    results = []
    for fpath in files:
        with open(fpath, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
                results.append(data)
            except json.JSONDecodeError:
                print(f'    [!] Invalid JSON: {fpath}')
    
    return results


def detect_lesson_boundaries(pages_data):
    """
    Phát hiện ranh giới bài học từ dữ liệu phân tích.
    
    Strategies:
    1. lesson_markers từ Gemini
    2. Tìm pattern "Bài [N]" trong elements heading
    3. Tìm pattern trong full_text
    
    Returns:
        List[dict] — mỗi bài gồm lesson_number, title, start_page, end_page, pages
    """
    # Collect all lesson markers
    markers = []
    
    for page_data in pages_data:
        page_num = page_data.get('_meta', {}).get('page_number', 0)
        page_type = page_data.get('page_type', '')
        
        # Skip TOC, title pages
        if page_type in ('toc', 'title_page'):
            continue
        
        # Strategy 1: From Gemini's lesson_markers
        for marker in page_data.get('lesson_markers', []):
            lesson_num = marker.get('lesson_number')
            lesson_title = marker.get('lesson_title', '')
            if lesson_num is not None:
                markers.append({
                    'lesson_number': lesson_num,
                    'lesson_title': lesson_title,
                    'page_number': page_num,
                    'source': 'gemini_marker'
                })
        
        # Strategy 2: From heading elements
        for elem in page_data.get('elements', []):
            if elem.get('type') == 'heading' and elem.get('level', 99) <= 1:
                text = elem.get('text', '')
                match = re.match(r'(?:BÀI|Bài)\s+(\d+)[.:]\s*(.*)', text, re.IGNORECASE)
                if match:
                    lesson_num = int(match.group(1))
                    lesson_title = match.group(2).strip()
                    # Avoid duplicates
                    if not any(m['lesson_number'] == lesson_num and m['page_number'] == page_num 
                              for m in markers):
                        markers.append({
                            'lesson_number': lesson_num,
                            'lesson_title': lesson_title,
                            'page_number': page_num,
                            'source': 'heading_element'
                        })
        
        # Strategy 3: From full_text
        full_text = page_data.get('full_text', '')
        if full_text:
            # Look for "Bài N." or "BÀI N:" patterns at line start
            for match in re.finditer(
                r'^(?:BÀI|Bài)\s+(\d+)[.:]\s*(.+?)$', full_text, re.MULTILINE | re.IGNORECASE
            ):
                lesson_num = int(match.group(1))
                lesson_title = match.group(2).strip()[:100]
                if not any(m['lesson_number'] == lesson_num and m['page_number'] == page_num
                           for m in markers):
                    markers.append({
                        'lesson_number': lesson_num,
                        'lesson_title': lesson_title,
                        'page_number': page_num,
                        'source': 'full_text_regex'
                    })
    
    # Deduplicate: keep earliest page for each lesson
    lesson_starts = {}
    for m in markers:
        num = m['lesson_number']
        if num not in lesson_starts or m['page_number'] < lesson_starts[num]['page_number']:
            lesson_starts[num] = m
    
    # Sort by lesson number
    sorted_lessons = sorted(lesson_starts.values(), key=lambda x: x['lesson_number'])
    
    return sorted_lessons


def build_lesson_manifest(grade, pages_data, lesson_starts):
    """
    Xây dựng lesson manifest: nhóm trang vào từng bài.
    
    Returns:
        dict — lesson manifest
    """
    # Get all available page numbers
    all_page_nums = sorted([
        p.get('_meta', {}).get('page_number', 0) for p in pages_data
        if p.get('page_type', '') not in ('toc', 'title_page', 'error')
    ])
    
    lessons = []
    
    for i, marker in enumerate(lesson_starts):
        start_page = marker['page_number']
        
        # End page: start of next lesson - 1, or last page
        if i + 1 < len(lesson_starts):
            end_page = lesson_starts[i + 1]['page_number'] - 1
        else:
            end_page = max(all_page_nums) if all_page_nums else start_page
        
        # Collect pages for this lesson
        lesson_pages = [p for p in all_page_nums if start_page <= p <= end_page]
        
        # Collect all elements for this lesson
        lesson_elements = []
        for page_data in pages_data:
            pnum = page_data.get('_meta', {}).get('page_number', 0)
            if pnum in lesson_pages:
                for elem in page_data.get('elements', []):
                    elem_copy = dict(elem)
                    elem_copy['_source_page'] = pnum
                    lesson_elements.append(elem_copy)
        
        # Collect images for this lesson
        lesson_images = []
        for page_data in pages_data:
            pnum = page_data.get('_meta', {}).get('page_number', 0)
            if pnum in lesson_pages:
                for elem in page_data.get('elements', []):
                    if elem.get('type') == 'image_region':
                        lesson_images.append({
                            'page_number': pnum,
                            'description': elem.get('description', ''),
                            'bbox': elem.get('bbox', {})
                        })
        
        # Build full text for the lesson
        full_text_parts = []
        for page_data in pages_data:
            pnum = page_data.get('_meta', {}).get('page_number', 0)
            if pnum in lesson_pages:
                ft = page_data.get('full_text', '')
                if ft:
                    full_text_parts.append(f'--- Trang {pnum} ---\n{ft}')
        
        lesson = {
            'lesson_number': marker['lesson_number'],
            'lesson_title': marker['lesson_title'],
            'start_page': start_page,
            'end_page': end_page,
            'pages': lesson_pages,
            'page_count': len(lesson_pages),
            'element_count': len(lesson_elements),
            'image_count': len(lesson_images),
            'images': lesson_images,
            'detection_source': marker['source'],
            'full_text': '\n\n'.join(full_text_parts)
        }
        lessons.append(lesson)
    
    # Identify pre-lesson pages (before Bài 1: cover, TOC, intro)
    first_lesson_page = lesson_starts[0]['page_number'] if lesson_starts else 999
    pre_lesson_pages = [p for p in all_page_nums if p < first_lesson_page]
    
    manifest = {
        'grade': grade,
        'total_pages_analyzed': len(pages_data),
        'total_lessons_found': len(lessons),
        'pre_lesson_pages': pre_lesson_pages,
        'lessons': lessons,
        'lesson_summary': [
            {
                'number': l['lesson_number'],
                'title': l['lesson_title'],
                'pages': f"{l['start_page']}-{l['end_page']}",
                'page_count': l['page_count']
            }
            for l in lessons
        ]
    }
    
    return manifest


def process_grade(grade):
    """Xử lý Stage 3 cho 1 khối lớp."""
    print(f'\n{"=" * 60}')
    print(f'  STAGE 3 — LỚP {grade}: Lesson Splitter')
    print(f'{"=" * 60}')
    
    # Load analysis data
    pages_data = load_analysis_files(grade)
    if not pages_data:
        print(f'  [!] Không có dữ liệu phân tích. Chạy Stage 2 trước.')
        return None
    
    print(f'  Loaded {len(pages_data)} analyzed pages')
    
    # Detect lesson boundaries
    lesson_starts = detect_lesson_boundaries(pages_data)
    print(f'\n  Detected {len(lesson_starts)} lesson boundaries:')
    for ls in lesson_starts:
        print(f'    Bài {ls["lesson_number"]}: "{ls["lesson_title"]}" — trang {ls["page_number"]} [{ls["source"]}]')
    
    # Build manifest
    manifest = build_lesson_manifest(grade, pages_data, lesson_starts)
    
    # Save manifest
    grade_dir = os.path.join(SGK_DIR, f'Lớp_{grade}', 'ocr_output')
    manifest_path = os.path.join(grade_dir, 'lesson_manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    
    print(f'\n  ✓ Lesson manifest: {manifest_path}')
    print(f'\n  Tổng kết Lớp {grade}:')
    print(f'  {"─" * 50}')
    print(f'  {"Bài":<6} {"Tên bài":<35} {"Trang":<10} {"Số trang":<8}')
    print(f'  {"─" * 50}')
    for ls in manifest['lesson_summary']:
        title = ls['title'][:33] + '..' if len(ls['title']) > 35 else ls['title']
        print(f'  {ls["number"]:<6} {title:<35} {ls["pages"]:<10} {ls["page_count"]:<8}')
    
    return manifest


def main():
    parser = argparse.ArgumentParser(description='SGK Lesson Splitter — Stage 3')
    parser.add_argument('--grade', type=int, choices=[3, 4, 5, 6, 7, 8],
                        help='Khối lớp cần tách bài')
    parser.add_argument('--all', action='store_true',
                        help='Tách bài tất cả các lớp')
    
    args = parser.parse_args()
    
    if args.all:
        grades = [3, 4, 5, 6, 7, 8]
    elif args.grade:
        grades = [args.grade]
    else:
        parser.print_help()
        return
    
    for grade in grades:
        process_grade(grade)
    
    print(f'\n{"=" * 60}')
    print('  STAGE 3 COMPLETE!')
    print(f'{"=" * 60}')


if __name__ == '__main__':
    main()
