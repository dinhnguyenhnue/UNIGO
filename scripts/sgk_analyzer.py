"""
Stage 2: SGK Gemini Vision Analyzer
Gửi trang PDF đã render lên Gemini Vision API để phân tích layout, OCR tiếng Việt,
và trả về structured JSON cho mỗi trang.

Hỗ trợ:
- 4 API keys rotation
- Rate limiting (15 RPM free, 1000 RPM paid)
- Caching (skip trang đã phân tích)
- Retry logic

Usage:
    python scripts/sgk_analyzer.py --grade 3
    python scripts/sgk_analyzer.py --grade 3 --pages 6-12
"""
import sys, os, io, json, time, glob, argparse, base64, re
if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SGK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'SGK')

# ─── Load API Keys ───────────────────────────────────────────────
def load_api_keys():
    """Load Gemini API keys từ .env file hoặc environment."""
    keys = []
    
    # 1. Từ environment variable
    env_keys = os.environ.get('GEMINI_API_KEYS', '')
    if env_keys:
        keys = [k.strip() for k in env_keys.split(',') if k.strip()]
    
    # 2. Từ .env file
    if not keys:
        env_path = r'D:\AI\local-ai-agent\.env'
        if os.path.exists(env_path):
            with open(env_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('GEMINI_API_KEYS='):
                        val = line.split('=', 1)[1].strip()
                        keys = [k.strip() for k in val.split(',') if k.strip()]
                        break
    
    # 3. Fallback: single key
    if not keys:
        single = os.environ.get('GOOGLE_API_KEY', '') or os.environ.get('GEMINI_API_KEY', '')
        if single:
            keys = [single.strip()]
    
    if not keys:
        raise ValueError(
            "Không tìm thấy Gemini API key! Cài đặt GEMINI_API_KEYS trong .env "
            "hoặc GOOGLE_API_KEY trong environment."
        )
    
    print(f'  [API] Loaded {len(keys)} API key(s)')
    return keys


# ─── Gemini Client with Key Rotation ─────────────────────────────
class GeminiAnalyzer:
    """Gemini Vision API client với key rotation và rate limiting."""
    
    ANALYSIS_PROMPT = """Bạn là chuyên gia phân tích sách giáo khoa Tin học Việt Nam. 
Hãy phân tích trang SGK trong ảnh và trả về JSON cấu trúc.

## Yêu cầu phân tích:

1. **Nhận dạng loại trang**: title_page (trang bìa/đầu chương), toc (mục lục), content (nội dung bài), exercise (bài tập), review (ôn tập)

2. **Phát hiện ranh giới bài học**: Tìm các marker như "Bài 1", "BÀI 1", "Bài 2"... hoặc tiêu đề chủ đề lớn

3. **Trích xuất toàn bộ nội dung**: OCR chính xác từng phần tử theo thứ tự từ trên xuống dưới

4. **Phân loại vùng hình ảnh**: Xác định bounding box (tỷ lệ % so với kích thước trang) của các hình minh họa

## Trả về JSON đúng format sau (KHÔNG thêm markdown code block):

{
  "page_type": "content",
  "lesson_markers": [
    {"lesson_number": 1, "lesson_title": "Thông tin và xử lý thông tin", "position": "top"}
  ],
  "chapter_info": {"chapter_number": 1, "chapter_title": "Chủ đề A. Máy tính và cộng đồng"},
  "elements": [
    {
      "type": "heading",
      "level": 1,
      "text": "Bài 1. Thông tin và xử lý thông tin",
      "y_position": 0.05
    },
    {
      "type": "paragraph",
      "text": "Nội dung đoạn văn...",
      "y_position": 0.15
    },
    {
      "type": "image_region",
      "description": "Hình minh họa: Các em học sinh đang sử dụng máy tính",
      "bbox": {"x": 0.1, "y": 0.3, "w": 0.8, "h": 0.25},
      "y_position": 0.3
    },
    {
      "type": "note",
      "text": "Ghi nhớ: Thông tin là...",
      "style": "highlight_box",
      "y_position": 0.6
    },
    {
      "type": "exercise",
      "text": "Câu 1: Em hãy cho biết...",
      "exercise_number": 1,
      "y_position": 0.75
    },
    {
      "type": "table",
      "caption": "Bảng so sánh...",
      "headers": ["Cột 1", "Cột 2"],
      "rows": [["Giá trị 1", "Giá trị 2"]],
      "y_position": 0.8
    },
    {
      "type": "activity",
      "activity_type": "practice",
      "text": "Thực hành: Em hãy thực hiện...",
      "y_position": 0.85
    },
    {
      "type": "code_example",
      "language": "scratch",
      "text": "Mã lệnh minh họa...",
      "y_position": 0.9
    }
  ],
  "full_text": "Toàn bộ text trên trang, giữ nguyên thứ tự...",
  "page_summary": "Tóm tắt 1-2 câu nội dung trang"
}

## Lưu ý quan trọng:
- OCR chính xác tiếng Việt (dấu, chữ hoa/thường đúng)
- Giữ nguyên các thuật ngữ Tin học
- Nhận dạng đúng các phần: Khám phá, Thực hành, Luyện tập, Vận dụng, Ghi nhớ
- Xác định rõ hình ảnh minh họa với bounding box tương đối (0.0-1.0)
- Trả về JSON thuần, KHÔNG wrap trong ```json```
"""
    
    def __init__(self, api_keys, model='gemini-2.5-flash', rpm_limit=14):
        self.api_keys = api_keys
        self.current_key_idx = 0
        self.model = model
        self.rpm_limit = rpm_limit
        self.request_times = []
        self._init_client()
    
    def _init_client(self):
        """Khởi tạo Google GenAI client."""
        from google import genai
        self.client = genai.Client(api_key=self.api_keys[self.current_key_idx])
        print(f'  [API] Using key #{self.current_key_idx + 1}')
    
    def _rotate_key(self):
        """Xoay sang API key tiếp theo."""
        self.current_key_idx = (self.current_key_idx + 1) % len(self.api_keys)
        self._init_client()
        self.request_times = []  # Reset rate limit cho key mới
    
    def _wait_for_rate_limit(self):
        """Chờ nếu đạt rate limit."""
        now = time.time()
        # Xóa request cũ hơn 60s
        self.request_times = [t for t in self.request_times if now - t < 60]
        
        if len(self.request_times) >= self.rpm_limit:
            wait_time = 60 - (now - self.request_times[0]) + 1
            if wait_time > 0:
                print(f'    [Rate limit] Waiting {wait_time:.0f}s...')
                time.sleep(wait_time)
    
    def analyze_page(self, image_path, page_number, max_retries=3):
        """
        Phân tích 1 trang SGK bằng Gemini Vision.
        
        Args:
            image_path: Đường dẫn ảnh trang đã render
            page_number: Số trang
            max_retries: Số lần retry tối đa
        
        Returns:
            dict — JSON phân tích
        """
        from google import genai
        from google.genai import types
        
        # Đọc ảnh
        with open(image_path, 'rb') as f:
            image_bytes = f.read()
        
        for attempt in range(max_retries):
            try:
                self._wait_for_rate_limit()
                
                response = self.client.models.generate_content(
                    model=self.model,
                    contents=[
                        types.Content(
                            role='user',
                            parts=[
                                types.Part.from_bytes(
                                    data=image_bytes,
                                    mime_type='image/png'
                                ),
                                types.Part.from_text(
                                    text=f"Đây là trang {page_number} của sách giáo khoa Tin học.\n\n{self.ANALYSIS_PROMPT}"
                                )
                            ]
                        )
                    ],
                    config=types.GenerateContentConfig(
                        temperature=0.1,
                        max_output_tokens=8192,
                        response_mime_type='application/json'
                    )
                )
                
                self.request_times.append(time.time())
                
                # Parse response
                text = response.text.strip()
                
                # Clean up potential markdown wrapping
                if text.startswith('```'):
                    text = re.sub(r'^```(?:json)?\s*', '', text)
                    text = re.sub(r'\s*```$', '', text)
                
                result = json.loads(text)
                result['_meta'] = {
                    'page_number': page_number,
                    'image_path': image_path,
                    'api_key_index': self.current_key_idx,
                    'timestamp': time.strftime('%Y-%m-%dT%H:%M:%S'),
                    'model': self.model
                }
                
                return result
                
            except json.JSONDecodeError as e:
                print(f'    [!] Page {page_number}: JSON parse error (attempt {attempt+1}): {e}')
                if attempt < max_retries - 1:
                    time.sleep(2)
                    
            except Exception as e:
                error_str = str(e).lower()
                if 'rate' in error_str or '429' in error_str or 'quota' in error_str:
                    print(f'    [!] Rate limited on key #{self.current_key_idx+1}, rotating...')
                    self._rotate_key()
                    time.sleep(5)
                elif 'api key' in error_str or '401' in error_str or '403' in error_str:
                    print(f'    [!] Invalid API key #{self.current_key_idx+1}, rotating...')
                    self._rotate_key()
                else:
                    print(f'    [!] Page {page_number}: Error (attempt {attempt+1}): {e}')
                    if attempt < max_retries - 1:
                        time.sleep(3)
        
        # All retries failed
        return {
            'page_type': 'error',
            'elements': [],
            'full_text': '',
            'error': f'Failed after {max_retries} retries',
            '_meta': {'page_number': page_number, 'image_path': image_path}
        }


def process_grade(grade, page_range=None, force=False):
    """
    Chạy Stage 2 cho 1 khối lớp.
    
    Args:
        grade: Khối lớp (3-8)
        page_range: Tuple (start, end) 1-indexed
        force: Force re-analyze trang đã có cache
    """
    grade_dir = os.path.join(SGK_DIR, f'Lớp_{grade}', 'ocr_output')
    pages_dir = os.path.join(grade_dir, 'pages')
    analysis_dir = os.path.join(grade_dir, 'analysis')
    os.makedirs(analysis_dir, exist_ok=True)
    
    # Load manifest from Stage 1
    manifest_path = os.path.join(grade_dir, 'page_manifest.json')
    if not os.path.exists(manifest_path):
        print(f'  [!] Chưa có page_manifest.json — Chạy Stage 1 trước: python scripts/sgk_renderer.py --grade {grade}')
        return None
    
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    
    print(f'\n{"=" * 60}')
    print(f'  STAGE 2 — LỚP {grade}: Gemini Vision Analysis')
    print(f'{"=" * 60}')
    print(f'  Total rendered pages: {manifest["rendered_pages"]}')
    
    # Filter pages
    pages = manifest['pages']
    if page_range:
        pages = [p for p in pages if page_range[0] <= p['page_number'] <= page_range[1]]
    
    # Init analyzer
    api_keys = load_api_keys()
    analyzer = GeminiAnalyzer(api_keys)
    
    results = []
    skipped = 0
    
    for i, page_meta in enumerate(pages):
        page_num = page_meta['page_number']
        img_path = page_meta['filepath']
        
        # Check cache
        cache_path = os.path.join(analysis_dir, f'page_{page_num:03d}.json')
        if os.path.exists(cache_path) and not force:
            skipped += 1
            with open(cache_path, 'r', encoding='utf-8') as f:
                results.append(json.load(f))
            continue
        
        # Check if image exists
        if not os.path.exists(img_path):
            print(f'    [!] Image not found: {img_path}')
            continue
        
        print(f'  [{i+1}/{len(pages)}] Analyzing page {page_num}...', end='', flush=True)
        
        result = analyzer.analyze_page(img_path, page_num)
        
        # Save to cache
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        results.append(result)
        
        page_type = result.get('page_type', '?')
        markers = result.get('lesson_markers', [])
        n_elements = len(result.get('elements', []))
        marker_info = ""
        if markers:
            # Handle both structures: {lesson_number: N} and {type: 'lesson', lesson_number: N}
            lesson_markers = [m for m in markers if m.get('lesson_number') is not None]
            if lesson_markers:
                marker_info = f" — Bài {lesson_markers[0]['lesson_number']}"
        print(f' ✓ [{page_type}] {n_elements} elements{marker_info}')
    
    if skipped:
        print(f'\n  [Cache] Skipped {skipped} already-analyzed pages')
    
    # Save analysis manifest
    analysis_manifest = {
        'grade': grade,
        'total_analyzed': len(results),
        'pages_analyzed': [r.get('_meta', {}).get('page_number', 0) for r in results],
        'lesson_markers_found': [],
        'page_type_summary': {}
    }
    
    # Aggregate lesson markers
    for r in results:
        for marker in r.get('lesson_markers', []):
            analysis_manifest['lesson_markers_found'].append({
                'lesson_number': marker.get('lesson_number'),
                'lesson_title': marker.get('lesson_title', ''),
                'page_number': r.get('_meta', {}).get('page_number', 0)
            })
        
        ptype = r.get('page_type', 'unknown')
        analysis_manifest['page_type_summary'][ptype] = \
            analysis_manifest['page_type_summary'].get(ptype, 0) + 1
    
    manifest_out = os.path.join(analysis_dir, 'analysis_manifest.json')
    with open(manifest_out, 'w', encoding='utf-8') as f:
        json.dump(analysis_manifest, f, ensure_ascii=False, indent=2)
    
    print(f'\n  ✓ Analysis manifest: {manifest_out}')
    print(f'  ✓ Lesson markers found: {len(analysis_manifest["lesson_markers_found"])}')
    print(f'  ✓ Page types: {analysis_manifest["page_type_summary"]}')
    
    return analysis_manifest


def parse_page_range(s):
    if '-' in s:
        parts = s.split('-')
        return (int(parts[0]), int(parts[1]))
    else:
        n = int(s)
        return (n, n)


def main():
    parser = argparse.ArgumentParser(description='SGK Gemini Vision Analyzer — Stage 2')
    parser.add_argument('--grade', type=int, choices=[3, 4, 5, 6, 7, 8],
                        help='Khối lớp cần phân tích')
    parser.add_argument('--all', action='store_true',
                        help='Phân tích tất cả các lớp')
    parser.add_argument('--pages', type=str, default=None,
                        help='Phạm vi trang (VD: 6-12)')
    parser.add_argument('--force', action='store_true',
                        help='Force re-analyze (bỏ qua cache)')
    parser.add_argument('--model', type=str, default='gemini-2.5-flash',
                        help='Gemini model (default: gemini-2.5-flash)')
    
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
        process_grade(grade, page_range, args.force)
    
    print(f'\n{"=" * 60}')
    print('  STAGE 2 COMPLETE!')
    print(f'{"=" * 60}')


if __name__ == '__main__':
    main()
