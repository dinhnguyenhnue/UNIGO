"""
Batch process all grades: Stage 2 → 3 → 4
Dùng sau khi Stage 1 (render) đã xong cho tất cả lớp.

Usage:
    python scripts/sgk_batch_analyze.py
    python scripts/sgk_batch_analyze.py --grades 4 5 6
"""
import sys, os, io, time, argparse
if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from sgk_analyzer import process_grade as analyze_grade
from sgk_splitter import process_grade as split_grade
from sgk_docx_builder import process_grade as build_grade


def batch_process(grades, force=False):
    total_start = time.time()
    results = {}
    
    for grade in grades:
        grade_start = time.time()
        print(f'\n{"█" * 60}')
        print(f'  PROCESSING LỚP {grade}')
        print(f'{"█" * 60}')
        
        # Stage 2: Analyze
        print(f'\n  ▶ Stage 2: Gemini Vision Analysis...')
        try:
            analyze_grade(grade, force=force)
        except Exception as e:
            print(f'  [!] Stage 2 error: {e}')
            results[grade] = {'status': 'FAILED', 'error': str(e)}
            continue
        
        # Stage 3: Split
        print(f'\n  ▶ Stage 3: Lesson Splitting...')
        try:
            split_grade(grade)
        except Exception as e:
            print(f'  [!] Stage 3 error: {e}')
            results[grade] = {'status': 'PARTIAL', 'error': str(e)}
            continue
        
        # Stage 4: Build DOCX
        print(f'\n  ▶ Stage 4: Building DOCX...')
        try:
            build_result = build_grade(grade)
            elapsed = time.time() - grade_start
            results[grade] = {
                'status': 'OK',
                'docx_count': build_result['total_docx_created'] if build_result else 0,
                'elapsed': f'{elapsed:.1f}s'
            }
        except Exception as e:
            print(f'  [!] Stage 4 error: {e}')
            results[grade] = {'status': 'PARTIAL', 'error': str(e)}
    
    total_elapsed = time.time() - total_start
    
    print(f'\n{"█" * 60}')
    print(f'  BATCH PROCESSING COMPLETE ({total_elapsed:.1f}s)')
    print(f'{"█" * 60}')
    
    for grade, info in results.items():
        status = info['status']
        if status == 'OK':
            print(f'  Lớp {grade}: ✓ {info["docx_count"]} DOCX ({info["elapsed"]})')
        else:
            print(f'  Lớp {grade}: ✗ {status} — {info.get("error", "")}')
    
    return results


def main():
    parser = argparse.ArgumentParser(description='Batch analyze & build DOCX for all grades')
    parser.add_argument('--grades', type=int, nargs='+', default=[3, 4, 5, 6, 7, 8],
                        help='Grades to process (default: 3 4 5 6 7 8)')
    parser.add_argument('--force', action='store_true',
                        help='Force re-analyze (bypass cache)')
    args = parser.parse_args()
    
    batch_process(args.grades, args.force)


if __name__ == '__main__':
    main()
