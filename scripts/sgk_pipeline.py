"""
SGK Pipeline Orchestrator
Điều phối toàn bộ 4 stage: Render → Analyze → Split → Build DOCX

Usage:
    # Xử lý toàn bộ 1 lớp
    python scripts/sgk_pipeline.py --grade 3

    # Xử lý phạm vi trang cụ thể
    python scripts/sgk_pipeline.py --grade 3 --pages 6-12

    # Xử lý bài cụ thể (stage 4)
    python scripts/sgk_pipeline.py --grade 3 --lessons 1-5

    # Xử lý tất cả các lớp
    python scripts/sgk_pipeline.py --all

    # Chạy 1 stage cụ thể
    python scripts/sgk_pipeline.py --grade 3 --stage render
    python scripts/sgk_pipeline.py --grade 3 --stage analyze
    python scripts/sgk_pipeline.py --grade 3 --stage split
    python scripts/sgk_pipeline.py --grade 3 --stage build-docx

    # Force re-analyze (bỏ qua cache)
    python scripts/sgk_pipeline.py --grade 3 --force
"""
import sys, os, io, time, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Add scripts dir to path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from sgk_renderer import process_grade as render_grade, parse_page_range
from sgk_analyzer import process_grade as analyze_grade
from sgk_splitter import process_grade as split_grade
from sgk_docx_builder import process_grade as build_grade, parse_lesson_range


def run_pipeline(grade, stage='all', page_range=None, lesson_range=None, force=False, dpi=300):
    """
    Chạy pipeline cho 1 khối lớp.
    
    Args:
        grade: Khối lớp (3-8)
        stage: 'all', 'render', 'analyze', 'split', 'build-docx'
        page_range: Tuple (start, end) cho Stage 1-2
        lesson_range: Tuple (start, end) cho Stage 4
        force: Force re-analyze
        dpi: Render DPI
    """
    start_time = time.time()
    
    print(f'\n{"█" * 60}')
    print(f'  SGK OCR PIPELINE — LỚP {grade}')
    print(f'  Stage: {stage} | DPI: {dpi} | Force: {force}')
    if page_range:
        print(f'  Page range: {page_range[0]}-{page_range[1]}')
    if lesson_range:
        print(f'  Lesson range: {lesson_range[0]}-{lesson_range[1]}')
    print(f'{"█" * 60}')
    
    stages_to_run = {
        'all': ['render', 'analyze', 'split', 'build-docx'],
        'render': ['render'],
        'analyze': ['analyze'],
        'split': ['split'],
        'build-docx': ['build-docx']
    }
    
    stages = stages_to_run.get(stage, ['all'])
    
    # ── Stage 1: Render ──
    if 'render' in stages:
        print(f'\n{"─" * 60}')
        print(f'  ▶ STAGE 1: Rendering PDF pages...')
        print(f'{"─" * 60}')
        manifest = render_grade(grade, page_range, dpi)
        if not manifest:
            print(f'  [!] Stage 1 failed for Lớp {grade}')
            return False
    
    # ── Stage 2: Analyze ──
    if 'analyze' in stages:
        print(f'\n{"─" * 60}')
        print(f'  ▶ STAGE 2: Gemini Vision Analysis...')
        print(f'{"─" * 60}')
        analysis = analyze_grade(grade, page_range, force)
        if not analysis:
            print(f'  [!] Stage 2 failed for Lớp {grade}')
            return False
    
    # ── Stage 3: Split ──
    if 'split' in stages:
        print(f'\n{"─" * 60}')
        print(f'  ▶ STAGE 3: Lesson Splitting...')
        print(f'{"─" * 60}')
        lessons = split_grade(grade)
        if not lessons:
            print(f'  [!] Stage 3 failed for Lớp {grade}')
            return False
    
    # ── Stage 4: Build DOCX ──
    if 'build-docx' in stages:
        print(f'\n{"─" * 60}')
        print(f'  ▶ STAGE 4: Building DOCX files...')
        print(f'{"─" * 60}')
        build_result = build_grade(grade, lesson_range)
        if not build_result:
            print(f'  [!] Stage 4 failed for Lớp {grade}')
            return False
    
    elapsed = time.time() - start_time
    print(f'\n{"█" * 60}')
    print(f'  ✅ LỚP {grade}: PIPELINE COMPLETE ({elapsed:.1f}s)')
    print(f'{"█" * 60}')
    
    return True


def main():
    parser = argparse.ArgumentParser(
        description='SGK OCR Pipeline — Chuyển đổi SGK PDF → DOCX theo bài/lớp',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scripts/sgk_pipeline.py --grade 3              # Full pipeline Lớp 3
  python scripts/sgk_pipeline.py --grade 3 --pages 6-20 # Chỉ trang 6-20
  python scripts/sgk_pipeline.py --grade 6 --lessons 1-5 # DOCX Bài 1-5
  python scripts/sgk_pipeline.py --all                   # Tất cả các lớp
  python scripts/sgk_pipeline.py --grade 3 --stage render # Chỉ Stage 1
        """
    )
    
    parser.add_argument('--grade', type=int, choices=[3, 4, 5, 6, 7, 8],
                        help='Khối lớp cần xử lý')
    parser.add_argument('--all', action='store_true',
                        help='Xử lý tất cả các lớp (3-8)')
    parser.add_argument('--stage', type=str, default='all',
                        choices=['all', 'render', 'analyze', 'split', 'build-docx'],
                        help='Stage cần chạy (default: all)')
    parser.add_argument('--pages', type=str, default=None,
                        help='Phạm vi trang 1-indexed (VD: 6-20)')
    parser.add_argument('--lessons', type=str, default=None,
                        help='Phạm vi bài (VD: 1-5)')
    parser.add_argument('--force', action='store_true',
                        help='Force re-analyze (bỏ qua cache)')
    parser.add_argument('--dpi', type=int, default=300,
                        help='Render DPI (default: 300)')
    
    args = parser.parse_args()
    
    page_range = parse_page_range(args.pages) if args.pages else None
    lesson_range = parse_lesson_range(args.lessons) if args.lessons else None
    
    if args.all:
        grades = [3, 4, 5, 6, 7, 8]
    elif args.grade:
        grades = [args.grade]
    else:
        parser.print_help()
        return
    
    total_start = time.time()
    success = 0
    
    for grade in grades:
        ok = run_pipeline(
            grade,
            stage=args.stage,
            page_range=page_range,
            lesson_range=lesson_range,
            force=args.force,
            dpi=args.dpi
        )
        if ok:
            success += 1
    
    total_elapsed = time.time() - total_start
    
    print(f'\n{"█" * 60}')
    print(f'  PIPELINE FINISHED: {success}/{len(grades)} grades OK ({total_elapsed:.1f}s)')
    print(f'{"█" * 60}')


if __name__ == '__main__':
    main()
