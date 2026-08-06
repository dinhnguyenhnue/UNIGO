"""Extract Bài 1 content from SGK Tin học for all grades (3-8)."""
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')

import pypdf

SGK_DIR = r'd:\UNIGO\SGK'

for grade in [3, 4, 5, 6, 7, 8]:
    folder = os.path.join(SGK_DIR, f'Lớp_{grade}')
    pdfs = glob.glob(os.path.join(folder, '*.pdf'))
    if not pdfs:
        print(f'=== LỚP {grade}: Không tìm thấy PDF ===\n')
        continue

    pdf_path = pdfs[0]
    print(f'=== LỚP {grade}: {os.path.basename(pdf_path)} ===')
    reader = pypdf.PdfReader(pdf_path)
    print(f'    Tổng: {len(reader.pages)} trang')

    # Find pages containing "Bài 1" and extract text
    found = False
    for i in range(min(25, len(reader.pages))):
        text = reader.pages[i].extract_text() or ''
        if 'Bài 1' in text or 'BÀI 1' in text:
            if not found:
                print(f'\n--- Bài 1 bắt đầu tại trang {i+1} ---')
                found = True
            print(f'\n[Trang {i+1}]:')
            print(text[:1500])
            if len(text) > 1500:
                print('... (truncated)')
        elif found:
            # Print one more page after last Bài 1 page
            text2 = reader.pages[i].extract_text() or ''
            if 'Bài 2' in text2 or 'BÀI 2' in text2:
                break
            print(f'\n[Trang {i+1}]:')
            print(text2[:1500])
            break
    print('\n')
