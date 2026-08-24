import fitz, os, sys
sys.stdout.reconfigure(encoding='utf-8')

def find_pages(pdf_path, name):
    doc = fitz.open(pdf_path)
    print(f"=== {name} ({len(doc)} pages) ===")
    for i, page in enumerate(doc):
        text = page.get_text()
        for kw in ["BÀI 3", "BÀI 4", "Bài 3", "Bài 4"]:
            if kw in text:
                lines = [l.strip() for l in text.splitlines() if l.strip()][:3]
                print(f"Page {i+1}: {kw} -> {' | '.join(lines)}")

find_pages(r'D:\UNIGO\SGK\Lớp_5\SGK Tin học 5 KNTT.pdf', 'Lớp 5')
find_pages(r'D:\UNIGO\SGK\Lớp_7\SGK Tin học 7 KNTT.pdf', 'Lớp 7')
