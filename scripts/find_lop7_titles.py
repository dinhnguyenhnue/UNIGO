import fitz, os, sys
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image

def find_titles(pdf_path):
    doc = fitz.open(pdf_path)
    for pno in range(len(doc)):
        text = doc[pno].get_text()
        if "BÀI" in text or "Bài" in text or "Chủ đề" in text or "CHỦ ĐỀ" in text:
            for l in text.splitlines():
                if any(k in l.upper() for k in ["BÀI 1", "BÀI 2", "BÀI 3", "BÀI 4", "BÀI 5"]):
                    print(f"Page {pno+1}: {l.strip()}")

print("=== Lớp 7 SGK Titles ===")
find_titles(r'D:\UNIGO\SGK\Lớp_7\SGK Tin học 7 KNTT.pdf')
