import os, sys
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image

def analyze_pages(folder):
    files = sorted([f for f in os.listdir(folder) if f.endswith('.png')])
    for f in files:
        p = os.path.join(folder, f)
        img = Image.open(p)
        print(f"{f}: size={img.size}")

print("Lớp 5 pages:")
analyze_pages(r'D:\UNIGO\SGK\Lớp_5\pages_10_30')

print("\nLớp 7 pages:")
analyze_pages(r'D:\UNIGO\SGK\Lớp_7\pages_10_30')
