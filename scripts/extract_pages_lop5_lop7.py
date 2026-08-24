import fitz, os, sys
sys.stdout.reconfigure(encoding='utf-8')

def extract_and_inspect(pdf_path, out_dir, start_p, end_p):
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    for p in range(start_p - 1, min(end_p, len(doc))):
        page = doc[p]
        pix = page.get_pixmap(dpi=150)
        out_path = os.path.join(out_dir, f"page_{p+1}.png")
        pix.save(out_path)
        print(f"Saved {out_path} ({pix.width}x{pix.height})")

# Let's extract pages 10 to 30 for Lop 5 and Lop 7
extract_and_inspect(r'D:\UNIGO\SGK\Lớp_5\SGK Tin học 5 KNTT.pdf', r'D:\UNIGO\SGK\Lớp_5\pages_10_30', 10, 26)
extract_and_inspect(r'D:\UNIGO\SGK\Lớp_7\SGK Tin học 7 KNTT.pdf', r'D:\UNIGO\SGK\Lớp_7\pages_10_30', 10, 28)
