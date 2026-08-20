# -*- coding: utf-8 -*-
"""
Extract images from SGK Tin học (KNTT) for Bài 2, Bài 3, Bài 4 for Grades 3 to 8.
"""
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')
import fitz  # PyMuPDF

SGK_DIR = r'D:\UNIGO\SGK'

GRADE_PDFS = {
    3: os.path.join(SGK_DIR, 'Lớp_3', 'SGK Tin học 3 KNTT.pdf'),
    4: os.path.join(SGK_DIR, 'Lớp_4', 'SGK Tin học 4 KNTT.pdf'),
    5: os.path.join(SGK_DIR, 'Lớp_5', 'SGK Tin học 5 KNTT.pdf'),
    6: os.path.join(SGK_DIR, 'Lớp_6', 'SGK Tin học  6 KNTT.pdf'),
    7: os.path.join(SGK_DIR, 'Lớp_7', 'SGK Tin học 7 KNTT.pdf'),
    8: os.path.join(SGK_DIR, 'Lớp_8', 'SGK TIN HỌC 8 KNTT.pdf'),
}

def extract_all_grade_images(grade, pdf_path):
    if not os.path.isfile(pdf_path):
        print(f"[-] Grade {grade}: PDF not found: {pdf_path}")
        return
    
    print(f"\n[+] Extracting Grade {grade} from {os.path.basename(pdf_path)}...")
    doc = fitz.open(pdf_path)
    out_dir = os.path.join(SGK_DIR, f'Lớp_{grade}', 'all_extracted_images')
    os.makedirs(out_dir, exist_ok=True)
    
    img_count = 0
    for page_idx in range(min(len(doc), 60)):  # First 60 pages cover Bài 1 to 5
        page = doc[page_idx]
        image_list = page.get_images(full=True)
        
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue
                
            img_bytes = base_image["image"]
            img_ext = base_image["ext"]
            w = base_image.get("width", 0)
            h = base_image.get("height", 0)
            
            if w < 80 or h < 80 or (w * h < 10000):
                continue
                
            fname = f"p{page_idx+1}_img{img_idx+1}_{w}x{h}.{img_ext}"
            fpath = os.path.join(out_dir, fname)
            with open(fpath, "wb") as f:
                f.write(img_bytes)
            img_count += 1
            
    print(f"    Saved {img_count} images for Grade {grade} into {out_dir}")

def main():
    for g, path in GRADE_PDFS.items():
        extract_all_grade_images(g, path)

if __name__ == '__main__':
    main()
