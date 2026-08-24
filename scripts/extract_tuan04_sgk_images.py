import fitz, os, io
from PIL import Image

def extract_page_images(pdf_path, start_page, end_page, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    count = 0
    for pno in range(start_page - 1, min(end_page, len(doc))):
        page = doc[pno]
        img_list = page.get_images(full=True)
        for img_idx, img_info in enumerate(img_list):
            xref = img_info[0]
            base_img = doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext = base_img["ext"]
            try:
                img = Image.open(io.BytesIO(img_bytes))
                # Only keep reasonable sized images (skip 1x1 or tiny icons)
                if img.width > 80 and img.height > 80:
                    out_name = f"p{pno+1}_img{img_idx+1}_{img.width}x{img.height}.{ext}"
                    out_path = os.path.join(out_dir, out_name)
                    img.save(out_path)
                    print(f"Page {pno+1}: Saved {out_name}")
                    count += 1
            except Exception as e:
                pass
    print(f"Total extracted: {count} images in {out_dir}")

extract_page_images(r'D:\UNIGO\SGK\Lớp_5\SGK Tin học 5 KNTT.pdf', 14, 23, r'D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images\sgk_extracted')
extract_page_images(r'D:\UNIGO\SGK\Lớp_7\SGK Tin học 7 KNTT.pdf', 12, 26, r'D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images\sgk_extracted')
