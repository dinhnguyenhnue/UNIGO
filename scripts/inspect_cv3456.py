import os
import pypdf
import pypdfium2 as pdfium

pdf_path = r"D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\3456-VV_huong_dan_trien_khai_Khung_nang_luc_so_cho_HS_885ca.pdf"

reader = pypdf.PdfReader(pdf_path)
pdf = pdfium.PdfDocument(pdf_path)

print(f"Total pages: {len(reader.pages)}")
for i, page in enumerate(reader.pages):
    text = page.extract_text() or ""
    print(f"Page {i+1}: text_len={len(text)}, images={len(page.images)}")
