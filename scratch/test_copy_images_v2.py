# -*- coding: utf-8 -*-
"""
TEST v2: Copy hình ảnh từ giáo án gốc → KHBD Robotics (1 file)
Sửa lỗi: dùng part.get_or_add_image() thay vì part.get_or_add_image_part()
"""
import sys, os, re, io, copy
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Inches, Emu, Pt
from lxml import etree

# Namespaces
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

def qn(tag):
    prefix, local = tag.split(':')
    nsmap = {'w': W_NS, 'r': R_NS, 'a': A_NS, 'wp': WP_NS, 'pic': PIC_NS}
    return f'{{{nsmap[prefix]}}}{local}'


def get_source_images_by_section(src_doc):
    """Parse source and collect images grouped by section."""
    sections = {'khoi_dong': [], 'kien_thuc': [], 'luyen_tap': [], 'van_dung': []}
    
    # Build rId → blob mapping
    rid_to_blob = {}
    for rel_id, rel in src_doc.part.rels.items():
        if 'image' in str(rel.reltype).lower():
            rid_to_blob[rel_id] = rel.target_part.blob
    
    current_section = None
    for i, p in enumerate(src_doc.paragraphs):
        txt = p.text.strip()
        
        if re.match(r'^1\.\s*(Khởi động|Let)', txt, re.IGNORECASE):
            current_section = 'khoi_dong'
        elif re.match(r'^2\.\s*(Kiến thức|Let)', txt, re.IGNORECASE):
            current_section = 'kien_thuc'
        elif re.match(r'^3\.\s*(Luyện tập|Let)', txt, re.IGNORECASE):
            current_section = 'luyen_tap'
        elif re.match(r'^[4-6]\.\s*(Vận dụng|Tự học|Sáng tạo|Tự kiểm tra|My own|Let)', txt, re.IGNORECASE):
            current_section = 'van_dung'
        
        drawings = p._element.findall(f'.//{qn("w:drawing")}')
        if drawings and current_section:
            for drawing in drawings:
                blips = drawing.findall(f'.//{qn("a:blip")}')
                drawing_blobs = {}
                for blip in blips:
                    rid = blip.get(qn('r:embed'))
                    if rid and rid in rid_to_blob:
                        drawing_blobs[rid] = rid_to_blob[rid]
                
                if drawing_blobs:
                    sections[current_section].append({
                        'para_idx': i,
                        'drawing_xml': drawing,
                        'blobs': drawing_blobs,
                    })
    
    return sections


def copy_images_to_khbd_table(dst_doc, section_images, table_idx):
    """Copy images from source section into KHBD table cell."""
    if table_idx >= len(dst_doc.tables):
        return 0
    
    table = dst_doc.tables[table_idx]
    if len(table.rows) < 2:
        return 0
    
    cell = table.rows[1].cells[0]  # GV-HS column
    count = 0
    
    for img_data in section_images:
        drawing_xml = img_data['drawing_xml']
        blobs = img_data['blobs']
        
        if not blobs:
            continue
        
        # Deep copy the drawing element
        new_drawing = copy.deepcopy(drawing_xml)
        
        # Update blip references with new rIds
        blips = new_drawing.findall(f'.//{qn("a:blip")}')
        all_ok = True
        for blip in blips:
            old_rid = blip.get(qn('r:embed'))
            if old_rid in blobs:
                image_blob = blobs[old_rid]
                try:
                    # Use the CORRECT method
                    rId, image = dst_doc.part.get_or_add_image(io.BytesIO(image_blob))
                    blip.set(qn('r:embed'), rId)
                except Exception as e:
                    print(f'    Error adding image: {e}')
                    all_ok = False
        
        if not all_ok:
            continue
        
        # Add drawing to cell
        new_para = cell.add_paragraph()
        new_para.alignment = 1  # CENTER
        run_element = etree.SubElement(new_para._element, qn('w:r'))
        run_element.append(new_drawing)
        count += 1
    
    return count


# ============================================================
# TEST
# ============================================================
src_path = r'D:\UNIGO\Phân phối chương trình\Robotics\Giáo án-Initiate\I 03 - 01 - đu xà.docx'
# Use BACKUP to avoid modifying already-changed file
dst_path = r'D:\UNIGO\KHBD_Robotics\Lớp_1\Tuần_02\KHBD_Robotics_Lớp_1_Tiet01_Bai_1_Tap_the_duc_nao.docx'

src_doc = Document(src_path)
dst_doc = Document(dst_path)

# Get images by section
sections = get_source_images_by_section(src_doc)
for sec_name, imgs in sections.items():
    print(f'  {sec_name}: {len(imgs)} image(s)')

# Copy images
total = 0
mapping = {'khoi_dong': 1, 'kien_thuc': 2, 'luyen_tap': 3, 'van_dung': 4}
for sec_name, table_idx in mapping.items():
    imgs = sections[sec_name]
    if imgs:
        count = copy_images_to_khbd_table(dst_doc, imgs, table_idx)
        total += count
        print(f'  Copied {count} image(s) → Table {table_idx} ({sec_name})')

# Save
test_out = r'D:\UNIGO\scratch\TEST_KHBD_with_images_v2.docx'
dst_doc.save(test_out)

# Verify
v = Document(test_out)
v_drawings = v.element.body.findall(f'.//{qn("w:drawing")}')
v_rels = [r for r in v.part.rels.values() if 'image' in str(r.reltype).lower()]
print(f'\nResult: {len(v_drawings)} drawings, {len(v_rels)} image rels')
print(f'Test file: {test_out}')
