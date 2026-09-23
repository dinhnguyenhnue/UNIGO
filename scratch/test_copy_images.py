# -*- coding: utf-8 -*-
"""
TEST: Copy hình ảnh + NLS từ 1 file giáo án gốc → 1 file KHBD
Kiểm tra kỹ thuật trước khi scale lên toàn bộ.
"""
import sys, os, re, io, copy
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Inches, Emu, Pt
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

# Namespaces
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

def qn(tag):
    """Qualified name helper."""
    prefix, local = tag.split(':')
    nsmap = {'w': W_NS, 'r': R_NS, 'a': A_NS, 'wp': WP_NS, 'pic': PIC_NS}
    return f'{{{nsmap[prefix]}}}{local}'


def get_source_images_by_section(src_doc):
    """
    Parse source document and collect images grouped by section.
    Returns dict: { 'khoi_dong': [(para_idx, drawing_element, rId_to_blob_map)], ... }
    """
    sections = {
        'khoi_dong': [],    # → KHBD Table 1
        'kien_thuc': [],    # → KHBD Table 2
        'luyen_tap': [],    # → KHBD Table 3
        'van_dung': [],     # → KHBD Table 4 (includes Tự học, Sáng tạo, Tự kiểm tra)
    }
    
    # Build rId → blob mapping from source
    rid_to_blob = {}
    for rel_id, rel in src_doc.part.rels.items():
        if 'image' in str(rel.reltype).lower():
            rid_to_blob[rel_id] = rel.target_part.blob
    
    # Detect sections and collect image paragraphs
    current_section = None
    for i, p in enumerate(src_doc.paragraphs):
        txt = p.text.strip()
        
        # Detect section boundaries
        if re.match(r'^1\.\s*(Khởi động|Let)', txt, re.IGNORECASE):
            current_section = 'khoi_dong'
        elif re.match(r'^2\.\s*(Kiến thức|Let)', txt, re.IGNORECASE):
            current_section = 'kien_thuc'
        elif re.match(r'^3\.\s*(Luyện tập|Let)', txt, re.IGNORECASE):
            current_section = 'luyen_tap'
        elif re.match(r'^4\.\s*(Vận dụng|My own)', txt, re.IGNORECASE):
            current_section = 'van_dung'
        elif re.match(r'^5\.\s*(Tự học|Sáng tạo|Let)', txt, re.IGNORECASE):
            current_section = 'van_dung'
        elif re.match(r'^6\.\s*(Tự kiểm tra|Let)', txt, re.IGNORECASE):
            current_section = 'van_dung'
        
        # Check if this paragraph has drawings
        drawings = p._element.findall(f'.//{qn("w:drawing")}')
        if drawings and current_section:
            for drawing in drawings:
                # Find blip rIds in this drawing
                blips = drawing.findall(f'.//{qn("a:blip")}')
                drawing_blobs = {}
                for blip in blips:
                    rid = blip.get(qn('r:embed'))
                    if rid and rid in rid_to_blob:
                        drawing_blobs[rid] = rid_to_blob[rid]
                
                sections[current_section].append({
                    'para_idx': i,
                    'drawing_xml': drawing,
                    'blobs': drawing_blobs,
                })
    
    return sections


def copy_images_to_khbd_table(dst_doc, section_images, table_idx):
    """
    Copy images from source section into the specified KHBD table cell (GV-HS column).
    """
    if table_idx >= len(dst_doc.tables):
        return 0
    
    table = dst_doc.tables[table_idx]
    if len(table.rows) < 2:
        return 0
    
    # GV-HS cell is column 0, row 1 (data row)
    cell = table.rows[1].cells[0]
    
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
        for blip in blips:
            old_rid = blip.get(qn('r:embed'))
            if old_rid in blobs:
                image_blob = blobs[old_rid]
                # Add image to destination document
                try:
                    rId, image_part = dst_doc.part.get_or_add_image_part(
                        io.BytesIO(image_blob)
                    )
                    # Update the rId reference
                    blip.set(qn('r:embed'), rId)
                except Exception as e:
                    print(f'    Warning: Could not add image: {e}')
                    continue
        
        # Add a new paragraph with the drawing to the cell
        new_para = cell.add_paragraph()
        new_para.alignment = 1  # CENTER
        run_element = etree.SubElement(new_para._element, qn('w:r'))
        run_element.append(new_drawing)
        
        # Set font for the run
        rpr = etree.SubElement(run_element, qn('w:rPr'))
        sz = etree.SubElement(rpr, qn('w:sz'))
        sz.set(qn('w:val'), '26')  # 13pt
        
        count += 1
    
    return count


def copy_nls_from_source(src_doc, dst_doc):
    """
    Copy năng lực descriptions from source into KHBD.
    Replaces specific NL content while keeping KHBD structure.
    """
    # Extract source's "Về năng lực" lines
    src_nl_lines = []
    src_pc_lines = []
    current = None
    
    for p in src_doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        if re.match(r'^2\.\s*Về năng lực', txt):
            current = 'nl'
            continue
        elif re.match(r'^3\.\s*Về phẩm chất', txt):
            current = 'pc'
            continue
        elif re.match(r'^(I{1,3}|IV)\.\s', txt) or re.match(r'^[1-6]\.\s*(Khởi động|Kiến thức|Luyện tập)', txt):
            current = None
            continue
        
        if current == 'nl':
            src_nl_lines.append(txt)
        elif current == 'pc':
            src_pc_lines.append(txt)
    
    return src_nl_lines, src_pc_lines


# ============================================================
# TEST: One file
# ============================================================

src_path = r'D:\UNIGO\Phân phối chương trình\Robotics\Giáo án-Initiate\I 03 - 01 - đu xà.docx'
dst_path = r'D:\UNIGO\KHBD_Robotics\Lớp_1\Tuần_02\KHBD_Robotics_Lớp_1_Tiet01_Bai_1_Tap_the_duc_nao.docx'

print(f'Source: {os.path.basename(src_path)}')
print(f'Dest: {os.path.basename(dst_path)}')

# Open documents
src_doc = Document(src_path)
dst_doc = Document(dst_path)

# Check source images
print(f'\nSource images before:')
src_rels = [r for r in src_doc.part.rels.values() if 'image' in str(r.reltype).lower()]
print(f'  Image rels: {len(src_rels)}')

# Get images by section
sections = get_source_images_by_section(src_doc)
for sec_name, imgs in sections.items():
    print(f'  {sec_name}: {len(imgs)} image(s)')

# Copy NLS
nl_lines, pc_lines = copy_nls_from_source(src_doc, dst_doc)
print(f'\nNLS from source:')
for l in nl_lines:
    print(f'  NL: {l[:100]}')
for l in pc_lines:
    print(f'  PC: {l[:100]}')

# Copy images to tables
total_images = 0
section_to_table = {'khoi_dong': 1, 'kien_thuc': 2, 'luyen_tap': 3, 'van_dung': 4}
for sec_name, table_idx in section_to_table.items():
    imgs = sections[sec_name]
    if imgs:
        count = copy_images_to_khbd_table(dst_doc, imgs, table_idx)
        total_images += count
        print(f'  Copied {count} image(s) to Table {table_idx} ({sec_name})')

# Save test file
test_output = r'D:\UNIGO\scratch\TEST_KHBD_with_images.docx'
dst_doc.save(test_output)
print(f'\nSaved test file: {test_output}')
print(f'Total images copied: {total_images}')

# Verify
verify_doc = Document(test_output)
verify_drawings = verify_doc.element.body.findall(f'.//{qn("w:drawing")}')
verify_rels = [r for r in verify_doc.part.rels.values() if 'image' in str(r.reltype).lower()]
print(f'\nVerification:')
print(f'  Drawings in output: {len(verify_drawings)}')
print(f'  Image rels in output: {len(verify_rels)}')
print(f'  Tables: {len(verify_doc.tables)}')
