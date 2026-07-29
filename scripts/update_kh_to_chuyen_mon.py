import os, sys, re, shutil, math
import docx
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

def set_font_run(run, font_name="Times New Roman", size_pt=13, bold=None, italic=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
        
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
    rPr.append(rFonts)

def format_paragraph(p, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=None, italic=None):
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    for run in p.runs:
        set_font_run(run, font_name=font_name, size_pt=size_pt, bold=bold, italic=italic)

def format_cell(cell, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=None, italic=None):
    tcPr = cell._element.get_or_add_tcPr()
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>')
    tcPr.append(vAlign)
    for p in cell.paragraphs:
        format_paragraph(p, font_name=font_name, size_pt=size_pt, align=align, bold=bold, italic=italic)

def set_cell_width(cell, width_cm):
    cell.width = Cm(width_cm)

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        # Check if borders already exist to avoid duplicate XML elements
        if not tblPr[0].xpath('w:tblBorders'):
            borders_xml = parse_xml(
                f'<w:tblBorders {nsdecls("w")}>\n'
                f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
                f'</w:tblBorders>'
            )
            tblPr[0].append(borders_xml)

def extract_rows_from_5col_khdh(file_path, table_index):
    doc = docx.Document(file_path)
    t = doc.tables[table_index]
    rows = []
    for r_idx in range(1, len(t.rows)):
        c = [cell.text.strip().replace('\n', ' ') for cell in t.rows[r_idx].cells]
        if len(c) < 5:
            continue
        stt, name, so_tiet, ppct, yccd = c[0], c[1], c[2], c[3], c[4]
        is_topic = ('chủ đề' in name.lower() or 'chủ đề' in stt.lower()) and ('bài ' not in name.lower())
        rows.append((is_topic, stt, name, so_tiet, ppct, yccd))
    return rows

def build_7col_table(doc, rows_data):
    new_table = doc.add_table(rows=0, cols=7)
    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(new_table)
    
    col_widths_cm = [1.0, 3.2, 1.2, 1.5, 1.6, 3.0, 5.0]
    
    hdr_row = new_table.add_row()
    hdr_titles = ['TT', 'Bài/chủ đề', 'Tổng số tiết', 'Tuần', 'Tiết theo PPCT', 'Nội dung', 'Mục tiêu bài học']
    for idx, title in enumerate(hdr_titles):
        cell = hdr_row.cells[idx]
        cell.text = title
        align = WD_ALIGN_PARAGRAPH.CENTER if idx in [0, 2, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
        format_cell(cell, font_name="Times New Roman", size_pt=13, align=align, bold=True)
        set_cell_width(cell, col_widths_cm[idx])
        
    hk1_row = new_table.add_row()
    c0 = hk1_row.cells[0]
    c0.merge(hk1_row.cells[6])
    c0.text = "HỌC KỲ I (Tiết 1 đến Tiết 19)"
    format_cell(c0, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    hk2_inserted = False
    
    for is_topic, stt, name, so_tiet, ppct, yccd in rows_data:
        nums = [int(n) for n in re.findall(r'\d+', ppct)]
        if not hk2_inserted and nums and nums[0] >= 20:
            hk2_row = new_table.add_row()
            c0 = hk2_row.cells[0]
            c0.merge(hk2_row.cells[6])
            c0.text = "HỌC KỲ II (Tiết 20 đến Tiết 35)"
            format_cell(c0, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
            hk2_inserted = True
            
        row = new_table.add_row()
        if is_topic:
            c0 = row.cells[0]
            c0.merge(row.cells[6])
            c0.text = name
            format_cell(c0, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=True)
        else:
            if ppct == '1':
                tuan_str = 'Tuần 1'
            elif ',' in ppct:
                tuan_str = f"Tuần {ppct}"
            elif '-' in ppct:
                tuan_str = f"Tuần {ppct}"
            else:
                tuan_str = f"Tuần {ppct}"
                
            row.cells[0].text = stt
            format_cell(row.cells[0], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_width(row.cells[0], col_widths_cm[0])
            
            row.cells[1].text = name
            format_cell(row.cells[1], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            set_cell_width(row.cells[1], col_widths_cm[1])
            
            row.cells[2].text = so_tiet
            format_cell(row.cells[2], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_width(row.cells[2], col_widths_cm[2])
            
            row.cells[3].text = tuan_str
            format_cell(row.cells[3], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_width(row.cells[3], col_widths_cm[3])
            
            row.cells[4].text = ppct
            format_cell(row.cells[4], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_width(row.cells[4], col_widths_cm[4])
            
            row.cells[5].text = name
            format_cell(row.cells[5], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            set_cell_width(row.cells[5], col_widths_cm[5])
            
            row.cells[6].text = yccd
            format_cell(row.cells[6], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            set_cell_width(row.cells[6], col_widths_cm[6])
            
    return new_table

def build_english_7col_table(doc, src_table):
    new_table = doc.add_table(rows=0, cols=7)
    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(new_table)
    
    col_widths_cm = [1.0, 3.2, 1.2, 1.5, 1.6, 3.0, 5.0]
    
    hdr_row = new_table.add_row()
    hdr_titles = ['TT', 'Bài/chủ đề', 'Tổng số tiết', 'Tuần', 'Tiết theo PPCT', 'Nội dung', 'Mục tiêu bài học']
    for idx, title in enumerate(hdr_titles):
        cell = hdr_row.cells[idx]
        cell.text = title
        align = WD_ALIGN_PARAGRAPH.CENTER if idx in [0, 2, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
        format_cell(cell, font_name="Times New Roman", size_pt=13, align=align, bold=True)
        set_cell_width(cell, col_widths_cm[idx])
        
    hk1_row = new_table.add_row()
    c0 = hk1_row.cells[0]
    c0.merge(hk1_row.cells[6])
    c0.text = "HỌC KỲ I"
    format_cell(c0, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    hk2_inserted = False
    
    for r_idx in range(1, len(src_table.rows)):
        c = [cell.text.strip().replace('\n', ' ') for cell in src_table.rows[r_idx].cells]
        if len(c) < 4:
            continue
        stt, name, so_tiet, yccd = c[0], c[1], c[2], c[3]
        
        try:
            p_num = int(stt)
        except:
            p_num = r_idx
            
        tuan_num = math.ceil(p_num / 10.0) if p_num <= 350 else 35
        tuan_str = f"Tuần {tuan_num}"
        
        if not hk2_inserted and p_num > 175:
            hk2_row = new_table.add_row()
            c0 = hk2_row.cells[0]
            c0.merge(hk2_row.cells[6])
            c0.text = "HỌC KỲ II"
            format_cell(c0, font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
            hk2_inserted = True
            
        row = new_table.add_row()
        row.cells[0].text = str(stt)
        format_cell(row.cells[0], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[0], col_widths_cm[0])
        
        row.cells[1].text = name
        format_cell(row.cells[1], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell_width(row.cells[1], col_widths_cm[1])
        
        row.cells[2].text = "1"
        format_cell(row.cells[2], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[2], col_widths_cm[2])
        
        row.cells[3].text = tuan_str
        format_cell(row.cells[3], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[3], col_widths_cm[3])
        
        row.cells[4].text = str(stt)
        format_cell(row.cells[4], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[4], col_widths_cm[4])
        
        row.cells[5].text = name
        format_cell(row.cells[5], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell_width(row.cells[5], col_widths_cm[5])
        
        row.cells[6].text = yccd
        format_cell(row.cells[6], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell_width(row.cells[6], col_widths_cm[6])
        
    return new_table

def build_english_assessment_table(doc, lop_str):
    new_table = doc.add_table(rows=0, cols=5)
    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(new_table)
    
    col_widths_cm = [1.0, 1.5, 3.0, 6.5, 4.4]
    
    hdr_row = new_table.add_row()
    hdr_titles = ['TT', 'Lớp', 'Bài kiểm tra', 'Nội dung', 'Hình thức']
    for idx, title in enumerate(hdr_titles):
        cell = hdr_row.cells[idx]
        cell.text = title
        align = WD_ALIGN_PARAGRAPH.CENTER if idx in [0, 1, 2] else WD_ALIGN_PARAGRAPH.LEFT
        format_cell(cell, font_name="Times New Roman", size_pt=13, align=align, bold=True)
        set_cell_width(cell, col_widths_cm[idx])
        
    assessments = [
        ("1", lop_str, "Đánh giá định kỳ 1", "Nắm vững từ vựng và ngữ pháp từ Unit 1 đến Unit 3; kỹ năng Nghe và Đọc cơ bản (Tuần 11).", "Trắc nghiệm và vấn đáp"),
        ("2", lop_str, "Đánh giá định kỳ 2", "Tổng hợp kiến thức và kỹ năng ngôn ngữ Học kỳ I (Tuần 16).", "Trắc nghiệm và vấn đáp"),
        ("3", lop_str, "Đánh giá định kỳ 3", "Nắm vững từ vựng và cấu trúc câu từ Unit 7 đến Unit 9; kỹ năng Viết và Nói (Tuần 28).", "Trắc nghiệm và vấn đáp"),
        ("4", lop_str, "Đánh giá định kỳ 4", "Tổng hợp toàn bộ kiến thức và kỹ năng ngôn ngữ Học kỳ II (Tuần 33).", "Trắc nghiệm và vấn đáp"),
    ]
    
    for tt, lp, bkt, nd, ht in assessments:
        row = new_table.add_row()
        row.cells[0].text = tt
        format_cell(row.cells[0], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[0], col_widths_cm[0])
        
        row.cells[1].text = lp
        format_cell(row.cells[1], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[1], col_widths_cm[1])
        
        row.cells[2].text = bkt
        format_cell(row.cells[2], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_width(row.cells[2], col_widths_cm[2])
        
        row.cells[3].text = nd
        format_cell(row.cells[3], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell_width(row.cells[3], col_widths_cm[3])
        
        row.cells[4].text = ht
        format_cell(row.cells[4], font_name="Times New Roman", size_pt=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_cell_width(row.cells[4], col_widths_cm[4])
        
    return new_table

def update_khtcm_file(main_khtcm_path):
    print(f"Updating: {main_khtcm_path}")
    doc = docx.Document(main_khtcm_path)
    
    tin_thcs_path = r".\Hệ thống mẫu văn bản\Nguyên đã làm\Kế hoạch dạy học môn Tin học (THCS) - 2026 - 2027.docx"
    robotics_thcs_path = r".\Hệ thống mẫu văn bản\Nguyên đã làm\Kế hoạch dạy học môn Robotics (THCS) - 2026 - 2027.docx"
    english_pl1_path = r".\Hệ thống mẫu văn bản\Nguyên đã làm\Kế hoạch tổ chuyên môn\PL1. Khung kế hoạch dạy học môn Tiếng Anh (THCS)-PPCT.docx"
    
    # Apply borders and formatting to all tables in doc
    for table in doc.tables:
        set_table_borders(table)
        
    # 1. Update Tin học 6, 7, 8 (Tables 9, 11, 13)
    tin_map = [(9, 3), (11, 4), (13, 5)]
    for target_t_idx, src_t_idx in tin_map:
        rows_data = extract_rows_from_5col_khdh(tin_thcs_path, src_t_idx)
        old_t = doc.tables[target_t_idx]
        old_elem = old_t._element
        parent = old_elem.getparent()
        pos = parent.index(old_elem)
        
        new_t = build_7col_table(doc, rows_data)
        parent.insert(pos, new_t._element)
        parent.remove(old_elem)
        print(f"  Updated Tin học table {target_t_idx}")

    # 2. Update Robotics 6, 7, 8 (Tables 15, 17, 19)
    rob_map = [(15, 3), (17, 4), (19, 5)]
    for target_t_idx, src_t_idx in rob_map:
        rows_data = extract_rows_from_5col_khdh(robotics_thcs_path, src_t_idx)
        old_t = doc.tables[target_t_idx]
        old_elem = old_t._element
        parent = old_elem.getparent()
        pos = parent.index(old_elem)
        
        new_t = build_7col_table(doc, rows_data)
        parent.insert(pos, new_t._element)
        parent.remove(old_elem)
        print(f"  Updated Robotics table {target_t_idx}")
        
    # Apply borders to all tables after updates
    for table in doc.tables:
        set_table_borders(table)

    doc.save(main_khtcm_path)
    print("  Saved updated Kế hoạch tổ chuyên môn (THCS) with borders on all tables.")

def main():
    path = r"D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Kế hoạch tổ chuyên môn\Kế hoạch tổ chuyên môn (THCS).docx"
    update_khtcm_file(path)
    
    sync_dst = r"D:\UNIGO\Hệ thống mẫu văn bản\09.07.26. Kế hoạch tổ chuyên môn (THCS).docx"
    shutil.copy2(path, sync_dst)
    print(f"Synced to: {sync_dst}")

if __name__ == "__main__":
    main()
