"""
generate_khbd_all.py — Tạo lại toàn bộ KHBD Bài 1 từ Tiền tiểu học đến Lớp 8
Dựa trên luật KHBD_TIEU_HOC.md và KHBD_THCS.md đã được phân tích từ template thực tế.
"""
import os, sys, shutil, copy
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Emu, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ─── TEMPLATE PATHS ────────────────────────────────────────────────────────
TPL_TH   = r'd:\UNIGO\Hệ thống mẫu văn bản\Khung  giáo án Unigo 2026-2027 Thang 7.2026.docx'
TPL_THCS = r'd:\UNIGO\Hệ thống mẫu văn bản\PL4-Khung kế hoạch bài dạy (THCS).docx'
OUT_BASE = r'd:\UNIGO\KHBD_Tin_học'

# ─── FONT HELPER ───────────────────────────────────────────────────────────
def afont(run, bold=None, italic=None, size_pt=13, color=None):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    if bold is not None: run.bold = bold
    if italic is not None: run.italic = italic
    if color: run.font.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rFonts.set(qn(attr), 'Times New Roman')

def set_line_spacing(para, ratio):
    """Set line spacing by ratio (1.5 → 360, 1.15 → 276)"""
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:line'), str(int(240 * ratio)))
    sp.set(qn('w:lineRule'), 'auto')

def add_para(doc, text='', bold=False, italic=False, align=None,
             indent_first=None, indent_left=None, line_ratio=None,
             sp_before=None, sp_after=None, size_pt=13):
    p = doc.add_paragraph()
    if text:
        r = p.add_run(text)
        afont(r, bold=bold, italic=italic, size_pt=size_pt)
    if align is not None: p.alignment = align
    pf = p.paragraph_format
    if indent_first is not None: pf.first_line_indent = Emu(indent_first)
    if indent_left is not None: pf.left_indent = Emu(indent_left)
    if sp_before is not None: pf.space_before = Pt(sp_before)
    if sp_after is not None: pf.space_after = Pt(sp_after)
    if line_ratio is not None: set_line_spacing(p, line_ratio)
    return p

# ─── TABLE BORDERS ─────────────────────────────────────────────────────────
def set_table_borders(table):
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        b.set(qn('w:space'), '0')
        tblBorders.append(b)
    tblPr.append(tblBorders)

def fill_cell(cell, text, bold=False, italic=False, line_ratio=1.5, align=None):
    for p in cell.paragraphs:
        for r in p.runs: r.text = ''
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i > 0: p = cell.add_paragraph()
        if line:
            r = p.add_run(line)
            afont(r, bold=bold, italic=italic)
    if line_ratio: set_line_spacing(cell.paragraphs[0], line_ratio)
    if align: cell.paragraphs[0].alignment = align

# ─── SAVE HELPER ───────────────────────────────────────────────────────────
def save_doc(doc, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        doc.save(path)
        return path
    except PermissionError:
        alt = path.replace('.docx', '_new.docx')
        doc.save(alt)
        return alt

# ─── CLEAN BODY ────────────────────────────────────────────────────────────
def clean_body(doc):
    """Xóa toàn bộ body elements, giữ sectPr"""
    body = doc.element.body
    for child in list(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag != 'sectPr':
            body.remove(child)

# ════════════════════════════════════════════════════════════════════════════
# PHẦN A: TẠO KHBD TIỂU HỌC
# Theo luật KHBD_TIEU_HOC.md:
# - Font: TNR 13pt, Line spacing: 1.5, Indent: 457200 EMU
# - Mục tiêu: Phẩm chất TRƯỚC → Năng lực SAU
# - Bảng 2 cột GV/HS với hàng gộp gridSpan=2 cho tiêu đề HĐ
# ════════════════════════════════════════════════════════════════════════════

def add_merged_row_th(table, text_lines, bold=True):
    """Thêm hàng tiêu đề gộp 2 cột cho TH"""
    row = table.add_row()
    tc0 = row.cells[0]._tc
    tc1 = row.cells[1]._tc
    # Set gridSpan=2
    tcPr = tc0.get_or_add_tcPr()
    gs = OxmlElement('w:gridSpan')
    gs.set(qn('w:val'), '2')
    tcPr.append(gs)
    # Remove 2nd cell
    row._tr.remove(tc1)
    # Clear existing paragraph in tc0 and fill text
    existing_p = tc0.find(qn('w:p'))
    if existing_p is not None:
        tc0.remove(existing_p)
    for i, line in enumerate(text_lines):
        new_p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        # line spacing 1.5
        sp_el = OxmlElement('w:spacing')
        sp_el.set(qn('w:line'), '360')
        sp_el.set(qn('w:lineRule'), 'auto')
        pPr.append(sp_el)
        new_p.append(pPr)
        if line:
            new_r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            if bold:
                b_el = OxmlElement('w:b')
                rPr.append(b_el)
            # Font
            rFonts = OxmlElement('w:rFonts')
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                rFonts.set(qn(attr), 'Times New Roman')
            rPr.append(rFonts)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '26')  # 13pt = 26 half-points
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), '26')
            rPr.append(szCs)
            new_r.append(rPr)
            t_el = OxmlElement('w:t')
            t_el.text = line
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            new_r.append(t_el)
            new_p.append(new_r)
        tc0.append(new_p)
    return row

def add_content_row_th(table, gv_lines, hs_lines):
    """Thêm hàng nội dung GV | HS"""
    row = table.add_row()
    fill_cell(row.cells[0], '\n'.join(gv_lines), line_ratio=1.5)
    fill_cell(row.cells[1], '\n'.join(hs_lines), line_ratio=1.5)
    return row

def create_khbd_th(grade_name, grade_folder, bai_so, ten_bai, chu_diem,
                   tiet_ppct, ngay_day,
                   pham_chat_items, nang_luc_mon_items, nang_luc_chung_items,
                   do_dung_gv, do_dung_hs,
                   phuong_phap, ki_thuat,
                   activities):
    """
    Tạo KHBD Tiểu học hoàn chỉnh.
    activities: list of dicts:
      {
        'ten': 'Tên hoạt động (...phút)',
        'muc_tieu': 'Mục tiêu',
        'sub': [  # Sub-hoạt động trong cùng nhóm (cho HĐ2, HĐ3)
          {'ten': '2.1. Tên HĐ', 'gv': [...], 'hs': [...]},
          ...
        ],
        'gv': ['GV làm...'],  # Nếu không có sub
        'hs': ['HS làm...'],
      }
    """
    doc = Document(TPL_TH)
    clean_body(doc)

    # ── Phần đầu ─────────────────────────────────────────────────────────
    # P[0]: Ngày
    p = add_para(doc, f'Thứ …… ngày {ngay_day}',
                 align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.5)
    # Italic cho phần thứ
    p.clear()
    r0 = p.add_run('Thứ …… ngày ')
    afont(r0, italic=True)
    r1 = p.add_run(ngay_day)
    afont(r1, italic=True)
    set_line_spacing(p, 1.5)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # P[1]: GV
    p1 = add_para(doc, 'Họ và tên Giáo viên: Đậu Đình Nguyên',
                  align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.5)

    # P[2]: Tên môn
    add_para(doc, 'KẾ HOẠCH DẠY HỌC MÔN TIN HỌC',
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.5)

    # P[3]: Chủ điểm
    p3 = add_para(doc, line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_cd_label = p3.add_run('CHỦ ĐIỂM: ')
    afont(r_cd_label, bold=True)
    r_cd_val = p3.add_run(chu_diem)
    afont(r_cd_val, bold=True)

    # P[4]: Bài + Tiết PPCT
    p4 = add_para(doc, line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_bai = p4.add_run(f'BÀI: {ten_bai}  ')
    afont(r_bai, bold=True)
    r_tiet_l = p4.add_run('(Tiết: ')
    afont(r_tiet_l)
    r_tiet_v = p4.add_run(str(tiet_ppct))
    afont(r_tiet_v, italic=True)
    r_tiet_r = p4.add_run(')')
    afont(r_tiet_r)

    # P[5]: Trống
    add_para(doc, '')

    # ── I. YÊU CẦU CẦN ĐẠT ──────────────────────────────────────────────
    add_para(doc, 'I. YÊU CẦU CẦN ĐẠT:', bold=True, line_ratio=1.5,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, '- Sau tiết học, học sinh sẽ:', bold=True,
             indent_first=457200, line_ratio=1.5,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 1. Phẩm chất (TRƯỚC)
    add_para(doc, '1. Phát triển phẩm chất', bold=True,
             indent_first=457200, line_ratio=1.5,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    for pc in pham_chat_items:
        add_para(doc, pc, indent_first=450215, line_ratio=1.33,
                 align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 2. Năng lực (SAU)
    add_para(doc, '2. Phát triển năng lực', bold=True,
             indent_first=457200, line_ratio=1.5,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, '2.1. Năng lực môn học:', indent_first=457200,
             line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    for nl in nang_luc_mon_items:
        add_para(doc, nl, indent_first=450215, line_ratio=1.33,
                 align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, '2.2. Năng lực chung và đặc thù:', indent_first=457200,
             line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    for nl in nang_luc_chung_items:
        add_para(doc, nl, indent_first=450215, line_ratio=1.33,
                 align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # ── II. ĐỒ DÙNG ──────────────────────────────────────────────────────
    add_para(doc, 'II. ĐỒ DÙNG DẠY HỌC :', bold=True, line_ratio=None)
    add_para(doc, f'\t1. Giáo viên: {do_dung_gv}', line_ratio=1.5,
             indent_left=228600)
    add_para(doc, f'2. Học sinh: {do_dung_hs}', line_ratio=1.5,
             indent_first=457200, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # ── III. PHƯƠNG PHÁP ─────────────────────────────────────────────────
    add_para(doc, 'III. PHƯƠNG PHÁP, KĨ THUẬT DẠY HỌC', bold=True,
             line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, f'- Phương pháp: {phuong_phap}',
             indent_first=457200, line_ratio=None,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, f'\t- Kĩ thuật: {ki_thuat}',
             line_ratio=None)

    # ── IV. CÁC HOẠT ĐỘNG DẠY-HỌC ────────────────────────────────────────
    add_para(doc, 'IV. CÁC HOẠT ĐỘNG DẠY - HỌC CHỦ YẾU: ', bold=True,
             line_ratio=1.5)

    # Tạo bảng 2 cột
    table = doc.add_table(rows=1, cols=2)
    set_table_borders(table)

    # Header row
    hdr_row = table.rows[0]
    fill_cell(hdr_row.cells[0], 'Hoạt động của GV', bold=True,
              line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(hdr_row.cells[1], 'Hoạt động của HS', bold=True,
              line_ratio=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Điền các hoạt động
    for act in activities:
        if act.get('sub'):
            # Tiêu đề nhóm HĐ (gộp 2 cột)
            header_lines = [act['ten']]
            if act.get('muc_tieu'):
                header_lines.append(f'*Mục tiêu: {act["muc_tieu"]}')
            add_merged_row_th(table, header_lines)
            for sub in act['sub']:
                sub_header = [sub['ten']]
                if sub.get('muc_tieu'):
                    sub_header.append(f'*Mục tiêu: {sub["muc_tieu"]}')
                add_merged_row_th(table, sub_header)
                add_content_row_th(table, sub.get('gv', ['']), sub.get('hs', ['']))
        else:
            # Hàng đơn gộp tiêu đề
            header_lines = [act['ten']]
            if act.get('muc_tieu'):
                header_lines.append(f'*Mục tiêu: {act["muc_tieu"]}')
            if len(activities) == 1 or not act.get('sub'):
                # Nếu là HĐ cuối (Vận dụng), không gộp mà để 2 cột
                add_content_row_th(table,
                    [act['ten']] + act.get('gv', ['']),
                    act.get('hs', ['...............']))
            else:
                add_merged_row_th(table, header_lines)
                add_content_row_th(table, act.get('gv', ['']), act.get('hs', ['']))

    # ── V. ĐIỀU CHỈNH BỔ SUNG — copy từ template gốc ─────────────────────
    doc_tpl = Document(TPL_TH)
    tpl_paras = doc_tpl.paragraphs
    # P[33] là "V. ĐIỀU CHỈNH..." đến P[43]
    tail_start = 33
    for i in range(tail_start, min(len(tpl_paras), 44)):
        new_p = copy.deepcopy(tpl_paras[i]._p)
        # Chèn trước sectPr
        sect_pr = doc.element.body.find(qn('w:sectPr'))
        if sect_pr is not None:
            doc.element.body.insert(list(doc.element.body).index(sect_pr), new_p)
        else:
            doc.element.body.append(new_p)

    # ── Save ──────────────────────────────────────────────────────────────
    out_dir = os.path.join(OUT_BASE, grade_folder, f'Bài_{bai_so:02d}')
    filename = f'KHBD_Tin_hoc_{grade_folder}_Bai{bai_so:02d}_{ten_bai.replace(" ", "_")[:40]}.docx'
    out_path = os.path.join(out_dir, filename)
    saved = save_doc(doc, out_path)

    # ── Verify ────────────────────────────────────────────────────────────
    verify = Document(saved)
    hdr_drawings = sum(1 for r in verify.sections[0].header.paragraphs[0].runs
                       if r._r.findall(qn('w:drawing')))
    ftr_paras = len(verify.sections[0].footer.paragraphs)
    empty_start = sum(1 for p in verify.paragraphs[:10] if not p.text.strip())
    print(f'  ✅ Saved: {saved}')
    print(f'     Header drawings: {hdr_drawings} | Footer paras: {ftr_paras} | Empty start: {empty_start}')
    return saved


# ════════════════════════════════════════════════════════════════════════════
# PHẦN B: TẠO KHBD THCS
# Theo luật KHBD_THCS.md (Phương án A):
# - Font: TNR 13pt, Line spacing: 1.15, Indent mục con: 180340 EMU
# - Mục tiêu: Kiến thức → Năng lực → Phẩm chất
# - Tiến trình: Paragraphs Bước 1-4 (KHÔNG dùng bảng)
# - GIỮ NGUYÊN Table[0] thông tin và Table[2] ký tên từ template
# ════════════════════════════════════════════════════════════════════════════

def create_khbd_thcs(grade_name, grade_folder, bai_so, ten_bai,
                     mon_hoc, lop, thoi_luong, tiet_ppct,
                     ngay_soan, ngay_day,
                     kien_thuc_items, nang_luc_chung_items,
                     nang_luc_dac_thu_items, nang_luc_so_items,
                     pham_chat_items,
                     thiet_bi, hoc_lieu,
                     hoat_dong_list):
    """
    Tạo KHBD THCS hoàn chỉnh theo Phương án A.
    hoat_dong_list: list of dicts:
      {
        'stt': 1,
        'ten': 'Khởi động (Xác định vấn đề/...)',
        'muc_tieu': '...',
        'noi_dung': '...',
        'san_pham': '...',
        'buoc1': '...',
        'buoc2': '...',
        'buoc3': '...',
        'buoc4': '...',
      }
    """
    doc = Document(TPL_THCS)
    clean_body(doc)

    # ── Table[0]: Thông tin trường/GV — copy từ template ──────────────────
    doc_tpl = Document(TPL_THCS)
    tbl_info = copy.deepcopy(doc_tpl.tables[0]._tbl)

    # Sửa ngày trong table copy
    # Row[1], Col[1] chứa ngày soạn/dạy
    rows = tbl_info.findall(qn('w:tr'))
    if len(rows) >= 2:
        cells_r1 = rows[1].findall(qn('w:tc'))
        if len(cells_r1) >= 2:
            tc_ngay = cells_r1[1]
            for p_el in tc_ngay.findall(qn('w:p')):
                for r_el in p_el.findall(qn('w:r')):
                    t_el = r_el.find(qn('w:t'))
                    if t_el is not None and t_el.text:
                        if 'Ngày soạn' in t_el.text:
                            t_el.text = f'Ngày soạn: {ngay_soan}   Ngày dạy: {ngay_day}'

    # Insert table vào body
    sect_pr = doc.element.body.find(qn('w:sectPr'))
    if sect_pr is not None:
        idx = list(doc.element.body).index(sect_pr)
        doc.element.body.insert(idx, tbl_info)
    else:
        doc.element.body.append(tbl_info)

    # ── Tên bài + Tiết PPCT ───────────────────────────────────────────────
    p_ten = add_para(doc, f'TÊN BÀI DẠY: {ten_bai.upper()}',
                     bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.15)
    add_para(doc, f'Môn học: {mon_hoc}   Lớp: {lop}   Thời lượng: {thoi_luong}',
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.15)
    add_para(doc, f'Tiết theo PPCT: {tiet_ppct}',
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.15)
    add_para(doc, f'Tên tiết: {ten_bai}',
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.15)

    # ── I. MỤC TIÊU ──────────────────────────────────────────────────────
    add_para(doc, 'I. Mục tiêu', bold=True, line_ratio=1.15)

    add_para(doc, '1. Kiến thức: ', bold=True,
             indent_first=180340, line_ratio=1.15)
    for kt in kien_thuc_items:
        add_para(doc, kt, indent_first=180340, sp_after=0, line_ratio=1.15)

    p_nl = add_para(doc, '', indent_first=180340, line_ratio=1.15)
    r_nl_label = p_nl.add_run('2. Năng lực: ')
    afont(r_nl_label, bold=True)
    r_nl_desc = p_nl.add_run('Nêu cụ thể yêu cầu học sinh làm được gì')
    afont(r_nl_desc)

    add_para(doc, '- Năng lực chung:', bold=True,
             indent_first=180340, sp_after=0, line_ratio=1.15)
    for nl in nang_luc_chung_items:
        add_para(doc, f'  {nl}', indent_first=180340, sp_after=0, line_ratio=1.15)

    add_para(doc, '- Năng lực đặc thù:', bold=True,
             indent_first=180340, sp_after=0, line_ratio=1.15)
    for nl in nang_luc_dac_thu_items:
        add_para(doc, f'  {nl}', indent_first=180340, sp_after=0, line_ratio=1.15)

    add_para(doc, '- Năng lực số:', bold=True,
             indent_first=180340, sp_after=0, line_ratio=1.15)
    for nl in nang_luc_so_items:
        add_para(doc, f'  {nl}', indent_first=180340, sp_after=0, line_ratio=1.15)

    add_para(doc, '3. Phẩm chất: ', bold=True,
             indent_first=180340, line_ratio=1.15)
    for pc in pham_chat_items:
        add_para(doc, pc, indent_first=180340, sp_after=0, line_ratio=1.15)

    # ── II. THIẾT BỊ ─────────────────────────────────────────────────────
    p_tb = add_para(doc, '', line_ratio=1.15)
    r_tb = p_tb.add_run('II. Thiết bị dạy học và học liệu: ')
    afont(r_tb, bold=True)

    add_para(doc, '1. Thiết bị', bold=True,
             indent_first=180340, sp_after=0, line_ratio=1.15)
    add_para(doc, thiet_bi, indent_first=180340, sp_after=0, line_ratio=1.15)

    add_para(doc, '2. Học liệu', bold=True,
             indent_first=180340, sp_after=0, line_ratio=1.15)
    add_para(doc, hoc_lieu, indent_first=180340, sp_after=0, line_ratio=1.15)

    # ── III. TIẾN TRÌNH DẠY HỌC (paragraphs, KHÔNG bảng) ─────────────────
    add_para(doc, 'III. Tiến trình dạy học', bold=True, line_ratio=1.15)

    for hd in hoat_dong_list:
        stt = hd['stt']
        ten_hd = hd['ten']
        is_last = (stt == len(hoat_dong_list))

        add_para(doc, f'{stt}. Hoạt động {stt}. {ten_hd}', bold=True,
                 indent_first=180340, line_ratio=1.15)

        # a) b) c) d)
        p_a = add_para(doc, '', indent_first=180340, sp_after=0, line_ratio=1.15)
        r_a_label = p_a.add_run('a) Mục tiêu: ')
        afont(r_a_label, italic=True)
        r_a_val = p_a.add_run(hd.get('muc_tieu', ''))
        afont(r_a_val)

        p_b = add_para(doc, '', indent_first=180340, sp_after=0, line_ratio=1.15)
        r_b_label = p_b.add_run('b) Nội dung: ')
        afont(r_b_label, italic=True)
        r_b_val = p_b.add_run(hd.get('noi_dung', ''))
        afont(r_b_val)

        p_c = add_para(doc, '', indent_first=180340, sp_after=0, line_ratio=1.15)
        r_c_label = p_c.add_run('c) Sản phẩm: ')
        afont(r_c_label, italic=True)
        r_c_val = p_c.add_run(hd.get('san_pham', ''))
        afont(r_c_val)

        p_d = add_para(doc, '', indent_first=180340, line_ratio=1.15)
        r_d_label = p_d.add_run('d) Tổ chức thực hiện: ')
        afont(r_d_label, italic=True)
        r_d_val = p_d.add_run(hd.get('to_chuc', ''))
        afont(r_d_val)

        # Bước 1-4
        buoc_labels_default = [
            'Bước 1. Chuyển giao nhiệm vụ học tập',
            'Bước 2. Học sinh tiếp nhận nhiệm vụ học tập',
            'Bước 3. Báo cáo kết quả hoạt động',
            'Bước 4. Đánh giá kết quả thực hiện nhiệm vụ',
        ]
        buoc_labels_last = [
            'Bước 1. Chuyển giao nhiệm vụ học tập',
            'Bước 2. Học sinh tiếp nhận nhiệm vụ học tập',
            'Bước 3. Báo cáo kết quả hoạt động',
            'Bước 4. Giáo viên nhắc nhở nhiệm vụ về nhà',
        ]
        buoc_labels = buoc_labels_last if is_last else buoc_labels_default
        buoc_contents = [
            hd.get('buoc1', ''), hd.get('buoc2', ''),
            hd.get('buoc3', ''), hd.get('buoc4', ''),
        ]

        for j, (label, content) in enumerate(zip(buoc_labels, buoc_contents)):
            p_buoc = add_para(doc, '', indent_first=360045,
                              sp_after=0, line_ratio=1.15)
            r_bl = p_buoc.add_run(label)
            afont(r_bl, italic=True)
            if content:
                r_bc = p_buoc.add_run(f'\n{content}')
                afont(r_bc)

    # ── RÚT KINH NGHIỆM ──────────────────────────────────────────────────
    add_para(doc, 'RÚT KINH NGHIỆM SAU BÀI DẠY', bold=True, line_ratio=1.15)
    add_para(doc, '….……………………………………………………………………………………………..', bold=True, line_ratio=1.15)
    add_para(doc, 'Lưu ý: Sau 1 tuần mới để phần kí',
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, line_ratio=1.15)
    add_para(doc, '')
    add_para(doc, '', indent_first=180340, line_ratio=1.15)

    # ── Table[2]: Bảng ký tên — copy từ template ─────────────────────────
    tbl_sign = copy.deepcopy(doc_tpl.tables[2]._tbl)
    sect_pr2 = doc.element.body.find(qn('w:sectPr'))
    if sect_pr2 is not None:
        idx2 = list(doc.element.body).index(sect_pr2)
        doc.element.body.insert(idx2, tbl_sign)
    else:
        doc.element.body.append(tbl_sign)

    # ── Save ──────────────────────────────────────────────────────────────
    out_dir = os.path.join(OUT_BASE, grade_folder, f'Bài_{bai_so:02d}')
    filename = f'KHBD_Tin_hoc_{grade_folder}_Bai{bai_so:02d}_{ten_bai.replace(" ", "_")[:40]}.docx'
    out_path = os.path.join(out_dir, filename)
    saved = save_doc(doc, out_path)

    # ── Verify ────────────────────────────────────────────────────────────
    verify = Document(saved)
    hdr = verify.sections[0].header
    hdr_drawings = sum(1 for p in hdr.paragraphs
                       for r in p.runs if r._r.findall(qn('w:drawing')))
    tbl_count = len(verify.tables)
    print(f'  ✅ Saved: {saved}')
    print(f'     Header drawings: {hdr_drawings} | Tables: {tbl_count}')
    return saved


# ════════════════════════════════════════════════════════════════════════════
# DỮ LIỆU NỘI DUNG CHO TỪNG LỚP - Bài 1
# ════════════════════════════════════════════════════════════════════════════

def build_all():
    results = []

    # ────────────────────────────────────────────────────────────────────
    # TIỀN TIỂU HỌC — Bài 1: Máy tính xung quanh em
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Tiền tiểu học - Bài 1...')
    saved = create_khbd_th(
        grade_name='Tiền tiểu học', grade_folder='Tiền_tiểu_học',
        bai_so=1, ten_bai='BÀI 1. MÁY TÍNH XUNG QUANH EM', chu_diem='Làm quen với công nghệ',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Tích cực tham gia các hoạt động học tập, hoàn thành nhiệm vụ được giao.',
            '- Nhân ái: Biết chia sẻ, hợp tác với bạn bè trong các hoạt động nhóm.',
            '- Trách nhiệm: Giữ gìn thiết bị máy tính, sử dụng đúng cách.',
        ],
        nang_luc_mon_items=[
            '- NLa (Sử dụng và quản lí thiết bị công nghệ số): Nhận diện được hình dạng và chức năng cơ bản của máy tính.',
            '- NLb (Kết nối và cộng tác): Biết làm việc cùng bạn trong các hoạt động khám phá thiết bị.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Quan sát, nhận biết các thiết bị xung quanh.',
            '- Giao tiếp và hợp tác: Chia sẻ những gì em biết về máy tính với bạn.',
        ],
        do_dung_gv='Máy chiếu, máy tính, tranh ảnh về các loại máy tính, thẻ từ hình ảnh thiết bị.',
        do_dung_hs='Bút, vở, tranh ảnh chuẩn bị sẵn về máy tính.',
        phuong_phap='vấn đáp, quan sát, trò chơi học tập, hoạt động nhóm đôi',
        ki_thuat='đặt câu hỏi, trình bày 1 phút, chia sẻ nhóm đôi',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Tạo hứng thú, kết nối bài học với cuộc sống.',
                'gv': [
                    'GV chiếu hình ảnh nhiều loại máy tính (laptop, máy tính bảng, điện thoại).',
                    'GV hỏi: "Em đã thấy những thiết bị này ở đâu chưa?"',
                    '* Nhận xét, dẫn vào bài: Hôm nay chúng ta cùng khám phá máy tính xung quanh em!',
                ],
                'hs': [
                    '- HS quan sát hình ảnh trên màn chiếu.',
                    '- HS xung phong trả lời: ở nhà, ở trường, trong siêu thị...',
                    '- HS lắng nghe, vỗ tay.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'muc_tieu': '',
                'sub': [
                    {
                        'ten': '2.1. Nhận biết các loại máy tính (10 phút)',
                        'muc_tieu': 'HS phân biệt được laptop, máy tính bảng và điện thoại thông minh.',
                        'gv': [
                            'GV giới thiệu từng loại máy tính qua hình ảnh lớn trên màn chiếu.',
                            'GV đặt câu hỏi: "Cái này gọi là gì? Để làm gì?"',
                            '* GV kết luận tên gọi và chức năng từng loại.',
                        ],
                        'hs': [
                            '- HS quan sát, lắng nghe.',
                            '- HS đoán tên và trả lời theo hiểu biết.',
                            '- HS lắp lại tên gọi theo GV.',
                        ],
                    },
                    {
                        'ten': '2.2. Các bộ phận của máy tính (5 phút)',
                        'muc_tieu': 'HS chỉ đúng màn hình, bàn phím và chuột máy tính.',
                        'gv': [
                            'GV chỉ vào máy tính thật, giới thiệu: màn hình, bàn phím, chuột.',
                            'GV cho HS lần lượt lên chỉ và gọi tên bộ phận.',
                        ],
                        'hs': [
                            '- HS quan sát, lắng nghe.',
                            '- Từng HS lên chỉ và nói tên bộ phận.',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'muc_tieu': '',
                'sub': [
                    {
                        'ten': '3.1. Trò chơi: Tìm đúng tên (7 phút)',
                        'muc_tieu': 'Củng cố nhận biết tên các thiết bị máy tính.',
                        'gv': [
                            'GV phát thẻ hình ảnh thiết bị, HS ghép với thẻ tên tương ứng.',
                            '* Nhận xét, chữa bài, tuyên dương nhóm đúng nhanh nhất.',
                        ],
                        'hs': [
                            '- HS làm việc nhóm đôi, ghép thẻ hình-tên.',
                            '- HS đổi bài, kiểm tra chéo.',
                        ],
                    },
                    {
                        'ten': '3.2. Em tô màu máy tính (3 phút)',
                        'muc_tieu': 'HS ghi nhớ hình dạng thiết bị qua hoạt động sáng tạo.',
                        'gv': [
                            'GV phát phiếu tô màu hình máy tính.',
                            '* Nhận xét sản phẩm của HS.',
                        ],
                        'hs': [
                            '- HS tô màu theo ý thích.',
                            '- HS giơ bài, cùng chia sẻ với bạn ngồi cạnh.',
                        ],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV hỏi: "Về nhà, em thấy máy tính/điện thoại ở đâu? Ai dùng?"',
                    'GV dặn dò: Về nhà kể cho bố mẹ nghe tên 3 bộ phận của máy tính.',
                ],
                'hs': [
                    '- HS chia sẻ trước lớp.',
                    '- HS lắng nghe dặn dò.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 1 — Bài 1
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Lớp 1 - Bài 1...')
    saved = create_khbd_th(
        grade_name='Lớp 1', grade_folder='Lớp_1',
        bai_so=1, ten_bai='BÀI 1. CHIẾC MÁY TÍNH CỦA EM', chu_diem='Em và máy tính',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Tích cực học tập, hoàn thành bài tập đúng thời gian.',
            '- Trách nhiệm: Giữ gìn và sử dụng thiết bị đúng cách.',
            '- Nhân ái: Biết nhường bạn, hợp tác trong học nhóm.',
        ],
        nang_luc_mon_items=[
            '- NLa: Nhận biết và gọi tên được các bộ phận chính của máy tính (màn hình, bàn phím, chuột, thân máy).',
            '- NLb: Biết cách bật/tắt máy tính đúng quy trình.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Chủ động thực hiện các thao tác được hướng dẫn.',
            '- Giao tiếp và hợp tác: Trao đổi, chia sẻ với bạn về các bộ phận máy tính.',
        ],
        do_dung_gv='Máy chiếu, máy tính demo, tranh poster các bộ phận máy tính.',
        do_dung_hs='SGK Tin học 1, bút chì.',
        phuong_phap='vấn đáp, quan sát trực quan, thực hành, hoạt động nhóm đôi',
        ki_thuat='đặt câu hỏi, tia chớp, chia sẻ nhóm đôi',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Kết nối kiến thức cũ, tạo hứng thú học bài mới.',
                'gv': [
                    'GV chiếu hình ảnh câu đố: "Đây là vật gì? Dùng để làm gì?"',
                    'GV mời HS trả lời, dẫn vào bài học.',
                ],
                'hs': [
                    '- HS quan sát, đoán và xung phong trả lời.',
                    '- HS cùng đọc tên bài học.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'sub': [
                    {
                        'ten': '2.1. Khám phá các bộ phận của máy tính (10 phút)',
                        'muc_tieu': 'HS nhận biết và gọi tên đúng các bộ phận chính của máy tính.',
                        'gv': [
                            'GV chỉ từng bộ phận trên máy tính thật và màn chiếu.',
                            'GV gọi HS lên chỉ và đọc tên.',
                            '* Kết luận: Máy tính gồm màn hình, bàn phím, chuột và thân máy.',
                        ],
                        'hs': [
                            '- HS quan sát, lắng nghe.',
                            '- HS lần lượt lên chỉ và đọc tên bộ phận.',
                            '- HS đọc đồng thanh tên các bộ phận.',
                        ],
                    },
                    {
                        'ten': '2.2. Quy trình bật và tắt máy tính (5 phút)',
                        'muc_tieu': 'HS biết thứ tự các bước bật và tắt máy tính đúng cách.',
                        'gv': [
                            'GV thao tác mẫu bật máy, giải thích từng bước.',
                            'GV thao tác mẫu tắt máy đúng quy trình.',
                        ],
                        'hs': [
                            '- HS quan sát theo dõi từng bước.',
                            '- HS nhắc lại thứ tự bước bật/tắt máy.',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'sub': [
                    {
                        'ten': '3.1. Nối hình với tên (5 phút)',
                        'muc_tieu': 'Củng cố nhận biết tên các bộ phận máy tính.',
                        'gv': [
                            'GV phát phiếu bài tập: HS nối hình bộ phận với tên gọi đúng.',
                            '* GV chữa bài, nhận xét.',
                        ],
                        'hs': [
                            '- HS làm bài tập nối hình.',
                            '- HS đổi bài kiểm tra chéo, báo cáo kết quả.',
                        ],
                    },
                    {
                        'ten': '3.2. Thực hành bật/tắt máy (5 phút)',
                        'muc_tieu': 'HS thực hiện được thao tác bật/tắt máy đúng quy trình.',
                        'gv': [
                            'GV hướng dẫn HS thực hành theo từng bước.',
                            '* GV quan sát, hỗ trợ HS gặp khó khăn.',
                        ],
                        'hs': [
                            '- HS thực hành bật máy theo hướng dẫn.',
                            '- HS thực hành tắt máy đúng quy trình.',
                        ],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV hỏi: "Em có thể kể tên 3 bộ phận của máy tính cho bố mẹ nghe không?"',
                    'GV dặn: Về nhà quan sát máy tính ở gia đình.',
                ],
                'hs': [
                    '- HS thi đua kể tên bộ phận nhanh nhất.',
                    '- HS lắng nghe và ghi nhớ nhiệm vụ về nhà.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 2 — Bài 1
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Lớp 2 - Bài 1...')
    saved = create_khbd_th(
        grade_name='Lớp 2', grade_folder='Lớp_2',
        bai_so=1, ten_bai='BÀI 1. MÁY TÍNH LÀ NGƯỜI BẠN CỦA EM', chu_diem='Máy tính và cuộc sống',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Chủ động tìm hiểu và ghi nhớ các ứng dụng của máy tính.',
            '- Trách nhiệm: Sử dụng máy tính đúng mục đích, không lạm dụng.',
            '- Nhân ái: Biết chia sẻ kiến thức về máy tính với bạn bè.',
        ],
        nang_luc_mon_items=[
            '- NLa: Biết được máy tính có thể giúp em học tập, giải trí và kết nối với mọi người.',
            '- NLb: Nhận biết được một số ứng dụng phổ biến trên máy tính.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Tự tìm hiểu và chia sẻ những điều em biết về máy tính.',
            '- Giao tiếp và hợp tác: Thảo luận nhóm về cách máy tính giúp ích cuộc sống.',
        ],
        do_dung_gv='Máy chiếu, video ngắn về ứng dụng máy tính trong cuộc sống, phiếu học tập.',
        do_dung_hs='SGK Tin học 2, bút màu.',
        phuong_phap='vấn đáp, xem video, thảo luận nhóm, trình bày',
        ki_thuat='động não, chia sẻ nhóm đôi, trình bày 1 phút',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Khơi gợi hiểu biết của HS về lợi ích của máy tính.',
                'gv': [
                    'GV hỏi: "Các em đã dùng máy tính để làm gì rồi?"',
                    'GV ghi nhanh các ý kiến lên bảng, dẫn vào bài.',
                ],
                'hs': [
                    '- HS tự do phát biểu.',
                    '- HS lắng nghe, bổ sung ý kiến của nhau.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'sub': [
                    {
                        'ten': '2.1. Máy tính giúp em học tập (8 phút)',
                        'muc_tieu': 'HS biết được máy tính hỗ trợ học tập như thế nào.',
                        'gv': [
                            'GV chiếu video ngắn: HS dùng máy tính tra từ điển, đọc sách điện tử, làm bài tập online.',
                            'GV đặt câu hỏi: "Máy tính giúp em học tập như thế nào?"',
                            '* Kết luận: Máy tính là công cụ học tập hiệu quả.',
                        ],
                        'hs': [
                            '- HS xem video, quan sát.',
                            '- HS thảo luận nhóm đôi trả lời câu hỏi.',
                            '- Đại diện nhóm chia sẻ.',
                        ],
                    },
                    {
                        'ten': '2.2. Máy tính giúp em vui chơi và kết nối (7 phút)',
                        'muc_tieu': 'HS biết máy tính có thể dùng để giải trí và kết nối an toàn.',
                        'gv': [
                            'GV giới thiệu: máy tính dùng nghe nhạc, xem phim, chơi game học tập, video call.',
                            'GV nhấn mạnh: cần dùng đúng cách, đúng thời gian.',
                        ],
                        'hs': [
                            '- HS lắng nghe, ghi nhớ.',
                            '- HS chia sẻ: "Em đã dùng máy tính để làm gì ở nhà?"',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'sub': [
                    {
                        'ten': '3.1. Phân loại việc dùng máy tính (8 phút)',
                        'muc_tieu': 'HS phân biệt được dùng máy tính đúng cách và chưa đúng cách.',
                        'gv': [
                            'GV phát phiếu: HS phân loại các tình huống vào 2 cột Đúng/Chưa đúng.',
                            '* GV nhận xét, chốt đáp án.',
                        ],
                        'hs': [
                            '- HS đọc và phân loại tình huống.',
                            '- HS giải thích lý do lựa chọn.',
                        ],
                    },
                    {
                        'ten': '3.2. Vẽ máy tính của em (2 phút)',
                        'muc_tieu': 'HS thể hiện sự sáng tạo và ghi nhớ hình dạng thiết bị.',
                        'gv': ['GV hướng dẫn HS vẽ nhanh máy tính yêu thích vào vở.'],
                        'hs': ['- HS vẽ và tô màu.'],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV: "Em hãy kể cho bạn nghe 3 điều máy tính giúp ích cho em."',
                    'Dặn dò: Về nhà nói chuyện với bố mẹ về việc dùng máy tính đúng cách.',
                ],
                'hs': [
                    '- HS chia sẻ cặp đôi.',
                    '- Đại diện một số HS chia sẻ trước lớp.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 3 — Bài 1: Thông tin và quyết định
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Lớp 3 - Bài 1...')
    saved = create_khbd_th(
        grade_name='Lớp 3', grade_folder='Lớp_3',
        bai_so=1, ten_bai='BÀI 1. THÔNG TIN VÀ QUYẾT ĐỊNH', chu_diem='Thông tin quanh ta',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Chủ động tìm hiểu thông tin trong cuộc sống hằng ngày.',
            '- Trung thực: Nhận biết và không lan truyền thông tin sai lệch.',
            '- Trách nhiệm: Biết kiểm tra thông tin trước khi ra quyết định.',
        ],
        nang_luc_mon_items=[
            '- NLa (Sử dụng và quản lí phương tiện kĩ thuật số): Hiểu được vai trò của thông tin trong cuộc sống.',
            '- NLc (Giải quyết vấn đề với sự hỗ trợ của công nghệ thông tin): Biết thu thập và sử dụng thông tin để đưa ra quyết định đơn giản.',
            '- Năng lực số (CV 3456): Biết thông tin có nhiều dạng khác nhau (hình ảnh, âm thanh, văn bản).',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Thu thập thông tin từ nhiều nguồn khác nhau.',
            '- Giao tiếp và hợp tác: Chia sẻ thông tin và thảo luận quyết định cùng nhóm.',
            '- Giải quyết vấn đề và sáng tạo: Vận dụng thông tin để giải quyết tình huống thực tế.',
        ],
        do_dung_gv='Máy chiếu, tranh ảnh minh họa thông tin, phiếu bài tập tình huống.',
        do_dung_hs='SGK Tin học 3, bút, vở bài tập.',
        phuong_phap='vấn đáp, thảo luận nhóm, giải quyết vấn đề, trò chơi học tập',
        ki_thuat='đặt câu hỏi, động não, chia sẻ nhóm đôi',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Kích hoạt suy nghĩ về thông tin trong cuộc sống hằng ngày.',
                'gv': [
                    'GV kể tình huống: "Sáng nay, trước khi ra khỏi nhà, em xem dự báo thời tiết thấy trời sẽ mưa. Em quyết định mang ô theo. Tại sao em làm vậy?"',
                    'GV hỏi: "Dự báo thời tiết là loại thông tin gì? Nó giúp em quyết định điều gì?"',
                ],
                'hs': [
                    '- HS lắng nghe tình huống.',
                    '- HS phát biểu suy nghĩ.',
                    '- HS thảo luận cặp đôi, chia sẻ đáp án.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'sub': [
                    {
                        'ten': '2.1. Thông tin là gì? (8 phút)',
                        'muc_tieu': 'HS nêu được khái niệm thông tin và các dạng thông tin.',
                        'gv': [
                            'GV chiếu hình ảnh: biển báo giao thông, tiếng chuông, sách vở, video.',
                            'GV hỏi: "Mỗi hình ảnh truyền đến em điều gì?"',
                            '* Kết luận: Thông tin là những gì cho chúng ta biết về sự vật, hiện tượng. Thông tin có 3 dạng: văn bản, hình ảnh, âm thanh.',
                        ],
                        'hs': [
                            '- HS quan sát từng hình ảnh.',
                            '- HS trả lời câu hỏi theo suy nghĩ.',
                            '- HS ghi nhớ 3 dạng thông tin.',
                        ],
                    },
                    {
                        'ten': '2.2. Thông tin giúp ta quyết định (7 phút)',
                        'muc_tieu': 'HS hiểu mối liên hệ giữa thông tin và quyết định.',
                        'gv': [
                            'GV nêu ví dụ: "Nhìn đèn giao thông đỏ → quyết định dừng lại."',
                            'GV phát phiếu tình huống: HS viết thông tin → quyết định.',
                            '* GV chốt: Thông tin đúng giúp ta quyết định đúng.',
                        ],
                        'hs': [
                            '- HS lắng nghe ví dụ.',
                            '- HS làm bài tập tình huống.',
                            '- HS chia sẻ đáp án.',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'sub': [
                    {
                        'ten': '3.1. Trò chơi "Thông tin – Quyết định" (8 phút)',
                        'muc_tieu': 'Củng cố hiểu biết về mối quan hệ thông tin-quyết định.',
                        'gv': [
                            'GV phát thẻ: mỗi nhóm nhận thẻ thông tin và thẻ quyết định, ghép đôi cho đúng.',
                            '* GV chữa bài, tuyên dương nhóm làm đúng và nhanh.',
                        ],
                        'hs': [
                            '- HS làm việc nhóm 4, ghép thẻ.',
                            '- Đại diện nhóm giải thích lý do ghép.',
                        ],
                    },
                    {
                        'ten': '3.2. Viết tình huống của em (2 phút)',
                        'muc_tieu': 'HS vận dụng kiến thức vào tình huống thực tế.',
                        'gv': ['GV yêu cầu HS tự viết 1 ví dụ thông tin → quyết định trong cuộc sống em.'],
                        'hs': ['- HS viết vào vở.', '- 2-3 HS đọc ví dụ của mình.'],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV: "Hôm nay em đã dùng thông tin nào để quyết định điều gì?"',
                    'Dặn dò: Quan sát và ghi lại 3 tình huống dùng thông tin để quyết định trong ngày mai.',
                ],
                'hs': [
                    '- HS suy nghĩ và chia sẻ.',
                    '- HS lắng nghe dặn dò.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 4 — Bài 1: Phần cứng và phần mềm máy tính
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Lớp 4 - Bài 1...')
    saved = create_khbd_th(
        grade_name='Lớp 4', grade_folder='Lớp_4',
        bai_so=1, ten_bai='BÀI 1. PHẦN CỨNG VÀ PHẦN MỀM MÁY TÍNH', chu_diem='Cấu tạo máy tính',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Tích cực tìm hiểu, ghi nhớ kiến thức về phần cứng và phần mềm.',
            '- Trách nhiệm: Sử dụng đúng cách, bảo quản thiết bị phần cứng.',
            '- Trung thực: Sử dụng phần mềm hợp pháp, không sao chép trái phép.',
        ],
        nang_luc_mon_items=[
            '- NLa: Phân biệt được phần cứng và phần mềm máy tính; liệt kê được ví dụ cụ thể.',
            '- NLc: Biết phần cứng và phần mềm phải phối hợp mới tạo nên hệ thống hoàn chỉnh.',
            '- Năng lực số (CV 3456): Hiểu được cơ bản về hệ thống máy tính (hardware/software).',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Phân loại thiết bị máy tính vào đúng nhóm phần cứng/mềm.',
            '- Giao tiếp và hợp tác: Thảo luận nhóm về các ví dụ phần cứng, phần mềm.',
        ],
        do_dung_gv='Máy chiếu, ảnh/video các thiết bị phần cứng, màn hình giới thiệu phần mềm.',
        do_dung_hs='SGK Tin học 4, phiếu bài tập.',
        phuong_phap='vấn đáp, quan sát, thảo luận nhóm, phân loại',
        ki_thuat='đặt câu hỏi, sơ đồ tư duy, chia sẻ nhóm',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Kích thích sự tò mò về cấu tạo của máy tính.',
                'gv': [
                    'GV đặt câu hỏi: "Máy tính gồm những gì? Em có thể cầm lấy phần mềm không?"',
                    'GV dẫn vào bài: Hôm nay chúng ta tìm hiểu sự khác nhau giữa phần cứng và phần mềm.',
                ],
                'hs': [
                    '- HS suy nghĩ và trả lời.',
                    '- HS tò mò và nêu câu hỏi.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'sub': [
                    {
                        'ten': '2.1. Phần cứng là gì? (8 phút)',
                        'muc_tieu': 'HS định nghĩa và liệt kê được ví dụ phần cứng máy tính.',
                        'gv': [
                            'GV chiếu ảnh các thiết bị: CPU, RAM, ổ cứng, màn hình, bàn phím, chuột.',
                            'GV hỏi: "Điểm chung của những thứ này là gì?"',
                            '* Kết luận: Phần cứng là những bộ phận vật lí của máy tính — có thể nhìn thấy và chạm vào được.',
                        ],
                        'hs': [
                            '- HS quan sát hình ảnh.',
                            '- HS phát biểu nhận xét.',
                            '- HS ghi nhớ định nghĩa phần cứng.',
                        ],
                    },
                    {
                        'ten': '2.2. Phần mềm là gì? (7 phút)',
                        'muc_tieu': 'HS định nghĩa và liệt kê được ví dụ phần mềm.',
                        'gv': [
                            'GV mở máy tính, giới thiệu: Word, Chrome, Paint, hệ điều hành Windows.',
                            'GV hỏi: "Những thứ này khác phần cứng như thế nào?"',
                            '* Kết luận: Phần mềm là các chương trình chạy trên máy tính — không thể chạm vào.',
                        ],
                        'hs': [
                            '- HS quan sát màn hình máy tính.',
                            '- HS phát biểu sự khác biệt.',
                            '- HS cho thêm ví dụ phần mềm em biết.',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'sub': [
                    {
                        'ten': '3.1. Phân loại phần cứng/phần mềm (8 phút)',
                        'muc_tieu': 'Củng cố khả năng phân biệt phần cứng và phần mềm.',
                        'gv': [
                            'GV chiếu danh sách 10 ví dụ, HS xếp vào bảng phân loại.',
                            '* GV chữa bài, giải thích các trường hợp khó.',
                        ],
                        'hs': [
                            '- HS làm việc cá nhân, điền vào phiếu bài tập.',
                            '- HS so sánh kết quả với bạn cạnh.',
                        ],
                    },
                    {
                        'ten': '3.2. Sơ đồ tư duy (2 phút)',
                        'muc_tieu': 'HS tổng hợp kiến thức bài học bằng sơ đồ.',
                        'gv': ['GV hướng dẫn HS vẽ sơ đồ Máy tính → Phần cứng / Phần mềm.'],
                        'hs': ['- HS vẽ sơ đồ vào vở.'],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV: "Về nhà, em hãy quan sát và liệt kê 5 phần cứng và 5 phần mềm."',
                    'Dặn dò: Chia sẻ danh sách với bố mẹ, giải thích sự khác nhau.',
                ],
                'hs': [
                    '- HS lắng nghe nhiệm vụ về nhà.',
                    '- HS ghi vào vở.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 5 — Bài 1: Em có thể làm gì với máy tính
    # ────────────────────────────────────────────────────────────────────
    print('\n🔵 Tạo KHBD Lớp 5 - Bài 1...')
    saved = create_khbd_th(
        grade_name='Lớp 5', grade_folder='Lớp_5',
        bai_so=1, ten_bai='BÀI 1. EM CÓ THỂ LÀM GÌ VỚI MÁY TÍNH', chu_diem='Máy tính và cuộc sống',
        tiet_ppct='1', ngay_day='   /   /2026',
        pham_chat_items=[
            '- Chăm chỉ: Tích cực khám phá các ứng dụng hữu ích của máy tính trong học tập.',
            '- Trách nhiệm: Sử dụng máy tính đúng mục đích, không lướt web vô ích.',
            '- Trung thực: Thành thật về thời gian và cách sử dụng máy tính.',
        ],
        nang_luc_mon_items=[
            '- NLa: Biết và sử dụng được máy tính cho các hoạt động học tập, sáng tạo.',
            '- NLb: Biết cách tìm kiếm thông tin học tập an toàn trên Internet.',
            '- NLe: Nhận biết được những nguy cơ khi sử dụng máy tính không đúng cách.',
            '- Năng lực số (CV 3456): Sử dụng thiết bị và phần mềm cơ bản phục vụ học tập và sáng tạo.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Lên kế hoạch sử dụng máy tính hợp lí.',
            '- Giao tiếp và hợp tác: Chia sẻ kinh nghiệm sử dụng máy tính học tập với bạn.',
            '- Giải quyết vấn đề và sáng tạo: Đề xuất các cách dùng máy tính sáng tạo.',
        ],
        do_dung_gv='Máy chiếu, máy tính kết nối mạng, slide ví dụ ứng dụng máy tính.',
        do_dung_hs='SGK Tin học 5, vở ghi.',
        phuong_phap='thảo luận nhóm, trình bày, thực hành, giải quyết vấn đề',
        ki_thuat='động não, chia sẻ nhóm đôi, đặt câu hỏi',
        activities=[
            {
                'ten': '1. Hoạt động MỞ ĐẦU (5 phút)',
                'muc_tieu': 'Khởi động tư duy về các ứng dụng đa dạng của máy tính.',
                'gv': [
                    'GV đặt câu hỏi: "Trong 1 ngày, em đã dùng máy tính/điện thoại bao nhiêu lần và để làm gì?"',
                    'GV ghi nhanh ý kiến lên bảng, dẫn vào bài.',
                ],
                'hs': [
                    '- HS phát biểu tự do.',
                    '- HS so sánh câu trả lời của nhau.',
                ],
            },
            {
                'ten': '2. HOẠT ĐỘNG HÌNH THÀNH KIẾN THỨC MỚI',
                'sub': [
                    {
                        'ten': '2.1. Máy tính trong học tập và sáng tạo (10 phút)',
                        'muc_tieu': 'HS biết được các ứng dụng máy tính hỗ trợ học tập và sáng tạo.',
                        'gv': [
                            'GV trình chiếu: Soạn văn bản, vẽ tranh, làm slide, học tiếng Anh online, lập trình Scratch.',
                            'GV thao tác nhanh demo 1-2 ứng dụng.',
                            '* Kết luận: Máy tính là công cụ mạnh mẽ cho học tập và sáng tạo.',
                        ],
                        'hs': [
                            '- HS quan sát và đặt câu hỏi.',
                            '- HS chia sẻ ứng dụng em đã dùng.',
                        ],
                    },
                    {
                        'ten': '2.2. Sử dụng máy tính an toàn (5 phút)',
                        'muc_tieu': 'HS nhận biết và tuân thủ các quy tắc dùng máy tính an toàn.',
                        'gv': [
                            'GV nêu 5 quy tắc: Giới hạn thời gian, bảo vệ mật khẩu, không chia sẻ thông tin cá nhân, không xem nội dung không phù hợp, giữ tư thế đúng.',
                        ],
                        'hs': [
                            '- HS lắng nghe và ghi nhớ 5 quy tắc.',
                            '- HS cho ví dụ về vi phạm quy tắc.',
                        ],
                    },
                ],
            },
            {
                'ten': '3. HĐ LUYỆN TẬP-THỰC HÀNH',
                'sub': [
                    {
                        'ten': '3.1. Lập kế hoạch sử dụng máy tính (8 phút)',
                        'muc_tieu': 'HS tự lập được lịch sử dụng máy tính hợp lí.',
                        'gv': [
                            'GV phát mẫu lịch tuần, HS điền thời gian và mục đích dùng máy tính.',
                            '* GV nhận xét, tư vấn điều chỉnh cho hợp lí.',
                        ],
                        'hs': [
                            '- HS điền vào mẫu lịch.',
                            '- HS chia sẻ kế hoạch với bạn cạnh.',
                        ],
                    },
                    {
                        'ten': '3.2. Thực hành mở ứng dụng (2 phút)',
                        'muc_tieu': 'HS thực hành mở và đóng một ứng dụng trên máy tính.',
                        'gv': ['GV hướng dẫn HS mở Paint, vẽ một hình tự do, lưu file.'],
                        'hs': ['- HS thực hành trên máy tính.'],
                    },
                ],
            },
            {
                'ten': '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                'gv': [
                    '4. HOẠT ĐỘNG VẬN DỤNG, TRẢI NGHIỆM (5 phút)',
                    'GV: "Em hãy chọn 1 ứng dụng trên máy tính để giúp em hoàn thành một bài tập ở nhà. Chia sẻ kết quả tuần sau."',
                    'Dặn dò: Áp dụng kế hoạch sử dụng máy tính đã lập, chia sẻ với bố mẹ.',
                ],
                'hs': [
                    '- HS lên kế hoạch cụ thể.',
                    '- HS lắng nghe và ghi lại nhiệm vụ.',
                ],
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 6 — THCS — Bài 1: Thông tin và dữ liệu
    # ────────────────────────────────────────────────────────────────────
    print('\n🟠 Tạo KHBD Lớp 6 - Bài 1 (THCS)...')
    saved = create_khbd_thcs(
        grade_name='Lớp 6', grade_folder='Lớp_6',
        bai_so=1, ten_bai='Thông tin và dữ liệu',
        mon_hoc='Tin học', lop='6', thoi_luong='1 tiết (45 phút)', tiet_ppct='1',
        ngay_soan='   /   /2026', ngay_day='   /   /2026',
        kien_thuc_items=[
            '- Sự hiểu biết về khái niệm thông tin và dữ liệu trong máy tính.',
            '- Khả năng phân biệt thông tin (ý nghĩa con người hiểu) với dữ liệu (biểu diễn trong máy tính).',
            '- Sự nhận biết các dạng dữ liệu cơ bản: số, văn bản, hình ảnh, âm thanh, video.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Chủ động tìm kiếm ví dụ về thông tin và dữ liệu trong cuộc sống. (Đạt được thông qua Hoạt động 1, 2)',
            '- Giao tiếp và hợp tác: Thảo luận nhóm, chia sẻ ví dụ và đưa ra kết luận chung. (Đạt được thông qua Hoạt động 2, 3)',
        ],
        nang_luc_dac_thu_items=[
            '- NLa (Hiểu biết máy tính): Mô tả được khái niệm thông tin và dữ liệu; phân biệt được các dạng dữ liệu. (Đạt được thông qua Hoạt động 2)',
            '- NLc (Giải quyết vấn đề): Áp dụng hiểu biết về dữ liệu để phân tích tình huống thực tế. (Đạt được thông qua Hoạt động 3)',
        ],
        nang_luc_so_items=[
            '- Năng lực 1.1 (Sử dụng thiết bị kĩ thuật số): Hiểu cách máy tính biểu diễn thông tin dưới dạng dữ liệu. (Đạt được thông qua Hoạt động 2)',
        ],
        pham_chat_items=[
            '- Chăm chỉ: Tích cực tìm hiểu và ghi nhớ khái niệm về thông tin và dữ liệu. (Thông qua Hoạt động 1, 2)',
            '- Trung thực: Đánh giá đúng giá trị của thông tin, không bịa đặt dữ liệu. (Thông qua Hoạt động 3)',
            '- Trách nhiệm: Sử dụng thông tin có trách nhiệm, không chia sẻ thông tin sai. (Thông qua Hoạt động 4)',
        ],
        thiet_bi='Máy tính GV, máy chiếu, bảng tương tác, phiếu học tập nhóm.',
        hoc_lieu='SGK Tin học 6 (trang 6-12), phiếu bài tập phân loại dữ liệu.',
        hoat_dong_list=[
            {
                'stt': 1,
                'ten': 'Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)',
                'muc_tieu': 'Giúp học sinh xác định được mối quan hệ giữa thông tin và dữ liệu trong thực tế.',
                'noi_dung': 'GV chiếu tình huống: "Em nhận được dãy số 36.5 — đây là dữ liệu gì? Có ý nghĩa gì?" HS thảo luận cặp đôi.',
                'san_pham': 'HS phát biểu được rằng dãy số 36.5 có thể là nhiệt độ cơ thể (thông tin: không sốt) hoặc điểm số.',
                'to_chuc': 'GV chiếu câu hỏi tình huống, yêu cầu HS thảo luận cặp đôi trong 2 phút.',
                'buoc1': 'GV chiếu hình ảnh: dãy số "36.5" và hỏi "Đây là gì? Có nghĩa là gì?"',
                'buoc2': 'HS nhận câu hỏi, thảo luận cặp đôi, mỗi cặp nêu ít nhất 1 cách hiểu.',
                'buoc3': '2-3 cặp HS trình bày, HS khác nhận xét.',
                'buoc4': 'GV nhận xét, nhấn mạnh: cùng 1 dữ liệu nhưng có thể mang nhiều ý nghĩa khác nhau tùy ngữ cảnh → Dẫn vào bài.',
            },
            {
                'stt': 2,
                'ten': 'Hình thành kiến thức mới/giải quyết vấn đề',
                'muc_tieu': 'HS phân biệt được thông tin và dữ liệu; nhận biết 5 dạng dữ liệu cơ bản.',
                'noi_dung': 'GV giảng khái niệm thông tin và dữ liệu qua SGK trang 6-8; HS quan sát ví dụ và phân loại.',
                'san_pham': 'HS ghi được vào vở: định nghĩa thông tin, dữ liệu và 5 dạng dữ liệu với ví dụ minh họa.',
                'to_chuc': 'GV trình bày kiến thức kết hợp đặt câu hỏi tương tác.',
                'buoc1': 'GV chiếu slide khái niệm: Thông tin (information) = ý nghĩa mà con người hiểu. Dữ liệu (data) = biểu diễn của thông tin trong máy tính.',
                'buoc2': 'HS đọc SGK tr.6, gạch chân định nghĩa và ví dụ. GV hỏi: "Cho ví dụ 1 thông tin và dữ liệu tương ứng."',
                'buoc3': 'HS trả lời cá nhân, GV gọi 3-4 HS. GV giới thiệu 5 dạng: số, văn bản, hình ảnh, âm thanh, video.',
                'buoc4': 'GV tổng kết, HS ghi vào vở định nghĩa và bảng 5 dạng dữ liệu.',
            },
            {
                'stt': 3,
                'ten': 'Luyện tập',
                'muc_tieu': 'Củng cố khả năng phân biệt thông tin/dữ liệu và phân loại dạng dữ liệu.',
                'noi_dung': 'HS thực hành phân loại 10 ví dụ vào đúng dạng dữ liệu; giải thích lý do.',
                'san_pham': 'Phiếu học tập hoàn chỉnh với 10/10 ví dụ được phân loại đúng và có giải thích.',
                'to_chuc': 'Hoạt động cá nhân, sau đó đổi phiếu kiểm tra chéo.',
                'buoc1': 'GV phát phiếu học tập với 10 ví dụ (ảnh chụp, file MP3, bảng Excel...). HS phân loại vào đúng dạng dữ liệu.',
                'buoc2': 'HS làm việc cá nhân trong 5 phút, phân loại và viết giải thích ngắn.',
                'buoc3': 'HS đổi phiếu theo cặp, kiểm tra chéo. GV gọi 2 cặp báo cáo kết quả.',
                'buoc4': 'GV chiếu đáp án, giải thích các trường hợp khó. Tổng kết điểm cần nhớ.',
            },
            {
                'stt': 4,
                'ten': 'Mở rộng (Nhiệm vụ về nhà)',
                'muc_tieu': 'HS vận dụng kiến thức vào thực tế và chuẩn bị bài học tiếp theo.',
                'noi_dung': 'HS quan sát dữ liệu trong cuộc sống (điện thoại, máy tính ở nhà) và ghi lại 5 ví dụ về các dạng dữ liệu khác nhau.',
                'san_pham': 'Danh sách 5 ví dụ dữ liệu từ cuộc sống thực với chú thích dạng dữ liệu.',
                'buoc1': 'GV giao nhiệm vụ: Quan sát điện thoại/máy tính ở nhà, tìm ví dụ về 5 dạng dữ liệu khác nhau.',
                'buoc2': 'HS ghi nhiệm vụ vào sổ tay.',
                'buoc3': 'HS sẽ nộp kết quả đầu tiết học sau.',
                'buoc4': 'GV nhắc nhở: Ghi rõ tên dữ liệu, dạng dữ liệu và ý nghĩa (thông tin) của nó.',
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 7 — THCS — Bài 1: Thiết bị vào - ra
    # ────────────────────────────────────────────────────────────────────
    print('\n🟠 Tạo KHBD Lớp 7 - Bài 1 (THCS)...')
    saved = create_khbd_thcs(
        grade_name='Lớp 7', grade_folder='Lớp_7',
        bai_so=1, ten_bai='Thiết bị vào - ra',
        mon_hoc='Tin học', lop='7', thoi_luong='1 tiết (45 phút)', tiet_ppct='1',
        ngay_soan='   /   /2026', ngay_day='   /   /2026',
        kien_thuc_items=[
            '- Sự hiểu biết về khái niệm thiết bị vào (Input devices) và thiết bị ra (Output devices).',
            '- Khả năng phân loại và liệt kê các thiết bị vào/ra phổ biến của máy tính.',
            '- Sự hiểu biết về chức năng cơ bản của từng loại thiết bị vào-ra.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Tự tìm hiểu thêm về các thiết bị vào-ra mới trên thị trường. (Đạt được thông qua HĐ 2, 4)',
            '- Giao tiếp và hợp tác: Thảo luận nhóm phân loại thiết bị và trình bày kết quả. (Đạt được thông qua HĐ 2, 3)',
        ],
        nang_luc_dac_thu_items=[
            '- NLa (Hiểu biết máy tính): Phân loại chính xác các thiết bị vào-ra của máy tính. (Đạt được thông qua HĐ 2)',
            '- NLb (Kết nối và cộng tác): Mô tả chức năng của từng thiết bị trong hệ thống máy tính. (Đạt được thông qua HĐ 3)',
        ],
        nang_luc_so_items=[
            '- Năng lực 1.2 (Hiểu biết về thiết bị): Nhận biết và sử dụng đúng cách các thiết bị vào-ra phổ biến. (Đạt được thông qua HĐ 3)',
        ],
        pham_chat_items=[
            '- Chăm chỉ: Tích cực tham gia thảo luận nhóm và hoàn thành phiếu học tập. (Thông qua HĐ 2, 3)',
            '- Trách nhiệm: Sử dụng và bảo quản thiết bị vào-ra đúng cách. (Thông qua HĐ 4)',
        ],
        thiet_bi='Máy chiếu, máy tính, bàn phím, chuột, loa ngoài, tai nghe, webcam (nếu có).',
        hoc_lieu='SGK Tin học 7 (trang 6-10), phiếu học tập phân loại thiết bị.',
        hoat_dong_list=[
            {
                'stt': 1,
                'ten': 'Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)',
                'muc_tieu': 'Kích hoạt hiểu biết về các thiết bị máy tính và tạo nhu cầu học.',
                'noi_dung': 'GV đặt câu hỏi: "Kể tên tất cả thiết bị em thấy xung quanh máy tính trong phòng học này." HS quan sát và kể.',
                'san_pham': 'Danh sách 5-10 thiết bị xung quanh máy tính được HS liệt kê.',
                'to_chuc': 'Hoạt động cá nhân nhanh (2 phút) sau đó chia sẻ với lớp.',
                'buoc1': 'GV yêu cầu HS quan sát phòng máy và liệt kê nhanh các thiết bị nhìn thấy.',
                'buoc2': 'HS quan sát xung quanh và ghi nhanh vào giấy nháp.',
                'buoc3': 'HS xung phong đọc danh sách. GV ghi lên bảng.',
                'buoc4': 'GV đặt câu hỏi: "Chúng ta sẽ phân loại những thiết bị này như thế nào?" → Dẫn vào bài.',
            },
            {
                'stt': 2,
                'ten': 'Hình thành kiến thức mới/giải quyết vấn đề',
                'muc_tieu': 'HS phân biệt và định nghĩa được thiết bị vào và thiết bị ra.',
                'noi_dung': 'GV giảng khái niệm: thiết bị vào (đưa thông tin vào máy), thiết bị ra (đưa thông tin ra); HS làm việc nhóm phân loại.',
                'san_pham': 'HS ghi được định nghĩa và ví dụ 5 thiết bị vào, 5 thiết bị ra vào vở.',
                'to_chuc': 'GV giảng kết hợp hỏi đáp; sau đó nhóm 4 làm bài tập phân loại.',
                'buoc1': 'GV chiếu slide: Thiết bị vào = đưa thông tin/lệnh vào máy (bàn phím, chuột, micro, camera). Thiết bị ra = đưa thông tin ra ngoài (màn hình, máy in, loa).',
                'buoc2': 'HS đọc SGK tr.6-8, gạch chân ví dụ. GV hỏi từng thiết bị: "Thiết bị này vào hay ra?"',
                'buoc3': 'Nhóm 4 người thảo luận phân loại danh sách 10 thiết bị trên phiếu học tập. Đại diện trình bày.',
                'buoc4': 'GV nhận xét, giải thích các trường hợp đặc biệt (màn hình cảm ứng = vừa vào vừa ra). HS ghi vở.',
            },
            {
                'stt': 3,
                'ten': 'Luyện tập',
                'muc_tieu': 'Củng cố kĩ năng phân loại thiết bị vào-ra; liên hệ thực tế.',
                'noi_dung': 'HS thực hành phân loại thiết bị và kết hợp thực hành sử dụng bàn phím, chuột.',
                'san_pham': 'Phiếu bài tập hoàn chỉnh; HS gõ đúng đoạn văn mẫu bằng bàn phím.',
                'to_chuc': 'Kết hợp bài tập lý thuyết và thực hành máy tính.',
                'buoc1': 'GV phát phiếu: 10 câu hỏi phân loại thiết bị + 2 câu tình huống. HS mở máy, thực hành gõ đoạn văn.',
                'buoc2': 'HS làm bài tập lý thuyết (5 phút) và thực hành gõ bàn phím (5 phút).',
                'buoc3': 'HS đổi phiếu kiểm tra chéo. GV chiếu đáp án. HS báo cáo kết quả gõ bàn phím.',
                'buoc4': 'GV nhận xét kết quả thực hành. Nhấn mạnh tư thế ngồi và cách đặt tay đúng.',
            },
            {
                'stt': 4,
                'ten': 'Mở rộng (Nhiệm vụ về nhà)',
                'muc_tieu': 'HS tìm hiểu thêm về các thiết bị vào-ra hiện đại.',
                'noi_dung': 'HS tìm hiểu 2 thiết bị vào-ra hiện đại (VR headset, bút stylus, máy in 3D...) và mô tả chức năng.',
                'san_pham': 'Đoạn văn ngắn (5-7 câu) mô tả 2 thiết bị vào-ra hiện đại em tìm được.',
                'buoc1': 'GV giao nhiệm vụ: Tìm hiểu 2 thiết bị vào-ra hiện đại chưa học, viết mô tả ngắn.',
                'buoc2': 'HS ghi nhiệm vụ vào sổ tay.',
                'buoc3': 'Nộp bài đầu tiết học sau, sẽ được chia sẻ với lớp.',
                'buoc4': 'GV gợi ý: Tìm kiếm trên Internet với từ khóa "thiết bị đầu vào máy tính mới nhất".',
            },
        ]
    )
    results.append(saved)

    # ────────────────────────────────────────────────────────────────────
    # LỚP 8 — THCS — Bài 1: Lược sử công cụ tính toán
    # ────────────────────────────────────────────────────────────────────
    print('\n🟠 Tạo KHBD Lớp 8 - Bài 1 (THCS)...')
    saved = create_khbd_thcs(
        grade_name='Lớp 8', grade_folder='Lớp_8',
        bai_so=1, ten_bai='Lược sử công cụ tính toán',
        mon_hoc='Tin học', lop='8', thoi_luong='1 tiết (45 phút)', tiet_ppct='1',
        ngay_soan='   /   /2026', ngay_day='   /   /2026',
        kien_thuc_items=[
            '- Sự hiểu biết về các mốc lịch sử phát triển công cụ tính toán từ thủ công đến hiện đại.',
            '- Khả năng mô tả đặc điểm và vai trò của từng thế hệ máy tính (Thế hệ 1-5).',
            '- Sự nhận biết xu hướng phát triển công nghệ tính toán hiện nay và tương lai.',
        ],
        nang_luc_chung_items=[
            '- Tự chủ và tự học: Chủ động tìm kiếm và hệ thống hóa kiến thức về lịch sử máy tính. (Đạt được thông qua HĐ 2, 4)',
            '- Giao tiếp và hợp tác: Thảo luận và thuyết trình về lịch sử phát triển công nghệ. (Đạt được thông qua HĐ 2, 3)',
        ],
        nang_luc_dac_thu_items=[
            '- NLa (Hiểu biết máy tính): Trình bày được lịch sử phát triển máy tính theo đúng thứ tự thời gian. (Đạt được thông qua HĐ 2)',
            '- NLe (Ứng xử phù hợp): Nhận thức được tác động của công nghệ máy tính đến cuộc sống xã hội. (Đạt được thông qua HĐ 4)',
        ],
        nang_luc_so_items=[
            '- Năng lực 5.1 (Hiểu về lịch sử công nghệ): Mô tả được sự phát triển của công nghệ số qua các thế hệ. (Đạt được thông qua HĐ 2, 3)',
        ],
        pham_chat_items=[
            '- Chăm chỉ: Tích cực nghiên cứu và ghi chép thông tin lịch sử máy tính. (Thông qua HĐ 2)',
            '- Trách nhiệm: Trân trọng thành tựu khoa học và có ý thức ứng dụng công nghệ đúng đắn. (Thông qua HĐ 4)',
        ],
        thiet_bi='Máy chiếu, máy tính, video lược sử máy tính (2-3 phút), tranh ảnh minh họa các thế hệ máy tính.',
        hoc_lieu='SGK Tin học 8 (trang 6-14), phiếu học tập dòng thời gian.',
        hoat_dong_list=[
            {
                'stt': 1,
                'ten': 'Khởi động (Xác định vấn đề/nhiệm vụ học tập/Mở đầu)',
                'muc_tieu': 'Tạo hứng thú và kích hoạt tư duy về lịch sử phát triển công nghệ tính toán.',
                'noi_dung': 'GV chiếu 2 hình ảnh: chiếc bàn tính abacus cổ đại và siêu máy tính hiện đại. Hỏi: "Có điểm gì giống nhau giữa 2 thiết bị này?"',
                'san_pham': 'HS nêu được ít nhất 1 điểm chung (đều dùng để tính toán) và nêu được sự khác biệt về công nghệ.',
                'to_chuc': 'Hoạt động quan sát và thảo luận cặp đôi (2 phút).',
                'buoc1': 'GV chiếu 2 hình ảnh (abacus và siêu máy tính) và đặt câu hỏi tư duy.',
                'buoc2': 'HS quan sát, thảo luận cặp đôi tìm điểm giống và khác nhau.',
                'buoc3': '3 cặp HS trình bày. GV ghi ý kiến lên bảng.',
                'buoc4': 'GV dẫn vào bài: "Hành trình từ abacus đến siêu máy tính là hành trình hàng ngàn năm. Hôm nay chúng ta cùng khám phá lược sử đó."',
            },
            {
                'stt': 2,
                'ten': 'Hình thành kiến thức mới/giải quyết vấn đề',
                'muc_tieu': 'HS nắm được các mốc chính trong lịch sử phát triển công cụ tính toán từ cổ đại đến hiện đại.',
                'noi_dung': 'GV trình bày lịch sử qua timeline: Abacus → Pascaline → ENIAC → Microprocessor → AI/Cloud. HS điền vào phiếu dòng thời gian.',
                'san_pham': 'Phiếu dòng thời gian hoàn chỉnh với 5 mốc lịch sử, đặc điểm và năm ra đời.',
                'to_chuc': 'GV trình bày timeline kết hợp video ngắn; HS điền phiếu đồng thời.',
                'buoc1': 'GV chiếu timeline và video lược sử máy tính (2 phút). Phát phiếu dòng thời gian cho HS.',
                'buoc2': 'HS theo dõi video, đọc SGK tr.6-10, điền vào phiếu dòng thời gian: tên, năm, đặc điểm.',
                'buoc3': 'HS so sánh phiếu với bạn cạnh, bổ sung thiếu sót. GV gọi 1 HS đọc timeline của mình.',
                'buoc4': 'GV chốt đáp án, giải thích thêm về ENIAC và chiếc máy tính đầu tiên dùng transistor. HS sửa phiếu.',
            },
            {
                'stt': 3,
                'ten': 'Luyện tập',
                'muc_tieu': 'Củng cố kiến thức về lịch sử máy tính và liên hệ thực tế.',
                'noi_dung': 'HS thực hành bài tập sắp xếp thứ tự thời gian và câu hỏi liên hệ về công nghệ hiện đại.',
                'san_pham': 'Phiếu bài tập hoàn chỉnh và đoạn văn ngắn liên hệ thực tế.',
                'to_chuc': 'Bài tập cá nhân sau đó thảo luận nhóm.',
                'buoc1': 'GV phát phiếu bài tập: Sắp xếp 10 sự kiện theo đúng thứ tự lịch sử; 2 câu liên hệ.',
                'buoc2': 'HS làm bài cá nhân (7 phút). GV quan sát, gợi ý HS gặp khó khăn.',
                'buoc3': 'HS đổi bài kiểm tra chéo. Nhóm 4 thảo luận câu hỏi liên hệ: "Theo em, 20 năm nữa máy tính sẽ như thế nào?"',
                'buoc4': 'Đại diện 2 nhóm trình bày dự đoán. GV nhận xét, giới thiệu xu hướng AI và Quantum Computing.',
            },
            {
                'stt': 4,
                'ten': 'Mở rộng (Nhiệm vụ về nhà)',
                'muc_tieu': 'HS tìm hiểu sâu hơn về một nhân vật lịch sử hoặc mốc phát triển máy tính mà em quan tâm.',
                'noi_dung': 'HS chọn 1 trong số: Alan Turing, Grace Hopper, Steve Jobs, hay sự ra đời của Internet; viết đoạn thuyết trình 7-10 câu.',
                'san_pham': 'Đoạn văn 7-10 câu về nhân vật/mốc lịch sử được chọn, nộp đầu tiết học sau.',
                'buoc1': 'GV giới thiệu 4 chủ đề lựa chọn và giao nhiệm vụ.',
                'buoc2': 'HS ghi nhiệm vụ vào sổ tay, chọn chủ đề yêu thích.',
                'buoc3': 'Tuần sau HS nộp bài, GV chọn 3-4 bài hay đọc trước lớp.',
                'buoc4': 'GV gợi ý nguồn tham khảo: Britannica, Wikipedia tiếng Việt, kênh YouTube khoa học.',
            },
        ]
    )
    results.append(saved)

    return results


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    print('🚀 Bắt đầu tạo toàn bộ KHBD Bài 1 (Tiền TH → Lớp 8)...\n')
    results = build_all()
    print(f'\n✅ HOÀN TẤT! Đã tạo {len(results)} file KHBD.')
    for r in results:
        print(f'  📄 {r}')
