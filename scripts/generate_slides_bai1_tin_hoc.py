"""
Generate Slide Bài 1 Tin học for all grades (Tiền TH, Lớp 1-8).
Reads KHBD .docx to extract lesson content, then creates .pptx slides.
"""
import sys, os, glob, copy, re
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from docx import Document
from lxml import etree

# ── Constants ──────────────────────────────────────────────────────
TEMPLATE = r'D:\UNIGO\Hệ thống mẫu văn bản\Mẫu slide có chân trang.pptx'
KHBD_ROOT = r'D:\UNIGO\KHBD_Tin_học'
SGK_ROOT = r'D:\UNIGO\SGK'

# Slide dimensions (matching existing Tiết 00 slides: 13.33 x 7.5 inches)
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# Safe zone
SAFE_TOP = Inches(1.15)
SAFE_BOTTOM = Inches(6.30)
SAFE_HEIGHT = Inches(5.15)
CONTENT_TOP = Inches(1.30)
CONTENT_LEFT = Inches(0.50)
CONTENT_RIGHT_MARGIN = Inches(0.50)

# ── 9 Color Palettes ──────────────────────────────────────────────
PALETTES = {
    'TTH': {  # Tiền tiểu học - Hồng cam vui nhộn
        'primary': RGBColor(0xFF, 0x6B, 0x6B),
        'secondary': RGBColor(0xFF, 0x8E, 0x53),
        'accent': RGBColor(0xFF, 0xD9, 0x3D),
        'bg': RGBColor(0xFF, 0xF5, 0xF5),
        'text_dark': RGBColor(0x2D, 0x13, 0x3A),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '1': {  # Lớp 1 - Xanh lá tươi
        'primary': RGBColor(0x4E, 0xCA, 0x54),
        'secondary': RGBColor(0x2E, 0x7D, 0x32),
        'accent': RGBColor(0xFF, 0xEB, 0x3B),
        'bg': RGBColor(0xF1, 0xF8, 0xE9),
        'text_dark': RGBColor(0x1B, 0x2D, 0x1B),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '2': {  # Lớp 2 - Tím hoa oải hương
        'primary': RGBColor(0x7C, 0x4D, 0xFF),
        'secondary': RGBColor(0x53, 0x6D, 0xFE),
        'accent': RGBColor(0xFF, 0xC1, 0x07),
        'bg': RGBColor(0xF3, 0xF0, 0xFF),
        'text_dark': RGBColor(0x1A, 0x12, 0x3E),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '3': {  # Lớp 3 - Xanh dương tươi
        'primary': RGBColor(0x21, 0x96, 0xF3),
        'secondary': RGBColor(0x15, 0x65, 0xC0),
        'accent': RGBColor(0xFF, 0xAB, 0x00),
        'bg': RGBColor(0xE3, 0xF2, 0xFD),
        'text_dark': RGBColor(0x0D, 0x1F, 0x3C),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '4': {  # Lớp 4 - Cam ấm
        'primary': RGBColor(0xFF, 0x98, 0x00),
        'secondary': RGBColor(0xEF, 0x6C, 0x00),
        'accent': RGBColor(0x00, 0xBC, 0xD4),
        'bg': RGBColor(0xFF, 0xF3, 0xE0),
        'text_dark': RGBColor(0x33, 0x1A, 0x00),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '5': {  # Lớp 5 - Hồng tím
        'primary': RGBColor(0xE9, 0x1E, 0x63),
        'secondary': RGBColor(0x88, 0x0E, 0x4F),
        'accent': RGBColor(0x00, 0xE5, 0xFF),
        'bg': RGBColor(0xFC, 0xE4, 0xEC),
        'text_dark': RGBColor(0x3E, 0x0A, 0x22),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '6': {  # Lớp 6 - Xanh teal
        'primary': RGBColor(0x00, 0x96, 0x88),
        'secondary': RGBColor(0x00, 0x4D, 0x40),
        'accent': RGBColor(0xFF, 0xD5, 0x4F),
        'bg': RGBColor(0xE0, 0xF2, 0xF1),
        'text_dark': RGBColor(0x0A, 0x2B, 0x28),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '7': {  # Lớp 7 - Xanh đậm (Navy)
        'primary': RGBColor(0x30, 0x3F, 0x9F),
        'secondary': RGBColor(0x1A, 0x23, 0x7E),
        'accent': RGBColor(0xFF, 0xC1, 0x07),
        'bg': RGBColor(0xE8, 0xEA, 0xF6),
        'text_dark': RGBColor(0x0F, 0x14, 0x3A),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
    '8': {  # Lớp 8 - Xanh ngọc bích
        'primary': RGBColor(0x00, 0x79, 0x7B),
        'secondary': RGBColor(0x00, 0x50, 0x52),
        'accent': RGBColor(0xFF, 0xB7, 0x4D),
        'bg': RGBColor(0xE0, 0xF7, 0xFA),
        'text_dark': RGBColor(0x00, 0x2B, 0x2C),
        'text_light': RGBColor(0xFF, 0xFF, 0xFF),
        'card_bg': RGBColor(0xFF, 0xFF, 0xFF),
    },
}

# ── Grade Config ──────────────────────────────────────────────────
GRADES = [
    {
        'key': 'TTH', 'folder': 'Tiền_tiểu_học', 'label': 'Tiền TH',
        'display': 'TIỀN TIỂU HỌC', 'sgk': False,
    },
    {
        'key': '1', 'folder': 'Lớp_1', 'label': 'Lớp 1',
        'display': 'LỚP 1', 'sgk': False,
    },
    {
        'key': '2', 'folder': 'Lớp_2', 'label': 'Lớp 2',
        'display': 'LỚP 2', 'sgk': False,
    },
    {
        'key': '3', 'folder': 'Lớp_3', 'label': 'Lớp 3',
        'display': 'LỚP 3', 'sgk': True, 'sgk_grade': 3,
    },
    {
        'key': '4', 'folder': 'Lớp_4', 'label': 'Lớp 4',
        'display': 'LỚP 4', 'sgk': True, 'sgk_grade': 4,
    },
    {
        'key': '5', 'folder': 'Lớp_5', 'label': 'Lớp 5',
        'display': 'LỚP 5', 'sgk': True, 'sgk_grade': 5,
    },
    {
        'key': '6', 'folder': 'Lớp_6', 'label': 'Lớp 6',
        'display': 'LỚP 6', 'sgk': True, 'sgk_grade': 6,
    },
    {
        'key': '7', 'folder': 'Lớp_7', 'label': 'Lớp 7',
        'display': 'LỚP 7', 'sgk': True, 'sgk_grade': 7,
    },
    {
        'key': '8', 'folder': 'Lớp_8', 'label': 'Lớp 8',
        'display': 'LỚP 8', 'sgk': True, 'sgk_grade': 8,
    },
]


# ══════════════════════════════════════════════════════════════════
#                     KHBD READER
# ══════════════════════════════════════════════════════════════════

def read_khbd(grade_cfg):
    """Read KHBD .docx and extract lesson data."""
    folder = grade_cfg['folder']
    files = glob.glob(os.path.join(KHBD_ROOT, folder, 'Bài_01', '*.docx'))
    if not files:
        raise FileNotFoundError(f"No KHBD found for {folder}")
    
    doc = Document(files[0])
    
    data = {
        'title': '',          # e.g. "BÀI 1. THÔNG TIN VÀ QUYẾT ĐỊNH"
        'topic': '',          # e.g. "Thông tin quanh ta"
        'objectives': [],     # Yêu cầu cần đạt
        'activities': [],     # List of activity dicts
    }
    
    # Extract title and topic from paragraphs
    is_thcs = grade_cfg['key'] in ('6', '7', '8')
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        if 'BÀI:' in txt:
            raw = txt.replace('BÀI:', '').strip()
            # Remove "(Tiết: X)" suffix
            raw = re.sub(r'\s*\(Tiết.*?\)', '', raw).strip()
            data['title'] = raw
        elif 'CHỦ ĐIỂM:' in txt or 'CHỦ ĐỀ:' in txt:
            data['topic'] = txt.split(':', 1)[1].strip()
        elif txt.startswith('TÊN BÀI DẠY:'):
            data['title'] = 'Bài 1. ' + txt.replace('TÊN BÀI DẠY:', '').strip()
        elif txt.startswith('Tên tiết:') and not data['topic']:
            data['topic'] = txt.replace('Tên tiết:', '').strip()
    
    # Extract objectives
    in_objectives = False
    for p in doc.paragraphs:
        txt = p.text.strip()
        if 'YÊU CẦU CẦN ĐẠT' in txt or 'Mục tiêu' == txt.strip():
            in_objectives = True
            continue
        if in_objectives:
            if txt.startswith('II.') or txt.startswith('III.') or txt.startswith('1. Kiến thức'):
                if txt.startswith('1. Kiến thức'):
                    in_objectives = True  # continue for THCS
                    continue
                break
            if txt.startswith('- ') and 'NL' not in txt and 'Năng lực' not in txt:
                obj = txt.lstrip('- ').strip()
                if len(obj) > 10:
                    data['objectives'].append(obj)
    
    # For THCS, extract objectives from "Kiến thức" section
    if is_thcs and not data['objectives']:
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt.startswith('- Sự ') or txt.startswith('- Khả năng') or txt.startswith('- Sự nhận biết'):
                data['objectives'].append(txt.lstrip('- '))
    
    # Extract activities from paragraphs (THCS format)
    if is_thcs:
        activities = _extract_thcs_activities(doc)
        data['activities'] = activities
    else:
        # TH format: single table with GV/HS columns
        activities = _extract_th_activities(doc)
        data['activities'] = activities
    
    return data


def _extract_th_activities(doc):
    """Extract activities from TH-format KHBD (single table with GV|HS)."""
    activities = []
    if not doc.tables:
        return activities
    
    table = doc.tables[0]
    current_act = None
    
    for row in table.rows:
        gv_text = row.cells[0].text.strip()
        hs_text = row.cells[1].text.strip() if len(row.cells) > 1 else ''
        
        # Extract multi-line content after the header line
        gv_lines = gv_text.split('\n')
        gv_extra = '\n'.join(gv_lines[1:]).strip() if len(gv_lines) > 1 else ''
        
        # Detect activity headers
        if 'MỞ ĐẦU' in gv_text or 'Hoạt động MỞ ĐẦU' in gv_text:
            current_act = {'type': 'khởi_động', 'title': 'Khởi động', 'gv': '', 'hs': ''}
            activities.append(current_act)
            # Header rows often contain content too
            if gv_extra:
                current_act['gv'] += gv_extra + '\n'
                current_act['hs'] += hs_text + '\n'
        elif 'KIẾN THỨC MỚI' in gv_text and not re.match(r'^2\.\d', gv_text):
            # Section header only, skip
            continue
        elif re.match(r'^2\.\d\.\s', gv_text):
            # Sub-activity of knowledge section (title + optional mục tiêu)
            match = re.match(r'^2\.\d\.\s*(.+?)(?:\s*\(\d+ phút\))?$', gv_lines[0])
            title = match.group(1) if match else gv_lines[0][:50]
            current_act = {'type': 'kiến_thức', 'title': title, 'gv': '', 'hs': ''}
            activities.append(current_act)
        elif 'LUYỆN TẬP' in gv_text and not re.match(r'^3\.\d', gv_text):
            continue
        elif re.match(r'^3\.\d\.\s', gv_text):
            match = re.match(r'^3\.\d\.\s*(.+?)(?:\s*\(\d+ phút\))?$', gv_lines[0])
            title = match.group(1) if match else gv_lines[0][:50]
            current_act = {'type': 'luyện_tập', 'title': title, 'gv': '', 'hs': ''}
            activities.append(current_act)
        elif 'VẬN DỤNG' in gv_text or 'TRẢI NGHIỆM' in gv_text:
            current_act = {'type': 'vận_dụng', 'title': 'Vận dụng', 'gv': '', 'hs': ''}
            activities.append(current_act)
            if gv_extra:
                current_act['gv'] += gv_extra + '\n'
                current_act['hs'] += hs_text + '\n'
        elif current_act is not None:
            # Content row - add to current activity
            current_act['gv'] += gv_text + '\n'
            current_act['hs'] += hs_text + '\n'
    
    return activities


def _extract_thcs_activities(doc):
    """Extract activities from THCS-format KHBD (paragraphs + multiple tables)."""
    activities = []
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        
        # Detect activity headers with various naming patterns
        if re.search(r'Hoạt động\s*1', txt) and any(kw in txt for kw in ['Khởi động', 'Mở đầu', 'Xác định']):
            act = {'type': 'khởi_động', 'title': 'Khởi động', 'content': '', 'product': ''}
            activities.append(act)
        elif re.search(r'Hoạt động\s*2', txt) and any(kw in txt.lower() for kw in ['hình thành', 'kiến thức', 'giải quyết']):
            act = {'type': 'kiến_thức', 'title': 'Hình thành kiến thức mới', 'content': '', 'product': ''}
            activities.append(act)
        elif re.search(r'Hoạt động\s*3', txt) and 'Luyện tập' in txt:
            act = {'type': 'luyện_tập', 'title': 'Luyện tập', 'content': '', 'product': ''}
            activities.append(act)
        elif re.search(r'Hoạt động\s*4', txt) and any(kw in txt for kw in ['Vận dụng', 'vận dụng', 'Mở rộng', 'mở rộng']):
            act = {'type': 'vận_dụng', 'title': 'Vận dụng – Mở rộng', 'content': '', 'product': ''}
            activities.append(act)
        elif activities and txt.startswith('a) Mục tiêu:'):
            activities[-1]['objective'] = txt.replace('a) Mục tiêu:', '').strip()
        elif activities and txt.startswith('b) Nội dung:'):
            activities[-1]['content'] = txt.replace('b) Nội dung:', '').strip()
        elif activities and txt.startswith('c) Sản phẩm:'):
            activities[-1]['product'] = txt.replace('c) Sản phẩm:', '').strip()
        elif activities and txt.startswith('d) Tổ chức thực hiện:'):
            activities[-1]['organize'] = txt.replace('d) Tổ chức thực hiện:', '').strip()
    
    # Also extract GV content from tables (tables 1-4 correspond to activities)
    act_tables = [t for t in doc.tables if len(t.columns) == 3 and len(t.rows) >= 4]
    for i, (act, tbl) in enumerate(zip(activities, act_tables)):
        gv_content = []
        for row in tbl.rows[1:]:  # Skip header
            gv_text = row.cells[1].text.strip() if len(row.cells) > 1 else ''
            if gv_text and len(gv_text) > 10:
                # Take first meaningful line
                for line in gv_text.split('\n'):
                    line = line.strip()
                    if line and len(line) > 15 and not line.startswith('HS '):
                        gv_content.append(line)
                        break
        if gv_content:
            act['gv_bullets'] = gv_content
    
    return activities


# ══════════════════════════════════════════════════════════════════
#                     SLIDE BUILDER
# ══════════════════════════════════════════════════════════════════

def create_presentation(grade_cfg, lesson_data):
    """Create a complete slide deck for a grade's Bài 1."""
    pal = PALETTES[grade_cfg['key']]
    
    # Create new presentation with correct dimensions
    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.33 inches
    prs.slide_height = Emu(6858000)   # 7.5 inches
    
    # We'll use blank layout
    blank_layout = prs.slide_layouts[6]  # Blank
    
    title_short = lesson_data['title']
    # Remove "BÀI 1. " prefix for display
    display_title = re.sub(r'^BÀI\s*\d+\.\s*', '', title_short, flags=re.IGNORECASE).strip()
    
    slides_created = []
    
    # ── Slide 1: TRANG BÌA ────────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    _add_cover_slide(slide, prs, grade_cfg, display_title, lesson_data['topic'], pal)
    slides_created.append('Trang bìa')
    
    # ── Slide 2: MỤC TIÊU ────────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    _add_objectives_slide(slide, prs, grade_cfg, lesson_data, pal)
    slides_created.append('Mục tiêu')
    
    # ── Slide 3: KHỞI ĐỘNG ────────────────────────────────────────
    kd_acts = [a for a in lesson_data['activities'] if a['type'] == 'khởi_động']
    if kd_acts:
        slide = prs.slides.add_slide(blank_layout)
        _add_activity_slide(slide, prs, '🎮  Khởi động', kd_acts[0], pal, 'left')
        slides_created.append('Khởi động')
    
    # ── Slide 4-7: KIẾN THỨC MỚI ─────────────────────────────────
    kt_acts = [a for a in lesson_data['activities'] if a['type'] == 'kiến_thức']
    for i, act in enumerate(kt_acts):
        slide = prs.slides.add_slide(blank_layout)
        side = 'right' if i % 2 == 0 else 'left'
        emoji = ['📖', '📝', '💡', '🔍'][i % 4]
        title_text = f'{emoji}  {act.get("title", "Kiến thức mới")}'
        _add_activity_slide(slide, prs, title_text, act, pal, side)
        slides_created.append(f'Kiến thức {i+1}')
    
    # If no knowledge activities found, add a generic one
    if not kt_acts:
        slide = prs.slides.add_slide(blank_layout)
        _add_generic_knowledge_slide(slide, prs, display_title, pal)
        slides_created.append('Kiến thức mới')
    
    # ── Slide 8-9: LUYỆN TẬP ─────────────────────────────────────
    lt_acts = [a for a in lesson_data['activities'] if a['type'] == 'luyện_tập']
    for i, act in enumerate(lt_acts):
        slide = prs.slides.add_slide(blank_layout)
        emoji = ['✏️', '🎯'][i % 2]
        title_text = f'{emoji}  {act.get("title", "Luyện tập")}'
        _add_activity_slide(slide, prs, title_text, act, pal, 'right' if i % 2 == 0 else 'left')
        slides_created.append(f'Luyện tập {i+1}')
    
    # ── Slide 10: VẬN DỤNG ───────────────────────────────────────
    vd_acts = [a for a in lesson_data['activities'] if a['type'] == 'vận_dụng']
    if vd_acts:
        slide = prs.slides.add_slide(blank_layout)
        _add_activity_slide(slide, prs, '🚀  Vận dụng', vd_acts[0], pal, 'left')
        slides_created.append('Vận dụng')
    
    # ── Slide 11: TỔNG KẾT ───────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    _add_summary_slide(slide, prs, display_title, lesson_data, pal)
    slides_created.append('Tổng kết')
    
    # ── Slide 12: CẢM ƠN ─────────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    _add_thankyou_slide(slide, prs, grade_cfg, pal)
    slides_created.append('Cảm ơn')
    
    # ── Add footer to ALL slides ──────────────────────────────────
    for slide in prs.slides:
        _add_footer(slide, prs)
    
    return prs, slides_created


def _set_slide_bg(slide, color):
    """Set slide background color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_shape(slide, shape_type, left, top, width, height, fill_color=None, line_color=None):
    """Add a shape with optional fill and border."""
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()  # No fill
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def _add_textbox(slide, left, top, width, height, text, font_size=Pt(18),
                 font_color=RGBColor(0, 0, 0), bold=False, alignment=PP_ALIGN.LEFT,
                 font_name='Times New Roman'):
    """Add a text box with formatted text."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def _add_bullet_textbox(slide, left, top, width, height, items, pal,
                        font_size=Pt(18), bullet_char='●', line_spacing=Pt(28)):
    """Add a text box with bullet points."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        # Add bullet
        run_bullet = p.add_run()
        run_bullet.text = f'{bullet_char}  '
        run_bullet.font.size = Pt(14)
        run_bullet.font.color.rgb = pal['primary']
        run_bullet.font.name = 'Times New Roman'
        
        # Add text
        run_text = p.add_run()
        run_text.text = item
        run_text.font.size = font_size
        run_text.font.color.rgb = pal['text_dark']
        run_text.font.name = 'Times New Roman'
        
        p.space_after = Pt(8)
        p.line_spacing = line_spacing
    
    return txBox


def _add_card_with_accent(slide, left, top, width, height, pal, accent_side='left'):
    """Add a white card with colored accent bar."""
    # White card
    card = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
                      fill_color=pal['card_bg'])
    # Make corners less rounded
    card.adjustments[0] = 0.02
    
    # Add shadow effect via XML
    spPr = card._element.spPr
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    effectLst = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst')
    outerShdw = etree.SubElement(effectLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}outerShdw',
                                  attrib={'blurRad': '50800', 'dist': '25400', 'dir': '5400000',
                                          'rotWithShape': '0'})
    srgbClr = etree.SubElement(outerShdw, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr',
                                attrib={'val': '000000'})
    etree.SubElement(srgbClr, '{http://schemas.openxmlformats.org/drawingml/2006/main}alpha',
                     attrib={'val': '25000'})
    
    # Accent bar
    bar_w = Inches(0.12)
    if accent_side == 'left':
        bar_left = left
    else:
        bar_left = left + width - bar_w
    
    bar = _add_shape(slide, MSO_SHAPE.RECTANGLE, bar_left, top, bar_w, height,
                     fill_color=pal['primary'])
    
    return card, bar


def _add_cover_slide(slide, prs, grade_cfg, display_title, topic, pal):
    """Create cover slide."""
    _set_slide_bg(slide, pal['primary'])
    
    sw = prs.slide_width
    
    # Large white card in center
    card_w = Inches(10)
    card_h = Inches(4.6)
    card_l = (sw - card_w) // 2
    card_t = Inches(1.40)
    
    card = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, card_l, card_t, card_w, card_h,
                      fill_color=RGBColor(0xFF, 0xFF, 0xFF))
    card.adjustments[0] = 0.03
    
    # Grade badge
    badge_w = Inches(4.5)
    badge_h = Inches(0.50)
    badge_l = card_l + Inches(0.5)
    badge_t = card_t + Inches(0.3)
    
    badge = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, badge_l, badge_t, badge_w, badge_h,
                       fill_color=pal['primary'])
    badge.adjustments[0] = 0.3
    badge_tf = badge.text_frame
    badge_tf.word_wrap = True
    badge_p = badge_tf.paragraphs[0]
    badge_p.alignment = PP_ALIGN.CENTER
    run = badge_p.add_run()
    run.text = f'TIN HỌC {grade_cfg["display"]}  •  NĂM HỌC 2026 – 2027'
    run.font.size = Pt(16)
    run.font.color.rgb = pal['text_light']
    run.font.bold = True
    run.font.name = 'Times New Roman'
    
    # Title
    title_l = card_l + Inches(0.5)
    title_t = badge_t + Inches(0.8)
    title_w = card_w - Inches(1)
    title_h = Inches(2.2)
    
    txBox = slide.shapes.add_textbox(title_l, title_t, title_w, title_h)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    # Line 1: "Bài 1 —"
    p1 = tf.paragraphs[0]
    r1 = p1.add_run()
    r1.text = 'Bài 1'
    r1.font.size = Pt(22)
    r1.font.color.rgb = pal['primary']
    r1.font.bold = True
    r1.font.name = 'Times New Roman'
    
    # Line 2: Title
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = display_title
    r2.font.size = Pt(28)
    r2.font.color.rgb = pal['text_dark']
    r2.font.bold = True
    r2.font.name = 'Times New Roman'
    p2.space_before = Pt(4)
    
    # Line 3: Topic
    if topic:
        p3 = tf.add_paragraph()
        r3 = p3.add_run()
        r3.text = f'Chủ đề: {topic}'
        r3.font.size = Pt(16)
        r3.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        r3.font.italic = True
        r3.font.name = 'Times New Roman'
        p3.space_before = Pt(8)
    
    # Teacher info
    info_l = card_l + Inches(0.5)
    info_t = card_t + card_h - Inches(1.2)
    info_w = card_w - Inches(1)
    info_h = Inches(1.0)
    
    txBox2 = slide.shapes.add_textbox(info_l, info_t, info_w, info_h)
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p_school = tf2.paragraphs[0]
    r_school = p_school.add_run()
    r_school.text = 'Trường TH & THCS UNIGO'
    r_school.font.size = Pt(16)
    r_school.font.color.rgb = pal['text_dark']
    r_school.font.name = 'Times New Roman'
    
    p_gv = tf2.add_paragraph()
    r_gv = p_gv.add_run()
    r_gv.text = 'Giáo viên: Đậu Đình Nguyên'
    r_gv.font.size = Pt(16)
    r_gv.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r_gv.font.name = 'Times New Roman'


def _add_objectives_slide(slide, prs, grade_cfg, lesson_data, pal):
    """Create objectives slide."""
    _set_slide_bg(slide, pal['bg'])
    sw = prs.slide_width
    
    # Card
    card_w = Inches(11.5)
    card_h = Inches(4.8)
    card_l = (sw - card_w) // 2
    card_t = CONTENT_TOP
    
    _add_card_with_accent(slide, card_l, card_t, card_w, card_h, pal, 'left')
    
    # Title
    _add_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.2),
                 Inches(10), Inches(0.5),
                 '🎯  Mục tiêu bài học', Pt(24), pal['primary'], bold=True)
    
    # Objectives
    objectives = lesson_data.get('objectives', [])
    if not objectives:
        objectives = ['Nắm được kiến thức cơ bản của bài học.',
                      'Vận dụng kiến thức vào thực hành.',
                      'Phát triển năng lực tư duy và hợp tác.']
    
    # Limit to 4 objectives max
    objectives = objectives[:4]
    
    _add_bullet_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.9),
                        Inches(10.5), Inches(3.6), objectives, pal,
                        font_size=Pt(18))


def _get_activity_content(act):
    """Extract content from an activity dict into bullet points."""
    bullets = []
    
    # For TH format (gv/hs keys)
    if 'gv' in act:
        gv_text = act['gv'].strip()
        for line in gv_text.split('\n'):
            line = line.strip()
            if line.startswith('*') or line.startswith('-'):
                line = line.lstrip('*- ').strip()
            # Strip "GV " prefix
            if line.startswith('GV '):
                line = line[3:]
            # Strip "Mục tiêu:" prefix
            if line.startswith('Mục tiêu:'):
                continue
            if len(line) > 15 and len(line) < 120:
                # Capitalize first letter
                line = line[0].upper() + line[1:] if line else line
                bullets.append(line)
    
    # For THCS format (content/product keys)
    if 'content' in act and act['content']:
        content = act['content']
        # Split by semicolons
        parts = re.split(r'[;]', content)
        for part in parts:
            part = part.strip().rstrip('.')
            if len(part) > 10 and len(part) < 120:
                part = part[0].upper() + part[1:] if part else part
                bullets.append(part)
    
    # For THCS format with gv_bullets extracted from tables
    if 'gv_bullets' in act and act['gv_bullets']:
        for b in act['gv_bullets']:
            b = b.strip()
            if b.startswith('GV '):
                b = b[3:]
            if len(b) > 10 and len(b) < 120:
                b = b[0].upper() + b[1:] if b else b
                if b not in bullets:
                    bullets.append(b)
    
    if 'product' in act and act['product']:
        product = act['product'][:100]
        product = product[0].upper() + product[1:] if product else product
        bullets.append(f'📋 Sản phẩm: {product}')
    
    # Limit and clean
    cleaned = []
    seen = set()
    for b in bullets:
        b = b.strip()
        if b and len(b) > 5 and b not in seen:
            seen.add(b)
            # Truncate long bullets
            if len(b) > 90:
                b = b[:87] + '...'
            cleaned.append(b)
    
    return cleaned[:4]  # Max 4 bullets per slide



def _add_activity_slide(slide, prs, title_text, act, pal, card_side='left'):
    """Create an activity slide with card layout."""
    _set_slide_bg(slide, pal['bg'])
    sw = prs.slide_width
    
    card_w = Inches(7.5)
    card_h = Inches(4.8)
    
    if card_side == 'left':
        card_l = Inches(0.50)
    else:
        card_l = sw - card_w - Inches(0.50)
    card_t = CONTENT_TOP
    
    accent_side = 'left' if card_side == 'left' else 'right'
    _add_card_with_accent(slide, card_l, card_t, card_w, card_h, pal, accent_side)
    
    # Title
    title_l = card_l + Inches(0.4)
    _add_textbox(slide, title_l, card_t + Inches(0.2),
                 card_w - Inches(0.8), Inches(0.5),
                 title_text, Pt(22), pal['primary'], bold=True)
    
    # Content bullets
    bullets = _get_activity_content(act)
    if not bullets:
        act_title = act.get('title', 'Hoạt động')
        bullets = [f'Thực hiện hoạt động: {act_title}',
                   'Học sinh làm việc theo hướng dẫn của giáo viên.',
                   'Hoàn thành sản phẩm học tập.']
    
    _add_bullet_textbox(slide, title_l, card_t + Inches(0.9),
                        card_w - Inches(0.8), Inches(3.6), bullets, pal,
                        font_size=Pt(18))
    
    # Add colored accent rectangle on the opposite side (decorative)
    if card_side == 'left':
        deco_l = sw - Inches(5.2)
    else:
        deco_l = Inches(0.50)
    deco_w = Inches(4.8)
    deco_h = Inches(4.8)
    deco = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, deco_l, card_t, deco_w, deco_h,
                      fill_color=pal['primary'])
    deco.adjustments[0] = 0.04
    # Make it semi-transparent by setting alpha
    fill_el = deco._element.spPr.solidFill
    if fill_el is not None:
        srgb = fill_el.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
        if srgb is not None:
            etree.SubElement(srgb, '{http://schemas.openxmlformats.org/drawingml/2006/main}alpha',
                             attrib={'val': '20000'})  # 20% opacity
    
    # Add emoji decoration on the decorative card
    emoji_map = {
        'khởi_động': '🎮',
        'kiến_thức': '📖',
        'luyện_tập': '✏️',
        'vận_dụng': '🚀',
    }
    emoji = emoji_map.get(act.get('type', ''), '📌')
    _add_textbox(slide, deco_l + deco_w//2 - Inches(0.5), card_t + deco_h//2 - Inches(0.5),
                 Inches(1), Inches(1),
                 emoji, Pt(48), pal['text_light'], alignment=PP_ALIGN.CENTER)


def _add_generic_knowledge_slide(slide, prs, display_title, pal):
    """Add a generic knowledge slide when no activities are extracted."""
    _set_slide_bg(slide, pal['bg'])
    sw = prs.slide_width
    
    card_w = Inches(12)
    card_h = Inches(4.8)
    card_l = (sw - card_w) // 2
    card_t = CONTENT_TOP
    
    _add_card_with_accent(slide, card_l, card_t, card_w, card_h, pal, 'left')
    
    _add_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.2),
                 Inches(11), Inches(0.5),
                 f'📖  {display_title}', Pt(24), pal['primary'], bold=True)
    
    bullets = [
        'Tìm hiểu nội dung bài học trong SGK.',
        'Thảo luận nhóm về các khái niệm chính.',
        'Ghi chép kiến thức trọng tâm.',
    ]
    
    _add_bullet_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.9),
                        Inches(11), Inches(3.6), bullets, pal)


def _add_summary_slide(slide, prs, display_title, lesson_data, pal):
    """Create summary slide."""
    _set_slide_bg(slide, pal['bg'])
    sw = prs.slide_width
    
    # Full-width card
    card_w = Inches(12.30)
    card_h = Inches(4.8)
    card_l = (sw - card_w) // 2
    card_t = CONTENT_TOP
    
    _add_card_with_accent(slide, card_l, card_t, card_w, card_h, pal, 'left')
    
    _add_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.2),
                 Inches(11), Inches(0.5),
                 f'📌  Ghi nhớ — {display_title}', Pt(24), pal['primary'], bold=True)
    
    # Summary bullets from objectives
    summary = lesson_data.get('objectives', [])[:3]
    if not summary:
        summary = ['Nắm vững kiến thức cơ bản của bài học.',
                   'Hoàn thành bài tập luyện tập.',
                   'Chuẩn bị trước nội dung bài tiếp theo.']
    
    summary.append('📚 Bài tập về nhà: Xem lại nội dung bài học và chuẩn bị bài tiếp theo.')
    
    _add_bullet_textbox(slide, card_l + Inches(0.5), card_t + Inches(0.9),
                        Inches(11), Inches(3.6), summary, pal)


def _add_thankyou_slide(slide, prs, grade_cfg, pal):
    """Create thank you slide."""
    _set_slide_bg(slide, pal['primary'])
    sw = prs.slide_width
    sh = prs.slide_height
    
    # Center card
    card_w = Inches(8)
    card_h = Inches(3.5)
    card_l = (sw - card_w) // 2
    card_t = Inches(2.0)
    
    card = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, card_l, card_t, card_w, card_h,
                      fill_color=RGBColor(0xFF, 0xFF, 0xFF))
    card.adjustments[0] = 0.04
    
    # Thank you text
    txBox = slide.shapes.add_textbox(card_l + Inches(0.5), card_t + Inches(0.5),
                                      card_w - Inches(1), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = 'CẢM ƠN CÁC EM! 🎉'
    r.font.size = Pt(32)
    r.font.color.rgb = pal['primary']
    r.font.bold = True
    r.font.name = 'Times New Roman'
    
    # Subtitle
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = 'Hẹn gặp lại ở tiết học sau!'
    r2.font.size = Pt(20)
    r2.font.color.rgb = pal['text_dark']
    r2.font.name = 'Times New Roman'
    p2.space_before = Pt(12)
    
    # School info
    txBox2 = slide.shapes.add_textbox(card_l + Inches(0.5), card_t + Inches(2.2),
                                       card_w - Inches(1), Inches(1.0))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p3 = tf2.paragraphs[0]
    p3.alignment = PP_ALIGN.CENTER
    r3 = p3.add_run()
    r3.text = 'Trường TH & THCS UNIGO  •  GV: Đậu Đình Nguyên'
    r3.font.size = Pt(14)
    r3.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    r3.font.name = 'Times New Roman'


def _add_footer(slide, prs):
    """Add UNIGO footer bar to slide."""
    sw = prs.slide_width
    
    footer_h = Inches(0.55)
    footer_t = prs.slide_height - footer_h
    
    # Blue bar
    bar = _add_shape(slide, MSO_SHAPE.RECTANGLE, 0, footer_t, sw, footer_h,
                     fill_color=RGBColor(0x00, 0x70, 0xC0))
    
    # Footer text
    txBox = slide.shapes.add_textbox(Inches(0.3), footer_t + Inches(0.05),
                                      sw - Inches(0.6), footer_h - Inches(0.1))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = 'TRƯỜNG TIỂU HỌC VÀ THCS UNIGO'
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.bold = True
    r.font.name = 'Times New Roman'


# ══════════════════════════════════════════════════════════════════
#                         MAIN
# ══════════════════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('  GENERATE SLIDES BÀI 1 TIN HỌC — TẤT CẢ CÁC LỚP')
    print('=' * 60)
    
    results = []
    
    for grade_cfg in GRADES:
        key = grade_cfg['key']
        folder = grade_cfg['folder']
        label = grade_cfg['label']
        
        print(f'\n{"─"*50}')
        print(f'  {label}: Đang xử lý...')
        print(f'{"─"*50}')
        
        try:
            # 1. Read KHBD
            lesson_data = read_khbd(grade_cfg)
            print(f'  ✓ KHBD: {lesson_data["title"]}')
            print(f'    Topic: {lesson_data["topic"]}')
            print(f'    Objectives: {len(lesson_data["objectives"])}')
            print(f'    Activities: {len(lesson_data["activities"])}')
            
            # 2. Create slides
            prs, slides_created = create_presentation(grade_cfg, lesson_data)
            print(f'  ✓ Slides: {len(slides_created)} slides')
            for s in slides_created:
                print(f'    • {s}')
            
            # 3. Save
            out_dir = os.path.join(KHBD_ROOT, folder, 'Bài_01')
            os.makedirs(out_dir, exist_ok=True)
            
            # Generate filename
            title_clean = lesson_data['title']
            title_clean = re.sub(r'^BÀI\s*\d+\.\s*', '', title_clean, flags=re.IGNORECASE)
            title_clean = re.sub(r'[^\w\s]', '', title_clean)
            title_clean = title_clean.strip().replace(' ', '_')
            if key == 'TTH':
                fname = f'Slide_Tin_hoc_Tien_TH_Bai01_{title_clean}.pptx'
            else:
                fname = f'Slide_Tin_hoc_Lop_{key}_Bai01_{title_clean}.pptx'
            
            out_path = os.path.join(out_dir, fname)
            prs.save(out_path)
            print(f'  ✓ Saved: {out_path}')
            
            results.append((label, '✅', fname, len(slides_created)))
            
        except Exception as e:
            print(f'  ✗ ERROR: {e}')
            import traceback
            traceback.print_exc()
            results.append((label, '❌', str(e), 0))
    
    # Summary
    print(f'\n{"="*60}')
    print('  KẾT QUẢ TỔNG HỢP')
    print(f'{"="*60}')
    for label, status, fname, count in results:
        print(f'  {status} {label}: {fname} ({count} slides)')
    
    print(f'\nHOÀN THÀNH! Tổng: {sum(1 for r in results if r[1]=="✅")}/{len(results)} file thành công.')


if __name__ == '__main__':
    main()
