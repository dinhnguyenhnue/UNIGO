# -*- coding: utf-8 -*-
"""
Script copy paste nội dung giáo án gốc (Initiate/Kinder/Excel) vào KHBD Robotics.

Logic chính:
- Đọc giáo án gốc → Parse sections (Mục tiêu kiến thức, Thiết bị, Tiến trình 6 mục)
- Mở KHBD đích → Thay thế nội dung paragraphs tương ứng
- Giữ nguyên: Table 0 (thông tin GV), Table 5 (chữ ký), Header/Footer, 
  Năng lực đặc thù/Năng lực số/Năng lực chung/Phẩm chất

Nội dung được copy:
- Phần I.1 Kiến thức (từ giáo án gốc)
- Phần III Tiến trình: Nội dung chi tiết câu hỏi, đáp án, kiến thức, 
  hướng dẫn lắp ráp → Điền vào bảng 2 cột (Hoạt động GV-HS | Kết quả cần đạt)
"""
import sys, os, re, shutil, traceback
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt
from copy import deepcopy

BASE_SRC = r'D:\UNIGO\Phân phối chương trình\Robotics'
BASE_DST = r'D:\UNIGO\KHBD_Robotics'

# ============================================================
# MAPPING: Giáo án gốc → Lớp + Bài số
# ============================================================

def build_initiate_mapping():
    """Initiate: I 03→I 10, 4 bài/bộ → Lớp 1 Bài 1-32"""
    init_dir = os.path.join(BASE_SRC, 'Giáo án-Initiate')
    files = sorted(os.listdir(init_dir))
    mapping = []
    for i, f in enumerate(files):
        bai_so = i + 1  # 1-based
        mapping.append({
            'src': os.path.join(init_dir, f),
            'bai_so': bai_so,
            'filename': f,
        })
    return mapping

def build_kinder_mapping():
    """Kinder: K+3→K+10, 4 bài/bộ → Lớp 2,3,4 Bài 1-32
    Sắp xếp theo thứ tự K+3 < K+4 < K+5 ... < K+10
    """
    kind_dir = os.path.join(BASE_SRC, 'Giáo án-Kinder')
    files = os.listdir(kind_dir)
    
    def sort_key(f):
        # Parse K+XX - YY pattern
        m = re.match(r'K\+(\d+)\s*-\s*(\d+)', f)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        return (999, 999)
    
    files = sorted(files, key=sort_key)
    mapping = []
    for i, f in enumerate(files):
        bai_so = i + 1
        mapping.append({
            'src': os.path.join(kind_dir, f),
            'bai_so': bai_so,
            'filename': f,
        })
    return mapping

def build_excel_mapping():
    """Excel: 01-01 → 1-12 → Lớp 5,6,7,8 Bài 1-12"""
    exc_dir = os.path.join(BASE_SRC, 'Giáo án-Excel')
    files = sorted(os.listdir(exc_dir))
    mapping = []
    for i, f in enumerate(files):
        bai_so = i + 1
        mapping.append({
            'src': os.path.join(exc_dir, f),
            'bai_so': bai_so,
            'filename': f,
        })
    return mapping


# ============================================================
# PARSE GIÁO ÁN GỐC
# ============================================================

def parse_source_content(src_path):
    """
    Đọc giáo án gốc và trích xuất nội dung theo sections.
    Returns dict with keys: 'kien_thuc', 'thiet_bi', 'tien_trinh_sections'
    """
    doc = Document(src_path)
    paragraphs = doc.paragraphs
    
    result = {
        'kien_thuc_lines': [],      # I.1 Về kiến thức
        'nang_luc_lines': [],       # I.2 Về năng lực  
        'pham_chat_lines': [],      # I.3 Về phẩm chất
        'thiet_bi_lines': [],       # II. Thiết bị
        'tien_trinh': {},           # III. Tiến trình - dict by section name
    }
    
    current_section = None  # 'kien_thuc', 'nang_luc', 'pham_chat', 'thiet_bi', 'tien_trinh'
    current_tt_section = None  # '1. Khởi động', '2. Kiến thức', etc.
    
    for i, p in enumerate(paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        
        # Skip header lines
        if txt.startswith('KHUNG KẾ HOẠCH BÀI DẠY') or txt.startswith('BÀI') or txt.startswith('Môn học'):
            continue
        
        # Detect major sections
        if re.match(r'^I\.\s*MỤC TIÊU', txt, re.IGNORECASE):
            current_section = 'pre_kienthuc'
            continue
        elif re.match(r'^1\.\s*Về kiến thức', txt, re.IGNORECASE):
            current_section = 'kien_thuc'
            continue
        elif re.match(r'^2\.\s*Về năng lực', txt, re.IGNORECASE):
            current_section = 'nang_luc'
            continue
        elif re.match(r'^3\.\s*Về phẩm chất', txt, re.IGNORECASE):
            current_section = 'pham_chat'
            continue
        elif re.match(r'^II\.\s*THIẾT BỊ', txt, re.IGNORECASE):
            current_section = 'thiet_bi'
            continue
        elif re.match(r'^III\.\s*TIẾN TRÌNH', txt, re.IGNORECASE):
            current_section = 'tien_trinh'
            continue
        
        # Detect tiến trình sub-sections
        if current_section == 'tien_trinh':
            m = re.match(r'^(\d+)\.\s*(Khởi động|Kiến thức|Luyện tập|Vận dụng|Tự học|Sáng tạo|Tự kiểm tra)', txt, re.IGNORECASE)
            if m:
                current_tt_section = m.group(0)
                if current_tt_section not in result['tien_trinh']:
                    result['tien_trinh'][current_tt_section] = []
                result['tien_trinh'][current_tt_section].append(txt)
                continue
        
        # Collect content
        if current_section == 'kien_thuc':
            result['kien_thuc_lines'].append(txt)
        elif current_section == 'nang_luc':
            result['nang_luc_lines'].append(txt)
        elif current_section == 'pham_chat':
            result['pham_chat_lines'].append(txt)
        elif current_section == 'thiet_bi':
            result['thiet_bi_lines'].append(txt)
        elif current_section == 'tien_trinh' and current_tt_section:
            result['tien_trinh'][current_tt_section].append(txt)
    
    return result


# ============================================================
# APPLY TO KHBD
# ============================================================

def find_khbd_files_for_lesson(lop, bai_so):
    """Tìm file KHBD tương ứng với bài số cho một lớp."""
    lop_dir = os.path.join(BASE_DST, f'Lớp_{lop}')
    if not os.path.exists(lop_dir):
        return []
    
    results = []
    for tuan in sorted(os.listdir(lop_dir)):
        tuan_dir = os.path.join(lop_dir, tuan)
        if not os.path.isdir(tuan_dir):
            continue
        for f in os.listdir(tuan_dir):
            if not f.endswith('.docx'):
                continue
            # Match pattern like: KHBD_Robotics_Lớp_1_Tiet01_Bai_1_...
            m = re.search(r'Tiet\d+_Bai_(\d+)_', f)
            if m and int(m.group(1)) == bai_so:
                results.append(os.path.join(tuan_dir, f))
    return results


def build_gv_hs_content(section_lines):
    """Chuyển đổi nội dung section thành text cho cột GV-HS trong bảng 2 cột."""
    lines = []
    for line in section_lines:
        lines.append(line)
    return '\n'.join(lines)


def build_ket_qua_content(section_name, section_lines):
    """Tạo nội dung cho cột Kết quả cần đạt từ nội dung section."""
    # Extract sản phẩm hoạt động and mục tiêu
    ket_qua = []
    in_san_pham = False
    in_muc_tieu = False
    
    for line in section_lines:
        if '- Mục tiêu' in line or '- Mục tiêu:' in line:
            in_muc_tieu = True
            in_san_pham = False
            continue
        elif '- Sản phẩm' in line:
            in_san_pham = True
            in_muc_tieu = False
            # Check if content is on same line
            m = re.match(r'^-\s*Sản phẩm[^:]*:\s*(.+)$', line)
            if m and m.group(1).strip():
                ket_qua.append(m.group(1).strip())
            continue
        elif re.match(r'^-\s*(Nội dung|Tổ chức thực hiện)', line):
            in_san_pham = False
            in_muc_tieu = False
            continue
        
        if in_san_pham and line.startswith('+'):
            ket_qua.append(line)
        elif in_muc_tieu and line.startswith('+'):
            ket_qua.append(line)
    
    if not ket_qua:
        # Fallback: use mục tiêu lines
        for line in section_lines:
            if line.startswith('- Mục tiêu:'):
                m = re.match(r'^-\s*Mục tiêu:\s*(.+)$', line)
                if m:
                    ket_qua.append(m.group(1).strip())
    
    return '\n'.join(ket_qua) if ket_qua else 'HS hoàn thành nhiệm vụ học tập.'


def build_hoat_dong_content(section_lines):
    """Xây dựng nội dung cho cột Hoạt động GV-HS từ section lines."""
    gv_hs = []
    in_tochuc = False
    
    for line in section_lines:
        # Check for Tổ chức thực hiện
        if re.match(r'^-\s*Tổ chức thực hiện', line):
            in_tochuc = True
            continue
        
        if in_tochuc:
            gv_hs.append(line)
        elif line.startswith('*') or line.startswith('=>') or line.startswith('=\u003e'):
            gv_hs.append(line)
    
    # If no Tổ chức lines found, include all content after Mục tiêu, Nội dung, Sản phẩm
    if not gv_hs:
        skip_prefixes = ['- Mục tiêu', '- Nội dung', '- Sản phẩm']
        for line in section_lines:
            if not any(line.startswith(p) for p in skip_prefixes):
                gv_hs.append(line)
    
    return '\n'.join(gv_hs)


def apply_source_to_khbd(src_content, khbd_path, dry_run=False):
    """
    Áp dụng nội dung giáo án gốc vào file KHBD.
    
    Thay thế:
    1. Phần "Yêu cầu cần đạt" / Kiến thức trong paragraphs
    2. Nội dung bảng 2 cột (Hoạt động GV-HS | Kết quả cần đạt) - Tables 1-4
    
    Giữ nguyên:
    - Table 0 (thông tin GV)
    - Table 5 (chữ ký)  
    - Header/Footer
    - Năng lực đặc thù/Năng lực số/Năng lực chung/Phẩm chất paragraphs
    """
    doc = Document(khbd_path)
    
    # === 1. Thay thế nội dung Kiến thức trong Mục tiêu ===
    # Tìm phần "Yêu cầu cần đạt" hoặc "Kiến thức" 
    found_kienthuc_start = -1
    found_kienthuc_end = -1
    
    for idx, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        # Tìm bắt đầu phần Yêu cầu cần đạt (sau "- Sau bài học này em sẽ:" hoặc "1. Kiến thức:")
        if txt.startswith('- Sau bài học này em sẽ:') or re.match(r'^1\.\s*Kiến thức', txt):
            found_kienthuc_start = idx + 1
        elif txt.startswith('1. Năng lực:') or re.match(r'^1\.\s*Năng lực:', txt):
            if found_kienthuc_start > 0:
                found_kienthuc_end = idx
                break
        elif re.match(r'^1\.1\.\s*Năng lực đặc thù', txt):
            if found_kienthuc_start > 0:
                found_kienthuc_end = idx
                break
        elif re.match(r'^2\.\s*Năng lực:', txt):
            if found_kienthuc_start > 0:
                found_kienthuc_end = idx
                break
    
    if found_kienthuc_start > 0 and found_kienthuc_end > 0 and src_content['kien_thuc_lines']:
        # Replace kiến thức paragraphs
        # Clear existing content in range
        for idx in range(found_kienthuc_start, found_kienthuc_end):
            p = doc.paragraphs[idx]
            for run in p.runs:
                run.text = ''
        
        # Fill new content
        new_lines = src_content['kien_thuc_lines']
        for i, line in enumerate(new_lines):
            if i < (found_kienthuc_end - found_kienthuc_start):
                target_p = doc.paragraphs[found_kienthuc_start + i]
                # Clear and set new text
                for run in target_p.runs:
                    run.text = ''
                if target_p.runs:
                    target_p.runs[0].text = line
                else:
                    target_p.text = line
    
    # === 2. Thay nội dung bảng Hoạt động (Tables 1-4) ===
    # Map giáo án gốc sections → 4 hoạt động KHBD
    # Source có 6 mục: Khởi động, Kiến thức, Luyện tập, Vận dụng, Tự học, Tự kiểm tra
    # KHBD có 4 hoạt động: Khởi động, Khám phá/Kiến thức, Luyện tập, Vận dụng/Sáng tạo
    
    tt_keys = list(src_content['tien_trinh'].keys())
    
    # Build 4 activity blocks from source
    activities = []
    
    # Hoạt động 1: Khởi động
    khoi_dong = None
    for k in tt_keys:
        if 'Khởi động' in k or 'Let\'s think' in k:
            khoi_dong = src_content['tien_trinh'][k]
            break
    activities.append(khoi_dong or [])
    
    # Hoạt động 2: Kiến thức / Khám phá
    kien_thuc = None
    for k in tt_keys:
        if 'Kiến thức' in k or 'Let\'s learn' in k:
            kien_thuc = src_content['tien_trinh'][k]
            break
    activities.append(kien_thuc or [])
    
    # Hoạt động 3: Luyện tập
    luyen_tap = None
    for k in tt_keys:
        if 'Luyện tập' in k or 'Let\'s build' in k:
            luyen_tap = src_content['tien_trinh'][k]
            break
    activities.append(luyen_tap or [])
    
    # Hoạt động 4: Vận dụng + Tự học + Tự kiểm tra (gộp lại)
    van_dung_lines = []
    for k in tt_keys:
        if any(x in k for x in ['Vận dụng', 'Tự học', 'Sáng tạo', 'Tự kiểm tra', 'My own', 'Let\'s move', 'Let\'s check']):
            van_dung_lines.extend(src_content['tien_trinh'][k])
    activities.append(van_dung_lines)
    
    # Apply to tables 1-4 (0-indexed: tables[1] through tables[4])
    for act_idx in range(4):
        table_idx = act_idx + 1  # Table 1-4
        if table_idx >= len(doc.tables):
            continue
        
        table = doc.tables[table_idx]
        if len(table.rows) < 2:
            continue
        
        act_lines = activities[act_idx]
        if not act_lines:
            continue
        
        # Build GV-HS content and Kết quả content
        gv_hs_text = build_hoat_dong_content(act_lines)
        ket_qua_text = build_ket_qua_content(f'HĐ{act_idx+1}', act_lines)
        
        # Replace content in row 1 (data row)
        row = table.rows[1]
        
        # Cell 0: GV-HS activities
        cell0 = row.cells[0]
        # Clear existing paragraphs
        for p in cell0.paragraphs:
            for run in p.runs:
                run.text = ''
        # Set new content
        if cell0.paragraphs:
            first_p = cell0.paragraphs[0]
            if first_p.runs:
                first_p.runs[0].text = gv_hs_text
            else:
                first_p.text = gv_hs_text
            # Set font
            for run in first_p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
        
        # Cell 1: Kết quả cần đạt
        cell1 = row.cells[1]
        for p in cell1.paragraphs:
            for run in p.runs:
                run.text = ''
        if cell1.paragraphs:
            first_p = cell1.paragraphs[0]
            if first_p.runs:
                first_p.runs[0].text = ket_qua_text
            else:
                first_p.text = ket_qua_text
            for run in first_p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
    
    if not dry_run:
        doc.save(khbd_path)
    
    return True


# ============================================================
# MAIN EXECUTION
# ============================================================

def process_all():
    stats = {'success': 0, 'skip': 0, 'error': 0, 'no_match': 0}
    
    # === LỚP 1: Initiate ===
    print('=' * 60)
    print('PROCESSING LỚP 1 (Initiate)')
    print('=' * 60)
    init_map = build_initiate_mapping()
    for entry in init_map:
        bai_so = entry['bai_so']
        khbd_files = find_khbd_files_for_lesson(1, bai_so)
        if not khbd_files:
            print(f'  [SKIP] Bài {bai_so} ({entry["filename"]}): Không tìm thấy KHBD tương ứng')
            stats['no_match'] += 1
            continue
        
        try:
            src_content = parse_source_content(entry['src'])
            for khbd in khbd_files:
                apply_source_to_khbd(src_content, khbd)
                print(f'  [OK] Bài {bai_so}: {entry["filename"]} → {os.path.basename(khbd)}')
                stats['success'] += 1
        except Exception as e:
            print(f'  [ERROR] Bài {bai_so}: {e}')
            traceback.print_exc()
            stats['error'] += 1
    
    # === LỚP 2, 3, 4: Kinder ===
    kind_map = build_kinder_mapping()
    for lop in [2, 3, 4]:
        print(f'\n{"=" * 60}')
        print(f'PROCESSING LỚP {lop} (Kinder)')
        print(f'{"=" * 60}')
        for entry in kind_map:
            bai_so = entry['bai_so']
            khbd_files = find_khbd_files_for_lesson(lop, bai_so)
            if not khbd_files:
                print(f'  [SKIP] Bài {bai_so} ({entry["filename"]}): Không tìm thấy KHBD')
                stats['no_match'] += 1
                continue
            
            try:
                src_content = parse_source_content(entry['src'])
                for khbd in khbd_files:
                    apply_source_to_khbd(src_content, khbd)
                    print(f'  [OK] Bài {bai_so}: {entry["filename"]} → {os.path.basename(khbd)}')
                    stats['success'] += 1
            except Exception as e:
                print(f'  [ERROR] Bài {bai_so}: {e}')
                traceback.print_exc()
                stats['error'] += 1
    
    # === LỚP 5, 6, 7, 8: Excel ===
    exc_map = build_excel_mapping()
    for lop in [5, 6, 7, 8]:
        print(f'\n{"=" * 60}')
        print(f'PROCESSING LỚP {lop} (Excel)')
        print(f'{"=" * 60}')
        for entry in exc_map:
            bai_so = entry['bai_so']
            khbd_files = find_khbd_files_for_lesson(lop, bai_so)
            if not khbd_files:
                print(f'  [SKIP] Bài {bai_so} ({entry["filename"]}): Không tìm thấy KHBD')
                stats['no_match'] += 1
                continue
            
            try:
                src_content = parse_source_content(entry['src'])
                for khbd in khbd_files:
                    apply_source_to_khbd(src_content, khbd)
                    print(f'  [OK] Bài {bai_so}: {entry["filename"]} → {os.path.basename(khbd)}')
                    stats['success'] += 1
            except Exception as e:
                print(f'  [ERROR] Bài {bai_so}: {e}')
                traceback.print_exc()
                stats['error'] += 1
    
    print(f'\n{"=" * 60}')
    print(f'SUMMARY')
    print(f'{"=" * 60}')
    print(f'  Success: {stats["success"]}')
    print(f'  Skipped (no match): {stats["no_match"]}')
    print(f'  Errors: {stats["error"]}')
    print(f'  Total processed: {stats["success"] + stats["error"]}')


if __name__ == '__main__':
    process_all()
