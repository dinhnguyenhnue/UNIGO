"""
Script trích xuất Yêu cầu cần đạt (YCCĐ) từ tất cả file KHDH (Kế hoạch dạy học)
trong thư mục "Nguyên đã làm".

Output: D:/UNIGO/.agents/skills/tao-khbd/references/yccd_khdh_data.json

Cấu trúc output JSON:
{
  "Tin học": {
    "Lớp 3": [
      {"stt": "1", "bai": "Bài 1: Thông tin và quyết định", "so_tiet": "1", "ppct": "1", "yccd": "..."},
      ...
    ],
    ...
  },
  "Robotics": {
    "Lớp 1": [...],
    ...
  }
}
"""
import sys
import io
import os
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document

BASE_DIR = r"D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm"
OUTPUT_PATH = r"D:\UNIGO\.agents\skills\tao-khbd\references\yccd_khdh_data.json"

def extract_yccd_from_khdh(filepath):
    """Trích xuất bảng PPCT (Table 3) từ file KHDH."""
    doc = Document(filepath)
    
    # Tìm bảng có header chứa "Yêu cầu cần đạt"
    target_table = None
    for t in doc.tables:
        if len(t.rows) < 2:
            continue
        header_cells = [c.text.strip() for c in t.rows[0].cells]
        if any("Yêu cầu cần đạt" in h for h in header_cells):
            # Kiểm tra đây là bảng PPCT (có STT, Bài học) chứ không phải bảng Đánh giá
            if any("STT" in h or "Bài học" in h for h in header_cells):
                target_table = t
                break
    
    if target_table is None:
        return []
    
    lessons = []
    header = [c.text.strip() for c in target_table.rows[0].cells]
    
    # Xác định index các cột
    col_map = {}
    for i, h in enumerate(header):
        h_lower = h.lower()
        if 'stt' in h_lower:
            col_map['stt'] = i
        elif 'bài học' in h_lower or 'bài' in h_lower:
            col_map['bai'] = i
        elif 'số tiết' in h_lower:
            col_map['so_tiet'] = i
        elif 'ppct' in h_lower or 'tiết theo' in h_lower:
            col_map['ppct'] = i
        elif 'yêu cầu cần đạt' in h_lower:
            col_map['yccd'] = i
    
    for row in target_table.rows[1:]:
        cells = [c.text.strip() for c in row.cells]
        
        # Kiểm tra nếu là dòng gộp (chủ đề) - tất cả cells có nội dung giống nhau
        unique_cells = set(cells)
        if len(unique_cells) == 1 and len(cells) > 1:
            # Dòng chủ đề gộp
            lessons.append({
                "type": "chu_de",
                "ten_chu_de": cells[0]
            })
            continue
        
        # Dòng bài học thường
        entry = {"type": "bai_hoc"}
        for key, idx in col_map.items():
            if idx < len(cells):
                entry[key] = cells[idx]
        
        # Bỏ qua dòng rỗng
        if all(not entry.get(k, '') for k in ['stt', 'bai', 'yccd']):
            continue
        
        lessons.append(entry)
    
    return lessons


def main():
    all_data = {}
    
    for mon in ["Tin học", "Robotics"]:
        all_data[mon] = {}
        subdir = f"Kế hoạch dạy học {mon} từng lớp"
        dir_path = os.path.join(BASE_DIR, subdir)
        
        if not os.path.exists(dir_path):
            print(f"  [SKIP] Thư mục không tồn tại: {dir_path}")
            continue
        
        for filename in sorted(os.listdir(dir_path)):
            if not filename.endswith('.docx') or filename.startswith('~'):
                continue
            
            # Xác định tên lớp từ filename
            # VD: "Kế hoạch dạy học môn Tin học - Lớp 3 - 2026 - 2027.docx"
            parts = filename.split(' - ')
            lop_name = None
            for p in parts:
                p = p.strip()
                if p.startswith('Lớp') or p.startswith('Tiền'):
                    lop_name = p
                    break
            
            if not lop_name:
                lop_name = filename  # fallback
            
            filepath = os.path.join(dir_path, filename)
            print(f"  Đang trích xuất: {mon} / {lop_name}...")
            
            try:
                lessons = extract_yccd_from_khdh(filepath)
                all_data[mon][lop_name] = lessons
                
                # Đếm số bài thực
                bai_count = sum(1 for l in lessons if l.get('type') == 'bai_hoc')
                chu_de_count = sum(1 for l in lessons if l.get('type') == 'chu_de')
                print(f"    → {bai_count} bài học, {chu_de_count} chủ đề")
                
            except Exception as e:
                print(f"    [ERROR] {e}")
    
    # Lưu JSON
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ Đã lưu: {OUTPUT_PATH}")
    print(f"   Tổng: {sum(len(v) for v in all_data.values())} khối lớp")


if __name__ == '__main__':
    main()
