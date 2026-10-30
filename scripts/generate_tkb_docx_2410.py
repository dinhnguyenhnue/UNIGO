import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def update_docx_tkb():
    src = r"D:\UNIGO\Thời khóa biểu giáo viên\Thời khóa biểu - Đậu Đình Nguyên.docx"
    dst = r"D:\UNIGO\Thời khóa biểu giáo viên\Thời khóa biểu - Đậu Đình Nguyên (Mới 24.10).docx"
    
    doc = docx.Document(src)
    table = doc.tables[0]
    
    # Update subtitle
    if len(doc.paragraphs) > 0:
        p0 = doc.paragraphs[0]
        p0.text = "TRƯỜNG TIỂU HỌC VÀ THCS UNIGO\nTHỜI KHOÁ BIỂU CÁ NHÂN GIÁO VIÊN\nNăm học 2026 – 2027 (Cập nhật 24/10/2026)"
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p0.runs:
            run.font.name = "Times New Roman"
            run.font.bold = True
            
    # Clear schedule cells (Rows 2 to 6, Rows 8 to 12; Cols 3 to 7)
    for r in range(2, 7):
        for c in range(3, 8):
            cell = table.cell(r, c)
            cell.text = ""
            set_cell_background(cell, "FFFFFF")
            
    for r in range(8, 13):
        for c in range(3, 8):
            cell = table.cell(r, c)
            cell.text = ""
            set_cell_background(cell, "FFFFFF")
            
    # Map for new schedule (Thứ Hai=3, Thứ Ba=4, Thứ Tư=5, Thứ Năm=6, Thứ Sáu=7)
    # Morning: T1=2, T2=3, T3=4, T4=5, T5=6
    # Afternoon: T1=8, T2=9, T3=10, T4=11, T5=12
    
    new_slots = {
        # Thứ Hai (Col 3)
        (4, 3):  ("🤖 Robotics\n(4C1)", "FFE5CC", "8C3C00"),
        (5, 3):  ("💻 Tin học\n(1A1)",  "D6E4FF", "003C8C"),
        (10, 3): ("💻 Tin học\n(2C1)",  "D6E4FF", "003C8C"),

        # Thứ Ba (Col 4)
        (2, 4):  ("💻 Tin học\n(1C1)",  "D6E4FF", "003C8C"),
        (5, 4):  ("🤖 Robotics\n(5C1)", "FFE5CC", "8C3C00"),
        (10, 4): ("💻 Tin học\n(7A1)",  "D6E4FF", "003C8C"),
        (11, 4): ("🤖 Robotics\n(7A1)", "FFE5CC", "8C3C00"),

        # Thứ Tư (Col 5)
        (3, 5):  ("🤖 Robotics\n(3A1)", "FFE5CC", "8C3C00"),
        (5, 5):  ("💻 Tin học\n(3C1)",  "D6E4FF", "003C8C"),
        (10, 5): ("💻 Tin học\n(4C1)",  "D6E4FF", "003C8C"),
        (11, 5): ("🤖 Robotics\n(2A1)", "FFE5CC", "8C3C00"),

        # Thứ Năm (Col 6)
        (2, 6):  ("💻 Tin học\n(TT3)",   "D6E4FF", "003C8C"),
        (3, 6):  ("💻 Tin học\n(TTH 1)", "D6E4FF", "003C8C"),
        (5, 6):  ("⚠️ TRÙNG LỊCH:\n🤖 Robotics (1A1)\n💻 Tin học (TTH2)", "FFD6D6", "B30000"),
        (8, 6):  ("💻 Tin học\n(3A1)",   "D6E4FF", "003C8C"),
        (9, 6):  ("💻 Tin học\n(5C1)",   "D6E4FF", "003C8C"),
        (10, 6): ("🤖 Robotics\n(2C1)",  "FFE5CC", "8C3C00"),
        (11, 6): ("🤖 Robotics\n(1C1)",  "FFE5CC", "8C3C00"),

        # Thứ Sáu (Col 7)
        (2, 7):  ("💻 Tin học\n(2A1)",  "D6E4FF", "003C8C"),
        (3, 7):  ("🤖 Robotics\n(3C1)", "FFE5CC", "8C3C00"),
        (5, 7):  ("💻 Tin học\n(6A1)",  "D6E4FF", "003C8C"),
        (6, 7):  ("🤖 Robotics\n(6A1)", "FFE5CC", "8C3C00"),
        (8, 7):  ("💻 Tin học\n(8A1)",  "D6E4FF", "003C8C"),
        (9, 7):  ("🤖 Robotics\n(8A1)", "FFE5CC", "8C3C00"),
    }
    
    for (r, c), (text, bg_hex, text_color_hex) in new_slots.items():
        cell = table.cell(r, c)
        cell.text = text
        set_cell_background(cell, bg_hex)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(9.5 if "TRÙNG" in text else 10.5)
            run.font.bold = True
            r_c = int(text_color_hex[0:2], 16)
            g_c = int(text_color_hex[2:4], 16)
            b_c = int(text_color_hex[4:6], 16)
            run.font.color.rgb = RGBColor(r_c, g_c, b_c)

    # Add explanatory note paragraph at the bottom
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(12)
    r_hdr = p_note.add_run("📌 Ghi chú điều chỉnh & Cảnh báo trùng lịch (TKB 24/10):\n")
    r_hdr.font.name = "Times New Roman"
    r_hdr.font.bold = True
    r_hdr.font.size = Pt(11)
    
    notes_lines = [
        "• Tin học 5C1: Đổi từ Thứ Ba Sáng T3 sang Thứ Năm Chiều T2.",
        "• Tin học TTH 1: Đổi từ Thứ Năm Chiều T2 lên Thứ Năm Sáng T2.",
        "• Robotics 2C1: Đổi từ Thứ Năm Sáng T4 xuống Thứ Năm Chiều T3.",
        "• ⚠️ CẢNH BÁO TRÙNG LỊCH: Thứ Năm Sáng Tiết 4 bị trùng 2 lớp [Robotics 1A1] và [Tin học TTH2].",
        "• 👉 Kiến nghị giải pháp: Đổi nội bộ lớp TTH2 giữa Tiết 3 (GVVN - Trang) và Tiết 4 (Tin - Nguyên). Khi đó Thầy Nguyên dạy TTH2 vào Tiết 3 (thầy đang trống) và dạy 1A1 Robotics vào Tiết 4."
    ]
    for line in notes_lines:
        r_line = p_note.add_run(line + "\n")
        r_line.font.name = "Times New Roman"
        r_line.font.size = Pt(10)
        if "⚠️" in line:
            r_line.font.bold = True
            r_line.font.color.rgb = RGBColor(204, 0, 0)
        elif "👉" in line:
            r_line.font.bold = True
            r_line.font.color.rgb = RGBColor(0, 51, 102)

    doc.save(dst)
    print("DOCX file saved successfully at:", dst)

if __name__ == "__main__":
    update_docx_tkb()
