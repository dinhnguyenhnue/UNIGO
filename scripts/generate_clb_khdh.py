# -*- coding: utf-8 -*-
"""
Script tạo Kế hoạch dạy học CLB Robotics UNIGO năm học 2026 - 2027
Bảo tồn Logo w:drawing ở Header và Footer nguyên bản.
Bảng phân phối 4 cột chuẩn XML viền, font Times New Roman 13pt.
"""

import sys, os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

DOCX_PATH = r"D:\UNIGO\CLB_Robotics\Kế hoạch dạy học CLB Robotics.docx"

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def format_all_fonts(doc):
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)
        for r in p.runs:
            r.font.name = "Times New Roman"
            r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
            if not r.font.size:
                r.font.size = Pt(13)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r._r.get_or_add_rPr().set(qn('w:rFonts'), 'Times New Roman')
                        if not r.font.size:
                            r.font.size = Pt(13)

DATES_35 = [
    "Thứ 7, 12/09/2026 (09h45 - 11h15)",
    "Thứ 7, 19/09/2026 (09h45 - 11h15)",
    "Thứ 7, 26/09/2026 (09h45 - 11h15)",
    "Thứ 7, 03/10/2026 (09h45 - 11h15)",
    "Thứ 7, 10/10/2026 (09h45 - 11h15)",
    "Thứ 7, 17/10/2026 (09h45 - 11h15)",
    "Thứ 7, 24/10/2026 (09h45 - 11h15)",
    "Thứ 7, 31/10/2026 (09h45 - 11h15)",
    "Thứ 7, 07/11/2026 (09h45 - 11h15)",
    "Thứ 7, 14/11/2026 (09h45 - 11h15)",
    "Thứ 7, 21/11/2026 (09h45 - 11h15)",
    "Thứ 7, 28/11/2026 (09h45 - 11h15)",
    "Thứ 7, 05/12/2026 (09h45 - 11h15)",
    "Thứ 7, 12/12/2026 (09h45 - 11h15)",
    "Thứ 7, 19/12/2026 (09h45 - 11h15)",
    "Thứ 7, 26/12/2026 (09h45 - 11h15)",
    "Thứ 7, 02/01/2027 (09h45 - 11h15)",
    # HK2
    "Thứ 7, 09/01/2027 (09h45 - 11h15)",
    "Thứ 7, 16/01/2027 (09h45 - 11h15)",
    "Thứ 7, 23/01/2027 (09h45 - 11h15)",
    "Thứ 7, 30/01/2027 (09h45 - 11h15)",
    "Thứ 7, 20/02/2027 (09h45 - 11h15)",
    "Thứ 7, 27/02/2027 (09h45 - 11h15)",
    "Thứ 7, 06/03/2027 (09h45 - 11h15)",
    "Thứ 7, 13/03/2027 (09h45 - 11h15)",
    "Thứ 7, 20/03/2027 (09h45 - 11h15)",
    "Thứ 7, 27/03/2027 (09h45 - 11h15)",
    "Thứ 7, 03/04/2027 (09h45 - 11h15)",
    "Thứ 7, 10/04/2027 (09h45 - 11h15)",
    "Thứ 7, 17/04/2027 (09h45 - 11h15)",
    "Thứ 7, 24/04/2027 (09h45 - 11h15)",
    "Thứ 7, 08/05/2027 (09h45 - 11h15)",
    "Thứ 7, 15/05/2027 (09h45 - 11h15)",
    "Thứ 7, 22/05/2027 (09h45 - 11h15)",
    "Thứ 7, 29/05/2027 (09h45 - 11h15)"
]

# Dữ liệu 35 buổi Nhóm Lớp 1 - 2 (HK1 Initiate -> HK2 Kinder)
DATA_LOP_1_2 = [
    # HK1: Initiate
    ("Buổi 1", "Bài 1: TẬP THỂ DỤC NÀO!\nNgày dạy: " + DATES_35[0], "90 phút\n(1 buổi)",
     "- Nhận biết một số dụng cụ thể dục tại nhà và tác dụng của bài tập đu xà đối với sức khỏe.\n- Kỹ năng nhận diện các linh kiện cơ bản và lắp ráp mô hình Đu xà OLLO Initiate.\n- Ý thức rèn luyện thân thể và an toàn lao động trong giờ học."),
    ("Buổi 2", "Bài 2: CHÚ CÚN DỄ THƯƠNG\nNgày dạy: " + DATES_35[1], "90 phút\n(1 buổi)",
     "- Tìm hiểu các đặc điểm sinh học nổi bật và tập tính thân thiện của loài chó.\n- Lắp ráp hoàn thiện robot Chú cún bốn chân với khớp cổ và đuôi cử động linh hoạt.\n- Bồi dưỡng tình yêu thương động vật và tinh thần chia sẻ đồ dùng học tập."),
    ("Buổi 3", "Bài 3: TĂNG CƯỜNG SỨC KHỎE\nNgày dạy: " + DATES_35[2], "90 phút\n(1 buổi)",
     "- Nhận biết lợi ích của hoạt động đi bộ hằng ngày đối với sự phát triển thể chất.\n- Lắp ráp mô hình Robot đi bộ với cơ cấu chân bước so le thăng bằng.\n- Hình thành thói quen rèn luyện sức khỏe và tính kiên trì trong thực hành."),
    ("Buổi 4", "Bài 4: CHÚ ỐC SÊN CHẬM CHẠP\nNgày dạy: " + DATES_35[3], "90 phút\n(1 buổi)",
     "- Quan sát cấu tạo vỏ xoắn ốc và cơ chế di chuyển bò trườn đặc trưng của loài ốc sên.\n- Lắp ráp mô hình Chú ốc sên thân mềm kết hợp bánh xe ẩn chuyển động êm ái.\n- Phát triển năng lực quan sát tự nhiên và tư duy sáng tạo hình khối."),
    ("Buổi 5", "Bài 5: XE CẢNH SÁT TUẦN TRA\nNgày dạy: " + DATES_35[4], "90 phút\n(1 buổi)",
     "- Nhận biết vai trò của lực lượng cảnh sát và phương tiện chuyên dụng giữ gìn trật tự.\n- Thực hành lắp ráp Xe cảnh sát bốn bánh vững chắc có còi báo hiệu mô phỏng.\n- Nâng cao ý thức chấp hành luật an toàn giao thông và quy tắc trường học."),
    ("Buổi 6", "Bài 6: KHÁM PHÁ XE CỨU HỎA\nNgày dạy: " + DATES_35[5], "90 phút\n(1 buổi)",
     "- Tìm hiểu cấu tạo cơ bản của xe chữa cháy: vòi phun, thang cứu hộ và đèn tín hiệu.\n- Lắp ráp mô hình Xe cứu hỏa OLLO Initiate với thang cơ học xoay và nâng hạ linh hoạt.\n- Nắm vững kỹ năng an toàn phòng cháy chữa cháy và xử lý khi có tình huống khẩn cấp."),
    ("Buổi 7", "Bài 7: DING DONG! NGƯỜI GIAO HÀNG ĐÃ ĐẾN\nNgày dạy: " + DATES_35[6], "90 phút\n(1 buổi)",
     "- Tìm hiểu quy trình vận chuyển hàng hóa và công việc của người giao hàng trong đời sống.\n- Lắp ráp mô hình Xe kéo/Xe chở hàng có thùng chứa đồ tiện lợi.\n- Phát triển kỹ năng làm việc nhóm và tính cẩn thận trong bảo quản chi tiết."),
    ("Buổi 8", "Bài 8: THẾ GIỚI KHỦNG LONG\nNgày dạy: " + DATES_35[7], "90 phút\n(1 buổi)",
     "- Khám phá thế giới cổ đại và đặc điểm nhận dạng của loài khủng long bạo chúa T-Rex.\n- Lắp ráp mô hình Khủng long với cơ cấu hàm đóng mở và chi trước mô phỏng.\n- Kích thích trí tưởng tượng khoa học và niềm đam mê nghiên cứu cổ sinh vật học."),
    ("Buổi 9", "Bài 9: HỒ BƠI MÙA HÈ (Đánh giá định kỳ 1)\nNgày dạy: " + DATES_35[8], "90 phút\n(1 buổi)",
     "- Tìm hiểu các kiểu bơi cơ bản, đồ dùng an toàn dưới nước và phòng chống đuối nước.\n- Lắp ráp Robot vận động viên bơi lội với cánh tay sải nước đối xứng.\n- ĐÁNH GIÁ ĐỊNH KỲ 1: Đánh giá kỹ năng nhận diện linh kiện Initiate và tốc độ lắp ráp chuẩn."),
    ("Buổi 10", "Bài 10: KHÁM PHÁ ĐẠI DƯƠNG\nNgày dạy: " + DATES_35[9], "90 phút\n(1 buổi)",
     "- Nhận biết đặc điểm loài cá mập, vai trò cân bằng sinh thái của sinh vật biển.\n- Lắp ráp mô hình Cá mập săn mồi với vây bơi linh hoạt và thân uốn lượn.\n- Nâng cao ý thức giữ gìn môi trường biển và bảo vệ các loài động vật quý hiếm."),
    ("Buổi 11", "Bài 11: CHÚ CUA CỨNG CÁP\nNgày dạy: " + DATES_35[10], "90 phút\n(1 buổi)",
     "- Khám phá cơ chế bò ngang đặc biệt và cấu tạo đôi càng kẹp bảo vệ của loài cua.\n- Lắp ráp mô hình Chú cua biển với cơ cấu đòn bẩy điều khiển càng kẹp mở/đóng.\n- Rèn luyện kỹ năng kết nối các trục xoay và thanh giằng cơ học."),
    ("Buổi 12", "Bài 12: CON MỰC THÔNG MINH\nNgày dạy: " + DATES_35[11], "90 phút\n(1 buổi)",
     "- Tìm hiểu cấu tạo các xúc tu mềm và cơ chế phản lực giúp mực di chuyển nhanh trong nước.\n- Lắp ráp robot Con mực đại dương với các xúc tu cử động linh hoạt.\n- Phát triển tư duy hình học không gian và sự khéo léo của đôi bàn tay."),
    ("Buổi 13", "Bài 13: HOẠT ĐỘNG MÙA HÈ\nNgày dạy: " + DATES_35[12], "90 phút\n(1 buổi)",
     "- Nhận biết các hoạt động thể thao dưới nước, kỹ thuật chèo thuyền kayak và rèn luyện thể lực.\n- Lắp ráp mô hình Thuyền chèo với hệ thống hai mái chèo cơ học liên động.\n- Bồi dưỡng tinh thần thể thao năng động và khả năng phối hợp đồng đội."),
    ("Buổi 14", "Bài 14: BẮT ĐẦU CHUYẾN HÀNH TRÌNH CÙNG TÀU HOẢ\nNgày dạy: " + DATES_35[13], "90 phút\n(1 buổi)",
     "- Tìm hiểu lịch sử ngành đường sắt, cấu tạo đầu tàu kéo và các toa chở khách/chở hàng.\n- Lắp ráp mô hình Đoàn tàu hỏa với các khớp nối toa tàu di chuyển nhịp nhàng.\n- Nhận thức về các phương tiện giao thông khối lượng lớn thân thiện với môi trường."),
    ("Buổi 15", "Bài 15: PHÍA TRÊN BẦU TRỜI\nNgày dạy: " + DATES_35[14], "90 phút\n(1 buổi)",
     "- Nhận biết cấu tạo thân máy bay, cánh nâng khí động học và vai trò của cánh quạt đẩy.\n- Lắp ráp mô hình Máy bay cánh quạt hai tầng cánh thăng bằng trên không trung.\n- Nuôi dưỡng ước mơ chinh phục bầu trời và khám phá kỹ thuật hàng không."),
    ("Buổi 16", "Bài 16: VŨ TRỤ RỘNG LỚN\nNgày dạy: " + DATES_35[15], "90 phút\n(1 buổi)",
     "- Khám phá các hành tinh trong Hệ Mặt Trời và nhiệm vụ của tàu vũ trụ thám hiểm không gian.\n- Lắp ráp mô hình Tàu vũ trụ OLLO Initiate với cánh pin năng lượng mặt trời.\n- Khơi gợi niềm đam mê khoa học thiên văn và công nghệ khám phá không gian."),
    ("Buổi 17", "Ôn tập & Triển lãm Robot sáng tạo Học kì 1 (Đánh giá định kỳ 2)\nNgày dạy: " + DATES_35[16], "90 phút\n(1 buổi)",
     "- Hệ thống hóa toàn bộ kiến thức, linh kiện và kỹ thuật lắp ráp bộ OLLO Initiate HK1.\n- Thực hành sáng tạo tự do hoặc nâng cấp mô hình robot yêu thích.\n- ĐÁNH GIÁ ĐỊNH KỲ 2: Báo cáo thuyết trình sản phẩm cá nhân/nhóm và triển lãm robot."),
    
    # HK2: Chuyển tiếp lên OLLO Kinder (Lớp 1, 2)
    ("Buổi 18", "Bài 1: HÃY NẤU NHỮNG MÓN ĂN NGON!\nNgày dạy: " + DATES_35[17], "90 phút\n(1 buổi)",
     "- Làm quen bộ Kit OLLO KINDER; nhận diện linh kiện, hộp pin và động cơ Kinder.\n- Tìm hiểu dụng cụ nhà bếp, nắp vung đóng mở và chất liệu cách nhiệt của nồi nấu ăn.\n- Lắp ráp mô hình Chiếc nồi thần kỳ có nắp tự động nâng hạ."),
    ("Buổi 19", "Bài 2: HÀM RĂNG TRẮNG SÁNG\nNgày dạy: " + DATES_35[18], "90 phút\n(1 buổi)",
     "- Khám phá cấu tạo hàm răng, thói quen vệ sinh răng miệng đúng cách để ngừa sâu răng.\n- Lắp ráp mô hình Bàn chải thông minh với cơ cấu xoay tròn làm sạch mảng bám.\n- Hình thành nền nếp giữ gìn vệ sinh cá nhân và bảo vệ nụ cười khỏe đẹp."),
    ("Buổi 20", "Bài 3: SỰ PHẢN XẠ ÁNH SÁNG\nNgày dạy: " + DATES_35[19], "90 phút\n(1 buổi)",
     "- Tìm hiểu hiện tượng phản xạ ánh sáng, ứng dụng gương phản chiếu và cấu tạo máy ảnh.\n- Lắp ráp mô hình Máy ảnh du lịch OLLO Kinder có nút bấm chụp và ống kính cơ học.\n- Rèn luyện kỹ năng quan sát chi tiết và thẩm mỹ thị giác."),
    ("Buổi 21", "Bài 4: ĐÀN GÀ CON ĐÁNG YÊU\nNgày dạy: " + DATES_35[20], "90 phút\n(1 buổi)",
     "- Tìm hiểu vòng đời của loài gà, đặc điểm của gà con và tập tính tìm thức ăn.\n- Lắp ráp mô hình Gà con mổ thóc với cơ chế đòn bẩy gật đầu và vẫy cánh đồng thời.\n- Phát triển tư duy liên hệ giữa cơ học máy và cử động sinh học tự nhiên."),
    ("Buổi 22", "Bài 5: NHỮNG BẠN NHỎ LỄ PHÉP\nNgày dạy: " + DATES_35[21], "90 phút\n(1 buổi)",
     "- Tìm hiểu văn hóa chào hỏi lễ phép của các quốc gia và hành vi ứng xử lịch thiệp.\n- Lắp ráp mô hình Robot cúi chào với cơ cấu gập thân góc 90 độ nhịp nhàng.\n- Giáo dục phẩm chất tôn trọng người lớn, lễ phép với thầy cô và thân thiện với bạn bè."),
    ("Buổi 23", "Bài 6: ĐỘNG VẬT THÂN MỀM\nNgày dạy: " + DATES_35[22], "90 phút\n(1 buổi)",
     "- Tìm hiểu các loài động vật thân mềm nâng cao, cấu tạo lớp vỏ bảo vệ và chất nhờn.\n- Lắp ráp mô hình Ốc sên Kinder với hệ truyền động bánh răng tạo bước bò lượn sóng.\n- Rèn luyện kỹ năng phân tích cơ cấu truyền động bánh răng trục tròn."),
    ("Buổi 24", "Bài 7: CHÚ THỎ TAI DÀI NHẢY NHÓT\nNgày dạy: " + DATES_35[23], "90 phút\n(1 buổi)",
     "- Tìm hiểu đặc điểm cấu tạo đôi chân sau dài khỏe giúp thỏ bật nhảy nhanh thoát hiểm.\n- Lắp ráp mô hình Chú thỏ trắng với cơ cấu đòn bẩy chân lò xo tạo lực nhún nhảy.\n- Khám phá nguyên lý chuyển đổi thế năng đàn hồi thành động năng bật nhảy."),
    ("Buổi 25", "Bài 8: HIỆP SĨ ĐẤU KIẾM\nNgày dạy: " + DATES_35[24], "90 phút\n(1 buổi)",
     "- Tìm hiểu tinh thần thượng võ, trang phục áo giáp và luật thi đấu môn đấu kiếm thể thao.\n- Lắp ráp mô hình Hiệp sĩ cầm kiếm với cơ chế tay đòn vung kiếm dứt khoát.\n- Nâng cao tính kỷ luật, tinh thần dũng cảm và lòng chính trực."),
    ("Buổi 26", "Bài 9: MÔI TRƯỜNG BIỂN & CHÚ CÁ BƠI LỘI\nNgày dạy: " + DATES_35[25], "90 phút\n(1 buổi)",
     "- Khám phá hệ sinh thái san hô đại dương và cấu tạo các vây điều hướng của loài cá.\n- Lắp ráp mô hình Chú cá nhiệt đới OLLO Kinder quẫy đuôi đẩy thân lướt đi.\n- Ý thức bảo vệ rạn san hô và giảm thiểu rác thải nhựa đại dương."),
    ("Buổi 27", "Bài 10: CÙNG CÂU CÁ NÀO!\nNgày dạy: " + DATES_35[26], "90 phút\n(1 buổi)",
     "- Tìm hiểu môn thể thao câu cá dã ngoại, cấu tạo cần câu và nguyên tắc an toàn bên mặt nước.\n- Lắp ráp mô hình Cần câu cá với cụm ròng rọc cơ học cuốn thả dây cáp có nam châm.\n- Vận dụng nguyên lý đòn bẩy và ròng rọc cố định nâng vật nặng dễ dàng."),
    ("Buổi 28", "Bài 11: VẬN CHUYỂN ĐỒ VẬT (Đánh giá định kỳ 3)\nNgày dạy: " + DATES_35[27], "90 phút\n(1 buổi)",
     "- Tìm hiểu các dòng xe nâng chuyên dụng trong nhà kho logistics và nguyên tắc an toàn hàng hóa.\n- Lắp ráp Xe nâng hàng OLLO Kinder với giá đỡ nâng hạ theo thanh ray trượt thẳng đứng.\n- ĐÁNH GIÁ ĐỊNH KỲ 3: Kiểm tra thao tác lắp ghép khung cơ khí Kinder và khả năng chịu tải của xe."),
    ("Buổi 29", "Bài 12: SỨC MẠNH CỦA MÁY ỦI\nNgày dạy: " + DATES_35[28], "90 phút\n(1 buổi)",
     "- Khám phá máy ủi bánh xích công trường, cơ chế lưỡi ben thép gạt đất đá san phẳng mặt bằng.\n- Lắp ráp mô hình Máy ủi dũng mãnh với hệ dẫn động đòn nâng lưỡi ben linh hoạt.\n- Nâng cao ý thức an toàn công trình xây dựng và bảo hộ lao động."),
    ("Buổi 30", "Bài 13: MÁY BẮN ĐÁ THẦN TỐC\nNgày dạy: " + DATES_35[29], "90 phút\n(1 buổi)",
     "- Tìm hiểu lịch sử các loại máy cơ khí cổ đại, nguyên lý tích lũy thế năng và đòn bẩy cánh tay dài.\n- Lắp ráp Máy bắn đá OLLO Kinder với đối trọng và chốt kích hoạt phóng bóng xa.\n- Thử nghiệm độ chính xác của tầm bắn qua việc điều chỉnh góc phóng."),
    ("Buổi 31", "Bài 14: CUNG TÊN ĐÀN HỒI\nNgày dạy: " + DATES_35[30], "90 phút\n(1 buổi)",
     "- Tìm hiểu môn bắn cung Olympic, lực đàn hồi của cánh cung và quỹ đạo bay hình parabol.\n- Lắp ráp mô hình Cung tên thể thao an toàn có ống định hướng mũi tên.\n- Rèn luyện tính kiên nhẫn, khả năng tập trung cao độ và căn chỉnh mục tiêu."),
    ("Buổi 32", "Bài 15: XE BUÝT HAI TẦNG THÀNH PHỐ\nNgày dạy: " + DATES_35[31], "90 phút\n(1 buổi)",
     "- Nhận biết đặc trưng xe buýt hai tầng phục vụ du lịch ngắm cảnh tại các đô thị hiện đại.\n- Lắp ráp mô hình Xe buýt hai tầng với khung gầm vững chắc, cầu thang xoắn và cửa tự động.\n- Bồi dưỡng ý thức văn minh giao thông công cộng và bảo vệ bầu không khí xanh."),
    ("Buổi 33", "Bài 16: CÁ SẤU THẬT NGẦU!\nNgày dạy: " + DATES_35[32], "90 phút\n(1 buổi)",
     "- Khám phá tập tính ẩn mình phục kích và sức mạnh cú đớp mồi của loài cá sấu đầm lầy.\n- Lắp ráp robot Cá sấu cơ bắp với hàm răng liên động đóng mở phát ra tiếng cơ khí mạnh mẽ.\n- Phát triển năng lực phân tích liên hệ sinh học cơ thể sống với thiết kế kỹ thuật."),
    ("Buổi 34", "Ôn tập & Thiết kế dự án Robot cuối năm (Đánh giá định kỳ 4)\nNgày dạy: " + DATES_35[33], "90 phút\n(1 buổi)",
     "- Tổng kết toàn bộ kỹ năng lắp ghép cơ bản (Initiate) và nâng cao (Kinder) trong cả năm học.\n- Thực hiện dự án chế tạo robot tự chọn theo nhóm: thiết kế mô hình, thử nghiệm và tối ưu.\n- ĐÁNH GIÁ ĐỊNH KỲ 4: Đánh giá dự án robot cuối năm, kỹ năng giải quyết sự cố và tinh thần đồng đội."),
    ("Buổi 35", "Ngày hội Robotics UNIGO & Bế mạc CLB\nNgày dạy: " + DATES_35[34], "90 phút\n(1 buổi)",
     "- Triển lãm Festival Robotics UNIGO: Trình diễn các sản phẩm robot xuất sắc của học sinh Lớp 1-2.\n- Tham gia thử thách vượt sa bàn robot và giao lưu chia sẻ cảm xúc học tập suốt năm.\n- Trao giấy chứng nhận hoàn thành chương trình CLB Robotics UNIGO và khen thưởng cá nhân xuất sắc.")
]

# Dữ liệu 35 buổi Nhóm Lớp 3 - 7 (HK1 Kinder -> HK2 Excel)
DATA_LOP_3_7 = [
    # HK1: Kinder nâng cao
    ("Buổi 1", "Bài 1: CỖ MÁY LÀM SẠCH ĐƯỜNG PHỐ!\nNgày dạy: " + DATES_35[0], "90 phút\n(1 buổi)",
     "- Khám phá quy trình vệ sinh môi trường đô thị hiện đại và cấu tạo robot xe gom rác tự động.\n- Kỹ năng nhận diện và lắp ráp hệ thống chổi quét xoay, thùng chứa rác và khung xe OLLO Kinder.\n- Hình thành ý thức bảo vệ môi trường, phân loại rác thải tại nguồn và văn minh học đường."),
    ("Buổi 2", "Bài 2: VÌ SAO CHIẾC XE NÀY KHÔNG CÓ MÁI?\nNgày dạy: " + DATES_35[1], "90 phút\n(1 buổi)",
     "- Tìm hiểu nguyên lý khí động học, trọng tâm xe thể thao mui trần giúp xe đạt tốc độ cao và ổn định.\n- Lắp ráp hoàn thiện Xe đua mui trần với hệ thống trục lái trước chuyển hướng linh hoạt.\n- Nâng cao tư duy phân tích cấu trúc cơ học và kỹ năng cân chỉnh trục bánh xe."),
    ("Buổi 3", "Bài 3: CỖ MÁY BAY GIỮA KHÔNG TRUNG\nNgày dạy: " + DATES_35[2], "90 phút\n(1 buổi)",
     "- Tìm hiểu nguyên lý cánh quạt kép đồng trục triệt tiêu mômen xoắn và tạo lực nâng trực thăng.\n- Lắp ráp mô hình Trực thăng trinh sát OLLO Kinder với hệ truyền động bánh răng cánh quạt.\n- Rèn luyện kỹ năng lắp ráp cơ cấu truyền động bánh răng nón và bánh răng trụ."),
    ("Buổi 4", "Bài 4: CHIẾC XE ĐẦU ĐỜI CỦA BÉ\nNgày dạy: " + DATES_35[3], "90 phút\n(1 buổi)",
     "- Khám phá cấu trúc tam giác cân bằng hình học và phân bổ trọng lượng trong thiết kế xe tập đi.\n- Lắp ráp mô hình Xe đẩy cân bằng đa hướng có tay vịn chống lật an toàn.\n- Vận dụng kiến thức vật lý về điểm tựa và diện tích mặt chân đế vào chế tạo mô hình."),
    ("Buổi 5", "Bài 5: CHÚ CHIM ĐỘI MŨ ĐỎ\nNgày dạy: " + DATES_35[4], "90 phút\n(1 buổi)",
     "- Tìm hiểu đặc tính gõ cây bắt sâu bảo vệ rừng và cấu tạo giảm chấn hộp sọ chim gõ kiến.\n- Lắp ráp cơ cấu thanh truyền biến chuyển động quay thành chuyển động gõ mỏ tịnh tiến liên tục.\n- Phát triển tư duy phỏng sinh học (Biomimicry) trong sáng tạo công nghệ robot."),
    ("Buổi 6", "Bài 6: BẬC THẦY BAY LƯỢN\nNgày dạy: " + DATES_35[5], "90 phút\n(1 buổi)",
     "- Khám phá cấu tạo sải cánh khí động học và kỹ thuật chao lượn tận dụng luồng nhiệt của đại bàng.\n- Lắp ráp mô hình Đại bàng săn mồi với cơ cấu khớp cánh đa tầng vỗ cánh uy lực.\n- Rèn luyện tính kiên nhẫn trong tinh chỉnh các khớp nối đa liên kết."),
    ("Buổi 7", "Bài 7: NHỮNG NGƯỜI BẠN ĐÁNG YÊU\nNgày dạy: " + DATES_35[6], "90 phút\n(1 buổi)",
     "- Tìm hiểu sức bền dẻo dai của giống chó tuyết Husky và cấu tạo xe trượt tuyết vùng cực Bắc.\n- Lắp ráp mô hình Chó kéo xe chở hàng với hệ dây kéo cân bằng lực hai bên sườn.\n- Nâng cao kỹ năng làm việc nhóm, phân công vai trò lắp ghép nhịp nhàng."),
    ("Buổi 8", "Bài 8: CÙNG KHÁM PHÁ ROBOT PIKACHU\nNgày dạy: " + DATES_35[7], "90 phút\n(1 buổi)",
     "- Phân tích nguyên lý trọng tâm lệch tâm (Cam lệch) tạo nên dao động nhún nhảy lắc lư ngộ nghĩnh.\n- Lắp ráp mô hình Robot Pikachu chuyển động lắc lư theo nhịp bước đi.\n- Vận dụng kiến thức về chuyển động dao động và sáng tạo nâng cấp chi tiết tai, đuôi."),
    ("Buổi 9", "Bài 9: KHÁM PHÁ SINH VẬT ĐUÔI CONG ĐẶC BIỆT (Đánh giá định kỳ 1)\nNgày dạy: " + DATES_35[8], "90 phút\n(1 buổi)",
     "- Khám phá cấu tạo đốt xương đuôi cong dạng móc bám và cách bơi thẳng đứng của cá ngựa.\n- Lắp ráp mô hình Cá ngựa đuôi cong với cơ cấu vây lưng rung tần số cao.\n- ĐÁNH GIÁ ĐỊNH KỲ 1: Đánh giá kỹ năng thiết kế cơ cấu máy Kinder và giải bài toán kỹ thuật nhóm."),
    ("Buổi 10", "Bài 10: TÌM HIỂU CƠ CHẾ CHUYỂN ĐỘNG TRÊN MẶT NƯỚC\nNgày dạy: " + DATES_35[9], "90 phút\n(1 buổi)",
     "- Tìm hiểu lực đẩy Ac-si-met, lực cản dòng nước và nguyên lý đẩy của cánh quạt chân vịt Cano.\n- Lắp ráp mô hình Tàu cao tốc Cano với bánh lái đuôi chuyển hướng dòng nước chuẩn xác.\n- Thử nghiệm trên thảm thực hành mô phỏng luồng di chuyển của tàu thuyền."),
    ("Buổi 11", "Bài 11: AI LÀ VẬN ĐỘNG VIÊN BƠI LỘI DƯỚI ĐẠI DƯƠNG\nNgày dạy: " + DATES_35[10], "90 phút\n(1 buổi)",
     "- Khám phá lớp mỡ dày giữ ấm và đôi chân vây dẻo dai giúp hải cẩu bơi lặn siêu hạng.\n- Lắp ráp mô hình Hải cẩu rẽ sóng với khớp cổ và vây trước uốn lượn nhịp nhàng.\n- Bồi dưỡng tình yêu thế giới sinh vật biển và trách nhiệm bảo tồn thiên nhiên hoang dã."),
    ("Buổi 12", "Bài 12: CÁNH TAY KHỔNG LỒ BIẾT NÂNG VẬT NẶNG\nNgày dạy: " + DATES_35[11], "90 phút\n(1 buổi)",
     "- Phân tích nguyên lý cánh tay đòn tải trọng lớn, đối trọng thăng bằng và đĩa quay 360 độ của cần cẩu.\n- Lắp ráp mô hình Cần cẩu tháp xây dựng OLLO Kinder với dây tời ròng rọc kép nâng hạ khối hàng.\n- Rèn luyện kỹ năng tính toán tải trọng giới hạn chống lật mô hình."),
    ("Buổi 13", "Bài 13: AI ĐANG DẤU MÌNH DƯỚI CÁT?\nNgày dạy: " + DATES_35[12], "90 phút\n(1 buổi)",
     "- Khám phá cơ chế thích nghi đổi màu ngụy trang và thân hình dẹt nằm sát đáy biển của cá bơn.\n- Lắp ráp mô hình Cá bơn ngụy trang với cơ cấu vây diềm gợn sóng đẩy cát phủ thân.\n- Nâng cao tư duy thiết kế mô phỏng các cơ chế ngụy trang thông minh trong kỹ thuật quân sự."),
    ("Buổi 14", "Bài 14: MẬT MÃ DISNEY\nNgày dạy: " + DATES_35[13], "90 phút\n(1 buổi)",
     "- Phân tích tỷ lệ hình học tạo hình nhân vật hoạt hình chú chuột Mickey nổi tiếng toàn cầu.\n- Lắp ráp mô hình Chú chuột Mickey thông minh với các khớp tay cầm nắm và bước chân linh hoạt.\n- Kết hợp nghệ thuật tạo hình (Art) với tư duy kỹ thuật lắp ráp STEM."),
    ("Buổi 15", "Bài 15: BÍ MẬT CỦA BỘ LÔNG TRẮNG!\nNgày dạy: " + DATES_35[14], "90 phút\n(1 buổi)",
     "- Tìm hiểu công nghệ dệt len tự nhiên, cấu trúc bộ lông cách nhiệt vượt trội của loài cừu tuyết.\n- Lắp ráp mô hình Chú cừu thông minh với thân hình bo tròn khí động học và bánh xe giảm chấn.\n- Rèn luyện kỹ năng kết nối các tấm ốp vỏ tạo độ chắc chắn cho khung robot."),
    ("Buổi 16", "Bài 16: ÂM THANH TỪ THẾ GIỚI HOẠT HÌNH\nNgày dạy: " + DATES_35[15], "90 phút\n(1 buổi)",
     "- Tìm hiểu nguyên lý phát thanh cơ học: gai gõ trên trống quay cọ xát vào thanh thép rung tạo nốt nhạc.\n- Lắp ráp mô hình Hộp âm thanh cơ học OLLO Kinder tạo chuỗi tiết tấu vui nhộn.\n- Tích hợp liên môn Âm nhạc và Kỹ thuật công nghệ phát triển cảm thụ nghệ thuật."),
    ("Buổi 17", "Ôn tập & Triển lãm Robot sáng tạo Học kì 1 (Đánh giá định kỳ 2)\nNgày dạy: " + DATES_35[16], "90 phút\n(1 buổi)",
     "- Hệ thống hóa toàn diện các cơ chế cơ khí Kinder: đòn bẩy, ròng rọc, bánh răng, trục lệch tâm, thanh truyền.\n- Thiết kế và chế tạo dự án robot giải quyết bài toán thực tế theo nhóm học sinh.\n- ĐÁNH GIÁ ĐỊNH KỲ 2: Trưng bày dự án, chấm điểm kỹ năng vận hành, thẩm mỹ kỹ thuật và thuyết trình."),
    
    # HK2: Đột phá công nghệ cùng OLLO EXCEL (Lớp 3 - 7)
    ("Buổi 18", "Bài 1: ĐỘNG CƠ LÀ GÌ? (OLLO EXCEL)\nNgày dạy: " + DATES_35[17], "90 phút\n(1 buổi)",
     "- Làm quen bộ thiết bị cao cấp OLLO EXCEL; tìm hiểu cấu tạo động cơ servo thông minh Dynamixel.\n- Nắm vững quy tắc cắm chốt chuyên dụng, căn chỉnh vạch chuẩn góc 0 độ của động cơ Dynamixel.\n- Thực hành lắp ráp Xe nâng hàng tự động vận hành bằng động cơ điện Dynamixel công suất lớn."),
    ("Buổi 19", "Bài 2: ROBOT CẦM NẮM VÀ CỐ ĐỊNH VẬT THỂ\nNgày dạy: " + DATES_35[18], "90 phút\n(1 buổi)",
     "- Khám phá cơ cấu ngàm kẹp cơ khí tự động, góc kẹp tối ưu và lực ma sát giữ đồ vật không trượt.\n- Lắp ráp mô hình Robot gắp hàng tự động có cánh tay kẹp điều khiển bởi động cơ Dynamixel.\n- Thử nghiệm gắp và vận chuyển các vật thể kích thước, khối lượng khác nhau vào ô mục tiêu."),
    ("Buổi 20", "Bài 3: ROBOT NHẬN BIẾT ÂM THANH\nNgày dạy: " + DATES_35[19], "90 phút\n(1 buổi)",
     "- Tìm hiểu nguyên lý cảm biến âm thanh (Sound Sensor / Micro) thu nhận sóng âm thanh biến đổi thành tín hiệu.\n- Lắp ráp Robot vỗ tay thông minh có khả năng di chuyển hoặc dừng lại khi nghe tiếng vỗ tay của học sinh.\n- Lập trình cấu hình độ nhạy cảm biến âm thanh phù hợp với độ ồn môi trường lớp học."),
    ("Buổi 21", "Bài 4: BỘ ĐIỀU KHIỂN TỪ XA RC-100\nNgày dạy: " + DATES_35[20], "90 phút\n(1 buổi)",
     "- Khám phá công nghệ truyền tín hiệu không dây hồng ngoại (IR) giữa tay cầm RC-100 và bộ điều khiển trung tâm.\n- Lắp ráp mô hình Pháo cao xạ tầm xa điều khiển từ xa quay nòng 360 độ và kích hoạt bắn đạn nhựa an toàn.\n- Rèn luyện kỹ năng điều khiển chính xác robot vượt chướng ngại vật và bắn hạ bia mục tiêu."),
    ("Buổi 22", "Bài 5: DYNAMIXEL HOẠT ĐỘNG NHƯ THẾ NÀO?\nNgày dạy: " + DATES_35[21], "90 phút\n(1 buổi)",
     "- Phân tích sâu nguyên lý hoạt động của động cơ bước / servo Dynamixel: mạch điều khiển tích hợp và hộp số giảm tốc.\n- Lắp ráp mô hình Robot khuếch đại chuyển động kiểm chứng sự thay đổi tốc độ và mômen xoắn.\n- Tính toán tỷ số truyền động bánh răng và quan sát thực tế sự đánh đổi giữa tốc độ và lực nâng."),
    ("Buổi 23", "Bài 6: PHƯƠNG TIỆN GIAO THÔNG QUA CÁC THỜI ĐẠI\nNgày dạy: " + DATES_35[22], "90 phút\n(1 buổi)",
     "- Tìm hiểu lịch sử tiến hóa phương tiện: từ cỗ xe ngựa thời cổ đến xe đua siêu tốc hiện đại.\n- Lắp ráp mô hình Cỗ xe ngựa đua tự động OLLO Excel với cơ cấu 4 chân phi nước đại nhịp nhàng.\n- Tổ chức giải đua xe địa hình tính thời gian giữa các đội trong câu lạc bộ."),
    ("Buổi 24", "Bài 7: CƠ CHẾ TĂNG GIẢM CHIỀU DÀI TỰ ĐỘNG\nNgày dạy: " + DATES_35[23], "90 phút\n(1 buổi)",
     "- Khám phá cơ cấu xếp hình quả trám (Lazy Tong / Scissors Mechanism) ứng dụng mở rộng tầm với trong robot gỡ bom.\n- Lắp ráp Cánh tay robot co giãn tự động vươn dài gắp vật thể ở cự ly xa hiểm trở.\n- Nâng cao tư duy thiết kế giải pháp robot phục vụ an ninh quốc phòng và cứu hộ cứu nạn."),
    ("Buổi 25", "Bài 8: ROBOT NHẬN BIẾT VẬT THỂ BẰNG CÁCH NÀO?\nNgày dạy: " + DATES_35[24], "90 phút\n(1 buổi)",
     "- Tìm hiểu nguyên lý phát xạ và thu nhận tia hồng ngoại (IR Sensor) đo khoảng cách chướng ngại vật.\n- Lắp ráp Xe robot tự hành thông minh tự động phát hiện bức tường phía trước và tự quay đầu tránh va chạm.\n- Vận dụng tư duy thuật toán rẽ nhánh điều kiện: IF (phát hiện vật cản) THEN (rẽ trái/lùi lại)."),
    ("Buổi 26", "Bài 9: SỐ NGẪU NHIÊN VÀ THUẬT TOÁN ĐƠN GIẢN\nNgày dạy: " + DATES_35[25], "90 phút\n(1 buổi)",
     "- Khám phá khái niệm số ngẫu nhiên (Random number) trong khoa học máy tính và hành vi sinh vật.\n- Lắp ráp Robot thám hiểm mê cung có hành vi ngẫu nhiên, mô phỏng phản xạ tự nhiên của động vật hoang dã.\n- Rèn luyện năng lực tư duy máy tính (Computational Thinking) và giải thuật logic."),
    ("Buổi 27", "Bài 10: ROBOT HÚT BỤI THÔNG MINH\nNgày dạy: " + DATES_35[26], "90 phút\n(1 buổi)",
     "- Phân tích công nghệ robot dịch vụ gia đình: kết hợp động cơ quét bụi, quạt hút chân không và cảm biến va chạm.\n- Lắp ráp mô hình Robot hút bụi tự động OLLO Excel tự di chuyển làm sạch sàn nhà không đụng chân bàn.\n- Thử nghiệm hiệu suất quét dọn mô phỏng và tối ưu hóa diện tích bao phủ mặt sàn."),
    ("Buổi 28", "Bài 11: LỰC HẤP DẪN TRONG ROBOTICS (Đánh giá định kỳ 3)\nNgày dạy: " + DATES_35[27], "90 phút\n(1 buổi)",
     "- Khám phá định luật vạn vật hấp dẫn, trọng lực tác dụng lên các bộ phận chuyển động và phản lực tiếp xúc.\n- Lắp ráp Robot đấu sĩ cưỡi ngựa Rodeo giữ thăng bằng trên lưng ngựa dao động mạnh.\n- ĐÁNH GIÁ ĐỊNH KỲ 3: Kiểm tra kỹ năng tích hợp cảm biến, động cơ Dynamixel và điều khiển không dây."),
    ("Buổi 29", "Bài 12: DOMINO TRONG ROBOTICS\nNgày dạy: " + DATES_35[28], "90 phút\n(1 buổi)",
     "- Khám phá cơ cấu cấp phôi tự động (Dispenser) và khoảng cách thả tối ưu tạo chuỗi phản ứng dây chuyền Domino.\n- Lắp ráp Xe robot tự động rải quân cờ Domino cách đều nhau trên sàn khi di chuyển thẳng.\n- Thực hiện thử thách xếp trận địa Domino dài nhất và kích hoạt đổ chuỗi hoàn hảo."),
    ("Buổi 30", "Bài 13: ROBOT DỌN RÁC THÔNG MINH\nNgày dạy: " + DATES_35[29], "90 phút\n(1 buổi)",
     "- Tìm hiểu hệ thống dẫn động 4 bánh độc lập (4WD), vi sai chuyển hướng và cơ cấu chổi cuốn rác xoay ngược chiều.\n- Lắp ráp Robot dọn rác công nghệ cao điều khiển từ xa bằng RC-100 vượt chướng ngại vật thu gom hạt nhựa.\n- Rèn luyện kỹ năng kết hợp nhuần nhuyễn giữa thao tác điều khiển tay bấm và khả năng phán đoán địa hình."),
    ("Buổi 31", "Bài 14: ROBOT CHIẾN BINH VÀ VŨ KHÍ TỰ VỆ\nNgày dạy: " + DATES_35[30], "90 phút\n(1 buổi)",
     "- Khám phá các chiến thuật chế tạo robot thi đấu: hạ thấp trọng tâm, tấm ốp mũi nêm lật đối thủ và cánh tay đập.\n- Lắp ráp Robot chiến binh giác đấu OLLO Excel được gia cố khung sườn kim loại và bánh xe bám cao su.\n- Nâng cao tư duy chiến thuật, phản xạ nhanh và sự tôn trọng đối thủ trong thi đấu thể thao điện tử/robot."),
    ("Buổi 32", "Bài 15: ĐẤU TRƯỜNG ROBOT ĐÁ BÓNG / THỂ THAO\nNgày dạy: " + DATES_35[31], "90 phút\n(1 buổi)",
     "- Tìm hiểu luật thi đấu Robot Soccer quốc tế, cấu tạo cơ cấu dẫn bóng và cánh gạt sút bóng uy lực.\n- Lắp ráp và tối ưu hóa Robot cầu thủ bóng đá điều khiển bằng remote RC-100 hai kênh độc lập.\n- Tổ chức giải bóng đá Robot mini 2 vs 2 rèn luyện khả năng phối hợp chiến thuật nhóm."),
    ("Buổi 33", "Bài 16: ĐẤU TRƯỜNG ROBOT SUMO & GIẢI CỨU MỤC TIÊU\nNgày dạy: " + DATES_35[32], "90 phút\n(1 buổi)",
     "- Nắm vững luật thi đấu Sumo Robot: hất đối thủ ra khỏi võ đài tròn Dohyo và giữ thăng bằng an toàn.\n- Chế tạo Robot Sumo công suất cực đại: tối ưu hóa lực kéo bánh xe, độ bám mặt sân và lưỡi xúc càn quét.\n- Tranh tài giải đấu Sumo Robot vòng bảng và bán kết nghẹt thở giữa các đấu thủ xuất sắc."),
    ("Buổi 34", "Ôn tập & Thiết kế dự án Robot cuối năm (Đánh giá định kỳ 4)\nNgày dạy: " + DATES_35[33], "90 phút\n(1 buổi)",
     "- Tổng kết toàn diện hành trình làm chủ công nghệ OLLO Excel: cơ khí, điện tử, cảm biến và điều khiển tự động.\n- Thiết kế dự án sáng tạo cuối năm theo đề tài: Thành phố thông minh, Nông nghiệp công nghệ cao, Cứu hộ khẩn cấp.\n- ĐÁNH GIÁ ĐỊNH KỲ 4: Thẩm định hồ sơ thiết kế kỹ thuật, sản phẩm thực tế, video vận hành và thuyết trình bảo vệ."),
    ("Buổi 35", "Ngày hội Robotics UNIGO & Bế mạc CLB\nNgày dạy: " + DATES_35[34], "90 phút\n(1 buổi)",
     "- Chung kết Đấu trường ROBOTICS UNIGO CUP 2026 - 2027: Tranh tài Sumo, Bóng đá robot và Thuyết trình dự án sáng tạo.\n- Lễ bế mạc tổng kết một năm học tập, nghiên cứu và bùng nổ đam mê công nghệ cùng CLB Robotics.\n- Trao cúp vô địch, huy chương danh dự, chứng nhận tốt nghiệp CLB và quà lưu niệm cho các thành viên.")
]

def generate_khdh_docx():
    print("Reading template:", DOCX_PATH)
    doc = docx.Document(DOCX_PATH)

    # Backup header paragraphs
    header_runs = []
    sec = doc.sections[0]
    for hp in sec.header.paragraphs:
        header_runs.append([r.text for r in hp.runs])

    # Remove all body children except sectPr
    body = doc.element.body
    children_to_remove = []
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag != 'sectPr':
            children_to_remove.append(child)
    
    for c in children_to_remove:
        body.remove(c)
    
    print(f"Removed {len(children_to_remove)} old body elements. Retained sectPr.")

    # 1. Header Table (Table 0: Quốc hiệu & Tên trường)
    tbl0 = doc.add_table(rows=2, cols=2)
    tbl0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl0)

    c00 = tbl0.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p00.add_run("TRƯỜNG TIỂU HỌC & THCS UNIGO\nCÂU LẠC BỘ ROBOTICS UNIGO")
    r.bold = True
    r.font.size = Pt(12)

    c01 = tbl0.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc")
    r.bold = True
    r.font.size = Pt(12)

    c10 = tbl0.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p10.add_run("---------------")
    r.font.size = Pt(11)

    c11 = tbl0.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p11.add_run("---------------")
    r.font.size = Pt(11)

    # 2. Tiêu đề Kế hoạch
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run("\nKẾ HOẠCH TỔ CHỨC DẠY HỌC\nCÂU LẠC BỘ ROBOTICS – NĂM HỌC 2026 - 2027")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0, 32, 96)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_sub.add_run("(Lộ trình tích hợp phân hóa 3 bộ thiết bị: OLLO Initiate, OLLO Kinder và OLLO Excel Level 1)")
    r.italic = True
    r.font.size = Pt(12)

    # I. ĐẶC ĐIỂM TÌNH HÌNH
    p = doc.add_paragraph()
    r = p.add_run("I. ĐẶC ĐIỂM TÌNH HÌNH")
    r.bold = True
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    r = p.add_run("1. Quy mô, Đối tượng tham gia và Cơ cấu phân hệ:")
    r.bold = True
    
    doc.add_paragraph("- Tên câu lạc bộ: Câu lạc bộ Robotics UNIGO.")
    doc.add_paragraph("- Mục tiêu đối tượng: Học sinh từ Lớp 1 đến Lớp 7 có niềm say mê khoa học chế tạo, sáng tạo công nghệ và tư duy kỹ thuật STEM.")
    doc.add_paragraph("- Cơ cấu phân hệ học tập (Phù hợp độ tuổi và sự phát triển nhận thức):")
    
    p = doc.add_paragraph()
    r = p.add_run("  + Phân hệ 1 (Dành cho học sinh Khối 1 - 2): ")
    r.bold = True
    p.add_run("Học kỳ 1 học bộ thiết bị ")
    r = p.add_run("OLLO Initiate")
    r.bold = True
    p.add_run(" (làm quen với linh kiện, khớp nối, chuyển động cơ học cơ bản, mô phỏng thế giới xung quanh). Sau khi hoàn thành nền tảng Initiate, sang Học kỳ 2 các em được chuyển tiếp lên học bộ thiết bị ")
    r = p.add_run("OLLO Kinder")
    r.bold = True
    p.add_run(" (khám phá các cơ cấu chuyển động phức tạp hơn, máy công trình và động vật).")

    p = doc.add_paragraph()
    r = p.add_run("  + Phân hệ 2 (Dành cho học sinh Khối 3 - 7): ")
    r.bold = True
    p.add_run("Học kỳ 1 học bộ thiết bị ")
    r = p.add_run("OLLO Kinder")
    r.bold = True
    p.add_run(" (cơ cấu truyền động cơ khí, đòn bẩy, ròng rọc, bánh răng, trục lệch tâm). Sau khi hoàn thành vững vàng Kinder, sang Học kỳ 2 các em được bứt phá nâng cấp lên học bộ thiết bị công nghệ cao ")
    r = p.add_run("OLLO Excel Level 1")
    r.bold = True
    p.add_run(" (tiếp cận động cơ servo thông minh Dynamixel, cảm biến âm thanh, cảm biến hồng ngoại, bộ điều khiển từ xa RC-100 và robot thực hiện nhiệm vụ).")

    doc.add_paragraph("- Thời gian sinh hoạt CLB: Sáng Thứ 7 hàng tuần, từ 09h45 đến 11h15 (Thời lượng: 90 phút/buổi).")
    doc.add_paragraph("- Tổng thời lượng chương trình: 35 buổi/năm học (Học kì 1: 17 buổi; Học kì 2: 18 buổi).")
    doc.add_paragraph("- Thời gian thực hiện: Bắt đầu từ ngày 12/09/2026 đến ngày 29/05/2027.")

    p = doc.add_paragraph()
    r = p.add_run("2. Tình hình đội ngũ giáo viên phụ trách:")
    r.bold = True
    doc.add_paragraph("- Giáo viên phụ trách chuyên môn: Đậu Đình Nguyên (Giáo viên phụ trách môn Robotics & Tin học).")
    doc.add_paragraph("- Giáo viên/Trợ giảng hỗ trợ: Cán bộ/giáo viên hỗ trợ theo dõi nền nếp, hỗ trợ thao tác và bảo đảm an toàn thực hành cho học sinh.")

    p = doc.add_paragraph()
    r = p.add_run("3. Cơ sở vật chất và thiết bị dạy học:")
    r.bold = True
    doc.add_paragraph("- Phòng chức năng Robotics UNIGO đạt chuẩn quốc tế, trang bị máy chiếu, màn hình tương tác thông minh, hệ thống bàn ghế làm việc nhóm và thảm sa bàn thử nghiệm robot.")
    doc.add_paragraph("- Thiết bị học tập đồng bộ: 15 bộ Kit OLLO Initiate, 20 bộ Kit OLLO Kinder (Tập 1 & Tập 2), 15 bộ Kit OLLO Excel Level 1 (kèm động cơ Dynamixel, cảm biến âm thanh, cảm biến hồng ngoại, bộ thu phát và tay cầm RC-100).")

    # II. MỤC TIÊU CỦA CÂU LẠC BỘ
    p = doc.add_paragraph()
    r = p.add_run("II. MỤC TIÊU CỦA CÂU LẠC BỘ")
    r.bold = True

    p = doc.add_paragraph()
    r = p.add_run("1. Về kiến thức và tư duy STEM:")
    r.bold = True
    doc.add_paragraph("- Giúp học sinh khám phá sâu sắc thế giới tự nhiên và xã hội xung quanh thông qua mô hình robot trực quan (phương tiện giao thông, máy xây dựng, sinh vật biển, động vật hoang dã...).")
    doc.add_paragraph("- Hiểu và vận dụng kiến thức tích hợp Khoa học (Science), Công nghệ (Technology), Kỹ thuật (Engineering), Nghệ thuật (Art) và Toán học (Math).")
    doc.add_paragraph("- Nhận diện cấu tạo và nguyên lý của các bộ phận cơ khí: trục xoay, bánh xe, cánh tay nâng, bánh răng, đòn bẩy, ròng rọc; tiến tới làm chủ động cơ điện thông minh và các loại cảm biến.")

    p = doc.add_paragraph()
    r = p.add_run("2. Về năng lực:")
    r.bold = True
    p = doc.add_paragraph()
    r = p.add_run("  2.1. Năng lực đặc thù Robotics (Nhận thức, Giao tiếp, Sử dụng, Đánh giá, Thiết kế công nghệ):")
    r.italic = True
    doc.add_paragraph("    + Rèn luyện kỹ năng quan sát bản vẽ kỹ thuật 2D/3D, phân loại linh kiện và thao tác lắp ráp chính xác, an toàn.")
    doc.add_paragraph("    + Tự phát hiện và khắc phục các lỗi cơ học; sáng tạo cải tiến mô hình theo ý tưởng riêng.")
    doc.add_paragraph("    + Năng lực giao tiếp, thuyết trình tự tin giới thiệu sản phẩm robot trước đám đông.")
    
    p = doc.add_paragraph()
    r = p.add_run("  2.2. Năng lực số (Theo Thông tư 02/2025/TT-BGDĐT – CV 3456/BGDĐT):")
    r.italic = True
    doc.add_paragraph("    + Miền I (Thông tin và dữ liệu): Tìm kiếm, nhận diện chính xác các linh kiện kỹ thuật số và bản vẽ lắp ráp điện tử (Bậc 1 đối với Lớp 1-2; Bậc 2 đối với Lớp 3-7).")
    doc.add_paragraph("    + Miền V (Giải quyết vấn đề): Ứng dụng sáng tạo công nghệ số, điều khiển thiết bị tự động và tư duy thuật toán để giải quyết các nhiệm vụ thực tiễn.")

    p = doc.add_paragraph()
    r = p.add_run("  2.3. Năng lực chung:")
    r.italic = True
    doc.add_paragraph("    + Tự chủ và tự học: Tự nghiên cứu sơ đồ hướng dẫn, tự kiểm tra và hoàn thiện sản phẩm.")
    doc.add_paragraph("    + Giao tiếp và hợp tác: Phát triển kỹ năng làm việc nhóm, chia sẻ linh kiện và tinh thần đồng đội.")

    p = doc.add_paragraph()
    r = p.add_run("3. Về phẩm chất:")
    r.bold = True
    doc.add_paragraph("- Chăm chỉ, kiên trì, tỉ mỉ trong từng thao tác lắp ghép linh kiện.")
    doc.add_paragraph("- Trách nhiệm: Giữ gìn cẩn thận bộ linh kiện, thu dọn vệ sinh ngăn nắp sau giờ học.")
    doc.add_paragraph("- Trung thực: Nghiêm túc tuân thủ luật thi đấu và tôn trọng sản phẩm sáng tạo của bạn.")

    # III. KẾ HOẠCH DẠY HỌC CHI TIẾT
    p = doc.add_paragraph()
    r = p.add_run("III. KẾ HOẠCH DẠY HỌC CHI TIẾT (35 BUỔI)")
    r.bold = True

    # ---------------- BẢNG 1: PHÂN HỆ LỚP 1 - 2 ----------------
    p = doc.add_paragraph()
    r = p.add_run("1. Bảng phân phối chương trình Phân hệ 1 (Khối lớp 1 - 2):")
    r.bold = True
    p_note = doc.add_paragraph()
    r = p_note.add_run("• Học kỳ 1: Khám phá cùng OLLO Initiate (17 buổi)  |  • Học kỳ 2: Nâng cấp cùng OLLO Kinder (18 buổi)")
    r.italic = True

    tbl1 = doc.add_table(rows=1, cols=4)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl1)

    headers = ["STT", "Bài học / Chủ đề (1)", "Thời lượng (2)", "Yêu cầu cần đạt (3)"]
    hdr_cells = tbl1.rows[0].cells
    widths = [Inches(0.9), Inches(2.3), Inches(1.1), Inches(2.7)]
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        set_cell_background(hdr_cells[i], "D9E1F2")
        set_cell_margins(hdr_cells[i])
        hdr_cells[i].width = widths[i]

    # Banner HK1 Initiate
    r_banner1 = tbl1.add_row()
    c_b1 = r_banner1.cells[0]
    r_banner1.cells[0].merge(r_banner1.cells[3])
    c_b1.text = "HỌC KÌ I: KHÁM PHÁ THẾ GIỚI CHUYỂN ĐỘNG CÙNG OLLO INITIATE (17 BUỔI)"
    c_b1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_b1.paragraphs[0].runs[0].bold = True
    set_cell_background(c_b1, "F2F2F2")
    set_cell_margins(c_b1)

    for stt, bai, tg, yccd in DATA_LOP_1_2[:17]:
        row = tbl1.add_row()
        for i, val in enumerate([stt, bai, tg, yccd]):
            cell = row.cells[i]
            cell.text = val
            set_cell_margins(cell)
            cell.width = widths[i]
            if i in [0, 2]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Banner HK2 Kinder
    r_banner2 = tbl1.add_row()
    c_b2 = r_banner2.cells[0]
    r_banner2.cells[0].merge(r_banner2.cells[3])
    c_b2.text = "HỌC KÌ II: NÂNG CẤP KỸ NĂNG CƠ CẤU MÁY CÙNG OLLO KINDER (18 BUỔI)"
    c_b2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_b2.paragraphs[0].runs[0].bold = True
    set_cell_background(c_b2, "F2F2F2")
    set_cell_margins(c_b2)

    for stt, bai, tg, yccd in DATA_LOP_1_2[17:]:
        row = tbl1.add_row()
        for i, val in enumerate([stt, bai, tg, yccd]):
            cell = row.cells[i]
            cell.text = val
            set_cell_margins(cell)
            cell.width = widths[i]
            if i in [0, 2]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph() # Spacing

    # ---------------- BẢNG 2: PHÂN HỆ LỚP 3 - 7 ----------------
    p = doc.add_paragraph()
    r = p.add_run("2. Bảng phân phối chương trình Phân hệ 2 (Khối lớp 3 - 7):")
    r.bold = True
    p_note = doc.add_paragraph()
    r = p_note.add_run("• Học kỳ 1: Khám phá cùng OLLO Kinder (17 buổi)  |  • Học kỳ 2: Bứt phá công nghệ cùng OLLO Excel Level 1 (18 buổi)")
    r.italic = True

    tbl2 = doc.add_table(rows=1, cols=4)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl2)

    hdr_cells2 = tbl2.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells2[i].text = h
        hdr_cells2[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells2[i].paragraphs[0].runs[0].bold = True
        set_cell_background(hdr_cells2[i], "D9E1F2")
        set_cell_margins(hdr_cells2[i])
        hdr_cells2[i].width = widths[i]

    # Banner HK1 Kinder
    r_banner21 = tbl2.add_row()
    c_b21 = r_banner21.cells[0]
    r_banner21.cells[0].merge(r_banner21.cells[3])
    c_b21.text = "HỌC KÌ I: CƠ CẤU CHUYỂN ĐỘNG VÀ KỸ THUẬT CÙNG OLLO KINDER (17 BUỔI)"
    c_b21.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_b21.paragraphs[0].runs[0].bold = True
    set_cell_background(c_b21, "F2F2F2")
    set_cell_margins(c_b21)

    for stt, bai, tg, yccd in DATA_LOP_3_7[:17]:
        row = tbl2.add_row()
        for i, val in enumerate([stt, bai, tg, yccd]):
            cell = row.cells[i]
            cell.text = val
            set_cell_margins(cell)
            cell.width = widths[i]
            if i in [0, 2]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Banner HK2 Excel
    r_banner22 = tbl2.add_row()
    c_b22 = r_banner22.cells[0]
    r_banner22.cells[0].merge(r_banner22.cells[3])
    c_b22.text = "HỌC KÌ II: ĐỘT PHÁ CÔNG NGHỆ CAO CÙNG OLLO EXCEL LEVEL 1 (18 BUỔI)"
    c_b22.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_b22.paragraphs[0].runs[0].bold = True
    set_cell_background(c_b22, "F2F2F2")
    set_cell_margins(c_b22)

    for stt, bai, tg, yccd in DATA_LOP_3_7[17:]:
        row = tbl2.add_row()
        for i, val in enumerate([stt, bai, tg, yccd]):
            cell = row.cells[i]
            cell.text = val
            set_cell_margins(cell)
            cell.width = widths[i]
            if i in [0, 2]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph() # Spacing

    # IV. KIỂM TRA, ĐÁNH GIÁ ĐỊNH KỲ
    p = doc.add_paragraph()
    r = p.add_run("IV. KẾ HOẠCH KIỂM TRA, ĐÁNH GIÁ ĐỊNH KỲ")
    r.bold = True

    tbl_dg = doc.add_table(rows=1, cols=5)
    tbl_dg.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_dg)

    headers_dg = ["Bài kiểm tra, đánh giá", "Thời gian (1)", "Thời điểm (2)", "Yêu cầu cần đạt (3)", "Hình thức (4)"]
    widths_dg = [Inches(1.2), Inches(0.8), Inches(1.1), Inches(2.7), Inches(1.2)]
    for i, h in enumerate(headers_dg):
        cell = tbl_dg.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "D9E1F2")
        set_cell_margins(cell)
        cell.width = widths_dg[i]

    DATA_DG = [
        ("Đánh giá định kỳ 1", "90 phút", "Buổi 9\n(07/11/2026)",
         "• Phân hệ 1 (Lớp 1-2): Đánh giá khả năng nhận diện linh kiện Initiate, phân loại chi tiết và tốc độ lắp ráp chuẩn mô hình theo bản vẽ.\n• Phân hệ 2 (Lớp 3-7): Đánh giá kỹ năng phân tích cơ cấu cơ học Kinder, kết nối trục xoay, bánh răng và khả năng giải bài toán kỹ thuật theo nhóm.",
         "Thực hành lắp ráp & Quan sát kỹ năng"),
        ("Đánh giá định kỳ 2", "90 phút", "Buổi 17\n(02/01/2027)",
         "• Phân hệ 1 (Lớp 1-2): Đánh giá khả năng sáng tạo cải tiến mô hình OLLO Initiate; sự tự tin giới thiệu sản phẩm và tinh thần làm việc nhóm.\n• Phân hệ 2 (Lớp 3-7): Đánh giá sản phẩm dự án cơ cấu máy OLLO Kinder hoàn chỉnh; thuyết trình nguyên lý vận hành và trả lời câu hỏi phản biện.",
         "Triển lãm sản phẩm & Báo cáo dự án"),
        ("Đánh giá định kỳ 3", "90 phút", "Buổi 28\n(03/04/2027)",
         "• Phân hệ 1 (Lớp 1-2): Đánh giá khả năng làm chủ thiết bị OLLO Kinder nâng cao, thao tác lắp ghép khung xe và cơ cấu chịu tải.\n• Phân hệ 2 (Lớp 3-7): Đánh giá năng lực tích hợp động cơ Dynamixel, cảm biến âm thanh, cảm biến hồng ngoại và kỹ năng điều khiển không dây RC-100 OLLO Excel.",
         "Thực hành kỹ thuật & Đấu thử nghiệm"),
        ("Đánh giá định kỳ 4", "90 phút", "Buổi 34\n(22/05/2027)",
         "• Phân hệ 1 (Lớp 1-2): Đánh giá dự án mô hình sáng tạo cuối năm OLLO Kinder, kỹ năng xử lý lỗi cơ học và tinh thần hợp tác đồng đội.\n• Phân hệ 2 (Lớp 3-7): Đánh giá dự án giải pháp công nghệ OLLO Excel hoàn thiện; năng lực tư duy thuật toán, tối ưu cơ khí và chiến thuật thi đấu.",
         "Bảo vệ dự án & Tranh tài võ đài")
    ]

    for dg_name, tg, td, yccd, ht in DATA_DG:
        row = tbl_dg.add_row()
        for i, val in enumerate([dg_name, tg, td, yccd, ht]):
            cell = row.cells[i]
            cell.text = val
            set_cell_margins(cell)
            cell.width = widths_dg[i]
            if i in [0, 1, 2, 4]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph() # Spacing

    # V. TỔ CHỨC THỰC HIỆN
    p = doc.add_paragraph()
    r = p.add_run("V. TỔ CHỨC THỰC HIỆN")
    r.bold = True
    doc.add_paragraph("1. Giáo viên phụ trách:")
    doc.add_paragraph("- Soạn giáo án (Kế hoạch bài dạy) chi tiết trước mỗi buổi học theo đúng quy chuẩn UNIGO và Công văn 5512/BGDĐT, bảo đảm mục tiêu Năng lực số Thông tư 02/2025.")
    doc.add_paragraph("- Kiểm tra, rà soát tình trạng hoạt động của bộ Kit (OLLO Initiate, Kinder, Excel), sạc pin đầy đủ và chuẩn bị đầy đủ vật liệu thử nghiệm trước mỗi ca học.")
    doc.add_paragraph("- Quản lý an toàn lớp học, hướng dẫn học sinh quy tắc giữ gìn thiết bị và kiểm đếm linh kiện cẩn thận sau khi kết thúc buổi sinh hoạt.")
    doc.add_paragraph("2. Học sinh tham gia CLB:")
    doc.add_paragraph("- Đi học đúng giờ (09h45 sáng Thứ 7), trang phục nghiêm chỉnh, chấp hành nội quy phòng thực hành Robotics.")
    doc.add_paragraph("- Có ý thức kỷ luật, bảo quản đồ dùng, linh kiện, không làm rơi vãi chốt ghép và tích cực hợp tác cùng bạn học.")
    doc.add_paragraph("3. Phối hợp với Ban Giám hiệu và Phụ huynh học sinh:")
    doc.add_paragraph("- Thường xuyên cập nhật hình ảnh, video hoạt động và tiến độ học tập của học sinh lên kênh thông tin nhà trường.")
    doc.add_paragraph("- Tổ chức các sự kiện Triển lãm giữa kỳ (Buổi 17) và Ngày hội Festival Robotics cuối năm (Buổi 35) để phụ huynh tham quan, trải nghiệm và chứng kiến sự tiến bộ vượt bậc của học sinh.")

    doc.add_paragraph() # Spacing

    # VI. BẢNG PHÊ DUYỆT (Table 3x3 chữ ký)
    tbl_sig = doc.add_table(rows=3, cols=3)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(tbl_sig)
    widths_sig = [Inches(2.3), Inches(2.3), Inches(2.4)]

    # Row 0
    headers_sig = ["DUYỆT CỦA BGH", "DUYỆT CỦA TỔ CM", "NGƯỜI LẬP KẾ HOẠCH"]
    for i, h in enumerate(headers_sig):
        c = tbl_sig.cell(0, i)
        c.text = h
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].bold = True
        c.width = widths_sig[i]
        set_cell_margins(c)

    # Row 1
    sub_sig = ["(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"]
    for i, s in enumerate(sub_sig):
        c = tbl_sig.cell(1, i)
        c.text = s
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].italic = True
        c.width = widths_sig[i]
        set_cell_margins(c)

    # Row 2
    names_sig = ["\n\n\n\n", "\n\n\n\n", "\n\n\n\nĐậu Đình Nguyên"]
    for i, n in enumerate(names_sig):
        c = tbl_sig.cell(2, i)
        c.text = n
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if "Đậu Đình Nguyên" in n:
            c.paragraphs[0].runs[0].bold = True
        c.width = widths_sig[i]
        set_cell_margins(c)

    # Apply font Times New Roman 13pt
    format_all_fonts(doc)

    doc.save(DOCX_PATH)
    print("Successfully generated and saved:", DOCX_PATH)

if __name__ == "__main__":
    generate_khdh_docx()
