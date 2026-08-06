"""
generate_lbg.py — Tu dong tao Lich bao giang theo tuan
======================================================

QUY TAC ROTATION:
  - Lop 5, 6, 7, 8: Moi lop co 2 tiet / tuan
      Tuan CHAN (2, 4, 6...): ca 2 tiet -> Tin hoc
      Tuan LE  (3, 5, 7...): ca 2 tiet -> Robotics
  - Cac lop con lai: giu nguyen mon hoc goc

QUY TAC PPCT (Khong offset):
  - Tuan 1: Tat ca = Tiet 0 (Dinh huong)
  - Tuan 2+: Tat ca cac Thu (T2 -> T6) đeu dong bo ppct = tuan_so - 1
  - Rotation classes (5,6,7,8): PPCT rieng cho Tin/Robotics
      Khi co 2 tiet lien tiep cung mon -> PPCT lien tiep (Tuan 2: PPCT 1,2; Tuan 4: PPCT 3,4...)

QUY TRINH (moi tuan):
  1. Doc file Tuan 01 lam template
  2. Cap nhat tieu de tuan + ngay
  3. Cap nhat nhan Thu/ngay trong bang
  4. Ap dung rotation cho lop 5,6,7,8
  5. Cap nhat PPCT + Ten bai tu PPCT_DATA
  6. Luu ban goc tong hop
  7. Tach 2 ban: TTH+TH va THCS
  8. Chen page break (sang/chieu/nhan xet = 3 trang)

CACH DUNG:
  python generate_lbg.py <so_tuan>
  Vi du: python generate_lbg.py 2
         python generate_lbg.py 3

OUTPUT (trong thu muc LBG_DIR):
  Lich bao giang - Tuan XX.docx          <- ban goc tong hop
  Lich bao giang - Tuan XX (TTH+TH).docx <- Tien TH + TH
  Lich bao giang - Tuan XX (THCS).docx   <- THCS
"""

import sys
import os
import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import date, timedelta

sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────────────────────────────────────
# CAU HINH
# ─────────────────────────────────────────────────────────────────────────────

LBG_DIR = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Lịch báo giảng'
TEMPLATE = os.path.join(LBG_DIR, 'Lịch báo giảng - Tuần 01.docx')

# Ngay dau Tuan 01 (Thu Hai 03/08/2026)
TUAN_01_START = date(2026, 8, 3)

# Lop co rotation Tin hoc / Robotics tuan chan-le (bat dau tu lop 5)
ROTATION_PREFIXES = ('5', '6', '7', '8')

THU_LABELS = {0: 'Hai', 1: 'Ba', 2: 'Tư', 3: 'Năm', 4: 'Sáu'}

# ─────────────────────────────────────────────────────────────────────────────
# PPCT_DATA: Mapping class -> {ppct_num: lesson_name}
# ppct_num starts at 0 (Dinh huong) then 1, 2, 3...
# In PPCT files, PPCT=1 = "Tiết 0: Định hướng", PPCT=2 = "Bài 1..."
# But in LBG, we use ppct=0 for orientation, ppct=1 for first real lesson
# So LBG ppct N maps to PPCT_DATA key N+1
# ─────────────────────────────────────────────────────────────────────────────

PPCT_TIN = {
    # Tiền Tiểu học (TT3, TTH 1, TTH2 - dùng PPCT chung TTH)
    'TT3': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Em làm quen với thế giới công nghệ',
        2: 'Máy tính quanh em',
        3: 'Em ngồi máy tính an toàn',
        4: 'Làm quen với chuột',
        5: 'Kéo – thả thật vui',
        6: 'Khám phá bàn phím',
        7: 'Chữ, hình và âm thanh',
        8: 'Màu sắc và hình dạng',
        9: 'Em tạo bức tranh đầu tiên',
        10: 'Ôn tập kỹ năng số',
        11: 'ĐÁNH GIÁ ĐỊNH KỲ 1',
        12: 'Thông tin quanh em',
        13: 'Máy tính làm việc theo lệnh',
        14: 'Một việc – nhiều bước',
        15: 'Chỉ đường cho nhân vật',
        16: 'Mê cung của em',
        17: 'Sửa lỗi đường đi',
        18: 'Làm quen lập trình trực quan',
        19: 'Ôn tập Coding cơ bản',
        20: 'ĐÁNH GIÁ ĐỊNH KỲ 2',
        21: 'Kể chuyện số',
        22: 'Tạo nhân vật',
        23: 'Tạo bối cảnh',
        24: 'Nhân vật chuyển động',
        25: 'Âm thanh và lời kể',
        26: 'Dự án: Câu chuyện số',
        27: 'Chia sẻ và làm việc nhóm',
        28: 'Hoàn thiện sản phẩm',
        29: 'ĐÁNH GIÁ ĐỊNH KỲ 3',
        30: 'Robot quanh em',
        31: 'Điều khiển robot',
        32: 'Robot thực hiện nhiệm vụ',
        33: 'AI quanh em',
        34: 'Sử dụng công nghệ an toàn',
        35: 'ĐÁNH GIÁ ĐỊNH KỲ 4',
    },
    # Lớp 1 (1A1, 1C1)
    '1': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Em làm quen với thế giới công nghệ',
        2: 'Ôn và nâng cấp kỹ năng chuột',
        3: 'Bàn phím và gõ câu ngắn',
        4: 'Mở, đóng và chuyển đổi ứng dụng',
        5: 'Tệp là gì?',
        6: 'Thư mục là gì?',
        7: 'Đặt tên và lưu sản phẩm',
        8: 'Mở lại và sắp xếp sản phẩm',
        9: 'Tạo một sản phẩm có tổ chức',
        10: 'Ôn tập tệp – thư mục',
        11: 'ĐÁNH GIÁ ĐỊNH KỲ 1',
        12: 'Thông tin và dữ liệu',
        13: 'Thông tin dạng chữ, hình, âm thanh, video',
        14: 'Thu nhận và xử lý thông tin',
        15: 'Một nhiệm vụ thành nhiều bước',
        16: 'Thuật toán bằng lời và hình',
        17: 'Lập trình trực quan: chuỗi lệnh',
        18: 'Lặp lại một hành động',
        19: 'Chạy thử và sửa lỗi',
        20: 'ĐÁNH GIÁ ĐỊNH KỲ 2',
        21: 'Thiết kế câu chuyện số',
        22: 'Thiết kế nhân vật',
        23: 'Thiết kế bối cảnh',
        24: 'Animation: chuyển động',
        25: 'Tương tác và âm thanh',
        26: 'Dự án: Hoạt hình ngắn',
        27: 'Thiết kế sản phẩm nhóm',
        28: 'Kiểm thử và hoàn thiện',
        29: 'ĐÁNH GIÁ ĐỊNH KỲ 3',
        30: 'Robot và cảm biến',
        31: 'Lập trình robot theo nhiệm vụ',
        32: 'Robot giải quyết vấn đề',
        33: 'AI có thể làm gì?',
        34: 'AI có thể sai – dữ liệu cần kiểm tra',
        35: 'ĐÁNH GIÁ ĐỊNH KỲ 4',
    },
    # Lớp 2 (2A1, 2C1)
    '2': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Em trở thành nhà sáng tạo số',
        2: 'Ôn và nâng cấp kỹ năng chuột',
        3: 'Bàn phím và gõ câu ngắn',
        4: 'Mở, đóng và chuyển đổi ứng dụng',
        5: 'Tệp là gì?',
        6: 'Thư mục là gì?',
        7: 'Đặt tên và lưu sản phẩm',
        8: 'Mở lại và sắp xếp sản phẩm',
        9: 'Tạo một sản phẩm có tổ chức',
        10: 'Ôn tập tệp – thư mục',
        11: 'ĐÁNH GIÁ ĐỊNH KỲ 1',
        12: 'Thông tin và dữ liệu',
        13: 'Thông tin dạng chữ, hình, âm thanh, video',
        14: 'Thu nhận và xử lý thông tin',
        15: 'Một nhiệm vụ thành nhiều bước',
        16: 'Thuật toán bằng lời và hình',
        17: 'Lập trình trực quan: chuỗi lệnh',
        18: 'Lặp lại một hành động',
        19: 'Chạy thử và sửa lỗi',
        20: 'ĐÁNH GIÁ ĐỊNH KỲ 2',
        21: 'Thiết kế câu chuyện số',
        22: 'Thiết kế nhân vật',
        23: 'Thiết kế bối cảnh',
        24: 'Animation: chuyển động',
        25: 'Tương tác và âm thanh',
        26: 'Dự án: Hoạt hình ngắn',
        27: 'Thiết kế sản phẩm nhóm',
        28: 'Kiểm thử và hoàn thiện',
        29: 'ĐÁNH GIÁ ĐỊNH KỲ 3',
        30: 'Robot và cảm biến',
        31: 'Lập trình robot theo nhiệm vụ',
        32: 'Robot giải quyết vấn đề',
        33: 'AI có thể làm gì?',
        34: 'AI có thể sai – dữ liệu cần kiểm tra',
        35: 'ĐÁNH GIÁ ĐỊNH KỲ 4',
    },
    # Lớp 3 (3A1, 3C1) - from Table 4 Tin TH
    '3': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1: Thông tin và quyết định',
        2: 'Bài 2: Xử lí thông tin',
        3: 'Bài 3: Máy tính và em',
        4: 'Bài 4: Làm việc với máy tính',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5: Sử dụng bàn phím',
        8: 'Bài 6: Khám phá thông tin trên Internet',
        9: 'Bài 7: Sắp xếp để dễ tìm',
        10: 'Bài 8: Sơ đồ hình cây. Tổ chức thông tin trong máy tính',
        11: 'Ôn tập Đánh giá định kỳ 2',
        12: 'Đánh giá định kỳ 2',
        13: 'Bài 11: Thực hành với tệp và thư mục trong máy tính',
        14: 'Bài 12: Bảo vệ thông tin khi dùng máy tính',
        15: 'Bài 13: Bài trình chiếu của em',
        16: 'Bài 14: Tìm hiểu về thế giới tự nhiên',
        17: 'Bài 15: Luyện tập sử dụng chuột',
        18: 'Ôn tập Đánh giá định kỳ 3',
        19: 'Đánh giá định kỳ 3',
        20: 'Bài 16: Em thực hiện công việc như thế nào?',
        21: 'Ôn tập và luyện tập',
        22: 'Ôn tập Đánh giá định kỳ 4',
        23: 'Đánh giá định kỳ 4',
        24: 'Tổng kết năm học',
    },
    # Lớp 4 (4C1) - from Table 5 Tin TH
    '4': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1: Phần cứng và phần mềm máy tính',
        2: 'Bài 2: Gõ bàn phím đúng cách',
        3: 'Bài 3: Thông tin trên trang Web',
        4: 'Bài 4: Tìm kiếm thông tin trên Internet',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5: Cây thư mục',
        8: 'Bài 6: Sử dụng phần mềm khi được phép',
        9: 'Bài 7: Tạo bài trình chiếu',
        10: 'Bài 8: Định dạng văn bản trên trang chiếu',
        11: 'Bài 9: Hiệu ứng chuyển trang',
        12: 'Ôn tập Đánh giá định kỳ 2',
        13: 'Đánh giá định kỳ 2',
        14: 'Bài 12: Phần mềm soạn thảo văn bản',
        15: 'Bài 13: Chỉnh sửa văn bản',
        16: 'Bài 14: Luyện tập gõ bàn phím / Đa phương tiện',
        17: 'Bài 15: Chơi với máy tính',
        18: 'Ôn tập Đánh giá định kỳ 3',
        19: 'Đánh giá định kỳ 3',
        20: 'Bài 16: Khám phá môi trường lập trình trực quan',
        21: 'Ôn tập và luyện tập',
        22: 'Ôn tập Đánh giá định kỳ 4',
        23: 'Đánh giá định kỳ 4',
        24: 'Tổng kết năm học',
    },
    # Lớp 5 (5C1) - from Table 6 Tin TH
    '5': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1: Em có thể làm gì với máy tính?',
        2: 'Bài 2: Tìm kiếm thông tin trên website',
        3: 'Bài 3: Tìm kiếm thông tin trong giải quyết vấn đề',
        4: 'Bài 4: Cây thư mục',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5: Bản quyền nội dung thông tin',
        8: 'Bài 6: Định dạng kí tự và bố trí hình ảnh trong văn bản',
        9: 'Bài 7: Thực hành soạn thảo văn bản',
        10: 'Bài 8: Sản phẩm đồ họa / Sản phẩm thủ công',
        11: 'Ôn tập Đánh giá định kỳ 2',
        12: 'Đánh giá định kỳ 2',
        13: 'Bài 11: Thực hành tạo sản phẩm số',
        14: 'Bài 12: Cấu trúc tuần tự',
        15: 'Bài 13: Cấu trúc lặp',
        16: 'Bài 14: Thực hành sử dụng lệnh lặp',
        17: 'Ôn tập Đánh giá định kỳ 3',
        18: 'Đánh giá định kỳ 3',
        19: 'Bài 15: Cấu trúc rẽ nhánh',
        20: 'Bài 16: Sử dụng biến trong chương trình',
        21: 'Ôn tập và luyện tập',
        22: 'Ôn tập Đánh giá định kỳ 4',
        23: 'Đánh giá định kỳ 4',
        24: 'Tổng kết năm học',
    },
    # Lớp 6 (6A1) - from Table 3 Tin THCS
    '6': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Thông tin và dữ liệu',
        2: 'Bài 2. Xử lí thông tin',
        3: 'Bài 3. Thông tin trong máy tính',
        4: 'Bài 4. Mạng máy tính',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5. Internet',
        8: 'Bài 6. Mạng thông tin toàn cầu',
        9: 'Bài 7. Tìm kiếm thông tin trên Internet',
        10: 'Bài 8. Thư điện tử',
        11: 'Ôn tập Đánh giá định kỳ 2',
        12: 'Đánh giá định kỳ 2',
        13: 'Bài 9. An toàn thông tin trên Internet',
        14: 'Bài 10. Sơ đồ tư duy',
        15: 'Bài 11. Định dạng văn bản',
        16: 'Bài 12. Trình bày thông tin ở dạng bảng',
        17: 'Ôn tập Đánh giá định kỳ 3',
        18: 'Đánh giá định kỳ 3',
        19: 'Bài 13. Tìm kiếm và thay thế',
        20: 'Ôn tập và luyện tập',
        21: 'Ôn tập Đánh giá định kỳ 4',
        22: 'Đánh giá định kỳ 4',
        23: 'Tổng kết năm học',
    },
    # Lớp 7 (7A1) - from Table 4 Tin THCS
    '7': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Thiết bị vào - ra',
        2: 'Bài 2. Phần mềm máy tính',
        3: 'Bài 3. Quản lý dữ liệu trong máy tính',
        4: 'Bài 4. Mạng xã hội và một số kênh trao đổi thông tin trên Internet',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5. Ứng xử trên mạng',
        8: 'Bài 6. Làm quen với phần mềm bảng tính',
        9: 'Bài 7. Tính toán tự động trên bảng tính',
        10: 'Bài 8. Công cụ hỗ trợ tính toán',
        11: 'Ôn tập Đánh giá định kỳ 2',
        12: 'Đánh giá định kỳ 2',
        13: 'Bài 9. Trình bày bảng tính',
        14: 'Bài 10. Hoàn thiện bảng tính',
        15: 'Bài 11. Tạo bài trình chiếu',
        16: 'Bài 12. Định dạng đối tượng trên trang chiếu',
        17: 'Bài 13. Thực hành tổng hợp: Hoàn thiện bài trình chiếu',
        18: 'Ôn tập Đánh giá định kỳ 3',
        19: 'Đánh giá định kỳ 3',
        20: 'Ôn tập và luyện tập',
        21: 'Ôn tập Đánh giá định kỳ 4',
        22: 'Đánh giá định kỳ 4',
        23: 'Tổng kết năm học',
    },
    # Lớp 8 (8A1) - from Table 5 Tin THCS
    '8': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Lược sử công cụ tính toán',
        2: 'Bài 2. Thông tin trong môi trường số',
        3: 'Bài 3. Thực hành khai thác thông tin số',
        4: 'Ôn tập Đánh giá định kỳ 1',
        5: 'Đánh giá định kỳ 1',
        6: 'Bài 4. Đạo đức và văn hóa trong sử dụng công nghệ số',
        7: 'Bài 5. Sử dụng bảng tính giải quyết bài toán thực tế',
        8: 'Bài 6. Sắp xếp và lọc dữ liệu',
        9: 'Bài 7. Trình bày dữ liệu bằng biểu đồ',
        10: 'Ôn tập Đánh giá định kỳ 2',
        11: 'Đánh giá định kỳ 2',
        12: 'Bài 8a. Làm việc với danh sách dạng liệt kê và hình ảnh trong văn bản',
        13: 'Bài 9a. Tạo đầu trang, chân trang cho văn bản',
        14: 'Bài 10a. Định dạng nâng cao cho trang chiếu',
        15: 'Bài 11a. Sử dụng bản mẫu cho bài trình chiếu',
        16: 'Ôn tập Đánh giá định kỳ 3',
        17: 'Đánh giá định kỳ 3',
        18: 'Bài 12. Từ thuật toán đến chương trình',
        19: 'Ôn tập và luyện tập',
        20: 'Ôn tập Đánh giá định kỳ 4',
        21: 'Đánh giá định kỳ 4',
        22: 'Bài 16. Tin học với nghề nghiệp',
    },
}

# Robotics PPCT (shared for 6,7,8; separate for TH classes)
PPCT_ROB = {
    # Robotics cho lop 1 (1A1, 1C1) - giong Rob TH Table 5
    '1': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Tập thể dục nào!',
        2: 'Bài 2. Chú cún dễ thương',
        3: 'Bài 3. Tăng cường sức khỏe',
        4: 'Bài 4. Chú ốc sên chậm chạp',
        5: 'Bài 5. Xe cảnh sát tuần tra',
        6: 'Bài 6. Khám phá xe cứu hoả',
        7: 'Bài 7. Người giao hàng đã đến',
        8: 'Bài 8. Thế giới khủng long',
        9: 'Ôn tập Đánh giá định kỳ 1',
        10: 'Đánh giá định kỳ 1',
        11: 'Bài 9. Hồ bơi mùa hè',
        12: 'Bài 10. Khám phá đại dương',
        13: 'Bài 11. Chú cua cứng cáp',
        14: 'Bài 12. Tôi có thể di chuyển đến bất cứ đâu',
        15: 'Bài 13. Hoạt động mùa hè',
        16: 'Bài 14. Bắt đầu chuyến hành trình cùng tàu hoả',
        17: 'Bài 15. Phía trên bầu trời',
        18: 'Ôn tập Đánh giá định kỳ 2',
        19: 'Đánh giá định kỳ 2',
        20: 'Bài 16. Khám phá vũ trụ rộng lớn',
        21: 'Bài 17. Máy bắn đá khổng lồ',
        22: 'Bài 18. Trò chơi dân gian',
        23: 'Bài 19. Khám phá trò chơi truyền thống các nước',
        24: 'Bài 20. Đấu vật thú vị',
        25: 'Bài 21. Sóc nhỏ dễ thương',
        26: 'Bài 22. Chú hươu tuyệt đẹp',
        27: 'Ôn tập Đánh giá định kỳ 3',
        28: 'Đánh giá định kỳ 3',
        29: 'Bài 23. Chú rùa thông minh',
        30: 'Bài 24. Đôi chân mạnh mẽ của chuột túi',
        31: 'Bài 25. Chú sư tử dũng mãnh',
        32: 'Bài 26. Chuột chũi, máy ủi dưới lòng đất',
        33: 'Ôn tập Đánh giá định kỳ 4',
        34: 'Đánh giá định kỳ 4',
        35: 'Tổng kết môn học & Triển lãm Robotics',
    },
    # Robotics cho lop 2 (2A1, 2C1) - giong Rob TH Table 6
    '2': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Hãy nấu những món ăn ngon',
        2: 'Bài 2. Hàm răng trắng sáng',
        3: 'Bài 3. Sự phản xạ ánh sáng',
        4: 'Bài 4. Đàn gà con',
        5: 'Bài 5. Những bạn nhỏ lễ phép',
        6: 'Bài 6. Động vật thân mềm',
        7: 'Bài 7. Khu phố của chúng ta',
        8: 'Bài 8. Rèn luyện sức khỏe',
        9: 'Ôn tập Đánh giá định kỳ 1',
        10: 'Đánh giá định kỳ 1',
        11: 'Bài 9. Môi trường biển',
        12: 'Bài 10. Cùng câu cá nào!',
        13: 'Bài 11. Thế giới khủng long',
        14: 'Bài 12. Người hiệp sĩ dũng cảm',
        15: 'Bài 13. Vận chuyển đồ vật',
        16: 'Bài 14. Sức mạnh của máy ủi',
        17: 'Bài 15. Máy xúc',
        18: 'Ôn tập Đánh giá định kỳ 2',
        19: 'Đánh giá định kỳ 2',
        20: 'Bài 16. Cùng nhau đi khắp thế giới',
        21: 'Bài 17. Nhắm và bắn!',
        22: 'Bài 18. Độ đàn hồi, lực đẩy của cung tên',
        23: 'Bài 19. Ba! Hai! Một! Bắn!!!',
        24: 'Bài 20. Cùng nhau tham quan thành phố',
        25: 'Bài 21. Khám phá trang phục người da đỏ',
        26: 'Bài 22. Cá sấu thật ngầu!',
        27: 'Ôn tập Đánh giá định kỳ 3',
        28: 'Đánh giá định kỳ 3',
        29: 'Bài 23. Những nốt nhạc vui',
        30: 'Bài 24. Choo Choo! Tàu hỏa',
        31: 'Bài 25. Có sáu chân thật tuyệt!!!',
        32: 'Bài 26. Bầu không khí trong lành',
        33: 'Ôn tập Đánh giá định kỳ 4',
        34: 'Đánh giá định kỳ 4',
        35: 'Tổng kết môn học & Triển lãm Robotics',
    },
    # Robotics cho lop 3 (3A1, 3C1) - giong Rob TH Table 7
    '3': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Hãy nấu những món ăn ngon',
        2: 'Bài 2. Hàm răng trắng sáng',
        3: 'Bài 3. Sự phản xạ ánh sáng',
        4: 'Bài 4. Đàn gà con',
        5: 'Bài 5. Những bạn nhỏ lễ phép',
        6: 'Bài 6. Động vật thân mềm',
        7: 'Bài 7. Khu phố của chúng ta',
        8: 'Bài 8. Rèn luyện sức khỏe',
        9: 'Ôn tập Đánh giá định kỳ 1',
        10: 'Đánh giá định kỳ 1',
        11: 'Bài 9. Môi trường biển',
        12: 'Bài 10. Cùng câu cá nào!',
        13: 'Bài 11. Thế giới khủng long',
        14: 'Bài 12. Người hiệp sĩ dũng cảm',
        15: 'Bài 13. Vận chuyển đồ vật',
        16: 'Bài 14. Sức mạnh của máy ủi',
        17: 'Bài 15. Máy xúc',
        18: 'Ôn tập Đánh giá định kỳ 2',
        19: 'Đánh giá định kỳ 2',
        20: 'Bài 16. Cùng nhau đi khắp thế giới',
        21: 'Bài 17. Nhắm và bắn!',
        22: 'Bài 18. Độ đàn hồi, lực đẩy của cung tên',
        23: 'Bài 19. Ba! Hai! Một! Bắn!!!',
        24: 'Bài 20. Cùng nhau tham quan thành phố',
        25: 'Bài 21. Khám phá trang phục người da đỏ',
        26: 'Bài 22. Cá sấu thật ngầu!',
        27: 'Ôn tập Đánh giá định kỳ 3',
        28: 'Đánh giá định kỳ 3',
        29: 'Bài 23. Những nốt nhạc vui',
        30: 'Bài 24. Choo Choo! Tàu hỏa',
        31: 'Bài 25. Có sáu chân thật tuyệt!!!',
        32: 'Bài 26. Bầu không khí trong lành',
        33: 'Ôn tập Đánh giá định kỳ 4',
        34: 'Đánh giá định kỳ 4',
        35: 'Tổng kết môn học & Triển lãm Robotics',
    },
    # Robotics cho lop 4 (4C1) - giong Rob TH Table 7 (same kit)
    '4': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Hãy nấu những món ăn ngon',
        2: 'Bài 2. Hàm răng trắng sáng',
        3: 'Bài 3. Sự phản xạ ánh sáng',
        4: 'Bài 4. Đàn gà con',
        5: 'Bài 5. Những bạn nhỏ lễ phép',
        6: 'Bài 6. Động vật thân mềm',
        7: 'Bài 7. Khu phố của chúng ta',
        8: 'Bài 8. Rèn luyện sức khỏe',
        9: 'Ôn tập Đánh giá định kỳ 1',
        10: 'Đánh giá định kỳ 1',
        11: 'Bài 9. Môi trường biển',
        12: 'Bài 10. Cùng câu cá nào!',
        13: 'Bài 11. Thế giới khủng long',
        14: 'Bài 12. Người hiệp sĩ dũng cảm',
        15: 'Bài 13. Vận chuyển đồ vật',
        16: 'Bài 14. Sức mạnh của máy ủi',
        17: 'Bài 15. Máy xúc',
        18: 'Ôn tập Đánh giá định kỳ 2',
        19: 'Đánh giá định kỳ 2',
        20: 'Bài 16. Cùng nhau đi khắp thế giới',
        21: 'Bài 17. Nhắm và bắn!',
        22: 'Bài 18. Độ đàn hồi, lực đẩy của cung tên',
        23: 'Bài 19. Ba! Hai! Một! Bắn!!!',
        24: 'Bài 20. Cùng nhau tham quan thành phố',
        25: 'Bài 21. Khám phá trang phục người da đỏ',
        26: 'Bài 22. Cá sấu thật ngầu!',
        27: 'Ôn tập Đánh giá định kỳ 3',
        28: 'Đánh giá định kỳ 3',
        29: 'Bài 23. Những nốt nhạc vui',
        30: 'Bài 24. Choo Choo! Tàu hỏa',
        31: 'Bài 25. Có sáu chân thật tuyệt!!!',
        32: 'Bài 26. Bầu không khí trong lành',
        33: 'Ôn tập Đánh giá định kỳ 4',
        34: 'Đánh giá định kỳ 4',
        35: 'Tổng kết môn học & Triển lãm Robotics',
    },
    # Robotics THCS (5,6,7,8) - shared PPCT from Rob THCS
    '5': {
        0: 'Tiết 0: Định hướng môn học',
        1: 'Bài 1. Động cơ là gì?',
        2: 'Bài 2. Robot cố định vật',
        3: 'Bài 3. Robot nhận biết âm thanh',
        4: 'Bài 4. Bộ điều khiển từ xa',
        5: 'Ôn tập Đánh giá định kỳ 1',
        6: 'Đánh giá định kỳ 1',
        7: 'Bài 5. Động cơ Dynamixel hoạt động thế nào?',
        8: 'Bài 6. Phương tiện giao thông qua các thời đại',
        9: 'Bài 7. Cơ chế tăng giảm chiều dài tự động',
        10: 'Bài 8. Robot nhận biết vật thể bằng cách nào?',
        11: 'Ôn tập Đánh giá định kỳ 2',
        12: 'Đánh giá định kỳ 2',
        13: 'Bài 9. Số ngẫu nhiên',
        14: 'Bài 10. Robot hút bụi',
        15: 'Bài 11. Lực hấp dẫn',
        16: 'Bài 12. Domino trong Robotics',
        17: 'Ôn tập Đánh giá định kỳ 3',
        18: 'Đánh giá định kỳ 3',
        19: 'Bài 13. Robot phục vụ đời sống',
        20: 'Luyện tập & Thực hành sáng tạo Robotics',
        21: 'Ôn tập Đánh giá định kỳ 4',
        22: 'Đánh giá định kỳ 4',
        23: 'Tổng kết môn học & Triển lãm Robotics',
    },
}
# 6,7,8 cùng chương trình Robotics THCS
PPCT_ROB['6'] = PPCT_ROB['5'].copy()
PPCT_ROB['7'] = PPCT_ROB['5'].copy()
PPCT_ROB['8'] = PPCT_ROB['5'].copy()
# TTH dùng chung Rob lớp 1
PPCT_ROB['TT'] = PPCT_ROB['1'].copy()


# ─────────────────────────────────────────────────────────────────────────────
# TIEN ICH
# ─────────────────────────────────────────────────────────────────────────────

def week_dates(tuan_so):
    delta = timedelta(weeks=tuan_so - 1)
    start = TUAN_01_START + delta
    end = start + timedelta(days=4)
    return start, end


def fmt_date(d):
    return f'{d.day:02d}/{d.month:02d}/{d.year}'


def day_label(d):
    thu = THU_LABELS[d.weekday()]
    return f'{thu} ({d.day:02d}/{d.month:02d})'


def classify_lop(lop):
    """'TTH_TH' hoac 'THCS' hoac None."""
    s = lop.strip().upper()
    if not s:
        return None
    if re.match(r'^[6789]', s):
        return 'THCS'
    return 'TTH_TH'


def is_rotation_lop(lop):
    """Tra ve True neu lop thuoc nhom 5,6,7,8 (co rotation)."""
    s = lop.strip().upper()
    return any(s.startswith(p) for p in ROTATION_PREFIXES)


def rotation_mon(tuan_so):
    """Mon hoc cho lop rotation theo tuan: chan=Tin hoc, le=Robotics."""
    return 'Tin học' if tuan_so % 2 == 0 else 'Robotics'


def rotation_do_dung(tuan_so):
    return 'Phòng Tin học' if tuan_so % 2 == 0 else 'Bộ Kit Robotics'


def get_grade_key(lop):
    """Extract grade key from class name for PPCT lookup.
    Examples: '1A1' -> '1', '7A1' -> '7', 'TT3' -> 'TT3', 'TTH 1' -> 'TT', 'TTH2' -> 'TT'
    """
    s = lop.strip().upper()
    if s.startswith('TT'):
        return 'TT3'  # TTH classes use TT3 PPCT
    m = re.match(r'^(\d)', s)
    if m:
        return m.group(1)
    return None


def get_ppct_lesson(lop, mon, ppct_num):
    """Get lesson name for a class at given PPCT number."""
    grade = get_grade_key(lop)
    if grade is None:
        return ''

    if mon == 'Tin học':
        data = PPCT_TIN.get(grade, {})
    elif mon == 'Robotics':
        data = PPCT_ROB.get(grade, {})
    else:
        return ''

    return data.get(ppct_num, '')


def compute_ppct(tuan_so, day_idx=0):
    """
    Tinh PPCT theo tuan (Khong ap dung offset ngay trong tuan).
    - Tuan 1: tat ca = 0 (Dinh huong)
    - Tuan 2+: tat ca cac ngay = tuan_so - 1
    """
    if tuan_so <= 1:
        return 0
    return tuan_so - 1


def compute_rotation_ppct(tuan_so, day_idx, is_even_week, period_index_in_day):
    """
    Tinh PPCT cho lop rotation (5,6,7,8) co 2 tiet lien tiep (Khong offset).
    - Tuan CHAN (2, 4, 6...): Tin hoc (Tuan 2 -> PPCT 1,2; Tuan 4 -> PPCT 3,4;...)
    - Tuan LE (3, 5, 7...): Robotics (Tuan 3 -> PPCT 1,2; Tuan 5 -> PPCT 3,4;...)
    """
    if tuan_so <= 1:
        return 0
    
    if is_even_week:
        # Tin hoc in even weeks (2, 4, 6...)
        if tuan_so < 2:
            return 0
        k = (tuan_so - 2) // 2 + 1  # Session index (1-based)
        return (k - 1) * 2 + period_index_in_day + 1
    else:
        # Robotics in odd weeks (3, 5, 7...)
        if tuan_so < 3:
            return 0
        k = (tuan_so - 3) // 2 + 1  # Session index (1-based)
        return (k - 1) * 2 + period_index_in_day + 1


def set_cell_text(cell, text):
    """Ghi text vao cell, giu dinh dang run dau tien."""
    for p in cell.paragraphs:
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ''
            break
        else:
            p.text = text
            break


def save_safe(doc, path):
    """Luu file; neu bi khoa thi luu sang _v2."""
    try:
        doc.save(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        alt = base + '_v2' + ext
        doc.save(alt)
        print(f'   ⚠ Bi khoa, luu: {os.path.basename(alt)}')
        return alt


# ─────────────────────────────────────────────────────────────────────────────
# CAP NHAT NOI DUNG
# ─────────────────────────────────────────────────────────────────────────────

def update_headers(doc, tuan_so, start_date, end_date):
    """Cap nhat tieu de tuan va ngay trong paragraph."""
    tuan_str = f'TUẦN {tuan_so:02d}'
    fs = fmt_date(start_date)
    fe = fmt_date(end_date)
    ngay_soan = start_date - timedelta(days=3)  # Thu Sau tuan truoc

    for p in doc.paragraphs:
        txt = p.text.strip()

        if 'TUẦN' in txt.upper() and 'từ ngày' in txt:
            new = f'{tuan_str} (từ ngày {fs} đến ngày {fe})'
        elif re.match(r'^Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+', txt) and len(txt) < 45:
            new = (f'Ngày {ngay_soan.day:02d} tháng '
                   f'{ngay_soan.month:02d} năm {ngay_soan.year}')
        elif 'sáng' in txt.lower() and 'tuần' in txt.lower():
            new = (f'Buổi…sáng……..Tuần…{tuan_so:02d}…'
                   f'(Từ ngày…{fs} …đến ngày:… {fe}….)')
        elif 'chiều' in txt.lower() and 'tuần' in txt.lower():
            new = (f'Buổi…chiều…..Tuần…{tuan_so:02d}…'
                   f'(Từ ngày…{fs}…đến ngày:… {fe}….)')
        else:
            continue

        if p.runs:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ''
        else:
            p.text = new


def update_day_labels(doc, start_date):
    """Cap nhat nhan Thu/ngay trong Table[1] va Table[3]."""
    for ti in [1, 3]:
        tbl = doc.tables[ti]
        for ri in range(1, len(tbl.rows)):
            day_idx = (ri - 1) // 5   # 0=Hai..4=Sau
            d = start_date + timedelta(days=day_idx)
            label = day_label(d)
            cell0 = tbl.rows[ri].cells[0]
            if cell0.text.strip():
                set_cell_text(cell0, label)


def get_day_idx_from_row(ri):
    """Get day index (0=T2..4=T6) from row index in LBG table."""
    return (ri - 1) // 5


def update_table_data(doc, tuan_so):
    """
    Cap nhat mon hoc, PPCT, ten bai, do dung cho cac hang co du lieu.
    - Lop 5,6,7,8: ap dung rotation + PPCT rieng
    - Tat ca: cap nhat PPCT + ten bai tu PPCT_DATA
    """
    if tuan_so == 1:
        return

    is_even = (tuan_so % 2 == 0)
    
    for ti in [1, 3]:
        tbl = doc.tables[ti]
        
        # First pass: identify rotation classes with multiple periods on same day
        # Build a map: (day_idx, lop) -> list of row indices
        day_lop_rows = {}
        for ri in range(1, len(tbl.rows)):
            row = tbl.rows[ri]
            lop = row.cells[4].text.strip()
            if not lop:
                continue
            day_idx = get_day_idx_from_row(ri)
            key = (day_idx, lop)
            if key not in day_lop_rows:
                day_lop_rows[key] = []
            day_lop_rows[key].append(ri)
        
        # Second pass: update each row
        for ri in range(1, len(tbl.rows)):
            row = tbl.rows[ri]
            lop = row.cells[4].text.strip()
            if not lop:
                continue
            
            day_idx = get_day_idx_from_row(ri)
            is_rot = is_rotation_lop(lop)
            
            # Determine the subject for this row
            if is_rot:
                mon = rotation_mon(tuan_so)
                dd = rotation_do_dung(tuan_so)
                set_cell_text(row.cells[3], mon)
                set_cell_text(row.cells[6], dd)
            else:
                mon = row.cells[3].text.strip()
                # Keep original mon (Tin hoc or Robotics)
            
            # Compute PPCT
            if is_rot:
                # Find period_index for this row within same (day, lop) group
                key = (day_idx, lop)
                rows_for_this = day_lop_rows.get(key, [ri])
                period_index = rows_for_this.index(ri) if ri in rows_for_this else 0
                ppct = compute_rotation_ppct(tuan_so, day_idx, is_even, period_index)
            else:
                ppct = compute_ppct(tuan_so, day_idx)
            
            # Set PPCT
            set_cell_text(row.cells[2], str(ppct))
            
            # Set lesson name
            lesson = get_ppct_lesson(lop, mon, ppct)
            if lesson:
                set_cell_text(row.cells[5], lesson)
            elif ppct == 0:
                set_cell_text(row.cells[5], 'Tiết 0: Định hướng môn học')


def update_ky_ten(doc, start_date):
    """Cap nhat ngay trong bang ky ten."""
    ngay_soan = start_date - timedelta(days=3)
    ngay_str = (f'Ngày {ngay_soan.day} tháng '
                f'{ngay_soan.month} năm {ngay_soan.year}')
    for ti in [2, 4]:
        if ti >= len(doc.tables):
            continue
        for row in doc.tables[ti].rows:
            for cell in row.cells:
                txt = cell.text
                if 'Ngày' in txt and 'tháng' in txt and 'năm' in txt:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if 'Ngày' in r.text and 'tháng' in r.text:
                                m = re.search(
                                    r'Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
                                    r.text)
                                if m:
                                    r.text = r.text[:m.start()] + ngay_str + r.text[m.end():]


# ─────────────────────────────────────────────────────────────────────────────
# PAGE BREAKS
# ─────────────────────────────────────────────────────────────────────────────

def add_page_break_after_table(tbl_element):
    """Them paragraph co page break ngay sau table."""
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    tbl_element.addnext(p)


def fix_page_breaks(doc):
    """
    Xoa empty paragraph thua giua cac section.
    Chen page break sau bang ky ten sang va chieu.
    Ket qua: 3 trang (sang | chieu | nhan xet BGH).
    """
    body = doc.element.body
    children = list(body)

    tbls = [c for c in children
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'tbl']

    sign_sang = tbls[2]   # Bang ky ten sang
    sign_chieu = tbls[4]  # Bang ky ten chieu

    def find_para(keyword):
        for c in list(body):
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'p':
                txt = ''.join(t.text or '' for t in c.findall('.//' + qn('w:t')))
                if keyword.lower() in txt.lower():
                    return c
        return None

    chieu_el = find_para('chiều')
    nhanxet_el = find_para('nhận xét của bgh')

    def remove_empty_between(el_a, el_b):
        ch = list(body)
        if el_a not in ch or el_b not in ch:
            return 0
        ia, ib = ch.index(el_a), ch.index(el_b)
        removed = 0
        for c in ch[ia + 1:ib]:
            if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'p':
                if not ''.join(t.text or '' for t in c.findall('.//' + qn('w:t'))).strip():
                    body.remove(c)
                    removed += 1
        return removed

    if chieu_el is not None:
        n = remove_empty_between(sign_sang, chieu_el)
        if n:
            print(f'   Xoa {n} para trong (sau ky ten sang)')
    if nhanxet_el is not None:
        n = remove_empty_between(sign_chieu, nhanxet_el)
        if n:
            print(f'   Xoa {n} para trong (sau ky ten chieu)')

    add_page_break_after_table(sign_sang)
    add_page_break_after_table(sign_chieu)
    print('   Page breaks: sang | chieu | nhan xet BGH')


# ─────────────────────────────────────────────────────────────────────────────
# TACH TTH+TH / THCS
# ─────────────────────────────────────────────────────────────────────────────

def clear_row_data(row):
    for ci in range(2, min(7, len(row.cells))):
        for p in row.cells[ci].paragraphs:
            for r in p.runs:
                r.text = ''
            if not p.runs and p.text.strip():
                p.text = ''


def filter_for_cap(doc, keep_cap):
    for ti in [1, 3]:
        for ri in range(1, len(doc.tables[ti].rows)):
            row = doc.tables[ti].rows[ri]
            lop = row.cells[4].text.strip()
            cap = classify_lop(lop)
            if cap is not None and cap != keep_cap:
                clear_row_data(row)


# ─────────────────────────────────────────────────────────────────────────────
# HAM CHINH
# ─────────────────────────────────────────────────────────────────────────────

def generate_lbg(tuan_so):
    start_date, end_date = week_dates(tuan_so)
    loai = 'CHẴN → Tin học' if tuan_so % 2 == 0 else 'LẺ → Robotics'

    print(f'\n📅 Tuan {tuan_so:02d}: {fmt_date(start_date)} -> {fmt_date(end_date)}')
    print(f'   Lop 5,6,7,8: {loai}')

    # B1: Load template
    doc = Document(TEMPLATE)

    # B2: Cap nhat tieu de
    update_headers(doc, tuan_so, start_date, end_date)
    print('   Cap nhat tieu de + ngay')

    # B3: Cap nhat nhan Thu/ngay
    update_day_labels(doc, start_date)
    print('   Cap nhat nhan Thu/ngay')

    # B4: Rotation + PPCT + Ten bai
    update_table_data(doc, tuan_so)
    print(f'   Rotation + PPCT + Ten bai updated')

    # B5: Ky ten
    update_ky_ten(doc, start_date)

    # B6: Luu ban goc
    main_name = f'Lịch báo giảng - Tuần {tuan_so:02d}.docx'
    main_path = os.path.join(LBG_DIR, main_name)
    saved_main = save_safe(doc, main_path)
    print(f'   Ban goc: {os.path.basename(saved_main)}')

    # B7+B8: Tach TTH+TH va THCS
    for keep_cap, suffix, label in [
        ('TTH_TH', 'TTH+TH', 'Tien TH + TH'),
        ('THCS', 'THCS', 'THCS'),
    ]:
        d2 = Document(saved_main)
        filter_for_cap(d2, keep_cap)
        fix_page_breaks(d2)

        split_name = f'Lịch báo giảng - Tuần {tuan_so:02d} ({suffix}).docx'
        split_path = os.path.join(LBG_DIR, split_name)
        saved = save_safe(d2, split_path)

        # Kiem tra
        dc = Document(saved)
        cs = sum(1 for ri in range(1, len(dc.tables[1].rows))
                 if dc.tables[1].rows[ri].cells[4].text.strip())
        cc = sum(1 for ri in range(1, len(dc.tables[3].rows))
                 if dc.tables[3].rows[ri].cells[4].text.strip())
        print(f'   {label}: sang={cs} tiet, chieu={cc} tiet -> {os.path.basename(saved)}')

    print(f'\n HOAN TAT Tuan {tuan_so:02d}! Thu muc: {LBG_DIR}')


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Cach dung: python generate_lbg.py <so_tuan>')
        print('Vi du:     python generate_lbg.py 2')
        sys.exit(1)
    try:
        tuan = int(sys.argv[1])
    except ValueError:
        print(f'Loi: "{sys.argv[1]}" khong phai so tuan hop le.')
        sys.exit(1)
    if not (1 <= tuan <= 52):
        print('Loi: So tuan phai tu 1 den 52.')
        sys.exit(1)
    generate_lbg(tuan)
