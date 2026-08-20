"""
generate_lbg.py — Tu dong tao Lich bao giang theo tuan tu TKB Giao vien
========================================================================

QUY TAC ROTATION:
  - Lop 5, 6, 7, 8: Moi lop co 2 tiet / tuan
      Tuan CHAN (2, 4, 6...): ca 2 tiet -> Tin hoc (Phong Tin hoc)
      Tuan LE  (3, 5, 7...): ca 2 tiet -> Robotics (Bo Kit Robotics)
  - Cac lop con lai: giu nguyen mon hoc goc tu TKB

QUY TAC PPCT (Khong offset):
  - Tuan 1: Tat ca = Tiet 0 (Dinh huong)
  - Tuan 2+: Tat ca cac Thu (T2 -> T6) deu dong bo ppct = tuan_so - 1
  - Rotation classes (5,6,7,8): PPCT rieng cho Tin/Robotics
      Khi co 2 tiet lien tiep cung mon -> PPCT lien tiep (Tuan 2: PPCT 1,2; Tuan 4: PPCT 3,4...)
      Tuan 3 (Le - Robotics): PPCT 1, 2; Tuan 5: PPCT 3, 4...

MASTER SCHEDULE (Tu 'Thoi khoa bieu - Dau Dinh Nguyen.xlsx'):
  SANG (13 tiet):
    - Thu Hai: T3 Rob 4C1, T4 Tin 1A1
    - Thu Ba: T1 Tin 1C1, T3-T4 Tin-Rob 5C1
    - Thu Tu: T2 Rob 3A1, T4 Tin 3C1
    - Thu Nam: T1 Tin TT3, T2 Rob 1A1, T4 Rob 2C1
    - Thu Sau: T1 Rob 3C1, T4-T5 Tin-Rob 6A1
  CHIEU (12 tiet):
    - Thu Hai: T3 Tin 2C1
    - Thu Ba: T1 Tin 2A1, T3-T4 Tin-Rob 7A1
    - Thu Tu: T3 Tin 4C1, T4 Rob 2A1
    - Thu Nam: T1 Tin 3A1, T2 Tin TTH 1, T3 Tin TTH2, T4 Rob 1C1
    - Thu Sau: T1-T2 Tin-Rob 8A1
"""

import sys
import os
import re
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import date, timedelta

sys.stdout.reconfigure(encoding='utf-8')

LBG_DIR = r'D:\UNIGO\Hệ thống mẫu văn bản\Nguyên đã làm\Lịch báo giảng'
TEMPLATE = os.path.join(LBG_DIR, 'Lịch báo giảng - Tuần 01.docx')
TUAN_01_START = date(2026, 8, 3)

ROTATION_PREFIXES = ('5', '6', '7', '8')
THU_LABELS = {0: 'Hai', 1: 'Ba', 2: 'Tư', 3: 'Năm', 4: 'Sáu'}

# Tên tổ trưởng theo cấp học
TO_TRUONG_THCS = 'Nguyễn Thị Ngọc Ánh'
TO_TRUONG_TH = 'Nguyễn Thị Ngọc'

# Số dòng chấm nhận xét tối đa trong bảng chữ ký (giữ gọn 1 trang)
MAX_DONG_CHAM = 4

PPCT_TIN = {
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

PPCT_ROB = {
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

PPCT_ROB['6'] = PPCT_ROB['5'].copy()
PPCT_ROB['7'] = PPCT_ROB['5'].copy()
PPCT_ROB['8'] = PPCT_ROB['5'].copy()
PPCT_ROB['TT'] = PPCT_ROB['1'].copy()


# ─────────────────────────────────────────────────────────────────────────────
# MASTER SCHEDULE (Tu 'Thoi khoa bieu - Dau Dinh Nguyen.xlsx')
# (day_idx 0..4, tiet_tkb 1..5): (lop, mon_display, is_rotation)
# ─────────────────────────────────────────────────────────────────────────────

MASTER_SCHEDULE_SANG = {
    (0, 3): ('4C1', 'Robotics', False),
    (0, 4): ('1A1', 'Tin học', False),
    (1, 1): ('1C1', 'Tin học', False),
    (1, 3): ('5C1', 'Tin - Robotics', True),
    (1, 4): ('5C1', 'Tin - Robotics', True),
    (2, 2): ('3A1', 'Robotics', False),
    (2, 4): ('3C1', 'Tin học', False),
    (3, 1): ('TT3', 'Tin học', False),
    (3, 2): ('1A1', 'Robotics', False),
    (3, 4): ('2C1', 'Robotics', False),
    (4, 1): ('3C1', 'Robotics', False),
    (4, 4): ('6A1', 'Tin - Robotics', True),
    (4, 5): ('6A1', 'Tin - Robotics', True),
}

MASTER_SCHEDULE_CHIEU = {
    (0, 3): ('2C1', 'Tin học', False),
    (1, 1): ('2A1', 'Tin học', False),
    (1, 3): ('7A1', 'Tin - Robotics', True),
    (1, 4): ('7A1', 'Tin - Robotics', True),
    (2, 3): ('4C1', 'Tin học', False),
    (2, 4): ('2A1', 'Robotics', False),
    (3, 1): ('3A1', 'Tin học', False),
    (3, 2): ('TTH 1', 'Tin học', False),
    (3, 3): ('TTH2', 'Tin học', False),
    (3, 4): ('1C1', 'Robotics', False),
    (4, 1): ('8A1', 'Tin - Robotics', True),
    (4, 2): ('8A1', 'Tin - Robotics', True),
}


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
    s = lop.strip().upper()
    if not s:
        return None
    if re.match(r'^[6789]', s):
        return 'THCS'
    return 'TTH_TH'


def is_rotation_lop(lop):
    s = lop.strip().upper()
    return any(s.startswith(p) for p in ROTATION_PREFIXES)


def rotation_mon(tuan_so):
    return 'Tin học' if tuan_so % 2 == 0 else 'Robotics'


def rotation_do_dung(tuan_so):
    return 'Phòng Tin học' if tuan_so % 2 == 0 else 'Bộ Kit Robotics'


def get_grade_key(lop):
    s = lop.strip().upper()
    if s.startswith('TT'):
        return 'TT3'
    m = re.match(r'^(\d)', s)
    if m:
        return m.group(1)
    return None


def get_ppct_lesson(lop, mon, ppct_num):
    grade = get_grade_key(lop)
    if grade is None:
        return ''
    if mon == 'Tin học':
        data = PPCT_TIN.get(grade, {})
    elif mon == 'Robotics':
        data = PPCT_ROB.get(grade, {})
    else:
        return ''
    if not data:
        return ''
    if ppct_num in data:
        return data[ppct_num]
    max_k = max(data.keys())
    if ppct_num > max_k:
        return data[max_k]
    return ''


def compute_ppct(tuan_so, day_idx=0):
    if tuan_so <= 1:
        return 0
    return tuan_so - 1


def compute_rotation_ppct(tuan_so, day_idx, is_even_week, period_index_in_day):
    if tuan_so <= 1:
        return 0
    if is_even_week:
        if tuan_so < 2:
            return 0
        k = (tuan_so - 2) // 2 + 1
        return (k - 1) * 2 + period_index_in_day + 1
    else:
        if tuan_so < 3:
            return 0
        k = (tuan_so - 3) // 2 + 1
        return (k - 1) * 2 + period_index_in_day + 1


def set_cell_text(cell, text):
    for p in cell.paragraphs:
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ''
            if p.runs[0].font.name != 'Times New Roman':
                p.runs[0].font.name = 'Times New Roman'
            if p.runs[0].font.size != Pt(13):
                p.runs[0].font.size = Pt(13)
            return
    p = cell.paragraphs[0]
    p.clear()
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)


def save_safe(doc, path):
    try:
        doc.save(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        alt = base + '_v2' + ext
        doc.save(alt)
        print(f'   ⚠ Bi khoa, luu: {os.path.basename(alt)}')
        return alt


def update_headers(doc, tuan_so, start_date, end_date):
    tuan_str = f'TUẦN {tuan_so:02d}'
    fs = fmt_date(start_date)
    fe = fmt_date(end_date)
    ngay_soan = start_date - timedelta(days=3)

    for p in doc.paragraphs:
        txt = p.text.strip()
        if 'TUẦN' in txt.upper() and 'từ ngày' in txt:
            new = f'{tuan_str} (từ ngày {fs} đến ngày {fe})'
        elif re.match(r'^Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+', txt) and len(txt) < 45:
            new = f'Ngày {ngay_soan.day:02d} tháng {ngay_soan.month:02d} năm {ngay_soan.year}'
        elif 'sáng' in txt.lower() and 'tuần' in txt.lower():
            new = f'Buổi…sáng……..Tuần…{tuan_so:02d}…(Từ ngày…{fs} …đến ngày:… {fe}….)'
        elif 'chiều' in txt.lower() and 'tuần' in txt.lower():
            new = f'Buổi…chiều…..Tuần…{tuan_so:02d}…(Từ ngày…{fs}…đến ngày:… {fe}….)'
        else:
            continue

        if p.runs:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ''
        else:
            p.text = new


def populate_lbg_table(tbl, schedule_map, tuan_so, start_date):
    is_even = (tuan_so % 2 == 0)
    for day_idx in range(5):
        d = start_date + timedelta(days=day_idx)
        d_lbl = day_label(d)
        for tiet in range(1, 6):
            ri = day_idx * 5 + tiet
            if ri >= len(tbl.rows):
                continue
            row = tbl.rows[ri]
            set_cell_text(row.cells[0], d_lbl)
            set_cell_text(row.cells[1], str(tiet))
            key = (day_idx, tiet)
            if key in schedule_map:
                lop, mon_display, is_rot = schedule_map[key]
                if is_rot:
                    mon_int = rotation_mon(tuan_so)
                    dd = rotation_do_dung(tuan_so)
                    mon_lbl = 'Tin - Robotics'
                    p_idx = 0 if (day_idx, tiet - 1) not in schedule_map else 1
                    ppct = compute_rotation_ppct(tuan_so, day_idx, is_even, p_idx)
                    lesson = get_ppct_lesson(lop, mon_int, ppct) if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                else:
                    mon_lbl = mon_display
                    dd = 'Phòng Tin học' if 'Tin' in mon_lbl else 'Bộ Kit Robotics'
                    ppct = compute_ppct(tuan_so, day_idx) if tuan_so > 1 else 0
                    lesson = get_ppct_lesson(lop, mon_lbl, ppct) if tuan_so > 1 else 'Tiết 0: Định hướng môn học'
                ppct_str = str(ppct) if tuan_so > 1 else '1'
                if not lesson and ppct == 0:
                    lesson = 'Tiết 0: Định hướng môn học'
                set_cell_text(row.cells[2], ppct_str)
                set_cell_text(row.cells[3], mon_lbl)
                set_cell_text(row.cells[4], lop)
                set_cell_text(row.cells[5], lesson)
                set_cell_text(row.cells[6], dd)
            else:
                set_cell_text(row.cells[2], '')
                set_cell_text(row.cells[3], '')
                set_cell_text(row.cells[4], '')
                set_cell_text(row.cells[5], '')
                set_cell_text(row.cells[6], '')


def update_table_data(doc, tuan_so, start_date=None):
    if start_date is None:
        start_date, _ = week_dates(tuan_so)
    lbg_tbls = [t for t in doc.tables
                if len(t.columns) == 7
                and 'Lớp' in [c.text.strip() for c in t.rows[0].cells]]
    if len(lbg_tbls) >= 1:
        populate_lbg_table(lbg_tbls[0], MASTER_SCHEDULE_SANG, tuan_so, start_date)
    if len(lbg_tbls) >= 2:
        populate_lbg_table(lbg_tbls[1], MASTER_SCHEDULE_CHIEU, tuan_so, start_date)


def update_ky_ten(doc, start_date):
    ngay_soan = start_date - timedelta(days=3)
    ngay_str = f'Ngày {ngay_soan.day} tháng {ngay_soan.month} năm {ngay_soan.year}'
    # Update all sign tables (1r x 2c)
    for tbl in doc.tables:
        if len(tbl.rows) != 1 or len(tbl.columns) != 2:
            continue
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text
                if 'Ngày' in txt and 'tháng' in txt and 'năm' in txt:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if 'Ngày' in r.text and 'tháng' in r.text:
                                m = re.search(r'Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+', r.text)
                                if m:
                                    r.text = r.text[:m.start()] + ngay_str + r.text[m.end():]


def update_sign_names(doc, to_truong_name):
    """Insert tổ trưởng name below 'Tổ trưởng' in all sign tables (1r x 2c).

    Sign table cell [0,1] run structure:
      Run 0: 'Ngày X tháng Y năm Z'
      Run 1: '\nTổ trưởng'
      Run 2: '\n' or '\n(Ký tên, đóng dấu)'
    We insert the name by modifying Run 1 to append '\n<name>'
    or adding a new run after Run 1.
    """
    for tbl in doc.tables:
        if len(tbl.rows) != 1 or len(tbl.columns) != 2:
            continue
        cell = tbl.rows[0].cells[1]  # Right cell (Tổ trưởng)
        for p in cell.paragraphs:
            runs = p.runs
            if not runs:
                continue
            # Find the run containing 'Tổ trưởng'
            for ri, r in enumerate(runs):
                if 'Tổ trưởng' in r.text:
                    # Check if name is already present
                    if to_truong_name in cell.text:
                        break
                    # Append name below 'Tổ trưởng'
                    r.text = r.text.rstrip() + '\n' + to_truong_name
                    break


def compact_sign_tables(doc):
    """Reduce dòng chấm nhận xét in sign tables to MAX_DONG_CHAM lines
    to ensure the LBG table + sign section fit within 1 page.

    Sign table cell [0,0] run structure:
      Run 0: 'Kiểm tra, nhận xét'
      Run 1..N: '\n…………………………………………………………………'
    We keep at most MAX_DONG_CHAM dòng chấm runs.
    """
    for tbl in doc.tables:
        if len(tbl.rows) != 1 or len(tbl.columns) != 2:
            continue
        cell = tbl.rows[0].cells[0]  # Left cell (Kiểm tra, nhận xét)
        for p in cell.paragraphs:
            runs = p.runs
            if not runs:
                continue
            # Find runs with dòng chấm
            dot_runs = []
            for ri, r in enumerate(runs):
                if '……' in r.text:
                    dot_runs.append((ri, r))
            # Keep only MAX_DONG_CHAM dòng chấm runs
            if len(dot_runs) > MAX_DONG_CHAM:
                for ri, r in dot_runs[MAX_DONG_CHAM:]:
                    r.text = ''

    # Also remove excessive empty paragraphs between tables
    body = doc.element.body
    children = list(body)
    for i in range(len(children) - 1):
        c = children[i]
        tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
        if tag == 'tbl':
            # Check if this is a sign table
            is_sign = False
            for tbl in doc.tables:
                if tbl._element is c and len(tbl.rows) == 1 and len(tbl.columns) == 2:
                    is_sign = True
                    break
            if is_sign:
                # Remove excessive empty paragraphs after sign table
                # Keep at most 0 empty paragraphs
                j = i + 1
                empty_count = 0
                while j < len(children):
                    nc = children[j]
                    nc_tag = nc.tag.split('}')[-1] if '}' in nc.tag else nc.tag
                    if nc_tag != 'p':
                        break
                    txt = ''.join(t.text or '' for t in nc.findall('.//' + qn('w:t')))
                    if txt.strip():
                        break
                    empty_count += 1
                    j += 1
                # Remove all empty paragraphs after sign table
                if empty_count > 0:
                    for nc in children[i+1:i+1+empty_count]:
                        nc_tag = nc.tag.split('}')[-1] if '}' in nc.tag else nc.tag
                        if nc_tag == 'p' and nc_tag != 'sectPr':
                            txt = ''.join(t.text or '' for t in nc.findall('.//' + qn('w:t')))
                            if not txt.strip():
                                body.remove(nc)


def remove_second_copy(doc):
    body = doc.element.body
    children = list(body)
    tuan_paras = []
    for c in children:
        tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
        if tag == 'p':
            txt = ''.join(t.text or '' for t in c.findall('.//' + qn('w:t')))
            if 'TUẦN' in txt.upper() and 'từ ngày' in txt:
                tuan_paras.append(c)
    if len(tuan_paras) < 2:
        return 0
    nhanxet_el = None
    for c in children:
        tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
        if tag == 'p':
            txt = ''.join(t.text or '' for t in c.findall('.//' + qn('w:t')))
            if 'nhận xét' in txt.lower() and 'bgh' in txt.lower():
                nhanxet_el = c
                break
    if nhanxet_el is None:
        return 0
    children = list(body)
    start_idx = children.index(tuan_paras[1])
    end_idx = children.index(nhanxet_el)
    removed = 0
    for el in children[start_idx:end_idx]:
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if tag != 'sectPr':
            body.remove(el)
            removed += 1
    return removed


def fix_all_fonts(doc):
    fixed = 0
    for tbl in doc.tables:
        if len(tbl.columns) != 7:
            continue
        header_texts = [c.text.strip() for c in tbl.rows[0].cells]
        if 'Lớp' not in header_texts:
            continue
        for ri in range(0, len(tbl.rows)):
            for ci in range(len(tbl.rows[ri].cells)):
                cell = tbl.rows[ri].cells[ci]
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.font.name != 'Times New Roman':
                            run.font.name = 'Times New Roman'
                            fixed += 1
                        if run.font.size is None or run.font.size != Pt(13):
                            run.font.size = Pt(13)
    return fixed


def add_page_break_after_table(tbl_element):
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    tbl_element.addnext(p)


def fix_page_breaks(doc):
    body = doc.element.body
    children = list(body)
    sign_tables = []
    for c in children:
        tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
        if tag == 'tbl':
            for tbl in doc.tables:
                if tbl._element is c and len(tbl.rows) == 1 and len(tbl.columns) == 2:
                    sign_tables.append(c)
                    break
    if len(sign_tables) < 2:
        return
    sign_sang = sign_tables[0]
    sign_chieu = sign_tables[1]

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
        remove_empty_between(sign_sang, chieu_el)
    if nhanxet_el is not None:
        remove_empty_between(sign_chieu, nhanxet_el)

    add_page_break_after_table(sign_sang)
    add_page_break_after_table(sign_chieu)


def clear_row_data(row):
    for ci in range(2, min(7, len(row.cells))):
        for p in row.cells[ci].paragraphs:
            for r in p.runs:
                r.text = ''
            if not p.runs and p.text.strip():
                p.text = ''


def filter_for_cap(doc, keep_cap):
    for tbl in doc.tables:
        if len(tbl.columns) != 7:
            continue
        header_texts = [c.text.strip() for c in tbl.rows[0].cells]
        if 'Lớp' not in header_texts:
            continue
        for ri in range(1, len(tbl.rows)):
            row = tbl.rows[ri]
            lop = row.cells[4].text.strip()
            cap = classify_lop(lop)
            if cap is not None and cap != keep_cap:
                clear_row_data(row)


def generate_lbg(tuan_so):
    start_date, end_date = week_dates(tuan_so)
    loai = 'CHẴN → Tin học' if tuan_so % 2 == 0 else 'LẺ → Robotics'
    print(f'📅 Tuần {tuan_so:02d}: {fmt_date(start_date)} -> {fmt_date(end_date)} [{loai}]')

    doc = Document(TEMPLATE)
    remove_second_copy(doc)
    update_headers(doc, tuan_so, start_date, end_date)
    update_table_data(doc, tuan_so, start_date)
    fix_all_fonts(doc)
    update_ky_ten(doc, start_date)
    compact_sign_tables(doc)

    main_name = f'Lịch báo giảng - Tuần {tuan_so:02d}.docx'
    main_path = os.path.join(LBG_DIR, main_name)
    saved_main = save_safe(doc, main_path)

    for keep_cap, suffix, label, to_truong in [
        ('TTH_TH', 'TTH+TH', 'TTH + TH', TO_TRUONG_TH),
        ('THCS', 'THCS', 'THCS', TO_TRUONG_THCS),
    ]:
        d2 = Document(saved_main)
        filter_for_cap(d2, keep_cap)
        update_sign_names(d2, to_truong)
        compact_sign_tables(d2)
        fix_page_breaks(d2)
        split_name = f'Lịch báo giảng - Tuần {tuan_so:02d} ({suffix}).docx'
        split_path = os.path.join(LBG_DIR, split_name)
        save_safe(d2, split_path)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Cách dùng: python generate_lbg.py <so_tuan> [den_tuan]')
        print('Ví dụ:     python generate_lbg.py 3')
        print('           python generate_lbg.py 3 35')
        sys.exit(1)
    try:
        t_start = int(sys.argv[1])
        t_end = int(sys.argv[2]) if len(sys.argv) > 2 else t_start
    except ValueError:
        print('Lỗi: Số tuần phải là số nguyên.')
        sys.exit(1)

    for w in range(t_start, t_end + 1):
        generate_lbg(w)
    print(f'\\n✅ Đã tạo xong từ tuần {t_start:02d} đến {t_end:02d}!')
