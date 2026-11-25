# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Cuối học kì 1 (Đánh giá định kỳ 2) & Đề cương 80 câu ôn tập
Môn: Tin học 7 (Bộ sách Kết nối tri thức với cuộc sống - UNIGO)
Phạm vi: Bài 1 đến Bài 8 (Tiết 1), trang 5 - 43 SGK Tin học 7.
"""

GRADE_DATA = {
    "lop": 7,
    "cap_hoc": "THCS",
    "lop_hien_thi": "7A1",
    "thoi_gian": "45 phút",
    "tuan_on_tap": 18,
    "tuan_kiem_tra": 20,
    "tiet_on_tap": 18,
    "tiet_kiem_tra": 19,

    "tom_tat_ly_thuyet": [
        {
            "bai": "Bài 1. Thiết bị vào – ra",
            "y_chinh": [
                "Thiết bị vào (Input): Tiếp nhận thông tin từ thế giới bên ngoài (văn bản, hình ảnh, âm thanh, thao tác điều khiển) đưa vào máy tính xử lí; gồm bàn phím, chuột, micro, máy quét (scanner), webcam, đầu đọc mã vạch.",
                "Thiết bị ra (Output): Đưa thông tin đã xử lí từ máy tính ra ngoài cho con người nhận biết; gồm màn hình thường, loa, tai nghe, máy in (printer), máy chiếu (projector).",
                "Thiết bị vừa vào vừa ra: Màn hình cảm ứng (touchscreen) trên điện thoại thông minh, máy tính bảng, cây ATM vừa hiển thị thông tin ra vừa nhận thao tác chạm đầu vào.",
                "An toàn thiết bị: Cắm đúng cổng kết nối phù hợp, thao tác nhẹ nhàng, không tự ý cắm dây nguồn khi tay ướt hoặc giật mạnh cáp khi máy đang chạy."
            ],
            "image_key": "g7_thiet_bi_va_phan_mem",
        },
        {
            "bai": "Bài 2. Phần mềm máy tính",
            "y_chinh": [
                "Hệ điều hành (Operating System): Phần mềm hệ thống quan trọng nhất quản lí toàn bộ phần cứng, điều phối hoạt động và tạo môi trường chạy các phần mềm ứng dụng (ví dụ: Windows, macOS, Linux trên máy tính; Android, iOS trên thiết bị di động).",
                "Phần mềm ứng dụng: Chương trình giúp người dùng thực hiện công việc cụ thể như soạn thảo Word, bảng tính Excel, trình chiếu PowerPoint, lập trình Scratch, trình duyệt Chrome.",
                "Phần mở rộng tên tệp (đuôi tệp): Cho hệ điều hành biết loại dữ liệu và phần mềm mở tương ứng, phân cách với phần tên bằng dấu chấm (.): .docx (Word), .xlsx (Excel), .pptx (PowerPoint), .sb3 (Scratch), .jpg/.png (ảnh), .mp3 (âm thanh), .mp4 (video), .zip/.rar (tệp nén)."
            ],
        },
        {
            "bai": "Bài 3. Quản lí dữ liệu trong máy tính",
            "y_chinh": [
                "Tổ chức cây thư mục: Đặt tên tệp, thư mục ngắn gọn, rõ nghĩa, theo cấu trúc phân cấp khoa học để tra cứu nhanh chóng.",
                "Các thao tác và phím tắt thông dụng trong File Explorer: Sao chép (Ctrl + C), Cắt/Di chuyển (Ctrl + X), Dán (Ctrl + V), Đổi tên (F2), Xóa vào thùng rác (Delete).",
                "Bảo vệ dữ liệu: Sao lưu dự phòng (Backup) sang ổ cứng ngoài, USB hoặc đám mây; dùng phần mềm diệt virus có bản quyền, đặt mật khẩu mạnh và không mở tệp lạ từ nguồn không rõ ràng."
            ],
        },
        {
            "bai": "Bài 4. Mạng xã hội và một số kênh trao đổi thông tin trên Internet",
            "y_chinh": [
                "Mạng xã hội (Facebook, Zalo, YouTube, Instagram...): Ứng dụng web kết nối cộng đồng có cùng sở thích hoặc mối quan hệ, cho phép tạo hồ sơ cá nhân, đăng bài, chia sẻ đa phương tiện, bình luận và nhắn tin trực tuyến.",
                "Các kênh trao đổi thông tin khác: Thư điện tử (Email), diễn đàn trực tuyến (Forum), trang tin điện tử, ứng dụng hội nghị truyền hình (Zoom, Google Meet, Microsoft Teams).",
                "Hai mặt của mạng xã hội: Giúp giao lưu, học tập nhanh chóng nhưng tiềm ẩn nguy cơ tin giả, lừa đảo trực tuyến, lộ thông tin cá nhân và mã độc."
            ],
        },
        {
            "bai": "Bài 5. Ứng xử trên mạng",
            "y_chinh": [
                "Giao tiếp văn minh trên mạng: Sử dụng ngôn từ lịch sự, tôn trọng người khác, xin phép trước khi đăng ảnh/thông tin riêng tư của bạn bè, kiểm chứng trước khi chia sẻ.",
                "Cách xử lí khi bị bắt nạt qua mạng (Cyberbullying): Giữ bình tĩnh, KHÔNG nhắn tin chửi bới trả đũa; chụp màn hình lưu giữ bằng chứng; chặn (Block) và báo cáo tài khoản xấu; lập tức chia sẻ và nhờ cha mẹ, thầy cô hoặc cơ quan chức năng hỗ trợ.",
                "Phòng tránh nghiện Internet: Sắp xếp thời gian sử dụng máy tính hợp lí, ưu tiên đọc sách, thể thao và giao tiếp trực tiếp cùng gia đình, bạn bè."
            ],
        },
        {
            "bai": "Bài 6. Làm quen với phần mềm bảng tính",
            "y_chinh": [
                "Giao diện trang tính (Excel): Gồm các cột kí hiệu bằng chữ cái (A, B, C...), các hàng đánh số (1, 2, 3...). Giao của cột và hàng là ô tính, địa chỉ ô viết dạng <Tên cột><Tên hàng> (ví dụ: C5).",
                "Vùng dữ liệu (khối ô): Tập hợp các ô liền kề tạo thành hình chữ nhật, viết dạng <Ô góc trái trên>:<Ô góc phải dưới> (ví dụ: B3:C6).",
                "Hộp tên (Name Box) hiển thị địa chỉ ô đang chọn; Thanh công thức (Formula Bar) dùng để nhập, hiển thị và chỉnh sửa dữ liệu hoặc công thức trong ô.",
                "Mặc định căn lề dữ liệu trong ô: Dữ liệu văn bản (chữ) tự động căn thẳng lề trái; dữ liệu số, ngày tháng tự động căn thẳng lề phải."
            ],
        },
        {
            "bai": "Bài 7. Tính toán tự động trên bảng tính",
            "y_chinh": [
                "Kí hiệu phép toán trong Excel: Cộng (+), Trừ (-), Nhân (*), Chia (/), Lũy thừa (^), Phần trăm (%). Mọi công thức tính toán bắt buộc bắt đầu bằng dấu bằng (=).",
                "Tính toán tự động với địa chỉ ô: Sử dụng địa chỉ ô (ví dụ: =B3+C3) thay vì hằng số giúp kết quả tự động cập nhật ngay khi dữ liệu nguồn thay đổi.",
                "Sao chép công thức: Khi sao chép công thức chứa địa chỉ tương đối (bằng Ctrl+C / Ctrl+V hoặc kéo nút điền tự động Fill Handle), địa chỉ các ô trong công thức tự động thay đổi tương ứng theo vị trí hàng và cột mới."
            ],
        },
        {
            "bai": "Bài 8. Công cụ hỗ trợ tính toán (Tiết 1)",
            "y_chinh": [
                "Hàm trong bảng tính: Là công thức được định nghĩa sẵn giúp tính toán nhanh gọn, chính xác trên vùng dữ liệu lớn.",
                "Cú pháp chung của hàm: =<Tên_hàm>(<các tham số>), trong đó các tham số có thể là số, địa chỉ ô hoặc vùng dữ liệu phân cách nhau bởi dấu phẩy (,) hoặc chấm phẩy (;).",
                "5 hàm thống kê cơ bản: SUM (tính tổng), AVERAGE (tính trung bình cộng), MAX (tìm giá trị lớn nhất), MIN (tìm giá trị nhỏ nhất), COUNT (đếm số ô chứa dữ liệu kiểu số)."
            ],
            "image_key": "g7_bang_tinh_excel_mau",
        },
    ],

    "ma_tran_kiem_tra": [
        [
            "Chủ đề 1. Máy tính và cộng đồng (Bài 1, 2, 3)",
            "8 câu (2,0đ)",
            "3 câu (0,75đ)",
            "—",
            "—",
            "—",
            "2,75 điểm",
        ],
        [
            "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin (Bài 4)",
            "3 câu (0,75đ)",
            "1 câu (0,25đ)",
            "—",
            "—",
            "—",
            "1,0 điểm",
        ],
        [
            "Chủ đề 3. Đạo đức, pháp luật và văn hóa trong môi trường số (Bài 5)",
            "2 câu (0,5đ)",
            "1 câu (0,25đ)",
            "1 câu (1,0đ)",
            "—",
            "—",
            "1,75 điểm",
        ],
        [
            "Chủ đề 4. Ứng dụng tin học – Phần mềm bảng tính (Bài 6, 7, 8)",
            "7 câu (1,75đ)",
            "3 câu (0,75đ)",
            "—",
            "1 câu (1,0đ)",
            "1 câu (1,0đ)",
            "4,5 điểm",
        ],
        [
            "TỔNG CỘNG",
            "20 câu (5,0đ)",
            "8 câu (2,0đ)",
            "1 câu (1,0đ)",
            "1 câu (1,0đ)",
            "1 câu (1,0đ)",
            "10,0 điểm",
        ],
    ],

    "exam_mcq": [
        # --- 20 CÂU NHẬN BIẾT (NB) ---
        {
            "bai": "Bài 1",
            "level": "NB",
            "q": "Thiết bị nào sau đây thuộc nhóm thiết bị vào (Input) của máy tính?",
            "options": [
                "Bàn phím máy tính",
                "Loa để bàn",
                "Máy in laser",
                "Máy chiếu lớp học",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "level": "NB",
            "q": "Thiết bị nào dưới đây dùng để xuất thông tin âm thanh từ máy tính ra bên ngoài?",
            "options": [
                "Tai nghe hoặc loa",
                "Máy quét hình ảnh",
                "Chuột quang không dây",
                "Đầu đọc mã vạch",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "level": "NB",
            "q": "Bộ phận nào trên máy tính bảng vừa đóng vai trò thiết bị vào vừa đóng vai trò thiết bị ra?",
            "options": [
                "Màn hình cảm ứng",
                "Cổng sạc nguồn điện",
                "Vỏ nhựa bảo vệ máy",
                "Nút nguồn bật tắt",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Phần mềm nào sau đây là hệ điều hành dành cho máy tính cá nhân?",
            "options": [
                "Microsoft Windows",
                "Microsoft Word",
                "Google Chrome",
                "Scratch Desktop",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Tệp chương trình lập trình trực quan được tạo bởi phần mềm Scratch 3.0 thường có phần mở rộng là gì?",
            "options": [
                ".sb3",
                ".xlsx",
                ".mp3",
                ".docx",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Phần mở rộng tên tệp nào dưới đây cho biết đó là tệp hình ảnh?",
            "options": [
                ".png",
                ".mp4",
                ".zip",
                ".pptx",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "level": "NB",
            "q": "Trong cửa sổ quản lí tệp File Explorer, phím tắt nào dùng để đổi tên thư mục đang được chọn?",
            "options": [
                "F2",
                "F5",
                "Ctrl + V",
                "Ctrl + Z",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "level": "NB",
            "q": "Hành động tạo thêm bản sao dữ liệu quan trọng cất giữ ở ổ cứng rời hoặc đám mây được gọi là gì?",
            "options": [
                "Sao lưu dữ liệu dự phòng (Backup)",
                "Giải nén tệp tin đa phương tiện",
                "Đổi phần mở rộng của tệp tin",
                "Gỡ cài đặt phần mềm ứng dụng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "NB",
            "q": "Dịch vụ nào sau đây là một mạng xã hội phổ biến giúp kết nối và chia sẻ thông tin trực tuyến?",
            "options": [
                "Facebook",
                "Microsoft Excel",
                "Windows Explorer",
                "Unikey",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "NB",
            "q": "Công cụ trực tuyến nào phù hợp nhất để tổ chức lớp học trực tuyến có hình ảnh và âm thanh theo thời gian thực?",
            "options": [
                "Google Meet",
                "Microsoft Paint",
                "Calculator",
                "Notepad",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "NB",
            "q": "Dấu hiệu nào cho thấy một trang thông tin trên mạng xã hội tiềm ẩn nguy cơ lừa đảo?",
            "options": [
                "Yêu cầu cung cấp mật khẩu và mã OTP ngân hàng để nhận quà trúng thưởng",
                "Đăng tải thông báo lịch học chính thức từ cổng thông tin nhà trường",
                "Chia sẻ bài giảng ôn tập môn Tin học của giáo viên bộ môn",
                "Cung cấp số điện thoại đường dây nóng bảo vệ trẻ em quốc gia 111",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "NB",
            "q": "Biểu hiện nào dưới đây thể hiện cách giao tiếp văn minh trên không gian mạng?",
            "options": [
                "Sử dụng ngôn từ lịch sự và tôn trọng ý kiến khác biệt của người đối thoại",
                "Tự ý đăng ảnh riêng tư của bạn lên nhóm công khai để trêu chọc",
                "Viết bình luận xúc phạm người khác bằng tài khoản ẩn danh",
                "Chia sẻ ngay các tin đồn giật gân chưa được kiểm chứng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "NB",
            "q": "Ai là người đáng tin cậy mà học sinh cần báo ngay khi bị kẻ xấu đe dọa hoặc bắt nạt qua mạng?",
            "options": [
                "Bố mẹ và thầy cô giáo",
                "Người lạ mới quen trong trò chơi trực tuyến",
                "Tài khoản ẩn danh gửi đường dẫn lạ",
                "Quản trị viên các nhóm cá cược trên mạng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Trong phần mềm bảng tính Excel, giao của cột C và hàng 5 được kí hiệu địa chỉ ô là gì?",
            "options": [
                "C5",
                "5C",
                "C:5",
                "C-5",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Kí hiệu nào dưới đây biểu diễn đúng địa chỉ vùng dữ liệu hình chữ nhật từ ô B3 đến ô C6 trong Excel?",
            "options": [
                "B3:C6",
                "B3-C6",
                "B3;C6",
                "B3..C6",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Theo mặc định trong Excel, khi nhập số 2026 vào một ô tính chưa định dạng thì dữ liệu sẽ tự động căn lề như thế nào?",
            "options": [
                "Căn thẳng lề phải của ô tính",
                "Căn thẳng lề trái của ô tính",
                "Căn chính giữa ô tính",
                "Căn đều hai bên lề ô tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "NB",
            "q": "Để phần mềm bảng tính nhận biết và thực hiện tính toán công thức, kí tự đầu tiên cần gõ vào ô là gì?",
            "options": [
                "Dấu bằng (=)",
                "Dấu cộng (+)",
                "Dấu hai chấm (:)",
                "Dấu sao (*)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "NB",
            "q": "Trong công thức bảng tính Excel, phép nhân và phép chia lần lượt được viết bằng cặp kí hiệu nào?",
            "options": [
                "Dấu * và dấu /",
                "Dấu x và dấu :",
                "Dấu . và dấu :",
                "Dấu ^ và dấu %",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "level": "NB",
            "q": "Hàm nào trong phần mềm bảng tính Excel được dùng để tính trung bình cộng của các giá trị số?",
            "options": [
                "AVERAGE",
                "SUM",
                "COUNT",
                "MAX",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "level": "NB",
            "q": "Trong Excel, hàm MIN có chức năng gì?",
            "options": [
                "Xác định giá trị nhỏ nhất trong danh sách các tham số",
                "Xác định giá trị lớn nhất trong danh sách các tham số",
                "Tính tổng tất cả các giá trị số trong vùng dữ liệu",
                "Đếm số lượng ô chứa dữ liệu văn bản trong vùng chọn",
            ],
            "ans": 0,
        },

        # --- 8 CÂU THÔNG HIỂU (TH) ---
        {
            "bai": "Bài 1",
            "level": "TH",
            "q": "Bạn Minh muốn thu giọng đọc thuyết trình của mình vào máy tính để lồng tiếng cho video. Minh cần kết nối và kiểm tra thiết bị nào?",
            "options": [
                "Micro (Microphone)",
                "Máy chiếu (Projector)",
                "Máy in (Printer)",
                "Loa ngoài (Speaker)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "TH",
            "q": "Vì sao máy tính bắt buộc phải cài đặt hệ điều hành trước khi sử dụng phần mềm soạn thảo Word hay bảng tính Excel?",
            "options": [
                "Vì hệ điều hành quản lí phần cứng và cung cấp môi trường để phần mềm ứng dụng chạy được",
                "Vì hệ điều hành tự động gõ toàn bộ văn bản thay cho con người mà không cần bàn phím",
                "Vì phần mềm ứng dụng chỉ lưu được tệp khi máy tính ngắt kết nối với hệ điều hành",
                "Vì hệ điều hành là thiết bị vật lí gắn bên trong vỏ thùng máy thay thế cho bộ xử lí",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "level": "TH",
            "q": "Biện pháp nào sau đây giúp bảo vệ dữ liệu học tập trong máy tính an toàn trước nguy cơ nhiễm virus và hỏng ổ cứng?",
            "options": [
                "Sao lưu định kì lên đám mây và không tải các tệp đính kèm không rõ nguồn gốc",
                "Đặt tất cả tệp tin của mọi môn học trực tiếp ngoài màn hình nền Desktop",
                "Tắt hoàn toàn phần mềm bảo vệ máy tính để mở được các phần mềm bẻ khóa",
                "Đổi phần mở rộng của tệp văn bản .docx sang .jpg trước khi tắt máy tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "TH",
            "q": "Khi cần gửi bài tập báo cáo gồm tệp văn bản .docx và bảng tính .xlsx cho giáo viên kèm tiêu đề trang trọng, kênh trao đổi nào phù hợp nhất?",
            "options": [
                "Thư điện tử (Email)",
                "Phần mềm xem ảnh Windows Photo Viewer",
                "Công cụ vẽ hình Microsoft Paint",
                "Trình phát nhạc Windows Media Player",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "TH",
            "q": "Nếu phát hiện một tài khoản lạ liên tục gửi tin nhắn vu khống và ghép ảnh trêu chọc em trên mạng xã hội, cách xử lí nào đúng đắn nhất?",
            "options": [
                "Không nhắn tin cãi lại, chụp màn hình lưu bằng chứng, chặn tài khoản đó và báo cho bố mẹ, thầy cô",
                "Lập nhiều tài khoản ảo khác để nhắn tin chửi bới và tung tin nói xấu lại người đó",
                "Giấu kín chuyện một mình, nghỉ học và xóa sạch mọi tin nhắn để không ai biết",
                "Chuyển tiền theo yêu cầu của tài khoản lạ để họ tự nguyện gỡ bài viết xuống",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "TH",
            "image_key": "g7_bang_tinh_excel_mau",
            "image_caption": "Hình minh họa: Bảng quyên góp sách lớp 7A1 (Quan sát bảng tính để trả lời Câu 26, Câu 27, Câu 28 và Câu 2 phần Tự luận)",
            "q": "Quan sát hình bảng tính 'BẢNG QUYÊN GÓP SÁCH LỚP 7A1' ở trên, vùng dữ liệu chứa số lượng Sách giáo khoa và Truyện thiếu nhi của cả 4 tổ (từ Tổ 1 đến Tổ 4) có địa chỉ là gì và gồm bao nhiêu ô tính?",
            "options": [
                "Vùng B3:C6 gồm 8 ô tính",
                "Vùng A3:D6 gồm 12 ô tính",
                "Vùng B2:C7 gồm 8 ô tính",
                "Vùng B3:B6 gồm 6 ô tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "TH",
            "q": "Trong bảng 'BẢNG QUYÊN GÓP SÁCH LỚP 7A1', tại ô D3 đang có công thức =B3+C3. Khi sao chép công thức từ ô D3 xuống ô D5 thì công thức tại ô D5 tự động chuyển thành gì và cho kết quả bao nhiêu?",
            "options": [
                "Công thức =B5+C5, kết quả bằng 45",
                "Công thức =B3+C3, kết quả bằng 40",
                "Công thức =B4+C4, kết quả bằng 50",
                "Công thức =B5*C5, kết quả bằng 500",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "level": "TH",
            "q": "Dựa vào số liệu cột C (Truyện thiếu nhi: C3=15, C4=20, C5=25, C6=10) trên bảng tính lớp 7A1, khi nhập công thức =MAX(C3:C6) và =AVERAGE(C3:C6) thì kết quả nhận được lần lượt là bao nhiêu?",
            "options": [
                "25 và 17,5",
                "35 và 27,5",
                "10 và 70",
                "25 và 70",
            ],
            "ans": 0,
        },
    ],

    "exam_tl": [
        {
            "level": "Thông hiểu",
            "points": "1,0 điểm",
            "lines": 5,
            "q": "Bạn Lan nhận được tin nhắn trên mạng xã hội từ một người lạ đe dọa sẽ tung ảnh chế xấu xí của Lan lên nhóm trường nếu Lan không gửi mật khẩu tài khoản mạng xã hội cho họ. Em hãy tư vấn cho Lan 4 bước xử lí an toàn, văn minh và nêu 2 thói quen giúp phòng tránh nghiện Internet.",
            "rubric": [
                "Tư vấn đúng quy trình 4 bước xử lí khi bị bắt nạt/đe dọa qua mạng (0,5 điểm): (1) Giữ bình tĩnh, tuyệt đối không gửi mật khẩu và không nhắn tin chửi bới trả đũa; (2) Chụp ảnh màn hình tin nhắn đe dọa để lưu giữ bằng chứng; (3) Sử dụng tính năng Chặn (Block) và Báo cáo vi phạm (Report) tài khoản đó; (4) Kể ngay sự việc cho bố mẹ, thầy cô giáo chủ nhiệm để được bảo vệ, can thiệp kịp thời.",
                "Nêu được ít nhất 2 thói quen phòng tránh nghiện Internet hợp lí (0,5 điểm): Quy định rõ khung giờ dùng máy tính mỗi ngày (không thức khuya lướt mạng); tắt thông báo mạng xã hội khi đang học bài; tích cực tham gia thể dục thể thao, đọc sách giấy và trò chuyện trực tiếp cùng gia đình, bạn bè.",
            ],
        },
        {
            "level": "Vận dụng",
            "points": "1,0 điểm",
            "lines": 6,
            "q": "Quan sát bảng tính 'BẢNG QUYÊN GÓP SÁCH LỚP 7A1' (biết số liệu Sách giáo khoa B3=25, B4=30, B5=20, B6=35; Truyện thiếu nhi C3=15, C4=20, C5=25, C6=10):\na) Em hãy viết công thức sử dụng hàm thích hợp tại ô B7 để tính tổng số Sách giáo khoa, tại ô C7 để tính tổng số Truyện thiếu nhi và tại ô D7 để tính tổng số sách quyên góp của cả lớp 7A1.\nb) Hãy tính giá trị kết quả số hiển thị cụ thể tại các ô D4, B7, C7 và D7.",
            "rubric": [
                "Ý a – Viết đúng cú pháp công thức hàm bắt đầu bằng dấu '=' (0,5 điểm): Ô B7: =SUM(B3:B6); Ô C7: =SUM(C3:C6); Ô D7: =SUM(D3:D6) (hoặc =SUM(B7:C7) hoặc =SUM(B3:C6)).",
                "Ý b – Tính đúng kết quả tại cả 4 ô (0,5 điểm, mỗi cặp đúng được 0,25 điểm): D4 = 30 + 20 = 50; B7 = 25 + 30 + 20 + 35 = 110; C7 = 15 + 20 + 25 + 10 = 70; D7 = 110 + 70 = 180 (quyển).",
            ],
        },
        {
            "level": "Vận dụng cao",
            "points": "1,0 điểm",
            "lines": 6,
            "q": "Vẫn trên bảng tính 'BẢNG QUYÊN GÓP SÁCH LỚP 7A1':\na) Sau khi kiểm đếm lại, Tổ 3 bổ sung thêm 10 quyển Sách giáo khoa nên giáo viên sửa giá trị ô B5 từ 20 thành 30. Hãy giải thích vì sao nên dùng công thức chứa địa chỉ ô thay vì gõ trực tiếp phép cộng =20+25 vào ô D5, và cho biết giá trị mới tại ô D5, B7, D7 tự động cập nhật thành bao nhiêu?\nb) Viết công thức hàm để tìm số lượng Sách giáo khoa trung bình mỗi tổ quyên góp được sau khi cập nhật và tính kết quả.",
            "rubric": [
                "Ý a – Giải thích bản chất tính toán tự động & tính đúng giá trị cập nhật (0,5 điểm): Dùng địa chỉ ô (=B5+C5, =SUM(B3:B6)...) giúp phần mềm bảng tính tự động tính lại tức thời khi sửa B5=30 mà không phải sửa lại công thức thủ công. Kết quả mới: D5 = 30 + 25 = 55; B7 = 25 + 30 + 30 + 35 = 120; D7 = 120 + 70 = 190.",
                "Ý b – Viết đúng hàm trung bình cộng & tính đúng kết quả mới (0,5 điểm): Công thức =AVERAGE(B3:B6); kết quả sau cập nhật = (25 + 30 + 30 + 35) / 4 = 120 / 4 = 30 (quyển/tổ).",
            ],
        },
    ],

    # ═════════════════════════════════════════════════════════════════════════
    # NGÂN HÀNG 80 CÂU HỎI ÔN TẬP CUỐI HỌC KÌ 1 (5 - 3 - 1 - 1 | 56 TN - 24 TL)
    # ═════════════════════════════════════════════════════════════════════════

    "review_nb_mcq": [
        # Bài 1 (5 câu NB)
        {
            "bai": "Bài 1",
            "q": "Nhóm thiết bị nào sau đây có chức năng tiếp nhận thông tin đưa vào máy tính?",
            "options": [
                "Thiết bị vào (Input)",
                "Thiết bị ra (Output)",
                "Bộ nguồn điện dự phòng",
                "Vỏ thùng máy tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Máy quét ảnh (Scanner) chuyển hình ảnh trên giấy vào máy tính nên được xếp vào loại thiết bị nào?",
            "options": [
                "Thiết bị vào",
                "Thiết bị ra",
                "Bộ nhớ ngoài",
                "Phần mềm ứng dụng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Máy chiếu (Projector) dùng trong phòng học thực hiện chức năng nào?",
            "options": [
                "Chiếu hình ảnh từ máy tính lên màn chiếu rộng cho cả lớp quan sát",
                "Thu âm giọng giảng bài của giáo viên lưu thành tệp âm thanh",
                "Quét mã vạch trên bìa sách giáo khoa đưa vào phần mềm thư viện",
                "Lưu trữ dự phòng các tệp văn bản khi máy tính bị mất điện",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Thiết bị nào giúp ghi nhận hình ảnh chuyển động của người dùng đưa vào máy tính khi học trực tuyến?",
            "options": [
                "Webcam",
                "Loa vi tính",
                "Máy in màu",
                "Tai nghe chụp tai",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Thao tác nào sau đây đảm bảo an toàn khi kết nối chuột có dây vào máy tính?",
            "options": [
                "Quan sát đúng hình dạng cổng USB và cắm thẳng nhẹ nhàng, vừa khít",
                "Dùng lực ấn thật mạnh đầu cắm USB vào cổng mạng LAN hoặc cổng HDMI",
                "Cắm và rút đầu dây liên tục nhiều lần khi tay đang ướt nước",
                "Nắm phần dây điện mảnh giật mạnh ra khỏi thùng máy khi tháo chuột",
            ],
            "ans": 0,
        },

        # Bài 2 (5 câu NB)
        {
            "bai": "Bài 2",
            "q": "Phần mềm máy tính được chia thành hai nhóm chính nào?",
            "options": [
                "Hệ điều hành và phần mềm ứng dụng",
                "Thiết bị vào và thiết bị ra",
                "Bộ nhớ trong và bộ nhớ ngoài",
                "Tệp văn bản và tệp hình ảnh",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Hệ điều hành nào dưới đây được cài đặt phổ biến trên điện thoại thông minh iPhone?",
            "options": [
                "iOS",
                "Linux Ubuntu",
                "Microsoft Excel",
                "Windows 11",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Phần mềm nào sau đây là phần mềm ứng dụng hỗ trợ soạn thảo văn bản?",
            "options": [
                "Microsoft Word",
                "Android",
                "macOS",
                "Windows 10",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Tệp bảng tính mặc định được lưu từ phần mềm Microsoft Excel (từ phiên bản 2007 trở lên) có phần mở rộng là gì?",
            "options": [
                ".xlsx",
                ".docx",
                ".pptx",
                ".mp3",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Phần mở rộng nào dưới đây thuộc về tệp tin đã được nén dung lượng?",
            "options": [
                ".zip",
                ".jpg",
                ".mp4",
                ".sb3",
            ],
            "ans": 0,
        },

        # Bài 3 (5 câu NB)
        {
            "bai": "Bài 3",
            "q": "Trong hệ điều hành Windows, dữ liệu trên ổ đĩa được tổ chức lưu trữ theo cấu trúc nào?",
            "options": [
                "Cấu trúc hình cây gồm các thư mục và tệp tin",
                "Cấu trúc vòng tròn khép kín không phân cấp",
                "Một danh sách phẳng duy nhất không cho phép tạo thư mục con",
                "Chỉ cho phép lưu tệp tin mà không có thư mục chứa",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Tổ hợp phím tắt Ctrl + C trong cửa sổ quản lí tệp dùng để thực hiện thao tác nào?",
            "options": [
                "Sao chép tệp hoặc thư mục đang chọn vào bộ nhớ tạm",
                "Xóa vĩnh viễn tệp hoặc thư mục khỏi ổ đĩa cứng",
                "Đổi tên tệp hoặc thư mục sang tên viết hoa",
                "Tạo ngay một thư mục con mới bên trong ổ đĩa C",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Để di chuyển một tệp từ thư mục hiện tại sang thư mục khác bằng bàn phím, ta dùng cặp tổ hợp phím nào?",
            "options": [
                "Ctrl + X rồi đến thư mục đích nhấn Ctrl + V",
                "Ctrl + C rồi nhấn phím Delete",
                "F2 rồi nhấn phím Esc",
                "Alt + F4 rồi nhấn phím Enter",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Khi chọn một tệp trên ổ đĩa cứng và nhấn phím Delete, tệp đó mặc định sẽ được chuyển đi đâu?",
            "options": [
                "Chuyển vào thùng rác (Recycle Bin) của máy tính",
                "Gửi tự động lên mạng xã hội Facebook",
                "Sao chép sang tất cả các ổ đĩa còn lại",
                "Đổi thành tệp hệ thống của Windows",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Kí tự nào sau đây KHÔNG được phép đặt trong tên tệp hoặc tên thư mục trên hệ điều hành Windows?",
            "options": [
                "Dấu hỏi chấm (?) hoặc dấu gạch chéo (/)",
                "Chữ số tự nhiên từ 0 đến 9",
                "Dấu gạch dưới (_)",
                "Dấu gạch ngang (-)",
            ],
            "ans": 0,
        },

        # Bài 4 (5 câu NB)
        {
            "bai": "Bài 4",
            "q": "Chức năng chính của mạng xã hội trên Internet là gì?",
            "options": [
                "Kết nối các thành viên để trao đổi, chia sẻ thông tin và tương tác cộng đồng",
                "Soạn thảo và in ấn hợp đồng văn bản khi máy tính không có mạng",
                "Quét và sửa lỗi vật lí trên bề mặt đĩa cứng của máy tính",
                "Điều khiển tốc độ quạt tản nhiệt bên trong thân máy tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Ứng dụng Zalo do doanh nghiệp công nghệ Việt Nam phát triển thuộc nhóm dịch vụ nào?",
            "options": [
                "Mạng xã hội và nhắn tin, gọi điện trực tuyến",
                "Hệ điều hành mã nguồn mở cho máy chủ",
                "Phần mềm xử lí bảng tính chuyên dụng",
                "Trình biên dịch ngôn ngữ lập trình máy",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Website nơi người dùng đăng câu hỏi theo từng chủ đề chuyên môn và cùng thảo luận trả lời được gọi là gì?",
            "options": [
                "Diễn đàn trực tuyến (Forum)",
                "Thanh công thức (Formula Bar)",
                "Thùng rác hệ thống (Recycle Bin)",
                "Trình quản lí tác vụ (Task Manager)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Nền tảng mạng xã hội nào dưới đây chuyên dùng để đăng tải, xem và chia sẻ các video?",
            "options": [
                "YouTube",
                "WinRAR",
                "Notepad",
                "WordPad",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Thông tin nào sau đây tuyệt đối KHÔNG nên công khai trên trang cá nhân mạng xã hội ở chế độ mọi người đều xem được?",
            "options": [
                "Số căn cước, mật khẩu tài khoản và địa chỉ nhà riêng kèm lịch vắng nhà",
                "Sở thích đọc sách khoa học thường thức vào dịp cuối tuần",
                "Bức tranh phong cảnh thiên nhiên do chính em tự vẽ bằng phần mềm Paint",
                "Lời chúc mừng ngày Nhà giáo Việt Nam 20/11 gửi tới thầy cô",
            ],
            "ans": 0,
        },

        # Bài 5 (5 câu NB)
        {
            "bai": "Bài 5",
            "q": "Hành vi cố ý sử dụng Internet để gửi tin nhắn đe dọa, làm nhục hoặc cô lập người khác được gọi là gì?",
            "options": [
                "Bắt nạt qua mạng (Cyberbullying)",
                "Sao lưu dữ liệu trực tuyến",
                "Hội nghị truyền hình từ xa",
                "Học tập hợp tác trên đám mây",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Trước khi đăng tải bức ảnh chụp chung có mặt bạn cùng lớp lên mạng xã hội, em cần làm gì?",
            "options": [
                "Hỏi ý kiến và chỉ đăng khi bạn đồng ý",
                "Chỉnh sửa mặt bạn thành hình hài hước rồi đăng ngay",
                "Đăng kèm số điện thoại riêng của bố mẹ bạn lên bài viết",
                "Gắn thẻ tài khoản của người lạ để câu lượt tương tác",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Tác hại nào thường xảy ra khi học sinh lạm dụng trò chơi trực tuyến và mạng xã hội quá nhiều giờ mỗi ngày?",
            "options": [
                "Mỏi mắt, thiếu ngủ, giảm sút kết quả học tập và hạn chế giao tiếp thực tế",
                "Tăng cường thị lực và giúp cột sống phát triển dẻo dai hơn",
                "Tự động hoàn thành toàn bộ bài tập về nhà các môn trên lớp",
                "Giúp dung lượng ổ cứng máy tính tự động tăng gấp đôi",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Khi vô tình nhận được một đường dẫn (link) lạ kèm lời mời 'Nhấn vào đây để xem bí mật động trời', em nên làm gì?",
            "options": [
                "Không nhấp vào đường dẫn đó và xóa tin nhắn để tránh mã độc, lừa đảo",
                "Nhấp chuột mở ngay lập tức và nhập tài khoản cá nhân để đăng nhập",
                "Chuyển tiếp đường dẫn đó cho toàn bộ danh bạ bạn bè trong lớp",
                "Tải tệp đính kèm về máy rồi tắt phần mềm diệt virus để chạy tệp",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Tổng đài điện thoại quốc gia bảo vệ trẻ em tại Việt Nam (miễn phí cước gọi, hoạt động 24/7) có số máy là bao nhiêu?",
            "options": [
                "111",
                "999",
                "1080",
                "1900",
            ],
            "ans": 0,
        },

        # Bài 6 (5 câu NB)
        {
            "bai": "Bài 6",
            "q": "Trong cửa sổ làm việc của phần mềm bảng tính Excel, tên các cột được kí hiệu như thế nào?",
            "options": [
                "Bằng các chữ cái in hoa A, B, C, ... Z, AA, AB...",
                "Bằng các chữ số tự nhiên 1, 2, 3, 4...",
                "Bằng các chữ số La Mã I, II, III, IV...",
                "Bằng các kí hiệu phép toán +, -, *, /",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Hộp tên (Name Box) nằm ở góc trên bên trái trang tính Excel có chức năng gì?",
            "options": [
                "Hiển thị địa chỉ của ô tính hiện thời đang được chọn",
                "Hiển thị dung lượng bộ nhớ RAM còn trống của máy tính",
                "Thay đổi màu nền cho toàn bộ màn hình Windows",
                "Kiểm tra tốc độ kết nối mạng Internet của phòng máy",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Thanh công thức (Formula Bar) trong phần mềm bảng tính có biểu tượng nhận diện nào ở đầu thanh?",
            "options": [
                "Kí hiệu fx",
                "Kí hiệu @",
                "Kí hiệu #",
                "Kí hiệu www",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Khi nhập dòng chữ 'Lớp 7A1' vào một ô tính mới chưa định dạng, phần mềm Excel mặc định căn lề dữ liệu đó sang phía nào?",
            "options": [
                "Căn thẳng lề trái của ô",
                "Căn thẳng lề phải của ô",
                "Căn giữa theo chiều ngang của ô",
                "Đẩy chữ ra ngoài phạm vi trang tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Một bảng tính (Workbook) trong Excel có thể bao gồm bao nhiêu trang tính (Sheet)?",
            "options": [
                "Một hoặc nhiều trang tính (Sheet1, Sheet2...)",
                "Bắt buộc chỉ có duy nhất một trang tính cố định",
                "Tối đa chỉ được hai trang tính cho mỗi tệp",
                "Không có trang tính nào cho đến khi in ra giấy",
            ],
            "ans": 0,
        },

        # Bài 7 (5 câu NB)
        {
            "bai": "Bài 7",
            "q": "Trong phần mềm bảng tính Excel, phép tính lũy thừa (ví dụ 2 mũ 3) sử dụng kí hiệu nào?",
            "options": [
                "Dấu mũ (^), viết là =2^3",
                "Dấu sao (*), viết là =2*3",
                "Dấu phần trăm (%), viết là =2%3",
                "Dấu và (&), viết là =2&3",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Nếu gõ chuỗi B2+C2 (không có dấu bằng ở đầu) vào ô D2 rồi nhấn Enter thì ô D2 sẽ hiển thị gì?",
            "options": [
                "Hiển thị nguyên văn dòng chữ B2+C2 như dữ liệu văn bản",
                "Hiển thị tổng giá trị hai số trong ô B2 và ô C2",
                "Hiển thị thông báo lỗi chia cho số 0 (#DIV/0!)",
                "Tự động xóa trắng dữ liệu ở cả hai ô B2 và C2",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Chấm vuông nhỏ màu xanh ở góc dưới bên phải của khung ô tính đang chọn dùng để kéo sao chép công thức nhanh gọi là gì?",
            "options": [
                "Nút điền tự động (Fill Handle)",
                "Hộp tên địa chỉ (Name Box)",
                "Thanh cuộn dọc (Scroll Bar)",
                "Nút phóng to cửa sổ (Maximize)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Biểu thức toán học (15 + 25) : 5 được chuyển sang công thức đúng trong Excel là cách viết nào?",
            "options": [
                "=(15+25)/5",
                "=(15+25):5",
                "[15+25]/5",
                "=15+25/5",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Khi ô D3 chứa công thức =B3*C3, nếu ta chọn ô D3 thì trên Thanh công thức (Formula Bar) hiển thị nội dung gì?",
            "options": [
                "Hiển thị công thức =B3*C3",
                "Chỉ hiển thị chữ D3 mà không hiện công thức",
                "Chỉ hiển thị tên tệp bảng tính đang mở",
                "Không hiển thị bất kì kí tự nào",
            ],
            "ans": 0,
        },

        # Bài 8 (5 câu NB)
        {
            "bai": "Bài 8",
            "q": "Cú pháp viết hàm chuẩn xác trong phần mềm bảng tính Excel là dạng nào sau đây?",
            "options": [
                "=<Tên_hàm>(<các tham số>)",
                "<Tên_hàm>=[<các tham số>]",
                "=([<Tên_hàm>] <các tham số>)",
                "=<các tham số>(<Tên_hàm>)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "q": "Hàm SUM trong phần mềm bảng tính được sử dụng để thực hiện công việc gì?",
            "options": [
                "Tính tổng các giá trị số trong danh sách tham số",
                "Tìm giá trị lớn nhất trong vùng dữ liệu được chọn",
                "Đếm số ô trống chưa nhập dữ liệu trong bảng",
                "Sắp xếp họ tên học sinh theo thứ tự bảng chữ cái",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "q": "Hàm MAX(B3:B10) trả về kết quả nào của vùng dữ liệu từ ô B3 đến ô B10?",
            "options": [
                "Giá trị số lớn nhất trong vùng B3:B10",
                "Giá trị số nhỏ nhất trong vùng B3:B10",
                "Trung bình cộng các số trong vùng B3:B10",
                "Số lượng ô có chứa chữ cái trong vùng B3:B10",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "q": "Hàm COUNT trong Excel dùng để đếm đối tượng nào trong vùng tham số?",
            "options": [
                "Đếm số lượng các ô chứa dữ liệu kiểu số",
                "Đếm tổng số chữ cái tiếng Việt trong một câu văn",
                "Đếm số lượng trang tính đang ẩn trong tệp",
                "Đếm tổng giá trị tiền tệ cộng dồn của các ô",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "q": "Cách viết hàm nào dưới đây đúng cú pháp của Excel để tính tổng vùng dữ liệu từ ô C2 đến ô C8?",
            "options": [
                "=SUM(C2:C8)",
                "=SUM(C2-C8)",
                "SUM=(C2:C8)",
                "=SUM[C2..C8]",
            ],
            "ans": 0,
        },
    ],

    "review_th_mcq": [
        # Bài 1 (2 câu TH)
        {
            "bai": "Bài 1",
            "q": "Khi máy tính để bàn vẫn sáng đèn nguồn trên thân máy nhưng màn hình tối đen kèm thông báo 'No Signal', nguyên nhân hợp lí nhất là gì?",
            "options": [
                "Dây cáp truyền tín hiệu hình ảnh (HDMI/VGA) nối từ thân máy sang màn hình bị lỏng hoặc tuột",
                "Chuột máy tính bị hết pin nên màn hình tự động xóa hệ điều hành",
                "Bàn phím bị kẹt phím Spacebar làm hỏng vĩnh viễn bóng đèn màn hình",
                "Thùng rác Recycle Bin chứa nhiều tệp văn bản làm ngắt điện màn hình",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Tại quầy thanh toán siêu thị, nhân viên chiếu thiết bị cầm tay vào mã sọc đen trắng trên hộp sữa để máy tính tự nhận tên hàng và giá tiền, sau đó in hóa đơn giấy. Cặp thiết bị vừa dùng lần lượt thuộc loại nào?",
            "options": [
                "Đầu đọc mã vạch là thiết bị vào; máy in hóa đơn là thiết bị ra",
                "Đầu đọc mã vạch là thiết bị ra; máy in hóa đơn là thiết bị vào",
                "Cả đầu đọc mã vạch và máy in hóa đơn đều là thiết bị vào",
                "Cả đầu đọc mã vạch và máy in hóa đơn đều là phần mềm hệ thống",
            ],
            "ans": 0,
        },

        # Bài 2 (2 câu TH)
        {
            "bai": "Bài 2",
            "q": "Bạn An nhận được 4 tệp tin trong thư mục Bài tập: Bai_hat.mp3, Thuyet_trinh.pptx, Bang_diem.xlsx, Phong_canh.jpg. Nếu An muốn mở tệp bài trình chiếu để luyện nói trước lớp thì An cần mở tệp nào?",
            "options": [
                "Thuyet_trinh.pptx",
                "Bai_hat.mp3",
                "Bang_diem.xlsx",
                "Phong_canh.jpg",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Điều gì xảy ra nếu người dùng tự ý đổi phần mở rộng của tệp ảnh Hinh_hoc.jpg thành Hinh_hoc.xlsx trong File Explorer?",
            "options": [
                "Hệ điều hành nhận nhầm loại tệp và gọi Excel mở tệp ảnh dẫn đến báo lỗi không đọc được định dạng",
                "Bức ảnh tự động biến thành bảng tính chứa đầy đủ các công thức tính chu vi, diện tích",
                "Dung lượng bức ảnh giảm xuống bằng 0 KB nhưng hình ảnh vẫn hiển thị nét hơn trong Excel",
                "Máy tính tự động cài đặt thêm phần mềm chỉnh sửa video chuyên nghiệp",
            ],
            "ans": 0,
        },

        # Bài 3 (2 câu TH)
        {
            "bai": "Bài 3",
            "q": "Em cần lưu các tệp bài tập của 3 môn Toán, Ngữ văn, Tin học trong học kì 1 sao cho dễ tìm nhất. Cách tạo cây thư mục nào dưới đây khoa học nhất?",
            "options": [
                "Tạo thư mục HOC_KI_1 trên ổ đĩa D:, bên trong tạo 3 thư mục con TOAN, NGU_VAN, TIN_HOC để lưu tệp đúng môn",
                "Lưu lẫn lộn tất cả các tệp vào ổ đĩa C:\\Windows\\System32 với tên tệp đặt là 1, 2, 3, 4",
                "Tạo 50 thư mục rỗng lồng nhau liên tiếp rồi giấu tệp ở thư mục cuối cùng không đặt tên gợi nhớ",
                "Lưu toàn bộ tệp vào Thùng rác (Recycle Bin) để tiết kiệm không gian màn hình nền",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Sự khác biệt cơ bản giữa thao tác Sao chép (Copy – Paste) và Di chuyển (Cut – Paste) một tệp tin là gì?",
            "options": [
                "Sao chép giữ nguyên tệp gốc ở thư mục nguồn, còn Di chuyển đưa tệp sang thư mục đích và không còn tệp ở thư mục nguồn",
                "Sao chép làm mất tệp ở thư mục nguồn, còn Di chuyển tạo ra hai tệp giống hệt nhau ở cả hai nơi",
                "Sao chép chỉ thực hiện được với tệp hình ảnh, còn Di chuyển chỉ thực hiện được với tệp âm thanh",
                "Sao chép làm giảm một nửa dung lượng tệp, còn Di chuyển làm tăng gấp đôi dung lượng tệp",
            ],
            "ans": 0,
        },

        # Bài 4 (2 câu TH)
        {
            "bai": "Bài 4",
            "q": "Khi đọc được một bài viết trên mạng xã hội thông báo 'Ngày mai toàn thành phố nghỉ học vì bão lớn' từ một trang cá nhân lạ không có dấu xác thực, em nên hành động như thế nào?",
            "options": [
                "Chưa vội tin hay chia sẻ, kiểm tra lại thông báo từ giáo viên chủ nhiệm và cổng thông tin Sở Giáo dục",
                "Chia sẻ ngay lập tức vào tất cả các nhóm lớp để các bạn cùng nghỉ học sáng mai",
                "Bình luận cãi nhau gay gắt dưới bài viết dù chưa kiểm tra nguồn tin chính thống",
                "Gửi thông tin đăng nhập tài khoản học tập của mình cho chủ bài viết để hỏi thêm",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "So với tin nhắn tức thời trên mạng xã hội, việc trao đổi qua Thư điện tử (Email) có ưu điểm nổi bật nào trong môi trường học tập và công việc?",
            "options": [
                "Trình bày nội dung có tiêu đề rõ ràng, quản lí thư khoa học và đính kèm nhiều tệp tài liệu chuẩn mực",
                "Tự động phát trực tiếp hình ảnh camera của người nhận ngay cả khi họ không mở máy",
                "Không cần kết nối mạng Internet mà vẫn gửi được tệp dung lượng hàng chục GB tức thời",
                "Cho phép chỉnh sửa trực tiếp các tệp đang lưu trong ổ cứng của người nhận mà không cần họ tải về",
            ],
            "ans": 0,
        },

        # Bài 5 (2 câu TH)
        {
            "bai": "Bài 5",
            "q": "Trong nhóm chat học tập của lớp 7A1, hai bạn xảy ra tranh luận gay gắt về cách giải bài toán và bắt đầu dùng lời lẽ thiếu kiềm chế. Là thành viên nhóm, em nên ứng xử thế nào?",
            "options": [
                "Khuyên hai bạn bình tĩnh, trao đổi lịch sự tập trung vào lời giải toán hoặc nhờ cô giáo bộ môn giải đáp",
                "Cổ vũ hai bạn cãi nhau to hơn rồi chụp ảnh đăng lên trang cá nhân để câu lượt thích",
                "Đặt biệt danh xấu cho một bạn và kêu gọi cả lớp tẩy chay bạn đó khỏi nhóm học tập",
                "Tung tin bịa đặt về điểm kiểm tra của hai bạn sang nhóm của lớp khác",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Vì sao khi bị bắt nạt qua mạng, nạn nhân KHÔNG nên nhắn tin chửi bới trả đũa kẻ bắt nạt?",
            "options": [
                "Vì trả đũa làm xung đột leo thang căng thẳng hơn và chính mình cũng vi phạm quy tắc ứng xử văn minh",
                "Vì bàn phím máy tính sẽ tự động khóa vĩnh viễn nếu gõ quá ba tin nhắn liên tiếp",
                "Vì nhà cung cấp dịch vụ Internet sẽ thu phí gấp mười lần đối với tin nhắn phản hồi",
                "Vì tin nhắn trả đũa sẽ tự động xóa mất bằng chứng trên máy của bố mẹ và thầy cô",
            ],
            "ans": 0,
        },

        # Bài 6 (2 câu TH)
        {
            "bai": "Bài 6",
            "q": "Một vùng dữ liệu trên trang tính được chọn từ ô A2 đến ô D5 (kí hiệu A2:D5). Hỏi vùng dữ liệu này bao gồm tổng cộng bao nhiêu ô tính?",
            "options": [
                "16 ô tính (4 cột × 4 hàng)",
                "9 ô tính (3 cột × 3 hàng)",
                "20 ô tính (4 cột × 5 hàng)",
                "8 ô tính (2 cột × 4 hàng)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Khi nhập điểm kiểm tra vào cột C trên Excel, bạn Nam vô tình gõ nhầm số không (0) thành chữ cái O trong giá trị '8.O'. Dấu hiệu trực quan nào giúp Nam nhận ra ngay ô đó đang bị hiểu là văn bản chứ không phải số?",
            "options": [
                "Giá trị '8.O' tự động căn thẳng sang lề trái của ô thay vì căn lề phải như các số khác",
                "Ô tính tự động đổi màu nền sang màu tím đậm và phát tiếng kêu cảnh báo",
                "Thanh tiêu đề của phần mềm Excel tự động đóng tệp bảng tính ngay lập tức",
                "Hộp tên Name Box biến mất hoàn toàn khỏi màn hình làm việc",
            ],
            "ans": 0,
        },

        # Bài 7 (2 câu TH)
        {
            "bai": "Bài 7",
            "q": "Cho các ô A1 = 10, B1 = 6, C1 = 2. Nếu tại ô D1 nhập công thức =A1+B1/C1 thì kết quả hiển thị trong ô D1 là bao nhiêu?",
            "options": [
                "13",
                "8",
                "18",
                "32",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Tại ô C2 có công thức =A2*B2. Khi sao chép ô C2 sang ô D3 (lùi sang phải 1 cột và xuống dưới 1 hàng), công thức tại ô D3 sẽ tự động điều chỉnh thành gì?",
            "options": [
                "=B3*C3",
                "=A2*B2",
                "=A3*B3",
                "=B2*C2",
            ],
            "ans": 0,
        },

        # Bài 8 (2 câu TH)
        {
            "bai": "Bài 8",
            "q": "Cho vùng ô A1:A5 lần lượt chứa các dữ liệu sau: A1 = 8, A2 = 10, A3 = 'Vắng', A4 = 6, A5 = 12. Khi thực hiện hàm =COUNT(A1:A5) thì Excel trả về kết quả là bao nhiêu?",
            "options": [
                "4 (vì chỉ có 4 ô chứa dữ liệu kiểu số, bỏ qua ô chữ 'Vắng')",
                "5 (đếm toàn bộ số ô trong vùng từ A1 đến A5)",
                "36 (tổng các giá trị số trong vùng A1:A5)",
                "9 (trung bình cộng các giá trị số trong vùng A1:A5)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8",
            "q": "Biết các ô B2 = 7, C2 = 8, D2 = 9. Công thức nào dưới đây cho kết quả KHÁC với ba công thức còn lại?",
            "options": [
                "=B2+C2+D2/3",
                "=AVERAGE(B2:D2)",
                "=(B2+C2+D2)/3",
                "=SUM(B2:D2)/COUNT(B2:D2)",
            ],
            "ans": 0,
        },
    ],

    "review_th_tl": [
        {
            "bai": "Bài 1",
            "q": "Phân biệt chức năng của thiết bị vào và thiết bị ra của máy tính. Kể tên 3 thiết bị vào, 3 thiết bị ra và giải thích vì sao màn hình cảm ứng được coi là thiết bị vừa vào vừa ra.",
            "answer": "Thiết bị vào tiếp nhận dữ liệu, lệnh điều khiển từ bên ngoài đưa vào máy tính xử lí (ví dụ: bàn phím, chuột, micro, máy quét, webcam). Thiết bị ra xuất thông tin đã xử lí từ máy tính ra dưới dạng hình ảnh, âm thanh, bản in cho con người nhận biết (ví dụ: màn hình thường, loa, máy in, máy chiếu). Màn hình cảm ứng vừa hiển thị hình ảnh kết quả ra cho người xem (Output), vừa nhận trực tiếp thao tác chạm, vuốt ngón tay làm dữ liệu điều khiển đầu vào (Input).",
        },
        {
            "bai": "Bài 2",
            "q": "Nêu vai trò khác nhau giữa hệ điều hành và phần mềm ứng dụng. Cho biết các phần mở rộng .docx, .xlsx, .pptx, .sb3, .mp3, .mp4 tương ứng với loại tệp tin nào?",
            "answer": "Hệ điều hành (Windows, macOS, Android, iOS...) là phần mềm hệ thống quản lí phần cứng, dữ liệu và tạo môi trường nền tảng cho toàn bộ máy tính hoạt động. Phần mềm ứng dụng (Word, Excel, Scratch...) chạy trên nền hệ điều hành để giải quyết từng nhu cầu công việc cụ thể của người dùng. Ý nghĩa đuôi tệp: .docx (tệp văn bản Word), .xlsx (tệp bảng tính Excel), .pptx (tệp trình chiếu PowerPoint), .sb3 (tệp dự án lập trình Scratch), .mp3 (tệp âm thanh), .mp4 (tệp video).",
        },
        {
            "bai": "Bài 3",
            "q": "Tại sao chúng ta cần tổ chức dữ liệu theo cây thư mục hợp lí và sao lưu dữ liệu (Backup) định kì? Nêu 3 cách sao lưu dữ liệu phổ biến hiện nay.",
            "answer": "Tổ chức cây thư mục khoa học giúp phân loại tài liệu ngăn nắp, tránh thất lạc và tìm kiếm nhanh chóng khi cần. Sao lưu dữ liệu định kì giúp khôi phục lại các tệp quan trọng nếu máy tính gặp sự cố hỏng ổ cứng, bị xóa nhầm hoặc nhiễm mã độc tống tiền. Ba cách sao lưu phổ biến: (1) Sao chép sang ổ cứng di động rời hoặc USB; (2) Đồng bộ lên dịch vụ lưu trữ đám mây (Google Drive, OneDrive); (3) Sao lưu sang một máy tính khác trong mạng.",
        },
        {
            "bai": "Bài 4",
            "q": "Trình bày 3 lợi ích thiết thực và 3 nguy cơ tiềm ẩn khi học sinh sử dụng mạng xã hội và các kênh trao đổi thông tin trên Internet.",
            "answer": "Ba lợi ích: (1) Kết nối, trao đổi bài học nhanh chóng với thầy cô và bạn bè; (2) Khai thác kho tài liệu, bài giảng trực tuyến phong phú; (3) Chia sẻ thông tin tích cực, tham gia câu lạc bộ năng khiếu. Ba nguy cơ tiềm ẩn: (1) Tiếp cận tin giả, thông tin sai lệch hoặc độc hại; (2) Bị lộ thông tin cá nhân, lừa đảo chiếm đoạt tài khoản hoặc nhiễm virus qua đường dẫn lạ; (3) Mất nhiều thời gian gây sao nhãng việc học và nguy cơ bị bắt nạt qua mạng.",
        },
        {
            "bai": "Bài 5",
            "q": "Em hãy nêu 4 quy tắc ứng xử văn minh khi tham gia bình luận, chia sẻ thông tin trên mạng xã hội.",
            "answer": "Bốn quy tắc ứng xử văn minh: (1) Sử dụng ngôn ngữ trong sáng, lịch sự, tôn trọng danh dự và nhân phẩm của người khác; (2) Xin phép trước khi đăng tải hình ảnh, thông tin riêng tư của người khác lên mạng; (3) Kiểm chứng độ chính xác của thông tin trước khi nhấn thích hoặc chia sẻ, không tiếp tay lan truyền tin đồn sai sự thật; (4) Tôn trọng bản quyền tác giả khi trích dẫn bài viết, hình ảnh, video.",
        },
        {
            "bai": "Bài 6",
            "q": "Giải thích khái niệm ô tính, địa chỉ ô, vùng dữ liệu trên trang tính Excel và nêu quy tắc căn lề mặc định khi nhập dữ liệu chữ so với dữ liệu số.",
            "answer": "Ô tính là vùng giao nhau giữa một cột và một hàng; địa chỉ ô gồm tên cột ghép liền tên hàng (ví dụ: B4). Vùng dữ liệu là khối các ô liền kề tạo thành hình chữ nhật, ghi bằng địa chỉ ô góc trái trên và ô góc phải dưới ngăn cách bởi dấu hai chấm (ví dụ: B3:D6). Quy tắc căn lề mặc định: dữ liệu kiểu văn bản (chữ) tự động căn thẳng lề trái; dữ liệu kiểu số, ngày tháng, thời gian tự động căn thẳng lề phải của ô.",
        },
        {
            "bai": "Bài 7",
            "q": "So sánh sự khác nhau giữa công thức dùng hằng số (ví dụ =25+15) và công thức dùng địa chỉ ô (ví dụ =B3+C3) tại ô D3. Khi sao chép công thức =B3+C3 từ ô D3 xuống ô D4 thì điều gì xảy ra?",
            "answer": "Công thức =25+15 cố định giá trị nên nếu số liệu trong ô B3 hoặc C3 thay đổi thì ô D3 vẫn giữ kết quả cũ 40 (không tự cập nhật). Ngược lại, công thức =B3+C3 tham chiếu đến địa chỉ ô nên mỗi khi giá trị trong B3 hoặc C3 thay đổi, Excel tự động tính lại kết quả mới chính xác tại D3. Khi sao chép =B3+C3 từ ô D3 xuống ô D4 (dịch xuống 1 hàng), địa chỉ tương đối tự động tăng chỉ số hàng thành =B4+C4.",
        },
        {
            "bai": "Bài 8",
            "q": "Nêu cú pháp tổng quát của hàm trong Excel và trình bày ý nghĩa của 5 hàm cơ bản: SUM, AVERAGE, MAX, MIN, COUNT kèm 1 ví dụ minh họa cho mỗi hàm.",
            "answer": "Cú pháp chung: =<Tên_hàm>(<các tham số>). Ý nghĩa và ví dụ trên vùng B3:B6: (1) SUM: tính tổng các giá trị số, ví dụ =SUM(B3:B6); (2) AVERAGE: tính trung bình cộng các giá trị số, ví dụ =AVERAGE(B3:B6); (3) MAX: tìm giá trị số lớn nhất, ví dụ =MAX(B3:B6); (4) MIN: tìm giá trị số nhỏ nhất, ví dụ =MIN(B3:B6); (5) COUNT: đếm số lượng ô chứa giá trị số trong vùng, ví dụ =COUNT(B3:B6).",
        },
    ],

    "review_vd_tl": [
        {
            "bai": "Bài 1",
            "q": "Phòng Tin học cần trang bị bộ máy tính phục vụ dạy học trực tuyến kết hợp trình chiếu cho cả lớp cùng nghe và nhìn rõ. Em hãy lập danh sách các thiết bị vào - ra cần kết nối thêm vào thân máy và nêu công dụng cụ thể của từng thiết bị trong tiết học đó.",
            "answer": "Danh sách thiết bị cần kết nối: (1) Thiết bị vào: Bàn phím và chuột (nhập lệnh, điều khiển bài giảng), Webcam (ghi hình lớp học/giáo viên truyền lên phòng học trực tuyến), Micro (thu giọng nói phát biểu của giáo viên và học sinh). (2) Thiết bị ra: Máy chiếu hoặc màn hình tương tác lớn (hiển thị nội dung bài giảng cho cả lớp quan sát), Loa phòng học (phát âm thanh từ đầu cầu trực tuyến và video minh họa).",
        },
        {
            "bai": "Bài 2",
            "q": "Trong thư mục Dự án Khoa học của nhóm có 6 tệp sau: Bao_cao_tong_ket.docx, So_lieu_do.xlsx, Slide_gioi_thieu.pptx, Mo_phong.sb3, Anh_thi_nghiem.png, Video_quay_cham.mp4. Em hãy phân loại từng tệp theo nhóm nội dung và chỉ rõ phần mềm ứng dụng phù hợp để mở chỉnh sửa từng tệp.",
            "answer": "Phân loại và phần mềm mở tương ứng: (1) Bao_cao_tong_ket.docx: tệp văn bản -> mở bằng Microsoft Word; (2) So_lieu_do.xlsx: tệp bảng tính -> mở bằng Microsoft Excel; (3) Slide_gioi_thieu.pptx: tệp bài trình chiếu -> mở bằng Microsoft PowerPoint; (4) Mo_phong.sb3: tệp chương trình lập trình trực quan -> mở bằng Scratch 3.0; (5) Anh_thi_nghiem.png: tệp hình ảnh -> mở sửa bằng Microsoft Paint hoặc Photos; (6) Video_quay_cham.mp4: tệp video -> mở xem/biên tập bằng Movies & TV hoặc trình biên tập video.",
        },
        {
            "bai": "Bài 3",
            "q": "Em hãy thiết kế sơ đồ cây thư mục trên ổ đĩa D: để quản lí tài liệu học tập năm Lớp 7 gồm 3 môn (Tin học, Toán, Tiếng Anh), trong đó thư mục Tin học chia thành 2 thư mục con là Bai_tap_Bang_tinh và Du_an_Trinh_chieu. Nêu quy trình phím tắt để di chuyển tệp Quyen_gop_sach.xlsx từ thư mục Downloads vào đúng thư mục con tương ứng.",
            "answer": "Cây thư mục chuẩn: D:\\NAM_HOC_LOP_7 gồm 3 thư mục con: TOAN, TIENG_ANH, TIN_HOC; bên trong TIN_HOC tạo 2 thư mục con: Bai_tap_Bang_tinh và Du_an_Trinh_chieu. Quy trình phím tắt di chuyển tệp: Bước 1: Mở thư mục Downloads, nhấp chọn tệp Quyen_gop_sach.xlsx và nhấn Ctrl + X (Cut); Bước 2: Mở theo đường dẫn D:\\NAM_HOC_LOP_7\\TIN_HOC\\Bai_tap_Bang_tinh và nhấn Ctrl + V (Paste).",
        },
        {
            "bai": "Bài 4",
            "q": "Lớp 7A1 chuẩn bị tổ chức buổi quyên góp áo ấm vùng cao. Em hãy đề xuất cách kết hợp 3 kênh trao đổi thông tin trên Internet (Email, Mạng xã hội nhóm lớp, Google Meet) sao cho kế hoạch triển khai hiệu quả và an toàn nhất.",
            "answer": "Cách phối hợp hiệu quả và an toàn: (1) Dùng Email gửi thư xin phép và bản kế hoạch chi tiết (đính kèm tệp .docx, .xlsx) tới thầy cô chủ nhiệm và Ban phụ huynh lớp; (2) Dùng nhóm kín Zalo/Facebook của lớp để đăng thông báo nhanh, cập nhật tiến độ tiếp nhận đồ quyên góp (không công khai số điện thoại cá nhân lên trang công cộng); (3) Tổ chức họp trực tuyến ngắn buổi tối qua Google Meet để thống nhất phân công nhiệm vụ giữa các tổ trưởng.",
        },
        {
            "bai": "Bài 5",
            "q": "Bạn Hùng thường thức đến 1 giờ sáng để xem video ngắn trên mạng xã hội khiến sáng hôm sau ngủ gật trên lớp; gần đây Hùng còn tự ý chụp ảnh bạn cùng bàn lúc đang ngủ rồi đăng lên nhóm công khai để trêu đùa. Em hãy phân tích các điểm sai của Hùng và đưa ra lời khuyên khắc phục cụ thể.",
            "answer": "Hai điểm chưa đúng của Hùng: (1) Thức khuya lướt mạng quá độ là dấu hiệu lệ thuộc/nghiện Internet, gây hại sức khỏe và giảm chất lượng học tập; (2) Tự ý chụp và đăng ảnh riêng tư của bạn lên mạng công khai để trêu chọc là hành vi thiếu văn hóa, xâm phạm quyền riêng tư và có thể trở thành bắt nạt qua mạng. Lời khuyên: Hùng cần gỡ ngay bức ảnh và xin lỗi bạn chân thành; đồng thời lập thời gian biểu tắt thiết bị điện tử trước 21h30, đi ngủ đúng giờ và tham gia vận động thể thao sau giờ học.",
        },
        {
            "bai": "Bài 6",
            "q": "Giáo viên yêu cầu lập bảng theo dõi nhiệt độ trung bình 5 ngày trong tuần (từ Thứ Hai đến Thứ Sáu) của hai thành phố Hà Nội và Đà Nẵng trên Excel. Em hãy thiết kế bố cục hàng, cột từ ô A1 đến ô D7 và xác định địa chỉ vùng chứa toàn bộ dữ liệu số đo nhiệt độ.",
            "answer": "Thiết kế bố cục trang tính: Ô A1 nhập tiêu đề 'BẢNG THEO DÕI NHIỆT ĐỘ 5 NGÀY'; Hàng 2 làm tiêu đề cột: A2='STT', B2='Thứ', C2='Hà Nội (°C)', D2='Đà Nẵng (°C)'; các hàng từ Hàng 3 đến Hàng 7 lần lượt nhập STT 1..5, Thứ Hai..Thứ Sáu và số đo nhiệt độ tương ứng ở cột C, cột D. Địa chỉ vùng chứa toàn bộ dữ liệu số đo nhiệt độ của hai thành phố là vùng C3:D7 (gồm 2 cột × 5 hàng = 10 ô tính).",
        },
        {
            "bai": "Bài 7",
            "q": "Cho bảng tính mua đồ dùng học tập: Hàng 3 là 'Vở kẻ ngang' có Số lượng tại ô B3 = 12, Đơn giá tại ô C3 = 8500; Hàng 4 là 'Bút bi xanh' có B4 = 5, C4 = 4000; ô E3 lưu mức giảm giá 2000 đồng. Hãy viết công thức tại ô D3 tính Thành tiền (= Số lượng × Đơn giá), tại ô F3 tính Thực trả (= Thành tiền - Giảm giá), và cho biết công thức cùng kết quả khi sao chép từ D3 xuống D4.",
            "answer": "Tại ô D3 viết công thức: =B3*C3 (kết quả = 12 × 8500 = 102000 đồng). Tại ô F3 viết công thức: =D3-E3 (kết quả = 102000 - 2000 = 100000 đồng). Khi sao chép công thức từ ô D3 xuống ô D4, địa chỉ tương đối tự động chuyển thành =B4*C4 và cho kết quả Thành tiền của Bút bi xanh là 5 × 4000 = 20000 đồng.",
        },
        {
            "bai": "Bài 8",
            "q": "Cho bảng điểm kiểm tra thường xuyên môn Tin học của Tổ 1 tại vùng C3:C8 gồm các giá trị: C3 = 8.5, C4 = 9.0, C5 = 7.5, C6 = 10.0, C7 = 8.0, C8 = 8.0. Hãy viết công thức hàm để tìm: (a) Điểm trung bình của tổ; (b) Điểm cao nhất; (c) Điểm thấp nhất; (d) Số bạn đã có điểm kiểm tra, kèm kết quả tính cụ thể.",
            "answer": "(a) Điểm trung bình của tổ: =AVERAGE(C3:C8) -> kết quả = (8.5 + 9.0 + 7.5 + 10.0 + 8.0 + 8.0) / 6 = 51 / 6 = 8.5. (b) Điểm cao nhất: =MAX(C3:C8) -> kết quả = 10.0. (c) Điểm thấp nhất: =MIN(C3:C8) -> kết quả = 7.5. (d) Số bạn đã có điểm kiểm tra: =COUNT(C3:C8) -> kết quả = 6 (bạn).",
        },
    ],

    "review_vdc_tl": [
        {
            "bai": "Bài 1",
            "q": "Một cửa hàng sách thông minh muốn thiết kế quầy tra cứu tự động cho khách hàng tự gõ hoặc chạm tìm vị trí kệ sách và nhận phiếu in sơ đồ vị trí kệ, nhưng không muốn dùng bàn phím và chuột rời để tránh thất lạc. Em hãy đề xuất giải pháp lựa chọn thiết bị vào - ra tối ưu nhất và giải thích cơ chế hoạt động của hệ thống.",
            "answer": "Giải pháp tối ưu: Trang bị 1 Màn hình cảm ứng (Touchscreen) kết hợp 1 Máy in nhiệt mini gắn liền thân quầy và 1 Đầu đọc mã vạch sách. Cơ chế hoạt động: Khách hàng chạm chọn tên sách trực tiếp trên bàn phím ảo của màn hình cảm ứng hoặc đưa mã vạch cuốn sách mẫu vào đầu đọc mã vạch (đầu vào Input); máy tính tra cứu cơ sở dữ liệu rồi hiển thị ngay bản đồ kệ sách lên chính màn hình cảm ứng đó, đồng thời xuất lệnh ra máy in nhiệt để in phiếu ghi số dãy kệ cho khách cầm theo (đầu ra Output).",
        },
        {
            "bai": "Bài 2",
            "q": "Máy tính của bạn Mai gần đây khởi động rất chậm, khi tải một trò chơi miễn phí từ trang web lạ về có tên Game_vui.jpg.exe thì trình duyệt tự động mở hàng loạt trang quảng cáo. Hãy giải thích mánh khóe đặt tên tệp của kẻ xấu và đề xuất quy trình xử lí an toàn cho máy tính của Mai.",
            "answer": "Giải thích mánh khóe: Phần mở rộng thực sự đứng sau dấu chấm cuối cùng là .exe (tệp chương trình thực thi có khả năng chứa virus/phần mềm độc hại), còn chữ .jpg phía trước chỉ là phần tên giả mạo để đánh lừa người dùng tưởng là tệp ảnh an toàn. Quy trình xử lí: Ngắt tạm kết nối mạng Internet; gỡ bỏ chương trình lạ vừa cài; cập nhật và chạy quét toàn bộ ổ đĩa bằng phần mềm diệt virus uy tín (như Windows Security/Defender); sao lưu các tệp học tập sạch và bật chế độ hiển thị phần mở rộng tên tệp (File name extensions) trong File Explorer để không bị lừa lần sau.",
        },
        {
            "bai": "Bài 3",
            "q": "Nhóm làm dự án của em gồm 4 thành viên cùng làm chung bài thuyết trình Lịch sử địa phương, mỗi lần sửa bài các bạn lại đặt tên Bai_lam_moi.pptx, Bai_lam_moi_nhat.pptx, Bai_chot_cuoi.pptx khiến cả nhóm nhầm lẫn phiên bản. Em hãy thiết kế quy ước đặt tên tệp và cấu trúc thư mục quản lí phiên bản chuyên nghiệp cho nhóm.",
            "answer": "Quy ước chuyên nghiệp: Tạo thư mục gốc DU_AN_LICH_SU_7A1 gồm 3 thư mục con: 01_Tu_lieu_Anh_Video (chứa ảnh, tư liệu nguồn), 02_Ban_nhap_Thanh_vien (chứa phần làm riêng), 03_Ban_chinh_thuc (chứa bản tổng hợp). Quy tắc đặt tên tệp theo chuẩn <TenDuAn>_v<SoPhienBan>_<NgayThangNam>_<NguoiSua>.pptx (ví dụ: LichSuDiaPhuong_v01_20261110_Nhom7A1.pptx, LichSuDiaPhuong_v02_20261115_Nhom7A1.pptx), kết hợp sao lưu đồng bộ lên thư mục chung Google Drive để luôn xác định chính xác bản mới nhất.",
        },
        {
            "bai": "Bài 4",
            "q": "Tài khoản mạng xã hội của một người bạn thân bỗng nhiên nhắn tin cho em: 'Tớ đang cấp cứu ngoài đường mà quên ví, chuyển gấp giúp tớ 500 nghìn vào số tài khoản lạ tên Nguyễn Văn X nhé, cấm gọi điện vì máy tớ hỏng loa rồi!'. Em hãy nhận diện dấu hiệu bất thường và nêu cách kiểm chứng thông minh trước khi hành động.",
            "answer": "Dấu hiệu lừa đảo chiếm đoạt tài khoản (hack nick): Hối thúc chuyển tiền gấp đánh vào tâm lí lo lắng, tên chủ tài khoản ngân hàng khác tên bạn mình, và cố tình ngăn cản gọi điện xác minh. Cách xử lí thông minh: Tuyệt đối không chuyển tiền ngay; gọi điện thoại di động thông thường trực tiếp vào số máy của bạn hoặc gọi cho bố mẹ/giáo viên chủ nhiệm của bạn để xác nhận tình hình thực tế; nếu xác nhận tài khoản bạn bị chiếm đoạt thì cảnh báo ngay trên nhóm lớp để các bạn khác không bị lừa.",
        },
        {
            "bai": "Bài 5",
            "q": "Một nhóm học sinh lập trang ẩn danh 'Confession khối 7' để đăng các bài bêu xấu ngoại hình và điểm số của một bạn học sinh mới chuyển trường khiến bạn ấy sợ hãi không dám đến lớp. Dưới góc độ đạo đức và văn hóa số, hãy phân tích hậu quả của hành vi này và đề xuất phương án hành động của tập thể lớp để bảo vệ bạn.",
            "answer": "Phân tích hậu quả: Đây là hành vi bắt nạt qua mạng tập thể nghiêm trọng, xúc phạm danh dự, xâm phạm đời tư và gây tổn thương tâm lí nặng nề cho người bị hại (dù ẩn danh vẫn để lại dấu vết số và vi phạm quy định nhà trường, pháp luật). Phương án hành động của tập thể lớp: Không nhấn thích/bình luận hùa theo bài viết xấu; chụp màn hình bài đăng làm bằng chứng và đồng loạt báo cáo (Report) trang vi phạm; báo cáo ngay với giáo viên chủ nhiệm và Ban giám hiệu để xử lí, gỡ trang; đồng thời cả lớp chủ động trò chuyện, động viên và đồng hành cùng bạn mới hòa nhập.",
        },
        {
            "bai": "Bài 6",
            "q": "Khi lập bảng thu quỹ kế hoạch nhỏ trong Excel, bạn Bình nhập cột B như sau: B3 = '15 kg', B4 = '20 kg', B5 = '25 kg' rồi viết công thức =SUM(B3:B5) tại ô B6 nhưng kết quả lại hiện số 0. Hãy giải thích nguyên nhân lỗi này và hướng dẫn Bình cách thiết kế lại bảng tính đúng chuẩn để tính toán tự động được.",
            "answer": "Nguyên nhân: Khi Bình gõ kèm đơn vị chữ 'kg' trực tiếp vào ô (ví dụ '15 kg'), phần mềm Excel nhận diện toàn bộ nội dung ô là dữ liệu văn bản (căn lề trái) chứ không phải dữ liệu số,而 hàm SUM bỏ qua các ô văn bản nên trả về kết quả bằng 0. Cách khắc phục chuẩn: Đưa đơn vị đo lên ô tiêu đề cột B2 thành 'Khối lượng giấy vụn (kg)', còn trong các ô dữ liệu B3, B4, B5 chỉ nhập thuần giá trị số 15, 20, 25; khi đó công thức =SUM(B3:B5) tại B6 sẽ tính đúng tổng bằng 60.",
        },
        {
            "bai": "Bài 7",
            "q": "Tại ô E4 có công thức =B4*C4-D4. Người dùng thực hiện thao tác Sao chép (Copy) ô E4 rồi Dán (Paste) vào ô F6. Không cần mở máy tính, em hãy xác định chính xác công thức xuất hiện tại ô F6 và giải thích quy luật dịch chuyển tọa độ của Excel. Nếu thay vì Copy-Paste, người dùng dùng Cut-Paste (Di chuyển) từ E4 sang F6 thì công thức tại F6 là gì?",
            "answer": "Khi Sao chép (Copy – Paste) từ E4 sang F6: vị trí ô đích dịch sang phải 1 cột (từ cột E sang cột F) và dịch xuống dưới 2 hàng (từ hàng 4 xuống hàng 6); do đó mọi địa chỉ tương đối trong công thức đều tịnh tiến sang phải 1 cột và xuống 2 hàng: B4 -> C6, C4 -> D6, D4 -> E6, vậy công thức tại F6 là =C6*D6-E6. Ngược lại, nếu dùng Di chuyển (Cut – Paste) từ E4 sang F6 thì Excel chuyển nguyên ô công thức đi chỗ khác nên giữ nguyên tham chiếu gốc là =B4*C4-D4.",
        },
        {
            "bai": "Bài 8",
            "q": "Trong cuộc thi Hùng biện Tiếng Anh, mỗi thí sinh được 5 giám khảo chấm điểm lưu tại các ô từ B3 đến F3. Theo thể lệ công bằng, Điểm chung cuộc của thí sinh tại ô G3 được tính bằng trung bình cộng của 3 điểm còn lại sau khi đã loại bỏ 1 điểm cao nhất và 1 điểm thấp nhất. Chỉ dùng các hàm cơ bản đã học (SUM, MAX, MIN, COUNT), em hãy lập công thức tổng quát tại ô G3 và kiểm chứng với bộ điểm: 8, 9, 9, 10, 7.",
            "answer": "Ý tưởng toán học: Tổng 3 điểm còn lại bằng Tổng cả 5 điểm trừ đi Điểm cao nhất và trừ đi Điểm thấp nhất, sau đó chia cho 3 (hoặc chia cho COUNT(B3:F3)-2). Công thức chuẩn tại ô G3: =(SUM(B3:F3)-MAX(B3:F3)-MIN(B3:F3))/3. Kiểm chứng với bộ điểm (8, 9, 9, 10, 7): SUM = 43, MAX = 10, MIN = 7 -> Kết quả G3 = (43 - 10 - 7) / 3 = 26 / 3 ≈ 8,67 điểm.",
        },
    ],
}
