# -*- coding: utf-8 -*-
"""Ngân hàng Đề kiểm tra & Đề cương 80 câu ôn tập Cuối học kì 1 (ĐGĐK 2) - Môn Tin học Lớp 2 UNIGO."""

GRADE_DATA = {
    "lop": 2,
    "cap_hoc": "Tiểu học",
    "lop_hien_thi": "2A1, 2C1",
    "thoi_gian": "35 phút",
    "tuan_on_tap": 20,
    "tuan_kiem_tra": 21,
    "tiet_on_tap": 19,
    "tiet_kiem_tra": 20,

    "tom_tat_ly_thuyet": [
        {
            "bai": "Chủ đề 1 (Bài 1 – Bài 4): Kỹ năng chuột, bàn phím gõ câu ngắn và chuyển đổi ứng dụng",
            "y_chinh": [
                "Nhà sáng tạo số lớp 2 tuân thủ nội quy phòng máy, ngồi đúng tư thế và sử dụng máy tính văn minh, an toàn.",
                "Nâng cấp kỹ năng chuột: Thực hiện chính xác các thao tác chọn (Click), mở nhanh (Double-click), mở bảng chọn tắt (nhấp nút phải chuột) và kéo – thả đối tượng đúng vị trí.",
                "Bàn phím và gõ câu ngắn: Phối hợp phím chữ, phím số với phím Cách (Space), phím Xuống dòng (Enter), phím Xóa (Backspace), phím Chữ hoa (Shift/Caps Lock) để gõ câu ngắn rõ ràng.",
                "Quản lý cửa sổ ứng dụng: Biết mở ứng dụng học tập, thu nhỏ cửa sổ (nút dấu trừ -), phóng to, chuyển đổi qua lại giữa các ứng dụng trên thanh tác vụ (Taskbar) và đóng ứng dụng (nút dấu X) sau khi lưu bài.",
            ],
        },
        {
            "bai": "Chủ đề 2 (Bài 5 – Bài 10): Tệp, thư mục và tổ chức lưu trữ sản phẩm số",
            "image_key": "g2_tep_va_thu_muc",
            "y_chinh": [
                "Tệp (File) là sản phẩm số được tạo ra và lưu trên máy tính, ví dụ: tệp tranh vẽ (.png, .jpg), tệp bài soạn thảo văn bản (.docx), tệp âm thanh, tệp chương trình.",
                "Thư mục (Folder) giống như ngăn cặp sách hoặc ngăn tủ trên máy tính, dùng để chứa các tệp và các thư mục con giúp sắp xếp gọn gàng, dễ tìm kiếm.",
                "Quy tắc đặt tên và lưu sản phẩm: Đặt tên tệp, tên thư mục ngắn gọn, rõ nghĩa, gợi nhớ đúng nội dung bài làm (ví dụ: Ngoi_truong_em, Tho_tang_me) và chọn đúng thư mục trước khi nhấn Save.",
                "Khi cất sản phẩm đúng thư mục theo chủ đề (như thư mục TRANH_VE, thư mục BAI_SOAN), em có thể tìm và mở lại bài làm rất nhanh chóng.",
            ],
        },
        {
            "bai": "Chủ đề 3 (Bài 12 – Bài 14): Thông tin, dữ liệu và quá trình thu nhận – xử lý thông tin",
            "y_chinh": [
                "Thông tin là những điều em biết, hiểu được về thế giới xung quanh; khi thông tin được ghi lại bằng chữ viết, con số, hình vẽ, âm thanh trong máy tính thì gọi là dữ liệu.",
                "Các dạng thông tin quen thuộc: Dạng chữ, dạng hình ảnh, dạng âm thanh và dạng video (kết hợp cả hình ảnh chuyển động, âm thanh và chữ).",
                "Quá trình xử lý thông tin gồm 3 bước liên tiếp: Đầu vào (tiếp nhận thông tin) -> Xử lý (suy nghĩ, tính toán hoặc máy tính xử lý) -> Kết quả / Đầu ra (hành động, câu trả lời hoặc hiển thị ra màn hình, loa).",
            ],
        },
        {
            "bai": "Chủ đề 4 (Bài 15 – Bài 16): Chia nhiệm vụ thành nhiều bước và thuật toán bằng lời, bằng hình",
            "y_chinh": [
                "Phân rã nhiệm vụ: Chia một công việc lớn thành các bước nhỏ, đơn giản, rõ ràng để dễ thực hiện và không bỏ sót việc.",
                "Thuật toán là một dãy các bước hướng dẫn rõ ràng, sắp xếp theo đúng trình tự từ đầu đến cuối để giải quyết một nhiệm vụ cụ thể.",
                "Em có thể mô tả thuật toán bằng lời (liệt kê Bước 1, Bước 2, Bước 3...) hoặc bằng sơ đồ hình vẽ nối tiếp nhau bằng các mũi tên chỉ hướng.",
            ],
        },
        {
            "bai": "Chủ đề 5 (Bài 17 – Bài 19): Lập trình trực quan với chuỗi lệnh, lặp lại hành động và chạy thử sửa lỗi",
            "image_key": "g2_thuat_toan_chuoi_lenh",
            "y_chinh": [
                "Trong lập trình trực quan, các khối lệnh ghép nối tiếp nhau tạo thành chuỗi lệnh; máy tính thực hiện lần lượt từng lệnh từ trên xuống dưới.",
                "Lặp lại một hành động: Khi một lệnh hoặc nhóm lệnh giống hệt nhau được thực hiện nhiều lần liên tiếp, ta dùng khối lệnh “Lặp lại ... lần” để chương trình ngắn gọn, dễ đọc.",
                "Chạy thử và sửa lỗi (Debugging): Luôn chạy thử chương trình để kiểm tra kết quả; nếu nhân vật làm chưa đúng yêu cầu, ta dò từng bước để phát hiện lệnh sai (sai hướng, sai số bước, thiếu lệnh) và sửa lại.",
            ],
        },
    ],

    "ma_tran_kiem_tra": [
        ["Chủ đề 1: Kỹ năng chuột, bàn phím & ứng dụng học tập (Bài 1 – 4)", "2 câu (1,0đ)", "1 câu (0,5đ)", "-", "-", "-", "1,5 điểm"],
        ["Chủ đề 2: Tệp, thư mục & quản lý sản phẩm số (Bài 5 – 10)", "3 câu (1,5đ)", "1 câu (0,5đ)", "1 câu (1,0đ)", "-", "-", "3,0 điểm"],
        ["Chủ đề 3: Thông tin, dữ liệu & xử lý thông tin (Bài 12 – 14)", "2 câu (1,0đ)", "1 câu (0,5đ)", "-", "1 câu (1,0đ)", "-", "2,5 điểm"],
        ["Chủ đề 4: Thuật toán, chuỗi lệnh, vòng lặp & sửa lỗi (Bài 15 – 19)", "3 câu (1,5đ)", "1 câu (0,5đ)", "-", "-", "1 câu (1,0đ)", "3,0 điểm"],
        ["TỔNG CỘNG", "10 câu (5,0đ)", "4 câu (2,0đ)", "1 câu (1,0đ)", "1 câu (1,0đ)", "1 câu (1,0đ)", "10,0 điểm"],
    ],

    # Đề kiểm tra chính thức: ĐÚNG 14 câu Trắc nghiệm (10 NB + 4 TH) = 7,0 điểm
    "exam_mcq": [
        # --- 10 câu Nhận biết (NB) ---
        {
            "level": "NB",
            "bai": "Bài 2",
            "q": "Để di chuyển một biểu tượng trò chơi học tập từ góc trái sang góc phải màn hình, em dùng thao tác chuột nào?",
            "options": [
                "Kéo – thả chuột",
                "Nhấp chuột phải một lần",
                "Lăn con lăn chuột tại chỗ",
                "Thả cả hai tay khỏi chuột",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 4",
            "q": "Khi muốn đóng hoàn toàn cửa sổ của một ứng dụng sau khi đã lưu bài, em nhấp chuột vào nút nào ở góc trên bên phải cửa sổ?",
            "options": [
                "Nút dấu nhân (X)",
                "Nút dấu trừ (-)",
                "Nút hình kính lúp",
                "Nút hình chiếc loa",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 5",
            "q": "Một bức tranh em vừa vẽ và lưu lại trên máy tính được gọi là gì?",
            "options": [
                "Một tệp sản phẩm số",
                "Một bàn phím máy tính",
                "Một thanh tác vụ",
                "Một dây cáp nguồn",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 6",
            "q": "Trên máy tính, đối tượng có biểu tượng hình chiếc kẹp giấy màu vàng (📁) dùng để chứa các tệp cho gọn gàng được gọi là gì?",
            "options": [
                "Thư mục (Folder)",
                "Con trỏ chuột",
                "Phím Cách (Space)",
                "Màn hình nền",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 7",
            "q": "Khi đặt tên cho tệp bài vẽ về gia đình của em, cách đặt tên nào sau đây rõ ràng và dễ tìm lại nhất?",
            "options": [
                "Gia_dinh_em",
                "123456789",
                "abcxyz_khong_nho",
                "bai_gi_do_000",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 13",
            "q": "Một đoạn phim hoạt hình vừa có hình ảnh chuyển động, vừa có lời thoại nhân vật và nhạc nền thuộc dạng thông tin nào?",
            "options": [
                "Dạng video",
                "Chỉ dạng chữ viết",
                "Chỉ dạng con số",
                "Chỉ dạng hình tĩnh",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 14",
            "q": "Mô hình đơn giản của quá trình xử lý thông tin gồm ba thành phần theo đúng thứ tự nào?",
            "options": [
                "Đầu vào -> Xử lý -> Kết quả (Đầu ra)",
                "Kết quả -> Đầu vào -> Xử lý",
                "Xử lý -> Kết quả -> Đầu vào",
                "Đầu vào -> Kết quả -> Tắt máy",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 16",
            "q": "Một dãy các bước hướng dẫn rõ ràng, sắp xếp theo đúng thứ tự từ đầu đến cuối để hoàn thành một nhiệm vụ được gọi là gì?",
            "options": [
                "Thuật toán",
                "Thư mục rỗng",
                "Hình nền máy tính",
                "Bàn phím chữ số",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 18",
            "q": "Trong lập trình trực quan, khối lệnh “Lặp lại 4 lần” giúp nhân vật thực hiện điều gì?",
            "options": [
                "Thực hiện các lệnh bên trong khối đó đúng 4 lần liên tiếp",
                "Tự động xóa bỏ 4 khối lệnh ở phía trên cùng",
                "Đứng yên tại chỗ và chờ trong 4 phút",
                "Đổi màu sắc của màn hình máy tính 4 lần rồi tắt",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 19",
            "q": "Việc chạy thử chương trình rồi tìm và chỉnh lại các khối lệnh bị ghép sai được gọi là gì?",
            "options": [
                "Phát hiện và sửa lỗi chương trình",
                "Đổi tên thư mục học tập",
                "Tắt nguồn máy tính để bàn",
                "Gõ câu văn ngắn bằng bàn phím",
            ],
            "ans": 0,
        },
        # --- 4 câu Thông hiểu (TH) ---
        {
            "level": "TH",
            "bai": "Bài 3",
            "q": "Bạn Minh đang mở đồng thời phần mềm Vẽ tranh và phần mềm Gõ chữ. Minh muốn chuyển nhanh sang cửa sổ Gõ chữ mà không đóng bài vẽ thì làm thế nào?",
            "options": [
                "Nhấp chuột vào biểu tượng phần mềm Gõ chữ trên thanh tác vụ (Taskbar) ở cạnh dưới màn hình",
                "Nhấn nút tắt nguồn trên thân máy tính rồi khởi động lại từ đầu",
                "Nhấn phím Backspace liên tục mười lần trên bàn phím",
                "Rút dây chuột máy tính ra rồi cắm sang máy tính khác",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 8",
            "image_key": "g2_tep_va_thu_muc",
            "image_caption": "Hình 1. Sơ đồ thư mục HOC_TAP_LOP_2 (dùng để trả lời Câu 12)",
            "q": "Quan sát Hình 1 ở trên, nếu em muốn mở lại tệp văn bản “Tho_tang_me.docx” để đọc cho cô giáo nghe, em cần tìm trong thư mục nào?",
            "options": [
                "Mở thư mục HOC_TAP_LOP_2 rồi mở tiếp thư mục BAI_SOAN",
                "Mở thư mục HOC_TAP_LOP_2 rồi mở tiếp thư mục TRANH_VE",
                "Tìm ở ngoài thùng rác Recycle Bin của máy tính",
                "Tạo một thư mục mới tinh có tên là AM_NHAC",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 14",
            "q": "Cô giáo hỏi: “5 cộng 3 bằng mấy?”, bạn Lan suy nghĩ trong đầu rồi giơ tay trả lời: “Thưa cô bằng 8 ạ!”. Trong tình huống này, hoạt động “Xử lý thông tin” là gì?",
            "options": [
                "Bộ não bạn Lan suy nghĩ và tính tổng của 5 cộng 3",
                "Tai bạn Lan nghe câu hỏi của cô giáo trên bục giảng",
                "Miệng bạn Lan nói to câu trả lời bằng 8",
                "Đôi mắt bạn Lan nhìn lên chiếc bảng đen của lớp",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 18",
            "q": "Thay vì ghép 5 khối lệnh “Nhảy lên 1 bước” giống hệt nhau nối tiếp nhau, cách ghép nào sau đây ngắn gọn hơn mà kết quả vẫn giống hệt?",
            "options": [
                "Dùng khối “Lặp lại 5 lần” bọc bên ngoài 1 khối “Nhảy lên 1 bước”",
                "Dùng khối “Lặp lại 2 lần” bọc bên ngoài 1 khối “Nhảy lên 1 bước”",
                "Chỉ dùng duy nhất 1 khối “Nhảy lên 1 bước” rồi dừng lại",
                "Dùng khối “Lặp lại 5 lần” bọc bên ngoài khối “Lùi lại 1 bước”",
            ],
            "ans": 0,
        },
    ],

    # Đề kiểm tra chính thức: ĐÚNG 3 câu Tự luận / Thực hành trên giấy = 3,0 điểm
    "exam_tl": [
        {
            "level": "Thông hiểu",
            "points": "1,0 điểm",
            "lines": 4,
            "q": "Trong thư mục HOC_TAP_LOP_2 có hai thư mục con là TRANH_VE và BAI_SOAN. Hôm nay em hoàn thành hai sản phẩm mới: (1) Bức tranh vẽ chú mèo tên là Con_meo_nha_em.png; (2) Bài gõ câu văn tên là Loi_chuc_20_11.docx. Em hãy cho biết mỗi tệp nên lưu vào thư mục con nào và giải thích vì sao cần phân loại như vậy.",
            "rubric": [
                "Xác định đúng nơi lưu tệp Con_meo_nha_em.png: Lưu vào thư mục con TRANH_VE (0,25 điểm).",
                "Xác định đúng nơi lưu tệp Loi_chuc_20_11.docx: Lưu vào thư mục con BAI_SOAN (0,25 điểm).",
                "Giải thích đúng lợi ích: Giúp sắp xếp các tệp sản phẩm số ngăn nắp, đúng chủ đề, không bị lẫn lộn và tìm mở lại rất nhanh khi cần (0,5 điểm).",
            ],
        },
        {
            "level": "Vận dụng",
            "points": "1,0 điểm",
            "lines": 4,
            "q": "Khi em dùng máy tính cầm tay hoặc phần mềm Máy tính (Calculator): Em bấm phép tính “12 + 5 =” trên bàn phím và màn hình hiện ra số “17”. Em hãy chỉ rõ đâu là Đầu vào, đâu là Xử lý và đâu là Kết quả (Đầu ra) trong tình huống trên.",
            "rubric": [
                "Xác định đúng Đầu vào: Phép tính “12 + 5 =” do em bấm đưa vào máy tính (0,25 điểm).",
                "Xác định đúng Xử lý: Máy tính thực hiện tính toán cộng hai số 12 và 5 (0,5 điểm).",
                "Xác định đúng Kết quả (Đầu ra): Số “17” hiển thị trên màn hình cho em nhìn thấy (0,25 điểm).",
            ],
        },
        {
            "level": "Vận dụng cao",
            "points": "1,0 điểm",
            "lines": 5,
            "image_key": "g2_thuat_toan_chuoi_lenh",
            "image_caption": "Hình 2. Sơ đồ thuật toán chuỗi lệnh & lặp lại hành động (dùng cho Câu 3 phần Tự luận)",
            "q": "Quan sát Hình 2 ở trên: (a) Ở Bước 2, nhân vật thực hiện lệnh “Đi thẳng 10 bước” tất cả bao nhiêu lần và đi tạo thành hình gì? (b) Nếu muốn nhân vật nói lời chào “Xin chào các bạn!” TRƯỚC khi đi theo hình vuông (thay vì đi xong mới chào), em cần sửa thứ tự các bước như thế nào?",
            "rubric": [
                "Ý (a): Trả lời đúng nhân vật thực hiện lệnh “Đi thẳng 10 bước” tất cả 4 lần (0,25 điểm) và đường đi tạo thành một hình vuông khép kín (0,25 điểm).",
                "Ý (b): Nêu đúng cách sửa thứ tự: Đổi chỗ Bước 3 lên trước Bước 2 (thứ tự mới: Bước 1 nhấn Lá cờ xanh -> Nói 'Xin chào các bạn!' trong 2 giây -> Lặp lại 4 lần [Đi thẳng 10 bước -> Rẽ phải góc vuông] -> Phát âm thanh 'Tiếng vỗ tay' và Dừng lại) (0,5 điểm).",
            ],
        },
    ],

    # ═════════════════════════════════════════════════════════════════════════
    # NGÂN HÀNG 80 CÂU HỎI ÔN TẬP CUỐI HỌC KÌ 1 (40 NB_TN + 16 TH_TN + 8 TH_TL + 8 VD_TL + 8 VDC_TL)
    # ═════════════════════════════════════════════════════════════════════════
    "review_nb_mcq": [
        # 1
        {"bai": "Bài 1", "q": "Thói quen nào sau đây thể hiện em là một “Nhà sáng tạo số” văn minh trong phòng Tin học?", "options": ["Giữ gìn máy tính sạch sẽ, làm bài nghiêm túc và tắt máy đúng cách", "Vẽ bậy bằng bút mực lên màn hình và bàn phím máy tính", "Tự ý xóa các bài vẽ lưu sẵn của bạn lớp khác", "Đập mạnh chuột máy tính xuống bàn khi chờ máy mở"], "ans": 0},
        # 2
        {"bai": "Bài 1", "q": "Khi ngồi làm việc với máy tính, hai bàn chân của em nên đặt như thế nào cho đúng tư thế?", "options": ["Đặt ngay ngắn trên sàn nhà hoặc giá để chân", "Co cả hai chân lên mặt ghế xoay", "Gác chân lên cạnh bàn phím máy tính", "Đung đưa đá mạnh vào dây điện dưới gầm bàn"], "ans": 0},
        # 3
        {"bai": "Bài 2", "q": "Thao tác nhấp nút phải chuột một lần vào đối tượng trên màn hình thường dùng để làm gì?", "options": ["Mở bảng chọn tắt chứa các lệnh thao tác nhanh", "Tắt ngay lập tức nguồn điện của cả phòng máy", "Gõ chữ cái in hoa vào trang văn bản", "Tăng độ sáng của bóng đèn trong lớp học"], "ans": 0},
        # 4
        {"bai": "Bài 2", "q": "Để mở nhanh một thư mục đang nằm trên màn hình nền (Desktop), em dùng thao tác chuột nào?", "options": ["Nhấp đúp nút trái chuột vào biểu tượng thư mục đó", "Trượt chuột đi vòng quanh màn hình nền ba vòng", "Nhấn giữ nút phải chuột và kéo ra góc màn hình", "Lăn con lăn giữa của chuột lên xuống liên tục"], "ans": 0},
        # 5
        {"bai": "Bài 2", "q": "Con lăn nhỏ nằm ở giữa nút trái và nút phải của chuột máy tính giúp em làm gì?", "options": ["Cuộn trang văn bản hoặc trang hình ảnh lên và xuống", "Xóa sạch toàn bộ các tệp có trong ổ đĩa", "Đổi màu sắc của vỏ màn hình máy tính", "Bật loa phát nhạc tự động của trường học"], "ans": 0},
        # 6
        {"bai": "Bài 3", "q": "Khi gõ câu ngắn “Em yêu trường UNIGO.”, sau từ cuối cùng em nhấn phím dấu nào để kết thúc câu?", "options": ["Phím dấu chấm (.)", "Phím dấu cộng (+)", "Phím dấu bằng (=)", "Phím mũi tên trái"], "ans": 0},
        # 7
        {"bai": "Bài 3", "q": "Phím nào trên bàn phím khi bật sáng đèn sẽ giúp em gõ liên tiếp nhiều chữ cái IN HOA mà không cần giữ phím Shift?", "options": ["Phím Caps Lock", "Phím Backspace", "Phím Space", "Phím Enter"], "ans": 0},
        # 8
        {"bai": "Bài 3", "q": "Khi gõ một câu văn ngắn, giữa hai từ liên tiếp nhau em cần gõ mấy dấu cách (Space)?", "options": ["Đúng 1 dấu cách", "5 dấu cách liên tiếp", "Không gõ dấu cách nào", "10 dấu cách liên tiếp"], "ans": 0},
        # 9
        {"bai": "Bài 4", "q": "Ở góc trên bên phải cửa sổ ứng dụng, nút có hình dấu trừ (-) dùng để làm gì?", "options": ["Thu nhỏ cửa sổ ứng dụng xuống thanh tác vụ bên dưới", "Xóa vĩnh viễn phần mềm đó khỏi máy tính", "In bài làm của em ra máy in màu", "Đổi tên tệp bài làm thành tên khác"], "ans": 0},
        # 10
        {"bai": "Bài 4", "q": "Thanh nằm ngang ở cạnh dưới cùng của màn hình Windows, nơi hiển thị các ứng dụng đang mở gọi là gì?", "options": ["Thanh tác vụ (Taskbar)", "Thanh dấu cách bàn phím", "Khung viền nhựa màn hình", "Dây kết nối mạng"], "ans": 0},
        # 11
        {"bai": "Bài 5", "q": "Sản phẩm số nào sau đây KHÔNG phải là một tệp lưu trong máy tính?", "options": ["Quyển vở giấy đặt trong ngăn bàn gỗ", "Tệp tranh vẽ ngôi nhà lưu từ phần mềm Paint", "Tệp bài thơ gõ bằng phần mềm Word", "Tệp bản nhạc thiếu nhi lưu trong máy tính"], "ans": 0},
        # 12
        {"bai": "Bài 5", "q": "Mỗi tệp trên máy tính đều có những đặc điểm nhận biết cơ bản nào?", "options": ["Có biểu tượng riêng và có tên tệp đi kèm", "Không có tên gọi và không có biểu tượng", "Luôn luôn tự đổi tên sau mỗi giờ học", "Chỉ mở được duy nhất một lần rồi tự biến mất"], "ans": 0},
        # 13
        {"bai": "Bài 5", "q": "Tệp có tên “Phong_canh_que_huong.png” thuộc loại tệp sản phẩm số nào?", "options": ["Tệp hình ảnh (tranh vẽ)", "Tệp soạn thảo văn bản chữ", "Thiết bị phần cứng bàn phím", "Ổ cắm điện của phòng học"], "ans": 0},
        # 14
        {"bai": "Bài 6", "q": "Công dụng chính của Thư mục (Folder) trên máy tính là gì?", "options": ["Chứa các tệp và thư mục con để quản lý gọn gàng, dễ tìm", "Dùng để gõ chữ thay cho bàn phím máy tính", "Dùng để phát ra âm thanh thay cho chiếc loa", "Dùng để chiếu hình ảnh lên bảng trắng"], "ans": 0},
        # 15
        {"bai": "Bài 6", "q": "Một thư mục nằm bên trong một thư mục lớn hơn thì được gọi là gì?", "options": ["Thư mục con", "Bàn phím phụ", "Màn hình phụ", "Con trỏ phụ"], "ans": 0},
        # 16
        {"bai": "Bài 7", "q": "Khi đặt tên cho tệp hoặc thư mục trên máy tính Windows, em KHÔNG được dùng ký tự đặc biệt nào sau đây?", "options": ["Dấu hỏi chấm (?) hoặc dấu gạch chéo (/)", "Các chữ cái từ A đến Z", "Các chữ số từ 0 đến 9", "Dấu gạch dưới (_) nối các từ"], "ans": 0},
        # 17
        {"bai": "Bài 7", "q": "Tổ hợp phím tắt nhanh nào trên bàn phím thường dùng để Lưu (Save) sản phẩm đang làm?", "options": ["Ctrl + S", "Ctrl + Z", "Alt + F4", "Shift + 1"], "ans": 0},
        # 18
        {"bai": "Bài 8", "q": "Muốn mở lại một tệp bài vẽ cũ đã lưu trong thư mục TRANH_VE, bước đầu tiên em cần làm là gì?", "options": ["Tìm và nhấp đúp chuột mở đúng thư mục TRANH_VE", "Tắt máy tính rồi ngồi đợi năm phút", "Nhấn phím Backspace xóa tên thư mục TRANH_VE", "Mở phần mềm Máy tính bỏ túi Calculator"], "ans": 0},
        # 19
        {"bai": "Bài 9", "q": "Trước khi bắt đầu làm một dự án nhỏ gồm cả tranh vẽ và bài viết, em nên làm gì để sản phẩm có tổ chức?", "options": ["Tạo một thư mục riêng mang tên dự án để lưu chung các tệp liên quan", "Lưu rải rác mỗi tệp ở một ổ đĩa bất kỳ không cần nhớ chỗ", "Đặt tên mọi tệp đều là so_1 mà không phân biệt nội dung", "Không cần lưu tệp nào cho đỡ tốn chỗ của máy tính"], "ans": 0},
        # 20
        {"bai": "Bài 10", "q": "Trong một thư mục duy nhất, hai tệp cùng loại (ví dụ cùng là .png) có được đặt tên giống hệt nhau hoàn toàn không?", "options": ["Không được trùng tên hệt nhau trong cùng một thư mục", "Luôn bắt buộc phải đặt trùng một tên duy nhất", "Chỉ được đặt tên bằng một dấu chấm duy nhất", "Không được phép có chữ cái nào trong tên tệp"], "ans": 0},
        # 21
        {"bai": "Bài 12", "q": "Những con số, chữ viết, hình ảnh được lưu trữ bên trong máy tính để biểu diễn thông tin được gọi là gì?", "options": ["Dữ liệu", "Bàn ghế học sinh", "Dây cáp mạng", "Vỏ thân máy"], "ans": 0},
        # 22
        {"bai": "Bài 12", "q": "Khi em nhìn tấm bảng dự báo thời tiết có số “28°C” và hình ông mặt trời, điều em hiểu được “Hôm nay trời nắng ấm” là gì?", "options": ["Thông tin em nhận được", "Chiếc bàn phím của em", "Thư mục rỗng trên máy", "Khối lệnh lặp lại"], "ans": 0},
        # 23
        {"bai": "Bài 13", "q": "Bản thu âm bài hát Quốc ca phát qua loa trường vào sáng thứ Hai thuộc dạng thông tin nào?", "options": ["Dạng âm thanh", "Dạng chữ viết trên giấy", "Dạng tranh vẽ tĩnh", "Dạng sơ đồ khối"], "ans": 0},
        # 24
        {"bai": "Bài 13", "q": "Để giới thiệu cho bạn bè thấy rõ từng động tác múa của đội văn nghệ kèm tiếng nhạc sôi động, em nên chọn dạng thông tin nào?", "options": ["Dạng video", "Chỉ một dòng chữ ngắn", "Chỉ một con số đơn lẻ", "Chỉ một bức ảnh đen trắng"], "ans": 0},
        # 25
        {"bai": "Bài 14", "q": "Con người thu nhận thông tin từ thế giới xung quanh thông qua đâu?", "options": ["Các giác quan (mắt nhìn, tai nghe, mũi ngửi, lưỡi nếm, da cảm nhận)", "Chỉ qua duy nhất đôi giày đang đi dưới chân", "Chỉ qua chiếc cặp sách đeo trên vai", "Qua nút nguồn của màn hình máy tính"], "ans": 0},
        # 26
        {"bai": "Bài 14", "q": "Bộ phận nào trên cơ thể con người đảm nhận việc suy nghĩ và xử lý thông tin?", "options": ["Bộ não", "Mái tóc", "Móng tay", "Khuỷu tay"], "ans": 0},
        # 27
        {"bai": "Bài 14", "q": "Đối với máy tính để bàn, hai thiết bị quen thuộc dùng để đưa thông tin ĐẦU VÀO cho máy là gì?", "options": ["Bàn phím và chuột máy tính", "Màn hình và loa máy tính", "Máy in và máy chiếu", "Tấm lót chuột và ghế ngồi"], "ans": 0},
        # 28
        {"bai": "Bài 15", "q": "“Phân rã nhiệm vụ” trong môn Tin học lớp 2 có nghĩa là gì?", "options": ["Chia một nhiệm vụ lớn thành nhiều bước nhỏ theo thứ tự hợp lý", "Bỏ qua nhiệm vụ khó không thực hiện nữa", "Gộp mười bài tập khác nhau làm cùng một giây", "Xóa hết các tệp bài tập đang có trong thư mục"], "ans": 0},
        # 29
        {"bai": "Bài 15", "q": "Lợi ích lớn nhất của việc chia một nhiệm vụ lớn thành từng bước nhỏ là gì?", "options": ["Giúp công việc trở nên dễ hiểu, dễ làm và không bỏ sót bước", "Làm công việc mất nhiều thời gian gấp mười lần", "Làm chúng ta quên mất mục tiêu cuối cùng của bài", "Giúp máy tính tự chạy không cần bật nguồn điện"], "ans": 0},
        # 30
        {"bai": "Bài 16", "q": "Khi biểu diễn thuật toán bằng sơ đồ hình vẽ, ký hiệu nào thường dùng để nối các bước theo đúng trình tự thực hiện?", "options": ["Mũi tên chỉ hướng (➔ hoặc ⬇)", "Dấu chấm hỏi lớn (?)", "Hình thùng rác nhỏ", "Dấu nhân màu đỏ (X)"], "ans": 0},
        # 31
        {"bai": "Bài 16", "q": "Trong một thuật toán pha nước cam, nếu đảo lộn bước “Uống hết cốc nước” lên làm Bước 1 đầu tiên thì điều gì xảy ra?", "options": ["Thuật toán bị sai trình tự và không thể thực hiện được", "Cốc nước cam tự động ngọt hơn bình thường", "Không cần chuẩn bị quả cam và cốc thủy tinh nữa", "Thuật toán vẫn chạy đúng và nhanh hơn trước"], "ans": 0},
        # 32
        {"bai": "Bài 17", "q": "Trong lập trình trực quan, khi ghép nhiều khối lệnh nối tiếp nhau từ trên xuống dưới, nhân vật sẽ thực hiện như thế nào?", "options": ["Thực hiện lần lượt từng khối lệnh từ trên xuống dưới", "Thực hiện khối lệnh cuối cùng trước rồi dừng hẳn", "Chỉ thực hiện các khối lệnh có màu tím", "Thực hiện ngẫu nhiên lúc lệnh trên lúc lệnh dưới"], "ans": 0},
        # 33
        {"bai": "Bài 17", "q": "Khối lệnh “Khi nhấn vào Lá cờ xanh” thường được đặt ở vị trí nào của chuỗi lệnh?", "options": ["Đặt ở vị trí đầu tiên trên cùng của chuỗi lệnh", "Đặt ở vị trí cuối cùng dưới đáy chuỗi lệnh", "Đặt tách rời ở góc xa không nối với lệnh nào", "Đặt bên trong tên tệp hình ảnh ngoài màn hình"], "ans": 0},
        # 34
        {"bai": "Bài 17", "q": "Khối lệnh “Nói 'Xin chào!' trong 2 giây” giúp nhân vật làm gì trên sân khấu?", "options": ["Hiện bóng thoại có chữ 'Xin chào!' trong 2 giây", "Di chuyển lùi về phía sau đúng 2 bước chân", "Xoay tròn tại chỗ liên tục trong 2 phút", "Biến mất hoàn toàn khỏi sân khấu chương trình"], "ans": 0},
        # 35
        {"bai": "Bài 18", "q": "Hoạt động nào trong đời sống hằng ngày mang ý tưởng “Lặp lại một hành động”?", "options": ["Vỗ tay theo nhịp bài hát 4 lần liên tiếp", "Mở công tắc đèn phòng học 1 lần duy nhất", "Viết họ và tên vào nhãn vở 1 lần", "Đội mũ bảo hiểm lên đầu trước khi ngồi xe"], "ans": 0},
        # 36
        {"bai": "Bài 18", "q": "Sử dụng khối lệnh “Lặp lại” khi lập trình mang lại lợi ích gì cho chuỗi lệnh?", "options": ["Giúp chuỗi lệnh ngắn gọn, rõ ràng, đỡ phải ghép lặp nhiều khối giống nhau", "Làm chương trình dài gấp ba lần so với ban đầu", "Làm nhân vật không thể di chuyển trên sân khấu", "Xóa hết âm thanh có trong câu chuyện số"], "ans": 0},
        # 37
        {"bai": "Bài 19", "q": "Sau khi ghép xong chuỗi lệnh cho nhân vật, bước tiếp theo em luôn cần làm là gì?", "options": ["Nhấn chạy thử để kiểm tra xem nhân vật đã làm đúng yêu cầu chưa", "Đóng ngay phần mềm lại mà không cần xem kết quả", "Kéo rời từng khối lệnh ra khắp màn hình", "Đổi tên thư mục của bạn ngồi bên cạnh"], "ans": 0},
        # 38
        {"bai": "Bài 19", "q": "Khi chạy thử chương trình mà thấy chú Mèo rẽ trái thay vì rẽ phải tới quả bóng, nguyên nhân là gì?", "options": ["Do trong chuỗi lệnh có khối lệnh rẽ hướng bị chọn nhầm", "Do màn hình máy tính bị đặt lệch trên mặt bàn", "Do bàn phím máy tính bị thiếu phím dấu cách", "Do chuột máy tính có dây nối quá dài"], "ans": 0},
        # 39
        {"bai": "Bài 19", "q": "Thái độ đúng đắn của người lập trình nhí khi phát hiện chương trình mình ghép bị lỗi là gì?", "options": ["Kiên nhẫn kiểm tra từng bước lệnh, tìm chỗ sai và sửa lại", "Bỏ cuộc ngay lập tức và không bao giờ mở lại bài", "Đổ lỗi cho máy tính bị hỏng và đập vào bàn phím", "Xóa bài của bạn bên cạnh để không ai làm xong"], "ans": 0},
        # 40
        {"bai": "Bài 10", "q": "Khi cần tạo một thư mục mới trong cửa sổ quản lý tệp, em thường chọn lệnh nào?", "options": ["New Folder (Thư mục mới)", "Delete (Xóa bỏ)", "Shut down (Tắt máy)", "Print (In ra giấy)"], "ans": 0},
    ],

    "review_th_mcq": [
        # 41
        {"bai": "Bài 3", "q": "Bạn Ngọc muốn gõ câu “Chúc mừng sinh nhật mẹ!” có dấu chấm than (!) nằm ở phía trên phím số 1. Ngọc cần bấm như thế nào để ra dấu chấm than?", "options": ["Nhấn giữ phím Shift đồng thời gõ phím số 1", "Nhấn giữ phím Space đồng thời gõ phím số 1", "Nhấn phím Backspace hai lần rồi gõ số 1", "Nhấn phím Caps Lock rồi gõ phím chữ M"], "ans": 0},
        # 42
        {"bai": "Bài 4", "q": "Nếu em nhấp chuột vào nút dấu trừ (-) ở góc trên bên phải cửa sổ phần mềm Vẽ tranh, bài vẽ của em có bị mất không?", "options": ["Không bị mất, cửa sổ chỉ tạm thu nhỏ xuống thanh tác vụ bên dưới", "Bị xóa sạch hoàn toàn khỏi bộ nhớ máy tính", "Tự động biến thành tệp âm thanh bài hát", "Máy tính lập tức khởi động lại từ đầu"], "ans": 0},
        # 43
        {"bai": "Bài 6", "q": "Em có 6 tệp tranh vẽ và 6 tệp bài thơ để chung lộn xộn ngoài màn hình nền. Cách sắp xếp nào khoa học nhất?", "options": ["Tạo thư mục HOC_TAP gồm 2 thư mục con TRANH_VE và BAI_THO rồi cất đúng loại tệp vào", "Đổi tên cả 12 tệp thành các con số ngẫu nhiên rồi để nguyên ngoài màn hình", "Kéo cả 12 tệp bỏ hết vào Thùng rác (Recycle Bin) cho màn hình trống", "Tạo 50 thư mục rỗng không có tên rồi để mặc các tệp ở chỗ cũ"], "ans": 0},
        # 44
        {
            "bai": "Bài 8",
            "image_key": "g2_tep_va_thu_muc",
            "image_caption": "Hình minh họa sơ đồ thư mục HOC_TAP_LOP_2 (dùng cho Câu 44)",
            "q": "Quan sát sơ đồ thư mục HOC_TAP_LOP_2 ở trên, phát biểu nào sau đây là ĐÚNG?",
            "options": [
                "Thư mục HOC_TAP_LOP_2 chứa 2 thư mục con và tổng cộng 4 tệp sản phẩm số",
                "Thư mục TRANH_VE chứa tệp văn bản Noi_quy_lop.docx",
                "Thư mục BAI_SOAN chứa hai tệp hình ảnh có đuôi .png",
                "Tệp Gia_dinh_em.png là thư mục mẹ chứa thư mục HOC_TAP_LOP_2",
            ],
            "ans": 0,
        },
        # 45
        {"bai": "Bài 7", "q": "Bạn Hải lưu bài vẽ chiếc thuyền buồm với tên tệp là “asdf123.png”. Một tuần sau Hải tìm mãi không nhớ bài vẽ ở tệp nào. Bài học rút ra là gì?", "options": ["Cần đặt tên tệp rõ nghĩa, gợi nhớ đúng nội dung sản phẩm (như Thuyen_buom.png)", "Lần sau chỉ nên đặt tên tệp bằng một chữ cái duy nhất", "Máy tính không cho phép mở lại bài vẽ sau một tuần", "Không bao giờ được lưu tệp hình ảnh vào trong thư mục"], "ans": 0},
        # 46
        {"bai": "Bài 9", "q": "Bạn Tùng muốn lưu bài mới vào thư mục LOP_2A1 nhưng chưa mở thư mục đó trong hộp thoại Save As mà đã bấm Save ngay. Điều gì xảy ra?", "options": ["Tệp bài làm bị lưu nhầm sang vị trí thư mục khác đang mở sẵn", "Máy tính tự động xóa thư mục LOP_2A1 đi", "Bài làm không được lưu mà tự động in ra giấy", "Màn hình máy tính bị khóa không dùng chuột được nữa"], "ans": 0},
        # 47
        {"bai": "Bài 12", "q": "Trên tờ phiếu điểm kiểm tra có ghi con số “10” màu đỏ, nhờ đó em biết mình đạt kết quả xuất sắc. Con số “10” được viết trên phiếu đóng vai trò là gì?", "options": ["Là dữ liệu biểu diễn thông tin điểm số của em", "Là thiết bị phần cứng của máy tính để bàn", "Là một thư mục con nằm trong ổ đĩa D", "Là khối lệnh lặp lại trong phần mềm lập trình"], "ans": 0},
        # 48
        {"bai": "Bài 13", "q": "Thầy giáo muốn gửi cho cả lớp nghe thử giai điệu tiếng đàn piano để đoán tên bài hát (không lộ hình ảnh tên bài). Thầy nên chọn dạng thông tin nào?", "options": ["Dạng âm thanh", "Dạng chữ viết in to tên bài hát", "Dạng hình chụp bìa bản nhạc có ghi tên bài", "Dạng sơ đồ cây thư mục"], "ans": 0},
        # 49
        {"bai": "Bài 14", "q": "Khi đi đến ngã tư, mắt bố nhìn thấy đèn giao thông chuyển màu vàng (1), bố suy nghĩ cần giảm tốc độ dừng trước vạch (2), tay bố bóp phanh dừng xe an toàn (3). Thứ tự (1) -> (2) -> (3) tương ứng với gì?", "options": ["Đầu vào (1) -> Xử lý (2) -> Kết quả hành động (3)", "Xử lý (1) -> Đầu vào (2) -> Kết quả hành động (3)", "Kết quả (1) -> Xử lý (2) -> Đầu vào (3)", "Đầu vào (1) -> Kết quả (2) -> Tắt nguồn (3)"], "ans": 0},
        # 50
        {"bai": "Bài 15", "q": "Nhiệm vụ “Làm tấm thiệp chúc mừng trên máy tính” gồm: (1) Lưu tệp thiệp vào thư mục; (2) Mở phần mềm vẽ và vẽ khung hoa; (3) Gõ lời chúc tốt đẹp lên thiệp. Thứ tự hợp lý nhất là gì?", "options": ["(2) -> (3) -> (1)", "(1) -> (2) -> (3)", "(3) -> (1) -> (2)", "(1) -> (3) -> (2)"], "ans": 0},
        # 51
        {"bai": "Bài 16", "q": "Thuật toán điều khiển Robot lấy sách gồm: “Bước 1: Tiến thẳng 3 ô -> Bước 2: Rẽ phải -> Bước 3: Tiến thẳng 1 ô -> Bước 4: Nhặt quyển sách”. Nếu bỏ quên Bước 2 thì chuyện gì xảy ra?", "options": ["Robot đi thẳng tổng cộng 4 ô lệch khỏi vị trí quyển sách", "Robot vẫn nhặt trúng quyển sách nhanh hơn ban đầu", "Robot tự động lùi về ô xuất phát ban đầu", "Quyển sách tự động di chuyển đến chỗ Robot"], "ans": 0},
        # 52
        {"bai": "Bài 17", "q": "Chú Cún thực hiện chuỗi lệnh: “Đi tới 2 ô -> Nói 'Gâu gâu!' -> Đi tới 3 ô”. Hỏi chú Cún nói 'Gâu gâu!' vào thời điểm nào?", "options": ["Sau khi đi được 2 ô đầu tiên và trước khi đi tiếp 3 ô sau", "Ngay trước khi bắt đầu bước đi ô đầu tiên", "Chỉ sau khi đã đi hết cả 5 ô đường thẳng", "Nói liên tục ở từng ô vuông trên suốt đường đi"], "ans": 0},
        # 53
        {
            "bai": "Bài 18",
            "image_key": "g2_thuat_toan_chuoi_lenh",
            "image_caption": "Hình minh họa thuật toán chuỗi lệnh & vòng lặp (dùng cho Câu 53)",
            "q": "Quan sát hình minh họa trên, nếu ở Bước 2 ta sửa từ “Lặp lại 4 lần” thành “Lặp lại 2 lần” thì nhân vật sẽ thực hiện mấy lần cặp lệnh [Đi thẳng 10 bước -> Rẽ phải góc vuông]?",
            "options": [
                "Thực hiện 2 lần (đi được 2 cạnh góc vuông)",
                "Thực hiện 4 lần (đi đủ 4 cạnh hình vuông)",
                "Thực hiện 10 lần liên tiếp không dừng",
                "Không thực hiện lần nào mà phát tiếng vỗ tay ngay",
            ],
            "ans": 0,
        },
        # 54
        {"bai": "Bài 18", "q": "Chuỗi hành động của Robot là: “Vỗ tay -> Nhảy cao -> Vỗ tay -> Nhảy cao -> Vỗ tay -> Nhảy cao”. Ta viết gọn chuỗi này bằng khối lệnh lặp như thế nào?", "options": ["Lặp lại 3 lần [ Vỗ tay -> Nhảy cao ]", "Lặp lại 6 lần [ Vỗ tay -> Nhảy cao ]", "Lặp lại 2 lần [ Vỗ tay -> Vỗ tay ]", "Lặp lại 3 lần [ Đứng yên -> Ngồi xuống ]"], "ans": 0},
        # 55
        {"bai": "Bài 19", "q": " Kho báu cách Robot 4 ô về phía trước, nhưng chương trình đang viết: “Lặp lại 3 lần [ Tiến thẳng 1 ô ]”. Lỗi của chương trình này là gì và sửa thế nào?", "options": ["Thiếu 1 bước tiến; sửa số lần lặp từ 3 lần thành 4 lần", "Thừa 2 bước tiến; sửa số lần lặp từ 3 lần thành 1 lần", "Sai hướng đi; đổi lệnh Tiến thẳng thành lệnh Đi lùi", "Không có lỗi nào vì Robot tự nhảy thêm 1 ô cuối"], "ans": 0},
        # 56
        {"bai": "Bài 19", "q": "Bạn Châu lập trình cho nhân vật Cá Vàng bơi sang phải 5 bước rồi kêu “Ủn ỉn” như chú lợn con. Đây là lỗi ở khối lệnh nào?", "options": ["Lỗi chọn nhầm tệp âm thanh phát ra không phù hợp với nhân vật Cá Vàng", "Lỗi dùng sai số bước di chuyển sang bên phải", "Lỗi chưa bật màn hình máy tính trước khi chạy bài", "Lỗi đặt tên thư mục lưu bài quá ngắn gọn"], "ans": 0},
    ],

    "review_th_tl": [
        # 57
        {
            "bai": "Bài 4",
            "q": "Em hãy nêu sự khác nhau về kết quả khi nhấp chuột vào nút dấu trừ (-) và nút dấu nhân (X) ở góc trên bên phải cửa sổ ứng dụng.",
            "answer": "- Nhấp nút dấu trừ (-): Cửa sổ ứng dụng chỉ tạm thu nhỏ xuống thanh tác vụ (Taskbar), bài làm vẫn đang mở và nhấp lại là hiện lên. - Nhấp nút dấu nhân (X): Đóng hẳn cửa sổ và thoát khỏi ứng dụng đó (nên cần lưu bài trước khi đóng).",
        },
        # 58
        {
            "bai": "Bài 5",
            "q": "Tệp (File) trên máy tính là gì? Em hãy kể tên 2 loại tệp sản phẩm số mà học sinh lớp 2 thường tạo ra.",
            "answer": "Tệp là sản phẩm số được tạo ra từ các phần mềm và lưu trữ trong máy tính. Hai loại tệp quen thuộc với học sinh lớp 2 là: tệp tranh vẽ hình ảnh (ví dụ: Ngoi_nha.png) và tệp bài soạn thảo văn bản (ví dụ: Bai_tho.docx).",
        },
        # 59
        {
            "bai": "Bài 6",
            "q": "Vì sao người ta thường ví Thư mục (Folder) trên máy tính giống như các ngăn trong chiếc cặp sách đi học của em?",
            "answer": "Vì cũng giống như các ngăn cặp giúp phân loại sách Toán, vở Tiếng Việt, hộp bút riêng biệt cho gọn gàng, Thư mục trên máy tính dùng để cất giữ và phân loại các tệp theo từng môn học hoặc chủ đề giúp dễ quản lý và tìm kiếm nhanh.",
        },
        # 60
        {
            "bai": "Bài 7",
            "q": "Em hãy nêu 2 quy tắc quan trọng khi đặt tên và lưu sản phẩm số vào máy tính.",
            "answer": "(1) Đặt tên tệp ngắn gọn, rõ nghĩa, đúng nội dung bài làm và không chứa các ký tự bị cấm như / \\ : * ? \" < > |. (2) Chọn đúng thư mục cần lưu trước khi nhấn nút Save để không bị thất lạc bài.",
        },
        # 61
        {
            "bai": "Bài 12",
            "q": "Em hãy phân biệt “Thông tin” và “Dữ liệu” qua ví dụ tấm biển ghi dòng chữ: “Thư viện mở cửa lúc 8 giờ sáng”.",
            "answer": "- Dữ liệu: Là các chữ cái và con số “Thư viện mở cửa lúc 8 giờ sáng” được viết trên tấm biển. - Thông tin: Là sự hiểu biết của em sau khi đọc dòng chữ đó (biết được thời gian thư viện bắt đầu mở cửa để đến mượn sách đúng giờ).",
        },
        # 62
        {
            "bai": "Bài 13",
            "q": "Kể tên 4 dạng thông tin cơ bản mà em đã học ở lớp 2 và cho biết dạng nào sinh động nhất khi muốn ghi lại cả cảnh buổi lễ chào cờ lẫn tiếng hát Quốc ca.",
            "answer": "4 dạng thông tin cơ bản gồm: Dạng chữ, dạng hình ảnh, dạng âm thanh và dạng video. Để ghi lại sinh động cả hình ảnh lễ chào cờ và âm thanh tiếng hát Quốc ca, em chọn dạng video.",
        },
        # 63
        {
            "bai": "Bài 16",
            "q": "Thuật toán là gì? Có những cách nào quen thuộc để em trình bày một thuật toán cho bạn cùng lớp hiểu?",
            "answer": "Thuật toán là một dãy các bước hướng dẫn rõ ràng, theo đúng thứ tự từ đầu đến cuối để giải quyết một nhiệm vụ. Em có thể trình bày thuật toán bằng lời (liệt kê Bước 1, Bước 2...) hoặc bằng sơ đồ hình vẽ có mũi tên nối tiếp nhau.",
        },
        # 64
        {
            "bai": "Bài 18",
            "q": "Khi nào em nên sử dụng khối lệnh “Lặp lại” trong lập trình trực quan? Cho 1 ví dụ ngắn gọn.",
            "answer": "Em nên dùng khối lệnh “Lặp lại” khi nhân vật cần thực hiện cùng một hành động (hoặc cùng một nhóm lệnh giống hệt nhau) nhiều lần liên tiếp. Ví dụ: Thay vì ghép 4 khối “Vỗ tay”, em dùng khối “Lặp lại 4 lần [ Vỗ tay ]”.",
        },
    ],

    "review_vd_tl": [
        # 65
        {
            "bai": "Bài 9",
            "q": "Đầu học kì, em muốn tổ chức góc học tập trên máy tính thật ngăn nắp cho 3 môn: Tin học, Mỹ thuật (vẽ máy) và Tiếng Việt. Em hãy vẽ hoặc mô tả cây thư mục hợp lý của riêng em.",
            "answer": "Em tạo 1 thư mục gốc mang tên em (ví dụ: HOC_TAP_LOP_2_AN). Bên trong thư mục gốc đó, em tạo 3 thư mục con riêng biệt là: TIN_HOC, MY_THUAT và TIENG_VIET để lưu sản phẩm tương ứng của từng môn.",
        },
        # 66
        {
            "bai": "Bài 3",
            "q": "Em cần gõ câu văn ngắn: “Lớp 2A1 đoàn kết.” rồi xuống dòng gõ tên mình. Em hãy nêu cách dùng các phím Shift, Space, dấu chấm và Enter khi thực hiện.",
            "answer": "- Giữ phím Shift + L để viết hoa chữ L đầu câu và Shift + A để viết hoa chữ A trong 2A1. - Nhấn phím Space (dấu cách) đúng 1 lần giữa các từ. - Cuối câu nhấn phím dấu chấm (.), sau đó nhấn phím Enter để xuống dòng mới gõ tên mình.",
        },
        # 67
        {
            "bai": "Bài 14",
            "q": "Trời đang nắng bỗng kéo mây đen và có tiếng sấm rền. Mẹ bảo Nam chạy ra sân cất quần áo khô vào nhà. Em hãy xác định Đầu vào, Xử lý và Kết quả (Đầu ra) của bạn Nam.",
            "answer": "- Đầu vào: Mắt thấy mây đen, tai nghe tiếng sấm và lời mẹ dặn. - Xử lý: Bộ não Nam hiểu trời sắp mưa to, cần thu quần áo ngay kẻo ướt. - Kết quả (Đầu ra): Nam nhanh nhẹn chạy ra sân gom quần áo khô mang vào nhà.",
        },
        # 68
        {
            "bai": "Bài 15",
            "q": "Em hãy phân rã nhiệm vụ “Trồng và tưới một chậu cây nhỏ” thành 4 bước nhỏ rõ ràng theo đúng trình tự.",
            "answer": "Bước 1: Cho đất tơi xốp vào chậu và đào một hốc nhỏ ở giữa. Bước 2: Đặt bầu cây con thẳng đứng vào giữa hốc đất. Bước 3: Lấp đất nhẹ nhàng quanh gốc cây cho chắc chắn. Bước 4: Dùng bình tưới một lượng nước vừa đủ làm ẩm đất.",
        },
        # 69
        {
            "bai": "Bài 16",
            "q": "Hãy viết thuật toán bằng lời (từng bước từ đầu đến cuối) hướng dẫn một bạn mới biết cách “Tạo thư mục TRANH_VE và lưu bức tranh Hoa_sen.png vào đúng thư mục đó”.",
            "answer": "Bước 1: Mở ổ đĩa học tập, chọn lệnh New Folder và đặt tên thư mục là TRANH_VE. Bước 2: Mở phần mềm Paint và vẽ bức tranh hoa sen. Bước 3: Nhấn nút Save (hoặc Ctrl + S), tìm nhấp đúp mở thư mục TRANH_VE. Bước 4: Gõ tên tệp là Hoa_sen rồi nhấn nút Save hoàn tất.",
        },
        # 70
        {
            "bai": "Bài 17",
            "q": "Trên sân khấu lập trình, chú Mèo đang đứng bên trái quả bóng 3 bước. Em hãy viết chuỗi lệnh từ lúc bấm Lá cờ xanh để chú Mèo đi tới quả bóng, sút bóng (phát tiếng “Bốp”) và nói “Vào rồi!”.",
            "answer": "Chuỗi lệnh theo thứ tự từ trên xuống: Bước 1: Khi nhấn vào Lá cờ xanh -> Bước 2: Đi sang phải 3 bước (hoặc ghép 3 khối Sang phải 1 bước) -> Bước 3: Phát âm thanh “Bốp” -> Bước 4: Nói “Vào rồi!” trong 2 giây.",
        },
        # 71
        {
            "bai": "Bài 18",
            "q": "Để vẽ một hình tam giác đều trên sân chơi, Robot thực hiện chuỗi lệnh dài sau: “Đi thẳng 5 bước -> Quay phải 1 góc -> Đi thẳng 5 bước -> Quay phải 1 góc -> Đi thẳng 5 bước -> Quay phải 1 góc”. Em hãy viết gọn lại chuỗi lệnh này bằng khối Lặp lại.",
            "answer": "Ta thấy cặp 2 lệnh [Đi thẳng 5 bước -> Quay phải 1 góc] được lặp đi lặp lại đúng 3 lần liên tiếp. Vậy chuỗi lệnh rút gọn là: Lặp lại 3 lần [ Đi thẳng 5 bước -> Quay phải 1 góc ].",
        },
        # 72
        {
            "bai": "Bài 13",
            "q": "Nhóm của em làm bài giới thiệu về “An toàn giao thông trước cổng trường”. Hãy nêu cách chọn dạng thông tin phù hợp cho từng việc: (a) Tên tiêu đề bài; (b) Minh họa cách đội mũ bảo hiểm cài quai đúng; (c) Ghi lại cảnh học sinh xếp hàng ra về trật tự.",
            "answer": "- Ý (a) Tên tiêu đề bài: Chọn thông tin dạng chữ to, rõ ràng. - Ý (b) Minh họa cài quai mũ bảo hiểm: Chọn thông tin dạng hình ảnh chụp/vẽ các bước. - Ý (c) Cảnh xếp hàng ra về trật tự: Chọn thông tin dạng video sinh động.",
        },
    ],

    "review_vdc_tl": [
        # 73
        {
            "bai": "Bài 19",
            "q": "Trên lưới ô vuông, Kho báu nằm ở vị trí đi thẳng 3 ô rồi rẽ phải đi tiếp 2 ô. Bạn Huy viết chương trình: “Bước 1: Khi bấm Lá cờ xanh -> Bước 2: Lặp lại 3 lần [ Đi thẳng 1 ô -> Rẽ phải ] -> Bước 3: Đi thẳng 2 ô”. Hãy tìm lỗi sai ở Bước 2 và sửa lại cho đúng.",
            "answer": "- Lỗi sai ở Bước 2: Lệnh “Rẽ phải” bị đặt nhầm vào BÊN TRONG vòng lặp 3 lần khiến cứ đi 1 ô Robot lại rẽ phải ngay (đi thành vòng tròn chứ không tiến thẳng 3 ô). - Cách sửa đúng: Đưa lệnh “Rẽ phải” ra ngoài vòng lặp -> Bước 2: Lặp lại 3 lần [ Đi thẳng 1 ô ] -> Bước 3: Rẽ phải -> Bước 4: Lặp lại 2 lần [ Đi thẳng 1 ô ].",
        },
        # 74
        {
            "bai": "Bài 18",
            "q": "Đường đi của Robot có dạng bậc cầu thang gồm 3 bậc giống nhau: mỗi bậc phải “Đi lên 1 ô” rồi “Sang phải 1 ô” để tới đích trên đỉnh cầu thang. Em hãy viết chương trình ngắn gọn nhất dùng khối lệnh Lặp lại để đưa Robot lên đỉnh cầu thang.",
            "answer": "Vì mỗi bậc thang đều gồm đúng cặp lệnh [Đi lên 1 ô -> Sang phải 1 ô] và có 3 bậc liên tiếp, chương trình ngắn gọn nhất là: Bước 1: Khi nhấn vào Lá cờ xanh -> Bước 2: Lặp lại 3 lần [ Đi lên 1 ô -> Sang phải 1 ô ].",
        },
        # 75
        {
            "bai": "Bài 10",
            "q": "Máy tính lớp học dùng chung cho hai lớp 2A1 và 2C1. Hôm trước bạn Minh lớp 2A1 lưu tệp “Bai_ve.png” ngay ngoài màn hình nền, hôm nay bạn Minh lớp 2C1 cũng lưu tệp tên “Bai_ve.png” ra màn hình nền và bấm Ghi đè (Replace). Chuyện gì đã xảy ra và làm sao để hai lớp không bao giờ bị mất bài của nhau?",
            "answer": "- Hậu quả: Bài vẽ cũ của bạn Minh lớp 2A1 đã bị ghi đè mất đi. - Giải pháp triệt để: Tạo hai thư mục riêng LOP_2A1 và LOP_2C1 trong ổ đĩa học tập; mỗi học sinh lưu vào đúng thư mục lớp mình và gắn thêm họ tên vào tên tệp (ví dụ: Bai_ve_Nguyen_Minh_2A1.png).",
        },
        # 76
        {
            "bai": "Bài 16",
            "q": "Cho thuật toán kiểm tra cặp sách trước khi đến trường gồm: Bước 1: Mở cặp sách và lấy thời khóa biểu hôm nay ra xem -> Bước 2: Kiểm tra từng quyển sách, vở xem đã đủ môn chưa -> Bước 3: Nếu thiếu môn nào thì lấy bổ sung ngay vào cặp -> Bước 4: Kiểm tra hộp bút và đóng cặp lại. Điều gì xảy ra nếu bỏ qua Bước 3?",
            "answer": "Nếu bỏ qua Bước 3, dù ở Bước 2 em đã phát hiện ra bị thiếu sách vở nhưng em lại đóng cặp đi học luôn mà không lấy bổ sung, dẫn đến lên lớp vẫn không có sách vở để học bài môn đó.",
        },
        # 77
        {
            "bai": "Bài 19",
            "q": "Nhân vật Chú Chó đang quay mặt sang bên phải. Sau khi chạy chương trình: “Lặp lại 2 lần [ Đi thẳng 2 ô -> Quay ngược 180 độ ]”, chú Chó đang đứng cách vị trí xuất phát mấy ô và quay mặt về hướng nào? Giải thích.",
            "answer": "Chú Chó đứng cách vị trí xuất phát 0 ô (quay về đúng vị trí ban đầu) và quay mặt về hướng bên phải như lúc đầu. Giải thích: Lần lặp 1 chú đi sang phải 2 ô rồi quay mặt sang trái; Lần lặp 2 chú đi ngược sang trái 2 ô về chỗ cũ rồi quay 180 độ lại hướng sang phải.",
        },
        # 78
        {
            "bai": "Bài 14",
            "q": "Em hãy so sánh điểm giống nhau giữa quá trình xử lý thông tin của Con người (khi giải bài Toán đố) và của Máy tính (khi chạy chương trình tính tổng) theo mô hình 3 thành phần.",
            "answer": "Điểm giống nhau là cả con người và máy tính đều thực hiện qua 3 bước: (1) Đầu vào: Tiếp nhận đề bài / dữ liệu đưa vào; (2) Xử lý: Bộ não con người suy nghĩ tính toán tương ứng với thân máy (CPU) xử lý theo lệnh; (3) Kết quả (Đầu ra): Viết/nói đáp số tương ứng với máy tính hiển thị kết quả lên màn hình.",
        },
        # 79
        {
            "bai": "Bài 18",
            "q": "Robot cần nhặt 3 ngôi sao nằm thẳng hàng ở các ô thứ 2, ô thứ 4 và ô thứ 6 trước mặt (mỗi ngôi sao cách nhau đúng 2 ô, Robot xuất phát ở ô 0). Em hãy thiết kế chuỗi lệnh dùng vòng lặp để Robot đi nhặt hết cả 3 ngôi sao.",
            "answer": "Mỗi lần nhặt 1 ngôi sao, Robot đều thực hiện giống nhau: tiến thẳng 2 ô rồi nhặt sao. Vậy chuỗi lệnh tối ưu là: Bước 1: Khi nhấn Lá cờ xanh -> Bước 2: Lặp lại 3 lần [ Đi thẳng 1 ô -> Đi thẳng 1 ô -> Nhặt ngôi sao ] (hoặc Lặp lại 3 lần [ Đi thẳng 2 ô -> Nhặt ngôi sao ]).",
        },
        # 80
        {
            "bai": "Bài 19",
            "q": "Em hãy tự sáng tạo kịch bản ngắn 4 bước cho nhân vật “Chú Ong Chăm Chỉ” đi hút mật 2 bông hoa nằm cách đều nhau 3 bước chân rồi mang mật về tổ reo vui, trong đó có sử dụng ít nhất 1 khối lệnh Lặp lại.",
            "answer": "Kịch bản thuật toán & chuỗi lệnh hợp lệ: Bước 1: Khi nhấn vào Lá cờ xanh (Chú Ong ở vạch xuất phát). Bước 2: Lặp lại 2 lần [ Bay thẳng 3 bước -> Hút mật hoa trong 1 giây ]. Bước 3: Bay tiếp 2 bước về tới Tổ ong. Bước 4: Phát âm thanh vui nhộn và nói “Tổ ong đầy mật ngọt rồi!”.",
        },
    ],
}
