# -*- coding: utf-8 -*-
"""Ngân hàng Đề kiểm tra & Đề cương 80 câu ôn tập Cuối học kì 1 (ĐGĐK 2) - Môn Tin học Lớp 1 UNIGO."""

GRADE_DATA = {
    "lop": 1,
    "cap_hoc": "Tiểu học",
    "lop_hien_thi": "1A1, 1C1",
    "thoi_gian": "35 phút",
    "tuan_on_tap": 20,
    "tuan_kiem_tra": 21,
    "tiet_on_tap": 19,
    "tiet_kiem_tra": 20,

    "tom_tat_ly_thuyet": [
        {
            "bai": "Chủ đề 1 (Bài 1 – Bài 3): Làm quen với máy tính, các bộ phận và tư thế an toàn",
            "image_key": "g1_bo_phan_may_tinh",
            "y_chinh": [
                "Máy tính giúp em học tập, vẽ tranh, nghe nhạc, xem video bổ ích và liên lạc với người thân.",
                "Máy tính để bàn có 4 bộ phận chính: Màn hình (hiển thị hình ảnh, chữ viết), Thân máy (bộ não xử lý của máy tính), Bàn phím (gõ chữ và số), Chuột máy tính (điều khiển con trỏ chọn đối tượng).",
                "Nội quy phòng Tin học: Đi nhẹ, nói khẽ, ngồi đúng chỗ, không mang đồ ăn thức uống vào phòng máy, bật và tắt máy tính đúng hướng dẫn của thầy cô.",
                "Tư thế ngồi đúng và an toàn: Ngồi thẳng lưng, hai vai thả lỏng, mắt cách màn hình từ 50 cm đến 80 cm, không dùng tay ướt chạm vào ổ điện hay dây điện.",
            ],
        },
        {
            "bai": "Chủ đề 2 (Bài 4 – Bài 5): Kỹ năng sử dụng chuột máy tính",
            "y_chinh": [
                "Cách cầm chuột đúng: Đặt bàn tay phải lên chuột, ngón trỏ đặt lên nút trái chuột, ngón giữa đặt lên nút phải chuột, các ngón còn lại giữ hai bên thân chuột.",
                "Di chuyển chuột: Giữ và trượt nhẹ chuột trên mặt phẳng để mũi tên (con trỏ chuột) di chuyển trên màn hình.",
                "Nhấp chuột (Click): Nhấn nhanh một lần nút trái chuột rồi thả ngón tay ra để chọn biểu tượng.",
                "Nhấp đúp chuột (Double-click): Nhấn nhanh hai lần liên tiếp nút trái chuột để mở chương trình.",
                "Kéo – thả chuột (Drag and Drop): Nhấn và giữ nút trái chuột vào đối tượng, di chuyển đến vị trí mới rồi thả ngón tay ra.",
            ],
        },
        {
            "bai": "Chủ đề 3 (Bài 6 – Bài 10): Bàn phím, gõ từ đơn giản và tạo sản phẩm số đầu tiên",
            "y_chinh": [
                "Trên bàn phím có các phím chữ cái (từ A đến Z) và các phím số (từ 0 đến 9) giúp em nhập chữ và số vào máy tính.",
                "Các phím đặc biệt em cần nhớ: Phím Cách (Space – phím dài nhất ở hàng dưới cùng) dùng tạo khoảng trắng giữa hai từ; phím Enter dùng xuống dòng; phím Backspace dùng xóa ký tự ngay bên trái con trỏ; phím Shift giữ kèm phím chữ để gõ chữ in hoa.",
                "Khi tạo sản phẩm số đầu tiên (như bức tranh vẽ hoặc dòng chữ tên em), em cần nhấn nút Lưu (Save) để cất giữ bài làm vào máy tính.",
            ],
        },
        {
            "bai": "Chủ đề 4 (Bài 12 – Bài 14): Thông tin quanh em, máy tính làm việc theo lệnh và chia nhỏ công việc",
            "y_chinh": [
                "Thông tin quanh em có 3 dạng quen thuộc: Dạng chữ (tên bài học, biển báo chữ), dạng hình ảnh (bức tranh, đèn giao thông, biển báo hình vẽ) và dạng âm thanh (tiếng trống trường, tiếng còi xe, tiếng chim hót).",
                "Máy tính không tự suy nghĩ như con người mà làm việc chính xác theo từng lệnh do con người đưa vào (như nhấp chuột mở bài vẽ, nhấn phím gõ chữ).",
                "Một công việc lớn thường gồm nhiều bước nhỏ thực hiện theo đúng thứ tự từ bước đầu đến bước cuối (ví dụ: Đánh răng, rửa tay, chuẩn bị cặp sách đi học).",
            ],
        },
        {
            "bai": "Chủ đề 5 (Bài 15 – Bài 19): Chỉ đường trong mê cung, sửa lỗi và làm quen lập trình trực quan",
            "image_key": "g1_me_cung_robot",
            "y_chinh": [
                "Để chỉ đường cho nhân vật hoặc Robot đi đến đích, em ghép các lệnh di chuyển (Lên trên 1 ô, Xuống dưới 1 ô, Sang trái 1 ô, Sang phải 1 ô) theo đúng thứ tự từng bước.",
                "Khi đi trong mê cung, nhân vật chỉ được đi vào các ô trống và phải tránh các ô có chướng ngại vật (như Đá cản, bức tường).",
                "Sửa lỗi đường đi (gỡ lỗi): Nếu nhân vật đi sai hướng hoặc chạm vật cản, em bình tĩnh tìm đúng bước lệnh bị sai và thay bằng lệnh đúng.",
                "Lập trình trực quan: Em dùng chuột kéo – thả và ghép các khối lệnh màu sắc lại với nhau thành một chuỗi lệnh hoàn chỉnh để điều khiển nhân vật.",
            ],
        },
    ],

    "ma_tran_kiem_tra": [
        ["Chủ đề 1: Làm quen máy tính, các bộ phận & an toàn (Bài 1 – 3)", "3 câu (1,5đ)", "1 câu (0,5đ)", "-", "-", "-", "2,0 điểm"],
        ["Chủ đề 2: Chuột, bàn phím & tạo sản phẩm số (Bài 4 – 10)", "4 câu (2,0đ)", "1 câu (0,5đ)", "1 câu (1,0đ)", "-", "-", "3,5 điểm"],
        ["Chủ đề 3: Thông tin quanh em & lệnh cho máy tính (Bài 12 – 14)", "1 câu (0,5đ)", "1 câu (0,5đ)", "-", "1 câu (1,0đ)", "-", "2,0 điểm"],
        ["Chủ đề 4: Mê cung, sửa lỗi & lập trình trực quan (Bài 15 – 19)", "2 câu (1,0đ)", "1 câu (0,5đ)", "-", "-", "1 câu (1,0đ)", "2,5 điểm"],
        ["TỔNG CỘNG", "10 câu (5,0đ)", "4 câu (2,0đ)", "1 câu (1,0đ)", "1 câu (1,0đ)", "1 câu (1,0đ)", "10,0 điểm"],
    ],

    # Đề kiểm tra chính thức: ĐÚNG 14 câu Trắc nghiệm (10 NB + 4 TH) = 7,0 điểm
    "exam_mcq": [
        # --- 10 câu Nhận biết (NB) ---
        {
            "level": "NB",
            "bai": "Bài 1",
            "q": "Việc làm nào sau đây được phép thực hiện trong phòng Tin học?",
            "options": [
                "Ngồi ngay ngắn và nghe lời thầy cô hướng dẫn",
                "Mang bánh kẹo và nước ngọt vào bàn máy tính",
                "Chạy nhảy, đùa nghịch quanh dây điện máy tính",
                "Tự ý rút dây cắm điện phía sau thân máy",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 2",
            "image_key": "g1_bo_phan_may_tinh",
            "image_caption": "Hình 1. Các bộ phận của máy tính để bàn (dùng để trả lời Câu 2)",
            "q": "Quan sát Hình 1 ở trên, bộ phận được đánh số (1) có tên gọi là gì?",
            "options": [
                "Màn hình máy tính",
                "Thân máy tính",
                "Bàn phím máy tính",
                "Chuột máy tính",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 3",
            "q": "Khi ngồi học với máy tính, lưng của em cần giữ ở tư thế nào để không bị cong vẹo cột sống?",
            "options": [
                "Ngồi thẳng lưng, hai vai thả lỏng tự nhiên",
                "Cúi sát mặt chạm vào màn hình máy tính",
                "Nằm dài người lên mặt bàn phím máy tính",
                "Nghiêng hẳn người sang một bên ghế ngồi",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 4",
            "q": "Khi cầm chuột máy tính bằng tay phải, ngón trỏ của em đặt lên vị trí nào?",
            "options": [
                "Đặt lên nút trái chuột",
                "Đặt lên nút phải chuột",
                "Đặt ở mặt dưới đáy chuột",
                "Đặt lên dây nối của chuột",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 5",
            "q": "Thao tác nhấn và giữ nút trái chuột, di chuyển đối tượng đến chỗ mới rồi thả tay ra được gọi là gì?",
            "options": [
                "Kéo – thả chuột",
                "Nhấp chuột phải",
                "Nhấp đúp chuột",
                "Tắt nguồn máy tính",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 6",
            "q": "Trên bàn phím máy tính có những nhóm phím quen thuộc nào giúp em gõ tên và tuổi của mình?",
            "options": [
                "Các phím chữ cái và các phím số",
                "Các nút chỉnh âm lượng của tivi",
                "Các bánh xe di chuyển của robot",
                "Các đầu cắm dây nguồn máy tính",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 7",
            "q": "Phím dài nhất nằm ở hàng dưới cùng của bàn phím, dùng để tạo khoảng cách giữa hai từ là phím nào?",
            "options": [
                "Phím Cách (Space)",
                "Phím Xuống dòng (Enter)",
                "Phím Xóa chữ (Backspace)",
                "Phím Chữ hoa (Shift)",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 12",
            "q": "Tiếng trống trường báo hiệu giờ vào lớp là thông tin ở dạng nào?",
            "options": [
                "Dạng âm thanh",
                "Dạng chữ viết",
                "Dạng hình ảnh",
                "Dạng mùi hương",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 15",
            "q": "Trong trò chơi chỉ đường trên lưới ô vuông, lệnh “Sang phải 1 ô” giúp nhân vật làm gì?",
            "options": [
                "Di chuyển sang ô ngay bên phải một bước",
                "Di chuyển lên ô ngay phía trên một bước",
                "Di chuyển xuống ô ngay phía dưới một bước",
                "Đứng yên tại chỗ không di chuyển",
            ],
            "ans": 0,
        },
        {
            "level": "NB",
            "bai": "Bài 18",
            "q": "Trong môi trường lập trình trực quan cho học sinh lớp 1, em tạo chương trình bằng cách nào?",
            "options": [
                "Kéo – thả và ghép các khối lệnh màu sắc với nhau",
                "Dùng bút màu vẽ trực tiếp lên màn hình máy tính",
                "Tháo rời bàn phím ra khỏi thân máy tính",
                "Tắt màn hình máy tính rồi bật lại nhiều lần",
            ],
            "ans": 0,
        },
        # --- 4 câu Thông hiểu (TH) ---
        {
            "level": "TH",
            "bai": "Bài 3",
            "q": "Vì sao em không được dùng tay ướt chạm vào nút nguồn hay dây điện của máy tính?",
            "options": [
                "Vì dễ gây chập điện và nguy hiểm cho bản thân",
                "Vì làm chữ trên màn hình bị đổi sang màu đỏ",
                "Vì làm bàn phím tự động gõ ra các con số",
                "Vì làm con trỏ chuột chạy nhanh gấp đôi",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 8",
            "q": "Bạn An muốn gõ chữ cái “A” in hoa ở đầu tên mình. Bạn An nên thực hiện thao tác nào?",
            "options": [
                "Nhấn giữ phím Shift đồng thời gõ phím chữ A",
                "Nhấn giữ phím Space đồng thời gõ phím chữ A",
                "Nhấn phím Backspace hai lần liên tiếp",
                "Nhấn phím Enter ba lần liên tiếp",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 13",
            "q": "Khi em nhấp đúp chuột vào biểu tượng phần mềm vẽ tranh trên màn hình nền, điều gì sẽ xảy ra?",
            "options": [
                "Máy tính nhận lệnh và mở cửa sổ phần mềm vẽ tranh",
                "Máy tính tự động tắt nguồn ngay lập tức",
                "Bàn phím tự động phát ra bản nhạc thiếu nhi",
                "Màn hình máy tính tự động thu nhỏ thành chiếc đồng hồ",
            ],
            "ans": 0,
        },
        {
            "level": "TH",
            "bai": "Bài 17",
            "q": "Bạn Minh xếp chuỗi lệnh cho chú mèo đi lấy bóng nhưng chú mèo lại đâm vào tường. Minh cần làm gì?",
            "options": [
                "Xem lại chuỗi lệnh, tìm bước lệnh sai và đổi lại cho đúng",
                "Xóa bỏ phần mềm học tập khỏi máy tính của trường",
                "Nhấn thật mạnh vào màn hình để đẩy chú mèo đi tiếp",
                "Rút dây chuột máy tính ra để chú mèo tự tìm bóng",
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
            "q": "Em hãy nối hoặc ghi rõ chức năng của 4 phím đặc biệt sau đây trên bàn phím: (a) Phím Space (phím Cách); (b) Phím Enter; (c) Phím Backspace; (d) Phím Shift.",
            "rubric": [
                "Nêu đúng chức năng phím Space (phím Cách): Tạo khoảng trắng (dấu cách) giữa hai từ khi gõ chữ (0,25 điểm).",
                "Nêu đúng chức năng phím Enter: Xuống dòng mới hoặc xác nhận thực hiện lệnh (0,25 điểm).",
                "Nêu đúng chức năng phím Backspace: Xóa ký tự viết sai nằm ngay bên trái con trỏ soạn thảo (0,25 điểm).",
                "Nêu đúng chức năng phím Shift: Nhấn giữ kèm phím chữ cái để gõ chữ in hoa (0,25 điểm).",
            ],
        },
        {
            "level": "Vận dụng",
            "points": "1,0 điểm",
            "lines": 4,
            "q": "Buổi sáng thức dậy, việc đánh răng gồm 4 bước đang bị xáo trộn thứ tự như sau: (1) Súc miệng thật sạch với nước; (2) Lấy bàn chải và bôi kem đánh răng; (3) Chải đều mặt ngoài, mặt trong và mặt nhai của răng; (4) Rửa sạch bàn chải và cất gọn lên giá. Em hãy sắp xếp lại thứ tự đúng của các bước và giải thích ngắn gọn vì sao cần làm đúng thứ tự đó.",
            "rubric": [
                "Sắp xếp đúng thứ tự 4 bước: Bước (2) -> Bước (3) -> Bước (1) -> Bước (4) (hoặc chấp nhận (2) -> (3) -> (4) -> (1)) (0,5 điểm).",
                "Giải thích hợp lý: Làm đúng thứ tự từng bước giúp răng sạch sẽ, không bỏ sót việc và hoàn thành nhiệm vụ nhanh gọn (0,5 điểm).",
            ],
        },
        {
            "level": "Vận dụng cao",
            "points": "1,0 điểm",
            "lines": 5,
            "image_key": "g1_me_cung_robot",
            "image_caption": "Hình 2. Mê cung Robot tìm kho báu (dùng để làm Câu 3 phần Tự luận)",
            "q": "Quan sát Hình 2 (Mê cung Robot tìm kho báu), biết Robot đang ở ô Xuất phát (góc trên bên trái) và KHÔNG được đi vào các ô ĐÁ CẢN màu đỏ. Em hãy viết theo thứ tự từng bước các lệnh điều khiển (Đi xuống 1 ô, Sang phải 1 ô, Đi lên 1 ô, Sang trái 1 ô) để đưa Robot đến đúng ô KHO BÁU.",
            "rubric": [
                "Xác định đúng hướng xuất phát tránh ô Đá cản: Đi xuống 1 ô -> Đi xuống 1 ô -> Đi xuống 1 ô (đưa Robot xuống ô góc dưới bên trái) (0,5 điểm).",
                "Viết đúng các bước tiếp theo đến Kho báu: Sang phải 1 ô -> Sang phải 1 ô -> Đi lên 1 ô (0,25 điểm).",
                "Trình bày đầy đủ chuỗi 6 bước lệnh theo đúng trình tự từ đầu đến cuối, không đi vào bất kỳ ô ĐÁ CẢN nào (0,25 điểm).",
            ],
        },
    ],

    # ═════════════════════════════════════════════════════════════════════════
    # NGÂN HÀNG 80 CÂU HỎI ÔN TẬP CUỐI HỌC KÌ 1 (40 NB_TN + 16 TH_TN + 8 TH_TL + 8 VD_TL + 8 VDC_TL)
    # ═════════════════════════════════════════════════════════════════════════
    "review_nb_mcq": [
        # 1
        {"bai": "Bài 1", "q": "Thiết bị nào sau đây giúp em học vẽ tranh, gõ chữ và xem bài giảng điện tử?", "options": ["Máy tính", "Cái chổi quét nhà", "Chiếc cặp sách vải", "Hộp bút chì gỗ"], "ans": 0},
        # 2
        {"bai": "Bài 1", "q": "Khi bước vào phòng Tin học, em cần thực hiện điều gì đầu tiên?", "options": ["Đi nhẹ, nói khẽ và ngồi đúng vị trí được phân công", "Chạy thật nhanh giành chiếc máy tính ở đầu bàn", "Bật tất cả công tắc điện có trong phòng học", "Đặt chai nước lọc mở nắp lên trên bàn phím"], "ans": 0},
        # 3
        {"bai": "Bài 1", "q": "Đồ vật nào sau đây em KHÔNG được mang vào bàn máy tính trong phòng Tin học?", "options": ["Đồ ăn vặt và ly nước ngọt", "Quyển vở ghi bài môn Tin học", "Chiếc bút chì và cục tẩy", "Thẻ học sinh của trường"], "ans": 0},
        # 4
        {"bai": "Bài 1", "q": "Muốn tắt máy tính đúng cách khi hết giờ học, em cần làm gì?", "options": ["Thực hiện thao tác tắt máy theo đúng hướng dẫn của thầy cô", "Rút phích cắm điện trực tiếp ra khỏi ổ cắm trên tường", "Nhấn giữ thật mạnh vào giữa màn hình máy tính", "Gập bàn phím úp ngược lên trên chuột máy tính"], "ans": 0},
        # 5
        {"bai": "Bài 2", "q": "Một bộ máy tính để bàn thông thường gồm mấy bộ phận chính?", "options": ["4 bộ phận chính", "2 bộ phận chính", "6 bộ phận chính", "9 bộ phận chính"], "ans": 0},
        # 6
        {"bai": "Bài 2", "q": "Bộ phận nào của máy tính để bàn trông giống như chiếc tivi, dùng để hiển thị chữ và hình ảnh?", "options": ["Màn hình", "Thân máy", "Bàn phím", "Chuột máy tính"], "ans": 0},
        # 7
        {"bai": "Bài 2", "q": "Bộ phận nào được ví như “bộ não” điều khiển mọi hoạt động của máy tính để bàn?", "options": ["Thân máy", "Bàn phím", "Chuột máy tính", "Tấm lót chuột"], "ans": 0},
        # 8
        {"bai": "Bài 2", "q": "Bộ phận nào của máy tính có nhiều phím nhỏ mang chữ cái và chữ số?", "options": ["Bàn phím", "Màn hình", "Thân máy", "Loa máy tính"], "ans": 0},
        # 9
        {"bai": "Bài 2", "q": "Bộ phận nhỏ gọn nằm gọn trong lòng bàn tay, giúp điều khiển mũi tên trên màn hình là gì?", "options": ["Chuột máy tính", "Thân máy tính", "Màn hình máy tính", "Máy chiếu lớp học"], "ans": 0},
        # 10
        {"bai": "Bài 3", "q": "Khoảng cách hợp lý từ mắt em đến màn hình máy tính khi ngồi học là bao nhiêu?", "options": ["Từ 50 cm đến 80 cm", "Từ 5 cm đến 10 cm", "Từ 2 m đến 4 m", "Sát chạm mũi vào màn hình"], "ans": 0},
        # 11
        {"bai": "Bài 3", "q": "Ngồi sai tư thế (cúi gằm mặt, vẹo lưng) khi dùng máy tính lâu ngày sẽ gây tác hại gì?", "options": ["Dễ bị mỏi mắt, cận thị và cong vẹo cột sống", "Giúp em gõ bàn phím nhanh hơn gấp nhiều lần", "Giúp màn hình máy tính sáng rõ hơn ban đầu", "Giúp máy tính tiết kiệm điện năng hơn"], "ans": 0},
        # 12
        {"bai": "Bài 3", "q": "Sau khi học trên máy tính khoảng 30 đến 35 phút, em nên làm gì để bảo vệ đôi mắt?", "options": ["Nghỉ giải lao, nhìn ra xa và chớp mắt thư giãn", "Nhìn sát hơn vào màn hình để đọc chữ nhỏ", "Tắt hết đèn trong phòng học cho thật tối", "Lấy tay dụi thật mạnh vào hai mắt"], "ans": 0},
        # 13
        {"bai": "Bài 4", "q": "Trên mặt trên của chuột máy tính thông thường có hai nút bấm chính nào?", "options": ["Nút trái chuột và nút phải chuột", "Nút tăng âm và nút giảm âm", "Nút bật đèn và nút tắt đèn", "Nút chữ cái và nút chữ số"], "ans": 0},
        # 14
        {"bai": "Bài 4", "q": "Khi cầm chuột đúng cách bằng tay phải, ngón giữa của em đặt lên đâu?", "options": ["Đặt lên nút phải chuột", "Đặt lên nút trái chuột", "Đặt dưới mặt bàn học", "Đặt lên phím Cách"], "ans": 0},
        # 15
        {"bai": "Bài 4", "q": "Hình mũi tên nhỏ chuyển động trên màn hình khi em trượt chuột được gọi là gì?", "options": ["Con trỏ chuột", "Nút nguồn máy tính", "Thanh dấu cách", "Khung màn hình"], "ans": 0},
        # 16
        {"bai": "Bài 4", "q": "Thao tác “Nhấp chuột” (Click) được thực hiện bằng cách nhấn một lần vào nút nào của chuột?", "options": ["Nút trái chuột", "Nút phải chuột", "Mặt dưới của chuột", "Đầu dây cắm chuột"], "ans": 0},
        # 17
        {"bai": "Bài 4", "q": "Thao tác nhấn nhanh hai lần liên tiếp vào nút trái chuột rồi thả tay ra gọi là gì?", "options": ["Nhấp đúp chuột", "Di chuyển chuột", "Nhấp chuột phải", "Xoay tròn chuột"], "ans": 0},
        # 18
        {"bai": "Bài 5", "q": "Trong trò chơi ghép tranh trên máy tính, em dùng thao tác chuột nào để đưa mảnh ghép vào khung?", "options": ["Kéo – thả chuột", "Lật ngửa chuột lên", "Gõ chuột xuống mặt bàn", "Nhấp nút phải chuột liên tục"], "ans": 0},
        # 19
        {"bai": "Bài 5", "q": "Khi đang kéo – thả một hình ảnh trên màn hình, ngón trỏ của em phải giữ nút nào cho đến khi tới đích?", "options": ["Giữ nút trái chuột", "Giữ nút phải chuột", "Giữ phím Enter", "Giữ nút nguồn màn hình"], "ans": 0},
        # 20
        {"bai": "Bài 6", "q": "Để gõ số tuổi “6” của em vào máy tính, em nhấn vào loại phím nào trên bàn phím?", "options": ["Phím số 6", "Phím chữ S", "Phím mũi tên lên", "Phím thoát Esc"], "ans": 0},
        # 21
        {"bai": "Bài 6", "q": "Các phím mang chữ cái A, B, C, D... trên bàn phím được gọi chung là gì?", "options": ["Các phím chữ", "Các phím số", "Các phím âm lượng", "Các nút nguồn"], "ans": 0},
        # 22
        {"bai": "Bài 7", "q": "Khi đang gõ văn bản, muốn đưa con trỏ xuống một dòng mới, em nhấn phím nào?", "options": ["Phím Enter", "Phím Space", "Phím Shift", "Phím số 0"], "ans": 0},
        # 23
        {"bai": "Bài 7", "q": "Em vừa gõ nhầm một chữ cái và muốn xóa chữ ngay bên trái con trỏ soạn thảo, em nhấn phím nào?", "options": ["Phím Backspace", "Phím Space", "Phím Enter", "Phím chữ Z"], "ans": 0},
        # 24
        {"bai": "Bài 7", "q": "Để gõ được chữ cái in hoa bằng phím Shift, em thực hiện như thế nào?", "options": ["Nhấn giữ phím Shift rồi gõ phím chữ cái cần viết hoa", "Nhấn phím Shift xong thả tay ra rồi mới gõ chữ cái", "Nhấn phím Enter cùng lúc với phím Backspace", "Nhấn phím Space hai lần trước khi gõ chữ cái"], "ans": 0},
        # 25
        {"bai": "Bài 8", "q": "Giữa hai tiếng trong cụm từ “EM YEU TRUONG”, em cần nhấn phím nào một lần để tách chữ?", "options": ["Phím Cách (Space)", "Phím Backspace", "Phím Enter", "Phím Shift"], "ans": 0},
        # 26
        {"bai": "Bài 9", "q": "Sau khi vẽ xong một bức tranh đẹp trên máy tính, em cần chọn lệnh nào để không bị mất bài?", "options": ["Lệnh Lưu (Save)", "Lệnh Xóa bỏ (Delete)", "Lệnh Tắt màn hình", "Lệnh Cắt bỏ (Cut)"], "ans": 0},
        # 27
        {"bai": "Bài 9", "q": "Biểu tượng nút Lưu bài (Save) trong nhiều phần mềm học tập thường có hình dạng giống vật gì?", "options": ["Hình chiếc đĩa mềm vuông nhỏ", "Hình chiếc kéo cắt giấy", "Hình thùng rác nhỏ", "Hình chiếc loa phát thanh"], "ans": 0},
        # 28
        {"bai": "Bài 10", "q": "Để đóng cửa sổ phần mềm sau khi đã lưu bài xong, em thường nhấp chuột vào nút nào ở góc trên bên phải?", "options": ["Nút dấu nhân (X)", "Nút dấu cộng (+)", "Nút dấu hỏi (?)", "Nút ngôi sao (*)"], "ans": 0},
        # 29
        {"bai": "Bài 12", "q": "Thông tin xung quanh chúng ta thường gặp ở ba dạng phổ biến nào?", "options": ["Dạng chữ, dạng hình ảnh và dạng âm thanh", "Dạng gỗ, dạng nhựa và dạng kim loại", "Dạng nắng, dạng mưa và dạng sương mù", "Dạng bàn, dạng ghế và dạng bảng đen"], "ans": 0},
        # 30
        {"bai": "Bài 12", "q": "Bức tranh vẽ bông hoa hướng dương treo trên tường lớp học là thông tin dạng nào?", "options": ["Dạng hình ảnh", "Dạng âm thanh", "Dạng chữ số", "Dạng tiếng chuông"], "ans": 0},
        # 31
        {"bai": "Bài 12", "q": "Dòng chữ “LỚP 1A1 CHĂM NGOAN” gắn trước cửa lớp là thông tin dạng nào?", "options": ["Dạng chữ", "Dạng âm thanh", "Dạng tiếng còi", "Dạng giai điệu nhạc"], "ans": 0},
        # 32
        {"bai": "Bài 13", "q": "Máy tính hoạt động và thực hiện công việc dựa vào đâu?", "options": ["Dựa vào các lệnh do con người điều khiển", "Tự làm theo ý thích riêng của màn hình", "Tự đoán suy nghĩ của người ngồi cạnh", "Hoạt động ngẫu nhiên không cần ai bấm"], "ans": 0},
        # 33
        {"bai": "Bài 13", "q": "Trong phần mềm vẽ, khi em nhấp chuột chọn ô màu đỏ rồi tô vào quả táo, đó là lúc em đang làm gì?", "options": ["Ra lệnh cho máy tính tô màu đỏ vào quả táo", "Tắt nguồn màn hình của máy tính", "Xóa toàn bộ bức tranh khỏi máy tính", "Gõ tên trường học bằng bàn phím"], "ans": 0},
        # 34
        {"bai": "Bài 14", "q": "Khi gặp một công việc gồm nhiều việc nhỏ (như chuẩn bị đi học), em nên thực hiện thế nào?", "options": ["Làm lần lượt từng bước theo đúng thứ tự", "Chỉ làm bước cuối cùng và bỏ qua bước đầu", "Làm lộn xộn bước nào nhớ ra trước thì làm", "Bỏ dở giữa chừng không cần hoàn thành"], "ans": 0},
        # 35
        {"bai": "Bài 15", "q": "Khối lệnh có hình mũi tên hướng thẳng lên trên (⬆) điều khiển nhân vật đi về đâu?", "options": ["Đi lên phía trên 1 ô", "Đi xuống phía dưới 1 ô", "Đi sang bên trái 1 ô", "Đi sang bên phải 1 ô"], "ans": 0},
        # 36
        {"bai": "Bài 15", "q": "Khối lệnh có hình mũi tên hướng sang bên trái (⬅) điều khiển nhân vật đi về đâu?", "options": ["Đi sang bên trái 1 ô", "Đi sang bên phải 1 ô", "Đi lên phía trên 1 ô", "Đi xuống phía dưới 1 ô"], "ans": 0},
        # 37
        {"bai": "Bài 16", "q": "Trong trò chơi Mê cung của em, nhân vật KHÔNG được đi vào loại ô nào?", "options": ["Ô có vật cản như đá tảng hoặc bức tường", "Ô xuất phát ban đầu của nhân vật", "Ô đường đi màu trắng trống trải", "Ô đích đến có ngôi sao hoặc kho báu"], "ans": 0},
        # 38
        {"bai": "Bài 17", "q": "Trong lập trình cơ bản, “lỗi đường đi” nghĩa là gì?", "options": ["Chuỗi lệnh có bước sai làm nhân vật không tới được đích", "Màn hình máy tính bị bụi bẩn bám bên ngoài", "Ghế ngồi của học sinh chưa kê ngay ngắn", "Chuột máy tính có màu sắc khác với bàn phím"], "ans": 0},
        # 39
        {"bai": "Bài 18", "q": "Trong phần mềm lập trình trực quan, nút có biểu tượng Lá cờ màu xanh lá cây thường dùng để làm gì?", "options": ["Bắt đầu chạy chuỗi lệnh của chương trình", "Xóa sạch các nhân vật trên sân khấu", "Tắt hoàn toàn nguồn điện của máy tính", "Đổi ngôn ngữ bàn phím sang tiếng Anh"], "ans": 0},
        # 40
        {"bai": "Bài 19", "q": "Một tập hợp nhiều lệnh được ghép nối tiếp nhau từ đầu đến cuối để nhân vật hoàn thành nhiệm vụ gọi là gì?", "options": ["Chuỗi lệnh", "Bàn phím số", "Màn hình nền", "Thân máy tính"], "ans": 0},
    ],

    "review_th_mcq": [
        # 41
        {
            "bai": "Bài 2",
            "image_key": "g1_bo_phan_may_tinh",
            "image_caption": "Hình minh họa các bộ phận máy tính để bàn (dùng cho Câu 41)",
            "q": "Quan sát hình minh họa trên, cặp số nào chỉ đúng hai bộ phận dùng để đưa lệnh và gõ chữ vào máy tính (Bàn phím và Chuột)?",
            "options": [
                "Bộ phận (3) và bộ phận (4)",
                "Bộ phận (1) và bộ phận (2)",
                "Bộ phận (1) và bộ phận (3)",
                "Bộ phận (2) và bộ phận (4)",
            ],
            "ans": 0,
        },
        # 42
        {"bai": "Bài 3", "q": "Bạn Lan ngồi học máy tính trong phòng tối không bật đèn và dí sát mắt vào màn hình. Điều gì dễ xảy ra với Lan?", "options": ["Mắt bạn Lan nhanh bị mỏi, nhức và dễ bị cận thị", "Bạn Lan nhìn chữ rõ hơn và mắt khỏe mạnh hơn", "Máy tính sẽ chạy nhanh hơn khi ở trong bóng tối", "Bàn phím máy tính sẽ tự phát sáng nhiều màu"], "ans": 0},
        # 43
        {"bai": "Bài 4", "q": "Để mở một phần mềm trò chơi học tập có biểu tượng ngoài màn hình nền, thao tác chuột nào nhanh và đúng nhất?", "options": ["Di chuyển con trỏ vào biểu tượng rồi nhấp đúp nút trái chuột", "Di chuyển con trỏ ra góc trống màn hình rồi nhấp nút phải chuột", "Nhấc bổng con chuột lên khỏi mặt bàn và lắc nhẹ hai lần", "Chỉ nhìn vào biểu tượng mà không chạm tay vào chuột"], "ans": 0},
        # 44
        {"bai": "Bài 5", "q": "Bạn Nam muốn xếp các quả táo số 1, 2, 3 vào giỏ trên màn hình nhưng vừa nhấn chuột đã thả tay ngay nên quả táo rơi lại chỗ cũ. Nam cần sửa thế nào?", "options": ["Nhấn và giữ chặt nút trái chuột khi kéo quả táo đến đúng giỏ rồi mới thả tay", "Nhấp nút phải chuột thật nhanh năm lần liên tiếp vào cái giỏ", "Nhấn phím Backspace trên bàn phím để đẩy quả táo vào giỏ", "Tắt màn hình đi rồi bật lại để quả táo tự nhảy vào giỏ"], "ans": 0},
        # 45
        {"bai": "Bài 7", "q": "Bạn Hà định gõ từ “HOA” nhưng lại gõ nhầm thành “HOÁ” có dư chữ K ở cuối thành “HOAK”. Con trỏ đang ở ngay sau chữ K, Hà nhấn phím nào một lần để xóa chữ K?", "options": ["Phím Backspace", "Phím Enter", "Phím Space", "Phím Shift"], "ans": 0},
        # 46
        {"bai": "Bài 8", "q": "Khi gõ cụm từ “LOP 1A1” không dấu, bạn Bình quên nhấn phím Space nên màn hình hiện ra kết quả nào?", "options": ["LOP1A1", "LOP   1A1", "lop 1a1", "1A1 LOP"], "ans": 0},
        # 47
        {"bai": "Bài 9", "q": "Nếu em vẽ xong bức tranh ngôi nhà rất đẹp nhưng tắt phần mềm ngay mà KHÔNG nhấn Lưu (Save), chuyện gì sẽ xảy ra?", "options": ["Bức tranh vừa vẽ sẽ bị mất và không mở lại được nữa", "Máy tính vẫn tự động in bức tranh ra giấy trắng", "Bức tranh tự động gửi sang máy tính của bạn bên cạnh", "Màn hình nền tự động đổi thành bức tranh ngôi nhà"], "ans": 0},
        # 48
        {"bai": "Bài 10", "q": "Để gõ tên mình bên dưới bức tranh vừa vẽ trong phần mềm học tập, em cần kết hợp hai thiết bị nào?", "options": ["Dùng chuột chọn công cụ viết chữ rồi dùng bàn phím gõ tên", "Chỉ dùng loa máy tính và dây nguồn của màn hình", "Chỉ nhấn nút bật nguồn trên thân máy tính hai lần", "Dùng cục tẩy chì chà nhẹ lên mặt bàn phím máy tính"], "ans": 0},
        # 49
        {"bai": "Bài 12", "q": "Khi xem một đoạn phim hoạt hình có nhân vật ca hát và có dòng phụ đề tên bài hát ở dưới, em nhận được những dạng thông tin nào?", "options": ["Cả dạng chữ, dạng hình ảnh và dạng âm thanh", "Chỉ có duy nhất thông tin dạng chữ viết", "Chỉ có duy nhất thông tin dạng âm thanh", "Không có bất kỳ dạng thông tin nào"], "ans": 0},
        # 50
        {"bai": "Bài 12", "q": "Đèn tín hiệu giao thông ở ngã tư chuyển sang màu đỏ báo hiệu các xe dừng lại. Tín hiệu màu đỏ đó thuộc dạng thông tin nào?", "options": ["Dạng hình ảnh", "Dạng âm thanh", "Dạng chữ viết dài", "Dạng tiếng trống"], "ans": 0},
        # 51
        {"bai": "Bài 13", "q": "Vì sao khi em gõ phím chữ “M” mà trên màn hình lại hiện ra chữ “N”?", "options": ["Vì ngón tay em đã nhấn nhầm vào phím chữ N nên máy làm đúng theo lệnh phím N", "Vì máy tính hôm nay bị mệt nên tự ý đổi chữ cái của em", "Vì màn hình máy tính không biết hiển thị chữ cái M", "Vì chuột máy tính đặt quá gần bàn phím của em"], "ans": 0},
        # 52
        {"bai": "Bài 14", "q": "Khi mặc giày đi học, nếu em xỏ giày buộc dây xong xuôi rồi mới nhớ chưa đi tất (vớ), điều đó cho thấy điều gì?", "options": ["Thứ tự thực hiện các bước bị ngược nên phải tháo giày làm lại", "Làm bước nào trước cũng cho kết quả giống hệt nhau", "Đi tất bên ngoài chiếc giày sẽ giúp chạy nhanh hơn", "Không cần chia việc đi giày thành các bước nhỏ"], "ans": 0},
        # 53
        {"bai": "Bài 15", "q": "Chú ong đang ở một ô vuông. Sau khi thực hiện chuỗi lệnh: “Sang phải 1 ô -> Sang phải 1 ô -> Đi lên 1 ô”, chú ong đã đi tổng cộng mấy ô?", "options": ["3 ô", "1 ô", "2 ô", "5 ô"], "ans": 0},
        # 54
        {
            "bai": "Bài 16",
            "image_key": "g1_me_cung_robot",
            "image_caption": "Hình mê cung Robot tìm kho báu (dùng cho Câu 54)",
            "q": "Quan sát hình Mê cung Robot ở trên, nếu từ ô Xuất phát Robot thực hiện ngay lệnh “Sang phải 1 ô”, điều gì xảy ra ở bước tiếp theo?",
            "options": [
                "Robot bị kẹt vì cả ô bên phải và ô phía dưới đều là ĐÁ CẢN",
                "Robot đến ngay được ô KHO BÁU chỉ sau một bước",
                "Robot đi thẳng xuống thẳng hàng đến ô KHO BÁU",
                "Robot tự động bay vọt qua những ô ĐÁ CẢN màu đỏ",
            ],
            "ans": 0,
        },
        # 55
        {"bai": "Bài 17", "q": "Đích đến nằm ngay bên phải chú thỏ 2 ô trống, nhưng chuỗi lệnh đang là: “Sang phải 1 ô -> Sang trái 1 ô”. Em sửa lệnh thứ hai thành gì để thỏ tới đích?", "options": ["Đổi lệnh “Sang trái 1 ô” thành lệnh “Sang phải 1 ô”", "Đổi lệnh “Sang trái 1 ô” thành lệnh “Đi lên 1 ô”", "Đổi lệnh “Sang trái 1 ô” thành lệnh “Đi xuống 1 ô”", "Xóa cả hai lệnh để chú thỏ đứng yên tại chỗ"], "ans": 0},
        # 56
        {"bai": "Bài 19", "q": "Trước khi nhấn nút chạy toàn bộ chuỗi lệnh dài để đưa Robot vượt mê cung, em nên làm gì để hạn chế sai sót?", "options": ["Dùng ngón tay dò thử từng bước trên lưới ô vuông xem có vướng vật cản không", "Nhắm mắt lại và ghép thật nhiều khối lệnh bất kỳ vào chương trình", "Xóa hết khối lệnh bắt đầu để chương trình ngắn gọn hơn", "Chỉ ghép toàn bộ các khối lệnh đi sang trái liên tiếp"], "ans": 0},
    ],

    "review_th_tl": [
        # 57
        {
            "bai": "Bài 2",
            "q": "Em hãy kể tên 4 bộ phận chính của một bộ máy tính để bàn và cho biết bộ phận nào dùng để hiển thị hình ảnh cho em xem.",
            "answer": "4 bộ phận chính của máy tính để bàn gồm: Màn hình, Thân máy, Bàn phím và Chuột máy tính. Trong đó, Màn hình là bộ phận dùng để hiển thị hình ảnh và chữ viết cho em quan sát.",
        },
        # 58
        {
            "bai": "Bài 3",
            "q": "Em hãy nêu 2 quy tắc về tư thế ngồi đúng và giữ an toàn khi học tập với máy tính.",
            "answer": "(1) Ngồi thẳng lưng, hai vai thả lỏng, mắt cách màn hình khoảng 50 cm đến 80 cm. (2) Không dùng tay ướt chạm vào thiết bị điện, không tự ý rút dây nguồn và không mang đồ ăn, nước uống vào bàn máy tính.",
        },
        # 59
        {
            "bai": "Bài 4",
            "q": "Em hãy phân biệt sự khác nhau giữa thao tác “Nhấp chuột” và thao tác “Nhấp đúp chuột”.",
            "answer": "- Nhấp chuột (Click): Nhấn nhanh một lần nút trái chuột rồi thả tay ra (thường dùng để chọn một biểu tượng hoặc nút lệnh). - Nhấp đúp chuột (Double-click): Nhấn nhanh hai lần liên tiếp nút trái chuột rồi thả tay ra (thường dùng để mở một chương trình ngoài màn hình nền).",
        },
        # 60
        {
            "bai": "Bài 5",
            "q": "Em hãy mô tả các bước thực hiện thao tác “Kéo – thả chuột” để di chuyển một bức ảnh nhỏ trên màn hình sang vị trí mới.",
            "answer": "Gồm 3 bước: Bước 1: Di chuyển con trỏ chuột đặt lên bức ảnh cần chuyển. Bước 2: Nhấn và giữ nguyên nút trái chuột, từ từ kéo bức ảnh đến vị trí mới. Bước 3: Thả ngón tay khỏi nút trái chuột để đặt bức ảnh xuống.",
        },
        # 61
        {
            "bai": "Bài 7",
            "q": "Khi gõ chữ trên máy tính, phím Enter và phím Backspace giúp ích gì cho em?",
            "answer": "- Phím Enter giúp em đưa con trỏ soạn thảo xuống dòng mới (hoặc đồng ý thực hiện một lệnh). - Phím Backspace giúp em xóa bỏ chữ cái hoặc con số bị gõ sai nằm ngay bên trái con trỏ soạn thảo.",
        },
        # 62
        {
            "bai": "Bài 12",
            "q": "Em hãy lấy mỗi loại 1 ví dụ gần gũi ở trường học cho: thông tin dạng chữ, thông tin dạng hình ảnh và thông tin dạng âm thanh.",
            "answer": "- Thông tin dạng chữ: Dòng chữ bảng tên lớp “Lớp 1A1” (hoặc nội quy lớp học). - Thông tin dạng hình ảnh: Biển báo bỏ rác vào thùng có hình vẽ minh họa (hoặc tranh treo tường). - Thông tin dạng âm thanh: Tiếng trống trường báo giờ ra chơi (hoặc giọng cô giáo giảng bài).",
        },
        # 63
        {
            "bai": "Bài 13",
            "q": "Em hiểu thế nào về câu nói: “Máy tính làm việc theo lệnh của con người”? Hãy nêu 1 ví dụ minh họa.",
            "answer": "Câu nói đó có nghĩa là máy tính không tự ý làm việc mà chỉ thực hiện đúng những thao tác do con người ra lệnh. Ví dụ: Khi em nhấp chuột chọn nút Lưu (Save), máy tính nhận lệnh và lưu bức tranh của em lại.",
        },
        # 64
        {
            "bai": "Bài 18",
            "q": "Trong bài học Làm quen lập trình trực quan, việc ghép các khối lệnh giống với trò chơi quen thuộc nào của em và giúp nhân vật làm được gì?",
            "answer": "Việc ghép các khối lệnh giống như trò chơi xếp hình Lego lắp các mảnh ghép khớp vào nhau. Khi ghép đúng thứ tự thành chuỗi lệnh, nhân vật trên màn hình sẽ di chuyển hoặc thực hiện nhiệm vụ theo đúng ý muốn của em.",
        },
    ],

    "review_vd_tl": [
        # 65
        {
            "bai": "Bài 1",
            "q": "Giờ thực hành Tin học, bạn ngồi cạnh định lấy hộp sữa tươi đặt ngay lên bàn phím máy tính để vừa uống vừa gõ chữ. Em sẽ khuyên bạn như thế nào?",
            "answer": "Em sẽ nhẹ nhàng khuyên bạn cất hộp sữa vào cặp hoặc để khu vực quy định bên ngoài bàn máy, vì nếu sữa đổ vào bàn phím hay thân máy sẽ làm hỏng thiết bị và có thể gây chập điện rất nguy hiểm.",
        },
        # 66
        {
            "bai": "Bài 8",
            "q": "Em muốn gõ dòng chữ không dấu “EM HOC LOP 1” mà chữ cái đầu tiên viết hoa hoặc cả dòng viết hoa bằng phím Shift. Em hãy nêu cách gõ khoảng cách giữa các từ và cách gõ số 1.",
            "answer": "- Để viết hoa chữ cái, em nhấn giữ phím Shift đồng thời gõ phím chữ cái tương ứng. - Giữa các từ EM, HOC, LOP, 1 em nhấn phím Cách (Space) đúng 1 lần để tạo khoảng trắng. - Cuối câu em nhấn phím số 1 ở hàng phím số trên bàn phím.",
        },
        # 67
        {
            "bai": "Bài 9",
            "q": "Em hãy nêu trình tự 3 bước cơ bản để tạo và cất giữ sản phẩm tranh vẽ đầu tiên của em trên máy tính.",
            "answer": "Bước 1: Nhấp đúp chuột vào biểu tượng phần mềm vẽ để mở công cụ. Bước 2: Dùng chuột chọn nét bút, hình mẫu và màu sắc để vẽ bức tranh theo ý tưởng (có thể gõ thêm tên em). Bước 3: Nhấp chuột vào nút Lưu (Save) để lưu sản phẩm vào máy tính.",
        },
        # 68
        {
            "bai": "Bài 14",
            "q": "Em hãy chia công việc “Rửa tay sạch bằng xà phòng” thành 4 bước nhỏ theo đúng thứ tự từ đầu đến cuối.",
            "answer": "Bước 1: Làm ướt hai bàn tay dưới vòi nước sạch. Bước 2: Lấy xà phòng vào lòng bàn tay và xoa đều các mặt tay, kẽ ngón tay. Bước 3: Rửa sạch hết bọt xà phòng dưới vòi nước và khóa vòi nước lại. Bước 4: Lau khô hai bàn tay bằng khăn sạch.",
        },
        # 69
        {
            "bai": "Bài 14",
            "q": "Các bước chuẩn bị cặp sách buổi tối trước khi đi học bị viết lộn xộn: (A) Đóng nắp cặp sách gọn gàng; (B) Xem thời khóa biểu ngày mai; (C) Xếp đủ sách vở và hộp bút vào cặp. Em hãy sắp xếp lại theo thứ tự đúng.",
            "answer": "Thứ tự đúng để không bị quên sách vở là: Bước (B) Xem thời khóa biểu ngày mai -> Bước (C) Xếp đủ sách vở và hộp bút vào cặp -> Bước (A) Đóng nắp cặp sách gọn gàng.",
        },
        # 70
        {
            "bai": "Bài 15",
            "q": "Trên lưới ô vuông không có vật cản, chú Rùa đang ở ô xuất phát. Ô lá rau nằm cách chú Rùa 2 ô về phía bên phải và 1 ô về phía dưới. Em hãy viết chuỗi lệnh đưa chú Rùa đến ô lá rau.",
            "answer": "Chuỗi lệnh đúng gồm 3 bước: Bước 1: Sang phải 1 ô -> Bước 2: Sang phải 1 ô -> Bước 3: Đi xuống 1 ô (Hoặc học sinh có thể viết: Đi xuống 1 ô -> Sang phải 1 ô -> Sang phải 1 ô cũng hoàn toàn chính xác).",
        },
        # 71
        {
            "bai": "Bài 16",
            "q": "Chú Thỏ cần đi đến củ cà rốt ở cách 2 ô bên phải, nhưng ô ngay bên phải chú Thỏ có một tảng đá chặn đường (các ô hàng dưới đều trống). Em hãy viết chuỗi lệnh giúp chú Thỏ đi vòng qua tảng đá để lấy cà rốt.",
            "answer": "Chuỗi lệnh đi vòng qua tảng đá gồm 4 bước: Bước 1: Đi xuống 1 ô -> Bước 2: Sang phải 1 ô -> Bước 3: Sang phải 1 ô -> Bước 4: Đi lên 1 ô.",
        },
        # 72
        {
            "bai": "Bài 10",
            "q": "Khi gõ tên lớp, bạn Nam gõ nhầm thành “1A12” (bị thừa số 2 ở cuối). Sau đó bạn muốn xuống dòng để gõ tên mình là “NAM”. Em hãy chỉ giúp bạn các phím cần bấm theo thứ tự.",
            "answer": "Thứ tự các phím bạn Nam cần bấm là: (1) Bấm phím Backspace 1 lần để xóa số 2 bị thừa; (2) Bấm phím Enter 1 lần để xuống dòng mới; (3) Giữ phím Shift và lần lượt gõ các phím chữ N, A, M.",
        },
    ],

    "review_vdc_tl": [
        # 73
        {
            "bai": "Bài 17",
            "q": "Trên hàng ngang gồm 4 ô liên tiếp (Ô 1 -> Ô 2 -> Ô 3 -> Ô 4), Robot ở Ô 1 và Ngôi sao ở Ô 4. Bạn Khoa ghép chuỗi lệnh: “Sang phải 1 ô -> Sang phải 1 ô -> Đi xuống 1 ô”. Em hãy chỉ ra bước lệnh bị sai và sửa lại chuỗi lệnh cho đúng.",
            "answer": "- Bước lệnh bị sai là lệnh thứ 3 (“Đi xuống 1 ô”) làm Robot chệch khỏi hàng ngang ở Ô 3 thay vì tiến vào Ô 4. - Chuỗi lệnh sửa lại đúng là: Sang phải 1 ô -> Sang phải 1 ô -> Sang phải 1 ô.",
        },
        # 74
        {
            "bai": "Bài 16",
            "q": "Từ ô Xuất phát đến ô Kho báu có hai con đường hợp lệ không vướng vật cản: Đường thứ nhất dài 4 bước lệnh, Đường thứ hai dài 8 bước lệnh đi vòng quanh. Nếu là người lập trình cho Robot, em chọn đường nào? Vì sao?",
            "answer": "Em chọn Đường thứ nhất (dài 4 bước lệnh). Vì cả hai đường đều an toàn tới đích nhưng đường 4 bước ngắn gọn hơn, giúp Robot đi nhanh hơn, tiết kiệm năng lượng và ít bị nhầm lẫn khi ghép khối lệnh.",
        },
        # 75
        {
            "bai": "Bài 19",
            "q": "Robot đi từ ô A theo chuỗi 4 lệnh: “Đi lên 1 ô -> Sang phải 1 ô -> Đi xuống 1 ô -> Sang trái 1 ô”. Hỏi sau khi thực hiện hết 4 lệnh đó, Robot đang đứng ở đâu? Giải thích ngắn gọn.",
            "answer": "Sau khi thực hiện hết 4 lệnh, Robot quay trở về đứng đúng tại ô A ban đầu. Giải thích: Robot đi lên 1 ô rồi lại xuống 1 ô, sang phải 1 ô rồi lại sang trái 1 ô (đi thành 1 vòng hình vuông nhỏ khép kín).",
        },
        # 76
        {
            "bai": "Bài 15",
            "q": "Robot đã đi từ Nhà đến Trường theo chuỗi 3 lệnh: “Sang phải 1 ô -> Sang phải 1 ô -> Đi lên 1 ô”. Tan học, muốn Robot đi từ Trường quay ngược về Nhà theo đúng con đường cũ, em cần viết chuỗi 3 lệnh nào?",
            "answer": "Để đi ngược lại đúng con đường cũ từ Trường về Nhà, em đảo ngược thứ tự và hướng đi thành chuỗi 3 lệnh: Đi xuống 1 ô -> Sang trái 1 ô -> Sang trái 1 ô.",
        },
        # 77
        {
            "bai": "Bài 12",
            "q": "Cô giáo muốn làm một tấm bảng nhắc nhở các bạn nhỏ lớp 1 (có cả những bạn mới học đánh vần chậm) biết “Bỏ rác đúng nơi quy định” ngay tại sân trường. Theo em, nên dùng dạng thông tin nào trên tấm bảng để bạn nào nhìn vào cũng hiểu ngay?",
            "answer": "Nên kết hợp cả thông tin dạng hình ảnh (hình vẽ bạn nhỏ bỏ rác vào thùng rõ ràng, màu sắc tươi sáng) kèm dòng chữ ngắn gọn. Vì hình ảnh trực quan giúp các bạn đọc chữ còn chậm vẫn hiểu ngay việc cần làm chỉ qua một ánh nhìn.",
        },
        # 78
        {
            "bai": "Bài 18",
            "q": "Trong phần mềm lập trình trực quan, bạn Mai kéo đủ 4 khối lệnh di chuyển ra màn hình nhưng để 4 khối lệnh nằm rời rạc xa nhau, không ghép dính dưới khối lệnh “Khi bấm Lá cờ xanh”. Khi bấm Lá cờ xanh chuyện gì xảy ra và Mai cần sửa thế nào?",
            "answer": "- Khi bấm Lá cờ xanh, nhân vật sẽ đứng yên không thực hiện 4 lệnh di chuyển vì các khối lệnh bị rời khỏi khối bắt đầu. - Cách sửa: Mai dùng thao tác kéo – thả chuột ghép lần lượt 4 khối lệnh dính khít nối tiếp ngay bên dưới khối “Khi bấm Lá cờ xanh”.",
        },
        # 79
        {
            "bai": "Bài 17",
            "q": "Trên lưới 3x3 (hàng 1 trên cùng, hàng 3 dưới cùng; cột 1 bên trái, cột 3 bên phải), Robot ở góc trên trái (hàng 1, cột 1), Đích ở góc dưới phải (hàng 3, cột 3). Ô chính giữa lưới (hàng 2, cột 2) có ĐÁ CẢN. Em hãy viết 2 chuỗi lệnh KHÁC NHAU (mỗi chuỗi 4 bước) đều đưa Robot tới Đích an toàn.",
            "answer": "- Cách 1 (đi men theo cạnh trên rồi xuống cạnh phải): Sang phải 1 ô -> Sang phải 1 ô -> Đi xuống 1 ô -> Đi xuống 1 ô. - Cách 2 (đi men theo cạnh trái rồi sang cạnh dưới): Đi xuống 1 ô -> Đi xuống 1 ô -> Sang phải 1 ô -> Sang phải 1 ô.",
        },
        # 80
        {
            "bai": "Bài 19",
            "q": "Em hãy tự thiết kế một trò chơi nhỏ “Robot tưới cây” dài 3 bước lệnh trên sân trường: nêu rõ vị trí xuất phát của Robot, vị trí bồn hoa, 1 vật cản cần tránh và viết chuỗi 3 lệnh hoàn chỉnh.",
            "answer": "Ví dụ thiết kế hợp lệ: Robot xuất phát ở Cổng lớp (ô 1,1); ghế đá cản đường ở ngay bên phải (ô 1,2); bồn hoa cần tưới nằm ở ô (2,3). Chuỗi 3 lệnh điều khiển Robot: Bước 1: Đi xuống 1 ô (tới ô 2,1) -> Bước 2: Sang phải 1 ô (tới ô 2,2) -> Bước 3: Sang phải 1 ô (tới bồn hoa ở ô 2,3 để tưới cây).",
        },
    ],
}
