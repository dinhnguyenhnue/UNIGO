# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Cuối học kì 1 (Đánh giá định kỳ 2) & Đề cương 80 câu ôn tập
Môn: Tin học 8 (Bộ sách Kết nối tri thức với cuộc sống - UNIGO)
Phạm vi: Bài 1 đến Bài 8a, trang 5 - 40 SGK Tin học 8.
"""

GRADE_DATA = {
    "lop": 8,
    "cap_hoc": "THCS",
    "lop_hien_thi": "8A1",
    "thoi_gian": "45 phút",
    "tuan_on_tap": 18,
    "tuan_kiem_tra": 20,
    "tiet_on_tap": 18,
    "tiet_kiem_tra": 19,

    "tom_tat_ly_thuyet": [
        {
            "bai": "Bài 1. Lược sử công cụ tính toán",
            "y_chinh": [
                "Máy tính cơ học đầu tiên: Năm 1642, nhà bác học Pháp Blaise Pascal chế tạo máy tính cơ học Pascaline dựa trên hệ thống bánh răng thực hiện phép cộng, trừ. Năm 1833, Charles Babbage thiết kế máy tính đa năng tính tự động (cùng nữ lập trình viên đầu tiên Ada Lovelace), đặt nền móng kiến trúc máy tính hiện đại.",
                "Năm thế hệ máy tính điện tử: Thế hệ 1 (1945–1955) dùng đèn điện tử chân không (tiêu biểu ENIAC); Thế hệ 2 (1955–1965) dùng bóng bán dẫn (Transistor); Thế hệ 3 (1965–1974) dùng mạch tích hợp (IC); Thế hệ 4 (1974–1990) dùng vi xử lí mật độ rất cao (VLSI), xuất hiện máy tính cá nhân; Thế hệ 5 (1990–nay) dùng siêu vi mạch (ULSI) và phát triển Trí tuệ nhân tạo (AI).",
                "Tác động xã hội: Máy tính thu nhỏ kích thước, tăng tốc độ vượt bậc, làm thay đổi sâu sắc y tế, giáo dục, sản xuất, giao tiếp và đời sống con người."
            ],
            "image_key": "g8_the_he_may_tinh_tin_cay",
        },
        {
            "bai": "Bài 2. Thông tin trong môi trường số",
            "y_chinh": [
                "Đặc điểm của thông tin số: Được mã hóa thành dãy bit, dễ dàng sao chép, lan truyền nhanh trên diện rộng, lưu trữ với dung lượng khổng lồ, đa dạng kiểu loại (chữ, ảnh, âm thanh, video), có công cụ tìm kiếm/xử lí nhanh nhưng khó thu hồi triệt để khi đã phát tán.",
                "Tiêu chí đánh giá độ tin cậy của thông tin số: (1) Nguồn gốc và thẩm quyền tác giả/tổ chức (ưu tiên tên miền .gov.vn của cơ quan nhà nước, .edu.vn của tổ chức giáo dục); (2) Mục đích bài viết (cung cấp tri thức khách quan đáng tin hơn quảng cáo bán hàng, giật gân); (3) Tính cập nhật (thời gian công bố mới nhất); (4) Trích dẫn minh chứng rõ ràng và kiểm chứng chéo nhiều nguồn chính thống."
            ],
        },
        {
            "bai": "Bài 3. Thực hành khai thác thông tin số",
            "y_chinh": [
                "Kĩ thuật tìm kiếm hiệu quả: Xác định từ khóa sát chủ đề, đặt cụm từ trong dấu ngoặc kép (\"...\") để tìm chính xác, kết hợp lọc theo thời gian hoặc tên miền uy tín.",
                "Chọn lọc và tổng hợp: Đối chiếu số liệu giữa các nguồn tin cậy, tóm tắt thành dàn ý logic phục vụ bài trình chiếu hoặc báo cáo học tập.",
                "Tuân thủ bản quyền khi khai thác: Luôn ghi rõ tên tác giả, tên tài liệu, địa chỉ nguồn (URL) và thời gian truy cập đối với thông tin, hình ảnh trích dẫn."
            ],
        },
        {
            "bai": "Bài 4. Đạo đức và văn hóa trong sử dụng công nghệ kĩ thuật số",
            "y_chinh": [
                "Biểu hiện đạo đức và pháp luật số: Tôn trọng quyền riêng tư và dữ liệu cá nhân; tôn trọng bản quyền phần mềm và tác phẩm số (không đạo văn, không bẻ khóa phần mềm lậu); trung thực khi học tập, kiểm tra trực tuyến.",
                "Ứng xử văn hóa khi dùng thiết bị số: Không quay phim, chụp ảnh ở khu vực có biển cấm hoặc khi chưa được sự đồng ý của người trong ảnh; không phát tán tin giả, thông tin xúc phạm danh dự người khác; đặt điện thoại ở chế độ im lặng nơi trang nghiêm, rạp chiếu phim, lớp học."
            ],
        },
        {
            "bai": "Bài 5. Sử dụng bảng tính giải quyết bài toán thực tế",
            "y_chinh": [
                "Địa chỉ tương đối (ví dụ: B3, D3): Tự động thay đổi cả tên cột và tên hàng tương ứng khi sao chép công thức sang vị trí mới.",
                "Địa chỉ tuyệt đối (ví dụ: $F$2): Có dấu đô la ($) đặt trước cả tên cột và tên hàng, giữ nguyên hoàn toàn không thay đổi khi sao chép công thức đến bất kì ô nào (nhấn phím F4 để chuyển đổi nhanh).",
                "Địa chỉ hỗn hợp (ví dụ: $B3 hoặc B$3): Chỉ có dấu $ trước tên cột (cố định cột, hàng thay đổi) hoặc trước tên hàng (cố định hàng, cột thay đổi) khi sao chép công thức."
            ],
        },
        {
            "bai": "Bài 6. Sắp xếp và lọc dữ liệu",
            "y_chinh": [
                "Sắp xếp dữ liệu (Thẻ Data -> Sort): Hoán đổi vị trí các hàng theo giá trị một hoặc nhiều cột tiêu chí (dùng nút Add Level để thêm tiêu chí phụ) theo thứ tự tăng dần (A to Z / Smallest to Largest) hoặc giảm dần (Z to A / Largest to Smallest).",
                "Lọc dữ liệu (Thẻ Data -> Filter): Nhấn biểu tượng phễu lọc để hiển thị nút mũi tên tại tiêu đề cột, chọn điều kiện lọc (theo giá trị hoặc Number Filters / Text Filters).",
                "Lưu ý quan trọng: Lọc dữ liệu chỉ hiển thị các hàng thỏa mãn điều kiện và tạm thời ẩn đi các hàng không thỏa mãn, tuyệt đối KHÔNG xóa mất dữ liệu gốc trên trang tính."
            ],
        },
        {
            "bai": "Bài 7. Trình bày dữ liệu bằng biểu đồ",
            "y_chinh": [
                "Ý nghĩa từng loại biểu đồ (Thẻ Insert -> nhóm Charts): Biểu đồ cột (Column Chart) dùng để so sánh độ lớn dữ liệu giữa các mục/đối tượng; Biểu đồ đoạn thẳng (Line Chart) dùng để thể hiện xu hướng tăng giảm, biến động theo thời gian; Biểu đồ hình quạt tròn (Pie Chart) dùng để biểu diễn tỉ lệ phần trăm (%) của từng thành phần so với tổng thể (100%).",
                "Các thành phần chính của biểu đồ: Tiêu đề biểu đồ (Chart Title), Chú giải (Legend), Nhãn số liệu (Data Labels), Trục ngang và Trục đứng (Axes) — bổ sung qua nút Chart Elements (+)."
            ],
            "image_key": "g8_excel_tuyet_doi_bieu_do",
        },
        {
            "bai": "Bài 8a. Làm việc với danh sách dạng liệt kê và hình ảnh trong văn bản",
            "y_chinh": [
                "Danh sách dạng liệt kê (Thẻ Home -> nhóm Paragraph): Danh sách dấu đầu dòng (Bullets) dùng cho các ý ngang hàng không phân biệt thứ tự; Danh sách có thứ tự (Numbering) dùng cho quy trình các bước tuần tự hoặc xếp hạng.",
                "Xử lí hình ảnh & hình đồ họa trong Word: Chèn ảnh (Insert -> Pictures) và hình khối (Insert -> Shapes, nháy phải chọn Add Text để viết chữ vào hình); cắt bỏ phần thừa của ảnh bằng công cụ Crop; chọn chế độ bố trí văn bản quanh ảnh bằng Wrap Text; thay đổi thứ tự lớp chồng nhau bằng Bring Forward (đưa lên trên) và Send Backward (đưa xuống dưới)."
            ],
        },
    ],

    "ma_tran_kiem_tra": [
        [
            "Chủ đề 1. Máy tính và cộng đồng (Bài 1: Lược sử công cụ tính toán)",
            "3 câu (0,75đ)",
            "1 câu (0,25đ)",
            "—",
            "—",
            "—",
            "1,0 điểm",
        ],
        [
            "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin (Bài 2, Bài 3)",
            "4 câu (1,0đ)",
            "1 câu (0,25đ)",
            "1 câu (1,0đ)",
            "—",
            "—",
            "2,25 điểm",
        ],
        [
            "Chủ đề 3. Đạo đức, pháp luật và văn hóa trong môi trường số (Bài 4)",
            "2 câu (0,5đ)",
            "1 câu (0,25đ)",
            "—",
            "—",
            "—",
            "0,75 điểm",
        ],
        [
            "Chủ đề 4. Ứng dụng tin học – Bảng tính & Soạn thảo nâng cao (Bài 5, 6, 7, 8a)",
            "11 câu (2,75đ)",
            "5 câu (1,25đ)",
            "—",
            "1 câu (1,0đ)",
            "1 câu (1,0đ)",
            "6,0 điểm",
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
            "q": "Chiếc máy tính cơ học Pascaline ra đời năm 1642 là phát minh của nhà bác học nào?",
            "options": [
                "Blaise Pascal",
                "Charles Babbage",
                "John von Neumann",
                "Alan Turing",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "level": "NB",
            "q": "Linh kiện điện tử đặc trưng của máy tính điện tử thế hệ thứ nhất (tiêu biểu là máy ENIAC) là gì?",
            "options": [
                "Đèn điện tử chân không",
                "Bóng bán dẫn (Transistor)",
                "Mạch tích hợp (IC)",
                "Siêu vi mạch (ULSI)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "level": "NB",
            "q": "Sự xuất hiện của máy tính cá nhân (PC) gắn liền với thành tựu công nghệ của máy tính điện tử thế hệ thứ mấy?",
            "options": [
                "Thế hệ thứ tư (dùng vi xử lí mật độ rất cao VLSI)",
                "Thế hệ thứ nhất (dùng đèn điện tử chân không)",
                "Thế hệ thứ hai (dùng bóng bán dẫn)",
                "Thời kì máy tính cơ học bánh răng thế kỉ XVII",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Phát biểu nào sau đây nêu ĐÚNG đặc điểm của thông tin trong môi trường số?",
            "options": [
                "Dễ dàng sao chép, truyền tải nhanh chóng và rất khó thu hồi hoàn toàn khi đã phát tán",
                "Chỉ tồn tại duy nhất ở dạng chữ viết đen trắng và không thể chứa âm thanh hay video",
                "Mỗi lần sao chép sang máy tính khác sẽ làm chất lượng thông tin gốc giảm đi một nửa",
                "Chỉ cho phép một người đọc tại một thời điểm trên toàn mạng Internet",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Tại Việt Nam, các trang web chính thống của cơ quan nhà nước và Chính phủ thường có đuôi tên miền nào?",
            "options": [
                ".gov.vn",
                ".com.vn",
                ".net",
                ".shop",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "level": "NB",
            "q": "Khi tìm kiếm thông tin trên máy tìm kiếm, để kết quả trả về chứa chính xác nguyên văn một cụm từ khóa, ta đặt cụm từ đó trong cặp dấu nào?",
            "options": [
                "Cặp dấu ngoặc kép \"...\"",
                "Cặp dấu ngoặc nhọn {...}",
                "Cặp dấu ngoặc vuông [...]",
                "Cặp dấu thăng #...#",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "level": "NB",
            "q": "Khi sử dụng số liệu và hình ảnh khai thác từ Internet đưa vào bài báo cáo thuyết trình của nhóm, học sinh cần thực hiện điều gì?",
            "options": [
                "Ghi rõ nguồn trích dẫn, tên tác giả và địa chỉ trang web cung cấp tư liệu",
                "Xóa tên tác giả trên ảnh và ghi tên mình là người chụp bức ảnh đó",
                "Chỉ cần đổi màu viền ảnh là mặc nhiên trở thành chủ sở hữu bản quyền",
                "Giấu kín nguồn tài liệu để thầy cô không đối chiếu được số liệu",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "NB",
            "q": "Hành động nào sau đây thể hiện sự tuân thủ đạo đức và văn hóa trong sử dụng công nghệ kĩ thuật số?",
            "options": [
                "Tắt chuông điện thoại và không quay lén phim trong rạp chiếu phim",
                "Tự ý quay video người khác tại khu vực có biển cấm quay phim, chụp ảnh",
                "Tải phần mềm bẻ khóa lậu để dùng miễn phí các tính năng trả phí",
                "Sao chép nguyên văn bài tiểu luận trên mạng rồi nộp làm bài kiểm tra của mình",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "NB",
            "q": "Khi mượn máy tính xách tay của bạn để gửi tệp bài tập nhưng thấy tài khoản mạng xã hội của bạn vẫn đang đăng nhập sẵn, cách hành xử trung thực nhất là gì?",
            "options": [
                "Không tự ý đọc tin nhắn riêng của bạn và nhắc bạn đăng xuất để bảo mật tài khoản",
                "Mở hộp thư riêng tư của bạn ra đọc rồi chụp màn hình gửi cho người khác",
                "Dùng tài khoản của bạn đăng bài viết trêu đùa lên tường trang cá nhân",
                "Thay đổi mật khẩu tài khoản của bạn mà không thông báo",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "NB",
            "q": "Trong phần mềm bảng tính Excel, cách viết nào dưới đây biểu diễn một địa chỉ tuyệt đối?",
            "options": [
                "$F$2",
                "F2",
                "$F2",
                "F$2",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "NB",
            "q": "Khi đang nhập địa chỉ ô trong công thức Excel, nhấn phím chức năng nào giúp chuyển đổi nhanh giữa địa chỉ tương đối, tuyệt đối và hỗn hợp?",
            "options": [
                "Phím F4",
                "Phím F1",
                "Phím F2",
                "Phím F5",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "NB",
            "q": "Địa chỉ ô có dạng $B3 hoặc B$3 (chỉ có một kí hiệu $ trước tên cột hoặc trước tên hàng) được gọi là loại địa chỉ gì?",
            "options": [
                "Địa chỉ hỗn hợp",
                "Địa chỉ tuyệt đối hoàn toàn",
                "Địa chỉ tương đối thuần túy",
                "Địa chỉ siêu liên kết",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Các nút lệnh Sắp xếp (Sort) và Lọc dữ liệu (Filter) trong phần mềm bảng tính Excel nằm trong thẻ dải lệnh nào?",
            "options": [
                "Thẻ Data",
                "Thẻ Insert",
                "Thẻ Page Layout",
                "Thẻ View",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Trong hộp thoại Sort của Excel, nút lệnh nào dùng để thêm tiêu chí sắp xếp tiếp theo khi cột tiêu chí thứ nhất có các giá trị bằng nhau?",
            "options": [
                "Add Level",
                "Delete Level",
                "Wrap Text",
                "Merge & Center",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "NB",
            "q": "Khi thực hiện lọc dữ liệu (Filter) trên trang tính Excel, những hàng không thỏa mãn điều kiện lọc sẽ như thế nào?",
            "options": [
                "Tạm thời bị ẩn đi chứ không bị xóa mất khỏi bảng tính",
                "Bị xóa vĩnh viễn khỏi tệp bảng tính và không khôi phục được",
                "Tự động chuyển sang một tệp văn bản Word mới",
                "Đổi toàn bộ số liệu trong hàng đó thành giá trị 0",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "NB",
            "q": "Loại biểu đồ nào thích hợp nhất để biểu diễn tỉ lệ phần trăm (%) đóng góp của từng phần so với tổng thể?",
            "options": [
                "Biểu đồ hình quạt tròn (Pie Chart)",
                "Biểu đồ đoạn thẳng (Line Chart)",
                "Biểu đồ cột ghép (Clustered Column)",
                "Biểu đồ phân tán (Scatter)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "NB",
            "q": "Để tạo biểu đồ trực quan từ bảng dữ liệu đã chọn trong Excel, người dùng mở thẻ nào trên thanh công cụ Ribbon?",
            "options": [
                "Thẻ Insert (nhóm Charts)",
                "Thẻ Review (nhóm Proofing)",
                "Thẻ Formulas (nhóm Function Library)",
                "Thẻ File (nhóm Print)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "NB",
            "q": "Thành phần giải thích ý nghĩa màu sắc hoặc kí hiệu của các chuỗi dữ liệu trên biểu đồ được gọi là gì?",
            "options": [
                "Chú giải biểu đồ (Legend)",
                "Thanh công thức (Formula Bar)",
                "Hộp tên địa chỉ (Name Box)",
                "Đường lưới trang in (Gridlines)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "level": "NB",
            "q": "Trong phần mềm soạn thảo Word, công cụ nào giúp cắt bỏ phần rìa không cần thiết xung quanh một bức ảnh?",
            "options": [
                "Crop (trong thẻ Picture Format)",
                "Bullets (trong thẻ Home)",
                "Sort (trong thẻ Layout)",
                "Find (trong thẻ Home)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "level": "NB",
            "q": "Khi muốn trình bày các bước thực hiện một quy trình thí nghiệm theo đúng trình tự Bước 1, Bước 2, Bước 3... trong Word, ta nên chọn kiểu danh sách nào?",
            "options": [
                "Danh sách liệt kê có thứ tự (Numbering)",
                "Danh sách dấu hoa thị không thứ tự (Bullets)",
                "Chèn chữ nghệ thuật (WordArt)",
                "Chèn tiêu đề chân trang (Footer)",
            ],
            "ans": 0,
        },

        # --- 8 CÂU THÔNG HIỂU (TH) ---
        {
            "bai": "Bài 1",
            "level": "TH",
            "q": "Vì sao dự án Máy tính giải tích (1833) của Charles Babbage được coi là bước ngoặt đặt nền móng cho máy tính hiện đại dù thời đó chưa hoàn thiện bằng linh kiện điện tử?",
            "options": [
                "Vì máy được thiết kế tính toán tự động đa năng theo chương trình và có cấu trúc bộ nhớ, bộ xử lí, vào – ra tương tự máy tính ngày nay",
                "Vì máy sử dụng siêu vi mạch ULSI và kết nối mạng không dây Wi-Fi tốc độ cao",
                "Vì đó là chiếc bàn tính gảy bằng hạt gỗ đầu tiên được dùng trong buôn bán cổ đại",
                "Vì máy hoạt động hoàn toàn bằng năng lượng mặt trời để xử lí trí tuệ nhân tạo",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "TH",
            "q": "Bạn Minh cần tìm thông tin hướng dẫn phòng dịch sốt xuất huyết để làm tờ rơi tuyên truyền cho trường. Nguồn thông tin nào dưới đây có độ tin cậy cao nhất?",
            "options": [
                "Bài hướng dẫn mới nhất đăng trên Cổng thông tin Bộ Y tế (moh.gov.vn)",
                "Bài quảng cáo bán thuốc gia truyền không rõ nhãn mác trên mạng xã hội",
                "Bình luận truyền miệng trong một nhóm trò chơi điện tử ẩn danh",
                "Một bài blog cá nhân viết từ 12 năm trước không ghi tên tác giả và nguồn trích dẫn",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "level": "TH",
            "q": "Hành vi nào dưới đây vi phạm quyền sở hữu trí tuệ và bản quyền trong môi trường số?",
            "options": [
                "Tải một bức ảnh thiết kế có bản quyền trên mạng, xóa logo tác giả rồi đem dự thi dưới tên của mình",
                "Trích dẫn một đoạn số liệu thống kê của Tổng cục Thống kê và ghi đầy đủ nguồn tham khảo ở cuối bài",
                "Sử dụng các biểu tượng đồ họa từ kho tài nguyên giấy phép mở miễn phí đúng theo điều khoản",
                "Tự vẽ sơ đồ tư duy bằng phần mềm trên máy tính cá nhân để thuyết trình trước lớp",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "TH",
            "image_key": "g8_excel_tuyet_doi_bieu_do",
            "image_caption": "Hình minh họa: Bảng tính Gian hàng hội chợ 8A1 (Quan sát bảng tính để trả lời Câu 24, Câu 25, Câu 26 và Câu 2 phần Tự luận)",
            "q": "Quan sát hình 'GIAN HÀNG HỘI CHỢ 8A1' ở trên (ô F2 lưu tỉ lệ quỹ cố định 10%, cột D là Thành tiền). Vì sao tại ô F3 phải viết công thức =D3*$F$2 thay vì viết =D3*F2 trước khi sao chép xuống F4 và F5?",
            "options": [
                "Để cố định địa chỉ ô $F$2 luôn lấy tỉ lệ 10%, tránh bị dịch chuyển thành F3, F4 khi sao chép xuống dưới",
                "Để cố định địa chỉ ô D3 không thay đổi sang D4, D5 khi sao chép công thức xuống các hàng dưới",
                "Vì phần mềm Excel bắt buộc mọi phép nhân đều phải có ít nhất hai kí hiệu $ mới thực hiện được",
                "Để tự động làm tròn kết quả tiền trích quỹ lên hàng triệu đồng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "level": "TH",
            "q": "Dựa vào số liệu trong bảng 'GIAN HÀNG HỘI CHỢ 8A1' (Sách cũ: B3=40, C3=15000; Đồ thủ công: B4=25, C4=20000; Cây sen đá: B5=30, C5=25000; Tỉ lệ quỹ F2=10%), giá trị Thành tiền tại ô D3 và tiền Quỹ tại ô F4 lần lượt bằng bao nhiêu?",
            "options": [
                "D3 = 600000 và F4 = 50000",
                "D3 = 500000 và F4 = 60000",
                "D3 = 750000 và F4 = 75000",
                "D3 = 60000 và F4 = 500000",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "level": "TH",
            "q": "Trên bảng 'GIAN HÀNG HỘI CHỢ 8A1' (Số lượng: Sách cũ = 40, Đồ thủ công = 25, Cây sen đá = 30), nếu bật Filter và lọc cột Số lượng với điều kiện 'Greater Than Or Equal To 30' (>= 30) thì kết quả hiển thị gồm những mặt hàng nào?",
            "options": [
                "Hiển thị 2 mặt hàng: Sách cũ và Cây sen đá (tạm ẩn hàng Đồ thủ công)",
                "Chỉ hiển thị duy nhất mặt hàng Sách cũ và xóa vĩnh viễn 2 mặt hàng còn lại",
                "Hiển thị 2 mặt hàng: Đồ thủ công và Cây sen đá",
                "Hiển thị cả 3 mặt hàng theo thứ tự Đơn giá giảm dần",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "level": "TH",
            "q": "Nhà trường muốn theo dõi xu hướng biến động số lượng học sinh đạt giải Tin học trẻ qua 5 năm học liên tiếp (từ 2022 đến 2026). Loại biểu đồ nào giúp nhận diện rõ nhất xu hướng tăng giảm qua các năm?",
            "options": [
                "Biểu đồ đoạn thẳng (Line Chart)",
                "Biểu đồ hình quạt tròn (Pie Chart)",
                "Danh sách dấu đầu dòng (Bullets)",
                "Hộp thoại lọc dữ liệu (AutoFilter)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "level": "TH",
            "q": "Khi thiết kế trang báo tường trong Word, em vẽ một khung hình chữ nhật màu nhạt đè lên làm che khuất bức ảnh kỷ niệm của lớp ở bên dưới. Để đưa khung hình chữ nhật ra phía sau bức ảnh mà không phải xóa đi vẽ lại, em dùng lệnh nào?",
            "options": [
                "Chọn khung hình chữ nhật rồi nhấp lệnh Send Backward / Send to Back",
                "Chọn khung hình chữ nhật rồi nhấp lệnh Crop cắt đôi bức ảnh",
                "Chọn đoạn văn bản tiêu đề rồi nhấn nút Numbering trên thẻ Home",
                "Chọn bức ảnh rồi nhấn phím Delete để xóa ảnh khỏi trang văn bản",
            ],
            "ans": 0,
        },
    ],

    "exam_tl": [
        {
            "level": "Thông hiểu",
            "points": "1,0 điểm",
            "lines": 5,
            "q": "Trong lúc tìm tư liệu làm bài thuyết trình về 'An toàn giao thông học đường', bạn Hà thấy hai nguồn: Nguồn A là văn bản hướng dẫn mới ban hành năm 2026 trên Cổng thông tin Ủy ban An toàn giao thông Quốc gia (.gov.vn); Nguồn B là bài đăng không rõ tác giả trên một trang mạng xã hội bán mũ bảo hiểm giá rẻ với tiêu đề giật gân. Dựa vào các tiêu chí đánh giá độ tin cậy của thông tin số, em hãy phân tích vì sao Hà nên chọn Nguồn A và nêu cách ghi trích dẫn đúng chuẩn.",
            "rubric": [
                "Phân tích đúng các tiêu chí khẳng định Nguồn A đáng tin cậy hơn Nguồn B (0,75 điểm): (1) Nguồn gốc & thẩm quyền: Nguồn A thuộc cơ quan nhà nước có thẩm quyền (.gov.vn), còn Nguồn B khuyết danh; (2) Mục đích: Nguồn A phổ biến quy định pháp luật khách quan, còn Nguồn B mang mục đích thương mại quảng cáo bán hàng và câu lượt xem; (3) Tính cập nhật: Nguồn A công bố mới nhất năm 2026.",
                "Nêu đúng cách ghi trích dẫn tôn trọng bản quyền (0,25 điểm): Ghi rõ tên bài viết/văn bản, tên cơ quan ban hành, đường dẫn URL trang web và ngày truy cập ở cuối trang chiếu/báo cáo.",
            ],
        },
        {
            "level": "Vận dụng",
            "points": "1,0 điểm",
            "lines": 6,
            "q": "Quan sát bảng tính 'GIAN HÀNG HỘI CHỢ 8A1' (biết Sách cũ: B3=40, C3=15000; Đồ thủ công: B4=25, C4=20000; Cây sen đá: B5=30, C5=25000; ô F2 lưu tỉ lệ trích quỹ 10%):\na) Hãy tính giá trị Thành tiền tại các ô D3, D4, D5, tổng Thành tiền cả gian hàng tại ô D6 (nêu công thức hàm tại D6) và tổng số tiền trích Quỹ thu được từ cả 3 mặt hàng.\nb) Nêu các bước thao tác sắp xếp danh sách 3 mặt hàng theo thứ tự Thành tiền giảm dần trên Excel.",
            "rubric": [
                "Ý a – Tính đúng số liệu và viết đúng công thức tổng (0,5 điểm): D3 = 40 × 15000 = 600.000; D4 = 25 × 20000 = 500.000; D5 = 30 × 25000 = 750.000; tại D6 viết =SUM(D3:D5) cho kết quả 1.850.000 đồng. Tiền trích quỹ: F3 = 60.000, F4 = 50.000, F5 = 75.000 -> Tổng quỹ = 185.000 đồng.",
                "Ý b – Nêu đúng quy trình sắp xếp giảm dần theo cột Thành tiền (0,5 điểm): Chọn vùng dữ liệu bảng A2:F5 -> mở thẻ Data -> chọn lệnh Sort -> tại mục Sort by chọn cột 'Thành tiền', tại mục Order chọn 'Largest to Smallest' (từ lớn đến nhỏ) -> nhấn OK (thứ tự sau sắp xếp: Cây sen đá -> Sách cũ -> Đồ thủ công).",
            ],
        },
        {
            "level": "Vận dụng cao",
            "points": "1,0 điểm",
            "lines": 6,
            "q": "Từ kết quả bảng tính 'GIAN HÀNG HỘI CHỢ 8A1':\na) Giả sử một bạn quên gõ dấu $ nên nhập công thức =D3*F2 vào ô F3 rồi kéo sao chép xuống ô F4. Hãy cho biết công thức sinh ra tại ô F4 là gì, kết quả sai lệch ra sao và có thể thay $F$2 bằng địa chỉ hỗn hợp nào mà khi sao chép dọc từ F3 xuống F5 vẫn cho kết quả hoàn toàn chính xác?\nb) Lớp trưởng muốn vẽ 1 biểu đồ so sánh trực quan doanh thu Thành tiền giữa 3 mặt hàng và 1 biểu đồ thể hiện tỉ trọng phần trăm (%) đóng góp doanh thu của mỗi mặt hàng trong tổng doanh thu gian hàng. Em hãy tư vấn chọn đúng tên 2 loại biểu đồ tương ứng và giải thích lí do.",
            "rubric": [
                "Ý a – Phân tích đúng lỗi sao chép địa chỉ tương đối & đề xuất địa chỉ hỗn hợp (0,5 điểm): Khi sao chép =D3*F2 từ F3 xuống F4 (tăng 1 hàng), công thức biến thành =D4*F3 = 500.000 × 60.000 = 30.000.000.000 (sai lệch hoàn toàn vì nhân với tiền quỹ hàng trên thay vì 10%). Vì chỉ sao chép theo chiều dọc cùng cột F (chỉ thay đổi chỉ số hàng), ta chỉ cần cố định hàng 2 nên dùng địa chỉ hỗn hợp F$2 (=D3*F$2) vẫn đúng 100%.",
                "Ý b – Tư vấn đúng 2 loại biểu đồ theo mục đích thống kê (0,5 điểm): (1) Dùng Biểu đồ cột (Column Chart) để so sánh độ lớn doanh thu Thành tiền giữa 3 mặt hàng (nhìn chiều cao cột thấy ngay Cây sen đá cao nhất); (2) Dùng Biểu đồ hình quạt tròn (Pie Chart) để thể hiện cơ cấu tỉ lệ phần trăm (%) đóng góp của từng mặt hàng so với tổng 100% doanh thu cả gian hàng.",
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
            "q": "Chiếc máy tính cơ học Pascaline (1642) hoạt động dựa trên cơ chế vật lí nào để thực hiện phép cộng và phép trừ?",
            "options": [
                "Hệ thống các bánh răng truyền động cơ khí",
                "Các bóng đèn điện tử chân không cỡ lớn",
                "Các vi mạch tích hợp siêu nhỏ bằng silicon",
                "Mạng nơ-ron nhân tạo kết nối đám mây",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Ai được lịch sử tin học ghi nhận là lập trình viên đầu tiên trên thế giới khi viết thuật toán cho máy tính của Charles Babbage?",
            "options": [
                "Ada Lovelace",
                "Marie Curie",
                "Grace Hopper",
                "Blaise Pascal",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Máy tính điện tử thế hệ thứ hai (khoảng 1955 – 1965) sử dụng linh kiện chính nào thay thế đèn điện tử chân không?",
            "options": [
                "Bóng bán dẫn (Transistor)",
                "Bánh răng đồng hồ cơ học",
                "Siêu vi mạch (ULSI)",
                "Bàn tính gảy hạt gỗ",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Mạch tích hợp (IC – Integrated Circuit) là công nghệ nền tảng của máy tính điện tử thế hệ thứ mấy?",
            "options": [
                "Thế hệ thứ ba",
                "Thế hệ thứ nhất",
                "Thế hệ thứ hai",
                "Thế hệ thứ năm",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Đặc trưng nổi bật của máy tính điện tử thế hệ thứ năm (từ khoảng 1990 đến nay) là gì?",
            "options": [
                "Ứng dụng siêu vi mạch (ULSI), xử lí song song và phát triển Trí tuệ nhân tạo (AI)",
                "Kích thước chiếm trọn cả căn phòng lớn và tiêu thụ điện năng khổng lồ",
                "Chỉ thực hiện được hai phép tính cộng, trừ các số nguyên có 6 chữ số",
                "Nhập dữ liệu hoàn toàn bằng bìa đục lỗ cơ học mà không có màn hình",
            ],
            "ans": 0,
        },

        # Bài 2 (5 câu NB)
        {
            "bai": "Bài 2",
            "q": "Trong máy tính và các thiết bị kĩ thuật số, thông tin số được lưu trữ dưới dạng cơ bản nào?",
            "options": [
                "Được mã hóa thành các dãy bit (gồm các kí hiệu 0 và 1)",
                "Được khắc trực tiếp bằng mực in lên vỏ màn hình",
                "Chỉ lưu bằng các xung âm thanh cơ học trên giấy",
                "Tồn tại dưới dạng các bánh răng kim loại quay liên tục",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Tên miền nào dưới đây thường thuộc về các trường đại học, viện nghiên cứu và tổ chức giáo dục tại Việt Nam?",
            "options": [
                ".edu.vn",
                ".biz.vn",
                ".info",
                ".store",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Yếu tố nào dưới đây là một trong những tiêu chí quan trọng để đánh giá độ tin cậy của thông tin trên Internet?",
            "options": [
                "Thẩm quyền, uy tín của tác giả/cơ quan công bố và ngày cập nhật bài viết",
                "Số lượng hình ảnh động nhấp nháy quảng cáo xuất hiện dọc hai bên trang web",
                "Màu sắc phông chữ của tiêu đề bài viết càng sặc sỡ thì càng chính xác",
                "Độ dài của đường dẫn trang web càng ngắn thì mọi thông tin đều đúng tuyệt đối",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Vì sao khi đăng tải một thông tin sai sự thật lên mạng Internet thì hậu quả rất khó khắc phục triệt để?",
            "options": [
                "Vì thông tin số dễ sao chép, chụp lại màn hình và lan truyền sang hàng nghìn nơi chỉ trong vài phút",
                "Vì các máy chủ Internet không cho phép người đăng xóa bất kì bài viết nào trên trang cá nhân",
                "Vì màn hình máy tính của người đọc sẽ lưu vết mực vĩnh viễn không tắt được",
                "Vì trình duyệt web tự động dịch bài viết đó sang mọi ngôn ngữ lập trình",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "level": "NB",
            "q": "Thao tác đối chiếu cùng một sự kiện hoặc số liệu trên nhiều tờ báo chính thống và cơ quan chuyên môn độc lập gọi là gì?",
            "options": [
                "Kiểm chứng chéo thông tin",
                "Định dạng trang tính tự động",
                "Nén dung lượng tệp đa phương tiện",
                "Sao lưu hệ điều hành máy tính",
            ],
            "ans": 0,
        },

        # Bài 3 (5 câu NB)
        {
            "bai": "Bài 3",
            "q": "Bước đầu tiên cần thực hiện khi khai thác thông tin số để làm bài thuyết trình theo chủ đề là gì?",
            "options": [
                "Xác định rõ yêu cầu nội dung và lựa chọn từ khóa tìm kiếm phù hợp",
                "Chèn ngay các hiệu ứng chuyển trang vào tệp trình chiếu trống",
                "In toàn bộ kết quả trang đầu tiên của Google ra giấy A4",
                "Xóa lịch sử trình duyệt web trước khi mở phần mềm Word",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Công cụ lọc 'Tools -> Time' (Công cụ -> Thời gian) trên máy tìm kiếm Google giúp người dùng làm gì?",
            "options": [
                "Giới hạn kết quả tìm kiếm mới được đăng tải trong 24 giờ qua, tuần qua hoặc năm qua",
                "Đo thời gian gõ bàn phím của người dùng tính bằng mili giây",
                "Hẹn giờ tự động tắt nguồn máy tính sau khi tải tệp xong",
                "Đổi múi giờ hệ thống trên thanh tác vụ của Windows",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Từ khóa nào dưới đây giúp tìm kiếm thu hẹp và sát nhất thông tin về điểm chuẩn lớp 10 tại Hà Nội năm 2026?",
            "options": [
                "\"điểm chuẩn lớp 10\" Hà Nội 2026",
                "trường học",
                "thi cử Việt Nam",
                "học sinh cấp ba",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Khi lưu trữ các tư liệu hình ảnh, văn bản tìm được từ Internet về máy tính để làm dự án nhóm, em nên làm gì cùng lúc?",
            "options": [
                "Đặt tên tệp gợi nhớ, xếp vào thư mục chủ đề và lưu lại đường dẫn nguồn trích dẫn",
                "Giữ nguyên tên tệp ngẫu nhiên dài dòng và để lẫn lộn ngoài màn hình Desktop",
                "Đổi đuôi tất cả các tệp văn bản và hình ảnh thành đuôi .exe",
                "Xóa ngay tên tác giả gốc để không ai biết tư liệu lấy từ đâu",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Thông tin tìm được trên mạng có tính 'thời sự, cập nhật' đặc biệt quan trọng đối với chủ đề nào sau đây?",
            "options": [
                "Dự báo thời tiết bão lũ và lịch thi tuyển sinh năm học hiện tại",
                "Công thức tính diện tích hình chữ nhật trong hình học phẳng",
                "Bảng cửu chương nhân từ 2 đến 9 của môn Toán tiểu học",
                "Năm sinh của nhà bác học Blaise Pascal ở thế kỉ XVII",
            ],
            "ans": 0,
        },

        # Bài 4 (5 câu NB)
        {
            "bai": "Bài 4",
            "q": "Hành vi tự ý lấy bài viết, đoạn văn hoặc sản phẩm sáng tạo của người khác rồi nhận làm tác phẩm của mình được gọi là gì?",
            "options": [
                "Đạo văn (vi phạm bản quyền tác giả)",
                "Kiểm chứng chéo thông tin",
                "Sao lưu dữ liệu đám mây",
                "Chia sẻ tài nguyên học liệu mở",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Khi tham quan bảo tàng hoặc khu di tích lịch sử gặp biển báo 'Cấm quay phim, chụp ảnh', cách ứng xử có văn hóa là gì?",
            "options": [
                "Nghiêm túc cất máy ảnh/điện thoại và tuân thủ quy định của ban quản lí",
                "Tắt đèn flash rồi lén chụp ảnh đăng lên mạng xã hội khoe chiến tích",
                "Đứng che khuất biển báo để bạn mình quay video phát trực tiếp",
                "Chụp ảnh hiện vật rồi chỉnh sửa màu khác đi để bảo vệ không nhận ra",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Hành động nào dưới đây vi phạm pháp luật và đạo đức khi sử dụng phần mềm máy tính?",
            "options": [
                "Phát tán công cụ bẻ khóa (crack) phần mềm thương mại có bản quyền",
                "Cài đặt phần mềm mã nguồn mở miễn phí LibreOffice phục vụ học tập",
                "Mua gói bản quyền chính hãng phần mềm diệt virus cho máy tính gia đình",
                "Sử dụng tài khoản Microsoft Office 365 do nhà trường cấp cho học sinh",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Biểu hiện nào thể hiện tính trung thực của học sinh trong môi trường học tập số?",
            "options": [
                "Tự lực làm bài kiểm tra trực tuyến, không mở tài liệu trái quy chế hay nhờ người làm hộ",
                "Nhờ trí tuệ nhân tạo viết toàn bộ bài văn nghị luận rồi chép nguyên xi nộp lấy điểm",
                "Đăng nhập hộ tài khoản học trực tuyến để điểm danh khống cho bạn trốn học",
                "Chụp đề thi đang làm trong phòng thi gửi ra nhóm chat bên ngoài nhờ giải bài",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Trước khi chia sẻ một đoạn video quay cảnh va chạm giao thông có rõ mặt và biển số xe người liên quan lên mạng, em cần lưu ý điều gì?",
            "options": [
                "Không tự ý đăng tải xâm phạm đời tư, tránh bình luận suy diễn sai sự thật gây tổn hại người trong cuộc",
                "Cắt ghép thêm nhạc giật gân và đặt tiêu đề bịa đặt để thu hút thật nhiều lượt xem",
                "Kêu gọi cộng đồng mạng truy lùng địa chỉ nhà riêng của người trong video",
                "Đăng kèm thông tin cá nhân chưa kiểm chứng để câu lượt chia sẻ",
            ],
            "ans": 0,
        },

        # Bài 5 (5 câu NB)
        {
            "bai": "Bài 5",
            "q": "Loại địa chỉ ô nào trong Excel tự động thay đổi cả chỉ số cột và chỉ số hàng khi sao chép công thức sang vị trí khác?",
            "options": [
                "Địa chỉ tương đối (ví dụ: C4)",
                "Địa chỉ tuyệt đối (ví dụ: $C$4)",
                "Hằng số văn bản cố định",
                "Tên trang tính mặc định Sheet1",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Kí hiệu nào được đặt trước tên cột và tên hàng để cố định địa chỉ ô trong phần mềm bảng tính Excel?",
            "options": [
                "Kí hiệu đô la ($)",
                "Kí hiệu phần trăm (%)",
                "Kí hiệu thăng (#)",
                "Kí hiệu a còng (@)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Trong địa chỉ hỗn hợp $D5, thành phần nào được giữ cố định không đổi khi sao chép công thức?",
            "options": [
                "Cố định cột D, còn hàng 5 có thể thay đổi",
                "Cố định hàng 5, còn cột D có thể thay đổi",
                "Cố định cả cột D lẫn hàng 5",
                "Không cố định thành phần nào cả",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Trong địa chỉ hỗn hợp E$2, thành phần nào được giữ cố định khi sao chép công thức sang ô khác?",
            "options": [
                "Cố định hàng 2, còn cột E thay đổi theo vị trí cột mới",
                "Cố định cột E, còn hàng 2 thay đổi theo vị trí hàng mới",
                "Cố định cả cột E lẫn hàng 2 ở mọi vị trí",
                "Tự động xóa dấu $ ngay sau khi nhấn phím Enter",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Tại ô C3 chứa công thức =$A$1*B3. Khi sao chép công thức này xuống ô C4 thì công thức tại ô C4 là gì?",
            "options": [
                "=$A$1*B4",
                "=$A$2*B4",
                "=$A$1*B3",
                "=A2*B4",
            ],
            "ans": 0,
        },

        # Bài 6 (5 câu NB)
        {
            "bai": "Bài 6",
            "q": "Nút lệnh có biểu tượng chữ A trên chữ Z kèm mũi tên đi xuống (Sort A to Z) trong thẻ Data dùng để làm gì?",
            "options": [
                "Sắp xếp dữ liệu theo thứ tự tăng dần (bảng chữ cái từ A đến Z hoặc số từ nhỏ đến lớn)",
                "Sắp xếp dữ liệu theo thứ tự giảm dần từ lớn đến nhỏ",
                "Xóa toàn bộ các chữ cái tiếng Anh trong trang tính",
                "Chuyển đổi chữ thường thành chữ in hoa trong ô tính",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Biểu tượng của nút lệnh Lọc dữ liệu (Filter) trong nhóm Sort & Filter trên thẻ Data có hình dạng gì?",
            "options": [
                "Hình chiếc phễu lọc",
                "Hình chiếc kính lúp",
                "Hình chiếc kéo cắt",
                "Hình cây chổi quét sơn",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Sau khi bật chế độ Filter cho bảng dữ liệu, dấu hiệu nào xuất hiện trên hàng tiêu đề các cột của bảng?",
            "options": [
                "Xuất hiện nút mũi tên nhỏ trỏ xuống ở góc phải mỗi ô tiêu đề cột",
                "Toàn bộ chữ trong bảng tự động chuyển sang phông chữ nghiêng màu đỏ",
                "Các đường kẻ ô lưới trên trang tính biến mất hoàn toàn",
                "Xuất hiện biểu đồ hình tròn đè lên chính giữa bảng dữ liệu",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Để hủy bỏ điều kiện lọc hiện tại và hiển thị trở lại đầy đủ tất cả các hàng dữ liệu trong bảng, ta thực hiện thao tác nào?",
            "options": [
                "Nhấn lệnh Clear trong nhóm Sort & Filter (hoặc chọn lại Select All trong phễu lọc)",
                "Nhấn phím Delete để xóa trắng vùng dữ liệu đang hiển thị",
                "Đóng phần mềm Excel mà không lưu tệp bảng tính",
                "Nhấn nút Add Level trong hộp thoại định dạng phông chữ",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Tùy chọn 'Top 10...' nằm trong danh mục Number Filters của công cụ lọc dữ liệu cho phép thực hiện điều gì?",
            "options": [
                "Lọc ra danh sách N hàng có giá trị số cao nhất (hoặc thấp nhất) trong cột được chọn",
                "Chỉ giữ lại đúng 10 chữ cái đầu tiên trong mỗi ô văn bản",
                "Tự động cộng thêm 10 điểm vào tất cả các ô trong cột điểm",
                "Chia bảng tính hiện tại thành 10 trang tính riêng biệt",
            ],
            "ans": 0,
        },

        # Bài 7 (5 câu NB)
        {
            "bai": "Bài 7",
            "q": "Ưu điểm nổi bật của việc trình bày dữ liệu bằng biểu đồ so với bảng số liệu thô nhiều hàng cột là gì?",
            "options": [
                "Trực quan, sinh động, giúp người xem dễ so sánh, nhận xét xu hướng và ghi nhớ nhanh",
                "Tự động sửa lại các lỗi chính tả tiếng Việt trong văn bản báo cáo",
                "Giúp tăng tốc độ vi xử lí CPU của máy tính lên gấp nhiều lần",
                "Ngăn không cho bất kì ai nhìn thấy con số thống kê của bảng",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Loại biểu đồ nào phù hợp nhất để so sánh số lượng huy chương đạt được giữa các khối lớp 6, 7, 8, 9 trong Hội khỏe Phù Đổng?",
            "options": [
                "Biểu đồ cột (Column Chart)",
                "Biểu đồ đoạn thẳng theo giây (Line Chart)",
                "Công cụ kiểm tra chính tả (Spelling)",
                "Hàm đếm ô số (COUNT)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Trên biểu đồ Excel, thành phần hiển thị con số giá trị cụ thể ngay trên đỉnh mỗi cột hoặc trên từng miếng quạt tròn gọi là gì?",
            "options": [
                "Nhãn dữ liệu (Data Labels)",
                "Tiêu đề biểu đồ (Chart Title)",
                "Tên trang tính (Sheet Tab)",
                "Thanh trạng thái (Status Bar)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Khi nhấp chọn một biểu đồ trong Excel, nút dấu cộng màu xanh (+) xuất hiện bên cạnh khung biểu đồ có tên là gì?",
            "options": [
                "Chart Elements (thêm/bớt các thành phần của biểu đồ)",
                "Fill Handle (sao chép công thức tự động)",
                "AutoSum (tính tổng nhanh)",
                "Zoom Slider (phóng to thu nhỏ trang tính)",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Nếu người dùng thay đổi một con số trong bảng dữ liệu nguồn đã dùng để vẽ biểu đồ trên Excel thì biểu đồ sẽ như thế nào?",
            "options": [
                "Biểu đồ tự động cập nhật lại hình dáng và độ cao/tỉ lệ tương ứng với số liệu mới",
                "Biểu đồ bị biến thành hình ảnh tĩnh chết cứng không bao giờ thay đổi",
                "Biểu đồ tự động bị xóa khỏi trang tính ngay lập tức",
                "Excel bắt buộc phải khởi động lại máy tính mới vẽ lại biểu đồ",
            ],
            "ans": 0,
        },

        # Bài 8a (5 câu NB)
        {
            "bai": "Bài 8a",
            "q": "Trong phần mềm Microsoft Word, hai nút lệnh tạo danh sách Bullets và Numbering nằm ở nhóm lệnh nào của thẻ Home?",
            "options": [
                "Nhóm Paragraph",
                "Nhóm Font",
                "Nhóm Clipboard",
                "Nhóm Editing",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "q": "Để chèn các hình đồ họa mẫu như hình chữ nhật bo góc, mũi tên khối, ngôi sao, hộp thoại đám mây vào văn bản Word, ta chọn lệnh nào?",
            "options": [
                "Thẻ Insert -> Shapes",
                "Thẻ Home -> Replace",
                "Thẻ Layout -> Margins",
                "Thẻ View -> Ruler",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "q": "Sau khi vẽ một hình khối (Shape) trong Word, để gõ trực tiếp dòng chữ tiêu đề vào bên trong lòng hình khối đó, ta nháy nút phải chuột vào hình và chọn lệnh nào?",
            "options": [
                "Add Text",
                "Crop to Shape",
                "Reset Picture",
                "Split Cells",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "q": "Lệnh Wrap Text trong thẻ Picture Format (hoặc Shape Format) của Word có tác dụng gì?",
            "options": [
                "Thiết lập cách bố trí dòng văn bản bao quanh hình ảnh hoặc hình đồ họa",
                "Đếm tổng số từ và số kí tự có trong toàn bộ tệp văn bản",
                "Kiểm tra độ phân giải màu sắc của màn hình máy tính",
                "Tự động dịch đoạn văn bản tiếng Việt sang tiếng Anh",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "level": "NB",
            "q": "Để đưa một đối tượng hình ảnh đang bị che khuất ở lớp dưới nổi lên lớp phía trên các hình vẽ khác trong Word, ta sử dụng lệnh nào?",
            "options": [
                "Bring Forward / Bring to Front",
                "Send Backward / Send to Back",
                "Align Bottom",
                "Compress Pictures",
            ],
            "ans": 0,
        },
    ],

    "review_th_mcq": [
        # Bài 1 (2 câu TH)
        {
            "bai": "Bài 1",
            "q": "Sự thay thế linh kiện từ Đèn điện tử chân không -> Bóng bán dẫn -> Mạch tích hợp IC -> Vi mạch VLSI/ULSI qua 5 thế hệ máy tính đã mang lại sự thay đổi như thế nào?",
            "options": [
                "Kích thước máy ngày càng nhỏ gọn, tiêu thụ ít điện năng hơn nhưng tốc độ tính toán và dung lượng lưu trữ tăng gấp hàng triệu lần",
                "Kích thước máy ngày càng to lớn hơn, tỏa nhiệt nhiều hơn và chỉ dành riêng cho các phòng thí nghiệm quân sự",
                "Tốc độ tính toán giảm dần để tiết kiệm điện nhưng vỏ máy được trang trí nhiều màu sắc hơn",
                "Loại bỏ hoàn toàn khả năng lưu trữ chương trình trong bộ nhớ của máy tính hiện đại",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 1",
            "q": "Điểm hạn chế lớn nhất của chiếc máy tính cơ học Pascaline so với ý tưởng máy tính đa năng của Charles Babbage là gì?",
            "options": [
                "Pascaline chỉ thực hiện cố định phép cộng, trừ số học cơ khí chứ chưa thể lập trình tự động giải quyết nhiều bài toán khác nhau",
                "Pascaline tốn quá nhiều điện năng làm cháy nổ cầu chì trong khi máy của Babbage chạy bằng pin tiểu",
                "Pascaline chỉ hiển thị kết quả trên màn hình cảm ứng mà không in ra giấy được",
                "Pascaline bắt buộc phải kết nối mạng Internet mới quay được các bánh răng",
            ],
            "ans": 0,
        },

        # Bài 2 (2 câu TH)
        {
            "bai": "Bài 2",
            "q": "Khi tra cứu thông tin về học bổng du học bậc phổ thông, bạn Nam đọc được một bài đăng quảng cáo 'Cam kết đậu học bổng 100% không cần xét điểm, chỉ cần đóng phí giữ chỗ trong hôm nay' trên một fanpage mới lập 2 ngày. Đánh giá nào sau đây chính xác nhất?",
            "options": [
                "Độ tin cậy rất thấp vì nguồn không thẩm quyền, nội dung phi thực tế đánh vào lòng tham và thúc ép nộp tiền lừa đảo",
                "Độ tin cậy tuyệt đối cao vì bài viết vừa mới được đăng tải cách đây đúng 2 ngày nên rất cập nhật",
                "Đáng tin cậy vì bài viết có sử dụng nhiều biểu tượng cảm xúc và phông chữ in hoa nổi bật",
                "Đáng tin cậy vì mọi thông tin xuất hiện trên mạng xã hội đều đã được Bộ Giáo dục kiểm duyệt trước",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 2",
            "q": "Bạn Lan tìm thấy hai bài báo cùng viết về quy định tuyển sinh lớp 10 của tỉnh nhà: Bài 1 đăng tháng 4/2021 và Bài 2 đăng tháng 3/2026 trên Cổng thông tin Sở GD&ĐT. Vì sao Lan phải ưu tiên sử dụng Bài 2?",
            "options": [
                "Vì quy chế tuyển sinh có thể thay đổi theo từng năm nên tiêu chí tính thời sự, cập nhật quyết định tính chính xác hiện hành",
                "Vì các bài viết đăng vào tháng 3 luôn chính xác hơn các bài viết đăng vào tháng 4 hằng năm",
                "Vì bài viết năm 2021 không thể sao chép văn bản sang phần mềm Microsoft Word",
                "Vì máy tính thế hệ thứ năm không đọc được các trang web xuất bản trước năm 2025",
            ],
            "ans": 0,
        },

        # Bài 3 (2 câu TH)
        {
            "bai": "Bài 3",
            "q": "Em cần tìm các tài liệu chính thống về 'Chuyển đổi số giáo dục' chỉ nằm trên các trang web của ngành giáo dục Việt Nam (.edu.vn). Cách gõ từ khóa tìm kiếm nâng cao nào trên Google hiệu quả nhất?",
            "options": [
                "\"Chuyển đổi số giáo dục\" site:edu.vn",
                "Chuyển đổi + số + giáo + dục .mp3",
                "site:com \"trò chơi giáo dục\"",
                "http://chuyendoiso@gmail.com",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 3",
            "q": "Trong quá trình làm dự án nhóm tìm hiểu về biến đổi khí hậu, bạn An tải 20 bài báo dài tổng cộng 100 trang rồi chép nguyên văn toàn bộ vào slide trình chiếu. Sai lầm chính trong kĩ năng khai thác thông tin của An là gì?",
            "options": [
                "Thiếu bước chọn lọc, đối chiếu và tổng hợp ý chính ngắn gọn phù hợp với mục đích trình chiếu",
                "Sử dụng máy tìm kiếm Google thay vì tự đi đo nhiệt độ ngoài trời bằng nhiệt kế",
                "Lưu tệp bài trình chiếu với phần mở rộng .pptx thay vì phần mở rộng .sb3",
                "Đặt cỡ chữ trên trang chiếu lớn hơn 10pt làm tốn dung lượng ổ cứng",
            ],
            "ans": 0,
        },

        # Bài 4 (2 câu TH)
        {
            "bai": "Bài 4",
            "q": "Bạn Khoa viết một đoạn nhạc điện tử ngắn bằng phần mềm trên máy tính và đăng công khai lên trang cá nhân. Bạn Tuấn tải đoạn nhạc đó về làm nhạc nền cho video bán hàng kiếm tiền của mình mà không xin phép và không ghi tên Khoa. Hành vi của Tuấn được đánh giá như thế nào?",
            "options": [
                "Vi phạm quyền tác giả vì tự ý khai thác thương mại sản phẩm trí tuệ của người khác khi chưa được phép",
                "Hoàn toàn hợp pháp vì cứ đăng lên mạng Internet ở chế độ công khai nghĩa là từ bỏ bản quyền",
                "Đúng chuẩn đạo đức số vì đoạn nhạc ngắn dưới 3 phút không được coi là tác phẩm số",
                "Chỉ vi phạm nếu Khoa đã in bản nhạc đó ra giấy và đóng dấu đỏ của nhà trường",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 4",
            "q": "Trong giờ tự học tại thư viện trường, cách sử dụng thiết bị kĩ thuật số nào sau đây thể hiện văn hóa ứng xử đúng mực nhất?",
            "options": [
                "Chuyển điện thoại sang chế độ im lặng, đeo tai nghe âm lượng vừa phải nếu cần nghe bài giảng trực tuyến",
                "Bật loa ngoài điện thoại thật to để xem video giải trí cùng các bạn ngồi bàn bên cạnh",
                "Gọi điện thoại video nói chuyện lớn tiếng suốt buổi trong phòng đọc sách chung",
                "Dùng điện thoại chụp lén các bạn đang đọc sách ở góc phòng để đăng ảnh chế",
            ],
            "ans": 0,
        },

        # Bài 5 (2 câu TH)
        {
            "bai": "Bài 5",
            "q": "Tại ô D2 có công thức =B2*$C$1+$A2. Khi sao chép công thức từ ô D2 sang ô E4 (dịch sang phải 1 cột và xuống dưới 2 hàng), công thức tại ô E4 sẽ là gì?",
            "options": [
                "=C4*$C$1+$A4",
                "=C4*$D$3+$B4",
                "=B4*$C$1+$A2",
                "=C2*$C$1+$B4",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 5",
            "q": "Bảng lương nhân viên có Hệ số lương ở cột B (B3:B20), Phụ cấp ở cột C (C3:C20), và mức Lương cơ sở dùng chung cho toàn cơ quan đặt cố định tại duy nhất ô F1. Công thức tính Lương tại ô D3 (= Hệ số lương × Lương cơ sở + Phụ cấp) viết như thế nào để kéo sao chép xuống D20 không bị sai?",
            "options": [
                "=B3*$F$1+C3",
                "=B3*F1+C3",
                "=$B$3*F1+$C$3",
                "=B3*$F1+C$3",
            ],
            "ans": 0,
        },

        # Bài 6 (2 câu TH)
        {
            "bai": "Bài 6",
            "q": "Giáo viên muốn sắp xếp danh sách học sinh lớp 8A1 sao cho học sinh có Điểm trung bình cao hơn xếp trên; nếu hai bạn có cùng Điểm trung bình thì bạn nào có Điểm Tin học cao hơn sẽ xếp trên. Thao tác nào trong hộp thoại Sort đáp ứng đúng yêu cầu này?",
            "options": [
                "Sort by chọn 'Điểm trung bình' (Largest to Smallest), nhấn Add Level rồi tại Then by chọn 'Điểm Tin học' (Largest to Smallest)",
                "Sort by chọn 'Điểm trung bình' (Smallest to Largest), sau đó bật nút Filter ẩn cột Điểm Tin học",
                "Thực hiện Sort cột 'Họ và tên' theo A to Z rồi nhấn Delete Level hai lần",
                "Dùng công cụ lọc Number Filters chọn Equals 10 trên cả hai cột cùng lúc",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 6",
            "q": "Một bảng danh sách gồm 35 học sinh từ hàng 3 đến hàng 37. Sau khi lọc cột 'Xếp loại rèn luyện' với điều kiện bằng 'Tốt', trên màn hình chỉ còn hiển thị 28 học sinh. Nhận định nào sau đây ĐÚNG về 7 học sinh còn lại?",
            "options": [
                "Các hàng của 7 học sinh đó đang tạm thời ẩn đi (số thứ tự hàng bên lề trái chuyển màu xanh và bị ngắt quãng), nhấn Clear Filter sẽ hiện lại đủ 35 bạn",
                "Dữ liệu của 7 học sinh đó đã bị xóa hẳn khỏi ổ cứng, muốn lấy lại bắt buộc phải gõ nhập lại từ đầu",
                "Excel đã tự động cắt 7 hàng đó chuyển sang lưu ở trang tính Sheet2 bên cạnh",
                "Điểm số của 7 học sinh đó đã bị hàm Filter đổi thành chữ 'Không đạt'",
            ],
            "ans": 0,
        },

        # Bài 7 (2 câu TH)
        {
            "bai": "Bài 7",
            "q": "Một bạn dùng Biểu đồ hình quạt tròn (Pie Chart) để biểu diễn đồng thời nhiệt độ cao nhất và nhiệt độ thấp nhất của 7 ngày trong tuần nhưng biểu đồ chỉ vẽ được một dãy số duy nhất và rất khó hiểu. Nguyên nhân là gì?",
            "options": [
                "Biểu đồ hình quạt tròn chỉ biểu diễn cơ cấu tỉ lệ % của một chuỗi dữ liệu tổng thể, muốn so sánh 2 chuỗi biến động theo ngày nên dùng Biểu đồ cột kép hoặc Biểu đồ đoạn thẳng",
                "Do phần mềm Excel không cho phép vẽ biểu đồ với các con số nhỏ hơn 50 độ C",
                "Do bạn chưa chuyển địa chỉ các ô nhiệt độ sang địa chỉ tuyệt đối có dấu $",
                "Do bảng dữ liệu chưa được lọc ẩn bớt các ngày cuối tuần Thứ Bảy và Chủ Nhật",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 7",
            "q": "Sau khi chèn xong một biểu đồ cột trong Excel, giáo viên nhận xét 'Người xem chưa biết biểu đồ này thống kê về chủ đề gì và chưa đọc được con số chính xác của từng cột'. Em cần bổ sung hai thành phần nào qua nút Chart Elements (+)?",
            "options": [
                "Chart Title (Tiêu đề biểu đồ) và Data Labels (Nhãn số liệu)",
                "Gridlines (Đường lưới mờ) và Error Bars (Thanh sai số)",
                "Formula Bar (Thanh công thức) và Name Box (Hộp tên)",
                "Sort (Sắp xếp) và Filter (Lọc dữ liệu)",
            ],
            "ans": 0,
        },

        # Bài 8a (2 câu TH)
        {
            "bai": "Bài 8a",
            "q": "Khi soạn thảo nội quy phòng máy gồm: (1) Danh sách 5 thiết bị có trong phòng (không phân biệt thứ tự trước sau) và (2) Quy trình 4 bước tắt máy tính an toàn (bắt buộc đúng trình tự). Cách chọn định dạng danh sách nào hợp lí nhất?",
            "options": [
                "Phần (1) dùng Bullets (dấu đầu dòng); Phần (2) dùng Numbering (đánh số thứ tự 1, 2, 3, 4)",
                "Phần (1) dùng Numbering đánh số La Mã; Phần (2) dùng công cụ cắt ảnh Crop",
                "Cả hai phần đều vẽ bằng bảng tính Excel rồi chụp ảnh dán vào mà không gõ chữ",
                "Phần (1) dùng Wrap Text; Phần (2) dùng lệnh Send to Back",
            ],
            "ans": 0,
        },
        {
            "bai": "Bài 8a",
            "q": "Em chèn một ảnh chân dung nhà bác học Blaise Pascal vào góc trái đoạn văn bản tiểu sử và muốn các dòng chữ tiểu sử tự động dàn đều bao sát quanh khung hình chữ nhật của bức ảnh. Em cần chọn chế độ nào trong mục Wrap Text?",
            "options": [
                "Square",
                "In Line with Text",
                "Behind Text",
                "Top and Bottom",
            ],
            "ans": 0,
        },
    ],

    "review_th_tl": [
        {
            "bai": "Bài 1",
            "q": "Trình bày tóm tắt 5 thế hệ máy tính điện tử (thời gian ước lượng, công nghệ linh kiện đặc trưng) và nêu vai trò lịch sử của dự án máy tính Charles Babbage.",
            "answer": "Năm thế hệ máy tính điện tử: (1) Thế hệ 1 (1945–1955): đèn điện tử chân không (máy ENIAC); (2) Thế hệ 2 (1955–1965): bóng bán dẫn Transistor; (3) Thế hệ 3 (1965–1974): mạch tích hợp IC; (4) Thế hệ 4 (1974–1990): vi mạch tích hợp cỡ rất lớn VLSI, ra đời vi xử lí và máy tính cá nhân; (5) Thế hệ 5 (1990–nay): siêu vi mạch ULSI, mạng toàn cầu và trí tuệ nhân tạo AI. Dự án máy tính giải tích của Charles Babbage (1833) lần đầu tiên đề xuất thiết kế máy tính đa năng thực hiện tự động theo chương trình với đầy đủ các bộ phận: thiết bị vào, bộ nhớ, bộ xử lí và thiết bị ra — trở thành mô hình kiến trúc nền tảng của máy tính hiện đại.",
        },
        {
            "bai": "Bài 2",
            "q": "Nêu các đặc điểm chính của thông tin số và trình bày 4 tiêu chí cơ bản giúp học sinh đánh giá độ tin cậy của thông tin tìm được trên môi trường số.",
            "answer": "Đặc điểm thông tin số: Mã hóa nhị phân, sao chép dễ dàng không giảm chất lượng, lan truyền cực nhanh, dung lượng lưu trữ lớn, đa phương tiện, dễ tìm kiếm/xử lí tự động nhưng rất khó thu hồi khi đã phát tán. Bốn tiêu chí đánh giá độ tin cậy: (1) Nguồn gốc & thẩm quyền (ưu tiên cơ quan nhà nước .gov.vn, cơ sở giáo dục .edu.vn, chuyên gia uy tín); (2) Mục đích thông tin (phổ biến khoa học, giáo dục khách quan đáng tin hơn bài quảng cáo, câu view); (3) Tính cập nhật (ngày đăng/sửa đổi mới nhất); (4) Trích dẫn chứng cứ rõ ràng và kiểm chứng chéo khớp với nhiều nguồn chính thống.",
        },
        {
            "bai": "Bài 3",
            "q": "Trình bày quy trình các bước tìm kiếm, chọn lọc và khai thác thông tin số có trách nhiệm để hoàn thành một bài báo cáo dự án học tập.",
            "answer": "Quy trình 4 bước chuẩn: Bước 1: Xác định chủ đề, mục tiêu báo cáo và lập bộ từ khóa chính xác (dùng dấu ngoặc kép \"...\", toán tử site: khi cần); Bước 2: Đánh giá và chọn lọc thông tin từ các nguồn uy tín (.gov.vn, .edu.vn, báo chính thống), đối chiếu số liệu mới nhất; Bước 3: Lưu trữ tư liệu vào cây thư mục khoa học và tổng hợp, diễn đạt lại thành dàn ý súc tích (không sao chép máy móc); Bước 4: Ghi đầy đủ nguồn trích dẫn (tên tác giả, tiêu đề, đường dẫn URL, ngày truy cập) để tôn trọng bản quyền.",
        },
        {
            "bai": "Bài 4",
            "q": "Nêu 4 hành vi cần tránh để đảm bảo đạo đức, pháp luật và văn hóa khi sử dụng điện thoại thông minh, máy tính và mạng Internet trong đời sống hằng ngày.",
            "answer": "Bốn hành vi cần tránh: (1) Không tự ý quay phim, chụp ảnh tại nơi có biển cấm hoặc đăng tải hình ảnh, thông tin cá nhân của người khác khi chưa được phép; (2) Không đạo văn, không dùng phần mềm bẻ khóa lậu hoặc xâm phạm bản quyền tác phẩm số; (3) Không bịa đặt, chia sẻ tin giả, bình luận vu khống, xúc phạm danh dự người khác trên mạng; (4) Không gian lận trong học tập/thi cử trực tuyến và không bật âm lượng thiết bị gây ồn ào nơi công cộng.",
        },
        {
            "bai": "Bài 5",
            "q": "Phân biệt sự khác nhau giữa địa chỉ tương đối, địa chỉ tuyệt đối và địa chỉ hỗn hợp trong Excel. Cho ví dụ minh họa từng loại địa chỉ.",
            "answer": "(1) Địa chỉ tương đối (ví dụ: B3): Không có dấu $, tự động thay đổi cả tên cột và tên hàng tương ứng theo độ dời khi sao chép công thức. (2) Địa chỉ tuyệt đối (ví dụ: $F$2): Có dấu $ trước cả tên cột và tên hàng, luôn giữ nguyên cố định tọa độ ô dù sao chép công thức đi bất cứ đâu. (3) Địa chỉ hỗn hợp (ví dụ: $B3 hoặc B$3): Chỉ đặt dấu $ trước một thành phần — $B3 cố định cột B (hàng thay đổi), còn B$3 cố định hàng 3 (cột thay đổi) khi sao chép công thức.",
        },
        {
            "bai": "Bài 6",
            "q": "So sánh mục đích và kết quả hiển thị của hai thao tác Sắp xếp dữ liệu (Sort) và Lọc dữ liệu (Filter) trong phần mềm bảng tính Excel.",
            "answer": "Sắp xếp dữ liệu (Sort) giữ nguyên toàn bộ số lượng hàng của bảng nhưng thay đổi thứ tự vị trí các hàng theo giá trị tăng dần hoặc giảm dần của một hay nhiều cột tiêu chí để dễ tra cứu, xếp hạng. Trong khi đó, Lọc dữ liệu (Filter) giữ nguyên thứ tự gốc nhưng chỉ hiển thị những hàng thỏa mãn điều kiện đặt ra và tạm thời ẩn đi các hàng không thỏa mãn (không làm mất dữ liệu gốc).",
        },
        {
            "bai": "Bài 7",
            "q": "Khi nào nên sử dụng Biểu đồ cột (Column), Biểu đồ đoạn thẳng (Line) và Biểu đồ hình quạt tròn (Pie)? Nêu ví dụ thực tế cho mỗi loại biểu đồ.",
            "answer": "(1) Biểu đồ cột (Column Chart): Dùng khi muốn so sánh độ lớn dữ liệu giữa các đối tượng độc lập (ví dụ: so sánh số cây xanh trồng được của 4 lớp 8A1, 8A2, 8A3, 8A4). (2) Biểu đồ đoạn thẳng (Line Chart): Dùng khi muốn theo dõi xu hướng tăng giảm, biến thiên liên tục theo thời gian (ví dụ: biến động lượng mưa trung bình 12 tháng trong năm). (3) Biểu đồ hình quạt tròn (Pie Chart): Dùng khi muốn thể hiện cơ cấu tỉ lệ phần trăm (%) của từng thành phần trong một tổng thể 100% (ví dụ: tỉ lệ % học sinh xếp loại Tốt, Khá, Đạt của khối 8).",
        },
        {
            "bai": "Bài 8a",
            "q": "Phân biệt trường hợp sử dụng danh sách dấu đầu dòng (Bullets) với danh sách có thứ tự (Numbering) trong Word, và nêu công dụng của các công cụ Crop, Wrap Text, Bring Forward / Send Backward khi xử lí hình ảnh.",
            "answer": "Bullets dùng liệt kê các mục ngang hàng không yêu cầu thứ tự trước sau (ví dụ: danh sách đồ dùng cần mang khi đi dã ngoại); Numbering dùng liệt kê các bước quy trình tuần tự hoặc bảng xếp hạng có thứ tự (ví dụ: các bước sơ cứu vết thương). Công cụ xử lí ảnh: Crop cắt bỏ vùng rìa thừa của ảnh; Wrap Text định kiểu chữ bao quanh ảnh hài hòa (Square, Tight, Top and Bottom...); Bring Forward / Send Backward điều chỉnh lớp hiển thị đưa đối tượng lên lớp trên hoặc lùi xuống lớp dưới khi chồng lên nhau.",
        },
    ],

    "review_vd_tl": [
        {
            "bai": "Bài 1",
            "q": "Hãy lấy 3 ví dụ thực tế trong trường học, bệnh viện và giao thông hiện nay cho thấy máy tính thế hệ thứ năm (tích hợp Trí tuệ nhân tạo AI) đang hỗ trợ con người hiệu quả như thế nào.",
            "answer": "Ba ví dụ thực tế: (1) Trong trường học: Hệ thống học tập thông minh cá nhân hóa bài tập theo năng lực từng học sinh, phần mềm nhận diện giọng nói chấm phát âm ngoại ngữ tức thời; (2) Trong bệnh viện: Hệ thống AI phân tích ảnh chụp X-quang, cắt lớp CT hỗ trợ bác sĩ phát hiện sớm tổn thương với độ chính xác cao; (3) Trong giao thông: Bản đồ số phân tích mật độ xe theo thời gian thực để gợi ý lộ trình tránh ùn tắc và hệ thống cảnh báo va chạm thông minh trên ô tô.",
        },
        {
            "bai": "Bài 2",
            "q": "Bạn Quân thấy trên mạng lan truyền thông tin 'Uống nước cốt chanh đậm đặc lúc đói thay bữa sáng giúp chữa khỏi mọi bệnh cận thị trong 7 ngày'. Vận dụng các tiêu chí đánh giá độ tin cậy của thông tin số, em hãy hướng dẫn Quân kiểm chứng và bác bỏ tin giả này.",
            "answer": "Áp dụng 4 tiêu chí kiểm chứng: (1) Nguồn gốc: Tin lan truyền trôi nổi, không phải khuyến cáo từ Viện Dinh dưỡng hay Bệnh viện Mắt Trung ương (.gov.vn); (2) Nội dung & chứng cứ khoa học: Khẳng định 'chữa khỏi mọi bệnh cận thị trong 7 ngày' là phóng đại phản khoa học, không có nghiên cứu y khoa trích dẫn (thậm chí axit chanh lúc đói gây hại dạ dày); (3) Mục đích: Câu tương tác hoặc dẫn dắt bán thực phẩm chức năng; (4) Kiểm chứng chéo: Tra cứu trên Cổng thông tin Bộ Y tế xác nhận cận thị khúc xạ không thể chữa khỏi bằng nước chanh -> Kết luận là tin sai sự thật, tuyệt đối không làm theo hay chia sẻ.",
        },
        {
            "bai": "Bài 3",
            "q": "Nhóm em được giao làm bài trình chiếu chủ đề 'Bảo vệ di sản thiên nhiên Vịnh Hạ Long'. Em hãy viết 3 câu lệnh tìm kiếm từ khóa hiệu quả trên Google để thu thập số liệu diện tích/đảo đá, giải pháp bảo vệ môi trường biển và mẫu ghi nguồn trích dẫn chuẩn cho trang chiếu cuối.",
            "answer": "Ba cụm từ khóa hiệu quả: (1) \"Di sản thiên nhiên thế giới Vịnh Hạ Long\" diện tích số lượng đảo site:gov.vn; (2) \"bảo vệ môi trường Vịnh Hạ Long\" giải pháp học sinh Quảng Ninh; (3) \"Ban Quản lí Vịnh Hạ Long\" số liệu du lịch 2025 2026. Mẫu ghi nguồn trích dẫn chuẩn ở trang cuối: 'Nguồn tham khảo: Cổng thông tin điện tử Ban Quản lí Vịnh Hạ Long (halongbay.com.vn / quangninh.gov.vn), bài viết Tổng quan giá trị di sản Vịnh Hạ Long, truy cập ngày 15/11/2026'.",
        },
        {
            "bai": "Bài 4",
            "q": "Trong buổi biểu diễn văn nghệ chào mừng 20/11 của trường, bạn Minh quay được đoạn video một bạn lớp bên cạnh vô tình bị vấp ngã trên sân khấu. Bạn bên cạnh xúi Minh ghép nhạc chế hài hước rồi đăng lên TikTok để lên xu hướng. Nếu là Minh, em sẽ xử lí như thế nào và giải thích lí do?",
            "answer": "Cách xử lí của Minh: Kiên quyết từ chối ghép nhạc chế hay đăng đoạn sự cố đó lên mạng xã hội, đồng thời xóa đoạn quay cảnh vấp ngã và chỉ giữ lại những khoảnh khắc biểu diễn đẹp để gửi tặng lớp bạn (sau khi xin phép). Lí do: Đăng cảnh sự cố ngoài ý muốn kèm nhạc chế làm bạn xấu hổ, tổn thương tâm lí, xâm phạm quyền hình ảnh cá nhân và trái với đạo đức, văn hóa ứng xử nhân văn trong môi trường số.",
        },
        {
            "bai": "Bài 5",
            "q": "Cửa hàng điện máy lập bảng tính tiền điện tiêu thụ cho 4 phòng học từ Hàng 4 đến Hàng 7: Cột B là Chỉ số cũ, Cột C là Chỉ số mới, Cột D là Điện năng tiêu thụ (kWh), Cột E là Thành tiền (đồng), biết Đơn giá 1 kWh điện lưu cố định tại ô E2 = 2500. Hãy viết công thức đúng tại ô D4, ô E4 và cho biết công thức tại D7, E7 sau khi kéo sao chép.",
            "answer": "Tại ô D4 (Điện năng tiêu thụ = Chỉ số mới - Chỉ số cũ): viết công thức =C4-B4 (dùng địa chỉ tương đối). Tại ô E4 (Thành tiền = Điện năng tiêu thụ × Đơn giá cố định tại E2): viết công thức =D4*$E$2 (dùng địa chỉ tuyệt đối $E$2). Khi kéo sao chép xuống hàng 7: ô D7 có công thức =C7-B7; ô E7 có công thức =D7*$E$2.",
        },
        {
            "bai": "Bài 6",
            "q": "Cho bảng kết quả Hội thao khối 8 (vùng A2:E12) gồm các cột: STT, Họ và tên, Lớp, Nội dung thi, Thành tích (điểm). Em hãy nêu thao tác chi tiết để: (a) Chỉ lọc ra danh sách các vận động viên của lớp '8A1' có Thành tích từ 8.5 trở lên; (b) Sắp xếp toàn bảng theo Lớp tăng dần (8A1 -> 8A4), trong cùng một lớp thì Thành tích giảm dần.",
            "answer": "(a) Lọc theo 2 điều kiện kết hợp: Chọn vùng A2:E12 -> thẻ Data -> nhấn Filter -> nhấp mũi tên tại cột 'Lớp' chỉ tích chọn '8A1' -> nhấp tiếp mũi tên tại cột 'Thành tích' chọn Number Filters -> Greater Than Or Equal To... -> nhập 8.5 -> nhấn OK. (b) Sắp xếp 2 cấp độ: Nhấn Clear hủy lọc trước -> chọn vùng A2:E12 -> thẻ Data -> Sort -> tại Sort by chọn cột 'Lớp' (Order: A to Z) -> nhấn Add Level -> tại Then by chọn cột 'Thành tích' (Order: Largest to Smallest) -> nhấn OK.",
        },
        {
            "bai": "Bài 7",
            "q": "Bảng thống kê Câu lạc bộ ngoại khóa của 40 học sinh lớp 8A1 như sau: CLB Tin học & Robotics: 14 bạn; CLB Tiếng Anh: 10 bạn; CLB Bóng rổ: 10 bạn; CLB Mỹ thuật: 6 bạn. Em hãy tính tỉ lệ phần trăm (%) của từng CLB và trình bày các bước vẽ biểu đồ hình quạt tròn hiển thị sẵn tỉ lệ % trên Excel.",
            "answer": "Tính tỉ lệ phần trăm trên tổng 40 học sinh: CLB Tin học & Robotics = 14/40 × 100% = 35%; CLB Tiếng Anh = 10/40 × 100% = 25%; CLB Bóng rổ = 10/40 × 100% = 25%; CLB Mỹ thuật = 6/40 × 100% = 15% (Tổng = 100%). Các bước vẽ trên Excel: Nhập bảng 2 cột (Tên CLB và Số lượng học sinh tại A2:B6) -> bôi đen vùng A2:B6 -> mở thẻ Insert -> nhấp biểu tượng Insert Pie Chart -> chọn kiểu 2-D Pie -> nhấp nút Chart Elements (+) tích chọn Data Labels (chọn Percentage) và sửa Chart Title thành 'CƠ CẤU THAM GIA CLB NGOẠI KHÓA LỚP 8A1'.",
        },
        {
            "bai": "Bài 8a",
            "q": "Em được phân công thiết kế một tờ rơi A4 trên Word tuyên truyền '5K An toàn trên không gian mạng' có 1 ảnh minh họa đặt bên phải, 5 quy tắc đánh số bên trái và 1 huy hiệu ngôi sao ghi chữ 'UNIGO 8A1' đè lên góc trên của bức ảnh. Hãy trình bày các công cụ Word cần phối hợp sử dụng.",
            "answer": "Quy trình phối hợp công cụ Word: (1) Gõ tiêu đề lớn căn giữa, sau đó dùng công cụ Numbering (thẻ Home -> Paragraph) để soạn danh sách 5 quy tắc đánh số tự động 1..5 rõ ràng; (2) Chọn Insert -> Pictures chèn ảnh minh họa, dùng Crop cắt gọn viền thừa và chọn Wrap Text -> Square rồi kéo ảnh sang lề phải ngang hàng với 5 quy tắc; (3) Chọn Insert -> Shapes vẽ hình ngôi sao, nháy phải chọn Add Text gõ 'UNIGO 8A1', chọn Wrap Text -> In Front of Text và dùng Bring to Front để huy hiệu nổi rõ lên trên góc bức ảnh.",
        },
    ],

    "review_vdc_tl": [
        {
            "bai": "Bài 1",
            "q": "Nhiều người lo ngại rằng sự bùng nổ của máy tính thế hệ thứ năm và Trí tuệ nhân tạo (AI) sẽ thay thế hoàn toàn con người trong mọi công việc. Dưới góc nhìn lịch sử phát triển công cụ tính toán, em hãy nêu quan điểm phản biện và đề xuất cách học sinh THCS cần chuẩn bị để làm chủ công nghệ.",
            "answer": "Quan điểm phản biện: Lịch sử từ máy cơ học Pascaline đến máy tính AI thế hệ 5 cho thấy máy tính là công cụ do con người sáng tạo để giải phóng lao động tính toán lặp lại và mở rộng năng lực trí tuệ, nhưng AI vẫn hoạt động dựa trên dữ liệu và thuật toán nên thiếu sự thấu cảm đạo đức, trách nhiệm pháp lí và tư duy sáng tạo nguyên bản của con người. Cách chuẩn bị của học sinh THCS: Nắm vững kiến thức nền tảng Tin học và tư duy thuật toán, rèn kĩ năng kiểm chứng kết quả do AI sinh ra, trau dồi ngoại ngữ, khả năng giải quyết vấn đề sáng tạo và đạo đức số để dùng AI làm trợ lí thông minh thay vì phụ thuộc thụ động.",
        },
        {
            "bai": "Bài 2",
            "q": "Một video cắt ghép bằng công nghệ Deepfake giả mạo giọng nói và khuôn mặt của một chuyên gia tài chính kêu gọi người dân đầu tư tiền qua ứng dụng lạ thu hút hàng chục nghìn lượt chia sẻ. Em hãy phân tích vì sao thông tin số đa phương tiện ngày nay dễ gây nhầm lẫn và xây dựng bộ '3 câu hỏi tự kiểm tra' trước khi tin vào bất kì video nào trên mạng.",
            "answer": "Nguyên nhân dễ gây nhầm lẫn: Công nghệ xử lí ảnh/âm thanh số và AI hiện đại có thể tạo ra video giả mạo sống động như thật, kết hợp hiệu ứng đám đông lượt xem ảo khiến người xem mất cảnh giác nếu chỉ tin vào mắt thấy tai nghe. Bộ '3 câu hỏi tự kiểm tra': (1) Nguồn đăng video này có phải kênh chính thức có xác thực của cơ quan nhà nước/đài truyền hình uy tín không? (2) Nội dung video có lợi ích bất thường, hối thúc chuyển tiền hay cài ứng dụng ngoài kho chính thống không? (3) Khi tra cứu tên sự kiện trên các báo chính thống (.gov.vn, VTV, TTXVN) có thông tin xác nhận hay cảnh báo lừa đảo không?",
        },
        {
            "bai": "Bài 3",
            "q": "Khi sử dụng công cụ tìm kiếm và chatbot AI để tra cứu tài liệu làm báo cáo Khoa học tự nhiên, bạn Hải nhận thấy chatbot đưa ra một con số thống kê kèm tên một cuốn sách tham khảo, nhưng khi tìm trên thư viện và Google Scholar thì cuốn sách đó hoàn toàn không tồn tại. Hãy giải thích hiện tượng này và nêu nguyên tắc khai thác thông tin an toàn.",
            "answer": "Giải thích hiện tượng: Các công cụ sinh văn bản tự động có thể gặp hiện tượng 'ảo giác thông tin' (tự ghép nối các từ ngữ nghe rất hợp lí nhưng số liệu sai lệch hoặc tự bịa ra nguồn trích dẫn không có thật). Nguyên tắc khai thác an toàn: Chỉ dùng công cụ gợi ý để định hướng ý tưởng ban đầu; mọi con số thống kê, mốc lịch sử, công thức khoa học và tài liệu trích dẫn bắt buộc phải mở kiểm tra trực tiếp từ nguồn gốc chính thống (.gov.vn, .edu.vn, SGK, cơ sở dữ liệu khoa học uy tín) trước khi đưa vào báo cáo.",
        },
        {
            "bai": "Bài 4",
            "q": "Câu lạc bộ Truyền thông của trường muốn chụp ảnh hoạt động giờ ra chơi của học sinh để đăng lên fanpage nhà trường lan tỏa năng lượng tích cực. Em hãy giúp CLB xây dựng 'Bộ quy tắc 4 điều chuẩn văn hóa & tôn trọng quyền riêng tư' cho các bạn cộng tác viên nhiếp ảnh học đường.",
            "answer": "Bộ quy tắc 4 điều chuẩn mực: (1) Xin phép và thông báo rõ mục đích truyền thông nhà trường cho thầy cô, các bạn học sinh trước khi chụp cận cảnh cá nhân; (2) Chỉ chọn đăng những hình ảnh đẹp, lịch sự, tích cực — tuyệt đối xóa và không phát tán các khoảnh khắc dìm hàng, nhạy cảm hoặc gây hiểu lầm; (3) Tôn trọng yêu cầu gỡ ảnh ngay lập tức nếu bạn có mặt trong ảnh không muốn công khai hình ảnh của mình; (4) Ghi rõ tên tác giả chụp ảnh/biên tập và không tự ý gắn kèm thông tin riêng tư (số điện thoại, địa chỉ nhà) của học sinh lên bài đăng.",
        },
        {
            "bai": "Bài 5",
            "q": "Người ta thiết kế bảng cửu chương nhân tự động dạng ma trận vuông trên Excel: Hàng 2 (từ B2 đến J2) lần lượt chứa các thừa số 1, 2, ..., 9; Cột A (từ A3 đến A11) cũng lần lượt chứa các thừa số 1, 2, ..., 9. Em hãy lập DUY NHẤT một công thức sử dụng địa chỉ hỗn hợp tại ô B3 sao cho khi kéo sao chép phủ kín toàn bộ vùng B3:J11 đều tự động tính đúng tích hàng × cột, và giải thích rõ cơ chế hoạt động.",
            "answer": "Công thức duy nhất tại ô B3: =$A3*B$2 (hoặc =B$2*$A3). Giải thích cơ chế: Thừa số thứ nhất nằm ở cột A (A3:A11) nên khi sao chép sang các cột C, D...J ta phải cố định cột A nhưng cho phép hàng chạy từ 3 đến 11 -> dùng $A3. Thừa số thứ hai nằm ở hàng 2 (B2:J2) nên khi sao chép xuống các hàng 4, 5...11 ta phải cố định hàng 2 nhưng cho phép cột chạy từ B đến J -> dùng B$2. Kiểm tra tại ô E6 (hàng 6 có A6=4, cột E có E2=4): công thức tự chuyển thành =$A6*E$2 = 4 × 4 = 16 chính xác tuyệt đối!",
        },
        {
            "bai": "Bài 6",
            "q": "Trên bảng điểm tổng kết 40 học sinh (vùng A3:F42, hàng 43 là dòng 'Điểm trung bình cả lớp' dùng hàm =AVERAGE(F3:F42)), bạn Bình bật Filter chọn lọc từ A3:F43 khiến hàng tổng kết thứ 43 bị lẫn vào dữ liệu lọc và sắp xếp nhảy lên đầu bảng. Hãy chỉ ra lỗi thao tác của Bình và hướng dẫn cách xử lí chuẩn xác.",
            "answer": "Nguyên nhân lỗi: Khi bôi đen vùng bao gồm cả hàng tổng kết cuối bảng (Hàng 43) để Sort/Filter, Excel coi hàng 'Điểm trung bình cả lớp' như một học sinh bình thường nên đem đi so sánh sắp xếp làm đảo lộn cấu trúc bảng. Cách xử lí chuẩn: Tách biệt hàng tổng kết khỏi vùng dữ liệu chi tiết (hoặc đặt ô thống kê ở khu vực riêng / để cách 1 hàng trống), khi thực hiện Sort hoặc Filter chỉ quét chọn đúng vùng bảng dữ liệu kèm tiêu đề cột là A2:F42 (không chọn hàng 43).",
        },
        {
            "bai": "Bài 7",
            "q": "Ban tổ chức chiến dịch 'Thu gom pin cũ bảo vệ môi trường' thu được số liệu của 4 tuần liên tiếp đối với 2 khối (Khối 7 và Khối 8): Tuần 1 (120 và 150 viên), Tuần 2 (180 và 210 viên), Tuần 3 (240 và 260 viên), Tuần 4 (310 và 380 viên). Hãy phân tích nên chọn biểu đồ nào nếu muốn: (TH1) Đối chiếu trực tiếp lượng pin giữa Khối 7 và Khối 8 ở từng tuần; (TH2) Nhấn mạnh đà tăng trưởng phong trào qua 4 tuần của cả hai khối; (TH3) Tính xem biểu đồ quạt tròn có vẽ chung trực tiếp bảng 2 cột × 4 hàng này được không?",
            "answer": "(TH1) Chọn Biểu đồ cột ghép đôi (Clustered Column Chart): đặt 2 cột Khối 7 và Khối 8 đứng sát nhau ở từng tuần giúp so sánh độ chênh lệch trực tiếp rất rõ. (TH2) Chọn Biểu đồ đoạn thẳng nhiều đường (Line Chart with Markers): hai đường gấp khúc đi lên qua 4 mốc tuần làm nổi bật xu hướng tăng trưởng mạnh mẽ của phong trào. (TH3) Không dùng trực tiếp 1 biểu đồ hình quạt tròn (Pie Chart) cho bảng nhiều chuỗi (2 khối × 4 tuần) được vì Pie Chart chỉ biểu diễn cơ cấu tỉ lệ % của duy nhất 1 chuỗi dữ liệu tổng hợp (ví dụ tổng cộng cả đợt Khối 7 gom 850 viên chiếm 45,95% và Khối 8 gom 1000 viên chiếm 54,05%).",
        },
        {
            "bai": "Bài 8a",
            "q": "Khi biên tập kỷ yếu lớp 8A1 trên Word, bạn Trang chèn 3 bức ảnh hoạt động vào 1 trang nhưng mỗi lần gõ thêm một dòng cảm nghĩ ở đầu trang thì cả 3 bức ảnh lại bị nhảy dồn sang trang sau làm vỡ hết bố cục, đồng thời 1 bức ảnh có chụp dính thùng rác ở góc phải. Em hãy hướng dẫn Trang quy trình xử lí dứt điểm cả hai vấn đề trên.",
            "answer": "Quy trình khắc phục triệt để: (1) Xử lí chi tiết thừa ở góc ảnh: Nhấp chọn bức ảnh -> mở thẻ Picture Format -> chọn công cụ Crop -> kéo thanh giới hạn cạnh phải thu hẹp vào trong để cắt bỏ phần thùng rác rồi nhấn Enter/Crop lần nữa. (2) Chống nhảy vỡ bố cục ảnh: Chuyển chế độ Wrap Text của các ảnh từ mặc định 'In Line with Text' sang 'Square' hoặc 'Tight' (kết hợp tùy chọn Fix Position on Page trong Layout Options nếu muốn khóa cố định vị trí ảnh trên trang giấy), đồng thời có thể nhóm (Group) các khung trang trí và ảnh liên quan lại với nhau để bố cục trang kỷ yếu luôn ổn định.",
        },
    ],
}
