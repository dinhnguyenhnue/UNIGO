# -*- coding: utf-8 -*-
"""
Script sinh tự động toàn bộ 70 Kế hoạch bài dạy (KHBD) cho CLB Robotics UNIGO
- Phân hệ 1 (Lớp 1 - 2): 35 buổi (Tuần 1-17: OLLO Initiate; Tuần 18-35: OLLO Kinder)
- Phân hệ 2 (Lớp 3 - 7): 35 buổi (Tuần 1-17: OLLO Kinder; Tuần 18-35: OLLO Excel)
"""

import sys, os, shutil
sys.path.append(r"D:\UNIGO\scripts")
from generate_clb_khbd_all import create_khbd_doc, CALENDAR, BASE_KHBD_DIR, DIR_LOP_1_2, DIR_LOP_3_7, sanitize_filename

# ----------------- TRACK 1: LỚP 1 - 2 -----------------
RAW_TRACK_1 = [
    # HK1: Initiate (1-17)
    (1, "Bài 1", "TẬP THỂ DỤC NÀO!", "Đu xà cơ học", "OLLO INITIATE",
     ["Kiến thức về các bài tập vận động thể chất và tác dụng của xà đơn, xà kép.",
      "Kỹ năng nhận diện linh kiện, chốt ghép thông minh trong bộ OLLO Initiate.",
      "Quy trình lắp ráp mô hình Đu xà thăng bằng vững chắc."],
     "Lắp ráp Đu xà cơ học mô phỏng chuyển động thể dục", "đòn bẩy và trục xoay đơn giản", "Bậc 1"),
    (2, "Bài 2", "CHÚ CÚN DỄ THƯƠNG", "Cún con bốn chân", "OLLO INITIATE",
     ["Đặc điểm sinh học nổi bật và tập tính thân thiện của loài chó.",
      "Kỹ năng tạo khung thân bốn chân vững chãi và các khớp nối cử động linh hoạt.",
      "Quy trình lắp ráp hoàn thiện robot Chú cún bốn chân vẫy đuôi."],
     "Lắp ráp Chú cún bốn chân với đầu và đuôi cử động", "khớp nối và phân bổ trọng tâm", "Bậc 1"),
    (3, "Bài 3", "TĂNG CƯỜNG SỨC KHỎE", "Robot đi bộ", "OLLO INITIATE",
     ["Lợi ích của hoạt động đi bộ hằng ngày đối với sức khỏe.",
      "Cơ chế chuyển động bước chân so le tiến về phía trước.",
      "Kỹ năng phối hợp các thanh giằng cơ học tạo bước đi thăng bằng."],
     "Lắp ráp Robot đi bộ hai chân so le", "cơ cấu thanh truyền và bước chân so le", "Bậc 1"),
    (4, "Bài 4", "CHÚ ỐC SÊN CHẬM CHẠP", "Robot ốc sên", "OLLO INITIATE",
     ["Đặc điểm cấu tạo vỏ xoắn ốc và cơ chế trườn bò của động vật thân mềm.",
      "Kỹ năng kết nối bánh xe ẩn dưới thân mô phỏng chuyển động bò êm ái.",
      "Tư duy liên hệ giữa hình dáng động vật và mô hình robot."],
     "Lắp ráp Chú ốc sên thân mềm với bánh xe trượt", "cơ cấu trượt và tạo hình vỏ xoắn", "Bậc 1"),
    (5, "Bài 5", "XE CẢNH SÁT TUẦN TRA", "Xe bốn bánh tuần tra", "OLLO INITIATE",
     ["Vai trò của lực lượng cảnh sát và các phương tiện chuyên dụng bảo đảm an ninh.",
      "Cấu tạo khung xe bốn bánh, trục xoay bánh xe và còi báo hiệu mô phỏng.",
      "Kỹ năng lắp ghép khung xe chịu lực và kiểm tra độ trơn tru của bánh xe."],
     "Lắp ráp Xe cảnh sát tuần tra bốn bánh vững chắc", "trục xoay bánh xe và khí động học", "Bậc 1"),
    (6, "Bài 6", "KHÁM PHÁ XE CỨU HỎA", "Xe thang cứu hỏa", "OLLO INITIATE",
     ["Cấu tạo cơ bản của xe chữa cháy: thang nâng cứu hộ, vòi phun và đèn tín hiệu.",
      "Cơ chế đòn bẩy nâng hạ và xoay thang cứu hộ bằng tay gạt cơ học.",
      "Quy tắc an toàn phòng cháy chữa cháy và xử lý khi gặp tình huống khẩn cấp."],
     "Lắp ráp Xe cứu hỏa với thang nâng vươn cao", "cơ cấu khớp xoay nâng hạ thang", "Bậc 1"),
    (7, "Bài 7", "DING DONG! NGƯỜI GIAO HÀNG ĐÃ ĐẾN", "Xe đẩy hàng", "OLLO INITIATE",
     ["Quy trình vận chuyển hàng hóa và công việc của người giao hàng trong xã hội.",
      "Cấu tạo thùng xe chứa hàng, tay nắm đẩy và bánh xe chịu tải.",
      "Kỹ năng phối hợp các tấm ghép tạo khoang chứa hàng chắc chắn."],
     "Lắp ráp Xe đẩy giao hàng tiện lợi", "khung chứa hàng và bánh xe chịu tải", "Bậc 1"),
    (8, "Bài 8", "THẾ GIỚI KHỦNG LONG", "Khủng long T-Rex", "OLLO INITIATE",
     ["Đặc điểm nhận dạng của loài khủng long bạo chúa T-Rex thời tiền sử.",
      "Cơ chế đối trọng giữa phần đầu to và chiếc đuôi dài giúp khủng long thăng bằng.",
      "Kỹ năng lắp ráp hàm răng đóng mở và chi trước mô phỏng."],
     "Lắp ráp Khủng long bạo chúa T-Rex với hàm đóng mở", "đối trọng thăng bằng và hàm kẹp", "Bậc 1"),
    (9, "Bài 9", "HỒ BƠI MÙA HÈ", "Vận động viên bơi lội", "OLLO INITIATE",
     ["Các kiểu bơi cơ bản, đồ dùng bảo hộ và kỹ năng phòng chống đuối nước.",
      "Cơ cấu quạt tay đối xứng rẽ nước tạo lực đẩy tiến tới.",
      "Đánh giá định kỳ 1: Kiểm tra kỹ năng nhận diện linh kiện và tốc độ thao tác."],
     "Lắp ráp Robot vận động viên bơi lội sải tay", "cơ cấu tay quay đối xứng", "Bậc 1"),
    (10, "Bài 10", "KHÁM PHÁ ĐẠI DƯƠNG", "Cá mập săn mồi", "OLLO INITIATE",
     ["Đặc điểm sinh học loài cá mập và vai trò quan trọng trong hệ sinh thái biển.",
      "Cấu tạo vây lưng, vây đuôi và cơ chế hàm răng đớp mồi uy lực.",
      "Ý thức bảo vệ môi trường biển và các loài sinh vật biển."],
     "Lắp ráp Cá mập săn mồi với thân uốn lượn", "khớp nối thân uốn và hàm mở", "Bậc 1"),
    (11, "Bài 11", "CHÚ CUA CỨNG CÁP", "Cua biển càng to", "OLLO INITIATE",
     ["Đặc điểm bò ngang đặc biệt và lớp vỏ giáp kiên cố của loài cua biển.",
      "Cơ chế đòn bẩy điều khiển đôi càng kẹp mở ra gập vào linh hoạt.",
      "Kỹ năng bố trí các chân bên sườn đối xứng giúp cua đứng vững."],
     "Lắp ráp Chú cua biển với đôi càng kẹp linh hoạt", "đòn bẩy cơ học điều khiển ngàm kẹp", "Bậc 1"),
    (12, "Bài 12", "CON MỰC THÔNG MINH", "Con mực đại dương", "OLLO INITIATE",
     ["Tập tính bơi lùi bằng phản lực nước và cơ chế phun mực tự vệ của loài mực.",
      "Cấu tạo thân thuôn dài và các xúc tu mềm dẻo bắt mồi.",
      "Kỹ năng ghép các thanh liên kết mềm tạo hình chuyển động uốn lượn."],
     "Lắp ráp Con mực đại dương với các xúc tu cử động", "khớp nối chuỗi và chuyển động đàn hồi", "Bậc 1"),
    (13, "Bài 13", "HOẠT ĐỘNG MÙA HÈ", "Thuyền chèo hai mái", "OLLO INITIATE",
     ["Kỹ thuật chèo thuyền kayak và rèn luyện thể lực trong các hoạt động dã ngoại.",
      "Hệ thống hai mái chèo liên động cơ học đẩy nước nhịp nhàng.",
      "Bồi dưỡng tinh thần thể thao và khả năng phối hợp đồng đội."],
     "Lắp ráp Thuyền chèo với hệ thống hai mái chèo đối xứng", "liên động đòn bẩy mái chèo", "Bậc 1"),
    (14, "Bài 14", "BẮT ĐẦU CHUYẾN HÀNH TRÌNH CÙNG TÀU HOẢ", "Đoàn tàu hỏa", "OLLO INITIATE",
     ["Lịch sử ngành đường sắt, cấu tạo đầu máy kéo và khớp nối các toa chở hàng.",
      "Nguyên lý chuyển động thẳng trên đường ray và bánh xe có gờ định hướng.",
      "Kỹ năng ghép nối nhiều module toa tàu chắc chắn không bị tuột."],
     "Lắp ráp Đoàn tàu hỏa với các toa kéo liên hoàn", "khớp nối cầu và chuyển động thẳng", "Bậc 1"),
    (15, "Bài 15", "PHÍA TRÊN BẦU TRỜI", "Máy bay cánh quạt", "OLLO INITIATE",
     ["Cấu tạo thân máy bay, cánh nâng khí động học và cánh quạt đẩy gió phía mũi.",
      "Nguyên lý tạo lực nâng của cánh máy bay khi lướt trong không khí.",
      "Nuôi dưỡng ước mơ chinh phục bầu trời và khám phá ngành hàng không."],
     "Lắp ráp Máy bay hai tầng cánh với cánh quạt xoay", "cánh nâng khí động và trục xoay cánh quạt", "Bậc 1"),
    (16, "Bài 16", "VŨ TRỤ RỘNG LỚN", "Tàu vũ trụ thám hiểm", "OLLO INITIATE",
     ["Các hành tinh trong Hệ Mặt Trời và nhiệm vụ nghiên cứu của tàu không gian.",
      "Cấu tạo tên lửa đẩy nhiều tầng, khoang điều khiển và cánh pin năng lượng.",
      "Khơi dậy niềm say mê thiên văn học và công nghệ vũ trụ."],
     "Lắp ráp Tàu vũ trụ thám hiểm không gian", "cấu trúc tháp đứng và cánh thu năng lượng", "Bậc 1"),
    (17, "Buổi 17", "ÔN TẬP VÀ TRIỂN LÃM ROBOT SÁNG TẠO HỌC KÌ 1", "Dự án Initiate HK1", "OLLO INITIATE",
     ["Hệ thống hóa toàn bộ kiến thức, linh kiện và kỹ thuật lắp ghép bộ OLLO Initiate.",
      "Kỹ năng thuyết trình tự tin giới thiệu mô hình robot cá nhân/nhóm.",
      "Đánh giá định kỳ 2: Đánh giá sản phẩm triển lãm và năng lực giải thích nguyên lý."],
     "Thiết kế và triển lãm mô hình sáng tạo OLLO Initiate", "tổng hợp kỹ thuật lắp ráp và thuyết trình", "Bậc 1"),

    # HK2: Nâng cấp lên Kinder (18-35)
    (18, "Bài 1", "HÃY NẤU NHỮNG MÓN ĂN NGON!", "Chiếc nồi thần kỳ", "OLLO KINDER",
     ["Làm quen bộ thiết bị OLLO Kinder; nhận diện linh kiện, hộp pin và động cơ Kinder.",
      "Tìm hiểu dụng cụ nhà bếp, nắp vung đóng mở và chất liệu cách nhiệt của nồi nấu ăn.",
      "Quy trình lắp ráp mô hình Chiếc nồi thần kỳ với nắp tự động mở/đóng."],
     "Lắp ráp Chiếc nồi thần kỳ với nắp nâng hạ tự động", "động cơ điện và cơ cấu bản lề mở nắp", "Bậc 1"),
    (19, "Bài 2", "HÀM RĂNG TRẮNG SÁNG", "Bàn chải thông minh", "OLLO KINDER",
     ["Cấu tạo hàm răng, thói quen chăm sóc sức khỏe và ngăn ngừa sâu răng đúng cách.",
      "Cơ cấu trục quay bàn chải làm sạch mảng bám trên bề mặt răng.",
      "Lắp ráp hoàn thiện robot Bàn chải vệ sinh răng miệng chuyển động xoay tròn."],
     "Lắp ráp Bàn chải thông minh xoay tròn", "trục quay bánh răng và thói quen vệ sinh", "Bậc 1"),
    (20, "Bài 3", "SỰ PHẢN XẠ ÁNH SÁNG", "Máy ảnh thông minh", "OLLO KINDER",
     ["Hiện tượng phản xạ ánh sáng, ống kính quang học và nút chụp hình cơ học.",
      "Cấu tạo buồng tối, gương phản xạ và cơ chế nút bấm lò xo.",
      "Kỹ năng lắp ghép mô hình Máy ảnh du lịch OLLO Kinder."],
     "Lắp ráp Máy ảnh thông minh có nút bấm cơ học", "cơ chế nút bấm đàn hồi và ống kính", "Bậc 1"),
    (21, "Bài 4", "ĐÀN GÀ CON ĐÁNG YÊU", "Gà con mổ thóc", "OLLO KINDER",
     ["Vòng đời của loài gà, đặc điểm của gà con và tập tính mổ hạt tìm mồi.",
      "Cơ chế đòn bẩy gật đầu mổ thóc và vẫy cánh đồng thời khi động cơ quay.",
      "Phát triển tư duy liên hệ giữa cơ cấu máy và cử động sinh học."],
     "Lắp ráp Gà con mổ thóc với cổ gật và cánh vỗ", "cơ cấu tay đòn lệch tâm tạo nhịp gật đầu", "Bậc 1"),
    (22, "Bài 5", "NHỮNG BẠN NHỎ LỄ PHÉP", "Robot cúi chào", "OLLO KINDER",
     ["Văn hóa chào hỏi lễ phép của các quốc gia và hành vi ứng xử văn minh.",
      "Cơ cấu gập thân 90 độ mô phỏng động tác cúi chào trang trọng.",
      "Giáo dục phẩm chất tôn trọng người lớn, lễ phép với thầy cô và bạn bè."],
     "Lắp ráp Robot cúi chào góc 90 độ văn minh", "khớp gập hông và cơ cấu hành trình chuyển động", "Bậc 1"),
    (23, "Bài 6", "ĐỘNG VẬT THÂN MỀM", "Ốc sên cơ học Kinder", "OLLO KINDER",
     ["Đặc điểm sinh học nâng cao của động vật thân mềm và lớp chất nhờn bảo vệ.",
      "Hệ truyền động bánh răng tạo bước trườn bò lượn sóng nhịp nhàng.",
      "Lắp ráp mô hình Ốc sên Kinder vận hành bằng động cơ điện."],
     "Lắp ráp Ốc sên cơ học Kinder bò tự động", "bánh răng giảm tốc và chuyển động lượn sóng", "Bậc 1"),
    (24, "Bài 7", "CHÚ THỎ TAI DÀI NHẢY NHÓT", "Thỏ bật nhảy", "OLLO KINDER",
     ["Cấu tạo đôi chân sau dài khỏe giúp loài thỏ bật nhảy thoát hiểm nhanh chóng.",
      "Cơ cấu đòn bẩy chân lò xo tích năng lượng tạo lực đẩy bật về phía trước.",
      "Quy trình lắp ráp và căn chỉnh thăng bằng khi tiếp đất của thỏ."],
     "Lắp ráp Chú thỏ bật nhảy với chân lò xo nhún", "cơ cấu đòn bẩy bật nhảy và thế năng đàn hồi", "Bậc 1"),
    (25, "Bài 8", "HIỆP SĨ ĐẤU KIẾM", "Hiệp sĩ dũng cảm", "OLLO KINDER",
     ["Tinh thần thượng võ, bộ áo giáp hiệp sĩ và luật thi đấu kiếm thể thao.",
      "Cơ cấu tay đòn vung kiếm dứt khoát điều khiển bằng thanh truyền.",
      "Rèn luyện tính kỷ luật, sự dũng cảm và chính trực."],
     "Lắp ráp Hiệp sĩ cầm kiếm với tay vung dứt khoát", "cơ cấu thanh truyền tay kiếm", "Bậc 1"),
    (26, "Bài 9", "MÔI TRƯỜNG BIỂN & CHÚ CÁ BƠI LỘI", "Cá rẽ sóng Kinder", "OLLO KINDER",
     ["Hệ sinh thái biển san hô và chức năng các vây điều hướng của loài cá.",
      "Cơ chế vây ngực và vây đuôi quẫy nước tạo lực đẩy tiến tới.",
      "Ý thức bảo vệ môi trường biển và hạn chế rác thải nhựa."],
     "Lắp ráp Chú cá nhiệt đới quẫy đuôi lướt sóng", "khớp đuôi uốn và động cơ truyền động", "Bậc 1"),
    (27, "Bài 10", "CÙNG CÂU CÁ NÀO!", "Cần câu ròng rọc", "OLLO KINDER",
     ["Môn thể thao câu cá dã ngoại, cấu tạo cần câu và an toàn bên bờ nước.",
      "Nguyên lý ròng rọc cuốn thả dây cáp và gắn đầu hút nam châm gắp cá.",
      "Vận dụng nguyên lý đòn bẩy và ròng rọc nâng vật nặng dễ dàng."],
     "Lắp ráp Cần câu cá với cụm ròng rọc cuốn dây", "ròng rọc cố định và dây tời cuốn cá", "Bậc 1"),
    (28, "Bài 11", "VẬN CHUYỂN ĐỒ VẬT", "Xe nâng hàng Kinder", "OLLO KINDER",
     ["Các dòng xe nâng trong kho logistics và nguyên tắc sắp xếp hàng an toàn.",
      "Cấu tạo khung nâng thẳng đứng và thanh trượt ray chuyển động êm ái.",
      "Đánh giá định kỳ 3: Kiểm tra kỹ năng lắp ráp khung cơ khí Kinder và tải trọng xe."],
     "Lắp ráp Xe nâng hàng OLLO Kinder nâng hạ kiện hàng", "ray trượt thẳng đứng và khung nâng pallet", "Bậc 1"),
    (29, "Bài 12", "SỨC MẠNH CỦA MÁY ỦI", "Xe ủi công trình", "OLLO KINDER",
     ["Máy ủi bánh xích công trường, cơ chế lưỡi ben thép san phẳng mặt bằng.",
      "Hệ thống đòn nâng lưỡi ben linh hoạt chịu lực va đập lớn.",
      "Nâng cao ý thức an toàn công trình xây dựng và bảo hộ lao động."],
     "Lắp ráp Máy ủi công trường với lưỡi ben nâng hạ", "cơ cấu đòn nâng thủy lực mô phỏng", "Bậc 1"),
    (30, "Bài 13", "MÁY BẮN ĐÁ THẦN TỐC", "Máy bắn đá công thành", "OLLO KINDER",
     ["Lịch sử máy cơ khí cổ đại, nguyên lý tích lũy thế năng và cánh tay đòn dài.",
      "Cơ chế khóa chốt kích hoạt phóng bóng xa chuẩn xác.",
      "Thử nghiệm góc phóng và đo khoảng cách bay của bóng."],
     "Lắp ráp Máy bắn đá công thành với chốt hãm phóng", "đòn bẩy thế năng và góc phóng parabol", "Bậc 1"),
    (31, "Bài 14", "CUNG TÊN ĐÀN HỒI", "Cung tên thể thao", "OLLO KINDER",
     ["Môn bắn cung Olympic, lực đàn hồi cánh cung và quỹ đạo bay chính xác.",
      "Ống định hướng mũi tên an toàn và cò kích hoạt giải phóng năng lượng.",
      "Rèn luyện sự kiên nhẫn, tập trung cao độ và căn chỉnh hồng tâm."],
     "Lắp ráp Cung tên thể thao có ống ngắm định hướng", "lực đàn hồi dây cung và khóa cò cơ học", "Bậc 1"),
    (32, "Bài 15", "XE BUÝT HAI TẦNG THÀNH PHỐ", "Xe buýt đô thị", "OLLO KINDER",
     ["Đặc trưng xe buýt hai tầng phục vụ du lịch đô thị và văn minh công cộng.",
      "Khung gầm kiên cố, cầu thang xoắn và cửa tự động đóng mở an toàn.",
      "Bồi dưỡng ý thức bảo vệ môi trường qua phương tiện giao thông xanh."],
     "Lắp ráp Xe buýt hai tầng vững chãi với cửa tự động", "kết cấu khung chịu lực hai tầng", "Bậc 1"),
    (33, "Bài 16", "CÁ SẤU THẬT NGẦU!", "Cá sấu đầm lầy", "OLLO KINDER",
     ["Tập tính phục kích dưới nước và sức mạnh cú đớp mồi của cá sấu.",
      "Hàm răng liên động đóng mở phát ra tiếng cơ khí mạnh mẽ.",
      "Phân tích liên hệ giữa cơ quan sinh học và thiết kế cơ cấu máy."],
     "Lắp ráp Robot Cá sấu cơ bắp với hàm đớp mồi uy lực", "bánh răng cam kẹp hàm đóng mở liên tục", "Bậc 1"),
    (34, "Buổi 34", "ÔN TẬP VÀ THIẾT KẾ DỰ ÁN ROBOT CUỐI NĂM", "Dự án cuối năm Kinder", "OLLO KINDER",
     ["Tổng kết toàn bộ kỹ năng lắp ghép cơ bản (Initiate) và nâng cao (Kinder).",
      "Thiết kế chế tạo mô hình robot tự chọn theo nhóm học sinh.",
      "Đánh giá định kỳ 4: Đánh giá dự án cuối năm, kỹ năng xử lý lỗi và tinh thần đồng đội."],
     "Thiết kế và hoàn thiện dự án robot sáng tạo cuối năm", "tích hợp đa cơ cấu máy và hoàn thiện sản phẩm", "Bậc 1"),
    (35, "Buổi 35", "NGÀY HỘI ROBOTICS UNIGO VÀ BẾ MẠC CLB", "Festival Robotics", "OLLO KINDER",
     ["Triển lãm Festival Robotics UNIGO: Trình diễn các sản phẩm xuất sắc.",
      "Tham gia thử thách vượt sa bàn robot và giao lưu chia sẻ thành tích.",
      "Lễ bế mạc, trao chứng nhận hoàn thành khóa học và khen thưởng."],
     "Tham gia Ngày hội Robotics UNIGO và bế mạc năm học", "trình diễn robot trên sa bàn và nhận chứng chỉ", "Bậc 1")
]

# ----------------- TRACK 2: LỚP 3 - 7 -----------------
RAW_TRACK_2 = [
    # HK1: Kinder (1-17)
    (1, "Bài 1", "CỖ MÁY LÀM SẠCH ĐƯỜNG PHỐ!", "Xe gom rác", "OLLO KINDER",
     ["Quy trình vệ sinh môi trường đô thị và cấu tạo robot xe gom rác tự động.",
      "Hệ thống chổi quét xoay, thùng chứa rác và khung xe dẫn động OLLO Kinder.",
      "Ý thức bảo vệ môi trường, phân loại rác thải tại nguồn."],
     "Lắp ráp Xe gom rác tự động với chổi quét xoay", "bánh răng truyền động chổi quét và dẫn động bánh xe", "Bậc 2"),
    (2, "Bài 2", "VÌ SAO CHIẾC XE NÀY KHÔNG CÓ MÁI?", "Xe đua mui trần", "OLLO KINDER",
     ["Nguyên lý khí động học, trọng tâm hạ thấp giúp xe thể thao ổn định ở tốc độ cao.",
      "Cấu tạo hệ thống trục lái trước chuyển hướng góc nghiêng linh hoạt.",
      "Kỹ năng cân chỉnh độ chụm bánh xe và cân bằng trọng lượng."],
     "Lắp ráp Xe đua thể thao mui trần với trục lái trước linh hoạt", "hệ trục lái chuyển hướng và khí động học khung gầm", "Bậc 2"),
    (3, "Bài 3", "CỖ MÁY BAY GIỮA KHÔNG TRUNG", "Trực thăng cánh quạt kép", "OLLO KINDER",
     ["Nguyên lý cánh quạt kép đồng trục triệt tiêu mômen quay tạo lực nâng trực thăng.",
      "Hệ thống truyền động bánh răng nón đổi hướng chuyển động 90 độ.",
      "Rèn luyện kỹ năng lắp ráp cơ cấu bánh răng nón và trục xoay đứng."],
     "Lắp ráp Trực thăng trinh sát với cánh quạt kép đồng trục", "bánh răng nón đổi hướng và trục đồng trục", "Bậc 2"),
    (4, "Bài 4", "CHIẾC XE ĐẦU ĐỜI CỦA BÉ", "Xe đẩy cân bằng", "OLLO KINDER",
     ["Cấu trúc tam giác cân bằng hình học và diện tích mặt chân đế chống lật xe.",
      "Cấu tạo tay vịn trợ lực và bốn bánh xe phân bố tải trọng đồng đều.",
      "Vận dụng kiến thức vật lý về trọng tâm vào thiết kế xe an toàn."],
     "Lắp ráp Xe đẩy cân bằng chống lật đa hướng", "diện tích mặt chân đế và phân bổ tải trọng", "Bậc 2"),
    (5, "Bài 5", "CHÚ CHIM ĐỘI MŨ ĐỎ", "Chim gõ kiến", "OLLO KINDER",
     ["Đặc tính bắt sâu bảo vệ rừng và cấu tạo giảm chấn xương sọ chim gõ kiến.",
      "Cơ cấu thanh truyền biến chuyển động quay tròn thành chuyển động gõ mỏ tịnh tiến.",
      "Tư duy phỏng sinh học trong thiết kế máy khoan đập công nghiệp."],
     "Lắp ráp Chim gõ kiến với mỏ gõ tịnh tiến liên tục", "cơ cấu biến chuyển động quay thành tịnh tiến", "Bậc 2"),
    (6, "Bài 6", "BẬC THẦY BAY LƯỢN", "Đại bàng sải cánh", "OLLO KINDER",
     ["Cấu tạo sải cánh khí động học và kỹ thuật chao lượn tận dụng luồng nhiệt.",
      "Khớp cánh nhiều bậc tự do mô phỏng nhịp đập cánh bay lượn uy lực.",
      "Rèn luyện kỹ năng ghép các thanh liên kết đa tầng đối xứng."],
     "Lắp ráp Đại bàng sải cánh với cơ cấu đập cánh đa khớp", "liên kết đa tầng tạo chuyển động lượn sóng cánh", "Bậc 2"),
    (7, "Bài 7", "NHỮNG NGƯỜI BẠN ĐÁNG YÊU", "Chó kéo xe tuyết", "OLLO KINDER",
     ["Sức bền của giống chó tuyết Husky và cấu tạo xe trượt tuyết vùng cực.",
      "Hệ dây kéo cân bằng lực kéo hai bên sườn xe và khả năng dẫn hướng.",
      "Nâng cao tinh thần làm việc nhóm và phân công vai trò lắp ghép."],
     "Lắp ráp Chó kéo xe chở hàng với dây giằng phân lực", "cơ cấu truyền lực kéo và ma sát trượt", "Bậc 2"),
    (8, "Bài 8", "CÙNG KHÁM PHÁ ROBOT PIKACHU", "Robot Pikachu lắc lư", "OLLO KINDER",
     ["Nguyên lý đĩa cam lệch tâm tạo nên dao động lắc lư vui nhộn.",
      "Cấu tạo khớp cổ và đuôi dao động theo chu kỳ bước chân của robot.",
      "Sáng tạo nâng cấp các chi tiết phụ trợ tăng tính thẩm mỹ."],
     "Lắp ráp Robot Pikachu chuyển động lắc lư theo nhịp", "cơ cấu cam lệch tâm tạo dao động chu kỳ", "Bậc 2"),
    (9, "Bài 9", "KHÁM PHÁ SINH VẬT ĐUÔI CONG ĐẶC BIỆT", "Cá ngựa đuôi cong", "OLLO KINDER",
     ["Cấu tạo xương đuôi dạng móc bám và tư thế bơi thẳng đứng của cá ngựa.",
      "Cơ cấu vây lưng rung tần số cao đẩy thân cá ngựa tiến thẳng.",
      "Đánh giá định kỳ 1: Kiểm tra kỹ năng phân tích cơ cấu cơ học và giải bài toán kỹ thuật."],
     "Lắp ráp Cá ngựa đuôi cong bơi thẳng đứng", "vây lưng rung tần số cao và đuôi bám giữ", "Bậc 2"),
    (10, "Bài 10", "TÌM HIỂU CƠ CHẾ CHUYỂN ĐỘNG TRÊN MẶT NƯỚC", "Cano lướt sóng", "OLLO KINDER",
     ["Lực đẩy Ac-si-met, lực cản dòng nước và nguyên lý đẩy của chân vịt Cano.",
      "Bánh lái phía sau chuyển hướng dòng nước chuẩn xác giúp Cano rẽ sóng.",
      "Thử nghiệm trên sa bàn mô phỏng luồng di chuyển đường thủy."],
     "Lắp ráp Cano lướt sóng chân vịt với bánh lái đuôi", "cánh quạt chân vịt và định hướng dòng đẩy", "Bậc 2"),
    (11, "Bài 11", "AI LÀ VẬN ĐỘNG VIÊN BƠI LỘI DƯỚI ĐẠI DƯƠNG", "Hải cẩu bơi lặn", "OLLO KINDER",
     ["Lớp mỡ dày giữ ấm và đôi chân vây dẻo dai giúp hải cẩu bơi lặn siêu hạng.",
      "Khớp cổ và vây trước uốn lượn nhịp nhàng đẩy cơ thể rẽ nước.",
      "Bồi dưỡng trách nhiệm bảo tồn thiên nhiên hoang dã."],
     "Lắp ráp Hải cẩu rẽ sóng với vây bơi linh hoạt", "cơ cấu vây bơi đa khớp và chuyển động uốn lượn", "Bậc 2"),
    (12, "Bài 12", "CÁNH TAY KHỔNG LỒ BIẾT NÂNG VẬT NẶNG", "Cần cẩu tháp 360", "OLLO KINDER",
     ["Cánh tay đòn tải trọng lớn, đối trọng thăng bằng và đĩa quay 360 độ của cần cẩu.",
      "Dây tời ròng rọc kép nâng hạ khối hàng nặng với lực kéo tối thiểu.",
      "Tính toán tải trọng giới hạn chống lật mô hình khi vươn cần xa."],
     "Lắp ráp Cần cẩu tháp xoay 360 độ với cụm ròng rọc kép", "ròng rọc kép giảm lực kéo và đối trọng thăng bằng", "Bậc 2"),
    (13, "Bài 13", "AI ĐANG DẤU MÌNH DƯỚI CÁT?", "Cá bơn ngụy trang", "OLLO KINDER",
     ["Cơ chế thích nghi đổi màu và thân hình dẹt nằm sát đáy biển của cá bơn.",
      "Cơ cấu vây diềm gợn sóng đẩy bùn cát phủ kín thân để ngụy trang.",
      "Tư duy thiết kế phỏng sinh học ứng dụng trong kỹ thuật ngụy trang."],
     "Lắp ráp Cá bơn ngụy trang với vây diềm lượn sóng", "cơ cấu gợn sóng bề mặt đáy", "Bậc 2"),
    (14, "Bài 14", "MẬT MÃ DISNEY", "Chuột Mickey cơ học", "OLLO KINDER",
     ["Tỷ lệ hình học tạo hình nhân vật chú chuột Mickey nổi tiếng toàn cầu.",
      "Cấu tạo khớp tay cầm nắm đồ vật và bước chân di chuyển linh hoạt.",
      "Kết hợp nghệ thuật tạo hình (Art) với tư duy kỹ thuật lắp ráp STEM."],
     "Lắp ráp Chuột Mickey thông minh với các khớp cử động", "tạo hình tỷ lệ vàng và khớp chuyển động", "Bậc 2"),
    (15, "Bài 15", "BÍ MẬT CỦA BỘ LÔNG TRẮNG!", "Chú cừu thông minh", "OLLO KINDER",
     ["Công nghệ dệt len tự nhiên và cấu trúc bộ lông cách nhiệt của loài cừu.",
      "Khung thân bo tròn khí động học kết hợp hệ thống bánh xe giảm chấn.",
      "Rèn luyện kỹ năng kết nối các tấm ốp vỏ tăng độ cứng cáp cho robot."],
     "Lắp ráp Chú cừu thông minh với thân bo tròn và bánh xe giảm chấn", "tấm ốp chịu lực và giảm chấn bánh xe", "Bậc 2"),
    (16, "Bài 16", "ÂM THANH TỪ THẾ GIỚI HOẠT HÌNH", "Hộp âm thanh cơ học", "OLLO KINDER",
     ["Nguyên lý phát thanh cơ học: gai gõ trên trống quay gảy thanh thép tạo nốt nhạc.",
      "Cơ chế bánh răng ăn khớp truyền chuyển động quay từ động cơ sang tang trống gai.",
      "Tích hợp liên môn Âm nhạc và Kỹ thuật công nghệ phát triển cảm thụ nghệ thuật."],
     "Lắp ráp Hộp âm thanh cơ học phát giai điệu vui tai", "tang trống gai gõ nốt nhạc và tỷ số truyền bánh răng", "Bậc 2"),
    (17, "Buổi 17", "ÔN TẬP VÀ TRIỂN LÃM ROBOT SÁNG TẠO HỌC KÌ 1", "Dự án Kinder HK1", "OLLO KINDER",
     ["Hệ thống hóa toàn bộ các cơ cấu cơ học: đòn bẩy, ròng rọc, bánh răng, trục lệch tâm.",
      "Thiết kế và chế tạo dự án robot giải quyết bài toán thực tế theo nhóm.",
      "Đánh giá định kỳ 2: Báo cáo dự án cơ cấu máy OLLO Kinder và kiểm tra toàn diện."],
     "Thiết kế và bảo vệ dự án cơ cấu máy OLLO Kinder", "tổng hợp cơ cấu cơ khí và thuyết trình giải pháp", "Bậc 2"),

    # HK2: Đột phá Excel Level 1 (18-35)
    (18, "Bài 1", "ĐỘNG CƠ LÀ GÌ? (OLLO EXCEL)", "Xe nâng Dynamixel", "OLLO EXCEL (LEVEL 1)",
     ["Làm quen bộ thiết bị cao cấp OLLO Excel; tìm hiểu động cơ thông minh Dynamixel.",
      "Quy tắc cắm chốt chuyên dụng, căn chỉnh vạch chuẩn góc 0 độ của động cơ Dynamixel.",
      "Thực hành lắp ráp Xe nâng hàng tự động vận hành bằng động cơ điện Dynamixel."],
     "Lắp ráp Xe nâng hàng tự động bằng động cơ servo Dynamixel", "động cơ Dynamixel, chốt đặc biệt và căn góc chuẩn", "Bậc 2"),
    (19, "Bài 2", "ROBOT CẦM NẮM VÀ CỐ ĐỊNH VẬT THỂ", "Robot gắp hàng", "OLLO EXCEL (LEVEL 1)",
     ["Cơ cấu ngàm kẹp cơ khí tự động, góc kẹp tối ưu và lực ma sát giữ đồ vật.",
      "Cánh tay kẹp đóng mở điều khiển bởi động cơ Dynamixel công suất lớn.",
      "Thử nghiệm gắp và vận chuyển các vật thể kích thước, khối lượng khác nhau."],
     "Lắp ráp Robot gắp hàng tự động có ngàm kẹp linh hoạt", "ngàm kẹp ma sát và điều khiển góc kẹp Dynamixel", "Bậc 2"),
    (20, "Bài 3", "ROBOT NHẬN BIẾT ÂM THANH", "Robot cảm biến âm thanh", "OLLO EXCEL (LEVEL 1)",
     ["Nguyên lý cảm biến âm thanh (Sound Sensor) thu nhận sóng âm biến đổi thành tín hiệu.",
      "Lắp ráp Robot vỗ tay thông minh di chuyển hoặc dừng lại khi nghe tiếng vỗ tay.",
      "Lập trình cấu hình độ nhạy cảm biến âm thanh phù hợp môi trường."],
     "Lắp ráp Robot phản ứng với tiếng vỗ tay bằng cảm biến âm thanh", "cảm biến âm thanh Micro và điều kiện ngưỡng âm", "Bậc 2"),
    (21, "Bài 4", "BỘ ĐIỀU KHIỂN TỪ XA RC-100", "Pháo cao xạ RC-100", "OLLO EXCEL (LEVEL 1)",
     ["Công nghệ truyền tín hiệu không dây hồng ngoại giữa tay cầm RC-100 và robot.",
      "Lắp ráp Pháo cao xạ tầm xa quay nòng 360 độ và kích hoạt bắn đạn nhựa an toàn.",
      "Rèn luyện kỹ năng điều khiển chính xác robot vượt chướng ngại vật."],
     "Lắp ráp Pháo cao xạ điều khiển từ xa bằng RC-100", "giao thức điều khiển không dây RC-100 và nòng xoay 360", "Bậc 2"),
    (22, "Bài 5", "DYNAMIXEL HOẠT ĐỘNG NHƯ THẾ NÀO?", "Robot khuếch đại lực", "OLLO EXCEL (LEVEL 1)",
     ["Phân tích sâu nguyên lý động cơ servo Dynamixel: mạch điều khiển và hộp số giảm tốc.",
      "Lắp ráp mô hình Robot khuếch đại chuyển động kiểm chứng tốc độ và mômen xoắn.",
      "Tính toán tỷ số truyền động bánh răng và sự đánh đổi giữa tốc độ và lực nâng."],
     "Lắp ráp Robot khuếch đại chuyển động kiểm chứng mômen xoắn", "hộp số giảm tốc servo và tỷ số truyền bánh răng", "Bậc 2"),
    (23, "Bài 6", "PHƯƠNG TIỆN GIAO THÔNG QUA CÁC THỜI ĐẠI", "Ngựa đua tự động", "OLLO EXCEL (LEVEL 1)",
     ["Lịch sử tiến hóa phương tiện: từ cỗ xe ngựa cổ đại đến xe đua siêu tốc.",
      "Lắp ráp mô hình Cỗ xe ngựa đua tự động với cơ cấu bốn chân phi nước đại liên động.",
      "Tổ chức giải đua xe địa hình tính thời gian giữa các đội."],
     "Lắp ráp Cỗ xe ngựa đua tự động bốn chân phi nước đại", "liên động bốn chân phi nước đại và cân bằng động", "Bậc 2"),
    (24, "Bài 7", "CƠ CHẾ TĂNG GIẢM CHIỀU DÀI TỰ ĐỘNG", "Cánh tay xếp quả trám", "OLLO EXCEL (LEVEL 1)",
     ["Cơ cấu xếp hình quả trám (Scissors Mechanism) mở rộng tầm với trong robot gỡ bom.",
      "Cánh tay robot co giãn tự động vươn dài gắp vật thể ở cự ly xa hiểm trở.",
      "Tư duy thiết kế robot phục vụ cứu hộ cứu nạn và an ninh."],
     "Lắp ráp Cánh tay robot co giãn tự động vươn xa gắp vật", "cơ cấu quả trám Lazy Tong và độ cứng vững khi vươn xa", "Bậc 2"),
    (25, "Bài 8", "ROBOT NHẬN BIẾT VẬT THỂ BẰNG CÁCH NÀO?", "Xe tự hành né vật cản", "OLLO EXCEL (LEVEL 1)",
     ["Nguyên lý phát xạ và thu nhận tia hồng ngoại (IR Sensor) đo khoảng cách vật cản.",
      "Lắp ráp Xe robot tự hành tự phát hiện bức tường và tự quay đầu tránh va chạm.",
      "Vận dụng tư duy thuật toán rẽ nhánh: IF (vật cản) THEN (rẽ hướng khác)."],
     "Lắp ráp Xe robot tự hành tránh vật cản bằng cảm biến hồng ngoại", "cảm biến hồng ngoại IR và thuật toán tránh va chạm", "Bậc 2"),
    (26, "Bài 9", "SỐ NGẪU NHIÊN VÀ THUẬT TOÁN ĐƠN GIẢN", "Robot mê cung Random", "OLLO EXCEL (LEVEL 1)",
     ["Khái niệm số ngẫu nhiên (Random number) trong khoa học máy tính và thuật toán.",
      "Lắp ráp Robot thám hiểm mê cung có hành vi ngẫu nhiên mô phỏng động vật hoang dã.",
      "Rèn luyện năng lực tư duy máy tính (Computational Thinking)."],
     "Lắp ráp Robot thám hiểm mê cung bằng thuật toán ngẫu nhiên", "thuật toán số ngẫu nhiên và phản xạ linh hoạt", "Bậc 2"),
    (27, "Bài 10", "ROBOT HÚT BỤI THÔNG MINH", "Robot hút bụi tự động", "OLLO EXCEL (LEVEL 1)",
     ["Phân tích công nghệ robot dịch vụ: động cơ quét bụi, quạt hút và cảm biến va chạm.",
      "Lắp ráp Robot hút bụi tự động OLLO Excel tự làm sạch sàn nhà không đụng chân bàn.",
      "Thử nghiệm hiệu suất quét dọn và tối ưu hóa diện tích bao phủ mặt sàn."],
     "Lắp ráp Robot hút bụi tự động tích hợp chổi quét và cảm biến", "công nghệ robot dịch vụ gia đình và đường đi zíc zắc", "Bậc 2"),
    (28, "Bài 11", "LỰC HẤP DẪN TRONG ROBOTICS", "Robot cưỡi ngựa Rodeo", "OLLO EXCEL (LEVEL 1)",
     ["Định luật vạn vật hấp dẫn, trọng lực tác dụng lên các bộ phận và phản lực tiếp xúc.",
      "Lắp ráp Robot đấu sĩ cưỡi ngựa Rodeo giữ thăng bằng trên lưng ngựa dao động mạnh.",
      "Đánh giá định kỳ 3: Kiểm tra kỹ năng tích hợp cảm biến, Dynamixel và điều khiển RC-100."],
     "Lắp ráp Robot cưỡi ngựa Rodeo kiểm chứng trọng lực thăng bằng", "con lắc thăng bằng cơ khí và kiểm soát dao động", "Bậc 2"),
    (29, "Bài 12", "DOMINO TRONG ROBOTICS", "Xe rải Domino", "OLLO EXCEL (LEVEL 1)",
     ["Cơ cấu cấp phôi tự động (Dispenser) và khoảng cách thả tối ưu tạo chuỗi đổ Domino.",
      "Lắp ráp Xe robot tự động rải quân cờ Domino cách đều nhau trên sàn khi di chuyển.",
      "Thực hiện thử thách xếp trận địa Domino dài nhất và kích hoạt đổ hoàn hảo."],
     "Lắp ráp Xe robot tự động rải quân cờ Domino thẳng tắp", "cơ cấu cấp phôi cơ học tuần tự và khoảng cách thả", "Bậc 2"),
    (30, "Bài 13", "ROBOT DỌN RÁC THÔNG MINH", "Robot 4WD dọn rác", "OLLO EXCEL (LEVEL 1)",
     ["Hệ thống dẫn động bốn bánh độc lập (4WD), vi sai chuyển hướng và chổi cuốn rác ngược chiều.",
      "Lắp ráp Robot dọn rác công nghệ cao điều khiển bằng RC-100 vượt chướng ngại vật.",
      "Kết hợp thao tác điều khiển tay cầm và phán đoán địa hình sa bàn."],
     "Lắp ráp Robot dọn rác 4 bánh dẫn động toàn phần điều khiển từ xa", "dẫn động 4WD độc lập và cơ cấu chổi cuốn đảo chiều", "Bậc 2"),
    (31, "Bài 14", "ROBOT CHIẾN BINH VÀ VŨ KHÍ TỰ VỆ", "Robot giác đấu", "OLLO EXCEL (LEVEL 1)",
     ["Chiến thuật chế tạo robot giác đấu: hạ thấp trọng tâm, mũi nêm lật và búa đập.",
      "Lắp ráp Robot chiến binh giác đấu OLLO Excel khung sườn gia cố chắc chắn.",
      "Tư duy chiến thuật, phản xạ nhanh và sự tôn trọng đối thủ trong thi đấu."],
     "Lắp ráp Robot chiến binh giác đấu với giáp chắn và búa đập", "tấm ốp mũi nêm lật đối thủ và tối ưu hóa trọng tâm thấp", "Bậc 2"),
    (32, "Bài 15", "ĐẤU TRƯỜNG ROBOT ĐÁ BÓNG / THỂ THAO", "Robot cầu thủ bóng đá", "OLLO EXCEL (LEVEL 1)",
     ["Luật thi đấu Robot Soccer quốc tế, cấu tạo càng dẫn bóng và chân sút uy lực.",
      "Lắp ráp Robot cầu thủ bóng đá điều khiển bằng remote RC-100 hai kênh độc lập.",
      "Tổ chức giải bóng đá Robot mini 2 vs 2 rèn luyện khả năng phối hợp chiến thuật nhóm."],
     "Lắp ráp Robot cầu thủ bóng đá với cánh gạt sút bóng tốc độ cao", "cơ cấu dẫn và sút bóng cơ học hai kênh RC-100", "Bậc 2"),
    (33, "Bài 16", "ĐẤU TRƯỜNG ROBOT SUMO & GIẢI CỨU MỤC TIÊU", "Robot Sumo", "OLLO EXCEL (LEVEL 1)",
     ["Luật thi đấu Sumo Robot: hất đối thủ ra khỏi võ đài Dohyo và giữ thăng bằng an toàn.",
      "Chế tạo Robot Sumo công suất cực đại: tối ưu hóa lực kéo bánh xe và lưỡi xúc càn quét.",
      "Tranh tài giải đấu Sumo Robot vòng bảng và bán kết nghẹt thở giữa các đấu thủ."],
     "Lắp ráp Robot Sumo công suất lớn tranh tài võ đài Dohyo", "tối ưu hóa lực kéo bám đường và chiến thuật càn quét", "Bậc 2"),
    (34, "Buổi 34", "ÔN TẬP VÀ THIẾT KẾ DỰ ÁN ROBOT CUỐI NĂM", "Dự án công nghệ Excel", "OLLO EXCEL (LEVEL 1)",
     ["Tổng kết toàn diện hành trình làm chủ công nghệ OLLO Excel: cơ khí, điện tử, cảm biến và điều khiển.",
      "Thiết kế dự án sáng tạo cuối năm theo đề tài: Thành phố thông minh, Cứu hộ khẩn cấp.",
      "Đánh giá định kỳ 4: Thẩm định hồ sơ thiết kế kỹ thuật, sản phẩm thực tế và thuyết trình bảo vệ."],
     "Thiết kế và lập hồ sơ kỹ thuật dự án công nghệ OLLO Excel cuối năm", "tích hợp đa hệ thống cảm biến, động cơ và điều khiển", "Bậc 2"),
    (35, "Buổi 35", "NGÀY HỘI ROBOTICS UNIGO VÀ BẾ MẠC CLB", "Chung kết UNIGO Cup", "OLLO EXCEL (LEVEL 1)",
     ["Chung kết Đấu trường ROBOTICS UNIGO CUP 2026 - 2027: Tranh tài Sumo, Bóng đá và Thuyết trình dự án.",
      "Lễ bế mạc tổng kết một năm học tập, nghiên cứu và bùng nổ đam mê công nghệ cùng CLB Robotics.",
      "Trao cúp vô địch, huy chương danh dự, chứng chỉ tốt nghiệp CLB và khen thưởng cá nhân xuất sắc."],
     "Thi đấu Chung kết Robotics UNIGO Cup và tham dự Lễ bế mạc CLB", "tranh tài đấu trường Robot Cup và nhận cúp chứng nhận", "Bậc 2")
]

def build_lesson(item, group_type):
    # group_type: "lop_1_2" hoặc "lop_3_7"
    b_num, b_code, t_bai, mo_ta, kit, yccd_kt, hd_build_title, co_cau_chinh, bac_str = item
    ngay_day, ngay_soan = CALENDAR[b_num - 1]

    if group_type == "lop_1_2":
        g_name = "CLB Robotics (Khối 1 - 2)"
        to_cm = "Tổ chuyên môn Tiểu học"
        dd_hs = f"Bộ Kit Robotics {kit} (theo nhóm 2-3 HS), khay phân loại linh kiện thông minh, dụng cụ tháo chốt kẹp an toàn."
        pc_cd = "Tích cực tìm hiểu bài học, kiên trì và tỉ mỉ trong từng thao tác kết nối chốt ghép mô hình. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
        pc_tn = f"Giữ gìn cẩn thận bộ linh kiện {kit}, phân loại đúng vị trí và dọn dẹp sạch sẽ không gian thực hành sau buổi học. (Đạt được thông qua Hoạt động 4)"
        pc_tt = "Tự giác thực hành, trung thực nhận xét và đánh giá sản phẩm của nhóm mình và nhóm bạn. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        nls_items = [
            ("Miền I. Thông tin và dữ liệu", "1.1. Duyệt, tìm kiếm và lọc dữ liệu, thông tin", "1",
             f"Quan sát, nhận diện chính xác các linh kiện kỹ thuật và sơ đồ hướng dẫn lắp ráp mô hình {t_bai}.", "Hoạt động 2"),
            ("Miền V. Giải quyết vấn đề", "5.3. Sử dụng sáng tạo công nghệ số", "1",
             f"Vận dụng các chi tiết cơ học thông minh để lắp ráp và vận hành hoàn thiện mô hình {t_bai}.", "Hoạt động 3, Hoạt động 4")
        ]
        nl_dt_items = [
            ("NL1", "Nhận thức công nghệ", f"Nhận diện đúng các linh kiện đặc trưng, khớp nối và cấu tạo mô hình {t_bai}.", "Hoạt động 2"),
            ("NL3", "Sử dụng công nghệ", f"Thao tác lắp ráp đúng quy trình, chắc chắn và vận hành trơn tru mô hình {t_bai}.", "Hoạt động 3"),
            ("NL5", "Giao tiếp công nghệ", f"Tự tin mô tả các bộ phận chính, nguyên lý hoạt động và giới thiệu sản phẩm trước lớp.", "Hoạt động 3, Hoạt động 4")
        ]
    else:
        g_name = "CLB Robotics (Khối 3 - 7)"
        to_cm = "Tổ chuyên môn THCS"
        dd_hs = f"Bộ Kit Robotics {kit} đồng bộ (theo nhóm 2-3 HS), động cơ/cảm biến/remote kèm theo, khay linh kiện, thảm thử nghiệm sa bàn."
        pc_cd = "Tích cực nghiên cứu bản vẽ kỹ thuật, kiên trì thử nghiệm nhiều phương án tối ưu cơ cấu cơ học. (Đạt được thông qua Hoạt động 2, Hoạt động 3)"
        pc_tn = f"Bảo quản an toàn thiết bị điện tử, động cơ, cảm biến bộ {kit}; kiểm đếm đầy đủ trước khi kết thúc ca học. (Đạt được thông qua Hoạt động 4)"
        pc_tt = "Nghiêm túc tuân thủ luật thử nghiệm/thi đấu sa bàn, trung thực ghi nhận kết quả và tôn trọng giải pháp sáng tạo của bạn bè. (Đạt được thông qua Hoạt động 3, Hoạt động 4)"
        nls_items = [
            ("Miền I. Thông tin và dữ liệu", "1.2. Đánh giá dữ liệu, thông tin và nội dung số", "2",
             f"Phân tích bản vẽ kỹ thuật 2D/3D, đánh giá thông số kích thước và góc quay linh kiện mô hình {t_bai}.", "Hoạt động 2"),
            ("Miền V. Giải quyết vấn đề", "5.3. Sử dụng sáng tạo công nghệ số", "2",
             f"Ứng dụng sáng tạo cơ cấu truyền động, cảm biến và động cơ để giải quyết bài toán nhiệm vụ thực tế của mô hình {t_bai}.", "Hoạt động 3, Hoạt động 4")
        ]
        nl_dt_items = [
            ("NL1", "Nhận thức công nghệ", f"Phân tích cấu tạo cơ khí, nguyên lý truyền lực và cơ chế hoạt động của {co_cau_chinh}.", "Hoạt động 2"),
            ("NL3", "Sử dụng công nghệ", f"Lắp ráp chính xác, kết nối đúng cực nguồn/cảm biến và vận hành hoàn hảo mô hình {t_bai}.", "Hoạt động 3"),
            ("NL4", "Đánh giá công nghệ", f"Thử nghiệm trên sa bàn, phát hiện lỗi kẹt khớp hoặc lệch hướng và cải tiến tối ưu hóa sản phẩm.", "Hoạt động 4"),
            ("NL5", "Giao tiếp công nghệ", f"Thuyết trình rõ ràng về giải pháp kỹ thuật, phân tích ưu điểm mô hình và bảo vệ ý tưởng thiết kế.", "Hoạt động 4")
        ]

    nl_ch_items = [
        ("Tự chủ và tự học", f"Chủ động nghiên cứu sơ đồ từng bước, tự lực lựa chọn linh kiện và lắp ráp mô hình {t_bai}.", "Hoạt động 2, Hoạt động 3"),
        ("Giao tiếp và hợp tác", "Tích cực thảo luận nhóm, phân công nhiệm vụ hợp lý và phối hợp nhịp nhàng trong hoạt động thử nghiệm thi đấu.", "Hoạt động 3, Hoạt động 4")
    ]

    dd_gv = f"Bộ Kit Robotics {kit} mẫu, mô hình hoàn thiện của {t_bai}, máy chiếu/màn hình tương tác trình chiếu video minh họa và sơ đồ lắp ráp từng bước, sa bàn thử nghiệm, phiếu tiêu chí đánh giá."

    # Xây dựng 4 hoạt động chi tiết
    hd1_title = "1. Hoạt động MỞ ĐẦU / KHỞI ĐỘNG (10 phút)"
    hd1_gv_hs = (
        f"Bước 1. Chuyển giao nhiệm vụ:\n"
        f"- GV trình chiếu video/hình ảnh thực tế liên quan đến chủ đề: {t_bai} ({mo_ta}).\n"
        f"- GV đặt câu hỏi gợi mở: 'Em có nhận xét gì về cấu tạo và cách chuyển động của {mo_ta} trong đời sống? Làm thế nào để chúng ta tạo ra một robot mô phỏng chuyển động này bằng bộ thiết bị {kit}?'\n\n"
        f"Bước 2. Thực hiện nhiệm vụ: HS quan sát video, suy nghĩ và trao đổi nhanh với bạn cùng bàn.\n\n"
        f"Bước 3. Báo cáo, thảo luận: Đại diện một số HS phát biểu nhận định về đặc điểm chuyển động, hình dáng và chức năng của {mo_ta}.\n\n"
        f"Bước 4. Kết luận, nhận định: GV tổng kết ý kiến, biểu dương tinh thần quan sát của HS và dẫn dắt vào bài học chế tạo robot {t_bai}."
    )
    hd1_kq = (
        f"- HS hào hứng, tập trung cao độ và sẵn sàng tham gia bài học.\n"
        f"- Nhận biết được hình dáng thực tế, công dụng và chuyển động đặc trưng của {mo_ta}."
    )

    hd2_title = "2. Hoạt động HÌNH THÀNH KIẾN THỨC MỚI / KHÁM PHÁ (20 phút)"
    hd2_gv_hs = (
        f"Bước 1. Chuyển giao nhiệm vụ:\n"
        f"- GV giới thiệu mô hình mẫu của robot {t_bai} và cho HS quan sát chuyển động thực tế.\n"
        f"- Hướng dẫn HS nhận diện các linh kiện cốt lõi: khung thân, bánh xe/chân đế, thanh liên kết và đặc biệt là cơ cấu {co_cau_chinh}.\n"
        f"- Trình chiếu bản vẽ phân rã linh kiện trên màn hình tương tác, hướng dẫn quy tắc ghép chốt an toàn và căn chỉnh khớp quay.\n\n"
        f"Bước 2. Thực hiện nhiệm vụ: Các nhóm HS mở hộp Kit {kit}, đối chiếu bản vẽ và tìm chọn các chi tiết chính ra khay phân loại.\n\n"
        f"Bước 3. Báo cáo, thảo luận: Đại diện nhóm gọi tên linh kiện, giải thích vai trò của cơ cấu {co_cau_chinh} trong mô hình.\n\n"
        f"Bước 4. Kết luận, nhận định: GV chuẩn hóa kiến thức về nguyên lý cơ học của bài học, nhấn mạnh các lưu ý kỹ thuật để tránh lắp ngược chiều hoặc kẹt khớp."
    )
    hd2_kq = (
        f"- Nắm vững tên gọi, hình dạng và công năng của các linh kiện trong bộ {kit}.\n"
        f"- Hiểu rõ nguyên lý hoạt động của cơ cấu {co_cau_chinh} trên mô hình {t_bai}."
    )

    hd3_title = "3. Hoạt động LUYỆN TẬP / THỰC HÀNH LẮP RÁP ROBOT (40 phút)"
    hd3_gv_hs = (
        f"Bước 1. Chuyển giao nhiệm vụ:\n"
        f"- GV giao nhiệm vụ cho từng nhóm: Thực hành lắp ráp hoàn chỉnh robot {t_bai} theo sơ đồ hướng dẫn từng bước (từ khung gầm -> thân máy -> cơ cấu chuyển động -> trang trí hoàn thiện).\n"
        f"- Nêu rõ các tiêu chí sản phẩm: Lắp đúng bản vẽ, kết cấu chắc chắn không lỏng lẻo, cơ cấu {co_cau_chinh} cử động êm ái, thẩm mỹ đẹp.\n\n"
        f"Bước 2. Thực hiện nhiệm vụ:\n"
        f"- HS làm việc theo nhóm 2-3 bạn: một bạn đọc bản vẽ và chọn linh kiện, một bạn thực hiện thao tác lắp ráp chốt, sau đó đổi vai trò.\n"
        f"- GV đi từng bàn quan sát, kịp thời nhắc nhở tư thế cầm dụng cụ tháo chốt, hỗ trợ các thao tác cắm chốt chặt hoặc căn chỉnh khớp quay cho các nhóm.\n\n"
        f"Bước 3. Báo cáo, thảo luận:\n"
        f"- Các nhóm hoàn thành tiến hành thử quay trục bằng tay hoặc bật công tắc nguồn để kiểm tra chuyển động ban đầu.\n"
        f"- Trao đổi nhanh trong nhóm nếu phát hiện bánh răng bị kênh hoặc khớp nối bị cọ xát.\n\n"
        f"Bước 4. Kết luận, nhận định:\n"
        f"- GV nhận xét tiến độ thực hành của các nhóm, tuyên dương các nhóm thao tác nhanh, chuẩn xác và phối hợp ăn ý."
    )
    hd3_kq = (
        f"- 100% các nhóm học sinh lắp ráp hoàn thiện mô hình robot {t_bai}.\n"
        f"- Mô hình có kết cấu kiên cố, các khớp chuyển động đúng theo thiết kế, không bị kẹt hay bung chốt."
    )

    hd4_title = "4. Hoạt động VẬN DỤNG, THỬ NGHIỆM & SÁNG TẠO (20 phút)"
    hd4_gv_hs = (
        f"Bước 1. Chuyển giao nhiệm vụ:\n"
        f"- GV tổ chức thử thách thử nghiệm / mini game trên sa bàn: Cho robot {t_bai} thực hiện nhiệm vụ vận hành (di chuyển vượt chướng ngại vật, thử nghiệm tốc độ, độ thăng bằng hoặc độ bền cơ học).\n"
        f"- Khuyến khích HS sáng tạo: Thử gắn thêm chi tiết phụ trợ, thay đổi vị trí chốt để cải tiến chuyển động hoặc làm cho robot đẹp hơn.\n\n"
        f"Bước 2. Thực hiện nhiệm vụ:\n"
        f"- Từng nhóm mang robot lên sa bàn kiểm thử, quan sát robot vận hành thực tế và ghi nhận kết quả.\n"
        f"- Các nhóm thảo luận cách tối ưu hóa: siết chặt chốt, bôi trơn tiếp điểm hoặc gia cố thanh chống lật.\n\n"
        f"Bước 3. Báo cáo, thảo luận:\n"
        f"- Đại diện các nhóm giới thiệu robot của mình, chia sẻ những điểm sáng tạo độc đáo hoặc cách nhóm đã sửa lỗi cơ học.\n"
        f"- Các nhóm bạn quan sát và chấm điểm đồng đẳng theo các tiêu chí đã thống nhất.\n\n"
        f"Bước 4. Kết luận, nhận định:\n"
        f"- GV đánh giá, chấm điểm và tổng kết buổi học. Khen ngợi tinh thần sáng tạo và nỗ lực của học sinh.\n"
        f"- Hướng dẫn học sinh quy trình tháo dỡ robot an toàn (hoặc bảo quản mô hình dự án), kiểm đếm đầy đủ số lượng linh kiện vào khay và xếp gọn hộp Kit {kit} vào tủ đồ dùng."
    )
    hd4_kq = (
        f"- Robot vận hành xuất sắc nhiệm vụ trên sa bàn thử nghiệm.\n"
        f"- HS rèn luyện kỹ năng giải quyết vấn đề, tự tin chia sẻ giải pháp kỹ thuật trước tập thể.\n"
        f"- Không gian lớp học sạch sẽ, toàn bộ linh kiện được kiểm đếm ngăn nắp, an toàn tuyệt đối."
    )

    return (
        b_num, ngay_day, ngay_soan, g_name, to_cm, kit, b_code, t_bai,
        yccd_kt, pc_cd, pc_tn, pc_tt, nl_dt_items, nls_items, nl_ch_items,
        dd_gv, dd_hs,
        hd1_title, hd1_gv_hs, hd1_kq,
        hd2_title, hd2_gv_hs, hd2_kq,
        hd3_title, hd3_gv_hs, hd3_kq,
        hd4_title, hd4_gv_hs, hd4_kq
    )

def main():
    print("Starting generation of 70 KHBD documents...")

    # Dọn dẹp các file cũ trong Tuần_01 đến Tuần_35
    for w in range(1, 36):
        w_dir = os.path.join(BASE_KHBD_DIR, f"Tuần_{w:02d}")
        os.makedirs(w_dir, exist_ok=True)
        # Xóa các file cũ
        for old_f in os.listdir(w_dir):
            if old_f.endswith(".docx"):
                try:
                    os.remove(os.path.join(w_dir, old_f))
                except Exception as e:
                    pass

    # Sinh 35 bài Track 1 (Lớp 1 - 2)
    print("\n--- Generating Track 1: Khối 1 - 2 (35 sessions) ---")
    for item in RAW_TRACK_1:
        args = build_lesson(item, "lop_1_2")
        buoi_num = args[0]
        bai_so = args[6]
        ten_bai = args[7]
        
        filename = f"KHBD_CLB_Robotics_Lop1_2_Buoi{buoi_num:02d}_{sanitize_filename(bai_so)}_{sanitize_filename(ten_bai)}.docx"
        
        # Đường dẫn trong thư mục Tuần tương ứng
        w_dir = os.path.join(BASE_KHBD_DIR, f"Tuần_{buoi_num:02d}")
        out_week_path = os.path.join(w_dir, filename)
        
        # Đường dẫn trong thư mục chuyên phân hệ Lớp 1 - 2
        out_group_path = os.path.join(DIR_LOP_1_2, filename)
        
        create_khbd_doc(*args, out_filepath=out_week_path)
        shutil.copy2(out_week_path, out_group_path)
        print(f"Created: Tuần_{buoi_num:02d} -> {filename}")

    # Sinh 35 bài Track 2 (Lớp 3 - 7)
    print("\n--- Generating Track 2: Khối 3 - 7 (35 sessions) ---")
    for item in RAW_TRACK_2:
        args = build_lesson(item, "lop_3_7")
        buoi_num = args[0]
        bai_so = args[6]
        ten_bai = args[7]
        
        filename = f"KHBD_CLB_Robotics_Lop3_7_Buoi{buoi_num:02d}_{sanitize_filename(bai_so)}_{sanitize_filename(ten_bai)}.docx"
        
        # Đường dẫn trong thư mục Tuần tương ứng
        w_dir = os.path.join(BASE_KHBD_DIR, f"Tuần_{buoi_num:02d}")
        out_week_path = os.path.join(w_dir, filename)
        
        # Đường dẫn trong thư mục chuyên phân hệ Lớp 3 - 7
        out_group_path = os.path.join(DIR_LOP_3_7, filename)
        
        create_khbd_doc(*args, out_filepath=out_week_path)
        shutil.copy2(out_week_path, out_group_path)
        print(f"Created: Tuần_{buoi_num:02d} -> {filename}")

    print("\nALL 70 KHBD DOCUMENTS SUCCESSFULLY GENERATED AND ORGANIZED!")

if __name__ == "__main__":
    main()
