/**
 * UNIGO — Dữ liệu Năng lực & Phẩm chất chuẩn Văn bản pháp quy
 * Phổ quát cho tất cả các môn học (Toán, Văn, Anh, KHTN, Lịch sử - Địa lý, Tin học, Robotics...)
 * 1. CV 3456/BGDĐT-CNTT (Thông tư 02/2025/TT-BGDĐT) — Khung Năng lực số
 * 2. Quyết định 3439/QĐ-BGDĐT — Khung Năng lực Trí tuệ nhân tạo (AI)
 * 3. Công văn 5512/BGDĐT-GDTrH & CT GDPT 2018 — Năng lực chung & 5 Phẩm chất
 * 4. CT GDPT 2018 môn Tin học (TT 32/2018/TT-BGDĐT) & Khung Robotics UNIGO
 */

// =============================================================================
// I. CÔNG VĂN 3456/BGDĐT — KHUNG NĂNG LỰC SỐ (6 MIỀN & 24 THÀNH TỐ)
// =============================================================================
const NLS_MIEN_DEFINITIONS = {
    1: {
        roman: "I",
        name: "Khai thác dữ liệu và thông tin",
        concept: "Xác định nhu cầu thông tin; tìm kiếm, truy cập dữ liệu, thông tin và nội dung trong môi trường số; đánh giá độ tin cậy và tính xác thực; lưu trữ, quản lý và tổ chức dữ liệu một cách có cấu trúc.",
        thanhToCodes: ["1.1", "1.2", "1.3"]
    },
    2: {
        roman: "II",
        name: "Giao tiếp và hợp tác trong môi trường số",
        concept: "Tương tác, giao tiếp và cộng tác thông qua các công nghệ số; nhận biết phương tiện truyền thông phù hợp; chia sẻ dữ liệu và nội dung; thực hiện trách nhiệm công dân; tuân thủ quy tắc ứng xử trên mạng (Netiquette) và quản lý danh tính số.",
        thanhToCodes: ["2.1", "2.2", "2.3", "2.4", "2.5", "2.6"]
    },
    3: {
        roman: "III",
        name: "Sáng tạo nội dung số",
        concept: "Tạo mới, chỉnh sửa và tích hợp dữ liệu, thông tin thành nội dung số dưới các định dạng khác nhau; hiểu biết và thực thi bản quyền, sở hữu trí tuệ; hiểu các nguyên lý của thuật toán và lập trình máy tính.",
        thanhToCodes: ["3.1", "3.2", "3.3", "3.4"]
    },
    4: {
        roman: "IV",
        name: "An toàn",
        concept: "Bảo vệ thiết bị, mạng và nội dung số; bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số; bảo vệ sức khỏe thể chất và an sinh số; nhận thức về tác động của công nghệ số đối với môi trường tự nhiên.",
        thanhToCodes: ["4.1", "4.2", "4.3", "4.4"]
    },
    5: {
        roman: "V",
        name: "Giải quyết vấn đề",
        concept: "Xác định các vấn đề kỹ thuật khi vận hành thiết bị số và giải quyết chúng; đánh giá nhu cầu và lựa chọn công cụ/giải pháp công nghệ phù hợp; sử dụng sáng tạo công nghệ số để tạo ra sản phẩm; xác định các khoảng trống năng lực số của bản thân để tự cải thiện.",
        thanhToCodes: ["5.1", "5.2", "5.3", "5.4"]
    },
    6: {
        roman: "VI",
        name: "Ứng dụng trí tuệ nhân tạo (AI)",
        concept: "Hiểu biết về khái niệm, nguyên lý hoạt động và giới hạn của AI; sử dụng các công cụ, ứng dụng tích hợp AI để giải quyết nhiệm vụ học tập và công việc; đánh giá tác động, độ tin cậy và khía cạnh đạo đức khi ứng dụng AI.",
        thanhToCodes: ["6.1", "6.2", "6.3"]
    }
};

const NLS_FULL_DATA = {
    "1.1": {
        code: "1.1",
        name: "Duyệt, tìm kiếm và lọc dữ liệu, thông tin và nội dung số",
        mienId: 1,
        mienName: "Khai thác dữ liệu và thông tin",
        mota: "Xác định được nhu cầu thông tin; tìm kiếm được dữ liệu, thông tin và nội dung trong môi trường số; truy cập chúng và khai thác được kết quả tìm kiếm. Tạo và cập nhật được chiến lược tìm kiếm.",
        descriptors: {
            "L1-3": "- Xác định được nhu cầu thông tin, tìm kiếm dữ liệu, thông tin và nội dung thông qua tìm kiếm đơn giản trong môi trường số,\n- Tìm được cách truy cập những dữ liệu, thông tin và nội dung này cũng như điều hướng giữa chúng,\n- Xác định được các chiến lược tìm kiếm đơn giản.",
            "L4-5": "- Xác định được nhu cầu thông tin,\n- Tìm được dữ liệu, thông tin và nội dung thông qua tìm kiếm đơn giản trong môi trường số,\n- Tìm được cách truy cập những dữ liệu, thông tin và nội dung này cũng như điều hướng giữa chúng,\n- Xác định được các chiến lược tìm kiếm đơn giản.",
            "L6-7": "- Giải thích được nhu cầu thông tin,\n- Thực hiện được rõ ràng và theo quy trình các tìm kiếm để tìm dữ liệu, thông tin và nội dung trong môi trường số,\n- Giải thích được cách truy cập và điều hướng các kết quả tìm kiếm,\n- Giải thích được rõ ràng và theo quy trình chiến lược tìm kiếm.",
            "L8-9": "- Minh họa được nhu cầu thông tin,\n- Tổ chức được tìm kiếm dữ liệu, thông tin và nội dung trong môi trường số,\n- Mô tả được cách truy cập những dữ liệu, thông tin và nội dung này cũng như điều hướng giữa chúng,\n- Tổ chức được các chiến lược tìm kiếm.",
            "L10-12": "- Đáp ứng được nhu cầu thông tin,\n- Áp dụng được kỹ thuật tìm kiếm để lấy được dữ liệu, thông tin và nội dung trong môi trường số,\n- Chỉ cho người khác cách truy cập những dữ liệu, thông tin và nội dung này cũng như điều hướng giữa chúng.\n- Tự đề xuất được chiến lược tìm kiếm."
        }
    },
    "1.2": {
        code: "1.2",
        name: "Đánh giá dữ liệu, thông tin và nội dung số",
        mienId: 1,
        mienName: "Khai thác dữ liệu và thông tin",
        mota: "Phân tích, so sánh và đánh giá được độ tin cậy và tính xác thực của nguồn dữ liệu, thông tin và nội dung số. Phân tích, giải thích và đánh giá được dữ liệu, thông tin và nội dung số.",
        descriptors: {
            "L1-3": "- Phát hiện được độ tin cậy và độ chính xác của các nguồn chung của dữ liệu, thông tin và nội dung số.",
            "L4-5": "- Phát hiện được độ tin cậy và độ chính xác của các nguồn chung của dữ liệu, thông tin và nội dung số.",
            "L6-7": "- Thực hiện phân tích, so sánh, đánh giá được độ tin cậy và độ chính xác của các nguồn dữ liệu, thông tin và nội dung số đã được tổ chức rõ ràng.\n- Thực hiện phân tích, diễn giải và đánh giá được dữ liệu, thông tin và nội dung số được xác định rõ ràng.",
            "L8-9": "- Thực hiện phân tích, so sánh và đánh giá được các nguồn dữ liệu, thông tin và nội dung số.\n- Thực hiện phân tích, diễn giải và đánh giá được dữ liệu, thông tin và nội dung số.",
            "L10-12": "- Thực hiện đánh giá được độ tin cậy và độ tin cậy của các nguồn dữ liệu, thông tin và nội dung số.\n- Tiến hành đánh giá được các dữ liệu, thông tin và nội dung số khác nhau."
        }
    },
    "1.3": {
        code: "1.3",
        name: "Quản lý dữ liệu, thông tin và nội dung số",
        mienId: 1,
        mienName: "Khai thác dữ liệu và thông tin",
        mota: "Tổ chức, lưu trữ và truy xuất được dữ liệu, thông tin và nội dung trong môi trường số. Tổ chức và sắp xếp được chúng trong một môi trường có cấu trúc.",
        descriptors: {
            "L1-3": "- Xác định được cách tổ chức, lưu trữ và truy xuất dữ liệu, thông tin và nội dung một cách đơn giản trong môi trường số.\n- Nhận biết được nơi để sắp xếp dữ liệu, thông tin và nội dung một cách đơn giản trong môi trường có cấu trúc.",
            "L4-5": "- Xác định được cách tổ chức, lưu trữ và truy xuất dữ liệu, thông tin và nội dung một cách đơn giản trong môi trường số.\n- Nhận biết được nơi để sắp xếp dữ liệu, thông tin và nội dung một cách đơn giản trong môi trường có cấu trúc.",
            "L6-7": "- Lựa chọn được dữ liệu, thông tin và nội dung để tổ chức, lưu trữ và truy xuất chúng một cách thường xuyên trong môi trường số.\n- Sắp xếp chúng một cách trật tự trong một môi trường có cấu trúc.",
            "L8-9": "- Sắp xếp được thông tin, dữ liệu, nội dung để dễ dàng lưu trữ và truy xuất.\n- Tổ chức được thông tin, dữ liệu và nội dung trong một môi trường có cấu trúc.",
            "L10-12": "- Thao tác được thông tin, dữ liệu và nội dung để tổ chức, lưu trữ và truy xuất dễ dàng hơn.\n- Triển khai được việc tổ chức và sắp xếp dữ liệu, thông tin và nội dung trong môi trường có cấu trúc."
        }
    },
    "2.1": {
        code: "2.1",
        name: "Tương tác thông qua công nghệ số",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Tương tác thông qua các công nghệ số khác nhau và nhận biết được phương tiện giao tiếp số nào phù hợp cho một bối cảnh nhất định.",
        descriptors: {
            "L1-3": "- Lựa chọn được các công nghệ số đơn giản để tương tác.\n- Xác định được các phương tiện giao tiếp đơn giản thích hợp cho một bối cảnh cụ thể.",
            "L4-5": "- Lựa chọn được các công nghệ số đơn giản để tương tác.\n- Xác định được các phương tiện giao tiếp đơn giản thích hợp cho một bối cảnh cụ thể.",
            "L6-7": "- Thực hiện được các tương tác được xác định rõ ràng và thường xuyên với các công nghệ số.\n- Lựa chọn được các phương tiện giao tiếp số phù hợp, được xác định rõ ràng cho phù hợp với bối cảnh nhất định.",
            "L8-9": "- Lựa chọn được nhiều công nghệ số để tương tác.\n- Lựa chọn được nhiều phương tiện truyền thông số cho phù hợp với bối cảnh nhất định.",
            "L10-12": "- Sử dụng được nhiều công nghệ số để tương tác.\n- Cho người khác thấy phương tiện giao tiếp số phù hợp nhất cho một bối cảnh cụ thể."
        }
    },
    "2.2": {
        code: "2.2",
        name: "Chia sẻ thông tin và nội dung thông qua công nghệ số",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Chia sẻ dữ liệu, thông tin và nội dung số với người khác thông qua các công nghệ số phù hợp. Đóng vai trò là người trung gian, hiểu biết và thực hành trích dẫn và ghi chú nguồn.",
        descriptors: {
            "L1-3": "- Nhận biết được các công nghệ số đơn giản, phù hợp để chia sẻ dữ liệu, thông tin và nội dung kỹ thuật số.\n- Nhận biết được phương pháp trích dẫn và ghi nguồn cơ bản.",
            "L4-5": "- Nhận biết được các công nghệ số đơn giản, phù hợp để chia sẻ dữ liệu, thông tin và nội dung kỹ thuật số.\n- Xác định được phương pháp trích dẫn và ghi nguồn cơ bản.",
            "L6-7": "- Lựa chọn các công nghệ số phù hợp được xác định rõ để trao đổi dữ liệu, thông tin và nội dung số.\n- Giải thích cách thức hoạt động như một trung gian để chia sẻ thông tin và nội dung thông qua các công nghệ kỹ thuật số được xác định rõ ràng và thường xuyên,\n- Minh họa rõ ràng và thường xuyên các phương pháp tham chiếu và ghi chú nguồn.",
            "L8-9": "- Vận dụng được các công nghệ số phù hợp để chia sẻ dữ liệu, thông tin và nội dung số.\n- Giải thích được cách đóng vai trò trung gian để chia sẻ thông tin và nội dung thông qua công nghệ số.\n- Áp dụng được các phương pháp tham chiếu và ghi chú nguồn.",
            "L10-12": "- Chia sẻ dữ liệu, thông tin và nội dung số thông qua nhiều công cụ số phù hợp,\n- Hướng dẫn người khác cách đóng vai trò trung gian để chia sẻ thông tin và nội dung thông qua công nghệ số.\n- Áp dụng được nhiều phương pháp tham chiếu và ghi nguồn khác nhau."
        }
    },
    "2.3": {
        code: "2.3",
        name: "Sử dụng công nghệ số để thực hiện trách nhiệm công dân",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Tham gia vào xã hội thông qua việc sử dụng các dịch vụ số công cộng và tư nhân. Tìm kiếm được cơ hội để trao quyền và thu hút công dân thông qua các công nghệ số phù hợp.",
        descriptors: {
            "L1-3": "- Xác định được các dịch vụ số đơn giản để có thể tham gia vào xã hội.\n- Nhận biết được các công nghệ số đơn giản, phù hợp để nâng cao năng lực cho bản thân và tham gia vào xã hội với tư cách là một công dân.",
            "L4-5": "- Xác định được các dịch vụ số đơn giản để có thể tham gia vào xã hội.\n- Nhận biết được các công nghệ số đơn giản, phù hợp để nâng cao năng lực cho bản thân và tham gia vào xã hội với tư cách là một công dân.",
            "L6-7": "- Lựa chọn được các dịch vụ số được xác định rõ ràng và phổ biến để tham gia vào xã hội.\n- Xác định được các công nghệ số rõ ràng và thích hợp để tự mình trang bị và tham gia vào xã hội như một công dân.",
            "L8-9": "- Lựa chọn được các dịch vụ số để tham gia vào xã hội.\n- Thảo luận về các công nghệ số phù hợp để nâng cao năng lực của bản thân và tham gia vào xã hội với tư cách là một công dân.",
            "L10-12": "- Đề xuất được các dịch vụ số khác nhau để tham gia vào xã hội.\n- Sử dụng được các công nghệ số thích hợp để tự mình trang bị và tham gia vào xã hội như một công dân."
        }
    },
    "2.4": {
        code: "2.4",
        name: "Hợp tác thông qua công nghệ số",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Sử dụng được các công cụ và công nghệ số cho các quá trình hợp tác cũng như để cùng xây dựng và đồng sáng tạo dữ liệu, tài nguyên và kiến thức.",
        descriptors: {
            "L1-3": "- Chọn được những công cụ và công nghệ số đơn giản cho các quá trình cộng tác.",
            "L4-5": "- Chọn được những công cụ và công nghệ số đơn giản cho các quá trình cộng tác.",
            "L6-7": "- Lựa chọn được các công cụ và công nghệ số được xác định rõ ràng và thường xuyên cho các quá trình hợp tác.",
            "L8-9": "- Lựa chọn được các công cụ và công nghệ số cho các quá trình hợp tác.",
            "L10-12": "- Đề xuất được các công cụ và công nghệ số khác nhau cho các quá trình hợp tác."
        }
    },
    "2.5": {
        code: "2.5",
        name: "Quy tắc ứng xử trên mạng (Netiquette)",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Nhận thức được các chuẩn mực hành vi và kiến thức khi sử dụng công nghệ số và tương tác trong môi trường số. Điều chỉnh các chiến lược giao tiếp phù hợp với đối tượng cụ thể và nhận thức được sự đa dạng văn hóa và thế hệ.",
        descriptors: {
            "L1-3": "- Phân biệt được các chuẩn mực hành vi đơn giản và biết cách sử dụng công nghệ số và tương tác trong môi trường số.\n- Chọn được các phương thức và chiến lược giao tiếp đơn giản phù hợp trong môi trường số.\n- Phân biệt các khía cạnh đơn giản của sự đa dạng về văn hóa và thế hệ cần được tính đến trong môi trường số.",
            "L4-5": "- Phân biệt được các chuẩn mực hành vi đơn giản và bí quyết sử dụng công nghệ số và tương tác trong môi trường số.\n- Chọn được các phương thức và chiến lược giao tiếp đơn giản phù hợp trong môi trường số.\n- Phân biệt các khía cạnh đơn giản của sự đa dạng về văn hóa và thế hệ cần được tính đến trong môi trường số.",
            "L6-7": "- Làm rõ được các chuẩn mực hành vi thường xuyên và được xác định rõ ràng cũng như bí quyết khi sử dụng công nghệ số và tương tác trong môi trường số.\n- Thể hiện được các chiến lược giao tiếp thường xuyên và xác định rõ ràng phương thức giao tiếp phù hợp trong môi trường số.\n- Mô tả các khía cạnh đa dạng về văn hóa và thế hệ được xác định rõ ràng và thông thường cần xem xét trong môi trường số.",
            "L8-9": "- Thảo luận về các chuẩn mực hành vi và cách sử dụng công nghệ số và tương tác trong môi trường số.\n- Thảo luận các chiến lược giao tiếp phù hợp trong môi trường số.\n- Thảo luận các khía cạnh đa dạng về văn hóa và thế hệ cần xem xét trong môi trường số.",
            "L10-12": "- Áp dụng được các chuẩn mực hành vi và bí quyết khác nhau khi sử dụng công nghệ số và tương tác trong môi trường số.\n- Áp dụng được các chiến lược giao tiếp khác nhau trong môi trường số một cách phù hợp.\n- Áp dụng được các khía cạnh đa dạng về văn hóa và thế hệ khác nhau để xem xét trong môi trường số."
        }
    },
    "2.6": {
        code: "2.6",
        name: "Quản lý danh tính số",
        mienId: 2,
        mienName: "Giao tiếp và hợp tác trong môi trường số",
        mota: "Tạo và quản lý được một hoặc nhiều danh tính số để bảo vệ danh tiếng của bản thân, làm việc với dữ liệu mà một người tạo ra bằng nhiều công cụ, môi trường và dịch vụ số.",
        descriptors: {
            "L1-3": "- Xác định được danh tính số.\n- Mô tả được những cách đơn giản để bảo vệ danh tiếng trực tuyến của bản thân.\n- Nhận biết được dữ liệu đơn giản do mình tạo ra thông qua các công cụ, môi trường hoặc dịch vụ số.",
            "L4-5": "- Xác định được danh tính số.\n- Mô tả được những cách đơn giản để bảo vệ danh tiếng trực tuyến của bản thân.\n- Nhận biết được dữ liệu đơn giản do mình tạo ra thông qua các công cụ, môi trường hoặc dịch vụ số.",
            "L6-7": "- Phân biệt được một loạt các danh tính số thông thường và được xác định rõ ràng.\n- Giải thích được những cách được xác định rõ ràng và thường xuyên để bảo vệ danh tính trực tuyến của bản thân.\n- Mô tả dữ liệu được xác định rõ ràng mà bạn thường xuyên thu được thông qua các công cụ, môi trường hoặc dịch vụ số.",
            "L8-9": "- Hiển thị được nhiều danh tính số cụ thể,\n- Thảo luận những cách cụ thể để bảo vệ danh tiếng trực tuyến của bản thân.\n- Thao tác dữ liệu cá nhân tạo ra thông qua các công cụ, môi trường hoặc dịch vụ số.",
            "L10-12": "- Sử dụng được nhiều danh tính số khác nhau.\n- Áp dụng được các cách khác nhau để bảo vệ danh tính trực tuyến của bản thân.\n- Sử dụng được dữ liệu tạo ra thông qua các công cụ, môi trường và một số dịch vụ số."
        }
    },
    "3.1": {
        code: "3.1",
        name: "Phát triển nội dung số",
        mienId: 3,
        mienName: "Sáng tạo nội dung số",
        mota: "Tạo và chỉnh sửa được nội dung số ở các định dạng khác nhau, thể hiện được bản thân thông qua các phương tiện kỹ thuật số.",
        descriptors: {
            "L1-3": "- Xác định được cách tạo và chỉnh sửa nội dung số ở các định dạng đơn giản.\n- Chọn được các cách thể hiện bản thân thông qua việc tạo ra các phương tiện kỹ thuật số đơn giản.",
            "L4-5": "- Xác định được cách tạo và chỉnh sửa nội dung số ở các định dạng đơn giản.\n- Chọn được các cách thể hiện bản thân thông qua việc tạo ra các phương tiện kỹ thuật số đơn giản.",
            "L6-7": "- Chỉ ra các cách tạo và chỉnh sửa nội dung ở các định dạng số thông thường và được xác định rõ ràng.\n- Thực hiện cách thể hiện bản thân thông qua việc sáng tạo các phương tiện kỹ thuật số được xác định rõ ràng và thường xuyên.",
            "L8-9": "- Tạo và chỉnh sửa được nội dung số ở các định dạng thông thường.\n- Thể hiện được bản thân thông qua việc tạo ra các phương tiện kỹ thuật số.",
            "L10-12": "- Tạo và chỉnh sửa được nội dung số ở nhiều định dạng khác nhau.\n- Thể hiện được bản thân thông qua việc tạo ra các phương tiện kỹ thuật số khác nhau."
        }
    },
    "3.2": {
        code: "3.2",
        name: "Tích hợp và tạo lập lại nội dung số",
        mienId: 3,
        mienName: "Sáng tạo nội dung số",
        mota: "Sửa đổi, tinh chỉnh, cải thiện và tích hợp được nội dung số mới vào những nội dung số hiện có để tạo ra nội dung số mới, nguyên bản và có liên quan.",
        descriptors: {
            "L1-3": "- Chọn được những cách đơn giản để sửa đổi, hoàn thiện, cải tiến và tích hợp các mục nội dung số mới vào các mục hiện có.",
            "L4-5": "- Chọn được những cách đơn giản để sửa đổi, hoàn thiện, cải tiến và tích hợp các mục nội dung số mới vào các mục hiện có.",
            "L6-7": "- Thực hiện được các thao tác sửa đổi, tinh chỉnh, cải thiện và tích hợp được xác định rõ ràng và thường xuyên các mục nội dung số mới vào những mục hiện có để tạo ra nội dung mới và nguyên bản.",
            "L8-9": "- Sửa đổi, tinh chỉnh, cải thiện và tích hợp được nội dung số mới vào những nội dung hiện có để tạo ra nội dung mới và nguyên bản.",
            "L10-12": "- Tạo ra được nội dung mới, nguyên bản và có liên quan bằng cách sửa đổi, tinh chỉnh, cải thiện và tích hợp nhiều nội dung số mới vào những nội dung hiện có."
        }
    },
    "3.3": {
        code: "3.3",
        name: "Thực thi bản quyền và giấy phép",
        mienId: 3,
        mienName: "Sáng tạo nội dung số",
        mota: "Hiểu được cách áp dụng bản quyền và giấy phép đối với dữ liệu, thông tin và nội dung số.",
        descriptors: {
            "L1-3": "- Xác định được các quy tắc đơn giản về bản quyền và giấy phép áp dụng cho dữ liệu, thông tin và nội dung số.",
            "L4-5": "- Xác định được các quy tắc đơn giản về bản quyền và giấy phép áp dụng cho dữ liệu, thông tin và nội dung số.",
            "L6-7": "- Làm rõ các quy tắc thông thường và được xác định rõ ràng về bản quyền và giấy phép áp dụng cho dữ liệu, thông tin và nội dung số.",
            "L8-9": "- Thảo luận về các quy tắc bản quyền và giấy phép áp dụng cho dữ liệu, thông tin và nội dung số.",
            "L10-12": "- Áp dụng được các quy tắc bản quyền và giấy phép khác nhau cho dữ liệu, thông tin và nội dung số."
        }
    },
    "3.4": {
        code: "3.4",
        name: "Lập trình",
        mienId: 3,
        mienName: "Sáng tạo nội dung số",
        mota: "Lập kế hoạch và phát triển được một chuỗi các hướng dẫn cho một hệ thống máy tính để giải quyết một vấn đề nhất định hoặc thực hiện một nhiệm vụ cụ thể.",
        descriptors: {
            "L1-3": "- Liệt kê được các hướng dẫn đơn giản cho một hệ thống máy tính nhằm giải quyết một vấn đề đơn giản hoặc thực hiện một nhiệm vụ đơn giản.",
            "L4-5": "- Liệt kê được các hướng dẫn đơn giản cho một hệ thống máy tính nhằm giải quyết một vấn đề đơn giản hoặc thực hiện một nhiệm vụ đơn giản.",
            "L6-7": "- Liệt kê được các hướng dẫn thông thường và được xác định rõ ràng cho một hệ thống máy tính để giải quyết các vấn đề thường gặp hoặc thực hiện một nhiệm vụ được xác định rõ ràng.",
            "L8-9": "- Lập được danh sách các hướng dẫn cho hệ thống máy tính để giải quyết vấn đề đã xác định hoặc thực hiện nhiệm vụ cụ thể.",
            "L10-12": "- Lập được kế hoạch và phát triển được chuỗi các hướng dẫn cho hệ thống máy tính nhằm giải quyết một loạt các vấn đề hoặc thực hiện một số nhiệm vụ cụ thể."
        }
    },
    "4.1": {
        code: "4.1",
        name: "Bảo vệ thiết bị",
        mienId: 4,
        mienName: "An toàn",
        mota: "Bảo vệ được các thiết bị và nội dung số, nhận thức được các rủi ro và mối đe dọa trong môi trường số, nhận biết được các biện pháp an toàn và an ninh và nhận thức đúng đắn về độ tin cậy và quyền riêng tư.",
        descriptors: {
            "L1-3": "- Xác định được các cách đơn giản để bảo vệ thiết bị và nội dung số,\n- Phân biệt được các rủi ro và mối đe dọa đơn giản trong môi trường số,\n- Chọn được các biện pháp an toàn và bảo mật đơn giản,\n- Nhận biết được các yếu tố đơn giản của độ tin cậy và quyền riêng tư.",
            "L4-5": "- Xác định được các cách đơn giản để bảo vệ thiết bị và nội dung số,\n- Phân biệt được các rủi ro và mối đe dọa đơn giản trong môi trường số,\n- Chọn được các biện pháp an toàn và bảo mật đơn giản,\n- Nhận biết được các yếu tố đơn giản của độ tin cậy và quyền riêng tư.",
            "L6-7": "- Chỉ ra được các cách thức thông thường và được xác định rõ ràng để bảo vệ thiết bị và nội dung số,\n- Phân biệt được các rủi ro và mối đe dọa được xác định rõ ràng và thông thường trong môi trường số,\n- Lựa chọn các biện pháp an toàn và bảo mật thường xuyên và được xác định rõ ràng,\n- Giải thích được các yếu tố thường xuyên và được xác định rõ ràng về độ tin cậy và quyền riêng tư.",
            "L8-9": "- Tổ chức được các cách để bảo vệ thiết bị và nội dung số,\n- Phân biệt được các rủi ro và mối đe dọa trong môi trường số,\n- Vận dụng được các biện pháp an toàn và bảo mật,\n- Thảo luận về các yếu tố của độ tin cậy và quyền riêng tư.",
            "L10-12": "- Áp dụng được các cách khác nhau để bảo vệ thiết bị và nội dung số,\n- Phân biệt được nhiều rủi ro và mối đe dọa trong môi trường số,\n- Sử dụng được nhiều biện pháp an toàn và bảo mật khác nhau,\n- Áp dụng được các yếu tố khác nhau về độ tin cậy và quyền riêng tư."
        }
    },
    "4.2": {
        code: "4.2",
        name: "Bảo vệ dữ liệu cá nhân và quyền riêng tư",
        mienId: 4,
        mienName: "An toàn",
        mota: "Bảo vệ được dữ liệu cá nhân và quyền riêng tư trong môi trường số. Nhận biết được cách sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời có thể tự bảo vệ bản thân và người khác khỏi bị tổn hại. Hiểu được các dịch vụ số sử dụng \"Chính sách quyền riêng tư\" để thông tin về cách sử dụng dữ liệu cá nhân.",
        descriptors: {
            "L1-3": "- Chọn các cách đơn giản để bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số,\n- Nhận biết các cách đơn giản để sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời bảo vệ bản thân và người khác khỏi bị tổn hại.\n- Xác định các tuyên bố đơn giản về quyền riêng tư của các dịch vụ số về cách sử dụng dữ liệu cá nhân.",
            "L4-5": "- Chọn các cách đơn giản để bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số,\n- Nhận biết các cách đơn giản để sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời bảo vệ bản thân và người khác khỏi bị tổn hại.\n- Xác định các tuyên bố đơn giản về quyền riêng tư của các dịch vụ số về cách sử dụng dữ liệu cá nhân.",
            "L6-7": "- Chỉ ra các cách thường xuyên và được xác định rõ ràng để bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số,\n- Thảo luận các cách thường xuyên và được xác định rõ ràng để sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời bảo vệ bản thân và người khác khỏi bị tổn hại.\n- Giải thích các tuyên bố về quyền riêng tư được xác định rõ ràng và thông thường của các dịch vụ số về cách sử dụng dữ liệu cá nhân.",
            "L8-9": "- Thảo luận về cách bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số,\n- Thảo luận về cách sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời bảo vệ bản thân và người khác khỏi bị tổn hại.\n- Giải thích các tuyên bố về quyền riêng tư của các dịch vụ số về cách sử dụng dữ liệu cá nhân.",
            "L10-12": "- Áp dụng được các cách khác nhau để bảo vệ dữ liệu cá nhân và quyền riêng tư trong môi trường số,\n- Áp dụng các cách khác nhau để sử dụng và chia sẻ thông tin nhận dạng cá nhân đồng thời bảo vệ bản thân và người khác khỏi bị tổn hại.\n- Sử dụng được các chính sách bảo mật thông tin khác nhau của các dịch vụ số về cách sử dụng dữ liệu cá nhân."
        }
    },
    "4.3": {
        code: "4.3",
        name: "Bảo vệ sức khỏe và an sinh số",
        mienId: 4,
        mienName: "An toàn",
        mota: "Tránh được các rủi ro về sức khỏe và các mối đe dọa đối với hạnh phúc thể chất và tâm lý khi sử dụng công nghệ số. Có khả năng tự bảo vệ mình và người khác khỏi các mối nguy hiểm tiềm ẩn trong môi trường số (ví dụ: bắt nạt trên mạng). Nhận thức được các công nghệ số đối với an sinh và hòa nhập xã hội.",
        descriptors: {
            "L1-3": "- Phân biệt được các cách đơn giản để tránh rủi ro về sức khỏe và các mối đe dọa đối với sức khỏe thể chất và tâm lý khi sử dụng công nghệ số.\n- Chọn được các cách đơn giản để bảo vệ bản thân khỏi những mối nguy hiểm tiềm ẩn trong môi trường số.\n- Nhận biết được các công nghệ số đơn giản cho sự an sinh và hòa nhập xã hội.",
            "L4-5": "- Phân biệt được các cách đơn giản để tránh rủi ro về sức khỏe và các mối đe dọa đối với sức khỏe thể chất và tâm lý khi sử dụng công nghệ số.\n- Chọn được các cách đơn giản để bảo vệ bản thân khỏi những mối nguy hiểm tiềm ẩn trong môi trường số.\n- Nhận biết được các công nghệ số đơn giản cho sự an sinh và hòa nhập xã hội.",
            "L6-7": "- Giải thích các cách thường xuyên và được xác định rõ ràng để tránh rủi ro về sức khỏe và các mối đe dọa đối với sức khỏe thể chất và tâm lý khi sử dụng công nghệ số.\n- Chỉ ra các cách thường xuyên và được xác định rõ ràng để bảo vệ bản thân và người khác khỏi các mối nguy hiểm trong môi trường số.\n- Giải thích các công nghệ số thường xuyên và được xác định rõ ràng về an sinh và hòa nhập xã hội.",
            "L8-9": "- Thảo luận về các cách để tránh rủi ro về sức khỏe và các mối đe dọa đối với sức khỏe thể chất và tâm lý khi sử dụng công nghệ số.\n- Thảo luận các cách để bảo vệ bản thân và người khác khỏi các nguy hiểm trong môi trường số.\n- Thảo luận về các công nghệ số cho an sinh và hòa nhập xã hội.",
            "L10-12": "- Áp dụng các cách khác nhau để tránh rủi ro về sức khỏe và các mối đe dọa đối với sức khỏe thể chất và tâm lý khi sử dụng công nghệ số.\n- Áp dụng các cách khác nhau để bảo vệ bản thân và người khác khỏi các mối nguy hiểm trong môi trường số.\n- Sử dụng được các công nghệ số khác nhau cho an sinh và hòa nhập xã hội."
        }
    },
    "4.4": {
        code: "4.4",
        name: "Bảo vệ môi trường",
        mienId: 4,
        mienName: "An toàn",
        mota: "Nhận thức được tác động môi trường của các công nghệ số và việc sử dụng chúng.",
        descriptors: {
            "L1-3": "- Nhận biết được các tác động đơn giản đến môi trường của công nghệ số và việc sử dụng chúng.",
            "L4-5": "- Nhận biết được các tác động đơn giản đến môi trường của công nghệ số và việc sử dụng chúng.",
            "L6-7": "- Giải thích được các tác động thông thường và được xác định rõ ràng đến môi trường của công nghệ số và việc sử dụng chúng.",
            "L8-9": "- Thảo luận về tác động môi trường của công nghệ số và việc sử dụng chúng.",
            "L10-12": "- Khuyến nghị được nhiều biện pháp khác nhau để giảm thiểu tác động môi trường của công nghệ số và việc sử dụng chúng."
        }
    },
    "5.1": {
        code: "5.1",
        name: "Giải quyết các vấn đề kỹ thuật",
        mienId: 5,
        mienName: "Giải quyết vấn đề",
        mota: "Xác định được các vấn đề kỹ thuật khi vận hành thiết bị và sử dụng môi trường số, và giải quyết chúng (từ xử lý sự cố đến giải quyết các vấn đề phức tạp hơn).",
        descriptors: {
            "L1-3": "- Xác định được các vấn đề kỹ thuật đơn giản khi vận hành thiết bị và sử dụng môi trường số.\n- Chọn được các giải pháp đơn giản cho chúng.",
            "L4-5": "- Xác định được các vấn đề kỹ thuật đơn giản khi vận hành thiết bị và sử dụng môi trường số.\n- Chọn được các giải pháp đơn giản cho chúng.",
            "L6-7": "- Chỉ ra các vấn đề kỹ thuật thông thường và được xác định rõ ràng khi vận hành thiết bị và sử dụng môi trường số.\n- Lựa chọn các giải pháp thường xuyên và được xác định rõ ràng cho chúng.",
            "L8-9": "- Phân biệt được các vấn đề kỹ thuật khi vận hành thiết bị và sử dụng môi trường số.\n- Vận dụng được các giải pháp cho chúng.",
            "L10-12": "- Phân biệt được nhiều vấn đề kỹ thuật khi vận hành thiết bị và sử dụng môi trường số.\n- Áp dụng được các giải pháp tối ưu cho chúng."
        }
    },
    "5.2": {
        code: "5.2",
        name: "Xác định nhu cầu và giải pháp công nghệ",
        mienId: 5,
        mienName: "Giải quyết vấn đề",
        mota: "Đánh giá được nhu cầu và xác định, đánh giá, lựa chọn và sử dụng được các công cụ số và các phản hồi công nghệ có thể để giải quyết chúng. Điều chỉnh và tùy chỉnh các môi trường số theo nhu cầu cá nhân.",
        descriptors: {
            "L1-3": "- Nhận biết được nhu cầu cá nhân và chọn các công cụ số đơn giản và các biện pháp phản hồi công nghệ đơn giản để giải quyết chúng.\n- Chọn được các cách đơn giản để điều chỉnh và tùy chỉnh môi trường số theo nhu cầu cá nhân.",
            "L4-5": "- Nhận biết được nhu cầu cá nhân và chọn các công cụ số đơn giản và các biện pháp phản hồi công nghệ đơn giản để giải quyết chúng.\n- Chọn được các cách đơn giản để điều chỉnh và tùy chỉnh môi trường số theo nhu cầu cá nhân.",
            "L6-7": "- Chỉ ra được nhu cầu rõ ràng và lựa chọn các công cụ số thông thường cũng như các phản hồi công nghệ thường xuyên để giải quyết chúng.\n- Lựa chọn các cách thông thường và được xác định rõ ràng để điều chỉnh và tùy biến môi trường số theo nhu cầu cá nhân.",
            "L8-9": "- Đánh giá được nhu cầu của bản thân và lựa chọn các công cụ số cũng như các phản hồi công nghệ để giải quyết chúng.\n- Vận dụng được các cách để điều chỉnh và tùy biến môi trường số theo nhu cầu cá nhân.",
            "L10-12": "- Đánh giá được nhu cầu và áp dụng các công cụ số khác nhau cũng như các phản hồi công nghệ khả thi để giải quyết chúng.\n- Áp dụng các cách khác nhau để điều chỉnh và tùy biến môi trường số theo nhu cầu cá nhân."
        }
    },
    "5.3": {
        code: "5.3",
        name: "Sử dụng sáng tạo công nghệ số",
        mienId: 5,
        mienName: "Giải quyết vấn đề",
        mota: "Sử dụng được các công cụ và công nghệ số để tạo ra tri thức và đổi mới các quá trình và sản phẩm. Tham gia vào việc nhận thức của cá nhân và tập thể để hiểu và giải quyết các vấn đề mang tính khái niệm và các tình huống có vấn đề trong môi trường số.",
        descriptors: {
            "L1-3": "- Nhận biết các công cụ và công nghệ số đơn giản có thể được sử dụng để tạo ra kiến thức và đổi mới các quy trình và sản phẩm.\n- Tuân theo các giải pháp đơn giản cho các vấn đề khái niệm và các tình huống có vấn đề trong môi trường số.",
            "L4-5": "- Nhận biết các công cụ và công nghệ số đơn giản có thể được sử dụng để tạo ra kiến thức và đổi mới các quy trình và sản phẩm.\n- Tuân theo các giải pháp đơn giản cho các vấn đề khái niệm và các tình huống có vấn đề trong môi trường số.",
            "L6-7": "- Chỉ ra các công cụ và công nghệ số được xác định rõ ràng và thông thường có thể được sử dụng để tạo ra kiến thức và đổi mới các quy trình và sản phẩm.\n- Lựa chọn các giải pháp thường xuyên và được xác định rõ ràng cho các vấn đề mang tính khái niệm và các tình huống có vấn đề trong môi trường số.",
            "L8-9": "- Vận dụng được các công cụ và công nghệ số để tạo ra kiến thức, đổi mới các quá trình và sản phẩm.\n- Giải quyết được các vấn đề mang tính khái niệm và các tình huống có vấn đề trong môi trường số.",
            "L10-12": "- Áp dụng được các công cụ và công nghệ số khác nhau để tạo ra tri thức và đổi mới các quá trình và sản phẩm.\n- Giải quyết các vấn đề phức tạp và các tình huống có vấn đề trong môi trường số."
        }
    },
    "5.4": {
        code: "5.4",
        name: "Xác định các vấn đề cần cải thiện về năng lực số",
        mienId: 5,
        mienName: "Giải quyết vấn đề",
        mota: "Hiểu được nơi nào cần cải thiện hoặc cập nhật năng lực số của bản thân. Có thể hỗ trợ người khác phát triển năng lực số của họ. Tìm kiếm được cơ hội phát triển bản thân và cập nhật được sự phát triển số.",
        descriptors: {
            "L1-3": "- Nhận biết được nơi mà năng lực số của bản thân cần được cải thiện hoặc cập nhật.\n- Xác định nơi để tìm kiếm các cơ hội phát triển bản thân và cập nhật thông tin về sự phát triển của công nghệ số.",
            "L4-5": "- Nhận biết được nơi mà năng lực số của bản thân cần được cải thiện hoặc cập nhật.\n- Xác định nơi để tìm kiếm các cơ hội phát triển bản thân và cập nhật thông tin về sự phát triển của công nghệ số.",
            "L6-7": "- Giải thích được nơi mà năng lực số của bản thân cần được cải thiện hoặc cập nhật.\n- Chỉ ra các cách thường xuyên và được xác định rõ ràng để hỗ trợ người khác phát triển năng lực số của họ.\n- Lựa chọn các cơ hội học tập được xác định rõ ràng và thường xuyên để tự phát triển và cập nhật sự phát triển của công nghệ số.",
            "L8-9": "- Thảo luận về nơi cần cải thiện hoặc cập nhật năng lực số của bản thân.\n- Thảo luận các cách để hỗ trợ người khác phát triển năng lực số của họ.\n- Đề xuất các cơ hội tự phát triển và cập nhật sự phát triển của công nghệ số.",
            "L10-12": "- Đánh giá được điểm mạnh và điểm yếu về năng lực số của bản thân.\n- Hướng dẫn và hỗ trợ người khác phát triển năng lực số.\n- Chủ động tìm kiếm và theo đuổi các cơ hội nâng cao năng lực số tiên tiến."
        }
    },
    "6.1": {
        code: "6.1",
        name: "Hiểu biết về trí tuệ nhân tạo (AI)",
        mienId: 6,
        mienName: "Ứng dụng trí tuệ nhân tạo (AI)",
        mota: "Hiểu được các khái niệm cơ bản về AI, nguyên lý hoạt động, các loại AI phổ biến và khả năng cũng như giới hạn của AI trong đời sống thực tế.",
        descriptors: {
            "L1-3": "- Nhận biết được các ví dụ đơn giản về AI trong đời sống hàng ngày (trợ lý giọng nói, robot hút bụi).\n- Nhận biết được AI là sản phẩm do con người tạo ra và lập trình.",
            "L4-5": "- Nhận biết và mô tả được một số ứng dụng AI quen thuộc.\n- Hiểu được AI cần có dữ liệu để học và hoạt động theo sự chỉ dẫn của con người.",
            "L6-7": "- Giải thích được khái niệm cơ bản về AI và sự khác biệt giữa phần mềm thông thường và AI.\n- Nhận biết được các thành phần chính của một hệ thống AI (thu thập dữ liệu, xử lý, đưa ra kết quả).",
            "L8-9": "- Trình bày được nguyên lý học máy cơ bản và vai trò cốt lõi của dữ liệu đối với chất lượng mô hình AI.\n- Phân tích được các ưu điểm và hạn chế của công nghệ AI trong các lĩnh vực.",
            "L10-12": "- Phân tích sâu về kiến trúc và các loại mô hình AI khác nhau.\n- Đánh giá được tiềm năng phát triển và giới hạn kỹ thuật của các hệ thống AI hiện đại."
        }
    },
    "6.2": {
        code: "6.2",
        name: "Sử dụng trí tuệ nhân tạo (AI)",
        mienId: 6,
        mienName: "Ứng dụng trí tuệ nhân tạo (AI)",
        mota: "Sử dụng được các công cụ và dịch vụ AI phục vụ học tập, giải quyết vấn đề và sáng tạo; biết cách đưa ra câu lệnh (prompt) hiệu quả và tinh chỉnh kết quả đầu ra của AI.",
        descriptors: {
            "L1-3": "- Sử dụng được các tính năng tương tác cơ bản bằng giọng nói hoặc hình ảnh với thiết bị AI theo hướng dẫn.",
            "L4-5": "- Sử dụng được công cụ AI đơn giản để tra cứu thông tin hoặc hỗ trợ học tập dưới sự giám sát của giáo viên/phụ huynh.",
            "L6-7": "- Soạn thảo được câu lệnh (prompt) rõ ràng, cụ thể để yêu cầu AI thực hiện nhiệm vụ học tập.\n- Sử dụng được các công cụ AI để tìm kiếm ý tưởng, tóm tắt tài liệu hoặc hỗ trợ viết mã đơn giản.",
            "L8-9": "- Áp dụng được kỹ thuật đặt câu lệnh (prompt engineering) để thu được kết quả chính xác từ mô hình ngôn ngữ lớn (LLM).\n- Tích hợp công cụ AI vào quy trình giải quyết vấn đề học tập và tạo sản phẩm số.",
            "L10-12": "- Tùy biến và kết hợp nhiều công cụ AI chuyên biệt để tối ưu hóa quy trình làm việc và giải quyết vấn đề phức tạp."
        }
    },
    "6.3": {
        code: "6.3",
        name: "Đánh giá trí tuệ nhân tạo (AI)",
        mienId: 6,
        mienName: "Ứng dụng trí tuệ nhân tạo (AI)",
        mota: "Phân tích, đánh giá tính chính xác, độ tin cậy và sự thiên vị của kết quả do AI tạo ra; nhận thức được các vấn đề đạo đức, quyền riêng tư và tác động xã hội khi sử dụng AI.",
        descriptors: {
            "L1-3": "- Nhận biết được thông tin do AI đưa ra có thể không chính xác và cần hỏi lại người lớn.",
            "L4-5": "- Biết kiểm chứng lại thông tin do AI cung cấp từ các nguồn đáng tin cậy khác (sách, thầy cô).",
            "L6-7": "- Đánh giá được tính hợp lý của kết quả đầu ra từ AI; phát hiện được các lỗi hiển nhiên hoặc ảo giác (hallucination) của AI.\n- Nhận thức được việc không chia sẻ thông tin cá nhân bí mật khi trò chuyện với AI.",
            "L8-9": "- Phân tích được các nguy cơ về bản quyền, tính thiên vị (bias) của dữ liệu huấn luyện và gian lận học thuật khi dùng AI.\n- Thực hiện trích dẫn minh bạch khi sử dụng nội dung do AI hỗ trợ tạo ra.",
            "L10-12": "- Đánh giá toàn diện tác động đạo đức, pháp lý và xã hội của các ứng dụng AI.\n- Đề xuất các nguyên tắc sử dụng AI có trách nhiệm và an toàn trong cộng đồng."
        }
    }
};

// =============================================================================
// II. QUYẾT ĐỊNH 3439/QĐ-BGDĐT — KHUNG NĂNG LỰC TRÍ TUỆ NHÂN TẠO (AI)
// =============================================================================
const QD3439_AI_DATA = {
    A: {
        code: "A",
        name: "Tư duy lấy con người làm trung tâm (Human-centred mindset)",
        concept: "Khẳng định con người là chủ thể sáng tạo, kiểm soát và thụ hưởng AI. Nhận thức rõ AI chỉ là công cụ mô phỏng trí tuệ và phản ứng, con người luôn chịu trách nhiệm tối thượng về các quyết định của mình.",
        subthemes: {
            "A1": {
                name: "Tính chủ động của con người",
                content: "Con người có cảm xúc thật và ý thức độc lập, AI không có cảm xúc thật mà chỉ mô phỏng phản ứng theo dữ liệu huấn luyện. Con người luôn là người ra quyết định và kiểm soát hệ thống."
            },
            "A2": {
                name: "AI vì sự tiến bộ của con người",
                content: "Nhận diện AI trong cuộc sống; hiểu mục đích của AI là hỗ trợ con người nâng cao năng suất, giải phóng sức lao động và phục vụ xã hội văn minh."
            }
        }
    },
    B: {
        code: "B",
        name: "Đạo đức AI (Ethics of AI)",
        concept: "Các nguyên tắc chỉ đạo về đạo đức, công bằng, minh bạch, bảo vệ quyền riêng tư và trách nhiệm giải trình khi thiết kế, triển khai và sử dụng hệ thống trí tuệ nhân tạo.",
        subthemes: {
            "B1": {
                name: "Các khía cạnh đạo đức của AI",
                content: "Phân biệt việc làm tốt và việc làm xấu bằng AI; không sử dụng AI để lừa đảo, giả mạo (deepfake), phát tán thông tin sai lệch hoặc gây tổn hại cho người khác."
            },
            "B2": {
                name: "Tác động của AI đối với xã hội",
                content: "Nhận thức về sự thiên vị dữ liệu (data bias), phân biệt đối xử của thuật toán, tác động đến thị trường việc làm và văn hóa số."
            },
            "B3": {
                name: "Nguyên tắc đạo đức và trách nhiệm xã hội",
                content: "Máy thông minh làm việc tốt; tôn trọng bản quyền, tính minh bạch và trách nhiệm giải trình của con người khi sử dụng sản phẩm từ AI."
            }
        }
    },
    C: {
        code: "C",
        name: "Kỹ thuật và ứng dụng AI (AI techniques & applications)",
        concept: "Hệ thống kiến thức và kỹ năng về các công nghệ nền tảng của AI (học máy, thị giác máy tính, NLP, mô hình ngôn ngữ lớn) và cách vận dụng chúng vào học tập và đời sống.",
        subthemes: {
            "C1": {
                name: "Đặc điểm chính của AI",
                content: "Nhận biết AI và ứng dụng; hiểu các chức năng cơ bản: nhận dạng mẫu, xử lý hình ảnh, xử lý âm thanh, văn bản."
            },
            "C2": {
                name: "Ứng dụng AI trong học tập và cuộc sống",
                content: "Sử dụng trợ lý ảo, dịch thuật thời gian thực, gợi ý học tập cá nhân hóa, công cụ sáng tạo nội dung số hỗ trợ học sinh."
            },
            "C3": {
                name: "Công nghệ AI",
                content: "Các trụ cột kỹ thuật: Học máy (Machine Learning), Thị giác máy tính (Computer Vision), Xử lý ngôn ngữ tự nhiên (NLP), Mạng nơ-ron nhân tạo."
            },
            "C4": {
                name: "Dữ liệu trong AI",
                content: "Dữ liệu là thức ăn của AI; vai trò thu thập, làm sạch, gắn nhãn dữ liệu huấn luyện; chất lượng dữ liệu quyết định độ chính xác của AI."
            },
            "C5": {
                name: "Kỹ thuật và thuật toán AI",
                content: "Nguyên lý thuật toán học có giám sát, học không giám sát, học tăng cường và mô hình dự đoán xác suất."
            }
        }
    },
    D: {
        code: "D",
        name: "Thiết kế hệ thống AI (AI system design)",
        concept: "Năng lực phân tích bài toán thực tế, lựa chọn giải pháp AI, thu thập dữ liệu, huấn luyện mô hình đơn giản, đánh giá độ chính xác và cải tiến hệ thống.",
        subthemes: {
            "D1": {
                name: "Nhận diện & hình thành giải pháp",
                content: "Nhận diện các vấn đề thực tiễn có thể giải quyết bằng AI (nhận diện rác thải, phân loại ảnh, chatbot hỏi đáp); hình thành ý tưởng giải pháp AI."
            },
            "D2": {
                name: "Cấu trúc, tương tác & cải tiến hệ thống",
                content: "Quy trình thiết kế hệ thống AI: Thu thập mẫu -> Huấn luyện mô hình -> Kiểm thử với mẫu mới -> Tinh chỉnh và tích hợp vào ứng dụng."
            }
        }
    }
};

// =============================================================================
// III. CHƯƠNG TRÌNH GDPT 2018 — 5 PHẨM CHẤT CHỦ YẾU (ÁP DỤNG MỌI MÔN HỌC)
// =============================================================================
const GDPT2018_5_PHAM_CHAT = [
    {
        id: "yeu-nuoc",
        name: "Yêu nước",
        emoji: "🇻🇳",
        concept: "Tình yêu đối với quê hương, đất nước, thiên nhiên; niềm tự hào về truyền thống dựng nước và giữ nước của dân tộc; có ý thức bảo vệ di sản, văn hóa và môi trường sống của Tổ quốc.",
        tieuHoc: "Yêu thiên nhiên, có những việc làm thiết thực bảo vệ cây xanh; Yêu quê hương, kính trọng Quốc kì, Quốc ca; Kính trọng, biết ơn thầy cô và những người có công với đất nước.",
        thcs: "Tích cực, chủ động tham gia các hoạt động bảo vệ môi trường; Tự hào về truyền thống và danh nhân lịch sử dân tộc; Có ý thức tìm hiểu, gìn giữ bản sắc văn hóa dân tộc và sẵn sàng tham gia xây dựng đất nước.",
        khbdExample: "Tự hào về lịch sử dân tộc qua bài học Lịch sử / Giới thiệu di sản văn hóa qua bài trình chiếu đa phương tiện (Đạt được thông qua Hoạt động 3, Hoạt động 4)."
    },
    {
        id: "nhan-ai",
        name: "Nhân ái",
        emoji: "❤️",
        concept: "Yêu quý mọi người; Tôn trọng sự khác biệt về hoàn cảnh, tính cách, quan điểm; Sẵn sàng cảm thông, chia sẻ, giúp đỡ bạn bè; Văn minh, lịch sự và không gây tổn thương cho người khác cả ngoài đời lẫn trên môi trường số.",
        tieuHoc: "Yêu quý bạn bè, kính trọng thầy cô; quan tâm, chăm sóc người thân trong gia đình; Không phân biệt đối xử; Biết nói lời xin lỗi, cảm ơn và biết tha thứ.",
        thcs: "Tôn trọng danh dự, nhân phẩm, quyền riêng tư của người khác; Không đồng tình với cái ác, sự bất công và bạo lực học đường; Sẵn sàng cảm thông, chia sẻ và hỗ trợ bạn bè cùng tiến bộ.",
        khbdExample: "Lắng nghe, tôn trọng quan điểm khác biệt trong thảo luận nhóm môn Ngữ văn/KHTN; sẵn sàng hướng dẫn bạn khi làm bài tập chung (Đạt được thông qua Hoạt động 2, Hoạt động 3)."
    },
    {
        id: "cham-chi",
        name: "Chăm chỉ",
        emoji: "📚",
        concept: "Chăm học, chăm làm, có tinh thần tự giác, say mê tìm tòi khám phá tri thức; kiên trì, nỗ lực vượt qua khó khăn để hoàn thành tốt nhiệm vụ học tập và rèn luyện.",
        tieuHoc: "Đi học đầy đủ, đúng giờ; thường xuyên hoàn thành bài tập; thích đọc sách và khám phá thế giới xung quanh; tự giác làm việc nhà vừa sức.",
        thcs: "Luôn cố gắng vươn lên đạt kết quả tốt trong học tập; có ý chí vượt khó; kiên trì làm bài tập thực hành, không nản lòng khi gặp bài toán khó hoặc thí nghiệm phức tạp.",
        khbdExample: "Chủ động chuẩn bị bài học trước giờ lên lớp; kiên trì giải quyết các bài tập tư duy và hoàn thành nhiệm vụ đúng thời hạn (Đạt được thông qua Hoạt động 3)."
    },
    {
        id: "trung-thuc",
        name: "Trung thực",
        emoji: "⚖️",
        concept: "Thật thà, ngay thẳng trong học tập, lao động và sinh hoạt; dũng cảm nhận lỗi và sửa sai; tôn trọng lẽ phải; tôn trọng sự thật và bản quyền sở hữu trí tuệ.",
        tieuHoc: "Thật thà, không nói dối; không tự tiện lấy đồ của người khác; biết nhận lỗi khi làm sai; trung thực trong kiểm tra, đánh giá.",
        thcs: "Luôn thống nhất giữa lời nói và việc làm; nghiêm túc trong học tập và thi cử; không chép bài bạn; tôn trọng bản quyền tác giả và trích dẫn rõ ràng nguồn tài liệu tham khảo.",
        khbdExample: "Tự giác làm bài kiểm tra; trung thực ghi chép số liệu quan sát thực tế trong thí nghiệm KHTN/Toán học; trích dẫn nguồn khi làm bài văn (Đạt được thông qua Hoạt động 3, Hoạt động 4)."
    },
    {
        id: "trach-nhiem",
        name: "Trách nhiệm",
        emoji: "🛡️",
        concept: "Có trách nhiệm với bản thân, gia đình, nhà trường, xã hội và môi trường; giữ lời hứa; dám nhận trách nhiệm về lời nói, hành vi của mình; có ý thức bảo vệ của công và giữ an toàn số.",
        tieuHoc: "Có ý thức giữ gìn vệ sinh cá nhân và lớp học; bảo vệ đồ dùng học tập của mình và của trường lớp; tuân thủ quy tắc phòng học và nơi công cộng.",
        thcs: "Bảo vệ tài sản chung của nhà trường; sử dụng thiết bị phòng chức năng cẩn thận, đúng quy trình; chấp hành nội quy nhà trường và pháp luật; chủ động hoàn thành nhiệm vụ được giao trong nhóm.",
        khbdExample: "Có ý thức giữ gìn vệ sinh và tài sản chung của phòng học; tắt các thiết bị điện khi ra về; chủ động hoàn thành phần việc được nhóm phân công (Đạt được thông qua Hoạt động 2, Hoạt động 4)."
    }
];

// =============================================================================
// IV. CÔNG VĂN 5512/BGDĐT & CT GDPT 2018 — 3 NĂNG LỰC CHUNG (ÁP DỤNG MỌI MÔN HỌC)
// =============================================================================
const GDPT2018_3_NL_CHUNG = [
    {
        id: "tu-chu-tu-hoc",
        name: "Tự chủ và tự học",
        emoji: "🎯",
        concept: "Khả năng tự quản lý bản thân, chủ động thực hiện nhiệm vụ học tập mà không cần chờ nhắc nhở; tự đặt ra mục tiêu, lập kế hoạch và kiên trì thực hiện kế hoạch tự học, tự hoàn thiện.",
        thanhTo: [
            "Tự lực: Tự giác làm bài, tự tìm hiểu tài liệu học tập trong SGK và thư viện.",
            "Tự khẳng định và bảo vệ quyền, nhu cầu chính đáng của bản thân.",
            "Tự điều chỉnh cảm xúc, thái độ, hành vi của bản thân trong lớp học.",
            "Tự học, tự hoàn thiện: Tự đánh giá và rút kinh nghiệm sau mỗi tiết học."
        ],
        khbdExample: "Chủ động nghiên cứu SGK, tự đọc tài liệu hướng dẫn trước khi giáo viên thuyết trình (Đạt được thông qua Hoạt động 1, Hoạt động 2)."
    },
    {
        id: "giao-tiep-hop-tac",
        name: "Giao tiếp và hợp tác",
        emoji: "🤝",
        concept: "Khả năng diễn đạt rõ ràng tư tưởng, ý kiến bằng lời nói, chữ viết hoặc sản phẩm học tập; biết lắng nghe, tôn trọng quan điểm khác biệt; phối hợp nhịp nhàng, hiệu quả với các thành viên trong nhóm để cùng đạt mục tiêu chung.",
        thanhTo: [
            "Xác định mục đích, nội dung, phương tiện và thái độ giao tiếp phù hợp.",
            "Thiết lập và duy trì mối quan hệ hợp tác tích cực với bạn bè.",
            "Xác định trách nhiệm và phân công công việc cụ thể trong nhóm.",
            "Biết lắng nghe, thảo luận dân chủ và giải quyết bất đồng hòa nhã."
        ],
        khbdExample: "Tích cực trao đổi, phân công nhiệm vụ rõ ràng trong nhóm để hoàn thành dự án/bài thuyết trình chung (Đạt được thông qua Hoạt động 2, Hoạt động 3)."
    },
    {
        id: "gqvd-sang-tao",
        name: "Giải quyết vấn đề và sáng tạo",
        emoji: "💡",
        concept: "Khả năng phát hiện vấn đề trong học tập và cuộc sống; phân tích, làm rõ nguyên nhân; đề xuất các giải pháp khả thi, độc đáo và lựa chọn giải pháp tối ưu để giải quyết vấn đề hiệu quả.",
        thanhTo: [
            "Nhận ra ý tưởng mới và phát hiện vấn đề cần giải quyết từ tình huống thực tế.",
            "Hình thành và phát triển ý tưởng sáng tạo trong sản phẩm học tập.",
            "Đề xuất, thử nghiệm và lựa chọn giải pháp tối ưu cho bài toán.",
            "Tư duy độc lập và linh hoạt khi xử lý các tình huống phát sinh."
        ],
        khbdExample: "Phát hiện mối liên hệ mới giữa các kiến thức, đề xuất phương pháp giải sáng tạo và tối ưu hơn (Đạt được thông qua Hoạt động 3, Hoạt động 4)."
    }
];

// =============================================================================
// V. NĂNG LỰC ĐẶC THÙ — THEO MÔN HỌC
// =============================================================================
const NL_DACTHU_TIN_HOC = [
    {
        code: "NLa",
        name: "Sử dụng và quản lý các phương tiện công nghệ thông tin và truyền thông",
        desc: "Biết sử dụng các thiết bị phần cứng, hệ điều hành, phần mềm ứng dụng thông dụng và dịch vụ mạng an toàn, hiệu quả để xử lý công việc và học tập.",
        khbdExample: "NLa (Đạt được thông qua Hoạt động 2, Hoạt động 3): Thành thạo thao tác khởi động máy tính, mở phần mềm và lưu tệp tin đúng quy định."
    },
    {
        code: "NLb",
        name: "Ứng xử phù hợp trong môi trường số",
        desc: "Tuân thủ các quy tắc đạo đức, pháp luật, chuẩn mực văn hóa khi giao tiếp, chia sẻ thông tin trên Internet; tôn trọng bản quyền và bảo vệ an toàn thông tin số.",
        khbdExample: "NLb (Đạt được thông qua Hoạt động 1, Hoạt động 4): Tuân thủ quy định an toàn mạng, tôn trọng bản quyền hình ảnh khi sử dụng tư liệu trên Internet."
    },
    {
        code: "NLc",
        name: "Khám phá tri thức, giải quyết vấn đề với sự trợ giúp của CNTT và truyền thông",
        desc: "Sử dụng máy tính và mạng Internet như một công cụ hỗ trợ tư duy, tìm kiếm, khai thác dữ liệu, phân tích thông tin và giải quyết vấn đề học tập.",
        khbdExample: "NLc (Đạt được thông qua Hoạt động 2, Hoạt động 3): Sử dụng công cụ tìm kiếm trên Internet để thu thập dữ liệu phục vụ bài học."
    },
    {
        code: "NLd",
        name: "Ứng dụng CNTT và truyền thông trong học và tự học",
        desc: "Chủ động sử dụng phần mềm, các khóa học trực tuyến và tài nguyên số để mở rộng kiến thức, tự học và nâng cao hiệu quả học tập suốt đời.",
        khbdExample: "NLd (Đạt được thông qua Hoạt động 3, Hoạt động 4): Sử dụng phần mềm học tập để tự rèn luyện kỹ năng."
    },
    {
        code: "NLe",
        name: "Hợp tác trong môi trường số",
        desc: "Sử dụng các công cụ truyền thông số, lưu trữ đám mây để trao đổi thông tin, chia sẻ tài liệu và cùng nhau thực hiện dự án học tập theo nhóm.",
        khbdExample: "NLe (Đạt được thông qua Hoạt động 3): Sử dụng công cụ chia sẻ trực tuyến để làm việc nhóm và đóng góp ý kiến cho sản phẩm chung."
    }
];

const NL_DACTHU_ROBOTICS = [
    {
        code: "NL1",
        name: "Nhận thức công nghệ và cơ chế kỹ thuật robot",
        desc: "Hiểu biết về các linh kiện điện tử, cảm biến, động cơ, mạch điều khiển và các cơ cấu chuyển động cơ học trong hệ thống robot.",
        khbdExample: "NL1 (Đạt được thông qua Hoạt động 1, Hoạt động 2): Nhận biết đúng tên gọi và chức năng của động cơ và cảm biến."
    },
    {
        code: "NL2",
        name: "Lắp ráp và chế tạo mô hình robot",
        desc: "Khả năng đọc bản vẽ kỹ thuật, chọn đúng linh kiện và lắp ráp mô hình robot chắc chắn, chính xác theo hướng dẫn hoặc sáng tạo mô hình mới.",
        khbdExample: "NL2 (Đạt được thông qua Hoạt động 2, Hoạt động 3): Đọc sơ đồ lắp ráp và thực hiện lắp ráp hoàn thiện khung xe robot đúng kỹ thuật."
    },
    {
        code: "NL3",
        name: "Lập trình điều khiển robot",
        desc: "Sử dụng ngôn ngữ lập trình khối hoặc văn bản để viết thuật toán điều khiển động cơ, đọc dữ liệu cảm biến và xử lý tình huống cho robot.",
        khbdExample: "NL3 (Đạt được thông qua Hoạt động 3): Lập trình thuật toán điều khiển robot tự động dừng lại khi phát hiện vật cản."
    },
    {
        code: "NL4",
        name: "Thử nghiệm, tinh chỉnh và xử lý sự cố (Troubleshooting)",
        desc: "Kiểm tra thực tế hoạt động của robot trên sa bàn, phát hiện lỗi phần cứng hoặc logic phần mềm, đo đạc thông số và điều chỉnh tối ưu hóa hoạt động.",
        khbdExample: "NL4 (Đạt được thông qua Hoạt động 3, Hoạt động 4): Thử nghiệm robot trên sa bàn, phát hiện nguyên nhân lỗi và tinh chỉnh thông số cảm biến."
    },
    {
        code: "NL5",
        name: "Hợp tác, giải quyết vấn đề kỹ thuật và tư duy sáng tạo",
        desc: "Phối hợp nhóm hiệu quả để giải quyết các thử thách nhiệm vụ robot thực tế; đề xuất cải tiến cấu trúc cơ khí hoặc chiến thuật lập trình sáng tạo.",
        khbdExample: "NL5 (Đạt được thông qua Hoạt động 3, Hoạt động 4): Thảo luận nhóm phân công vai trò lập trình - cơ khí để vượt qua thử thách sa bàn thi đấu."
    }
];

// =============================================================================
// VI. HỆ THỐNG CÂU HỎI ÔN TẬP PHỔ QUÁT THEO TỪNG VĂN BẢN (QUIZ BANK)
// Đảm bảo bám sát văn bản, ví dụ đa dạng các môn học, không thiên lệch Tin/Robotics
// =============================================================================
const QUIZ_BANK = {
    // -------------------------------------------------------------------------
    // BỘ 1: CÔNG VĂN 3456/BGDĐT & THÔNG TƯ 02/2025 (NĂNG LỰC SỐ PHỔ QUÁT)
    // -------------------------------------------------------------------------
    "cv3456": {
        title: "Công văn 3456/BGDĐT & TT 02/2025 — Năng lực số",
        badge: "💻 Khung NLS",
        subsections: {
            "cv3456_bac": {
                name: "Phần 1: Nguyên tắc 5 Bậc & Ký hiệu CB",
                desc: "Quy chuẩn phân bậc theo khối lớp, mức độ tự chủ và ký hiệu CB1 đến CB5.",
                questions: [
                    {
                        question: "Theo CV 3456, học sinh Lớp 6 và Lớp 7 thuộc Bậc năng lực số nào?",
                        options: ["Bậc 1 (Cơ bản 1 - CB1)", "Bậc 2 (Cơ bản 2 - CB2)", "Bậc 3 (Trung cấp 1 - CB3)", "Bậc 4 (Trung cấp 2 - CB4)"],
                        correct: 2,
                        explanation: "Căn cứ bảng phân Bậc chuẩn: Lớp 1-3 = Bậc 1 (CB1); Lớp 4-5 = Bậc 2 (CB2); Lớp 6-7 = Bậc 3 (CB3); Lớp 8-9 = Bậc 4 (CB4); Lớp 10-12 = Bậc 5 (CB5)."
                    },
                    {
                        question: "Khối lớp nào sau đây áp dụng ký hiệu chuẩn Bậc là CB4 trong KHBD?",
                        options: ["Lớp 4 và Lớp 5", "Lớp 6 và Lớp 7", "Lớp 8 và Lớp 9", "Lớp 10 và Lớp 11"],
                        correct: 2,
                        explanation: "CB4 (Trung cấp 2) áp dụng cho học sinh Lớp 8 và Lớp 9 theo đúng lộ trình phân Bậc của CV 3456."
                    },
                    {
                        question: "Mức độ tự chủ của học sinh Tiểu học ở Bậc 1 (Lớp 1-3) được quy định như thế nào?",
                        options: ["Tự chủ hoàn toàn không cần trợ giúp", "Dưới sự hướng dẫn của giáo viên hoặc người hỗ trợ", "Có thể tự đề xuất chiến lược mới", "Làm chủ hoàn toàn và hướng dẫn người khác"],
                        correct: 1,
                        explanation: "Ở Bậc 1 (Cơ bản 1), học sinh thực hiện các thao tác số đơn giản với sự hướng dẫn của giáo viên hoặc người lớn."
                    },
                    {
                        question: "Giáo viên THCS dạy Lớp 6 khi viết mục tiêu NLS cần lưu ý tránh lỗi phổ biến nào?",
                        options: ["Ghi nhầm thành Bậc 1", "Ghi nhầm thành Bậc 2 (CB2) do nghĩ Lớp 6 mới bắt đầu cấp 2", "Ghi nhầm thành Bậc 5", "Quên không ghi tên giáo viên"],
                        correct: 1,
                        explanation: "Giáo viên thường nhầm Lớp 6 là CB2, nhưng chuẩn theo CV 3456 Lớp 6-7 là Bậc 3 (Trung cấp 1 - CB3)."
                    },
                    {
                        question: "Ký hiệu 'CB5' trong hệ thống chuẩn Bậc NLS đại diện cho đối tượng người học nào?",
                        options: ["Học sinh THCS (Lớp 6-9)", "Học sinh THPT (Lớp 10, 11, 12)", "Cán bộ quản lý trường học", "Học sinh năng khiếu Tiểu học"],
                        correct: 1,
                        explanation: "CB5 (Nâng cao 1) là Bậc quy định cho học sinh cấp THPT gồm các Lớp 10, 11 và 12."
                    }
                ]
            },
            "cv3456_mien": {
                name: "Phần 2: 6 Miền & 24 Thành tố NLS",
                desc: "Phân biệt chính xác khái niệm 6 Miền và ứng dụng thực tiễn trong mọi môn học.",
                questions: [
                    {
                        question: "Khung Năng lực số của người học theo Thông tư 02/2025 gồm bao nhiêu Miền và bao nhiêu Thành tố?",
                        options: ["5 Miền và 20 Thành tố", "6 Miền và 24 Thành tố", "6 Miền và 18 Thành tố", "4 Miền và 16 Thành tố"],
                        correct: 1,
                        explanation: "Khung NLS theo Thông tư 02/2025 và CV 3456 gồm chính xác 6 Miền năng lực và 24 Thành tố năng lực cốt lõi."
                    },
                    {
                        question: "Khi học sinh tìm kiếm tư liệu lịch sử, văn học hoặc khoa học trên Internet và đánh giá tính xác thực của trang web, học sinh đang phát triển Miền NLS nào?",
                        options: ["Miền I: Khai thác dữ liệu và thông tin", "Miền III: Sáng tạo nội dung số", "Miền IV: An toàn", "Miền V: Giải quyết vấn đề"],
                        correct: 0,
                        explanation: "Miền I (Khai thác dữ liệu và thông tin) bao gồm tìm kiếm (1.1), đánh giá độ tin cậy và tính xác thực của thông tin (1.2) và quản lý lưu trữ (1.3)."
                    },
                    {
                        question: "Học sinh thiết kế bài thuyết trình PowerPoint/Canva cho môn Ngữ văn hoặc Tiếng Anh, biết trích dẫn nguồn hình ảnh tôn trọng bản quyền thuộc Miền nào?",
                        options: ["Miền II: Giao tiếp và hợp tác", "Miền III: Sáng tạo nội dung số (thành tố 3.1 & 3.3)", "Miền V: Giải quyết vấn đề", "Miền VI: Trí tuệ nhân tạo"],
                        correct: 1,
                        explanation: "Tạo sản phẩm thuyết trình (3.1 Phát triển nội dung số) và trích dẫn bản quyền (3.3 Thực thi bản quyền) thuộc Miền III: Sáng tạo nội dung số."
                    },
                    {
                        question: "Thành tố 2.5 'Quy tắc ứng xử trên mạng' (Netiquette) thể hiện rõ nhất qua hành vi nào của học sinh?",
                        options: [
                            "Biết cách gõ bàn phím thật nhanh",
                            "Giao tiếp lịch sự, hòa nhã, tôn trọng sự khác biệt văn hóa khi trao đổi nhóm trực tuyến",
                            "Tải thật nhiều tài liệu về máy tính",
                            "Cài đặt phần mềm diệt virus"
                        ],
                        correct: 1,
                        explanation: "Thành tố 2.5 thuộc Miền II quy định về chuẩn mực hành vi, văn hóa giao tiếp và tôn trọng sự đa dạng trong môi trường số."
                    },
                    {
                        question: "Hành động đặt mật khẩu mạnh, không chia sẻ số điện thoại và địa chỉ nhà lên diễn đàn công cộng thuộc Miền NLS nào?",
                        options: ["Miền I: Khai thác dữ liệu", "Miền IV: An toàn (thành tố 4.2 Bảo vệ dữ liệu cá nhân & quyền riêng tư)", "Miền III: Sáng tạo nội dung", "Miền VI: Ứng dụng AI"],
                        correct: 1,
                        explanation: "Bảo vệ thông tin định danh cá nhân và quyền riêng tư thuộc Thành tố 4.2 trong Miền IV: An toàn."
                    }
                ]
            },
            "cv3456_cuphap": {
                name: "Phần 3: Cú pháp viết mã NLS vào KHBD",
                desc: "Quy chuẩn viết mục 2.2 Năng lực số, tra cứu descriptor và liên kết hoạt động.",
                questions: [
                    {
                        question: "Trong cấu trúc mục tiêu KHBD theo chuẩn công văn hiện hành, Năng lực số nằm ở mục số mấy?",
                        options: ["Mục 1. Kiến thức", "Mục 2.1. Năng lực đặc thù", "Mục 2.2. Năng lực số", "Mục 3. Phẩm chất"],
                        correct: 2,
                        explanation: "Thứ tự chuẩn mục Năng lực: 2.1 = NL Đặc thù, 2.2 = NL Số (Thông tư 02/2025 - CV 3456), 2.3 = NL Chung, 2.4 = NL AI (nếu có). Phẩm chất ở Mục 3."
                    },
                    {
                        question: "Nội dung mô tả yêu cầu cần đạt trong mục NLS của KHBD bắt buộc phải trích xuất từ đâu?",
                        options: ["Giáo viên tự sáng tác theo ý hiểu cá nhân", "Sao chép nguyên văn descriptor từ CV 3456 theo đúng Bậc khối lớp", "Sao chép một đoạn văn từ SGK", "Chỉ cần ghi tên Miền là đủ"],
                        correct: 1,
                        explanation: "Bắt buộc trích xuất chính xác descriptor từ CV 3456 theo đúng Bậc khối lớp của học sinh, không tự ý thêm bớt từ ngữ pháp quy."
                    },
                    {
                        question: "Yếu tố bắt buộc phải có ở cuối mỗi gạch đầu dòng mô tả Năng lực số trong KHBD là gì?",
                        options: ["Điểm số đánh giá", "Mốc '(Đạt được thông qua Hoạt động X, Hoạt động Y)'", "Tên tác giả SGK", "Thời gian dạy tiết học"],
                        correct: 1,
                        explanation: "Mọi mục tiêu năng lực và phẩm chất đều bắt buộc phải chỉ rõ mốc hoạt động học tập tương ứng trong tiến trình bài dạy."
                    },
                    {
                        question: "Đối với bài dạy Lớp 7 có hoạt động tìm kiếm tài liệu trên mạng, cách viết mục 2.2 nào sau đây ĐÚNG chuẩn?",
                        options: [
                            "- Miền I. Khai thác dữ liệu và thông tin (thành tố 1.1. Duyệt, tìm kiếm và lọc dữ liệu, thông tin và nội dung số – Bậc 3): Giải thích được nhu cầu thông tin; Thực hiện được rõ ràng và theo quy trình các tìm kiếm để tìm dữ liệu, thông tin và nội dung trong môi trường số (Đạt được thông qua Hoạt động 2).",
                            "- Miền I: Học sinh biết dùng Google để tra cứu bài học.",
                            "- Năng lực số: Em tìm kiếm được thông tin trên mạng rất tốt.",
                            "- Miền 1 – Bậc 1: Tìm kiếm đơn giản dưới sự hướng dẫn."
                        ],
                        correct: 0,
                        explanation: "Đáp án A ghi đúng tên Miền La Mã, mã và tên thành tố, Bậc chuẩn (Lớp 7 = Bậc 3), trích nguyên văn descriptor từ CV 3456 và gắn mốc hoạt động."
                    }
                ]
            }
        }
    },

    // -------------------------------------------------------------------------
    // BỘ 2: QUYẾT ĐỊNH 3439/QĐ-BGDĐT (NĂNG LỰC TRÍ TUỆ NHÂN TẠO - AI)
    // -------------------------------------------------------------------------
    "qd3439": {
        title: "Quyết định 3439/QĐ-BGDĐT — Năng lực Trí tuệ nhân tạo (AI)",
        badge: "🤖 Khung NL AI",
        subsections: {
            "qd3439_chude": {
                name: "Phần 1: 4 Chủ đề trụ cột A, B, C, D",
                desc: "Khung năng lực AI gồm 4 chủ đề lớn: Tư duy, Đạo đức, Kỹ thuật và Thiết kế.",
                questions: [
                    {
                        question: "Khung năng lực AI cho học sinh phổ thông theo QĐ 3439 gồm mấy chủ đề chính?",
                        options: ["3 chủ đề", "4 chủ đề (A, B, C, D)", "5 chủ đề", "6 chủ đề"],
                        correct: 1,
                        explanation: "QĐ 3439/QĐ-BGDĐT quy định 4 chủ đề trụ cột: A. Tư duy lấy con người làm trung tâm, B. Đạo đức AI, C. Kỹ thuật và ứng dụng AI, D. Thiết kế hệ thống AI."
                    },
                    {
                        question: "Chủ đề A 'Tư duy lấy con người làm trung tâm' nhấn mạnh điều cốt lõi nào?",
                        options: [
                            "AI sẽ thay thế hoàn toàn con người trong tương lai",
                            "Con người luôn là chủ thể ra quyết định, kiểm soát công nghệ và chịu trách nhiệm tối thượng",
                            "Con người phải phục tùng mọi quyết định của máy móc",
                            "AI có thể tự chịu trách nhiệm pháp lý thay con người"
                        ],
                        correct: 1,
                        explanation: "Chủ đề A khẳng định con người là trung tâm: AI là công cụ hỗ trợ, con người giữ quyền quyết định và chịu trách nhiệm cao nhất."
                    },
                    {
                        question: "Chủ đề C 'Kỹ thuật và ứng dụng AI' bao gồm những nội dung trọng tâm nào?",
                        options: [
                            "Lịch sử các triều đại phong kiến",
                            "Đặc điểm AI, ứng dụng trong học tập/đời sống, công nghệ học máy, dữ liệu và thuật toán",
                            "Quy định xử phạt vi phạm giao thông",
                            "Phương pháp chăm sóc cây xanh"
                        ],
                        correct: 1,
                        explanation: "Chủ đề C tập trung vào công nghệ AI: nhận dạng mẫu, xử lý ngôn ngữ tự nhiên, thị giác máy tính, dữ liệu huấn luyện và ứng dụng thực tiễn."
                    },
                    {
                        question: "Hoạt động học sinh cùng nhau xây dựng ý tưởng ứng dụng AI phân loại rác thải tại trường học thuộc Chủ đề nào?",
                        options: ["Chủ đề A", "Chủ đề B", "Chủ đề C", "Chủ đề D: Thiết kế hệ thống AI (thành tố D1)"],
                        correct: 3,
                        explanation: "Xác định bài toán thực tiễn và hình thành ý tưởng giải pháp sử dụng AI là nội dung của Chủ đề D: Thiết kế hệ thống AI (D1)."
                    }
                ]
            },
            "qd3439_daoduc": {
                name: "Phần 2: Tư duy nhân văn & Đạo đức AI",
                desc: "Nguyên tắc cốt lõi: AI không có cảm xúc thật, phòng tránh thiên vị dữ liệu và liêm chính học thuật.",
                questions: [
                    {
                        question: "Khi tương tác với trợ lý ảo hoặc chatbot AI, học sinh cần nhận thức đúng đắn điều gì về cảm xúc của AI?",
                        options: [
                            "Chatbot thực sự yêu mến và có cảm xúc như một người bạn thân",
                            "AI chỉ mô phỏng phản ứng ngôn ngữ dựa trên dữ liệu lập trình sẵn, AI không có cảm xúc thật",
                            "AI có thể buồn bã khi học sinh tắt máy tính",
                            "AI tự hình thành ý thức và tình cảm sau một thời gian trò chuyện"
                        ],
                        correct: 1,
                        explanation: "Căn cứ Chủ đề A1: AI không có cảm xúc thật, con người không được nhầm lẫn giữa phản ứng mô phỏng của máy móc với cảm xúc con người."
                    },
                    {
                        question: "Hành vi nào sau đây vi phạm nguyên tắc Đạo đức AI (Chủ đề B) trong học tập?",
                        options: [
                            "Dùng AI để dịch một đoạn văn và kiểm tra lại từ vựng",
                            "Sử dụng AI viết hộ toàn bộ bài văn/bài luận rồi nộp như thể chính mình viết mà không ghi nguồn",
                            "Yêu cầu AI giải thích một khái niệm khoa học khó hiểu",
                            "Nhờ AI gợi ý dàn ý cho bài thuyết trình"
                        ],
                        correct: 1,
                        explanation: "Chủ đề B yêu cầu tính liêm chính học thuật: việc dùng AI làm hộ bài tập và giả mạo tác quyền là vi phạm đạo đức số và gian lận học thuật."
                    },
                    {
                        question: "Hiện tượng AI đưa ra câu trả lời nghe rất thuyết phục nhưng thông tin lại hoàn toàn sai sự thật được gọi là gì?",
                        options: ["Siêu trí tuệ", "Ảo giác của AI (AI Hallucination)", "Trí nhớ vĩnh cửu", "Thuật toán tối ưu"],
                        correct: 1,
                        explanation: "Ảo giác của AI (AI Hallucination) là hiện tượng mô hình AI tạo ra thông tin sai lệch nhưng diễn đạt tự tin, đòi hỏi học sinh phải luôn kiểm chứng lại với nguồn sách chính thống."
                    }
                ]
            }
        }
    },

    // -------------------------------------------------------------------------
    // BỘ 3: CÔNG VĂN 5512/BGDĐT & CT GDPT 2018 (NL CHUNG & 5 PHẨM CHẤT)
    // -------------------------------------------------------------------------
    "cv5512": {
        title: "Công văn 5512/BGDĐT & CT GDPT 2018 — NL Chung & Phẩm chất",
        badge: "⭐ 5 Phẩm chất & NL Chung",
        subsections: {
            "cv5512_5pc": {
                name: "Phần 1: 5 Phẩm chất chủ yếu theo CT GDPT 2018",
                desc: "Yêu nước, Nhân ái, Chăm chỉ, Trung thực, Trách nhiệm trong môi trường học đường.",
                questions: [
                    {
                        question: "Chương trình Giáo dục phổ thông 2018 quy định có bao nhiêu Phẩm chất chủ yếu?",
                        options: ["3 phẩm chất", "4 phẩm chất", "5 phẩm chất", "6 phẩm chất"],
                        correct: 2,
                        explanation: "Chương trình GDPT 2018 (Thông tư 32/2018/TT-BGDĐT) xác định 5 phẩm chất chủ yếu: Yêu nước, Nhân ái, Chăm chỉ, Trung thực, Trách nhiệm."
                    },
                    {
                        question: "Học sinh tự giác làm bài kiểm tra, không quay cóp và trung thực ghi lại số liệu thí nghiệm thực tế thể hiện phẩm chất nào?",
                        options: ["Yêu nước", "Nhân ái", "Chăm chỉ", "Trung thực"],
                        correct: 3,
                        explanation: "Thật thà, ngay thẳng, không gian lận trong học tập và thi cử là biểu hiện trực tiếp của phẩm chất Trung thực."
                    },
                    {
                        question: "Có ý thức tắt quạt, tắt đèn khi ra khỏi phòng học và giữ gìn bàn ghế, cơ sở vật chất của nhà trường thể hiện phẩm chất nào?",
                        options: ["Trách nhiệm", "Nhân ái", "Trung thực", "Yêu nước"],
                        correct: 0,
                        explanation: "Bảo vệ của công, giữ gìn tài sản chung và có ý thức tiết kiệm năng lượng là biểu hiện rõ nét của phẩm chất Trách nhiệm."
                    },
                    {
                        question: "Học sinh kiên trì giải bài tập khó, tích cực tìm tòi đọc thêm sách tham khảo để mở rộng kiến thức thể hiện phẩm chất nào?",
                        options: ["Chăm chỉ", "Nhân ái", "Trung thực", "Yêu nước"],
                        correct: 0,
                        explanation: "Chăm học, chăm làm, có tinh thần tự giác, kiên trì vượt khó trong học tập là biểu hiện của phẩm chất Chăm chỉ."
                    },
                    {
                        question: "Biết lắng nghe, thông cảm, sẵn sàng giúp đỡ bạn bè cùng tiến bộ và không có lời nói gây tổn thương bạn thể hiện phẩm chất nào?",
                        options: ["Yêu nước", "Nhân ái", "Trung thực", "Trách nhiệm"],
                        correct: 1,
                        explanation: "Yêu quý mọi người, tôn trọng sự khác biệt, đồng cảm và sẵn sàng chia sẻ, giúp đỡ người khác là biểu hiện của phẩm chất Nhân ái."
                    },
                    {
                        question: "Phẩm chất 'Yêu nước' được bồi dưỡng qua các môn học phổ thông thông qua những hoạt động nào?",
                        options: [
                            "Tìm hiểu về lịch sử, địa lý quê hương; tự hào về văn hóa dân tộc và tham gia bảo vệ môi trường",
                            "Chỉ khi học sinh tham gia duyệt binh",
                            "Chỉ thể hiện khi học sinh đạt giải quốc tế",
                            "Học thuộc lòng bảng cửu chương"
                        ],
                        correct: 0,
                        explanation: "Yêu nước bắt đầu từ tình yêu quê hương, gia đình, thiên nhiên, niềm tự hào về truyền thống dân tộc và ý thức bảo vệ môi trường sống."
                    }
                ]
            },
            "cv5512_nlchung": {
                name: "Phần 2: 3 Năng lực chung theo CT GDPT 2018",
                desc: "Tự chủ - Tự học, Giao tiếp - Hợp tác, Giải quyết vấn đề - Sáng tạo xuyên suốt các môn.",
                questions: [
                    {
                        question: "3 Năng lực chung cốt lõi áp dụng cho tất cả các môn học trong CT GDPT 2018 là gì?",
                        options: [
                            "Tự chủ và tự học; Giao tiếp và hợp tác; Giải quyết vấn đề và sáng tạo",
                            "Năng lực Toán học; Năng lực Ngôn ngữ; Năng lực Tin học",
                            "Năng lực Thể chất; Năng lực Thẩm mỹ; Năng lực Kỹ thuật",
                            "Năng lực Lãnh đạo; Năng lực Thuyết trình; Năng lực Quản lý"
                        ],
                        correct: 0,
                        explanation: "3 năng lực chung cốt lõi theo CT GDPT 2018 là: Tự chủ & Tự học, Giao tiếp & Hợp tác, Giải quyết vấn đề & Sáng tạo."
                    },
                    {
                        question: "Khi học sinh chủ động nghiên cứu SGK, chuẩn bị câu hỏi trước bài học mới mà không cần nhắc nhở, học sinh đang phát triển năng lực gì?",
                        options: ["Giao tiếp và hợp tác", "Tự chủ và tự học", "Giải quyết vấn đề", "Năng lực đặc thù"],
                        correct: 1,
                        explanation: "Hành động chủ động tự giác nghiên cứu tài liệu, tự đặt mục tiêu và thực hiện kế hoạch học tập thuộc năng lực Tự chủ và Tự học."
                    },
                    {
                        question: "Khi nhóm học sinh cùng nhau thảo luận, phân chia nhiệm vụ để làm một bài báo cáo chung môn KHTN/Lịch sử, các em đang rèn luyện năng lực nào?",
                        options: [
                            "Giao tiếp và hợp tác",
                            "Năng lực cá nhân độc lập",
                            "Năng lực phòng thủ",
                            "Năng lực thi đấu thể thao"
                        ],
                        correct: 0,
                        explanation: "Trao đổi, thảo luận, phân công công việc và cùng nhau hoàn thành mục tiêu nhóm là biểu hiện của năng lực Giao tiếp và Hợp tác."
                    }
                ]
            },
            "cv5512_format": {
                name: "Phần 3: Quy tắc soạn Mục tiêu KHBD theo CV 5512 & UNIGO",
                desc: "Quy chuẩn viết Mục 1 Kiến thức, Mục 2 Năng lực và Mục 3 Phẩm chất.",
                questions: [
                    {
                        question: "Quy định nào sau đây là BẮT BUỘC khi viết Mục 1 (Kiến thức) trong KHBD chuẩn UNIGO?",
                        options: [
                            "Bắt đầu bằng các động từ: 'Hiểu được', 'Biết được', 'Nắm được'",
                            "Sử dụng Danh từ / Cụm danh từ trực tiếp; TUYỆT ĐỐI KHÔNG dùng động từ",
                            "Viết một đoạn văn dài miên man không gạch đầu dòng",
                            "Chèn bảng điểm kiểm tra học sinh"
                        ],
                        correct: 1,
                        explanation: "Theo quy chuẩn UNIGO: Mục tiêu Kiến thức dùng Danh từ / Cụm danh từ trực tiếp. Tuyệt đối KHÔNG dùng động từ (hiểu, biết, nhận diện, nêu...)."
                    },
                    {
                        question: "Thứ tự các mục tiêu trong KHBD cấp THCS (Lớp 6-8) được sắp xếp như thế nào?",
                        options: [
                            "1. Phẩm chất -> 2. Năng lực -> 3. Kiến thức",
                            "1. Kiến thức -> 2. Năng lực (2.1 đến 2.4) -> 3. Phẩm chất",
                            "1. Năng lực -> 2. Kiến thức -> 3. Phẩm chất",
                            "Không cần theo thứ tự nào cả"
                        ],
                        correct: 1,
                        explanation: "Thứ tự chuẩn THCS: 1. Kiến thức -> 2. Năng lực (2.1 Đặc thù, 2.2 NLS, 2.3 NL Chung, 2.4 NL AI) -> 3. Phẩm chất."
                    }
                ]
            }
        }
    },

    // -------------------------------------------------------------------------
    // BỘ 4: NĂNG LỰC ĐẶC THÙ (TIN HỌC, ROBOTICS & CÁC MÔN HỌC)
    // -------------------------------------------------------------------------
    "dacthu": {
        title: "Năng lực Đặc thù môn học & Tích hợp liên môn",
        badge: "⚙️ NL Đặc thù",
        subsections: {
            "dacthu_tinhoc": {
                name: "Phần 1: 5 Năng lực đặc thù Tin học (NLa - NLe)",
                desc: "Hệ thống 5 năng lực cốt lõi theo Chương trình GDPT 2018 môn Tin học.",
                questions: [
                    {
                        question: "Mã hiệu nào sau đây đại diện cho năng lực 'Ứng xử phù hợp trong môi trường số' môn Tin học?",
                        options: ["NLa", "NLb", "NLc", "NLe"],
                        correct: 1,
                        explanation: "Quy chuẩn CT GDPT 2018 môn Tin học: NLa (Sử dụng CNTT), NLb (Ứng xử số), NLc (Giải quyết vấn đề với CNTT), NLd (Học & tự học với CNTT), NLe (Hợp tác trong môi trường số)."
                    },
                    {
                        question: "Năng lực NLc trong môn Tin học nhấn mạnh vào kỹ năng nào của học sinh?",
                        options: [
                            "Sử dụng và quản lý các phương tiện CNTT thông thường",
                            "Khám phá tri thức, giải quyết vấn đề với sự trợ giúp của CNTT và truyền thông",
                            "Lắp ráp linh kiện máy tính phần cứng",
                            "Mua bán thiết bị điện tử trên mạng"
                        ],
                        correct: 1,
                        explanation: "NLc = Khám phá tri thức, giải quyết vấn đề với sự trợ giúp của CNTT và truyền thông: dùng máy tính như công cụ tư duy."
                    }
                ]
            },
            "dacthu_robotics": {
                name: "Phần 2: 5 Năng lực môn Robotics UNIGO (NL1 - NL5)",
                desc: "Khung năng lực kỹ thuật và chế tạo robot giáo dục từ NL1 đến NL5.",
                questions: [
                    {
                        question: "Năng lực NL2 trong môn Robotics UNIGO quy định về kỹ năng nào?",
                        options: [
                            "Nhận thức công nghệ",
                            "Lắp ráp và chế tạo mô hình robot (đọc bản vẽ, lắp ráp cơ khí chính xác)",
                            "Lập trình thuật toán",
                            "Thuyết trình sản phẩm"
                        ],
                        correct: 1,
                        explanation: "NL2 = Lắp ráp và chế tạo mô hình robot: kỹ năng cơ khí, đọc sơ đồ và thao tác lắp ráp các khối chi tiết thành sản phẩm robot hoàn chỉnh."
                    },
                    {
                        question: "Kỹ năng thử nghiệm thực tế, phát hiện lỗi phần cứng/phần mềm và tinh chỉnh thông số kỹ thuật (Troubleshooting) là năng lực nào?",
                        options: ["NL1", "NL3", "NL4 (Thử nghiệm, tinh chỉnh và xử lý sự cố)", "NL5"],
                        correct: 2,
                        explanation: "NL4 = Thử nghiệm, tinh chỉnh và xử lý sự cố (Troubleshooting): chạy thử sa bàn, phát hiện lỗi và tối ưu hóa hệ thống robot."
                    }
                ]
            }
        }
    }
};

// Flattened helper for General Practice (toàn diện 20 câu ngẫu nhiên)
function getAllQuestions() {
    const list = [];
    Object.keys(QUIZ_BANK).forEach(docKey => {
        const doc = QUIZ_BANK[docKey];
        Object.keys(doc.subsections).forEach(subKey => {
            const sub = doc.subsections[subKey];
            sub.questions.forEach((q, idx) => {
                list.push({
                    ...q,
                    docKey,
                    docTitle: doc.title,
                    subKey,
                    subName: sub.name,
                    id: `${subKey}_${idx}`
                });
            });
        });
    });
    return list;
}
