---
name: nang-luc-so
description: >
  Viết Năng lực số (NLS) theo mã chuẩn CB (Chuẩn Bậc) cho KHBD Tin học & Robotics.
  Tuân thủ Thông tư 02/2025 – CV 3456/BGDĐT. Áp dụng khi cần viết/cập nhật
  phần 2.2. Năng lực số trong bất kỳ KHBD nào. Skill tự động xác định Bậc theo lớp,
  chọn Miền NLS phù hợp nội dung bài, tra descriptor từ CV 3456, và viết mã CB.
---

# Skill Năng lực số (NLS) — Format mã Chuẩn Bậc (CB)

## I. Format mã NLS chuẩn

### Cú pháp: `[Thành tố].CB[Bậc][letter]:`

```
[Mã thành tố].CB[Số bậc][a/b/c]: [Descriptor từ CV 3456]. [Bổ sung ngữ cảnh bài học].
(Đạt được thông qua Hoạt động X, Hoạt động Y)
```

### Giải thích các thành phần:

| Phần | Ý nghĩa | Ví dụ |
|:-----|:---------|:------|
| `4.3` | Mã thành tố (Miền.Thành tố) — xem Bảng 24 thành tố bên dưới | `4.3` = Bảo vệ sức khỏe và an sinh số |
| `CB` | Viết tắt cố định "Chuẩn Bậc" | `CB` |
| `2` | Số Bậc (1-5) — xem Bảng Bậc theo lớp | `2` = Bậc 2 (Lớp 4-5) |
| `a` | Thứ tự descriptor trong bậc đó | `a` = bullet thứ 1, `b` = bullet thứ 2, ... |

### Ví dụ minh họa:

```
2.2. Năng lực số (Thông tư 02/2025 – CV 3456):
- 5.2.CB3a: Chỉ ra được những nhu cầu được xác định rõ ràng và thường xuyên, 
  và chọn được các công cụ số thông thường để giải quyết (Lắp ráp robot điều khiển).
  (Đạt được thông qua Hoạt động 2, Hoạt động 3)
- 4.3.CB3a: Giải thích được cách tránh các rủi ro về sức khỏe và mối đe dọa đến an sinh
  khi sử dụng thiết bị điện tử và robot (Tư thế ngồi, an toàn điện).
  (Đạt được thông qua Hoạt động 1, Hoạt động 4)
```

---

## II. Bảng Bậc theo Khối lớp

| Khối lớp | Bậc | Tên Bậc | Ký hiệu CB |
|:---------|:----|:--------|:------------|
| Tiền TH + Lớp 1, 2, 3 | **Bậc 1** | Cơ bản 1 | `CB1` |
| Lớp 4, 5 | **Bậc 2** | Cơ bản 2 | `CB2` |
| Lớp 6, 7 | **Bậc 3** | Trung cấp 1 | `CB3` |
| Lớp 8, 9 | **Bậc 4** | Trung cấp 2 | `CB4` |
| Lớp 10-12 | **Bậc 5** | Nâng cao 1 | `CB5` |

> [!CAUTION]
> **Sai phổ biến:** Lớp 6-7 = **CB3** (KHÔNG phải CB2). Lớp 8-9 = **CB4** (KHÔNG phải CB3).

---

## III. 24 Thành tố NLS (6 Miền)

| Mã | Tên thành tố | Miền |
|:---|:-------------|:-----|
| **1.1** | Duyệt, tìm kiếm và lọc dữ liệu, thông tin và nội dung số | I. Khai thác dữ liệu |
| **1.2** | Đánh giá dữ liệu, thông tin và nội dung số | I |
| **1.3** | Quản lý dữ liệu, thông tin và nội dung số | I |
| **2.1** | Tương tác thông qua công nghệ số | II. Giao tiếp & Hợp tác |
| **2.2** | Chia sẻ thông tin và nội dung thông qua công nghệ số | II |
| **2.3** | Sử dụng công nghệ số để thực hiện trách nhiệm công dân | II |
| **2.4** | Hợp tác thông qua công nghệ số | II |
| **2.5** | Quy tắc ứng xử trên mạng | II |
| **2.6** | Quản lý danh tính số | II |
| **3.1** | Phát triển nội dung số | III. Sáng tạo nội dung số |
| **3.2** | Tích hợp và tạo lập lại nội dung số | III |
| **3.3** | Thực thi bản quyền và giấy phép | III |
| **3.4** | Lập trình | III |
| **4.1** | Bảo vệ thiết bị | IV. An toàn |
| **4.2** | Bảo vệ dữ liệu cá nhân và quyền riêng tư | IV |
| **4.3** | Bảo vệ sức khỏe và an sinh số | IV |
| **4.4** | Bảo vệ môi trường | IV |
| **5.1** | Giải quyết các vấn đề kỹ thuật | V. Giải quyết vấn đề |
| **5.2** | Xác định nhu cầu và giải pháp công nghệ | V |
| **5.3** | Sử dụng sáng tạo công nghệ số | V |
| **5.4** | Xác định các vấn đề cần cải thiện về NLS | V |
| **6.1** | Hiểu biết về trí tuệ nhân tạo | VI. Ứng dụng AI |
| **6.2** | Sử dụng trí tuệ nhân tạo | VI |
| **6.3** | Đánh giá trí tuệ nhân tạo | VI |

---

## IV. Bảng Mapping nội dung bài → Miền NLS

> [!IMPORTANT]
> Agent PHẢI đọc nội dung bài học trước, xác định chủ đề chính, rồi chọn Miền NLS phù hợp. KHÔNG được dùng cùng 1 bộ NLS cho mọi bài.

| Nhóm bài | Miền chính (Primary) | Miền phụ (Secondary) |
|:---------|:--------------------|:--------------------|
| Làm quen máy tính, thiết bị | **5.1** (GQVĐ kỹ thuật) | **4.1** (Bảo vệ thiết bị) |
| Lập trình, thuật toán (Scratch, Python) | **3.4** (Lập trình) | **5.3** (Sáng tạo CN số) |
| Internet, mạng, tìm kiếm | **1.1** (Duyệt, tìm kiếm) | **2.1** (Tương tác CN số) |
| An toàn, đạo đức, văn hóa số | **4.2** (Bảo vệ dữ liệu) | **2.5** (Quy tắc ứng xử) |
| Soạn thảo văn bản, trình chiếu | **3.1** (Phát triển nội dung số) | **1.3** (Quản lý dữ liệu) |
| Bảng tính, xử lý dữ liệu | **1.3** (Quản lý dữ liệu) | **3.1** (Phát triển nội dung số) |
| **Robotics** (động cơ, sensor, lắp ráp) | **5.2** (Nhu cầu & giải pháp CN) | **3.4** (Lập trình) |
| AI / Trí tuệ nhân tạo | **6.1** (Hiểu biết về AI) | **6.2** (Sử dụng AI) |
| Sáng tạo nội dung (vẽ, multimedia) | **3.1** (Phát triển nội dung số) | **3.2** (Tích hợp nội dung) |
| Quản lý file, thư mục | **1.3** (Quản lý dữ liệu) | **5.1** (GQVĐ kỹ thuật) |
| Tư duy logic, quy luật (unplugged) | **3.4** (Lập trình) | **5.3** (Sáng tạo CN số) |

---

## V. Quy trình viết NLS cho 1 bài (BẮT BUỘC)

### Bước 1: Xác định Bậc
```
Lớp → Bậc → Ký hiệu CB
Ví dụ: Lớp 6 → Bậc 3 → CB3
```

### Bước 2: Đọc nội dung bài và xác định chủ đề
- Đọc SGK/giáo án → Xác định bài thuộc nhóm nào (Bảng Phần IV)
- Ví dụ: Bài "Thư mục và tệp tin" → Nhóm "Quản lý file, thư mục"

### Bước 3: Chọn 2 Miền NLS (Primary + Secondary)
- Tra Bảng Mapping (Phần IV) theo nhóm bài
- Ví dụ: Quản lý file → Primary: **1.3**, Secondary: **5.1**

### Bước 4: Tra descriptor từ JSON
```python
import json
with open(r'D:\UNIGO\.agents\skills\tao-khbd\references\cv3456_full_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Key mapping cho Bậc
bac_key = {'CB1': 'L1-3', 'CB2': 'L4-5', 'CB3': 'L6-7', 'CB4': 'L8-9', 'CB5': 'L10-12'}

# Ví dụ: Thành tố 1.3, Bậc 3 (Lớp 6-7)
thanh_to = '1.3. Quản lý dữ liệu, thông tin và nội dung số'
descriptor_text = data[thanh_to]['descriptors']['L6-7']
# Tách thành các bullet (a, b, c...)
bullets = [b.strip() for b in descriptor_text.split('\n') if b.strip().startswith('-')]
# bullets[0] → letter 'a', bullets[1] → letter 'b', ...
```

### Bước 5: Viết mã CB
```
- [Thành tố].CB[Bậc][letter]: [Descriptor đúng]. [Bổ sung ngữ cảnh bài].
  (Đạt được thông qua Hoạt động X, Hoạt động Y)
```

Ví dụ hoàn chỉnh:
```
- 1.3.CB3a: Lựa chọn được dữ liệu, thông tin và nội dung để tổ chức, lưu trữ 
  và truy xuất chúng một cách thường xuyên trong môi trường số 
  (Tạo và quản lý thư mục/tệp tin trên máy tính).
  (Đạt được thông qua Hoạt động 2, Hoạt động 3)
- 5.1.CB3a: Chỉ ra được và giải quyết được các vấn đề kỹ thuật thường xuyên 
  khi vận hành thiết bị (Xử lý lỗi khi tạo/xóa/đổi tên thư mục).
  (Đạt được thông qua Hoạt động 3)
```

---

## VI. Checklist viết NLS

1. ✅ Bậc đúng theo lớp (Bảng Phần II)
2. ✅ Chọn 2 Miền phù hợp nội dung bài (Bảng Phần IV)
3. ✅ Tra descriptor từ `cv3456_full_data.json` đúng thành tố + bậc
4. ✅ Viết 2 NLS items (primary + secondary)
5. ✅ Format đúng mã `X.Y.CBZl:`
6. ✅ Bổ sung ngữ cảnh bài học trong ngoặc đơn
7. ✅ Gắn mốc Hoạt động (Đạt được thông qua HĐ X, HĐ Y)
8. ✅ KHÔNG lặp nội dung giữa 2 items
9. ✅ KHÔNG tự bịa descriptor — lấy từ JSON
10. ✅ KHÔNG dùng cùng bộ NLS cho mọi bài

---

## VII. Dữ liệu tham chiếu

| Tài liệu | Đường dẫn |
|:----------|:----------|
| Dữ liệu NLS đầy đủ (JSON) | `D:\UNIGO\.agents\skills\tao-khbd\references\cv3456_full_data.json` |
| Hướng dẫn NLS chi tiết | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_SO_CV3456.md` |
| CV 3456 gốc | `D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\3456-VV_huong_dan_trien_khai_Khung_nang_luc_so_cho_HS_885ca.docx` |
