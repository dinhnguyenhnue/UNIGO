# Hướng dẫn tra cứu YCCĐ từ KHDH khi soạn KHBD

> File này hướng dẫn agent cách tra cứu và áp dụng **Yêu cầu cần đạt (YCCĐ)** từ Kế hoạch dạy học (KHDH)
> khi soạn Kế hoạch bài dạy (KHBD).

---

## 1. Nguồn dữ liệu

**File JSON:** `D:\UNIGO\.agents\skills\tao-khbd\references\yccd_khdh_data.json`

**Cấu trúc:**
```json
{
  "Tin học": {
    "Tiền tiểu học": [...],
    "Lớp 1": [...], "Lớp 2": [...], ..., "Lớp 8": [...]
  },
  "Robotics": {
    "Lớp 1": [...], ..., "Lớp 8": [...]
  }
}
```

**Mỗi entry có 2 loại:**
- `{"type": "chu_de", "ten_chu_de": "Chủ đề 1: Máy tính và cộng đồng"}` — Dòng tiêu đề chủ đề (THCS)
- `{"type": "bai_hoc", "stt": "1", "bai": "Bài 1. Thông tin và dữ liệu", "so_tiet": "1", "ppct": "1", "yccd": "- Nhận biết được..."}` — Dòng bài học

---

## 2. Quy trình tra cứu (BẮT BUỘC trong Bước 3 của SKILL.md)

### Bước 3.0: Tra cứu YCCĐ từ KHDH

```python
# Pseudocode — Agent thực hiện trong đầu
data = load_json("yccd_khdh_data.json")
lessons = data[mon_hoc][ten_lop]  # VD: data["Tin học"]["Lớp 3"]

# Tìm bài khớp
target = None
for entry in lessons:
    if entry["type"] == "bai_hoc" and tên_bài_KHBD in entry["bai"]:
        target = entry
        break

yccd_text = target["yccd"]  # Đây là YCCĐ chính thức
```

### Quy tắc matching tên bài:
- So khớp **tên bài trong KHBD** với trường `"bai"` trong JSON
- Bỏ qua sự khác nhau về dấu `:` vs `.` (VD: `Bài 1:` vs `Bài 1.`)
- Bỏ qua khoảng trắng thừa
- Nếu KHBD yêu cầu "Tiết 0: Định hướng môn học" → tìm entry có `"bai"` chứa "Định hướng"

---

## 3. Áp dụng YCCĐ theo cấp học

### 3.1. Tiểu học (Tiền TH + Lớp 1-5)

YCCĐ từ KHDH được chèn **nguyên văn** vào phần:

```
I. YÊU CẦU CẦN ĐẠT:
   - Sau bài học này em sẽ:
   + [YCCĐ dòng 1 từ KHDH]
   + [YCCĐ dòng 2 từ KHDH]
   + [YCCĐ dòng 3 từ KHDH]
```

**Quy tắc:**
- Mỗi mục YCCĐ bắt đầu bằng dấu `-` trong JSON → chuyển thành dấu `+` trong KHBD
- Tách từng mục YCCĐ theo dấu `- ` (gạch ngang + khoảng trắng) hoặc dấu `. ` + viết hoa
- GIỮ NGUYÊN nội dung, KHÔNG viết lại, KHÔNG rút gọn, KHÔNG thêm bớt
- Nếu YCCĐ dài hơn 3 dòng → vẫn giữ nguyên, KHÔNG cắt bớt

**Ví dụ:**
- KHDH ghi: `"- Nêu ví dụ đơn giản minh họa vai trò của thông tin trong việc ra quyết định.  - Nhận biết được thông tin và quyết định trong ví dụ cụ thể."`
- KHBD viết:
  ```
  - Sau bài học này em sẽ:
  + Nêu ví dụ đơn giản minh họa vai trò của thông tin trong việc ra quyết định.
  + Nhận biết được thông tin và quyết định trong ví dụ cụ thể.
  ```

### 3.2. THCS (Lớp 6-8)

YCCĐ từ KHDH được **chuyển đổi thành danh từ/cụm danh từ** cho mục:

```
I. Mục tiêu
   1. Kiến thức:
   - [YCCĐ chuyển đổi thành danh từ]
   - [...]
```

**Quy tắc chuyển đổi YCCĐ → Mục tiêu Kiến thức:**

| YCCĐ gốc (động từ) | Mục tiêu Kiến thức (danh từ) |
|:---|:---|
| Nhận biết được sự khác nhau giữa thông tin và dữ liệu | Sự khác nhau giữa thông tin và dữ liệu |
| Phân biệt được thông tin và vật mang thông tin | Cách phân biệt thông tin và vật mang thông tin |
| Giải thích được việc có thể biểu diễn thông tin chỉ với hai kí hiệu 0 và 1 | Nguyên lý biểu diễn thông tin bằng hai kí hiệu 0 và 1 |
| Nêu được các bước cơ bản trong xử lí thông tin | Các bước cơ bản trong xử lí thông tin |
| Biết bit là đơn vị đo thông tin | Đơn vị đo thông tin: bit |

**Nguyên tắc chuyển đổi:**
1. Bỏ động từ mở đầu (nhận biết, nêu, giải thích, phân biệt, biết, hiểu, trình bày...)
2. Chuyển thành cụm danh từ bắt đầu bằng danh từ chính
3. KHÔNG dùng cấu trúc "Sự hiểu biết về...", "Khả năng nhận diện..."
4. Mỗi mục `-` xuống dòng riêng

---

## 4. Fallback khi không tìm thấy YCCĐ

Nếu bài cần soạn KHÔNG có trong JSON (VD: bài mới thêm, bài đặc biệt):

1. **Ưu tiên 1:** Đọc SGK tương ứng tại `D:\UNIGO\SGK\Lớp_{X}\`
2. **Ưu tiên 2:** Đọc KHDH tổng hợp (file .docx) thay vì JSON
3. **Ưu tiên 3:** Tự soạn YCCĐ dựa trên nội dung bài, nhưng phải GHI CHÚ rõ ràng cho user rằng YCCĐ chưa có trong KHDH

---

---

## 6. Tra cứu YCCĐ & Năng lực cho môn ROBOTICS

### 6.1. Nguồn dữ liệu Robotics
1. **Dữ liệu số hóa:** `D:\UNIGO\.agents\skills\tao-khbd\references\robotics_khung_ct_data.json`
2. **Khung chương trình chuẩn:** `D:\UNIGO\Phân phối chương trình\Robotics\KHUNG CHƯƠNG TRÌNH ROBOTICS TIỂU HỌC & THCS UNIGO.docx` (.pdf)
3. **Giáo trình chuyên môn theo khối:**
   - Khối 1, 2: `Giáo trình - OLLO Initiate.docx` (.pdf) — 32 bài
   - Khối 3, 4: `Giáo trình - OLLO Kinder.docx` (.pdf) — 32 bài
   - Khối 5, 6, 7, 8: `Giáo trình - OLLO Excel 1.docx` (.pdf) — 16 bài
   - Nâng cao: `Giáo trình - OLLO Spark.docx` (.pdf) — 16 bài
4. **Quy chuẩn năng lực đặc thù Robotics (NL1 - NL5):**
   - Đọc: `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_DAC_THU_ROBOTICS.md`

### 6.2. Cấu trúc Mục tiêu bài dạy môn Robotics
- **Kiến thức:**
  - Tiểu học: Giữ nguyên các gạch đầu dòng YCCĐ từ giáo trình/khung CT.
  - THCS: Chuyển đổi thành cụm danh từ (khái niệm động cơ, cảm biến, cơ cấu liên kết, v.v.).
- **Năng lực đặc thù (2.1):** Sử dụng các mã **NL1 - NL5** kèm tên đầy đủ:
  - `NL1 (Nhận thức công nghệ)`
  - `NL2 (Sử dụng công nghệ)`
  - `NL3 (Thiết kế kĩ thuật)`
  - `NL4 (Đánh giá công nghệ)`
  - `NL5 (Giao tiếp công nghệ)`
- **Năng lực số (2.2):** Tra đúng Bậc từ `cv3456_full_data.json`.
- **Năng lực chung (2.3):** Tra từ `KHBD_NANG_LUC_CHUNG.md`.

