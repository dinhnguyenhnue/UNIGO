# Quy tắc Ngày soạn / Ngày dạy & Tên lớp trong KHBD

> Tài liệu này áp dụng cho TẤT CẢ các KHBD (Kế hoạch bài dạy) tạo ra.

---

## 1. Quy tắc Ngày soạn / Ngày dạy

### Nguyên tắc chung
- **Ngày dạy**: Chiếu theo **Lịch báo giảng (LBG)**, lấy đúng ngày thực tế dạy lớp đó trong tuần.
- **Ngày soạn**: Luôn là **Thứ 7 (Saturday) của tuần trước** tuần dạy.

### Công thức tính
```python
from datetime import date, timedelta

TUAN_01_START = date(2026, 8, 3)  # Thứ Hai đầu tiên của năm học

def compute_dates(tuan_so, day_of_week):
    """
    tuan_so: số tuần (1, 2, 3, ...)
    day_of_week: 0=Thứ Hai, 1=Thứ Ba, ..., 4=Thứ Sáu
    Returns: (ngay_soan, ngay_day) dạng 'DD/MM/YYYY'
    """
    week_start = TUAN_01_START + timedelta(weeks=tuan_so - 1)
    ngay_day = week_start + timedelta(days=day_of_week)
    ngay_soan = week_start - timedelta(days=2)  # Saturday tuần trước
    return ngay_soan.strftime('%d/%m/%Y'), ngay_day.strftime('%d/%m/%Y')
```

### Ví dụ minh họa
| Tuần | Thứ Hai | Ngày soạn (Thứ 7 trước) | Lớp 5 dạy Thứ Ba | Lớp 8 dạy Thứ Sáu |
|:-----|:--------|:-------------------------|:------------------|:-------------------|
| 1    | 03/08   | 01/08/2026               | 04/08/2026        | 07/08/2026         |
| 2    | 10/08   | 08/08/2026               | 11/08/2026        | 14/08/2026         |
| 3    | 17/08   | 15/08/2026               | 18/08/2026        | 21/08/2026         |

---

## 2. Mapping Lớp → Tên lớp & Ngày dạy trong tuần (Tin học & Robotics)

### Môn Tin học
| Khối lớp | Tên lớp (ghi trong KHBD) | Ngày dạy trong tuần | day_of_week |
|:---------|:--------------------------|:---------------------|:------------|
| Tiền TH  | **TT3**                   | Thứ Năm              | 3           |
| Lớp 1    | **1A1**                   | Thứ Hai              | 0           |
| Lớp 2    | **2A1**                   | Thứ Ba               | 1           |
| Lớp 3    | **3A1**                   | Thứ Năm              | 3           |
| Lớp 4    | **4C1**                   | Thứ Tư               | 2           |
| Lớp 5    | **5C1**                   | Thứ Ba               | 1           |
| Lớp 6    | **6A1**                   | Thứ Sáu              | 4           |
| Lớp 7    | **7A1**                   | Thứ Ba               | 1           |
| Lớp 8    | **8A1**                   | Thứ Sáu              | 4           |

### Môn Robotics
| Khối lớp | Tên lớp (ghi trong KHBD) | Ngày dạy trong tuần | day_of_week |
|:---------|:--------------------------|:---------------------|:------------|
| Lớp 1    | **1A1**                   | Thứ Năm              | 3           |
| Lớp 2    | **2A1**                   | Thứ Tư               | 2           |
| Lớp 3    | **3A1**                   | Thứ Tư               | 2           |
| Lớp 4    | **4C1**                   | Thứ Sáu              | 4           |
| Lớp 5    | **5C1**                   | Thứ Ba (Tuần lẻ)     | 1           |
| Lớp 6    | **6A1**                   | Thứ Sáu (Tuần lẻ)    | 4           |
| Lớp 7    | **7A1**                   | Thứ Ba (Tuần lẻ)     | 1           |
| Lớp 8    | **8A1**                   | Thứ Sáu (Tuần lẻ)    | 4           |

> **LƯU Ý:**
> - Tiền TH bao gồm 3 lớp (TT3, TTH 1, TTH2) — cùng dạy Thứ Năm.
> - Lớp 1 (1A1, 1C1): Dùng **1A1** làm đại diện KHBD.
> - Lớp 2 (2A1, 2C1): Dùng **2A1** làm đại diện KHBD.
> - Lớp 3 (3A1, 3C1): Dùng **3A1** làm đại diện KHBD.
> - Lớp 4: **4C1**.
> - Lớp 5: **5C1**.
> - Lớp 6, 7, 8 thuộc cấp **THCS**, tên lớp format `XA1` (6A1, 7A1, 8A1).

---

## 3. Format hiển thị trong file KHBD

### Cấp THCS (Lớp 6, 7, 8)
Table[0] Row[1] Cell[1] chỉ có **2 dòng** (KHÔNG có dòng "Lớp" thừa):
```
Ngày soạn: DD/MM/YYYY   Ngày dạy: DD/MM/YYYY
Lớp: 6A1
```

Paragraph tiêu đề bài dạy:
```
Môn học: Tin học   Lớp: 6A1   Thời lượng: 1 tiết (45 phút)
```

### Cấp Tiểu học (Lớp 5)
Paragraph P[0]:
```
Ngày soạn: DD/MM/YYYY   Ngày dạy: DD/MM/YYYY
```

---

## 4. Áp dụng trong code

Constant `LOP_SCHEDULE` trong `generate_khbd_all.py`:
```python
LOP_SCHEDULE = {
    'TTH':  (3, 'TT3'),    # Thứ Năm
    '1':  (0, '1A1'),       # Thứ Hai
    '2':  (1, '2A1'),       # Thứ Ba
    '3':  (3, '3A1'),       # Thứ Năm
    '4':  (2, '4C1'),       # Thứ Tư
    '5':  (1, '5C1'),       # Thứ Ba
    '6':  (4, '6A1'),       # Thứ Sáu
    '7':  (1, '7A1'),       # Thứ Ba
    '8':  (4, '8A1'),       # Thứ Sáu
}
```

Khi gọi hàm tạo KHBD cho Lớp 5-8, luôn truyền:
- `ngay_soan` và `ngay_day` tính từ `compute_dates(tuan_so, day_of_week)`
- `lop` = tên lớp (VD: `'6A1'`, `'7A1'`, `'8A1'`, `'5C1'`)
