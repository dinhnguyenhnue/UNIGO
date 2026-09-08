# Bổ sung Mục 2.4 Năng lực AI — Patch cho SKILL.md tao-khbd

> [!IMPORTANT]
> File này bổ sung cho `tao-khbd/SKILL.md` — Agent PHẢI đọc file này CÙNG VỚI SKILL.md chính.

---

## Thay đổi 1: Bước 3 — Thêm nhóm năng lực thứ 4 (mục 2.4)

**Cấu trúc năng lực trong KHBD giờ gồm 4 nhóm bắt buộc:**

```
2. Năng lực:
   2.1. Năng lực đặc thù (Tin học / Robotics):  ← Như cũ
   2.2. Năng lực số (NLS) — TT 02/2025 – CV 3456:  ← Như cũ
   2.3. Năng lực chung:  ← Như cũ
   2.4. Năng lực Trí tuệ nhân tạo (AI) được phát triển:  ← MỚI
```

### Cách viết mục 2.4:

**BẮT BUỘC đọc:** `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_AI_CV3439.md`
**BẮT BUỘC tra:** `D:\UNIGO\.agents\skills\tao-khbd\references\cv3439_ai_data.json`

**Quy trình:**
1. Xác định **khối lớp** → chọn đúng **Bậc** (Bảng trong KHBD_NANG_LUC_AI_CV3439.md Phần II).
2. Xác định **nhóm bài** → chọn **Chủ đề AI phù hợp** (A/B/C/D) (Bảng Phần IV).
3. Tra YCCD từ `cv3439_ai_data.json`: `data["lop_X"]["themes"]["A/B/C/D"]["yccd"]`.
4. Ghi 2 NL AI items, format:

```
2.4. Năng lực Trí tuệ nhân tạo (AI) được phát triển:
- [Bậc].[Mã thành phần].[Mã cụ thể]: [Biểu hiện cụ thể từ QĐ 3439]
  (Đạt được thông qua Hoạt động X, Hoạt động Y).
```

**Ví dụ Lớp 3:**
```
2.4. Năng lực Trí tuệ nhân tạo (AI) được phát triển:
- 3.A1.1: Nhận biết được rằng AI là công cụ do con người tạo ra để hỗ trợ
  học tập, như trợ lý học tập thông minh, ứng dụng học ngôn ngữ
  (Đạt được thông qua Hoạt động 2, Hoạt động 3).
- 3.C1.1: Nhận biết được một số sản phẩm có sử dụng AI trong đời sống
  (Đạt được thông qua Hoạt động 2).
```

**Ví dụ Lớp 6:**
```
2.4. Năng lực Trí tuệ nhân tạo (AI) được phát triển:
- 6.A1.1: Nhận biết được vai trò của con người trong việc thiết kế, vận hành
  và sử dụng AI; biết rằng con người chịu trách nhiệm với các tác động AI
  (Đạt được thông qua Hoạt động 2).
- 6.C1.1: Giải thích được hai thành phần chính để "dạy" cho AI là Dữ liệu
  và Thuật toán (Đạt được thông qua Hoạt động 2, Hoạt động 3).
```

> [!CAUTION]
> - **CẤM tự bịa YCCD AI** — phải lấy từ cv3439_ai_data.json.
> - **KHÔNG ghi NL AI cho bài Ôn tập / Đánh giá định kỳ** (chỉ cho bài nội dung mới).

---

## Thay đổi 2: Bước 2 — Nguồn SGV từ NotebookLM

**Khi đọc SGK, Agent CÓ THỂ truy vấn SGV (Sách giáo viên) từ NotebookLM:**

- Notebook ID: `f6c754f2-0291-40f3-bdc8-8b8cd37ef396`
- Các source SGV:
  - SGV Lớp 3: `758b8616-8e91-4e0d-9756-0596287a2298`
  - SGV Lớp 4: `878fbee1-8cf2-42a1-9f36-84c135c18994`
  - SGV Lớp 5: `7678b4cb-8096-4338-b858-7d8a25d0db5a`
  - SGV Lớp 6: `d50aac9c-91b3-4ae8-a5ea-a0227761ce3c`
  - SGV Lớp 7: `17590b51-757a-45a0-98a5-ceeffe1a8679`
  - SGV Lớp 8: `aad2b61a-9a4e-4670-8429-95d7009028aa`
- Các source SGK:
  - SGK Lớp 3: `7c79c430-0e40-4a91-92af-53ee80c06644`
  - SGK Lớp 4: `7300e43e-7710-4bbf-8afb-5e74cb4d66dd`
  - SGK Lớp 5: `a9f8bf8e-f6e0-4877-833f-154b17866255`
  - SGK Lớp 6: `7a8ae116-e5cb-474e-a08e-ad4628bd3fd1`
  - SGK Lớp 7: `883a4a8d-0356-4399-93c1-03a77c2cd9e1`
  - SGK Lớp 8: `5ba03c44-1e50-4328-8371-528702b7d96a`

**Cách query:**
```python
notebook_query(
    notebook_id="f6c754f2-0291-40f3-bdc8-8b8cd37ef396",
    query="SGV Tin học [Lớp]: Bài [X] - Mục tiêu, Hoạt động dạy học, Chốt kiến thức",
    source_ids=["<SGV_source_id>", "<SGK_source_id>"]
)
```

**KHBD tham khảo (KHÔNG tiên quyết, CHỈ tham khảo):**
- `D:\UNIGO\Phân phối chương trình\Tin học\KHBD TIN HỌC 3,4,5 KNTT (1)\`

---

## Thay đổi 3: Bước 2 — Quy trình đọc SGK chi tiết

Khi đọc SGK, Agent PHẢI xác định chính xác:
1. **Mục tiêu bài học** (khung "Sau bài học này em sẽ")
2. **Tiêu đề các mục** (Mục 1, Mục 2, Mục 3...)
3. **Hoạt động** (Hoạt động 1, 2, 3... với câu hỏi dẫn dắt)
4. **Hộp kiến thức** (phần chốt kiến thức / ghi nhớ — QUAN TRỌNG NHẤT)
5. **Câu hỏi củng cố** (ngay sau mỗi mục kiến thức)
6. **Luyện tập** (câu hỏi, bài tập cuối bài)
7. **Vận dụng** (câu hỏi thực tế, trò chơi)

**Tương tác với user:** Nếu SGK dạng scan/ảnh khó đọc, HỎI user cung cấp ảnh.
**Ghi tham chiếu:** Mỗi hoạt động trong KHBD PHẢI ghi rõ `(Trang SGK: X)`.

---

## Thay đổi 4: Bước 5 — Checklist mở rộng (17 tiêu chí)

Checklist đã được cập nhật lên 17 tiêu chí (thêm #6 NL AI, #11 Chốt KT, #12 Trang SGK).

**Quy trình lặp (Iterative Loop):**
1. Nếu FAIL → Quay lại Bước 4 sửa.
2. Sau xuất KHBD → Chuyển sang tạo Slide.
3. Sau tạo Slide → Kiểm tra khớp KHBD ↔ Slide.
4. Nếu không khớp → Quay lại sửa cho đến khi đồng bộ.

---

## Thay đổi 5: Tài liệu tham chiếu mới

| Tài liệu | Đường dẫn |
|:---|:---|
| **NL AI (QĐ 3439)** | `D:\UNIGO\.agents\skills\tao-khbd\references\KHBD_NANG_LUC_AI_CV3439.md` |
| **Dữ liệu AI JSON** | `D:\UNIGO\.agents\skills\tao-khbd\references\cv3439_ai_data.json` |
| **CV 3439 DOCX đã OCR** | `D:\UNIGO\.agents\skills\tao-khbd\references\cv3439_ai_curriculum.docx` |
| **CV 3439 Raw Text** | `D:\UNIGO\.agents\skills\tao-khbd\references\cv3439_raw_text.json` |
