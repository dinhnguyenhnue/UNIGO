/**
 * ============================================================================
 * MÃ.GS CHUẨN HÓA CHO FILE GOOGLE SHEETS "gk1-6-10" (UNIGO TIN HỌC)
 * ============================================================================
 * CÁCH DÁN & CẬP NHẬT TRONG CỬA SỔ APPS SCRIPT "Điểm thi" (1 PHÚT):
 * 1. Trong cửa sổ Apps Script "Điểm thi" (tệp Mã.gs), xóa toàn bộ 23 dòng cũ và dán toàn bộ mã này vào -> Nhấn Ctrl + S (Lưu).
 * 2. (Khuyên dùng) Trên thanh công cụ bên cạnh nút "▷ Chạy", đổi hàm từ "doPost" sang "khoiTaoTrangTinhThu2"
 *    rồi bấm "▷ Chạy" 1 lần -> Ngay lập tức tab thứ 2 "OnTap_HK1" sẽ xuất hiện cạnh "Trang tính1" với đủ 13 cột tiêu đề!
 * 3. Nhấn nút màu xanh "Triển khai" (góc trên bên phải) -> "Quản lý các bản triển khai"
 *    -> Bấm biểu tượng cây bút ✏️ (Chỉnh sửa) -> Ở mục "Phiên bản", chọn "Phiên bản mới" -> Bấm "Triển khai".
 *    (Giữ nguyên 100% đường link Webhook cũ đang dùng trên web!).
 */

var SHEET1_NAME = "Trang tính1";
var SHEET2_NAME = "OnTap_HK1";

var SHEET2_HEADERS = [
  "Thời gian nộp",
  "Họ và tên học sinh",
  "Lớp",
  "Khối",
  "Số máy",
  "Nội dung bài ôn tập",
  "Lần làm thứ",
  "Tổng số lượt đã làm",
  "Điểm số",
  "Số câu đúng",
  "Tổng số câu",
  "Thời gian làm bài",
  "Xếp loại"
];

function getOrCreateSheet2(ss) {
  var sheet2 = ss.getSheetByName(SHEET2_NAME);
  if (!sheet2) {
    var insertIdx = Math.min(1, ss.getSheets().length);
    sheet2 = ss.insertSheet(SHEET2_NAME, insertIdx);
  }
  if (sheet2.getLastRow() === 0) {
    sheet2.appendRow(SHEET2_HEADERS);
    var headerRange = sheet2.getRange(1, 1, 1, SHEET2_HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1d4ed8");
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
    sheet2.setFrozenRows(1);
    sheet2.setColumnWidth(1, 145);
    sheet2.setColumnWidth(2, 185);
    sheet2.setColumnWidth(3, 75);
    sheet2.setColumnWidth(4, 60);
    sheet2.setColumnWidth(5, 85);
    sheet2.setColumnWidth(6, 230);
    sheet2.setColumnWidth(7, 95);
    sheet2.setColumnWidth(8, 130);
    sheet2.setColumnWidth(9, 85);
    sheet2.setColumnWidth(10, 95);
    sheet2.setColumnWidth(11, 95);
    sheet2.setColumnWidth(12, 120);
    sheet2.setColumnWidth(13, 115);
  }
  return sheet2;
}

function normText(s) {
  return String(s || "").trim().toLowerCase();
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(raw);

    var machineVal = data.machineNumber !== undefined ? data.machineNumber : (data.pcNumber || "Máy tự do");
    var scoreVal = data.score !== undefined ? data.score : (data.score10 !== undefined ? data.score10 : 0);

    // 1. NẾU LÀ DỮ LIỆU ÔN TẬP / KIỂM TRA HỌC KÌ 1 -> GHI VÀO TRANG TÍNH THỨ 2 ("OnTap_HK1")
    var isHk1 = (
      data.targetSheet === SHEET2_NAME ||
      data.periodId === "CK1" ||
      data.periodId === "HK1_ROBO"
    );

    if (isHk1) {
      var sheet2 = getOrCreateSheet2(ss);
      var lastRow = sheet2.getLastRow();

      var sName = String(data.studentName || "Học sinh").trim();
      var sClass = String(data.studentClass || ("Lớp " + (data.grade || ""))).trim();
      var actTitle = String(data.activityTitle || "Ôn tập Cuối kì 1").trim();

      // Tự động đếm số lần làm của học sinh ngay trên Trang tính thứ 2 (chính xác kể cả khi đổi máy trong phòng Tin!)
      var sameActivityPrev = 0;
      var totalStudentPrev = 0;
      if (lastRow >= 2) {
        // Đọc cột B (Họ tên), C (Lớp), D (Khối), E (Số máy), F (Nội dung bài ôn tập) -> từ cột 2, rộng 5 cột
        var rows = sheet2.getRange(2, 2, lastRow - 1, 5).getValues();
        var nName = normText(sName);
        var nClass = normText(sClass);
        var nAct = normText(actTitle);
        for (var i = 0; i < rows.length; i++) {
          if (normText(rows[i][0]) === nName && normText(rows[i][1]) === nClass) {
            totalStudentPrev++;
            if (normText(rows[i][4]) === nAct) {
              sameActivityPrev++;
            }
          }
        }
      }

      var attemptNum = Math.max(Number(data.attemptNumber) || 1, sameActivityPrev + 1);
      var totalAttempts = Math.max(Number(data.totalAttemptsAllModes) || 1, totalStudentPrev + 1);

      sheet2.appendRow([
        new Date(),
        sName,
        sClass,
        data.grade || "",
        machineVal,
        actTitle,
        attemptNum,
        totalAttempts,
        Number(scoreVal),
        data.correctCount !== undefined ? data.correctCount : "",
        data.totalQuestions !== undefined ? data.totalQuestions : "",
        data.timeSpent || "",
        data.rank || ""
      ]);

      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", sheet: SHEET2_NAME, attemptNumber: attemptNum }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. DỮ LIỆU CŨ (GK1): GHI VÀO ĐÚNG "Trang tính1" THEO Y HỆT 10 CỘT A -> J HIỆN CÓ!
    var sheet1 = ss.getSheetByName(SHEET1_NAME) || ss.getSheets()[0];
    sheet1.appendRow([
      new Date(),
      data.studentName,
      data.studentClass,
      data.grade,
      machineVal,
      scoreVal,
      data.correctCount,
      data.totalQuestions,
      data.timeSpent,
      data.rank
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", sheet: sheet1.getName() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Chọn hàm "khoiTaoTrangTinhThu2" trên thanh công cụ rồi bấm "▷ Chạy" để tạo ngay tab thứ 2 "OnTap_HK1"
 */
function khoiTaoTrangTinhThu2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet2(ss);
}
