/**
 * Nhật ký sử dụng app "Hoàn công CSKVTT" -> Google Sheet.
 * CHỈ GHI THÊM dòng (không có hàm đọc/sửa/xoá). Triển khai: xem HUONG_DAN.md cùng thư mục.
 */
var TEN_SHEET = 'NhatKy';
var COT = ['Thời điểm (máy chủ)', 'Thời điểm (máy người dùng)', 'Họ tên', 'Đơn vị / bộ phận', 'Số điện thoại',
           'Việc', 'Tên file', 'Chi tiết', 'Phiên bản app', 'Mã máy'];
var DAI_TOI_DA = 300;                        // cắt chuỗi quá dài (chống ghi rác)

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!d || typeof d !== 'object' || !d.viec) return traVe_({ok: false, loi: 'thiếu việc'});
    lock.waitLock(10000);
    var sh = sheet_();
    var c = function (v) { v = (v === undefined || v === null) ? '' : String(v); v = v.slice(0, DAI_TOI_DA); return /^[=+\-@]/.test(v) ? "'" + v : v; };
    sh.appendRow([new Date(), c(d.luc), c(d.ten), c(d.don_vi), c(d.sdt), c(d.viec), c(d.file), c(d.chi_tiet), c(d.phien_ban), c(d.may)]);
    return traVe_({ok: true});
  } catch (err) {
    return traVe_({ok: false, loi: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() { return traVe_({ok: true, dich_vu: 'nhat-ky-hoan-cong'}); }   // chỉ để kiểm tra đường dẫn sống

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TEN_SHEET);
  if (!sh) { sh = ss.insertSheet(TEN_SHEET); }
  if (sh.getLastRow() === 0) { sh.appendRow(COT); sh.setFrozenRows(1); sh.getRange(1, 1, 1, COT.length).setFontWeight('bold'); }
  return sh;
}

function traVe_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
