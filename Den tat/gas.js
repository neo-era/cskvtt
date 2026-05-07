// Google Apps Script — Quản lý đèn tắt CSCC
// Deploy: Extensions → Apps Script → Deploy as Web App
// Execute as: Me | Who has access: Anyone

const SHEET_NAME = 'DanhSachDen'; // ← Kiểm tra đúng tên tab Sheet

// Map camelCase JS → tên cột Sheet chính xác
const FIELD_MAP = {
  'id':            'ID',
  'soTru':         'Số trụ',
  'tenTu':         'Tên tủ',
  'lat':           'latitude',
  'latitude':      'latitude',
  'lon':           'lontitude',
  'longitude':     'lontitude',
  'lontitude':     'lontitude',
  'loaiDen':       'Loại đèn',
  'congSuat':      'Công suất',
  'trangThai':     'Trang thai',
  'duong':         'Đường',
  'phuong':        'Phường',
  'ngayPhatHien':  'Ngày phát hiện',
  'nguoiPhatHien': 'Người phát hiện',
  'ngaySua':       'Ngày sửa',
  'nguoiSua':      'Người sửa',
  'vatTuSua':      'Vật tư sửa',
  'hinhAnh':       'HÌnh ảnh',
  'ghiChu':        'Ghi chú',
  'vn2000x':       'VN2000-X',
  'vn2000y':       'VN2000-Y',
};

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// Chuẩn hóa chuỗi để so sánh (NFC + trim + lowercase)
function norm(s) {
  return String(s || '').normalize('NFC').trim().toLowerCase();
}

// Xây dựng index: norm(header) → col index (0-based)
function buildHeaderIndex(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[norm(h)] = i; });
  return idx;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    if (!sheet) throw new Error('Không tìm thấy sheet: ' + SHEET_NAME);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hIdx = buildHeaderIndex(headers);

    if (data.action === 'full_update') {
      const rowNum = findRowNum(sheet, headers, hIdx, data);
      if (rowNum > 0) {
        updateRow(sheet, hIdx, rowNum, data);
      } else {
        appendRow(sheet, headers, hIdx, data);
      }
    } else {
      // GPS-only: chỉ cập nhật lat/lon
      const rowNum = findRowNum(sheet, headers, hIdx, data);
      if (rowNum > 0) {
        const updates = {};
        const latVal = data.lat || data.latitude || '';
        const lonVal = data.lon || data.longitude || data.lontitude || '';
        if (latVal !== '') updates['latitude']  = latVal;
        if (lonVal !== '') updates['lontitude'] = lonVal;
        updateRowFields(sheet, hIdx, rowNum, updates);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Tìm số hàng theo ID (ưu tiên) hoặc Số trụ
function findRowNum(sheet, headers, hIdx, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const idColIdx    = hIdx[norm('ID')];
  const soTruColIdx = hIdx[norm('Số trụ')];

  const searchId    = norm(data.id    || data['ID']    || '');
  const searchSoTru = norm(data.soTru || data['Số trụ'] || '');
  if (!searchId && !searchSoTru) return -1;

  const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < allData.length; i++) {
    const rowId    = idColIdx    !== undefined ? norm(allData[i][idColIdx])    : '';
    const rowSoTru = soTruColIdx !== undefined ? norm(allData[i][soTruColIdx]) : '';
    if ((searchId    && rowId    && rowId    === searchId)    ||
        (searchSoTru && rowSoTru && rowSoTru === searchSoTru)) {
      return i + 2; // i+2: bỏ header (hàng 1) + offset 0-based
    }
  }
  return -1;
}

// Cập nhật toàn bộ fields của 1 hàng
function updateRow(sheet, hIdx, rowNum, data) {
  updateRowFields(sheet, hIdx, rowNum, buildFieldValues(data));
}

// Ghi { 'Tên cột Sheet': value } vào đúng cột bằng hIdx (normalized)
function updateRowFields(sheet, hIdx, rowNum, fieldValues) {
  for (const [header, value] of Object.entries(fieldValues)) {
    const col = hIdx[norm(header)];
    if (col !== undefined && value !== null && value !== undefined) {
      sheet.getRange(rowNum, col + 1).setValue(value);
    }
  }
}

// Thêm hàng mới theo đúng thứ tự cột
function appendRow(sheet, headers, hIdx, data) {
  const fieldValues = buildFieldValues(data);
  const row = headers.map(h => {
    const col = hIdx[norm(h)];
    return fieldValues[h] !== undefined ? fieldValues[h] : '';
  });
  sheet.appendRow(row);
}

// Chuyển payload camelCase → { 'Tên cột Sheet': giá trị }
function buildFieldValues(data) {
  const result = {};
  for (const [key, value] of Object.entries(data)) {
    if (key === 'action') continue;
    // Map camelCase → header, hoặc giữ nguyên nếu đã là tên cột
    const header = FIELD_MAP[key] || key;
    if (value !== undefined && value !== null && value !== '') {
      result[header] = value;
    }
  }
  return result;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Den tat GAS v2' }))
    .setMimeType(ContentService.MimeType.JSON);
}
