// Google Apps Script — Quản lý dữ liệu khảo sát chiếu sáng (index.html)
// ══════════════════════════════════════════════════════════════════════════
// HƯỚNG DẪN TRIỂN KHAI:
//   1. Mở Google Sheet mới (đặt tên tùy ý, khác "khaosat")
//   2. Tạo tab đầu tiên, đặt tên: DanhSachTru
//   3. Hàng 1 phải là header đúng thứ tự (xem HEADER bên dưới)
//   4. Vào Extensions → Apps Script → dán toàn bộ file này vào
//   5. Deploy → New deployment → Web App:
//        Execute as: Me
//        Who has access: Anyone
//   6. Copy URL → dán vào KHAOSAT_GAS_URL trong index.html
//   7. Publish Google Sheet dưới dạng CSV:
//        File → Share → Publish to web → Sheet: DanhSachTru → CSV → Publish
//   8. Copy URL CSV → dán vào KHAOSAT_CSV_URL trong index.html
//   9. Vào Script Properties (⚙️ Project Settings → Script Properties):
//        Key: GITHUB_TOKEN   Value: ghp_xxxxxxxxxxxx (PAT với scope contents:write)
// ══════════════════════════════════════════════════════════════════════════

const SHEET_NAME = 'DanhSachTru'; // ← Tên tab trong Google Sheet

// Header khớp chính xác với Google Sheet thực tế
const HEADER = [
  'ID', 'Tên trụ', 'Lat', 'Lon', 'Ghi chú', 'Người KS',
  'Loại', 'Tủ điều khiển', 'Loại trụ', 'Loại cần', 'Loại đèn',
  'Ảnh', 'Thời gian cập nhật', 'Marker gốc', 'Khoảng cách (m)'
];

// Map key payload JS → tên cột trong Sheet
const FIELD_MAP = {
  'id':          'ID',
  'tenTru':      'Tên trụ',
  'lat':         'Lat',
  'lon':         'Lon',
  'ghiChu':      'Ghi chú',
  'nguoiKS':     'Người KS',
  'loai':        'Loại',
  'tuDieuKhien': 'Tủ điều khiển',
  'loaiTru':     'Loại trụ',
  'loaiCan':     'Loại cần',
  'loaiDen':     'Loại đèn',
  'hinhAnh':     'Ảnh',
  'capNhat':     'Thời gian cập nhật',
  'markerGoc':   'Marker gốc',
  'khoangCach':  'Khoảng cách (m)',
};

// ── UTILS ──────────────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

// Tự động tạo header nếu trống, hoặc bổ sung cột còn thiếu ở cuối
function ensureHeader() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    // Sheet hoàn toàn trống — ghi toàn bộ header
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    _formatHeader(sheet, 1, HEADER.length);
    sheet.setFrozenRows(1);
    return;
  }

  // Đọc header hiện tại
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // Tìm cột còn thiếu (so sánh theo tên chuẩn hóa)
  const existingNorm = existing.map(h => norm(h));
  const missing = HEADER.filter(h => !existingNorm.includes(norm(h)));

  if (missing.length > 0) {
    const startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    _formatHeader(sheet, startCol, missing.length);
    sheet.autoResizeColumns(startCol, missing.length);
  }
}

function _formatHeader(sheet, startCol, numCols) {
  sheet.getRange(1, startCol, 1, numCols)
       .setFontWeight('bold')
       .setBackground('#cfe2f3')
       .setHorizontalAlignment('center')
       .setWrap(false);
}

function norm(s) {
  return String(s || '').normalize('NFC').trim().toLowerCase();
}

function buildHeaderIndex(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[norm(h)] = i; });
  return idx;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── UPLOAD ẢNH LÊN GITHUB (lưu vào thư mục images/ gốc của repo) ──────────

function handleImageUpload(imageBase64, soTru, ext) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    return jsonResponse({ status: 'error', message: 'GITHUB_TOKEN chưa được cài trong Script Properties.' });
  }

  const ts = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH-mm-ss") + 'Z';
  const safeName = (soTru || 'img').replace(/[^a-zA-Z0-9]/g, '-').toLowerCase();
  const fileName  = safeName + '-' + ts + '.' + (ext || 'jpg');
  const filePath  = 'images/' + fileName; // ← lưu vào thư mục images/ gốc

  const apiUrl = 'https://api.github.com/repos/neo-era/cskvtt/contents/'
    + filePath.split('/').map(encodeURIComponent).join('/');

  const res = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      message: 'Upload ảnh khảo sát: ' + fileName,
      content: imageBase64,
      branch: 'sub1'
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 201) {
    return jsonResponse({ status: 'error', message: 'GitHub API lỗi ' + code + ': ' + res.getContentText().slice(0, 300) });
  }
  return jsonResponse({ status: 'ok', path: 'images/' + fileName });
}

// ── MAIN HANDLER ───────────────────────────────────────────────────────────

// ── ĐĂNG NHẬP ──────────────────────────────────────────────────────────────

function handleLogin(username, password) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TaiKhoan');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet TaiKhoan chưa được tạo' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'error', message: 'Sai tên đăng nhập hoặc mật khẩu' });

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  // Cột: A=tenDangNhap, B=matKhau, C=hoTen, D=vaiTro
  for (const row of rows) {
    const dn = String(row[0] || '').trim();
    const mk = String(row[1] || '').trim();
    if (dn.toLowerCase() === username.toLowerCase() && mk === password) {
      return jsonResponse({
        status: 'ok',
        user: {
          username:    dn,
          displayName: String(row[2] || dn).trim(),
          role:        String(row[3] || 'user').trim()
        }
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'Sai tên đăng nhập hoặc mật khẩu' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'login') {
      return handleLogin((data.username || '').trim(), data.password || '');
    }

    ensureHeader();

    if (data.action === 'upload_image') {
      return handleImageUpload(data.imageBase64 || '', data.soTru || '', data.ext || 'jpg');
    }

    if (data.action === 'full_update') {
      const sheet = getSheet();
      if (!sheet) throw new Error('Không tìm thấy sheet: ' + SHEET_NAME);

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const hIdx    = buildHeaderIndex(headers);

      const rowNum = findRowNum(sheet, headers, hIdx, data);
      if (rowNum > 0) {
        updateRow(sheet, hIdx, rowNum, data);
      } else {
        appendRow(sheet, headers, hIdx, data);
      }
      return jsonResponse({ status: 'ok' });
    }

    return jsonResponse({ status: 'error', message: 'action không hợp lệ: ' + data.action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── ROW CRUD ───────────────────────────────────────────────────────────────

function findRowNum(sheet, headers, hIdx, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const idCol    = hIdx[norm('ID')];
  const nameCol  = hIdx[norm('Tên trụ')];
  const searchId = norm(data.id || '');
  const searchNm = norm(data.tenTru || '');
  if (!searchId && !searchNm) return -1;

  const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < allData.length; i++) {
    const rowId  = idCol   !== undefined ? norm(allData[i][idCol])   : '';
    const rowNm  = nameCol !== undefined ? norm(allData[i][nameCol]) : '';
    if ((searchId && rowId && rowId === searchId) ||
        (searchNm && rowNm && rowNm === searchNm)) {
      return i + 2;
    }
  }
  return -1;
}

function updateRow(sheet, hIdx, rowNum, data) {
  updateRowFields(sheet, hIdx, rowNum, buildFieldValues(data));
}

function updateRowFields(sheet, hIdx, rowNum, fieldValues) {
  for (const [header, value] of Object.entries(fieldValues)) {
    const col = hIdx[norm(header)];
    if (col !== undefined && value !== null && value !== undefined) {
      sheet.getRange(rowNum, col + 1).setValue(value);
    }
  }
}

function appendRow(sheet, headers, hIdx, data) {
  const fieldValues = buildFieldValues(data);
  const row = headers.map(h => fieldValues[h] !== undefined ? fieldValues[h] : '');
  sheet.appendRow(row);
}

function buildFieldValues(data) {
  const result = {};
  for (const [key, value] of Object.entries(data)) {
    if (key === 'action') continue;
    const header = FIELD_MAP[key] || key;
    if (value !== undefined && value !== null && value !== '') {
      result[header] = value;
    }
  }
  return result;
}

// ── HEALTH CHECK ───────────────────────────────────────────────────────────

function doGet(e) {
  ensureHeader();
  return jsonResponse({ status: 'ok', message: 'KhaoSat GAS v2 — login ready, header ensured' });
}
