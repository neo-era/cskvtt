// Kiem Code.gs (Apps Script) bang gia lap SpreadsheetApp/LockService. Chay: node tests/test_apps_script.js
const fs = require('fs'), vm = require('vm'), path = require('path');
let rows = [], frozen = 0, ok = 0, fail = 0;
const T = (ten, dk, x) => { if (dk) { ok++; console.log('  ĐẠT  ', ten); } else { fail++; console.log('  LỖI  ', ten, x === undefined ? '' : JSON.stringify(x)); } };
const sh = { getLastRow: () => rows.length, appendRow: (r) => rows.push(r), setFrozenRows: (n) => { frozen = n; }, getRange: () => ({ setFontWeight() {} }) };
const ctx = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => (rows.created ? sh : null), insertSheet: () => { rows.created = true; return sh; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (s) => ({ setMimeType: () => ({ body: s }) }) },
  JSON, String, Date,
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8'), ctx);
const post = (o) => JSON.parse(ctx.doPost({ postData: { contents: typeof o === 'string' ? o : JSON.stringify(o) } }).body);
let r = post({ luc: '2026-10-24T09:00:00+07:00', ten: 'Nguyễn Văn A', don_vi: 'Đội QLVH Q8', sdt: '0909', viec: 'Đăng ký', phien_ban: '2.6.0', may: 'abc' });
T('ghi dòng hợp lệ: ok', r.ok === true, r);
T('lần đầu tạo dòng tiêu đề + cố định', rows.length === 2 && rows[0][2] === 'Họ tên' && frozen === 1, rows[0]);
T('đúng thứ tự cột', rows[1][2] === 'Nguyễn Văn A' && rows[1][3] === 'Đội QLVH Q8' && rows[1][5] === 'Đăng ký' && rows[1][9] === 'abc', rows[1]);
r = post({ ten: 'x' }); T('thiếu "việc": từ chối, không ghi', r.ok === false && rows.length === 2, r);
r = post('không phải json'); T('dữ liệu hỏng: báo lỗi, không ghi', r.ok === false && rows.length === 2, r);
r = post({ viec: 'Lập hồ sơ', ten: '=HYPERLINK("x")', chi_tiet: 'a'.repeat(1000) });
const last = rows[rows.length - 1];
T('chặn công thức chèn vào ô (=…)', last[2].startsWith("'="), last[2]);
T('cắt chuỗi quá dài (300 ký tự)', last[7].length === 300, last[7].length);
T('không có hàm đọc/sửa/xoá công khai', ['doGet', 'doPost'].every(f => typeof ctx[f] === 'function') && !/getValues|deleteRow|clear\(|setValue/.test(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8')));
console.log('\nKết quả: ' + ok + '/' + (ok + fail) + ' đạt'); process.exit(fail ? 1 : 0);
