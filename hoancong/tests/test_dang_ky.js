// Kiem giao dien dang ky + nhat ky su dung. Chay (tu thu muc hoancong, may chu tinh o cong 8765):
//   python3 -m http.server 8765 &  NODE_PATH=$(npm root -g) node tests/test_dang_ky.js <bang tra .xlsx> <BCSC nho .pdf>
const {chromium} = require('playwright');
const URL = 'http://localhost:8765/index.html', LOG = 'https://log.test/exec';
const [BT, BCSC] = process.argv.slice(2);
let ok = 0, fail = 0;
const T = (ten, dk, x) => { if (dk) { ok++; console.log('  ĐẠT  ', ten); } else { fail++; console.log('  LỖI  ', ten, x === undefined ? '' : JSON.stringify(x).slice(0, 300)); } };
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({viewport: {width: 1200, height: 850}});
  await ctx.addInitScript((u) => { window.HC_LOG_URL = u; }, LOG);
  let mang = true; const nhan = [];
  await ctx.route('**/fonts.googleapis.com/**', r => r.abort());
  await ctx.route(LOG + '*', async r => { if (!mang) return r.abort(); nhan.push(JSON.parse(r.request().postData())); r.fulfill({status: 200, body: '{"ok":true}'}); });
  const p = await ctx.newPage();
  p.on('pageerror', e => console.log('PAGEERR', e.message));
  await p.goto(URL);
  await p.waitForSelector('#dkDlg', {state: 'visible', timeout: 15000}).catch(() => {});
  T('lần đầu mở: hiện hộp thoại đăng ký', await p.isVisible('#dkDlg'));
  T('con trỏ ở ô Họ tên', await p.evaluate(() => document.activeElement && document.activeElement.id) === 'dkTen');
  await p.click('#dkSau');
  T('"Để sau": đóng hộp thoại', !(await p.isVisible('#dkDlg')));
  T('thanh tiêu đề ghi "Chưa đăng ký"', (await p.textContent('#whoBtn')).includes('Chưa đăng ký'));
  await p.reload(); await p.waitForTimeout(800);
  T('mở lại sau "Để sau": hỏi lại', await p.isVisible('#dkDlg'));
  await p.click('#dkLuu');
  T('bỏ trống Họ tên: không lưu, báo lỗi', await p.isVisible('#dkDlg') && (await p.textContent('#dkLoi')).length > 0);
  await p.fill('#dkTen', '  Nguyễn Văn A '); await p.fill('#dkDv', 'Đội QLVH Q8'); await p.fill('#dkSdt', '0909 123 456');
  await p.click('#dkLuu'); await p.waitForTimeout(800);
  T('lưu: đóng hộp thoại, tên lên thanh tiêu đề', !(await p.isVisible('#dkDlg')) && (await p.textContent('#whoBtn')).includes('Nguyễn Văn A'));
  const dk = nhan.find(x => x.viec === 'Đăng ký');
  T('ghi Sheet dòng "Đăng ký" (tên đã cắt khoảng trắng, đơn vị, SĐT, mã máy, phiên bản)', dk && dk.ten === 'Nguyễn Văn A' && dk.don_vi === 'Đội QLVH Q8' && dk.sdt === '0909 123 456' && dk.may && dk.phien_ban, dk);
  await p.reload(); await p.waitForTimeout(800);
  T('mở lại sau khi đăng ký: không hỏi nữa', !(await p.isVisible('#dkDlg')) && (await p.textContent('#whoBtn')).includes('Nguyễn Văn A'));
  await p.click('#whoBtn');
  T('bấm tên: mở hộp thoại, điền sẵn', await p.isVisible('#dkDlg') && await p.inputValue('#dkTen') === 'Nguyễn Văn A');
  await p.keyboard.press('Escape');
  T('Esc: đóng', !(await p.isVisible('#dkDlg')));
  // mat mang: xep hang, co mang lai thi gui
  mang = false; const truoc = nhan.length;
  await p.click('#whoBtn'); await p.fill('#dkDv', 'Đội QLVH Q10'); await p.click('#dkLuu'); await p.waitForTimeout(1200);
  const hang = await p.evaluate(() => window.__hcNhatKy && window.__hcNhatKy.hang().then(h => h.length));
  T('mất mạng: app vẫn chạy, dòng ghi nằm trong hàng đợi', hang === 1 && nhan.length === truoc, hang);
  mang = true; await p.evaluate(() => window.dispatchEvent(new Event('online'))); await p.waitForTimeout(1500);
  const hang2 = await p.evaluate(() => window.__hcNhatKy.hang().then(h => h.length));
  T('có mạng lại: gửi bù, hàng đợi rỗng', hang2 === 0 && nhan.some(x => x.viec === 'Đổi thông tin' && x.don_vi === 'Đội QLVH Q10'), {hang2, n: nhan.length});
  // chuyen BCSC -> ghi "Chuyển BCSC"
  if (BT && BCSC) {
    await p.click('#tabSC');
    await p.setInputFiles('#master', BT); await p.setInputFiles('#scPdf', BCSC);
    await p.waitForFunction(() => !document.getElementById('scRun').disabled, null, {timeout: 300000});
    await p.click('#scRun');
    await p.waitForFunction(() => /xong|không xử lý/.test(document.getElementById('scText').textContent), null, {timeout: 300000});
    await p.waitForTimeout(800);
    const sc = nhan.find(x => x.viec === 'Chuyển BCSC');
    T('chuyển BCSC: ghi tên file + số báo cáo, không gửi nội dung file', sc && sc.file && /báo cáo/.test(sc.chi_tiet) && JSON.stringify(sc).length < 2000, sc);
  }
  // chua dang ky + mobile
  const c2 = await b.newContext({viewport: {width: 390, height: 780}}); await c2.addInitScript((u) => { window.HC_LOG_URL = u; }, LOG);
  await c2.route('**/fonts.googleapis.com/**', r => r.abort());
  const p2 = await c2.newPage(); await p2.goto(URL); await p2.waitForTimeout(1000);
  const tran = await p2.evaluate(() => document.documentElement.scrollWidth > window.innerWidth);
  const box = await p2.locator('.dlg').boundingBox();
  T('điện thoại 390px: hộp thoại nằm gọn, không tràn ngang', !tran && box && box.x >= 0 && box.x + box.width <= 390, box);
  await b.close();
  console.log('\nKết quả: ' + ok + '/' + (ok + fail) + ' đạt'); process.exit(fail ? 1 : 0);
})();
