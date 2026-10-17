/* Service worker Hoàn công CSKVTT
   - Vỏ trang (index, manifest, icon): network-first, chờ tối đa 6 giây rồi mới dùng bản đã lưu.
   - Mã xử lý py/pylibs.zip?v=<phiên bản>: lưu THEO ĐÚNG phiên bản (khớp cả ?v=) → đổi phiên bản là bắt buộc tải bản mới,
     không bao giờ trả bản cũ cho trang mới. Chỉ giữ 1 phiên bản.
   - pyodide/*, font Google và font vẽ PDF (fonts/*.ttf): cache-first, giữ qua các lần cập nhật. Lưu sẵn khi cài để mở được khi mất mạng. */
const VERSION = 'hc-shell-2.4.2';
const PYO = 'hc-pyodide-0.26.4';
const LIBS = 'hc-libs';
const FONTS = 'hc-pdffonts';   // font de ve PDF: luu lau dai (cache-first)
const FONT_FILES = ['./fonts/LiberationSerif-Regular.ttf', './fonts/LiberationSerif-Bold.ttf', './fonts/LiberationSerif-Italic.ttf', './fonts/LiberationSerif-BoldItalic.ttf', './fonts/Carlito-Regular.ttf', './fonts/Carlito-Bold.ttf'];
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png'];
const PYO_FILES = ['./pyodide/pyodide.js', './pyodide/pyodide.asm.js', './pyodide/pyodide.asm.wasm', './pyodide/python_stdlib.zip', './pyodide/pyodide-lock.json'];
self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    await (await caches.open(VERSION)).addAll(SHELL);
    try { await (await caches.open(PYO)).addAll(PYO_FILES); } catch (err) { /* lưu sau khi dùng */ }
    try { await (await caches.open(FONTS)).addAll(FONT_FILES); } catch (err) { /* lưu khi tạo PDF lần đầu */ }
    await self.skipWaiting();
  })());
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('hc-') && ![VERSION, PYO, LIBS, FONTS, 'hc-fonts'].includes(k)).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
const put = (c, req, r) => { if (r && r.status === 200) c.put(req, r.clone()).catch(() => {}); };
async function shell(req) {
  const c = await caches.open(VERSION);
  const net = fetch(req).then((r) => { put(c, req, r); return r; });
  const hit = await c.match(req, {ignoreSearch: true});
  if (!hit) return net;
  const slow = new Promise((res) => setTimeout(() => res(hit), 6000));
  try { return await Promise.race([net, slow]); } catch (err) { return hit; }
}
async function libs(req) {                      // khớp ĐÚNG phiên bản (không ignoreSearch)
  const c = await caches.open(LIBS);
  const hit = await c.match(req); if (hit) return hit;
  const r = await fetch(req);
  if (r && r.status === 200) { for (const k of await c.keys()) await c.delete(k); await c.put(req, r.clone()).catch(() => {}); }
  return r;
}
async function cacheFirst(name, req) {
  const c = await caches.open(name);
  const hit = await c.match(req, {ignoreSearch: true}); if (hit) return hit;
  const r = await fetch(req); put(c, req, r); return r;
}
self.addEventListener('fetch', (e) => {
  const req = e.request; if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (u.origin === location.origin && u.pathname.endsWith('/py/pylibs.zip')) return e.respondWith(libs(req));
  if (u.origin === location.origin && u.pathname.includes('/pyodide/')) return e.respondWith(cacheFirst(PYO, req));
  if (u.origin === location.origin && u.pathname.includes('/fonts/') && u.pathname.endsWith('.ttf')) return e.respondWith(cacheFirst(FONTS, req));
  if (u.hostname === 'fonts.googleapis.com' || u.hostname === 'fonts.gstatic.com') return e.respondWith(cacheFirst('hc-fonts', req));
  if (u.origin === location.origin) return e.respondWith(shell(req));
});
