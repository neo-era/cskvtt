/* Service worker Hoàn công CSKVTT
   - Vỏ trang (index, manifest, icon) và pylibs.zip (mã xử lý, đổi mỗi lần cập nhật):
     network-first có thời gian chờ 4 giây → mạng tốt luôn lấy bản mới, mạng yếu/mất mạng dùng bản đã lưu.
   - pyodide/* (bộ chạy Python, ~20 MB, phiên bản cố định) và font: cache-first, lưu riêng theo phiên bản Pyodide,
     KHÔNG xóa khi cập nhật code. Lưu sẵn ngay khi cài để lần sau mở được khi mất mạng.
   Đổi VERSION khi đổi vỏ trang; đổi PYO khi nâng Pyodide. */
const VERSION = 'hc-shell-2026-10-17b';
const PYO = 'hc-pyodide-0.26.4';
const SHELL = ['./', './index.html', './manifest.webmanifest', './pylibs.zip', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png'];
const PYO_FILES = ['./pyodide/pyodide.js', './pyodide/pyodide.asm.js', './pyodide/pyodide.asm.wasm', './pyodide/python_stdlib.zip', './pyodide/pyodide-lock.json'];
self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    await (await caches.open(VERSION)).addAll(SHELL);
    try { await (await caches.open(PYO)).addAll(PYO_FILES); } catch (err) { /* lưu sau khi dùng */ }
    await self.skipWaiting();
  })());
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('hc-') && k !== VERSION && k !== PYO && k !== 'hc-fonts').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
const put = (c, req, r) => { if (r && r.status === 200) c.put(req, r.clone()).catch(() => {}); };
async function networkFirst(req) {
  const c = await caches.open(VERSION);
  const net = fetch(req).then((r) => { put(c, req, r); return r; });
  const hit = await c.match(req, {ignoreSearch: true});
  if (!hit) return net;
  const slow = new Promise((res) => setTimeout(() => res(hit), 4000));
  try { return await Promise.race([net, slow]); } catch (err) { return hit; }
}
async function cacheFirst(name, req) {
  const c = await caches.open(name);
  const hit = await c.match(req, {ignoreSearch: true}); if (hit) return hit;
  const r = await fetch(req); put(c, req, r); return r;
}
self.addEventListener('fetch', (e) => {
  const req = e.request; if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (u.origin === location.origin && u.pathname.includes('/pyodide/')) return e.respondWith(cacheFirst(PYO, req));
  if (u.hostname === 'fonts.googleapis.com' || u.hostname === 'fonts.gstatic.com') return e.respondWith(cacheFirst('hc-fonts', req));
  if (u.origin === location.origin) return e.respondWith(networkFirst(req));
});
