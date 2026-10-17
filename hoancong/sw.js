/* Service worker Hoàn công CSKVTT
   - Vỏ trang (index, manifest, icon): network-first → luôn lấy bản mới khi có mạng, mất mạng thì dùng bản đã lưu.
   - pylibs.zip (mã xử lý): network-first (đổi mỗi lần cập nhật code).
   - pyodide/* và font: cache-first (file lớn, có phiên bản cố định).
   Đổi VERSION khi đổi danh sách vỏ trang. */
const VERSION = 'hc-v2-2026-10-17';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png'];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSION && k.startsWith('hc-')).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
async function networkFirst(req) {
  const c = await caches.open(VERSION);
  try { const r = await fetch(req); if (r && r.ok) c.put(req, r.clone()); return r; }
  catch (err) { const hit = await c.match(req, {ignoreSearch: true}); if (hit) return hit; throw err; }
}
async function cacheFirst(req) {
  const c = await caches.open('hc-static');
  const hit = await c.match(req); if (hit) return hit;
  const r = await fetch(req); if (r && (r.ok || r.type === 'opaque')) c.put(req, r.clone()); return r;
}
self.addEventListener('fetch', (e) => {
  const req = e.request; if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (u.origin === location.origin && u.pathname.includes('/pyodide/')) return e.respondWith(cacheFirst(req));
  if (u.hostname === 'fonts.googleapis.com' || u.hostname === 'fonts.gstatic.com') return e.respondWith(cacheFirst(req));
  if (u.origin === location.origin) return e.respondWith(networkFirst(req));
});
