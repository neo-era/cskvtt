const CACHE_NAME = 'lighting-system-v1.1';
const TILE_CACHE_NAME = 'lighting-tiles-v1';
const TILE_CACHE_MAX = 200;

const urlsToCache = [
  './',
  'index.html',
  'manifest.json',
  'data/khaosat.xlsx',
  'css/bootstrap.min.css',
  'css/style.css',
  'css/responsive.css',
  'css/jquery.mCustomScrollbar.min.css',
  'css/owl.carousel.min.css',
  'css/owl.theme.default.min.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.1/exceljs.min.js',
  'https://netdna.bootstrapcdn.com/font-awesome/4.0.3/css/font-awesome.css',
  'images/fevicon.png'
];

const tileHosts = [
  'tile.openstreetmap.org',
  'mt0.google.com', 'mt1.google.com', 'mt2.google.com', 'mt3.google.com'
];

function isTileRequest(url) {
  try {
    const host = new URL(url).hostname;
    return tileHosts.some(h => host.includes(h));
  } catch { return false; }
}

async function trimTileCache() {
  const cache = await caches.open(TILE_CACHE_NAME);
  const keys = await cache.keys();
  if (keys.length > TILE_CACHE_MAX) {
    const toDelete = keys.slice(0, keys.length - TILE_CACHE_MAX);
    await Promise.all(toDelete.map(k => cache.delete(k)));
  }
}

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME && k !== TILE_CACHE_NAME).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = event.request.url;

  // Map tiles: cache-first, giới hạn 200 tile để tiết kiệm storage
  if (isTileRequest(url)) {
    event.respondWith(
      caches.open(TILE_CACHE_NAME).then(async cache => {
        const cached = await cache.match(event.request);
        if (cached) return cached;
        const response = await fetch(event.request);
        if (response.ok) {
          cache.put(event.request, response.clone());
          trimTileCache();
        }
        return response;
      }).catch(() => caches.match(event.request))
    );
    return;
  }

  // Static assets: cache-first
  event.respondWith(
    caches.match(event.request).then(response => response || fetch(event.request))
  );
});
