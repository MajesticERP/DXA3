const CACHE_NAME = 'dxa3-v1';
const URLS_TO_CACHE = [
  '/DXA3/',
  '/DXA3/index.html',
  '/DXA3/thai-cargo-logo.png',
  '/DXA3/manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(URLS_TO_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // ไม่ cache คำขอไปยัง GAS API
  if (event.request.url.includes('script.google.com')) return;
  event.respondWith(
    caches.match(event.request).then(response => response || fetch(event.request))
  );
});
