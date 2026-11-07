/* Screening Atlas app shell only. Imported Excel/JSON data is never cached here. */
const CACHE = 'screening-atlas-shell-v5-20261107';
const FILES = [
  './', './index.html', './styles.css', './js/app.js', './js/excel.js',
  './js/model.js', './assets/icon.svg', './favicon.png',
  './apple-touch-icon.png', './apple-touch-icon-screening-v4.png',
  './assets/icon-192-v4.png', './assets/icon-512-v4.png',
  './manifest.webmanifest'
];
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(key => key.startsWith('screening-atlas-shell-') && key !== CACHE).map(key => caches.delete(key))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (request.mode === 'navigate') {
    event.respondWith(fetch(request).then(response => {
      if (response.ok) {
        const copy = response.clone();
        event.waitUntil(caches.open(CACHE).then(cache => cache.put(request, copy)));
      }
      return response;
    }).catch(() => caches.match(request).then(cached => cached || caches.match('./index.html'))));
    return;
  }
  event.respondWith(caches.match(request).then(cached => cached || fetch(request)));
});
