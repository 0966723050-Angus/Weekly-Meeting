// sw.js - PWA 離線快取:只快取網站本身的靜態檔(網路優先),Google API 與檔案資料一律走網路
const CACHE_NAME = 'atk-weekly-v1';
const APP_SHELL = [
  './', './index.html', './css/style.css', './js/config.js', './js/xlsx-model.js',
  './js/drive.js', './js/app.js', './manifest.webmanifest', './icons/icon.svg',
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) =>
    Promise.all(APP_SHELL.map((url) => fetch(url, { cache: 'reload' }).then((r) => r.ok && cache.put(url, r)).catch(() => {})))
  ));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then((resp) => {
        if (resp && resp.ok) { const c = resp.clone(); caches.open(CACHE_NAME).then((cache) => cache.put(event.request, c)); }
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
