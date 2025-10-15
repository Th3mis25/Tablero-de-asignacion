// Service Worker para cachear assets de la PWA
const CACHE_NAME = 'cargas-pwa-v13';
const DYNAMIC_CACHE = 'cargas-pwa-dynamic-v1';
const ASSETS = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './assets/logo.png',
  './manifest.json',
  './offline.html',
];

const ASSET_URLS = ASSETS.map(a => new URL(a, self.location).pathname);

self.addEventListener('install', e => {
  e.waitUntil(
    caches
      .open(CACHE_NAME)
      .then(c => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => ![CACHE_NAME, DYNAMIC_CACHE].includes(k)).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const { request } = e;

  if (request.headers.has('Authorization')) {
    e.respondWith(fetch(request));
    return;
  }

  if (request.method !== 'GET') {
    e.respondWith(fetch(request));
    return;
  }

  const url = new URL(request.url);

  if (url.origin !== self.location.origin) {
    e.respondWith(fetch(request));
    return;
  }
  if (url.pathname.endsWith('/config.js')) {
    e.respondWith(fetch(request));
    return;
  }

  if (ASSET_URLS.includes(url.pathname)) {
    e.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(request);
      const fetchPromise = fetch(request)
        .then(response => {
          cache.put(request, response.clone());
          return response;
        })
        .catch(async () =>
          cached ||
          (await caches.match(new URL('./offline.html', self.location).pathname))
        );
      return cached || fetchPromise;
    })());
    return;
  }

  e.respondWith((async () => {
    try {
      const response = await fetch(request);
      const cache = await caches.open(DYNAMIC_CACHE);
      cache.put(request, response.clone());
      return response;
    } catch {
      const cached = await caches.match(request);
      return (
        cached ||
        (await caches.match(new URL('./offline.html', self.location).pathname))
      );
    }
  })());
});
