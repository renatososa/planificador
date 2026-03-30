const CACHE_NAME = 'planificador-actividades-v2';
const RUNTIME_CACHE_NAME = 'planificador-actividades-runtime-v2';
const APP_SHELL = [
  './style.css',
  './app.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
  './icons/apple-touch-icon.svg'
];
const STATIC_CDN_PREFIXES = [
  'https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME && key !== RUNTIME_CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') {
    return;
  }

  const requestUrl = new URL(event.request.url);
  const isSameOrigin = requestUrl.origin === self.location.origin;
  const esCDNEstatico = STATIC_CDN_PREFIXES.some((prefix) => event.request.url.startsWith(prefix));

  if (!isSameOrigin && !esCDNEstatico) {
    return;
  }

  if (event.request.mode === 'navigate' && isSameOrigin) {
    event.respondWith(
      fetch(event.request)
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }

      return fetch(event.request).then((response) => {
        const responseClone = response.clone();
        const cacheDestino = esCDNEstatico ? RUNTIME_CACHE_NAME : CACHE_NAME;
        caches.open(cacheDestino).then((cache) => cache.put(event.request, responseClone));
        return response;
      });
    })
  );
});
