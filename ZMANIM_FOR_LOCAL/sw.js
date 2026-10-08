const CACHE_NAME = 'shabbat-schedule-v1';
const urlsToCache = [
  '/',
  '/index.html',
  '/latest-schedule.jpg',
  '/manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) {
          return response;
        }
        return fetch(event.request)
          .then(response => {
            if (response.status === 404) {
              return caches.match('index.html');
            }
            return response;
          });
      })
  );
});

// Vérifier les nouvelles versions toutes les heures
self.addEventListener('periodicsync', event => {
  if (event.tag === 'update-schedule') {
    event.waitUntil(
      fetch('/latest-schedule.jpg')
        .then(response => {
          if (response.ok) {
            return caches.open(CACHE_NAME)
              .then(cache => cache.put('/latest-schedule.jpg', response));
          }
        })
    );
  }
});
