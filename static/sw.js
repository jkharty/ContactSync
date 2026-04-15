// Service Worker for InVision Contact Sync PWA
const CACHE_NAME = 'contactsync-v1';

// Assets to pre-cache on install
const PRE_CACHE = [
  '/static/style.css',
  '/static/manifest.json',
  '/static/icons/icon-192.png',
  '/static/icons/icon-512.png',
];

// Install — pre-cache static assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRE_CACHE))
      .then(() => self.skipWaiting())
  );
});

// Activate — clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Fetch — stale-while-revalidate for static assets, network-first for pages/API
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Static assets: serve from cache immediately (fast), then fetch fresh copy
  // in the background so the next load always gets up-to-date assets.
  // This means CSS/icon deploys are live on the very next page load —
  // no need to ever bump CACHE_NAME.
  if (url.pathname.startsWith('/static/')) {
    event.respondWith(
      caches.open(CACHE_NAME).then(cache =>
        cache.match(event.request).then(cached => {
          const networkFetch = fetch(event.request).then(response => {
            cache.put(event.request, response.clone());
            return response;
          });
          return cached || networkFetch;
        })
      )
    );
    return;
  }

  // Everything else (pages, API calls): network-first, cache as fallback
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
