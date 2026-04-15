// ContactSync — minimal service worker
// Network-first: no offline caching (live data app, always needs the server).
// The service worker exists solely to satisfy PWA installability requirements.

const VERSION = "cs-1.0";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("fetch", event => {
  // Pass all requests straight through to the network.
  event.respondWith(fetch(event.request));
});
