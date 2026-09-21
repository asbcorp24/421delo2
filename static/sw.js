const CACHE_NAME = "otd421-shell-v2";
const APP_SHELL = [
  "/",
  "/manifest.webmanifest",
  "/static/bootstrap.min.css",
  "/static/bootstrap.bundle.min.js",
  "/static/bootstrap-icons.css",
  "/static/jquery-3.6.0.min.js",
  "/static/pwa-icon.svg",
  "/static/pwa-icon-192.png",
  "/static/pwa-icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (event.request.url.startsWith(self.location.origin)) {
          const cloned = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, cloned));
        }
        return response;
      })
      .catch(() => caches.match(event.request).then((cached) => cached || caches.match("/")))
  );
});
