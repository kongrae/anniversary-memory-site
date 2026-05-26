const CACHE_NAME = "anniversary-memory-v55";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./assets/vendor/react.production.min.js",
  "./assets/vendor/react-dom.production.min.js",
  "./assets/vendor/babel.min.js",
  "./assets/main-hero-cropped.jpg",
  "./assets/letter-photo.jpg",
  "./assets/hero-stack/stack-1.jpg",
  "./assets/hero-stack/stack-2.jpg",
  "./assets/hero-stack/stack-3.jpg",
  "./assets/hero-stack/stack-4.jpg",
  "./assets/icons/apple-touch-icon.png",
  "./assets/icons/icon-192.png",
  "./assets/icons/icon-512.png",
  "./assets/seasons/summer-1.jpg",
  "./assets/seasons/summer-2.jpg",
  "./assets/seasons/summer-3.jpg",
  "./assets/seasons/summer-4.jpg",
  "./assets/seasons/summer-5.jpg",
  "./assets/seasons/summer-6.jpg",
  "./assets/seasons/autumn-1.jpg",
  "./assets/seasons/autumn-2.jpg",
  "./assets/seasons/autumn-3.mp4",
  "./assets/seasons/autumn-4.jpg",
  "./assets/seasons/autumn-5.jpg",
  "./assets/seasons/autumn-6.jpg",
  "./assets/seasons/autumn-7.jpg",
  "./assets/seasons/autumn-8.jpg",
  "./assets/seasons/winter-1.jpg",
  "./assets/seasons/winter-2.mov",
  "./assets/seasons/winter-3.jpg",
  "./assets/seasons/winter-4.jpg",
  "./assets/seasons/winter-5.jpg",
  "./assets/seasons/spring-1.jpg",
  "./assets/seasons/spring-2.jpg",
  "./assets/seasons/spring-3.jpg",
  "./assets/seasons/spring-4.jpg",
  "./assets/seasons/spring-5.jpg",
  "./assets/seasons/spring-6.jpg",
  "./assets/seasons/spring-7.jpg",
  "./assets/seasons/spring-8.jpg",
  "./assets/seasons/spring-9.jpg",
  "./assets/seasons/spring-10.jpg",
  "./assets/seasons/spring-11.jpg",
  "./assets/seasons/spring-12.jpg"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys
        .filter(key => key !== CACHE_NAME)
        .map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put("./index.html", copy));
          return response;
        })
        .catch(() => caches.match("./index.html"))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;

      return fetch(event.request).then(response => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return response;
      });
    })
  );
});
