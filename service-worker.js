const CACHE_NAME = 'consulta-eep-v2'; // bump this on every future deploy that changes app.js/index.html/style.css
const ASSETS = ['./index.html', './style.css', './app.js', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // Let cross-origin requests (e.g. the xlsx CDN library) go straight to network
  if (url.origin !== self.location.origin) return;
  // Never cache API calls — those must always try the network first
  if (url.pathname.startsWith('/api/')) return;

  // Network-first for the app shell: always try to get the latest version first,
  // only falling back to the cached copy when there's no connection (offline).
  // This is what prevents a stale app.js from getting stuck forever after a deploy.
  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const clone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
