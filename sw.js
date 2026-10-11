// Daybook service worker: offline app shell.
// Online, every request goes to the network first (revalidating past the browser's HTTP cache,
// which GitHub Pages sets to 10 minutes), so a new deploy shows up on the next launch.
// Offline, the last good copy is served from the cache.
// Bump VERSION whenever you deploy changed files.
const VERSION = "daybook-v8";
const SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./config.js",
  "./manifest.webmanifest",
  "./vendor/msal-browser.min.js",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/apple-touch-icon.png",
  "./icons/favicon.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(VERSION)
      .then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: "reload" }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  const url = new URL(req.url);
  // Only handle our own GET requests; sign-in and Graph calls go straight to the network.
  if (req.method !== "GET" || url.origin !== self.location.origin) return;

  // Pages share one cached shell, whatever their query string (?demo, sign-in redirects).
  const key = req.mode === "navigate" ? "./index.html" : url.pathname;
  e.respondWith(
    fetch(req.url, { cache: "no-cache", credentials: "same-origin" })
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(VERSION).then((c) => c.put(key, copy));
        }
        return res;
      })
      .catch(() => caches.match(key, { ignoreSearch: true }).then((hit) => hit || Response.error()))
  );
});
