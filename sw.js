/* MyBudjet service worker: lets the app open instantly and install on Android.
   The page opens from the cache at once and is refreshed in the background (a new version shows on the next open).
   Calls to Apps Script and Google Fonts are never cached or touched. */
const CACHE = 'mybudjet-v3';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.allSettled(SHELL.map(u => c.add(u)))).then(() => self.skipWaiting()));   // one missing file must not break the install
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;     // backend + fonts go straight to the network
  e.respondWith(caches.match(req, {ignoreSearch: true}).then(cached => {
    const net = fetch(req).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => cached || caches.match('./index.html'));
    return cached || net;
  }));
});
