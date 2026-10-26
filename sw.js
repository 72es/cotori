// kotori offline cache: app shell cache-first, data.json network-first (falls back to last copy).
const V = 'kotori-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.pathname.endsWith('/data.json') || u.pathname.endsWith('/index.html') || u.pathname.endsWith('/')) {
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r; })
      .catch(() => caches.match(e.request, {ignoreSearch: true})));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
