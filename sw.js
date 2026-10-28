// cotori offline cache: page and data.json always fetched fresh from the network (skipping the 10-minute HTTP cache), last copy used only when offline.
const V = 'cotori-v6';
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
    e.respondWith(fetch(u.href, {cache: 'no-store', credentials: 'same-origin'}).then(r => { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r; })
      .catch(() => caches.match(e.request, {ignoreSearch: true})));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
