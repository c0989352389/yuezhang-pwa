const CACHE  = 'yuezhang-v1.0';
const STATIC = ['./', './index.html', './manifest.json'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (url.includes('script.google.com') || url.includes('googleapis.com')) return;
  if (url.includes('fonts.googleapis') || url.includes('cdn.jsdelivr')) {
    e.respondWith(fetch(e.request).then(res => { const c = res.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); return res; }).catch(() => caches.match(e.request)));
    return;
  }
  e.respondWith(caches.match(e.request).then(cached => { if (cached) return cached; return fetch(e.request).then(res => { if (e.request.method === 'GET' && res.status === 200) { const c = res.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); } return res; }); }));
});
