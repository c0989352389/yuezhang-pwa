const CACHE = 'yuezhang-v1.4';
const STATIC = ['./manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (url.includes('script.google.com') || url.includes('googleapis.com')) return;

  // HTML 與根目錄走 network-first(網路優先,失敗才用快取),確保 PWA 拿到最新
  if (e.request.mode === 'navigate' || url.endsWith('/') || url.endsWith('.html')) {
    e.respondWith(
      fetch(e.request).then(res => {
        const c = res.clone();
        caches.open(CACHE).then(ca => ca.put(e.request, c));
        return res;
      }).catch(() => caches.match(e.request).then(r => r || caches.match('./')))
    );
    return;
  }

  // CDN 字型/Chart.js — stale-while-revalidate
  if (url.includes('fonts.googleapis') || url.includes('cdn.jsdelivr')) {
    e.respondWith(
      fetch(e.request).then(res => {
        const c = res.clone();
        caches.open(CACHE).then(ca => ca.put(e.request, c));
        return res;
      }).catch(() => caches.match(e.request))
    );
    return;
  }

  // 其他靜態資源 — cache-first
  e.respondWith(caches.match(e.request).then(cached => {
    if (cached) return cached;
    return fetch(e.request).then(res => {
      if (e.request.method === 'GET' && res.status === 200) {
        const c = res.clone();
        caches.open(CACHE).then(ca => ca.put(e.request, c));
      }
      return res;
    });
  }));
});
