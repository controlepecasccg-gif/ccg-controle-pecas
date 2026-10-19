// sw.js — permite abrir o app sem sinal (coloque ao lado do index.html no Firebase Hosting)
const CACHE = 'ccg-pecas-v1';
const HOSTS = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com', 'www.gstatic.com', 'tessdata.projectnaptha.com'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './index.html'])).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.hostname.includes('firestore.googleapis.com') || u.hostname.includes('identitytoolkit') || u.hostname.includes('securetoken')) return;
  const mesmo = u.origin === location.origin;
  if (mesmo && (u.pathname.endsWith('/') || u.pathname.endsWith('.html'))) {
    // página: tenta a rede primeiro (pega versão nova), sem sinal usa o cache
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
    return;
  }
  if (mesmo || HOSTS.includes(u.hostname)) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return res; })));
  }
});
