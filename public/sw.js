// sw.js — deixa o app instalável e abre sem sinal
const CACHE = 'ccg-pecas-v2';
const BASE = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];
const HOSTS = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com', 'www.gstatic.com', 'tessdata.projectnaptha.com'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.hostname.includes('firestore.googleapis.com') || u.hostname.includes('identitytoolkit') || u.hostname.includes('securetoken')) return;
  const mesmo = u.origin === location.origin;
  if (e.request.mode === 'navigate' || (mesmo && (u.pathname.endsWith('/') || u.pathname.endsWith('.html')))) {
    // página: rede primeiro (pega a versão nova); sem sinal usa a cópia guardada
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
    return;
  }
  if (mesmo || HOSTS.includes(u.hostname)) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { if (res.ok || res.type === 'opaque') { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } return res; })));
  }
});
