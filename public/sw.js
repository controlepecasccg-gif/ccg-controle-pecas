// sw.js — deixa o app instalável e faz ele abrir sem sinal
const CACHE = 'ccg-pecas-v4';
const BASE = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './jsqr.js'];
// bibliotecas que o app precisa para abrir e entrar: já ficam guardadas na instalação
const LIBS = [
  'https://cdnjs.cloudflare.com/ajax/libs/react/18.3.1/umd/react.production.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.3.1/umd/react-dom.production.min.js',
  'https://cdn.jsdelivr.net/npm/htm@3.1.1/dist/htm.umd.js',
  'https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js',
  'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth-compat.js',
  'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore-compat.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap'
];
const HOSTS = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com', 'www.gstatic.com', 'tessdata.projectnaptha.com'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(async c => {
    await c.addAll(BASE);
    await Promise.all(LIBS.map(u => fetch(u, { mode: 'cors' }).then(r => r.ok ? c.put(u, r) : null).catch(() => null)));
  }).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.hostname.includes('firestore.googleapis.com') || u.hostname.includes('identitytoolkit') || u.hostname.includes('securetoken') || u.pathname.startsWith('/__/')) return;
  const mesmo = u.origin === location.origin;
  if (e.request.mode === 'navigate' || (mesmo && (u.pathname.endsWith('/') || u.pathname.endsWith('.html')))) {
    // página: rede primeiro (pega a versão nova); sem sinal (ou rede muito lenta) usa a cópia guardada
    e.respondWith(new Promise(resolve => {
      let feito = false; const usarCache = () => caches.match(e.request).then(r => r || caches.match('./index.html')).then(r => { if (!feito && r) { feito = true; resolve(r); } });
      const t = setTimeout(usarCache, 4000);
      fetch(e.request).then(r => { clearTimeout(t); if (r.ok) { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } if (!feito) { feito = true; resolve(r); } })
        .catch(() => { clearTimeout(t); usarCache().then(() => { if (!feito) { feito = true; resolve(Response.error()); } }); });
    }));
    return;
  }
  if (mesmo || HOSTS.includes(u.hostname)) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { if (res.ok || res.type === 'opaque') { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } return res; })));
  }
});
