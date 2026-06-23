// ── VERSIÓN: cambia este número cada vez que hagas un deploy ──
// Ej: poritos-v2, poritos-v3, poritos-v4...
const CACHE = 'metta-v26';

const ASSETS = [
  '/Finanzas-Wiedman-Placencia/',
  '/Finanzas-Wiedman-Placencia/index.html',
  '/Finanzas-Wiedman-Placencia/manifest.json',
  '/Finanzas-Wiedman-Placencia/icon-192.png',
  '/Finanzas-Wiedman-Placencia/icon-512.png',
  '/Finanzas-Wiedman-Placencia/apple-touch-icon.png'
];

self.addEventListener('install', e => {
  // No skipWaiting: el SW nuevo espera y la app muestra el banner "Actualizar".
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).catch(() => {}));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Escucha el mensaje desde index.html para activar la nueva versión
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  // No cachear APIs externas ni en tiempo real (Supabase, Claude, Apps Script, CDN)
  if (url.includes('supabase.co') ||
      url.includes('supabase.in') ||
      url.includes('script.google.com') ||
      url.includes('anthropic.com') ||
      url.includes('cdn.jsdelivr.net')) {
    return; // deja pasar a la red sin tocar la caché
  }
  // index.html / navegación: network-first para no quedar con versión vieja
  if (e.request.mode === 'navigate' || url.endsWith('/index.html')) {
    e.respondWith(
      fetch(e.request)
        .then(resp => {
          const copy = resp.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
          return resp;
        })
        .catch(() => caches.match(e.request).then(c => c || caches.match('/Finanzas-Wiedman-Placencia/index.html')))
    );
    return;
  }
  // Resto de assets: cache-first
  e.respondWith(
    caches.match(e.request).then(cached =>
      cached || fetch(e.request).then(resp => {
        if (resp.ok && e.request.method === 'GET') {
          const copy = resp.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        }
        return resp;
      }).catch(() => cached)
    )
  );
});
