// ── VERSIÓN: cambia este número cada vez que hagas un deploy ──
// Ej: poritos-v2, poritos-v3, poritos-v4...
const CACHE = 'poritos-v4';

const ASSETS = [
  '/Finanzas-Wiedman-Placencia/',
  '/Finanzas-Wiedman-Placencia/index.html',
  '/Finanzas-Wiedman-Placencia/manifest.json',
  '/Finanzas-Wiedman-Placencia/icon-192.png',
  '/Finanzas-Wiedman-Placencia/icon-512.png'
];

self.addEventListener('install', e => {
  // NO llamar skipWaiting() aquí — así el nuevo SW espera
  // y la app puede mostrar el banner de actualización
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
  // No cachear APIs externas (tiempo real)
  if (url.includes('script.google.com') || url.includes('anthropic.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached =>
      cached || fetch(e.request).then(resp => {
        if (resp.ok && e.request.method === 'GET') {