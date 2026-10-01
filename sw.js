/* Recogida de Leche: la app funciona sin cobertura. Los datos van a la cola del móvil, no aquí. */
const CACHE = 'recogida-v22';
const ARCHIVOS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',
  'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(k => Promise.all(k.filter(n => n !== CACHE).map(n => caches.delete(n)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || r.url.includes('script.google.com') || r.url.includes('googleusercontent.com')) return;
  const esPagina = r.mode === 'navigate' || r.url.endsWith('/index.html');
  if (esPagina) {   // la app: primero red (para recibir mejoras), si no hay, la guardada
    e.respondWith(fetch(r).then(res => { caches.open(CACHE).then(c => c.put('./index.html', res.clone())); return res; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => {
    const copia = res.clone(); caches.open(CACHE).then(c => c.put(r, copia)); return res; })));
});
