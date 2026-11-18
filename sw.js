// Guarda la aplicación para que abra sin señal. Los datos de las órdenes NO pasan por
// aquí: los guarda la propia aplicación (IndexedDB) y Graph nunca se cachea.
const VERSION = "oc-movil-11";
const ARCHIVOS = ["./", "index.html", "app.css", "app.js", "config.js", "msal-browser.min.js", "lugares.js", "mapa.json",
  "manifest.webmanifest", "iconos/icono-192.png", "iconos/icono-512.png", "iconos/icono-180.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// lo propio: primero la red (así llegan las versiones nuevas) y, sin señal, lo guardado
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then(r => {
      if (r.ok) { const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); }
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match("index.html")))
  );
});

// al tocar un aviso: abrir (o traer al frente) la aplicación
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const destino = new URL(e.notification.data?.url || "./", self.registration.scope).href;
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(l => {
    const abierta = l.find(c => c.url.startsWith(self.registration.scope));
    return abierta ? abierta.focus().then(c => c.navigate?.(destino)) : self.clients.openWindow(destino);
  }));
});
