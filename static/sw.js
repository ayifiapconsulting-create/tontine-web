// Service worker minimal : ne met rien en cache (l'appli est dynamique, connectée
// en direct à la base de données), il sert uniquement à rendre l'appli "installable".
self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});
