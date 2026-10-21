const VERSION = 'financas-plus-v2';
const SHELL = [
  '/',
  '/index.html',
  '/manifest.webmanifest',
  '/offline.html',
  '/icons/icon-192.png',
  '/icons/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(cache => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== VERSION).map(key => caches.delete(key)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname === '/api/sheets') return;
  event.respondWith(fetch(request).then(response => {
    if (response.ok && request.destination !== 'document') {
      const copy = response.clone();
      caches.open(VERSION).then(cache => cache.put(request, copy));
    }
    return response;
  }).catch(() => caches.match(request).then(cached => cached || (request.destination === 'document' ? caches.match('/offline.html') : Response.error()))));
});
