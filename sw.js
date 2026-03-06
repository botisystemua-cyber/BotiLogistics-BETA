// BOTI-Logistics Service Worker — мінімальний, для PWA install
const CACHE_NAME = 'boti-v1';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', () => self.clients.claim());

self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
