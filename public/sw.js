const CACHE_NAME = 'nma-offers-v1'

self.addEventListener('install', () => {
  self.skipWaiting()
})

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim())
})

// تمرير الطلبات عادي، مع رجوع للكاش لو النت انقطع (دعم بسيط للعمل بدون نت)
self.addEventListener('fetch', (event) => {
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  )
})