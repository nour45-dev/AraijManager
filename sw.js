/**
 * Araij Manager Pro - PWA Service Worker (100% Offline Engine)
 */

const CACHE_NAME = 'araij-manager-v4.1';
const CORE_ASSETS = [
  './',
  './index.html',
  './app.js',
  './data.js',
  './tailwind.min.js',
  './lucide.min.js',
  './chart.min.js',
  './html2pdf.bundle.min.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon.svg'
];

// External fonts to pre-cache
const CDN_ASSETS = [
  'https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&family=Tajawal:wght@400;500;700;800&display=swap'
];

// 1. INSTALL EVENT - Pre-cache core files
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      console.log('[ServiceWorker] Pre-caching core offline assets...');
      
      // Cache local core assets
      try {
        await cache.addAll(CORE_ASSETS);
      } catch (err) {
        console.warn('[ServiceWorker] Core asset caching note:', err);
      }

      // Try caching CDN assets gracefully
      for (const url of CDN_ASSETS) {
        try {
          const res = await fetch(url, { mode: 'cors' });
          if (res && res.ok) {
            await cache.put(url, res);
          }
        } catch (e) {
          console.warn('[ServiceWorker] CDN cache skipped for:', url);
        }
      }
    })
  );
});

// 2. ACTIVATE EVENT - Clean up outdated caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[ServiceWorker] Removing old cache:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 3. FETCH EVENT - Network-First for fresh updates, Cache fallback for Offline
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  event.respondWith(
    fetch(req)
      .then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const clone = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
        }
        return networkResponse;
      })
      .catch(() => {
        // Offline Fallback
        return caches.match(req).then((cached) => {
          if (cached) return cached;
          if (req.mode === 'navigate') {
            return caches.match('./index.html') || caches.match('./');
          }
        });
      })
  );
});
