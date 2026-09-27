/**
 * Araij Manager Pro - PWA Service Worker (100% Offline Engine)
 * v4.4 — Fixed: API exclusion, Background Sync, proper cache strategy
 */

const CACHE_NAME = 'araij-manager-v4.4';
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

// ✅ مسارات يجب استثناؤها من الكاش دائماً (API calls تُعيد دايماً من الشبكة)
const BYPASS_CACHE_PATTERNS = [
  /\/api\//,
  /\/ws$/,
  /script\.google\.com/,
  /docs\.google\.com/,
  /railway\.app\/api/
];

function shouldBypassCache(url) {
  return BYPASS_CACHE_PATTERNS.some(pattern => pattern.test(url));
}

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
        // Cache individually to avoid one failure killing all
        for (const asset of CORE_ASSETS) {
          try { await cache.add(asset); } catch(e) {}
        }
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

// 3. FETCH EVENT - Network-First مع استثناء الـ API routes
self.addEventListener('fetch', (event) => {
  const req = event.request;

  // ✅ FIX: تجاهل غير GET requests
  if (req.method !== 'GET') return;

  const url = req.url;

  // ✅ FIX: API calls و WebSocket تُعيد دائماً من الشبكة — لا كاش أبداً
  if (shouldBypassCache(url)) {
    event.respondWith(
      fetch(req).catch(() => {
        // إذا فشل الـ API وهو أوفلاين، نرجع response فارغ مناسب
        if (url.includes('/api/students') || url.includes('/api/get_students')) {
          return new Response(JSON.stringify({ status: 'offline', students: [], count: 0 }), {
            headers: { 'Content-Type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ status: 'offline' }), {
          headers: { 'Content-Type': 'application/json' }
        });
      })
    );
    return;
  }

  // ✅ Static Assets: Network-First مع Cache Fallback
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
        // Offline Fallback — نرجع من الكاش
        return caches.match(req).then((cached) => {
          if (cached) return cached;
          // للصفحات، نرجع index.html
          if (req.mode === 'navigate') {
            return caches.match('./index.html') || caches.match('./');
          }
        });
      })
  );
});

// 4. ✅ BACKGROUND SYNC EVENT — يُرسل العمليات المتراكمة أوفلاين لما يرجع النت
self.addEventListener('sync', (event) => {
  console.log('[ServiceWorker] Background sync triggered:', event.tag);

  if (event.tag === 'araij-offline-sync') {
    event.waitUntil(
      self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clients => {
        // إرسال رسالة لكل نافذة مفتوحة لتُطلق الـ flush
        clients.forEach(client => {
          client.postMessage({ type: 'BACKGROUND_SYNC_TRIGGER', tag: 'araij-offline-sync' });
        });
      })
    );
  }
});

// 5. ✅ PUSH NOTIFICATIONS (للإشعارات المستقبلية)
self.addEventListener('push', (event) => {
  if (!event.data) return;
  try {
    const data = event.data.json();
    const title = data.title || 'سنتر الارائج ⚡';
    const options = {
      body: data.body || 'تم تحديث البيانات',
      icon: './icon-192.png',
      badge: './icon-192.png',
      dir: 'rtl',
      lang: 'ar',
      vibrate: [100, 50, 150],
      data: { url: data.url || '/' }
    };
    event.waitUntil(self.registration.showNotification(title, options));
  } catch (e) {}
});

// فتح التطبيق عند الضغط على الإشعار
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const url = event.notification.data?.url || '/';
  event.waitUntil(
    self.clients.matchAll({ type: 'window' }).then(clients => {
      const existing = clients.find(c => c.url.includes(self.location.origin));
      if (existing) {
        existing.focus();
        existing.navigate(url);
      } else {
        self.clients.openWindow(url);
      }
    })
  );
});
