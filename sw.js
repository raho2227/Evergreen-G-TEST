const CACHE_VERSION = 'evergreen-v3';
const STATIC_CACHE = 'evergreen-static-v3';
const DYNAMIC_CACHE = 'evergreen-dynamic-v3';

const STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

const CDN_ASSETS = [
  'https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css',
  'https://www.gstatic.com/firebasejs/10.9.0/firebase-app-compat.js',
  'https://www.gstatic.com/firebasejs/10.9.0/firebase-firestore-compat.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Install: Pre-cache static assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(STATIC_CACHE).then(cache => {
      // Cache static assets first (critical)
      return cache.addAll(STATIC_ASSETS).then(() => {
        // Then try CDN assets (non-blocking)
        return Promise.allSettled(
          CDN_ASSETS.map(url => 
            cache.add(url).catch(err => {
              console.warn('CDN cache skip:', url, err.message);
            })
          )
        );
      });
    })
  );
  self.skipWaiting();
});

// Activate: Clean old caches
self.addEventListener('activate', event => {
  const currentCaches = [STATIC_CACHE, DYNAMIC_CACHE];
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => !currentCaches.includes(k))
          .map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// Fetch: Smart caching strategy
self.addEventListener('fetch', event => {
  const { request } = event;

  // Skip non-GET, chrome-extension, and WebSocket requests
  if (request.method !== 'GET') return;
  if (request.url.startsWith('chrome-extension')) return;
  if (request.url.includes('firestore.googleapis.com')) return;
  if (request.url.includes('googleapis.com/google.firestore')) return;

  // Strategy for API calls (Google Apps Script) - Network only
  if (request.url.includes('script.google.com')) {
    event.respondWith(fetch(request).catch(() => new Response('{"error":"offline"}', {
      headers: { 'Content-Type': 'application/json' }
    })));
    return;
  }

  // Strategy for CDN assets - Cache first, then network update
  if (request.url.startsWith('https://fonts.') || 
      request.url.startsWith('https://cdnjs.') || 
      request.url.startsWith('https://www.gstatic.com')) {
    event.respondWith(
      caches.match(request).then(cached => {
        const networkFetch = fetch(request).then(response => {
          if (response && response.ok) {
            const clone = response.clone();
            caches.open(STATIC_CACHE).then(cache => cache.put(request, clone));
          }
          return response;
        }).catch(() => cached);

        return cached || networkFetch;
      })
    );
    return;
  }

  // Strategy for static assets - Cache first, network fallback
  event.respondWith(
    caches.match(request).then(cached => {
      if (cached) {
        // Return cache immediately, update in background
        fetch(request).then(response => {
          if (response && response.ok) {
            caches.open(STATIC_CACHE).then(cache => cache.put(request, response));
          }
        }).catch(() => {});
        return cached;
      }

      // Not in cache - fetch from network
      return fetch(request).then(response => {
        if (response && response.ok) {
          const clone = response.clone();
          caches.open(DYNAMIC_CACHE).then(cache => cache.put(request, clone));
        }
        return response;
      }).catch(() => {
        // Offline fallback for navigation
        if (request.mode === 'navigate') {
          return caches.match('./index.html');
        }
        return new Response('', { status: 408 });
      });
    })
  );
});

// Listen for skip waiting message from client
self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
