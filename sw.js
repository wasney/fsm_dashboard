// sw.js - Basic Caching Service Worker

const CACHE_NAME = 'fsm-dashboard-cache-v1'; // Keep or update cache name as needed
// List of files to cache immediately upon installation (Relative Paths)
const urlsToCache = [
  '.',             // Represents index.html within /fsm_dashboard/
  'index.html',    // Relative to sw.js
  'style.css',     // Relative to sw.js
  'script.js',     // Relative to sw.js
  'manifest.json', // Relative to sw.js
  'icons/icon.svg',// Relative to sw.js (assumes icons folder inside fsm_dashboard)
  // External libraries remain absolute
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@latest/dist/chart.umd.js'
];

// Install event: Cache the core assets
self.addEventListener('install', event => {
  console.log('[Service Worker] Install');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching app shell');
        const promises = urlsToCache.map(url => {
            // For relative URLs, fetch respects the service worker's path
            return fetch(url, { cache: 'reload' })
                .then(response => {
                    if (!response.ok) {
                        console.error(`Request failed for ${url}: ${response.statusText}`);
                        return Promise.resolve(); // Continue caching other files
                    }
                    if(response.status === 200 || response.type === 'basic' || response.type === 'cors') {
                         return cache.put(url, response);
                    } else {
                         console.warn(`Skipping caching for ${url} due to response status/type: ${response.status} / ${response.type}`);
                         return Promise.resolve();
                    }
                }).catch(fetchError => {
                    console.error(`Fetch error for ${url}:`, fetchError);
                    return Promise.resolve();
                });
        });
        return Promise.all(promises);
      })
      .then(() => {
        console.log('[Service Worker] App shell caching attempted, skipping waiting.');
        return self.skipWaiting();
      })
      .catch(error => {
         console.error('[Service Worker] Caching failed during install setup:', error);
      })
  );
});

// Activate event: Clean up old caches
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activate');
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (!cacheWhitelist.includes(cacheName)) {
            console.log('[Service Worker] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
         console.log('[Service Worker] Claiming clients.');
         return self.clients.claim();
    })
  );
});

// Fetch event: Serve from cache first, fall back to network
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') { return; }
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) { return response; } // Serve from cache
        // Fetch from network
        return fetch(event.request).then( networkResponse => {
                // Optional: Cache dynamic GET requests here if needed
                return networkResponse;
             }
        ).catch(error => {
            console.error('[Service Worker] Network fetch failed:', error, event.request.url);
            // Optional: Return offline fallback
        });
      })
  );
});
