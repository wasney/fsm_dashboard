// sw.js - Basic Caching Service Worker
// Timestamp: 2025-05-24T15:10:00EDT
// Summary: Added Geographic Map View (Leaflet.js) and minor UI consistency improvements.

const CACHE_NAME = 'fsm-dashboard-cache-v1'; // Keep or update cache name as needed
// List of files to cache immediately upon installation (Relative Paths)
const urlsToCache = [
  '.',             // Represents index.html
  'index.html',
  'style.css',
  'script.js',
  'manifest.json',
  'icons/icon.svg',
  // External libraries
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@latest/dist/chart.umd.js',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css', // Leaflet CSS
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'    // Leaflet JS
];

// Install event: Cache the core assets
self.addEventListener('install', event => {
  console.log('[Service Worker] Install');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching app shell');
        const promises = urlsToCache.map(url => {
            return fetch(url, { cache: 'reload' }) 
                .then(response => {
                    if (!response.ok) {
                        console.error(`[Service Worker] Request failed for ${url} during cache: ${response.statusText}`);
                        return Promise.resolve(); 
                    }
                    if(response.status === 200 || response.type === 'basic' || response.type === 'cors') {
                         return cache.put(url, response);
                    } else {
                         console.warn(`[Service Worker] Skipping caching for ${url} due to response status/type: ${response.status} / ${response.type}`);
                         return Promise.resolve();
                    }
                }).catch(fetchError => {
                    console.error(`[Service Worker] Fetch error for ${url} during cache:`, fetchError);
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
      .then(cachedResponse => {
        if (cachedResponse) {
          return cachedResponse;
        }
        return fetch(event.request).then(
          networkResponse => {
            return networkResponse;
          }
        ).catch(error => {
          console.error('[Service Worker] Network fetch failed for:', event.request.url, error);
          if (event.request.mode === 'navigate') {
            console.log('[Service Worker] Fetch failed for navigation, attempting to serve cached index.html');
            return caches.match('index.html'); 
          }
          return undefined; 
        });
      })
  );
});
