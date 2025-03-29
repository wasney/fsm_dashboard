// sw.js - Basic Caching Service Worker

const CACHE_NAME = 'fsm-dashboard-cache-v1'; // Updated cache name
// List of files to cache immediately upon installation
const urlsToCache = [
  '.', // Represents index.html
  'index.html',
  'style.css',
  'script.js',
  'manifest.json',
  'icons/icon.svg', // Added SVG icon path
  // Add external libraries (use specific versions for better cache control)
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@latest/dist/chart.umd.js' // Use specific version if possible, e.g., chart.js@4.4.1
];

// Install event: Cache the core assets
self.addEventListener('install', event => {
  console.log('[Service Worker] Install');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching app shell');
        // Use fetch with cache: 'reload' to bypass browser cache during install
        const promises = urlsToCache.map(url => {
            return fetch(url, { cache: 'reload' })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`Request failed for ${url}: ${response.statusText}`);
                    }
                    return cache.put(url, response);
                });
        });
        return Promise.all(promises);
      })
      .then(() => {
        // Force the waiting service worker to become the active service worker.
        console.log('[Service Worker] App shell cached, skipping waiting.');
        return self.skipWaiting();
      })
      .catch(error => {
         console.error('[Service Worker] Cache addAll/put failed during install:', error);
      })
  );
});

// Activate event: Clean up old caches
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activate');
  const cacheWhitelist = [CACHE_NAME]; // Keep only the current cache version
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
         // Tell the active service worker to take control of the page immediately.
         console.log('[Service Worker] Claiming clients.');
         return self.clients.claim();
    })
  );
});

// Fetch event: Serve from cache first, fall back to network
self.addEventListener('fetch', event => {
  // Only handle GET requests for caching strategy
  if (event.request.method !== 'GET') {
      return;
  }

  // Cache-first strategy
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) {
          // console.log('[Service Worker] Found in cache:', event.request.url);
          return response; // Serve from cache
        }

        // console.log('[Service Worker] Not in cache, fetching from network:', event.request.url);
        // Optional: Clone the request to store in cache later if needed
        // const fetchRequest = event.request.clone();

        return fetch(event.request).then(
             // Return the network response directly
             networkResponse => {
                // Optional: Cache the newly fetched resource if needed
                // Check if we received a valid response
                // if(!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
                //   return networkResponse;
                // }
                // const responseToCache = networkResponse.clone();
                // caches.open(CACHE_NAME)
                //   .then(cache => {
                //     cache.put(event.request, responseToCache);
                //   });
                return networkResponse;
             }
        );
      })
      .catch(error => {
        console.error('[Service Worker] Fetch failed:', error);
        // Optional: Return a fallback offline page/resource
        // For example: return caches.match('/offline.html');
      })
  );
});
