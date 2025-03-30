// sw.js - Basic Caching Service Worker

const CACHE_NAME = 'fsm-dashboard-cache-v1'; // Updated cache name
// List of files to cache immediately upon installation
const urlsToCache = [
  '.', // Represents index.html
  'index.html',
  'style.css',
  'script.js',
  'manifest.json',
  'icons/icon.svg', // Path to your SVG icon
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
                        // Log the error but don't necessarily fail the whole install
                        // Sometimes external CDNs might be momentarily unavailable
                        console.error(`Request failed for ${url}: ${response.statusText}`);
                        // Optionally, you could throw an error here to fail the install
                        // if a resource is absolutely critical:
                        // throw new Error(`Request failed for ${url}: ${response.statusText}`);
                        return Promise.resolve(); // Continue caching other files
                    }
                    // Check if response can be cached (e.g., avoid opaque responses if strict)
                    // Basic check: Ensure we have a valid response to cache
                    if(response.status === 200 || response.type === 'basic' || response.type === 'cors') {
                         return cache.put(url, response);
                    } else {
                         console.warn(`Skipping caching for ${url} due to response status/type: ${response.status} / ${response.type}`);
                         return Promise.resolve();
                    }
                }).catch(fetchError => {
                    console.error(`Fetch error for ${url}:`, fetchError);
                    return Promise.resolve(); // Continue caching other files
                });
        });
        // Wait for all essential puts to complete
        return Promise.all(promises);
      })
      .then(() => {
        // Force the waiting service worker to become the active service worker.
        console.log('[Service Worker] App shell caching attempted, skipping waiting.');
        return self.skipWaiting();
      })
      .catch(error => {
         // This catch might handle errors from caches.open()
         console.error('[Service Worker] Caching failed during install setup:', error);
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
  // Only handle GET requests for this caching strategy
  if (event.request.method !== 'GET') {
      // Let non-GET requests pass through to the network
      return;
  }

  // Cache-first strategy
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Return response from cache if found
        if (response) {
          // console.log('[Service Worker] Found in cache:', event.request.url);
          return response;
        }

        // If not in cache, fetch from network
        // console.log('[Service Worker] Not in cache, fetching from network:', event.request.url);
        return fetch(event.request).then(
             networkResponse => {
                // Optionally cache the network response here if desired for future offline use.
                // Be careful caching everything, especially if URLs change or have query params.
                // Example: If you want to cache other successful GET requests dynamically:
                // if (networkResponse && networkResponse.ok && networkResponse.type === 'basic') { // Only cache basic requests from your origin
                //   const responseToCache = networkResponse.clone();
                //   caches.open(CACHE_NAME).then(cache => {
                //     cache.put(event.request, responseToCache);
                //   });
                // }
                return networkResponse;
             }
        ).catch(error => {
            // Handle network fetch errors
            console.error('[Service Worker] Network fetch failed:', error, event.request.url);
            // Optional: Return a custom offline fallback page/resource
            // For example: if (event.request.mode === 'navigate') { // Only for page navigations
            //                 return caches.match('/offline.html'); // Need to cache offline.html during install
            //             }
            // If no fallback, the browser's default offline error will show.
        });
      })
  );
});