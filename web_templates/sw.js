const CACHE_NAME = 'attendance-tracker-v1.0.0';
const urlsToCache = [
  '/',
  '/static/icon-192.png',
  '/static/icon-512.png',
  '/manifest.json',
  'https://cdnjs.cloudflare.com/ajax/libs/socket.io/4.7.2/socket.io.js'
];

// Install event - cache resources
self.addEventListener('install', event => {
  console.log('Service Worker: Installing...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Service Worker: Caching files');
        return cache.addAll(urlsToCache);
      })
      .then(() => {
        console.log('Service Worker: Cached all files successfully');
        return self.skipWaiting();
      })
      .catch(err => {
        console.error('Service Worker: Cache failed', err);
      })
  );
});

// Activate event - clean up old caches
self.addEventListener('activate', event => {
  console.log('Service Worker: Activating...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('Service Worker: Deleting old cache', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
      console.log('Service Worker: Activated');
      return self.clients.claim();
    })
  );
});

// Fetch event - serve from cache, fallback to network
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);
  
  // Handle API requests - always try network first, then fallback
  if (url.pathname.startsWith('/api/')) {
    event.respondWith(
      fetch(request)
        .then(response => {
          // Clone the response before returning it
          const responseClone = response.clone();
          
          // Cache successful API responses for offline fallback
          if (response.status === 200) {
            caches.open(CACHE_NAME).then(cache => {
              cache.put(request, responseClone);
            });
          }
          
          return response;
        })
        .catch(() => {
          // If network fails, try to serve from cache
          return caches.match(request).then(response => {
            if (response) {
              // Add offline indicator to cached response
              return response;
            }
            
            // Return offline fallback for API requests
            return new Response(
              JSON.stringify({
                success: false,
                error: 'You are offline. Please check your connection.',
                offline: true
              }),
              {
                status: 503,
                statusText: 'Service Unavailable',
                headers: {
                  'Content-Type': 'application/json'
                }
              }
            );
          });
        })
    );
    return;
  }
  
  // Handle static assets and pages - cache first strategy
  event.respondWith(
    caches.match(request)
      .then(response => {
        // Return cached version if available
        if (response) {
          console.log('Service Worker: Serving from cache', request.url);
          return response;
        }
        
        // Otherwise fetch from network
        console.log('Service Worker: Fetching from network', request.url);
        return fetch(request)
          .then(response => {
            // Don't cache non-successful responses
            if (!response || response.status !== 200 || response.type !== 'basic') {
              return response;
            }
            
            // Clone the response
            const responseToCache = response.clone();
            
            // Add to cache
            caches.open(CACHE_NAME)
              .then(cache => {
                cache.put(request, responseToCache);
              });
            
            return response;
          })
          .catch(() => {
            // Return offline page for navigation requests
            if (request.mode === 'navigate') {
              return caches.match('/').then(response => {
                return response || new Response(
                  `<!DOCTYPE html>
                  <html>
                  <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Offline - Attendance Tracker</title>
                    <style>
                      body {
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        min-height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        text-align: center;
                        padding: 20px;
                      }
                      .offline-container {
                        max-width: 400px;
                        background: rgba(255,255,255,0.1);
                        padding: 40px;
                        border-radius: 20px;
                        backdrop-filter: blur(10px);
                      }
                      h1 { margin-bottom: 20px; }
                      .retry-btn {
                        background: #007bff;
                        color: white;
                        border: none;
                        padding: 12px 24px;
                        border-radius: 8px;
                        cursor: pointer;
                        margin-top: 20px;
                      }
                    </style>
                  </head>
                  <body>
                    <div class="offline-container">
                      <h1>📱 You're Offline</h1>
                      <p>The attendance tracker is currently unavailable. Please check your internet connection and try again.</p>
                      <button class="retry-btn" onclick="window.location.reload()">🔄 Try Again</button>
                    </div>
                  </body>
                  </html>`,
                  {
                    headers: {
                      'Content-Type': 'text/html'
                    }
                  }
                );
              });
            }
            
            // For other requests, return a generic offline response
            return new Response('Offline', { status: 503 });
          });
      })
  );
});

// Background sync for offline actions
self.addEventListener('sync', event => {
  if (event.tag === 'attendance-sync') {
    console.log('Service Worker: Background sync for attendance');
    event.waitUntil(
      // Implement background sync logic here if needed
      Promise.resolve()
    );
  }
});

// Push notifications (for future use)
self.addEventListener('push', event => {
  const options = {
    body: event.data ? event.data.text() : 'Attendance reminder',
    icon: '/static/icon-192.png',
    badge: '/static/icon-72.png',
    vibrate: [100, 50, 100],
    data: {
      dateOfArrival: Date.now(),
      primaryKey: 'attendance-notification'
    },
    actions: [
      {
        action: 'check-in',
        title: 'Check In',
        icon: '/static/icon-96.png'
      },
      {
        action: 'check-out',
        title: 'Check Out',
        icon: '/static/icon-96.png'
      }
    ]
  };
  
  event.waitUntil(
    self.registration.showNotification('Attendance Tracker', options)
  );
});

// Handle notification clicks
self.addEventListener('notificationclick', event => {
  console.log('Service Worker: Notification click received');
  
  event.notification.close();
  
  if (event.action === 'check-in') {
    event.waitUntil(
      clients.openWindow('/?action=checkin')
    );
  } else if (event.action === 'check-out') {
    event.waitUntil(
      clients.openWindow('/?action=checkout')
    );
  } else {
    event.waitUntil(
      clients.openWindow('/')
    );
  }
});

// Message handling from main thread
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    console.log('Service Worker: Received skip waiting message');
    self.skipWaiting();
  }
});

// Error handling
self.addEventListener('error', event => {
  console.error('Service Worker: Error occurred', event.error);
});

self.addEventListener('unhandledrejection', event => {
  console.error('Service Worker: Unhandled promise rejection', event.reason);
});

console.log('Service Worker: Script loaded');
