/* Presente service worker — lets the app install and load offline. */
var CACHE = 'presente-v9';
var PRECACHE = ['./', 'manifest.json', 'icon-192.png', 'icon-512.png',
  'checkin.html', 'enroll.html', 'submit.html', 'jsqr.js', 'qrcode.min.js', 'jspdf.umd.min.js', 'jspdf.plugin.autotable.min.js'];
var CDN_HOSTS = ['cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(PRECACHE); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  var sameOrigin = url.origin === self.location.origin;
  var isCdn = CDN_HOSTS.indexOf(url.hostname) !== -1;
  /* Anything else (e.g. Google Sheet sync) goes straight to the network. */
  if (!sameOrigin && !isCdn) return;

  if (req.mode === 'navigate') {
    /* The app page: fresh from the network when online, cached copy when offline. */
    e.respondWith(
      fetch(req).then(function (res) {
        if (!res.ok) {
          /* Server is broken (e.g. 404 during an outage): show the good cached app instead. */
          return caches.match(req).then(function (m) {
            return m || caches.match('./') || res;
          });
        }
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put('./', copy); });
        return res;
      }).catch(function () {
        return caches.match(req).then(function (m) { return m || caches.match('./'); });
      })
    );
    return;
  }

  /* Libraries, fonts, icons: cache-first. */
  e.respondWith(
    caches.match(req).then(function (m) {
      if (m) return m;
      return fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
        return res;
      });
    })
  );
});
