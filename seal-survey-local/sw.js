/* Offline fallback for this survey page only (network first). Touches no other page's caches. */
var FAMILY = 'seal-survey-local-';
var CACHE = FAMILY + 'v3';
var LEGACY = [];
var PAGE = './index.html';

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE)
      .then(function (cache) { return cache.add(new Request(PAGE, { cache: 'no-cache' })); })
      .catch(function () {})
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.filter(function (k) {
          if (k === CACHE) return false;
          return k.indexOf(FAMILY) === 0 || LEGACY.indexOf(k) >= 0;
        }).map(function (k) { return caches.delete(k); }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

function clean(response) {
  if (!response || !response.redirected) return response;
  return response.blob().then(function (body) {
    return new Response(body, { status: response.status, statusText: response.statusText, headers: response.headers });
  });
}

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  var scopePath = new URL('./', self.location.href).pathname;
  if (url.pathname !== scopePath && url.pathname !== scopePath + 'index.html') return;
  event.respondWith(
    fetch(new Request(url.origin + url.pathname + url.search, { cache: 'no-cache', credentials: 'same-origin' }))
      .then(function (res) {
        if (res && res.ok && res.type === 'basic') {
          var copy = res.clone();
          caches.open(CACHE).then(function (cache) { return cache.put(PAGE, copy); }).catch(function () {});
        }
        return clean(res);
      })
      .catch(function () {
        return caches.open(CACHE)
          .then(function (cache) { return cache.match(PAGE); })
          .then(function (hit) { return hit || Response.error(); });
      })
  );
});
