/* Offline fallback for the survey page only (network first). Touches no other app's caches. */
var CACHE = 'seal-survey-v2-1';
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
        return Promise.all(keys.filter(function (k) { return k.indexOf('seal-survey-') === 0 && k !== CACHE; })
          .map(function (k) { return caches.delete(k); }));
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
  var isPage = req.mode === 'navigate' || url.pathname === scopePath || url.pathname === scopePath + 'index.html';
  if (!isPage) return;
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
