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
  event.respondWith(networkFirst(url));
});

// Network first, but a slow network (over 4 s) or a server error page falls back to the saved copy.
function networkFirst(url) {
  return new Promise(function (resolve) {
    var settled = false;
    function finish(r) { if (!settled) { settled = true; resolve(r); } }
    function cached() { return caches.open(CACHE).then(function (c) { return c.match(PAGE); }).catch(function () { return undefined; }); }
    var timer = setTimeout(function () { cached().then(function (hit) { if (hit) finish(hit); }); }, 4000);
    fetch(new Request(url.origin + url.pathname + url.search, { cache: 'no-cache', credentials: 'same-origin' }))
      .then(function (res) {
        if (res && res.ok && res.type === 'basic') {
          var copy = res.clone();
          caches.open(CACHE).then(function (cache) { return cache.put(PAGE, copy); }).catch(function () {});
          clearTimeout(timer);
          return Promise.resolve(clean(res)).then(finish);
        }
        return cached().then(function (hit) { clearTimeout(timer); finish(hit || res); });
      })
      .catch(function () {
        return cached().then(function (hit) { clearTimeout(timer); finish(hit || Response.error()); });
      });
  });
}

