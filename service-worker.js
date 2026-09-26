/* =========================================================================
   SERVICE WORKER — Personal Library Manager & Spine Studio
   -------------------------------------------------------------------------
   Strategy:
     • App shell (index.html, manifest)  →  cache-first with network fallback
     • CDN assets (SheetJS)              →  cache-first, cached on first use
     • Google Sheets / GViz endpoints    →  network only (never cached)
     • Anything else                     →  network with cached fallback

   Bump SW_VERSION below whenever index.html changes significantly, so
   clients pick up the new app shell on their next visit.
   ========================================================================= */

const SW_VERSION = 'v1.0.3';

const SHELL_CACHE   = `plm-shell-${SW_VERSION}`;
const RUNTIME_CACHE = `plm-runtime-${SW_VERSION}`;

/* App shell — cached during install */
const SHELL_ASSETS = [
  './',
  './index.html',
  './manifest.json'
];

/* External data hosts that must NEVER be cached */
const NETWORK_ONLY_HOSTS = [
  'docs.google.com',
  'googleusercontent.com',
  'script.google.com'
];

/* ============================ INSTALL =================================== */

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(SHELL_CACHE);

    // Cache each asset independently so a single 404 doesn't abort install
    await Promise.all(SHELL_ASSETS.map(async (url) => {
      try {
        const res = await fetch(url, { cache: 'no-cache' });
        if (res.ok) await cache.put(url, res);
      } catch (err) {
        console.warn('[SW] Pre-cache failed for', url, err);
      }
    }));

    // Activate immediately, don't wait for old tabs to close
    await self.skipWaiting();
  })());
});

/* ============================ ACTIVATE ================================== */

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    // Purge caches from older SW versions
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter(k => k !== SHELL_CACHE && k !== RUNTIME_CACHE)
        .map(k => caches.delete(k))
    );

    // Take control of currently open pages
    await self.clients.claim();
  })());
});

/* ============================ FETCH ===================================== */

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only intercept GET requests
  if (req.method !== 'GET') return;

  let url;
  try { url = new URL(req.url); }
  catch { return; }

  // Skip non-HTTP(S) schemes (chrome-extension, file, data, etc.)
  if (!/^https?:$/.test(url.protocol)) return;

  // Google Sheets / GViz — never intercept; let the browser handle it live
  if (NETWORK_ONLY_HOSTS.some(host => url.hostname === host || url.hostname.endsWith('.' + host))) {
    return;
  }

  /* ---------- Same-origin (the app itself) — cache-first ---------- */
  if (url.origin === self.location.origin) {
    event.respondWith((async () => {
      const cached = await caches.match(req, { ignoreSearch: true });
      if (cached) return cached;

      try {
        const res = await fetch(req);
        // Cache successful responses for later
        if (res && res.ok && res.type === 'basic') {
          const cache = await caches.open(SHELL_CACHE);
          cache.put(req, res.clone());
        }
        return res;
      } catch (err) {
        // Offline and not cached — fall back to index.html for navigations
        if (req.mode === 'navigate') {
          const fallback = await caches.match('./index.html', { ignoreSearch: true });
          if (fallback) return fallback;
        }
        throw err;
      }
    })());
    return;
  }

  /* ---------- Cross-origin (CDN assets) — cache-first ---------- */
  event.respondWith((async () => {
    const cached = await caches.match(req);
    if (cached) return cached;

    try {
      const res = await fetch(req);
      if (res && (res.ok || res.type === 'opaque')) {
        const cache = await caches.open(RUNTIME_CACHE);
        cache.put(req, res.clone());
      }
      return res;
    } catch (err) {
      // Last-ditch effort: return whatever we cached earlier
      const fallback = await caches.match(req);
      if (fallback) return fallback;
      throw err;
    }
  })());
});

/* ============================ MESSAGES ================================== */

self.addEventListener('message', (event) => {
  const data = event.data || {};
  if (data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (data.type === 'GET_VERSION' && event.ports && event.ports[0]) {
    event.ports[0].postMessage({ version: SW_VERSION });
  }
});
