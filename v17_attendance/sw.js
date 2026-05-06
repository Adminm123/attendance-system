// ═══════════════════════════════════════════════
// Service Worker — Attendance System
// Cache AI models เท่านั้น — index.html ดึงจาก network ตลอด
// ═══════════════════════════════════════════════

const MODEL_CACHE = 'attendance-models-v2';
const CDN_CACHE   = 'attendance-cdn-v2';

const MODEL_BASE = 'https://cdn.jsdelivr.net/npm/@vladmandic/face-api/model/';
const MODEL_FILES = [
  'ssd_mobilenetv1_model-weights_manifest.json',
  'ssd_mobilenetv1_model-shard1',
  'face_landmark_68_model-weights_manifest.json',
  'face_landmark_68_model-shard1',
  'face_recognition_model-weights_manifest.json',
  'face_recognition_model-shard1',
  'face_recognition_model-shard2',
];
const MODEL_URLS = MODEL_FILES.map(f => MODEL_BASE + f);

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(MODEL_CACHE).then(cache =>
      Promise.allSettled(MODEL_URLS.map(url => cache.add(url).catch(() => {})))
    )
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== MODEL_CACHE && k !== CDN_CACHE)
          .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = event.request.url;

  // ⚠️ index.html → network-only (ไม่ cache → user ได้ update ทันที)
  if (event.request.mode === 'navigate' ||
      url.endsWith('/') ||
      url.endsWith('/index.html')) {
    return; // ปล่อย browser fetch เอง
  }

  // Models → cache-first
  if (url.includes('vladmandic/face-api/model') ||
      url.includes('@vladmandic/face-api/model')) {
    event.respondWith(
      caches.open(MODEL_CACHE).then(cache =>
        cache.match(event.request).then(cached => {
          if (cached) return cached;
          return fetch(event.request).then(res => {
            if (res.ok) cache.put(event.request, res.clone());
            return res;
          }).catch(() => cached);
        })
      )
    );
    return;
  }

  // face-api.js / Tailwind / Fonts → cache-first
  if (url.includes('face-api.min.js') ||
      url.includes('cdn.tailwindcss') ||
      url.includes('fonts.googleapis') ||
      url.includes('fonts.gstatic')) {
    event.respondWith(
      caches.open(CDN_CACHE).then(cache =>
        cache.match(event.request).then(cached => {
          if (cached) return cached;
          return fetch(event.request).then(res => {
            if (res.ok) cache.put(event.request, res.clone());
            return res;
          }).catch(() => cached);
        })
      )
    );
    return;
  }
});
