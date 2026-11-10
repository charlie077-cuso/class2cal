/* 课表转日历 · Service Worker（离线可用 + 网络优先避免版本错配） */
'use strict';

const CACHE = 'class2cal-v9';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './css/styles.css',
  './js/parser.js',
  './js/ics.js',
  './js/app.js',
  './vendor/xlsx.full.min.js',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

function networkFirst(req, fallbackKey) {
  return fetch(req)
    .then((res) => {
      if (res && res.ok && res.type === 'basic') {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
      }
      return res;
    })
    .catch(() => caches.match(fallbackKey || req));
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // 页面导航：网络优先，离线回退到缓存的首页
  if (req.mode === 'navigate') {
    event.respondWith(networkFirst(req, './index.html'));
    return;
  }

  // 静态资源也用网络优先：更新部署后不会出现「新页面配旧脚本」的版本错配；离线时回退缓存
  event.respondWith(networkFirst(req));
});
