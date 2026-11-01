/* 일일계획표 서비스워커
   목적: 홈화면 설치(PWA) + 오프라인에서도 앱이 열리게 하기
   원칙: Firebase 실시간 통신과 로그인 요청은 절대 가로채지 않는다. */

const CACHE = 'planner-shell-v1';

/* 앱 껍데기 - 설치 시 미리 받아둔다 */
const CORE = [
  './', './index.html', './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './maskable-192.png',
  './maskable-512.png', './apple-touch-icon.png', './favicon-32.png'
];

/* 캐시해도 되는 외부 스크립트 (Firebase SDK 파일 자체) */
const CDN_OK = ['www.gstatic.com'];

/* 절대 건드리면 안 되는 호스트 - 실시간 DB와 인증 */
const BYPASS = [
  'firebasedatabase.app', 'firebaseio.com',
  'identitytoolkit.googleapis.com', 'securetoken.googleapis.com', 'googleapis.com'
];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await Promise.all(CORE.map(u => c.add(new Request(u, { cache: 'reload' })).catch(() => {})));
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map(k => k === CACHE ? null : caches.delete(k)));
    await self.clients.claim();
  })());
});

/* 새 버전을 즉시 적용하라는 신호 (페이지에서 보냄) */
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  let url;
  try { url = new URL(req.url); } catch (_) { return; }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return;

  /* 실시간 DB / 인증은 손대지 않고 그대로 통과 */
  if (BYPASS.some(h => url.hostname.endsWith(h))) return;

  const sameOrigin = url.origin === self.location.origin;
  const cdnOk = CDN_OK.includes(url.hostname);
  if (!sameOrigin && !cdnOk) return;

  /* 페이지 문서는 네트워크 우선 - 수정본이 바로 반영되게 */
  if (req.mode === 'navigate' || (sameOrigin && url.pathname.endsWith('.html'))) {
    e.respondWith((async () => {
      try {
        const fresh = await fetch(req);
        const c = await caches.open(CACHE);
        c.put('./index.html', fresh.clone());
        return fresh;
      } catch (_) {
        const cached = await caches.match('./index.html');
        return cached || new Response(
          '<meta charset="utf-8"><h1>오프라ᅵᆫ</h1>',
          { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }
    })());
    return;
  }

  /* 그 외 정적 파일은 캐시 우선 + 뒤에서 조용히 갱신 */
  e.respondWith((async () => {
    const cached = await caches.match(req);
    const net = fetch(req).then(res => {
      if (res && (res.ok || res.type === 'opaque')) {
        caches.open(CACHE).then(c => c.put(req, res.clone()));
      }
      return res;
    }).catch(() => null);
    return cached || (await net) || new Response('', { status: 504 });
  })());
});
