/* 일일계획표 서비스워커
   목적: 홈화면 설치(PWA) + 오프라인에서도 앱이 열리게 하기
   원칙: Firebase 실시간 통신과 로그인 요청은 절대 가로채지 않는다. */

const CACHE = 'planner-shell-v2';

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
  /* ★ 새 워커가 '대기' 상태에 머무르지 않고 곧바로 교대하게 한다.
     이게 없으면 index.html은 새 버전인데 서비스워커만 옛것인 상태가
     오래 간다. 푸시 수신 코드가 옛 워커에 없으면 알림이 조용히
     사라지는데, 화면상으로는 아무 이상이 없어 원인을 찾기 어렵다. */
  self.skipWaiting();
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

/* ══════════════════════════════════════════════════════════════════
   푸시 알림 수신
   발송기(Cloudflare Worker)가 보낸 암호화 메시지를 받아 안드로이드
   상단바 알림으로 띄운다. 본문 첫 줄이 제목, 나머지가 내용이 된다.
====================================================================== */
self.addEventListener('push', e => {
  let text = '';
  try { text = e.data ? e.data.text() : ''; } catch (_) {}
  if (!text) text = '일일계획표 알림';

  const lines = text.split('\n');
  const title = lines[0] || '일일계획표';
  const body  = lines.slice(1).join('\n').trim();

  e.waitUntil(self.registration.showNotification(title, {
    body,
    icon: './icon-192.png',
    badge: './favicon-32.png',
    tag: 'planner',           // 같은 tag는 덮어써서 알림이 쌓이지 않는다
    renotify: true,
    requireInteraction: false,
    data: { url: './index.html' }
  }));
});

/* 알림을 누르면 이미 열린 계획표로 가고, 없으면 새로 연다 */
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const target = new URL('./index.html', self.location.href).href;
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of all) {
      if (c.url.startsWith(self.registration.scope) && 'focus' in c) return c.focus();
    }
    if (self.clients.openWindow) return self.clients.openWindow(target);
  })());
});

/* 구독이 만료·교체되면 앱이 다음 실행 때 다시 등록하도록 표시만 남긴다 */
self.addEventListener('pushsubscriptionchange', e => {
  e.waitUntil(self.registration.showNotification('일일계획표', {
    body: '알림 등록을 갱신해야 합니다. 앱을 한 번 열어주세요.',
    icon: './icon-192.png', tag: 'planner-resub'
  }));
});
