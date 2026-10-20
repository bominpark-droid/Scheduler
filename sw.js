// v4: 예전 버전 캐시가 남아 옛 화면이 뜨는 문제를 막기 위해, 이 서비스워커는 스스로를 지우고 캐시를 비운다.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map(k => caches.delete(k)));
    await self.registration.unregister();
    const clients = await self.clients.matchAll({ type: 'window' });
    clients.forEach(c => c.navigate(c.url));
  })());
});
