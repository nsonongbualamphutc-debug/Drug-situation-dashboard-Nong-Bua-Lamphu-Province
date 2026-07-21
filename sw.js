/* Service Worker — แดชบอร์ดยาเสพติด นภ.
   กลยุทธ์: network-first (เอาของใหม่จากเน็ตก่อนเสมอ) ถ้าออฟไลน์ค่อยใช้แคช
   *** เลื่อนเลขเวอร์ชันทุกครั้งที่อยากบังคับล้างแคชเก่า *** */
const CACHE = 'nbl-drug-v5';
const SHELL = ['./','./index.html','./input.html','./import.html',
  './manifest.json','./icon-192.png','./icon-512.png','./favicon.png'];

self.addEventListener('install', e => {
  // cache แบบไม่ atomic: ไฟล์ไหนหายก็ไม่ทำให้ติดตั้งล้มทั้งหมด
  e.waitUntil(
    caches.open(CACHE)
      .then(c => Promise.allSettled(SHELL.map(u => c.add(u))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return; // คำขอข้ามโดเมน (Apps Script/CDN/แผนที่) ปล่อยผ่าน

  // network-first: ลองดึงของใหม่ก่อน สำเร็จก็อัปเดตแคช, ล้มเหลว(ออฟไลน์)ค่อยใช้แคช
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});
