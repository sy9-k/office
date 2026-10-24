// SK Hub Office のサービスワーカー（PWA。一度開いたら、オフラインでも開ける）
//   ・ページ・CSS・JS・アイコン（sk-hub.jp のもの）: ネットにつながっていれば新しいものを使い、つながらない（3 秒かかる）ときは前に取っておいたものを出す
//     （直したものが、次に開いたときにすぐ出るように）
//   ・フォント（Google Fonts）: 取っておいたものをすぐ使い、後ろで新しいものに入れかえる
//   ・置く場所は Office のいちばん上（office.sk-hub.jp/sw.js・手元では /office/sw.js）。scope も同じ
//   画面の部品を大きく変えたときは VERSION を上げる（古い取っておきを消す）
const VERSION = "sk-office-v18";
const BASE = new URL("./", self.location).pathname; // "/" または "/office/"
const SHELL = ["", "doc/", "sheet/", "slide/", "about/", "manifest.json", "doc/manifest.json", "sheet/manifest.json", "slide/manifest.json",
  "shared/office.css", "shared/office.js", "shared/store.js", "shared/drive.js", "shared/templates.js",
  "shared/doc.js", "shared/sheet.js", "shared/slide.js", "shared/formula.js", "shared/chart.js", "shared/zip.js", "shared/vendor/opentype.min.js",
  "icons/icon.svg", "icons/icon-192.png", "icons/doc-192.png", "icons/sheet-192.png", "icons/slide-192.png"].map((p) => BASE + p)
  .concat(["/toolbox/shared/m3.css", "/toolbox/shared/m3.js", "/assets/i18n.js", "/assets/fonts.js", "/assets/coming-soon.js", "/assets/maintenance.js"]);
const FONT_HOSTS = ["fonts.googleapis.com", "fonts.gstatic.com"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((cache) => Promise.allSettled(SHELL.map((u) => cache.add(u)))).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k.startsWith("sk-office-") && k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const same = url.origin === self.location.origin;

  // sk-hub.jp のもの: ネットを先に（3 秒待ってもだめなら、取っておいたもの）
  if (same) {
    const page = req.mode === "navigate";
    if (page && !url.pathname.startsWith(BASE)) return;
    const key = page ? url.pathname : url.pathname + url.search;
    e.respondWith((async () => {
      const cache = await caches.open(VERSION);
      try {
        // cache: "no-cache" … ブラウザの取っておき（HTTP キャッシュ）を使わず、変わったかをいつも聞く（変わっていなければ 304 で軽い）
        const fresh = page ? fetch(req.url, { cache: "no-cache", credentials: "same-origin" }) : fetch(req, { cache: "no-cache" });
        const res = await Promise.race([fresh, new Promise((_, rej) => setTimeout(() => rej(new Error("slow")), 3000))]);
        // ページが移った（/office/doc → /office/doc/ など）ときは、ブラウザにも移ってもらう
        if (page && res.redirected) return Response.redirect(res.url, 302);
        if (res.ok) cache.put(key, res.clone());
        return res;
      } catch (err) {
        return (await cache.match(key)) || (await cache.match(url.pathname)) || (page && (await cache.match(BASE))) || Response.error();
      }
    })());
    return;
  }

  // フォント: 取っておいたものをすぐ使い、後ろで新しくする
  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith((async () => {
      const cache = await caches.open(VERSION);
      const hit = await cache.match(req);
      const fresh = fetch(req).then((res) => {
        if (res.ok || res.type === "opaque") cache.put(req, res.clone());
        return res;
      }).catch(() => hit || Response.error());
      return hit || fresh;
    })());
  }
});
