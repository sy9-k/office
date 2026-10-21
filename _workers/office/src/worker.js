// SK Hub Office（https://office.sk-hub.jp/）
// ページは置かず、sk-hub.jp（GitHub Pages）の /office/ をそのつど取ってきて返す中継役（admin.sk-hub.jp・support.sk-hub.jp と同じしくみ）。
// なので、Office を直すときは sk-hub.jp のリポジトリにプッシュするだけでよい（この Worker の deploy はいらない）。
//
//   office.sk-hub.jp/             → sk-hub.jp/office/（入り口）
//   office.sk-hub.jp/doc/         → sk-hub.jp/office/doc/（Sheet・Slide も同じ）
//   office.sk-hub.jp/shared/…     → sk-hub.jp/office/shared/…
//   office.sk-hub.jp/assets/…     → sk-hub.jp/assets/…（サイトと共通の部品は同じパスのまま。SHARED）
//   office.sk-hub.jp/office/…     → /… に移す
//
// sk-hub.jp とは別の場所なので、ファイル（localStorage の sk_office_files）・このブラウザの設定はサイトと分かれる。
const ORIGIN = "https://sk-hub.jp";
const SHARED = [/^\/assets\//, /^\/toolbox\/shared\//, /^\/frameworks\//];
const CACHE_SECONDS = 60; // GitHub Pages から取ったものを Cloudflare に置いておく時間（直してから出るまで最大これだけかかる）

export default {
  async fetch(request) {
    const url = new URL(request.url);
    if (request.method !== "GET" && request.method !== "HEAD") return new Response("Method Not Allowed", { status: 405, headers: { Allow: "GET, HEAD" } });

    const path = url.pathname;
    if (path === "/office" || path.startsWith("/office/")) {
      url.pathname = path.slice("/office".length) || "/";
      return Response.redirect(url.toString(), 301);
    }
    // 開発中は検索に出さない（公開するときに Allow にして、ページの noindex も外す）
    if (path === "/robots.txt") return new Response("User-agent: *\nDisallow: /\n", { headers: { "Content-Type": "text/plain; charset=utf-8" } });

    const shared = SHARED.some((re) => re.test(path));
    const upstream = new URL((shared ? "" : "/office") + path + url.search, ORIGIN);
    const res = await fetch(upstream.toString(), {
      method: request.method,
      redirect: "manual",
      headers: { "User-Agent": "sk-office-worker", "Accept": request.headers.get("Accept") || "*/*" },
      cf: { cacheTtl: CACHE_SECONDS, cacheEverything: true }
    });

    // GitHub Pages の移動（/office/doc → /office/doc/ など）を、このアドレスに合わせる
    if (res.status >= 300 && res.status < 400) {
      const loc = res.headers.get("Location") || "";
      let target = loc;
      try {
        const u = new URL(loc, upstream);
        if (u.origin === ORIGIN && (u.pathname === "/office" || u.pathname.startsWith("/office/"))) {
          target = url.origin + (u.pathname.slice("/office".length) || "/") + u.search + u.hash;
        } else if (u.origin === ORIGIN && shared) {
          target = url.origin + u.pathname + u.search + u.hash;
        } else {
          target = u.toString();
        }
      } catch (e) { /* そのまま返す */ }
      return new Response(null, { status: res.status, headers: { Location: target } });
    }

    const out = new Response(res.body, res);
    out.headers.set("Cache-Control", `public, max-age=${CACHE_SECONDS}`);
    out.headers.set("X-Robots-Tag", "noindex, nofollow");
    out.headers.set("Referrer-Policy", "strict-origin-when-cross-origin");
    out.headers.delete("x-github-request-id");
    out.headers.delete("x-served-by");
    out.headers.delete("x-cache");
    out.headers.delete("x-cache-hits");
    out.headers.delete("x-timer");
    out.headers.delete("x-fastly-request-id");
    out.headers.delete("via");
    return out;
  }
};
