// テスト用の Google ドライブ（本物の代わり）。drive.js より先に読み込む
//   ・OFFICE_GOOGLE_CLIENT_ID を入れて、Google ドライブを使える状態にする
//   ・google.accounts.oauth2（ログイン）を、すぐトークンを返すものにする
//   ・fetch のうち www.googleapis.com/drive・upload だけを、このファイルの中のドライブで答える（ほかはそのまま）
//   window.DRIVE_MOCK … { files: { id: { name, mimeType, parents, trashed, text, appProperties, modifiedTime, createdTime } }, calls: [], token, nextToken }
(function () {
  "use strict";
  window.OFFICE_GOOGLE_CLIENT_ID = "test-client";
  var M = window.DRIVE_MOCK = { files: {}, calls: [], token: "tok-1", nextToken: "tok-1", seq: 0 };
  window.google = { accounts: { oauth2: {
    initTokenClient: function (cfg) {
      return { requestAccessToken: function () { setTimeout(function () { M.token = M.nextToken; cfg.callback({ access_token: M.nextToken, expires_in: 3600, scope: cfg.scope }); }, 0); } };
    },
    hasGrantedAllScopes: function () { return true; },
    revoke: function (tk, cb) { if (cb) cb(); }
  } } };
  function json(o, status) { return Promise.resolve(new Response(JSON.stringify(o), { status: status || 200, headers: { "Content-Type": "application/json" } })); }
  function meta(id) {
    var f = M.files[id];
    return { id: id, name: f.name, mimeType: f.mimeType, modifiedTime: f.modifiedTime, createdTime: f.createdTime, size: String((f.text || "").length), appProperties: f.appProperties, ownedByMe: true, owners: [{ emailAddress: "test@example.com", displayName: "テストの人" }], capabilities: { canEdit: true }, trashed: !!f.trashed };
  }
  function touch(f) { f.modifiedTime = new Date(Date.now() + (M.seq++)).toISOString(); }
  var real = window.fetch.bind(window);
  window.fetch = function (url, opts) {
    url = String(url); opts = opts || {};
    var m = /^https:\/\/www\.googleapis\.com\/(upload\/)?drive\/v3(\/[^?]*)(?:\?(.*))?$/.exec(url);
    if (!m) return real(url, opts);
    var method = opts.method || "GET", path = m[2], qs = new URLSearchParams(m[3] || ""), auth = (opts.headers || {}).Authorization || "";
    M.calls.push({ method: method, path: path, upload: !!m[1], qs: m[3] || "" });
    if (auth !== "Bearer " + M.token) return json({ error: { code: 401 } }, 401);
    if (path === "/about") return json({ user: { emailAddress: "test@example.com", displayName: "テストの人" } });
    var fm = /^\/files\/([^/]+)$/.exec(path), id = fm && decodeURIComponent(fm[1]);
    if (path === "/files" && method === "GET") {
      var q = qs.get("q") || "", list = Object.keys(M.files).filter(function (k) {
        var f = M.files[k];
        if (f.trashed) return false;
        var mt = /mimeType='([^']+)'/.exec(q), nm = /name='([^']+)'/.exec(q), par = /'([^']+)' in parents/.exec(q);
        if (mt && f.mimeType !== mt[1]) return false;
        if (nm && f.name !== nm[1]) return false;
        if (par && (f.parents || []).indexOf(par[1]) < 0) return false;
        return true;
      });
      return json({ files: list.map(meta) });
    }
    if (path === "/files" && method === "POST") {
      var body = opts.body, f;
      if (m[1]) {
        // multipart: 1 つ目がメタデータ、2 つ目が中身
        var bd = /boundary=(\S+)/.exec((opts.headers || {})["Content-Type"])[1], parts = body.split("--" + bd).slice(1, -1).map(function (p) { return p.replace(/^\r\n/, "").split("\r\n\r\n").slice(1).join("\r\n\r\n").replace(/\r\n$/, ""); });
        f = JSON.parse(parts[0]); f.text = parts[1];
      } else f = JSON.parse(body);
      var nid = "f" + (M.seq++);
      f.createdTime = new Date().toISOString(); touch(f);
      M.files[nid] = f;
      return json(meta(nid));
    }
    if (id && !M.files[id]) return json({ error: { code: 404 } }, 404);
    if (id && method === "GET") { if (qs.get("alt") === "media") return Promise.resolve(new Response(M.files[id].text || "")); return json(meta(id)); }
    if (id && method === "PATCH") {
      var x = M.files[id];
      if (m[1]) { x.text = opts.body; M.saves = (M.saves || 0) + 1; }
      else { var p = JSON.parse(opts.body); if (p.name) x.name = p.name; if (p.trashed) x.trashed = true; }
      touch(x);
      // 保存は少し時間がかかることにする（続けて保存したときに、まとめるかを確かめる）
      return new Promise(function (res) { setTimeout(function () { res(json(meta(id))); }, 30); }).then(function (r) { return r; });
    }
    return json({ error: { code: 400 } }, 400);
  };
})();
