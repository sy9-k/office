// SK Hub Office の Google ドライブ（window.OfficeDrive）。store.js が「保存先の 1 つ」として使う
//   ・使う人の Google ドライブの「SK Hub Office」フォルダに、1 ファイルずつ保存する（中身は PC のフォルダと同じ JSON。名前も「名前.skdoc」など）
//   ・権限は drive.file だけ（このアプリが作ったファイル・このアプリで開いたファイルだけを読み書きできる。ほかのファイルは見えない）
//   ・ログインは Google Identity Services のトークン（https://accounts.google.com/gsi/client）。1 時間で切れるので、切れたら
//     ボタンを押してもらって取りなおす（connect()。ポップアップを開くので、押したときにしか呼べない）。トークンはこのブラウザの localStorage に、切れるまで覚える
//   ・CLIENT_ID は Google Cloud の「OAuth クライアント ID（ウェブ）」。空のあいだは Google ドライブを使えない（state() が "unconfigured"）
//     作り方は office/ROADMAP.md の「Google ドライブ」
//   ・ファイルの ID は Google ドライブのファイル ID（store.js では「g:」を前に付ける）
//   どれも Promise を返す（state()・account()・CLIENT_ID だけはすぐ返す）
(function () {
  "use strict";
  // Google Cloud の OAuth クライアント ID（「….apps.googleusercontent.com」）。公開するときは、ここに入れる（クライアント ID は秘密ではない）
  var GOOGLE_CLIENT_ID = "";
  // ここに書く前に試すときは、ブラウザの localStorage の sk_office_google_client に入れてもよい
  var CLIENT_ID = window.OFFICE_GOOGLE_CLIENT_ID || (function () { try { return localStorage.getItem("sk_office_google_client") || ""; } catch (e) { return ""; } })() || GOOGLE_CLIENT_ID;
  var SCOPE = "https://www.googleapis.com/auth/drive.file";
  var API = "https://www.googleapis.com/drive/v3", UPLOAD = "https://www.googleapis.com/upload/drive/v3";
  var FOLDER_NAME = "SK Hub Office", FOLDER_MIME = "application/vnd.google-apps.folder", FILE_MIME = "application/json";
  var KEY = "sk_office_drive"; // { token, exp, email, folder, on }

  function load() { try { return JSON.parse(localStorage.getItem(KEY) || "{}") || {}; } catch (e) { return {}; } }
  function store(o) { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) { /* 覚えられなくても、このページでは使える */ } }
  var saved = load();
  function tokenOk() { return saved.token && saved.exp > Date.now() + 60000; }

  // "unconfigured"（CLIENT_ID がない）・"off"（つないでいない）・"needs-auth"（つないだことはあるが、トークンが切れた）・"ready"
  function state() { return !CLIENT_ID ? "unconfigured" : !saved.on ? "off" : tokenOk() ? "ready" : "needs-auth"; }

  // ---------- ログイン（Google Identity Services） ----------
  var gis = null;
  function loadGis() {
    if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
    if (gis) return gis;
    gis = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = "https://accounts.google.com/gsi/client";
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () { gis = null; reject(new Error("offline")); };
      document.head.appendChild(s);
    });
    return gis;
  }
  // つなぐ（ボタンを押したときに呼ぶ）。hint: ログインの画面で先に選んでおくメールアドレス。やめたら false
  function connect(hint) {
    if (!CLIENT_ID) return Promise.reject(new Error("unconfigured"));
    return loadGis().then(function () {
      return new Promise(function (resolve) {
        var client = google.accounts.oauth2.initTokenClient({
          client_id: CLIENT_ID, scope: SCOPE, hint: hint || saved.email || "", prompt: saved.on ? "" : "consent",
          callback: function (r) {
            if (!r || r.error || !r.access_token || !google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) { resolve(false); return; }
            saved.token = r.access_token; saved.exp = Date.now() + (+r.expires_in || 3600) * 1000; saved.on = true;
            store(saved);
            resolve(true);
          },
          error_callback: function () { resolve(false); }
        });
        client.requestAccessToken();
      });
    }).then(function (ok) {
      if (!ok) return false;
      // だれのドライブか（アカウントの画面に出す）。drive.file でも about の user は読める
      return api("GET", "/about?fields=user(emailAddress,displayName)").then(function (a) {
        saved.email = a.user && a.user.emailAddress || ""; saved.name = a.user && a.user.displayName || "";
        store(saved);
        return true;
      }, function () { return true; });
    });
  }
  function disconnect() {
    var tk = saved.token;
    saved = {};
    store(saved);
    if (tk && window.google && google.accounts && google.accounts.oauth2) google.accounts.oauth2.revoke(tk, function () {});
    return Promise.resolve();
  }

  // ---------- Drive API ----------
  function fail(status, text) { var e = new Error(status === 401 ? "auth" : status === 404 ? "not-found" : "drive-" + status); e.status = status; e.detail = text; return e; }
  function api(method, path, body, opts) {
    opts = opts || {};
    if (!tokenOk()) return Promise.reject(fail(401, "expired"));
    var headers = { Authorization: "Bearer " + saved.token };
    if (opts.type) headers["Content-Type"] = opts.type;
    else if (body !== undefined && typeof body !== "string") { headers["Content-Type"] = "application/json; charset=UTF-8"; body = JSON.stringify(body); }
    return fetch((opts.upload ? UPLOAD : API) + path, { method: method, headers: headers, body: body }).then(function (res) {
      // 401: トークンが切れた・取り消された。つなぎなおしてもらう
      if (res.status === 401) { saved.exp = 0; store(saved); }
      if (!res.ok) return res.text().then(function (x) { throw fail(res.status, x); });
      if (opts.raw) return res.text();
      return res.status === 204 ? null : res.json();
    });
  }
  function q(s) { return "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'"; }
  // 「SK Hub Office」フォルダ（なければ作る。このアプリが作ったフォルダなので drive.file で見える）
  var folderP = null;
  function folder() {
    if (folderP) return folderP;
    folderP = (saved.folder
      ? api("GET", "/files/" + encodeURIComponent(saved.folder) + "?fields=id,trashed").then(function (f) { return f.trashed ? null : f.id; }, function () { return null; })
      : Promise.resolve(null)
    ).then(function (id) {
      if (id) return id;
      return api("GET", "/files?q=" + encodeURIComponent("mimeType=" + q(FOLDER_MIME) + " and name=" + q(FOLDER_NAME) + " and trashed=false") + "&fields=files(id)&spaces=drive").then(function (r) {
        if (r.files && r.files.length) return r.files[0].id;
        return api("POST", "/files?fields=id", { name: FOLDER_NAME, mimeType: FOLDER_MIME }).then(function (f) { return f.id; });
      });
    }).then(function (id) { saved.folder = id; store(saved); return id; }, function (e) { folderP = null; throw e; });
    return folderP;
  }
  var FIELDS = "id,name,modifiedTime,createdTime,size,appProperties,ownedByMe,owners(emailAddress,displayName),capabilities(canEdit)";
  function meta(f) {
    return { id: f.id, name: f.name, app: f.appProperties && f.appProperties.skApp || "", updated: Date.parse(f.modifiedTime) || Date.now(), created: Date.parse(f.createdTime) || Date.now(),
      size: +f.size || 0, mine: f.ownedByMe !== false, owner: f.owners && f.owners[0] ? f.owners[0].displayName || f.owners[0].emailAddress : "", canEdit: !f.capabilities || f.capabilities.canEdit !== false };
  }
  // multipart/related（メタデータと中身を 1 回で送る）
  function multipart(metaObj, text) {
    var b = "skoffice" + Math.random().toString(36).slice(2);
    return { type: "multipart/related; boundary=" + b,
      body: "--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(metaObj) + "\r\n--" + b + "\r\nContent-Type: " + FILE_MIME + "\r\n\r\n" + text + "\r\n--" + b + "--" };
  }

  var Drive = {
    CLIENT_ID: CLIENT_ID,
    state: state,
    account: function () { return { email: saved.email || "", name: saved.name || "" }; },
    connect: connect,
    disconnect: disconnect,
    // 一覧（「SK Hub Office」フォルダの中と、共有されて開いたことのあるもの）: [{ id, name（拡張子つき）, app, updated, created, size, mine, owner, canEdit }]
    list: function () {
      return folder().then(function (fid) {
        var out = [];
        function page(token) {
          return api("GET", "/files?q=" + encodeURIComponent("(" + q(fid) + " in parents or sharedWithMe) and trashed=false and mimeType=" + q(FILE_MIME)) +
            "&fields=" + encodeURIComponent("nextPageToken,files(" + FIELDS + ")") + "&pageSize=200&spaces=drive" + (token ? "&pageToken=" + encodeURIComponent(token) : "")).then(function (r) {
            out = out.concat((r.files || []).map(meta));
            return r.nextPageToken ? page(r.nextPageToken) : out;
          });
        }
        return page();
      });
    },
    // 中身（JSON の文字列）とメタデータ
    get: function (id) {
      return Promise.all([
        api("GET", "/files/" + encodeURIComponent(id) + "?fields=" + encodeURIComponent(FIELDS)),
        api("GET", "/files/" + encodeURIComponent(id) + "?alt=media", undefined, { raw: true })
      ]).then(function (r) { var m = meta(r[0]); m.text = r[1]; return m; });
    },
    create: function (name, app, text) {
      return folder().then(function (fid) {
        var mp = multipart({ name: name, parents: [fid], mimeType: FILE_MIME, appProperties: { skApp: app } }, text);
        return api("POST", "/files?uploadType=multipart&fields=" + encodeURIComponent(FIELDS), mp.body, { upload: true, type: mp.type }).then(meta);
      });
    },
    save: function (id, text) {
      return api("PATCH", "/files/" + encodeURIComponent(id) + "?uploadType=media&fields=" + encodeURIComponent(FIELDS), text, { upload: true, type: FILE_MIME }).then(meta);
    },
    rename: function (id, name) { return api("PATCH", "/files/" + encodeURIComponent(id) + "?fields=" + encodeURIComponent(FIELDS), { name: name }).then(meta); },
    // ごみ箱に入れる（Google ドライブのごみ箱から 30 日は戻せる）
    trash: function (id) { return api("PATCH", "/files/" + encodeURIComponent(id) + "?fields=id", { trashed: true }); }
  };
  window.OfficeDrive = Drive;
})();
