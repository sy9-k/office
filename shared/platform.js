// SK Hub Office が動いている場所（window.OfficePlatform）
//   ・"web": ブラウザ（PWA を含む）。 ・"desktop": Tauri のデスクトップ版（desktop/）
//   画面と保存のしくみは 1 本のまま、この機能のフラグ（features）で分ける。デスクトップ版にしかできないことは、ここに足していく
//   Web 版では何も足さず、ほかの部品は features を見るだけ（desktop 以外では invoke を呼ばない）
//   Rust 側の命令は desktop/src-tauri/src/lib.rs
(function () {
  "use strict";
  var T = window.__TAURI__;
  var desktop = !!(T && T.core && typeof T.core.invoke === "function");
  var Platform = {
    kind: desktop ? "desktop" : "web",
    isDesktop: desktop,
    // デスクトップ版にだけある機能（Web 版との違い。about の比較表と合わせる）
    features: {
      nativeFs: desktop,      // フォルダ・ファイルをパスで直接読み書きする（毎回の許可なし。Safari・Firefox 相当の環境でも同じ）
      launchFiles: desktop,   // .skdoc などをダブルクリックして起動する（起動したウィンドウに渡るファイルは openArgs）
      multiWindow: desktop,   // 別のウィンドウで開く（openWindow）
      systemFonts: desktop,   // PC に入っているフォントすべて（systemFonts）
      silentPdf: desktop,     // 印刷のダイアログなしで PDF に書き出す（savePdf）
      externalLogin: desktop, // いつものブラウザでログインする（SK Hub Systems・Google。externalLogin）
      secureTokens: desktop,  // Google ドライブの長いログイン（リフレッシュトークン）を OS の保管庫に入れ、1 時間で切れない
      offlineQueue: desktop   // オフラインで保存した Google ドライブのファイルを、このアプリを閉じても覚えておき、つながったら送る
    },
    // 起動で渡されたファイルのパス（ダブルクリックしたとき。Rust がウィンドウを開くときに入れる）
    openArgs: desktop && Array.isArray(window.__SK_OPEN__) ? window.__SK_OPEN__ : [],
    invoke: function (cmd, args) {
      if (!desktop) return Promise.reject(new Error("not-desktop"));
      return T.core.invoke(cmd, args || {});
    },
    // 別のウィンドウで開く。app: "doc"・"sheet"・"slide"・"hub"。fileId があれば、そのファイルを開く
    openWindow: function (app, fileId) {
      var page = app === "doc" || app === "sheet" || app === "slide" ? app + "/index.html" : "index.html";
      return Platform.invoke("new_window", { page: page, hash: fileId ? "f=" + encodeURIComponent(fileId) : null });
    },
    // いつものブラウザでログインする。urlTemplate の {port}・{redirect} は、このアプリが待っているアドレスに置きかわる
    //   ログインが終わってブラウザが戻ってきたときの検索文字列（code・state・hint…と、_redirect）を返す。やめた・時間切れは reject
    //   inApp: true なら、いつものブラウザではなく、アプリの中の小さなログイン用のウィンドウで開く（SK Hub Systems のページ用。Google のログインは組み込みの画面を断るので、drive.js は使わない）
    externalLogin: function (urlTemplate, timeoutSecs, inApp) { return Platform.invoke("external_login", { urlTemplate: urlTemplate, timeoutSecs: timeoutSecs || 600, inApp: !!inApp }); },
    cancelLogin: function () { return Platform.invoke("external_login_cancel"); },
    openExternal: function (url) { return Platform.invoke("open_external", { url: url }); },
    // Google のトークン（リフレッシュトークンは Rust が OS の保管庫に入れ、ここには来ない）
    googleToken: function (params, useStoredRefresh) { return Platform.invoke("google_token", { params: params, useStoredRefresh: !!useStoredRefresh }); },
    googleHasRefresh: function () { return Platform.invoke("google_has_refresh"); },
    googleSignout: function () { return Platform.invoke("google_signout"); },
    // PC に入っているフォントの名前（五十音と ABC 順）
    systemFonts: function () { return Platform.invoke("system_fonts"); },
    // 印刷用の見た目にしたあとの画面を、PDF にして保存する。css は @page のある印刷用の CSS（用紙の大きさを読む）
    //   保存する場所を選んでもらう。やめたら false
    pickPdfPath: function (name) { return Platform.invoke("pick_pdf_path", { name: name }); },
    printPdf: function (path, css) {
      var s = Platform.pageSize(css);
      return Platform.invoke("print_pdf", { path: path, widthMm: s.w, heightMm: s.h });
    },
    // @page { size: … } から用紙の大きさ（mm）を読む。"A4 landscape"・"letter"・"254mm 142.9mm" など。読めないときは A4 縦
    pageSize: function (css) {
      var NAMED = { a3: [297, 420], a4: [210, 297], a5: [148, 210], b5: [182, 257], letter: [215.9, 279.4], legal: [215.9, 355.6] };
      var m = /@page\s*\{[^}]*?size\s*:\s*([^;}]+)/i.exec(css || "");
      var v = m ? m[1].trim().toLowerCase() : "", land = /landscape/.test(v), dims = null, mm = /([\d.]+)mm\s+([\d.]+)mm/.exec(v);
      if (mm) dims = [parseFloat(mm[1]), parseFloat(mm[2])];
      else for (var k in NAMED) if (new RegExp("\\b" + k + "\\b").test(v)) { dims = NAMED[k]; break; }
      dims = dims || NAMED.a4;
      return land && dims[0] < dims[1] ? { w: dims[1], h: dims[0] } : { w: dims[0], h: dims[1] };
    }
  };
  window.OfficePlatform = Platform;
  document.documentElement.setAttribute("data-platform", Platform.kind);

  // このアプリが Doc・Sheet・Slide のどれか（アプリごとに別々にインストールする。Rust がウィンドウを開くときに入れる）。開発用のすべて入りは "hub"
  Platform.app = desktop && typeof window.__SK_APP__ === "string" ? window.__SK_APP__ : "";

  // ---------- アプリのあいだで共有する設定 ----------
  //   アプリごとに画面の保存場所（localStorage）が分かれるので、ログイン・Google ドライブ・保存先のフォルダ・設定は、共有のファイル（Rust の shared_all・shared_set）に写す
  //   開くときに共有のファイルから localStorage に戻し、変えたら共有のファイルにも書く。ほかのアプリで変えたら、前に出したときに取りこむ
  var SHARED = ["skhub_account", "sk_office_drive", "sk_office_place", "sk_office_native_dir", "sk_office_settings"];
  if (desktop) {
    var rawSet = Storage.prototype.setItem, rawRemove = Storage.prototype.removeItem;
    var apply = function (all) {
      var changed = false;
      SHARED.forEach(function (k) {
        try {
          var now = localStorage.getItem(k), next = Object.prototype.hasOwnProperty.call(all, k) ? all[k] : null;
          if (now === next) return;
          if (next === null) rawRemove.call(localStorage, k); else rawSet.call(localStorage, k, next);
          changed = true;
        } catch (e) { /* 戻せなくても、このウィンドウでは使える */ }
      });
      return changed;
    };
    if (window.__SK_SHARED__) apply(window.__SK_SHARED__);
    Storage.prototype.setItem = function (k, v) {
      rawSet.call(this, k, v);
      if (this === window.localStorage && SHARED.indexOf(k) >= 0) Platform.invoke("shared_set", { key: k, value: String(v) }).catch(function () {});
    };
    Storage.prototype.removeItem = function (k) {
      rawRemove.call(this, k);
      if (this === window.localStorage && SHARED.indexOf(k) >= 0) Platform.invoke("shared_set", { key: k, value: null }).catch(function () {});
    };
    // ほかのアプリで変えたもの（ログイン・Google ドライブなど）を、前に出したときに取りこむ。変わったら sk-shared を知らせる
    var pull = function () {
      Platform.invoke("shared_all").then(function (all) { if (apply(all)) window.dispatchEvent(new Event("sk-shared")); }).catch(function () {});
    };
    window.addEventListener("focus", pull);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) pull(); });
  }

  // ---------- タイトルバーと「－ □ ×」の統合 ----------
  //   デスクトップ版は Windows の標準のタイトルバーを使わない（Rust 側で decorations を切っている）。
  //   かわりに、Office のタイトルバー（エディタの .titlebar・スタート画面の .wco-bar。Web 版の Window Controls Overlay と同じ）を帯にして、右の端に「－ □ ×」を置く
  //   帯の何もないところをドラッグするとウィンドウが動き、ダブルクリックで最大化・元に戻す
  //   office.js が <html class="is-wco"> を付け、office.css が帯の高さと右のあき（--wco-h・--wco-r）を決める
  if (desktop && T.window && T.window.getCurrentWindow) {
    var cur = T.window.getCurrentWindow();
    var SVG = {
      min: '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M0 5h10" stroke="currentColor" stroke-width="1" fill="none"/></svg>',
      max: '<svg viewBox="0 0 10 10" aria-hidden="true"><rect x=".5" y=".5" width="9" height="9" stroke="currentColor" stroke-width="1" fill="none"/></svg>',
      restore: '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M2.5 2.5V.5h7v7h-2M.5 2.5h7v7h-7z" stroke="currentColor" stroke-width="1" fill="none"/></svg>',
      close: '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M0 0l10 10M10 0L0 10" stroke="currentColor" stroke-width="1" fill="none"/></svg>'
    };
    var buildControls = function () {
      if (document.querySelector(".win-ctrl")) return;
      var box = document.createElement("div");
      box.className = "win-ctrl";
      box.innerHTML = '<button type="button" class="win-ctrl__b" data-win="min" aria-label="最小化" title="最小化">' + SVG.min + "</button>" +
        '<button type="button" class="win-ctrl__b" data-win="max" aria-label="最大化" title="最大化">' + SVG.max + "</button>" +
        '<button type="button" class="win-ctrl__b win-ctrl__b--close" data-win="close" aria-label="閉じる" title="閉じる">' + SVG.close + "</button>";
      document.body.appendChild(box);
      var maxBtn = box.querySelector('[data-win="max"]');
      var syncMax = function () {
        cur.isMaximized().then(function (m) {
          maxBtn.innerHTML = m ? SVG.restore : SVG.max;
          var label = m ? "元に戻す" : "最大化";
          maxBtn.setAttribute("aria-label", label); maxBtn.title = label;
          document.documentElement.classList.toggle("is-maximized", m);
        }).catch(function () {});
      };
      box.addEventListener("click", function (e) {
        var b = e.target.closest("[data-win]");
        if (!b) return;
        var w = b.getAttribute("data-win");
        if (w === "min") cur.minimize();
        else if (w === "max") cur.toggleMaximize().then(syncMax, syncMax);
        else cur.close();
      });
      window.addEventListener("resize", syncMax);
      syncMax();
    };
    // 帯の何もないところ（ボタン・リンク・入力欄の上ではないところ）をドラッグ → ウィンドウを動かす。ダブルクリック → 最大化・元に戻す
    var DRAG = ".titlebar, .wco-bar";
    var NO_DRAG = "button, a, input, select, textarea, label, [role=button], [role=menuitem], [contenteditable], .tbar-search, .tbar-name, .win-ctrl";
    var inBar = function (e) { return e.button === 0 && e.target.closest && e.target.closest(DRAG) && !e.target.closest(NO_DRAG); };
    document.addEventListener("mousedown", function (e) {
      if (!inBar(e) || e.detail > 1) return;
      cur.startDragging().catch(function () {});
    });
    document.addEventListener("dblclick", function (e) {
      if (!inBar(e)) return;
      cur.toggleMaximize().then(function () { window.dispatchEvent(new Event("resize")); }).catch(function () {});
    });
    if (document.body) buildControls(); else document.addEventListener("DOMContentLoaded", buildControls);
  }

  // ウィンドウのタイトルを、画面の題名（ファイル名 — Doc | SK Hub Office）に合わせる
  if (desktop && T.window && T.window.getCurrentWindow) {
    var win = T.window.getCurrentWindow(), last = "";
    var sync = function () { if (document.title && document.title !== last) { last = document.title; win.setTitle(last).catch(function () {}); } };
    new MutationObserver(sync).observe(document.querySelector("title") || document.documentElement, { childList: true, subtree: true, characterData: true });
    document.addEventListener("DOMContentLoaded", sync);
  }
})();
