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
    externalLogin: function (urlTemplate, timeoutSecs) { return Platform.invoke("external_login", { urlTemplate: urlTemplate, timeoutSecs: timeoutSecs || 600 }); },
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

  // ウィンドウのタイトルを、画面の題名（ファイル名 — Doc | SK Hub Office）に合わせる
  if (desktop && T.window && T.window.getCurrentWindow) {
    var win = T.window.getCurrentWindow(), last = "";
    var sync = function () { if (document.title && document.title !== last) { last = document.title; win.setTitle(last).catch(function () {}); } };
    new MutationObserver(sync).observe(document.querySelector("title") || document.documentElement, { childList: true, subtree: true, characterData: true });
    document.addEventListener("DOMContentLoaded", sync);
  }
})();
