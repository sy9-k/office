// SK Hub Office の共通（Doc・Sheet・Slide と入り口。開くのは https://office.sk-hub.jp/。中身は sk-hub.jp の /office/ を _workers/office/ が中継する）
//   画面の形は Microsoft Office（左に色付きのメニューがあるスタート画面・タイトルバー・タブとグループのリボン・ズーム付きのステータスバー）。
//   見た目は SK Material（部品は /toolbox/shared/m3.css、色は office.css）。
//   <html data-app="doc|sheet|slide|hub"> で、どのアプリかを決める。画面の中身はここで組み立てる（各ページは起動画面と入れ物だけ）
//   ・起動画面（#splash）→ スタート画面（ホーム・新規・開く）→ エディタ（#f=<ファイルの ID>）。エディタの「ファイル」でスタート画面に戻る
//   ・ファイルは PC の選んだフォルダに 1 つずつ保存する（フォルダに保存できないブラウザは、ブラウザの中）。中身は /office/shared/store.js
//   ・アカウント（SK Hub Systems アカウントのログインの状態・製品情報）とオプション（ユーザー名・テーマ・動き・保存先・自動保存）は、スタート画面の左のメニューから開く。
//     オプションは localStorage の sk_office_settings に保存する
//   ・PWA: /office/manifest.json（アプリ全体を 1 つの「SK Hub Office」として入れる。Doc・Sheet・Slide はショートカット）と /office/sw.js（オフライン）。
//     入れたあとは、.skdoc・.sksheet・.skslide をダブルクリックすると、そのアプリで開く（file_handlers・launchQueue）
//   ・まだできない操作のボタンは、押すと「まだ使えません」と出す（data-todo）
//   ・ダイアログは /toolbox/shared/m3.js の M3.prompt・M3.choose・M3.confirm
(function () {
  "use strict";
  var root = document.documentElement;
  var APP = root.getAttribute("data-app") || "hub";
  var I18N = window.SKI18N;
  function t(s, p) { return I18N ? I18N.t(s, p) : s.replace(/\{(\w+)\}/g, function (m, k) { return p && k in p ? p[k] : m; }); }
  var reduced = root.getAttribute("data-motion") === "reduce" || window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  var TEMPLATES = window.OFFICE_TEMPLATES || {};
  // Office のいちばん上のパス。office.sk-hub.jp では "/"、sk-hub.jp（手元のサーバー）では "/office/"（このファイルの 1 つ上）
  var BASE = (function () { var s = document.currentScript; try { return s ? new URL("..", s.src).pathname : "/office/"; } catch (e) { return "/office/"; } })();
  // サイト（sk-hub.jp）のページへのリンク。office.sk-hub.jp からは sk-hub.jp を付ける
  var SITE = /(^|\.)sk-hub\.jp$/.test(location.hostname) && !/^(www\.)?sk-hub\.jp$/.test(location.hostname) ? "https://sk-hub.jp" : "";

  var APPS = {
    doc: { name: "Doc", icon: "description", path: BASE + "doc/", untitled: t("無題の文書"), blank: t("白紙の文書") },
    sheet: { name: "Sheet", icon: "table_chart", path: BASE + "sheet/", untitled: t("無題の表"), blank: t("空白のブック") },
    slide: { name: "Slide", icon: "slideshow", path: BASE + "slide/", untitled: t("無題のプレゼンテーション"), blank: t("空白のプレゼンテーション") }
  };
  var CATS = [["all", t("すべて")], ["school", t("学校")], ["work", t("仕事")], ["life", t("くらし")]];

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s === undefined || s === null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function copy(o) { return JSON.parse(JSON.stringify(o)); }
  function icon(name, fill) { return '<span class="msr' + (fill ? " msr--fill" : "") + '" aria-hidden="true">' + name + "</span>"; }

  // ---------- 保存（/office/shared/store.js）と設定 ----------
  var Store = window.OfficeStore;
  var VERSION = "0.1";
  // SK Hub Office の設定（このブラウザの localStorage の sk_office_settings）
  //   { theme: "auto|light|dark", motion: "auto|reduce", autosave: true|false, user: "名前", pinned: [ファイルの ID] }
  var SET_KEY = "sk_office_settings";
  var settings = (function () {
    var d = {};
    try { d = JSON.parse(localStorage.getItem(SET_KEY) || "{}") || {}; } catch (e) { /* 初期値 */ }
    return {
      theme: ["auto", "light", "dark"].indexOf(d.theme) >= 0 ? d.theme : "auto",
      motion: d.motion === "reduce" ? "reduce" : "auto",
      autosave: d.autosave !== false,
      user: typeof d.user === "string" ? d.user : "",
      pinned: Array.isArray(d.pinned) ? d.pinned : []
    };
  })();
  function saveSettings() { try { localStorage.setItem(SET_KEY, JSON.stringify(settings)); } catch (e) { /* 保存できなくても、このページでは反映する */ } }
  // ライト・ダークは、「端末に合わせる」のときだけ、同じ場所にある SK's Toolbox の設定（sk_toolbox）にそろえる
  function applySettings() {
    var theme = settings.theme;
    if (theme === "auto") { try { var tb = JSON.parse(localStorage.getItem("sk_toolbox") || "{}"); if (tb.theme === "light" || tb.theme === "dark") theme = tb.theme; } catch (e) { /* なし */ } }
    if (theme === "light" || theme === "dark") root.setAttribute("data-theme", theme); else root.removeAttribute("data-theme");
    if (settings.motion === "reduce") root.setAttribute("data-motion", "reduce"); else root.removeAttribute("data-motion");
    reduced = settings.motion === "reduce" || window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    fillAvatars();
    syncThemeColor();
  }
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", function () { syncThemeColor(); });
  function accountHint() { try { return JSON.parse(localStorage.getItem("skhub_account") || "null"); } catch (e) { return null; } }
  function userName() { var h = accountHint(); return (settings.user || (h && h.name) || "").trim(); }
  // アカウントのアイコン（名前の 1 文字目。名前がなければ人のアイコン）
  function fillAvatars() {
    var n = userName();
    document.querySelectorAll("[data-avatar]").forEach(function (el) { el.innerHTML = n ? esc(Array.from(n)[0].toUpperCase()) : icon("person"); });
  }
  function isPinned(id) { return settings.pinned.indexOf(id) >= 0; }
  function setPinned(id, on) {
    settings.pinned = settings.pinned.filter(function (x) { return x !== id; });
    if (on) settings.pinned.push(id);
    saveSettings();
  }
  function where() { return Store.state() === "ready" ? Store.folderName() : Store.supported ? t("フォルダ") : t("このブラウザ"); }
  function when(ts) {
    var d = new Date(ts), now = new Date();
    var hm = d.getHours() + ":" + ("0" + d.getMinutes()).slice(-2);
    if (d.toDateString() === now.toDateString()) return t("今日 {time}", { time: hm });
    var y = new Date(now); y.setDate(now.getDate() - 1);
    if (d.toDateString() === y.toDateString()) return t("昨日 {time}", { time: hm });
    if (d.getFullYear() === now.getFullYear()) return t("{m}月{d}日", { m: d.getMonth() + 1, d: d.getDate() });
    return t("{y}年{m}月{d}日", { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() });
  }
  // ---------- パソコンでアプリとして入れたときのタイトルバー（Window Controls Overlay） ----------
  //   manifest の display_override で、ウィンドウのいちばん上の帯をページが使う。使っているあいだは <html class="is-wco">。
  //   エディタは、タイトルバー（アプリのアイコン・保存・ファイル名・検索・共有・アカウント）をそのまま帯にする。
  //   スタート画面は、同じ色の帯（.wco-bar）を出す。見た目は office.css。タイトルバーの右の「∧」で、ふつうのタイトルバーに戻せる
  var wco = navigator.windowControlsOverlay;
  function syncWco() { root.classList.toggle("is-wco", !!(wco && wco.visible)); }
  if (wco) { syncWco(); wco.addEventListener("geometrychange", syncWco); }
  // ウィンドウのボタン（閉じるなど）とアドレスバーの色を、アプリの色（タイトルバーの色）にそろえる
  function syncThemeColor() {
    var meta = document.querySelector('meta[name="theme-color"]');
    var c = getComputedStyle(root).getPropertyValue("--md-primary-container").trim();
    if (meta && c) meta.setAttribute("content", c);
  }

  // ---------- アプリとして入れる（PWA。/office/manifest.json・/office/sw.js） ----------
  var installEvt = null;
  var standalone = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
  window.addEventListener("beforeinstallprompt", function (e) {
    e.preventDefault();
    installEvt = e;
    if ($("home") && !$("home").hidden) renderStart();
  });
  window.addEventListener("appinstalled", function () {
    installEvt = null;
    notify(t("SK Hub Office をインストールしました"));
    if ($("home") && !$("home").hidden) renderStart();
  });
  function installText() {
    if (standalone) return t("インストールして使っています。");
    if (installEvt) return t("Windows・Mac・Chromebook に、アプリとして入れます。オフラインでも開けて、.skdoc などのファイルをダブルクリックで開けるようになります。");
    return t("ブラウザのメニューから入れられます（Chrome・Edge はアドレスバーの右のインストールのボタン、Safari は共有ボタンの「Dock に追加」や「ホーム画面に追加」）。");
  }
  function failed(e) {
    if (window.console) console.error(e);
    notify(Store.supported ? t("保存できませんでした。フォルダを使う許可を確かめてください") : t("保存できませんでした。端末の空きを確かめてください"));
  }
  // フォルダを使えるようにする（ボタンを押したときに呼ぶ。フォルダを選ぶ・許可をもらう）。使えるなら true
  function ensureFolder() {
    var s = Store.state();
    if (s === "unsupported" || s === "ready") return Promise.resolve(true);
    return (s === "none" ? Store.pick() : Store.allow()).then(function (ok) { if (ok) afterFolder(true); return ok; });
  }

  // ---------- スナックバー ----------
  var snackTimer;
  function notify(msg) {
    var snack = $("toast");
    if (!snack) return;
    $("toast-text").textContent = msg;
    snack.classList.add("is-on");
    clearTimeout(snackTimer);
    snackTimer = setTimeout(function () { snack.classList.remove("is-on"); }, 3200);
  }

  // ---------- 小さなプレビュー（テンプレート・最近のファイル） ----------
  function thumb(app, data) {
    if (app === "doc") return '<div class="th th--doc" aria-hidden="true"><div class="th-doc">' + (data.html || "") + "</div></div>";
    if (app === "sheet") return '<div class="th th--sheet" aria-hidden="true">' + sheetTable(data, 7, 5, false) + "</div>";
    var s = (data.slides || [])[0] || { layout: "title" };
    return '<div class="th th--slide" aria-hidden="true">' + slideHtml(s, false) + "</div>";
  }

  // ================================================================
  //   スタート画面（Office の「ホーム・新規・開く・アカウント・オプション」）
  // ================================================================
  var view = "home", fileTab = "recent", query = "", current = null;

  function railItem(v, ic, label) {
    return '<button type="button" class="rail__item" data-view="' + v + '"' + (view === v ? ' aria-current="page"' : "") + ">" +
      '<span class="rail__pill">' + icon(ic, view === v) + "</span><span>" + esc(label) + "</span></button>";
  }
  function railBtn(attr, ic, label) {
    return '<button type="button" class="rail__item" ' + attr + '><span class="rail__pill">' + icon(ic) + "</span><span>" + esc(label) + "</span></button>";
  }
  var animateFiles = false;
  function renderStart() {
    var a = APPS[APP];
    // 「ファイル」で開いたときの滑りこみは 1 回だけ（ほかの画面に移ったら、ふつうの出方にもどす）
    if ($("home")) $("home").classList.remove("is-backstage");
    animateFiles = true;
    var rail = '<nav class="rail" aria-label="' + esc(t("メニュー")) + '">' +
      (current ? railBtn("data-back-edit", "arrow_back", t("戻る")) : "") +
      '<a class="rail__brand" href="' + BASE + '" title="SK Hub Office"><span class="rail__logo">' + icon(a ? a.icon : "apps", true) + "</span><span>" + (a ? a.name : "Office") + "</span></a>" +
      railItem("home", "home", t("ホーム")) +
      (a ? railItem("new", "note_add", t("新規")) + railItem("open", "folder_open", t("開く")) : "") +
      (current ? railBtn("data-save-as", "save_as", t("コピーを保存")) + railBtn("data-close-file", "close", t("閉じる")) : "") +
      '<span class="rail__space"></span>' +
      (a ? '<a class="rail__item" href="' + BASE + '"><span class="rail__pill">' + icon("apps") + "</span><span>" + esc(t("アプリ")) + "</span></a>"
         : '<a class="rail__item" href="' + SITE + '/products/office/"><span class="rail__pill">' + icon("info") + "</span><span>" + esc(t("紹介")) + "</span></a>") +
      (installEvt && !standalone ? railBtn("data-install", "install_desktop", t("インストール")) : "") +
      railItem("account", "account_circle", t("アカウント")) +
      railItem("options", "settings", t("オプション")) +
      "</nav>";
    var main = view === "new" ? viewNew() : view === "open" ? viewOpen() : view === "account" ? viewAccount() : view === "options" ? viewOptions() : viewHome();
    // アプリとして入れたときだけ見える、ウィンドウのいちばん上の帯（is-wco）
    var bar = '<header class="wco-bar"><span class="wco-bar__logo">' + icon(a ? a.icon : "apps", true) + '</span><span class="wco-bar__name">' + (a ? a.name : "SK Hub Office") + "</span>" +
      (a ? '<span class="wco-bar__sub">SK Hub Office</span>' : "") +
      '<span class="wco-bar__space"></span><button type="button" class="wco-search m3-state" data-todo="' + esc(t("検索")) + '">' + icon("search") + "<span>" + esc(t("検索")) + "</span></button>" +
      '<span class="wco-bar__space"></span><button type="button" class="tbar-avatar" data-view="account" data-avatar aria-label="' + esc(t("アカウント")) + '" title="' + esc(t("アカウント")) + '"></button></header>';
    $("home").innerHTML = bar + rail + '<main class="start__main" id="start-main">' + main + "</main>";
    if (view === "new") filterTemplates("all");
    if (view === "home" || view === "open") renderFileTable();
    renderLegacy();
    fillAvatars();
  }
  // フォルダを選んでいない・許可がいるときのお知らせ（showUnsupported: フォルダに保存できないブラウザのお知らせも出す）
  function banner(showUnsupported) {
    var s = Store.state();
    if (s === "ready") return "";
    if (s === "unsupported") return showUnsupported ? '<div class="note-card">' + icon("info") + "<p>" + esc(t("このブラウザでは PC のフォルダに保存できないので、ブラウザの中に保存します。パソコンの Chrome か Edge なら、フォルダにファイルとして保存できます。")) + "</p></div>" : "";
    if (s === "needs-permission") return '<div class="folder-card">' + icon("folder_open") + '<div class="folder-card__text"><strong>' + esc(t("フォルダ「{name}」を使う許可が必要です", { name: Store.folderName() })) + "</strong><p>" + esc(t("ブラウザを開き直すと、もう一度許可を聞かれることがあります。")) + '</p></div><button type="button" class="m3-btn m3-state" data-allow>' + esc(t("許可する")) + "</button></div>";
    return '<div class="folder-card">' + icon("create_new_folder") + '<div class="folder-card__text"><strong>' + esc(t("保存するフォルダを選んでください")) + "</strong><p>" + esc(t("ファイルは、この PC のフォルダに 1 つずつ保存されます。最初に 1 回だけ、保存するフォルダを選んでください（例：ドキュメントの中に「SK Hub Office」フォルダを作って選ぶ）。")) + '</p></div><button type="button" class="m3-btn m3-state" data-pick>' + esc(t("フォルダを選ぶ")) + "</button></div>";
  }
  function greet() {
    var h = new Date().getHours(), n = userName();
    var g = h < 5 ? t("こんばんは") : h < 11 ? t("おはようございます") : h < 18 ? t("こんにちは") : t("こんばんは");
    return n ? t("{greet}、{name} さん", { greet: g, name: n }) : g;
  }
  function viewHome() {
    var a = APPS[APP], h = '<h1 class="start__h1">' + esc(greet()) + "</h1>" + banner(false);
    if (a) {
      var tpls = (TEMPLATES[APP] || []).slice(0, 6);
      h += '<section class="start__sec"><div class="tpl-row">' + tpls.map(function (tp) { return tplCard(APP, tp); }).join("") + "</div>" +
        '<button type="button" class="more-link m3-state" data-view="new">' + esc(t("その他のテンプレート")) + icon("arrow_forward") + "</button></section>";
    } else {
      h += '<section class="start__sec"><h2 class="start__h2">' + esc(t("新しく作る")) + '</h2><div class="apps">' +
        ["doc", "sheet", "slide"].map(function (k) {
          var x = APPS[k];
          return '<a class="app-tile app-tile--' + k + ' m3-state" href="' + x.path + '"><span class="app-tile__ico">' + icon(x.icon, true) + "</span><span><strong>" + x.name + "</strong><small>" + esc(x.blank) + "</small></span></a>";
        }).join("") + "</div></section>";
    }
    return h + fileSection(true);
  }
  function viewNew() {
    return '<h1 class="start__h1">' + esc(t("新規")) + "</h1>" + banner(false) + '<section class="start__sec">' +
      '<div class="chips" id="cat-chips" role="radiogroup" aria-label="' + esc(t("種類")) + '">' + CATS.map(function (c, i) {
        return '<button type="button" class="chip m3-state" role="radio" aria-checked="' + (i === 0) + '" data-cat="' + c[0] + '">' + icon("check") + esc(c[1]) + "</button>";
      }).join("") + '</div><div class="tpl-grid" id="tpl-grid">' + (TEMPLATES[APP] || []).map(function (tp) { return tplCard(APP, tp); }).join("") + "</div></section>";
  }
  function viewOpen() {
    var s = Store.state();
    var here = s === "unsupported" ? ["public", t("このブラウザ"), t("ブラウザの中に保存したファイル")]
      : s === "ready" ? ["computer", t("この PC"), t("フォルダ「{name}」", { name: Store.folderName() })]
      : ["computer", t("この PC"), t("フォルダを選んでいません")];
    return '<h1 class="start__h1">' + esc(t("開く")) + "</h1>" + banner(true) +
      '<div class="open">' +
        '<div class="open__places" role="list">' +
          '<button type="button" class="place is-on m3-state" role="listitem">' + icon(here[0]) + "<span><strong>" + esc(here[1]) + "</strong><small>" + esc(here[2]) + "</small></span></button>" +
          (Store.supported ? '<button type="button" class="place m3-state" role="listitem" data-pick>' + icon("drive_folder_upload") + "<span><strong>" + esc(t("フォルダを変更")) + "</strong><small>" + esc(t("ほかのフォルダに保存する")) + "</small></span></button>" : "") +
          '<button type="button" class="place m3-state" role="listitem" data-import>' + icon("upload_file") + "<span><strong>" + esc(t("参照")) + "</strong><small>" + esc(t(".skdoc・.sksheet・.skslide を開く")) + "</small></span></button>" +
          '<button type="button" class="place m3-state" role="listitem" data-todo="' + esc(t("SK Hub Systems アカウント")) + '">' + icon("cloud") + "<span><strong>" + esc(t("SK Hub Systems アカウント")) + "</strong><small>" + esc(t("もうすぐ使えます")) + "</small></span></button>" +
        '</div><div class="open__main"><div id="legacy"></div>' + fileSection(false) + "</div>" +
      "</div>";
  }
  function viewAccount() {
    var hint = accountHint(), name = userName(), a = APPS[APP] || { icon: "apps" };
    return '<h1 class="start__h1">' + esc(t("アカウント")) + "</h1>" +
      '<section class="acct-card"><span class="acct-avatar" data-avatar aria-hidden="true"></span>' +
        '<div class="acct-card__text"><strong>' + esc(name || t("名前が設定されていません")) + "</strong><small>" + esc(hint && hint.email ? hint.email : t("SK Hub Systems アカウントにログインしていません")) + "</small></div>" +
        '<button type="button" class="text-btn m3-state" data-view="options">' + esc(t("名前を変更")) + "</button></section>" +
      '<section class="start__sec"><h2 class="start__h2">SK Hub Systems ' + esc(t("アカウント")) + '</h2><p class="sec-text">' +
        esc(hint ? t("sk-hub.jp に {name} さんとしてログインしています。SK Hub Office とつなぐ機能は、もうすぐ使えます。", { name: hint.name || hint.email }) : t("ログインすると、いろいろな端末で同じファイルを開けるようになる予定です。いまは、ファイルはこの PC に保存されるので、ログインしなくても使えます。")) +
        '</p><div class="btn-row"><a class="m3-btn m3-btn--tonal m3-state" href="' + SITE + '/account/" target="_blank" rel="noopener">' + icon("open_in_new") + esc(t("アカウントのページを開く")) + "</a></div></section>" +
      '<section class="start__sec"><h2 class="start__h2">' + esc(t("製品情報")) + '</h2>' +
        '<div class="about"><span class="about__logo">' + icon(a.icon, true) + '</span><div><strong>SK Hub Office</strong><small>' + esc(t("バージョン {v}（開発中）", { v: VERSION })) + "</small></div></div>" +
        '<p class="sec-text">' + esc(t("文書の Doc・表計算の Sheet・スライドの Slide をまとめた、SK のオフィスアプリです。見た目は SK Material（Google の Material Design 3 がもと）です。")) + "</p>" +
        '<div class="links"><a href="' + SITE + '/products/office/" target="_blank" rel="noopener">' + esc(t("SK Hub Office について")) + '</a><a href="' + SITE + '/policies/" target="_blank" rel="noopener">' + esc(t("利用規約とプライバシーポリシー")) + '</a><a href="https://support.sk-hub.jp/" target="_blank" rel="noopener">' + esc(t("サポート")) + "</a></div></section>";
  }
  function optRow(ic, title, sub, control) {
    return '<div class="opt">' + icon(ic) + '<div class="opt__text"><strong>' + esc(title) + "</strong>" + (sub ? "<small>" + esc(sub) + "</small>" : "") + "</div>" + control + "</div>";
  }
  function viewOptions() {
    var s = Store.state();
    var place = s === "unsupported" ? t("このブラウザの中（このブラウザでは PC のフォルダに保存できません）")
      : s === "none" ? t("まだ選んでいません") : Store.folderName() + (s === "needs-permission" ? t("（許可が必要です）") : "");
    var themes = [["auto", "brightness_auto", t("端末に合わせる")], ["light", "light_mode", t("ライト")], ["dark", "dark_mode", t("ダーク")]];
    return '<h1 class="start__h1">' + esc(t("オプション")) + "</h1>" +
      '<section class="opts"><h2 class="opts__title">' + esc(t("全般")) + "</h2>" +
        optRow("badge", t("ユーザー名"), t("あいさつとアカウントのアイコンに出ます"), '<input class="opt-input" id="opt-user" maxlength="40" autocomplete="nickname" placeholder="' + esc(t("例：SK")) + '" value="' + esc(settings.user) + '" aria-label="' + esc(t("ユーザー名")) + '">') +
      "</section>" +
      '<section class="opts"><h2 class="opts__title">' + esc(t("表示")) + "</h2>" +
        optRow("contrast", t("テーマ"), "", '<div class="tbs-seg" role="radiogroup" aria-label="' + esc(t("テーマ")) + '">' + themes.map(function (x) {
          return '<button type="button" class="tbs-seg__btn m3-state" role="radio" aria-checked="' + (settings.theme === x[0]) + '" data-theme-set="' + x[0] + '"><span class="msr tbs-seg__check" aria-hidden="true">check</span><span class="msr tbs-seg__icon" aria-hidden="true">' + x[1] + "</span>" + esc(x[2]) + "</button>";
        }).join("") + "</div>") +
        optRow("animation", t("動きを減らす"), t("起動画面や、画面の切りかえの動きを止めます"), '<input type="checkbox" class="tbs-switch" id="opt-motion"' + (settings.motion === "reduce" ? " checked" : "") + ' aria-label="' + esc(t("動きを減らす")) + '">') +
      "</section>" +
      '<section class="opts"><h2 class="opts__title">' + esc(t("保存")) + "</h2>" +
        optRow("folder", t("保存先"), place, s === "unsupported" ? "" : '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-pick>' + esc(s === "none" ? t("フォルダを選ぶ") : t("変更")) + "</button>") +
        optRow("save", t("自動保存"), t("変えたらすぐに保存します。オフにすると、保存のボタンか Ctrl+S で保存します"), '<input type="checkbox" class="tbs-switch" id="opt-autosave"' + (settings.autosave ? " checked" : "") + ' aria-label="' + esc(t("自動保存")) + '">') +
        '<div id="legacy"></div>' +
      "</section>" +
      '<section class="opts"><h2 class="opts__title">' + esc(t("アプリ")) + "</h2>" +
        optRow("install_desktop", t("アプリとしてインストール"), installText(), installEvt && !standalone ? '<button type="button" class="m3-btn m3-state" data-install>' + esc(t("インストール")) + "</button>" : "") +
      "</section>" +
      '<p class="start__note">' + esc(t("オプションは、このブラウザに保存されます。")) + "</p>";
  }
  // 前にブラウザの中に保存したファイル（フォルダに移せる）
  function renderLegacy() {
    var box = $("legacy");
    if (!box) return;
    var n = Store.legacy().length;
    if (!n) { box.innerHTML = ""; return; }
    box.innerHTML = '<div class="note-card">' + icon("move_to_inbox") + "<p>" + esc(Store.state() === "ready"
      ? t("前にこのブラウザの中に保存したファイルが {n} 件あります。フォルダ「{name}」に移せます。", { n: n, name: Store.folderName() })
      : t("前にこのブラウザの中に保存したファイルが {n} 件あります。フォルダを選ぶと、そこに移せます。", { n: n })) + "</p>" +
      (Store.state() === "ready" ? '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-migrate>' + esc(t("フォルダに移す")) + "</button>" : "") + "</div>";
  }
  function fileSection(withTabs) {
    return '<section class="start__sec files-sec">' +
      '<div class="files-head">' +
        (withTabs ? '<div class="tabs" role="tablist">' +
          '<button type="button" class="tab m3-state" role="tab" data-ftab="recent" aria-selected="' + (fileTab === "recent") + '"><span>' + esc(t("最近使ったアイテム")) + "</span></button>" +
          '<button type="button" class="tab m3-state" role="tab" data-ftab="pinned" aria-selected="' + (fileTab === "pinned") + '"><span>' + esc(t("ピン留め")) + "</span></button></div>"
          : '<h2 class="start__h2">' + esc(Store.state() === "unsupported" ? t("このブラウザのファイル") : t("フォルダのファイル")) + "</h2>") +
        '<label class="search">' + icon("search") + '<input type="search" id="file-search" placeholder="' + esc(t("検索")) + '" value="' + esc(query) + '" autocomplete="off" aria-label="' + esc(t("ファイルを検索")) + '"></label>' +
      "</div>" +
      '<div class="ftable" role="table" aria-label="' + esc(t("ファイル")) + '">' +
        '<div class="ftable__head" role="row"><span role="columnheader">' + esc(t("名前")) + '</span><span role="columnheader" class="ftable__when">' + esc(t("更新日時")) + '</span><span role="columnheader"></span></div>' +
        '<div id="files" data-i18n-skip><div class="files-loading">' + esc(t("読み込んでいます…")) + "</div></div>" +
      "</div>" +
      '<p class="files-empty" id="files-empty" hidden></p>' +
      '<p class="start__note">' + esc(Store.state() === "unsupported" ? t("ファイルは、このブラウザの中に保存されます。SK Hub Office は開発中です。") : t("ファイルは、この PC の選んだフォルダに保存されます。SK Hub Office は開発中です。")) + "</p></section>";
  }
  function tplCard(app, tp) {
    var blank = tp.cat === "blank";
    return '<button type="button" class="tpl m3-state" data-tpl="' + esc(tp.id) + '" data-cat="' + esc(tp.cat) + '">' +
      thumb(app, blank ? { html: "", cells: {}, slides: [{ layout: "blank" }] } : tp.data) +
      '<span class="tpl__name">' + esc(blank ? APPS[app].blank : t(tp.name)) + "</span></button>";
  }
  function filterTemplates(cat) {
    var chips = $("cat-chips");
    if (!chips) return;
    chips.querySelectorAll(".chip").forEach(function (c) { c.setAttribute("aria-checked", String(c.getAttribute("data-cat") === cat)); });
    $("tpl-grid").querySelectorAll(".tpl").forEach(function (x) {
      var c = x.getAttribute("data-cat");
      x.hidden = !(cat === "all" || c === cat || c === "blank");
    });
  }
  var listToken = 0;
  function renderFileTable() {
    if (!$("files")) return Promise.resolve();
    var my = ++listToken;
    return Store.list().then(function (all) {
      var box = $("files");
      if (my !== listToken || !box) return;
      all = all.filter(function (f) { return APP === "hub" ? !!APPS[f.app] : f.app === APP; });
      var list = all.filter(function (f) {
        if (view === "home" && fileTab === "pinned" && !isPinned(f.id)) return false;
        return !query || f.name.toLowerCase().indexOf(query.toLowerCase()) >= 0;
      }).sort(function (a, b) { return (isPinned(b.id) ? 1 : 0) - (isPinned(a.id) ? 1 : 0) || b.updated - a.updated; });
      var empty = $("files-empty"), s = Store.state();
      empty.hidden = list.length > 0;
      empty.textContent = s === "none" || s === "needs-permission" ? t("フォルダを使えるようにすると、ここにファイルが出ます。")
        : query ? t("「{q}」に合うファイルはありません", { q: query })
        : view === "home" && fileTab === "pinned" ? t("ピン留めしたファイルはありません。ファイルの右のピンを押すと、ここに出ます。")
        : APPS[APP] ? t("まだファイルがありません。上のテンプレートを選ぶと、新しく作れます。") : t("まだファイルがありません。上のアプリから作れます。");
      // 画面を開いたときだけ、1 つずつずらして出す（検索の文字を打つたびには動かさない）
      box.classList.toggle("is-anim", animateFiles);
      animateFiles = false;
      box.innerHTML = list.map(function (f) {
        var a = APPS[f.app], pin = isPinned(f.id);
        return '<div class="frow" role="row" data-app="' + f.app + '">' +
          '<button type="button" class="frow__open m3-state" role="cell" data-open="' + esc(f.id) + '" data-open-app="' + f.app + '">' +
            '<span class="frow__ico">' + icon(a.icon, true) + '</span><span class="frow__text"><span class="frow__name">' + esc(f.name) + '</span><span class="frow__meta">' + esc(a.name) + " ・ " + esc(where()) + '<span class="frow__when-sm"> ・ ' + esc(when(f.updated)) + "</span></span></span></button>" +
          '<span class="ftable__when" role="cell">' + esc(when(f.updated)) + "</span>" +
          '<span class="frow__act" role="cell">' +
            '<button type="button" class="icon-btn m3-state" data-pin="' + esc(f.id) + '" aria-pressed="' + pin + '" aria-label="' + esc(pin ? t("ピン留めを外す") : t("ピン留め")) + '" title="' + esc(pin ? t("ピン留めを外す") : t("ピン留め")) + '">' + icon("push_pin") + "</button>" +
            '<button type="button" class="icon-btn m3-state" data-more="' + esc(f.id) + '" data-name="' + esc(f.name) + '" aria-label="' + esc(t("ファイルの操作")) + '">' + icon("more_horiz") + "</button></span></div>";
      }).join("");
    }).catch(function (e) {
      if (window.console) console.error(e);
      var box = $("files");
      if (box) box.innerHTML = '<div class="files-loading">' + esc(t("フォルダを読めませんでした。オプションから、フォルダを選び直してください。")) + "</div>";
    });
  }

  function openFile(id, app) {
    if (app && app !== APP && APPS[app]) { location.href = APPS[app].path + "#f=" + encodeURIComponent(id); return; }
    history.pushState(null, "", "#f=" + encodeURIComponent(id));
    route();
  }
  function createFrom(tplId) {
    var tp = (TEMPLATES[APP] || []).filter(function (x) { return x.id === tplId; })[0];
    if (!tp) return;
    ensureFolder().then(function (ok) {
      if (!ok) return;
      return Store.create(APP, tp.cat === "blank" ? APPS[APP].untitled : t(tp.name), tp.id, copy(tp.data)).then(function (f) { openFile(f.id); });
    }).catch(failed);
  }
  function fileActions(id, name) {
    if (!window.M3) return;
    M3.choose({ title: name, options: [
      { value: "open", label: t("開く"), icon: "open_in_new" },
      { value: "rename", label: t("名前を変更"), icon: "edit" },
      { value: "copy", label: t("コピーを作る"), icon: "content_copy" },
      { value: "export", label: t("ファイルとして書き出す"), icon: "download" },
      { value: "delete", label: t("削除"), icon: "delete" }
    ] }).then(function (v) {
      if (!v) return;
      if (v === "open") return Store.get(id).then(function (f) { if (f) openFile(f.id, f.app); });
      if (v === "rename") return Store.get(id).then(function (f) { if (f) renameFile(f, renderFileTable); });
      if (v === "copy") return Store.get(id).then(function (f) {
        return f && Store.copy(f, t("{name} のコピー", { name: f.name })).then(function () { renderFileTable(); notify(t("コピーを作りました")); });
      });
      if (v === "export") return Store.get(id).then(function (f) { return f && Store.exportFile(f); }).then(function (ok) { if (ok) notify(t("書き出しました")); });
      if (v === "delete") {
        return M3.confirm({ title: t("「{name}」を削除しますか？", { name: name }),
          text: Store.state() === "unsupported" ? t("このブラウザから消えます。元に戻すことはできません。") : t("フォルダ「{folder}」からファイルを削除します。ごみ箱には入らず、元に戻すことはできません。", { folder: Store.folderName() }),
          ok: t("削除"), danger: true, icon: "delete" })
          .then(function (ok) { if (!ok) return; return Store.remove(id).then(function () { setPinned(id, false); renderFileTable(); notify(t("削除しました")); }); });
      }
    }).catch(failed);
  }
  function renameFile(f, done) {
    return M3.prompt({ title: t("名前を変更"), label: t("ファイルの名前"), value: f.name, ok: t("変更") }).then(function (v) {
      v = (v || "").trim();
      if (!v || v === f.name) return;
      var old = f.id;
      return Store.rename(f, v).then(function () {
        if (old !== f.id && isPinned(old)) { setPinned(old, false); setPinned(f.id, true); }
        if (done) done();
      });
    }).catch(function (e) {
      if (e && e.message === "outside") notify(t("フォルダの外のファイルは、名前を変えられません。「コピーを保存」で、新しい名前で保存できます"));
      else failed(e);
    });
  }
  function importFile() {
    ensureFolder().then(function (ok) {
      if (!ok) return;
      var pick = typeof window.showOpenFilePicker === "function"
        ? window.showOpenFilePicker({ types: [{ description: "SK Hub Office", accept: { "application/json": [".skdoc", ".sksheet", ".skslide"] } }] }).then(function (hs) { return hs[0].getFile(); }, function () { return null; })
        : new Promise(function (resolve) {
            var input = document.createElement("input");
            input.type = "file"; input.accept = ".skdoc,.sksheet,.skslide";
            input.addEventListener("change", function () { resolve(input.files[0] || null); });
            input.click();
          });
      return pick.then(function (file) {
        if (!file) return;
        return Store.importFile(file).then(function (f) { notify(t("「{name}」を読み込みました", { name: f.name })); openFile(f.id, f.app); },
          function () { notify(t("SK Hub Office のファイルではないので、開けませんでした")); });
      });
    }).catch(failed);
  }

  // ================================================================
  //   エディタ（タイトルバー・リボン・ステータスバー）
  // ================================================================
  // リボン: [{ tab, groups: [{ label, items }] }]
  //   items の 1 つ: { big, icon, label } … 大きなボタン / { stack: [{ icon, label }] } … 文字つきの小さなボタンを縦に
  //                  { rows: [[{ icon, label } か { chip, w }]] } … アイコンだけのボタンの行 / { styles: [{ tag, label }] } … スタイルの一覧（Doc）
  //   cmd があれば動く（COMMANDS）。todo: 1 は「まだ使えません」。toggle: "bold" などは、選んだところの状態でオンに見せる（Doc）
  var CLIP = { label: t("クリップボード"), items: [
    { big: 1, icon: "content_paste", label: t("貼り付け"), todo: 1 },
    { stack: [{ icon: "content_cut", label: t("切り取り"), todo: 1 }, { icon: "content_copy", label: t("コピー"), todo: 1 }, { icon: "format_paint", label: t("書式のコピー"), todo: 1 }] }
  ] };
  var ZOOM = { label: t("ズーム"), items: [
    { big: 1, icon: "search", label: t("ズーム"), cmd: "zoom" },
    { stack: [{ icon: "zoom_in", label: t("拡大"), cmd: "zoomIn" }, { icon: "zoom_out", label: t("縮小"), cmd: "zoomOut" }, { icon: "fit_screen", label: "100%", cmd: "zoom100" }] }
  ] };
  var RIBBON = {
    doc: [
      { tab: t("ホーム"), groups: [CLIP,
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "M PLUS 1", w: 132, todo: 1 }, { chip: "10.5", w: 64, todo: 1 }],
          [{ icon: "format_bold", label: t("太字"), cmd: "bold", toggle: "bold" }, { icon: "format_italic", label: t("斜体"), cmd: "italic", toggle: "italic" },
           { icon: "format_underlined", label: t("下線"), cmd: "underline", toggle: "underline" }, { icon: "format_strikethrough", label: t("取り消し線"), cmd: "strikeThrough", toggle: "strikeThrough" },
           { icon: "format_color_text", label: t("文字の色"), todo: 1 }, { icon: "ink_highlighter", label: t("蛍光ペン"), todo: 1 }]
        ] }] },
        { label: t("段落"), items: [{ rows: [
          [{ icon: "format_list_bulleted", label: t("箇条書き"), cmd: "insertUnorderedList", toggle: "insertUnorderedList" }, { icon: "format_list_numbered", label: t("段落番号"), cmd: "insertOrderedList", toggle: "insertOrderedList" },
           { icon: "format_indent_decrease", label: t("インデントを減らす"), cmd: "outdent" }, { icon: "format_indent_increase", label: t("インデントを増やす"), cmd: "indent" }],
          [{ icon: "format_align_left", label: t("左揃え"), cmd: "justifyLeft", toggle: "justifyLeft" }, { icon: "format_align_center", label: t("中央揃え"), cmd: "justifyCenter", toggle: "justifyCenter" },
           { icon: "format_align_right", label: t("右揃え"), cmd: "justifyRight", toggle: "justifyRight" }, { icon: "format_align_justify", label: t("両端揃え"), cmd: "justifyFull", toggle: "justifyFull" }]
        ] }] },
        { label: t("スタイル"), items: [{ styles: [{ tag: "p", label: t("標準") }, { tag: "h1", label: t("表題") }, { tag: "h2", label: t("見出し 1") }, { tag: "h3", label: t("見出し 2") }, { tag: "blockquote", label: t("引用") }] }] },
        { label: t("編集"), items: [{ stack: [{ icon: "search", label: t("検索"), todo: 1 }, { icon: "find_replace", label: t("置換"), todo: 1 }, { icon: "select_all", label: t("すべて選択"), cmd: "selectAll" }] }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("ページ"), items: [{ big: 1, icon: "insert_page_break", label: t("改ページ"), todo: 1 }] },
        { label: t("表"), items: [{ big: 1, icon: "table", label: t("表"), cmd: "table" }] },
        { label: t("図"), items: [{ big: 1, icon: "image", label: t("画像"), todo: 1 }, { stack: [{ icon: "category", label: t("図形"), todo: 1 }, { icon: "bar_chart", label: t("グラフ"), todo: 1 }] }] },
        { label: t("リンク"), items: [{ big: 1, icon: "link", label: t("リンク"), cmd: "link" }] },
        { label: t("テキスト"), items: [{ stack: [{ icon: "horizontal_rule", label: t("横線"), cmd: "hr" }, { icon: "today", label: t("日付"), cmd: "date" }] }] }
      ] },
      { tab: t("レイアウト"), groups: [
        { label: t("ページ設定"), items: [{ big: 1, icon: "margin", label: t("余白"), todo: 1 }, { big: 1, icon: "crop_rotate", label: t("印刷の向き"), todo: 1 }, { big: 1, icon: "crop_portrait", label: t("サイズ"), todo: 1 }] }
      ] },
      { tab: t("校閲"), groups: [
        { label: t("文章校正"), items: [{ big: 1, icon: "spellcheck", label: t("スペル チェック"), todo: 1 }, { big: 1, icon: "123", label: t("文字カウント"), cmd: "count" }] },
        { label: t("コメント"), items: [{ big: 1, icon: "add_comment", label: t("新しいコメント"), todo: 1 }] }
      ] },
      { tab: t("表示"), groups: [ZOOM] }
    ],
    sheet: [
      { tab: t("ホーム"), groups: [CLIP,
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "M PLUS 1", w: 132, todo: 1 }, { chip: "11", w: 56, todo: 1 }],
          [{ icon: "format_bold", label: t("太字"), todo: 1 }, { icon: "format_italic", label: t("斜体"), todo: 1 }, { icon: "format_underlined", label: t("下線"), todo: 1 },
           { icon: "border_all", label: t("罫線"), todo: 1 }, { icon: "format_color_fill", label: t("塗りつぶしの色"), todo: 1 }]
        ] }] },
        { label: t("配置"), items: [{ rows: [
          [{ icon: "vertical_align_top", label: t("上揃え"), todo: 1 }, { icon: "vertical_align_center", label: t("上下中央揃え"), todo: 1 }, { icon: "vertical_align_bottom", label: t("下揃え"), todo: 1 }],
          [{ icon: "format_align_left", label: t("左揃え"), todo: 1 }, { icon: "format_align_center", label: t("中央揃え"), todo: 1 }, { icon: "format_align_right", label: t("右揃え"), todo: 1 }]
        ] }] },
        { label: t("数値"), items: [{ rows: [
          [{ chip: t("標準"), w: 120, todo: 1 }],
          [{ icon: "currency_yen", label: t("通貨表示形式"), cmd: "yen" }, { icon: "percent", label: t("パーセント"), todo: 1 }, { icon: "decimal_increase", label: t("小数点以下の表示桁数を増やす"), todo: 1 }]
        ] }] },
        { label: t("編集"), items: [{ big: 1, icon: "functions", label: t("オート SUM"), cmd: "fn:SUM" }, { stack: [{ icon: "sort", label: t("並べ替え"), todo: 1 }, { icon: "filter_alt", label: t("フィルター"), todo: 1 }, { icon: "search", label: t("検索"), todo: 1 }] }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("テーブル"), items: [{ big: 1, icon: "table", label: t("テーブル"), todo: 1 }] },
        { label: t("グラフ"), items: [{ big: 1, icon: "bar_chart", label: t("縦棒"), todo: 1 }, { big: 1, icon: "show_chart", label: t("折れ線"), todo: 1 }, { big: 1, icon: "pie_chart", label: t("円"), todo: 1 }] }
      ] },
      { tab: t("数式"), groups: [
        { label: t("関数ライブラリ"), items: [{ big: 1, icon: "functions", label: t("オート SUM"), cmd: "fn:SUM" },
          { stack: [{ icon: "avg_pace", label: t("平均"), cmd: "fn:AVERAGE" }, { icon: "arrow_upward", label: t("最大値"), cmd: "fn:MAX" }, { icon: "arrow_downward", label: t("最小値"), cmd: "fn:MIN" }] },
          { stack: [{ icon: "tag", label: t("数値の個数"), cmd: "fn:COUNT" }] }] }
      ] },
      { tab: t("データ"), groups: [
        { label: t("並べ替えとフィルター"), items: [{ big: 1, icon: "sort_by_alpha", label: t("並べ替え"), todo: 1 }, { big: 1, icon: "filter_alt", label: t("フィルター"), todo: 1 }] }
      ] },
      { tab: t("表示"), groups: [ZOOM] }
    ],
    slide: [
      { tab: t("ホーム"), groups: [CLIP,
        { label: t("スライド"), items: [{ big: 1, icon: "add_box", label: t("新しいスライド"), cmd: "add" },
          { stack: [{ icon: "dashboard", label: t("レイアウト"), cmd: "layout" }, { icon: "content_copy", label: t("複製"), cmd: "dup" }, { icon: "delete", label: t("削除"), cmd: "del" }] }] },
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "M PLUS 1", w: 132, todo: 1 }, { chip: "28", w: 56, todo: 1 }],
          [{ icon: "format_bold", label: t("太字"), todo: 1 }, { icon: "format_italic", label: t("斜体"), todo: 1 }, { icon: "format_underlined", label: t("下線"), todo: 1 }, { icon: "format_color_text", label: t("フォントの色"), todo: 1 }]
        ] }] },
        { label: t("図形描画"), items: [{ big: 1, icon: "category", label: t("図形"), todo: 1 }, { big: 1, icon: "palette", label: t("クイック スタイル"), todo: 1 }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("スライド"), items: [{ big: 1, icon: "add_box", label: t("新しいスライド"), cmd: "add" }] },
        { label: t("画像"), items: [{ big: 1, icon: "image", label: t("画像"), todo: 1 }] },
        { label: t("図"), items: [{ big: 1, icon: "category", label: t("図形"), todo: 1 }, { big: 1, icon: "bar_chart", label: t("グラフ"), todo: 1 }] },
        { label: t("テキスト"), items: [{ big: 1, icon: "text_fields", label: t("テキスト ボックス"), todo: 1 }] }
      ] },
      { tab: t("デザイン"), groups: [
        { label: t("テーマ"), items: [{ big: 1, icon: "palette", label: t("テーマ"), todo: 1 }, { big: 1, icon: "aspect_ratio", label: t("スライドのサイズ"), todo: 1 }] }
      ] },
      { tab: t("スライド ショー"), groups: [
        { label: t("スライド ショーの開始"), items: [{ big: 1, icon: "play_arrow", label: t("最初から"), cmd: "present:0" }, { big: 1, icon: "slideshow", label: t("このスライドから"), cmd: "present:cur" }] }
      ] },
      { tab: t("表示"), groups: [ZOOM] }
    ]
  };
  var ribbonTab = 0, ribbonOpen = true, zoom = 100;

  function btnAttrs(x) { return x.todo ? ' data-todo="' + esc(x.label || x.chip) + '"' : ' data-cmd="' + esc(x.cmd) + '"' + (x.toggle ? ' data-toggle="' + x.toggle + '" aria-pressed="false"' : ""); }
  function ribbonItem(x) {
    if (x.big) return '<button type="button" class="rb-big m3-state"' + btnAttrs(x) + ">" + icon(x.icon) + "<span>" + esc(x.label) + "</span></button>";
    if (x.stack) return '<div class="rb-stack">' + x.stack.map(function (s) { return '<button type="button" class="rb-small m3-state"' + btnAttrs(s) + ">" + icon(s.icon) + "<span>" + esc(s.label) + "</span></button>"; }).join("") + "</div>";
    if (x.rows) return '<div class="rb-rows">' + x.rows.map(function (row) {
      return '<div class="rb-row">' + row.map(function (b) {
        if (b.chip) return '<button type="button" class="rb-chip m3-state" style="width:' + b.w + 'px"' + btnAttrs(b) + "><span>" + esc(b.chip) + "</span>" + icon("arrow_drop_down") + "</button>";
        return '<button type="button" class="rb-icon m3-state"' + btnAttrs(b) + ' aria-label="' + esc(b.label) + '" title="' + esc(b.label) + '">' + icon(b.icon) + "</button>";
      }).join("") + "</div>";
    }).join("") + "</div>";
    if (x.styles) return '<div class="rb-styles">' + x.styles.map(function (s) {
      return '<button type="button" class="rb-style rb-style--' + s.tag + ' m3-state" data-cmd="fmt:' + s.tag + '"><span class="rb-style__sample">' + esc(t("あア亜Aa")) + "</span><span>" + esc(s.label) + "</span></button>";
    }).join("") + "</div>";
    return "";
  }
  // prevTab: 切りかえる前のタブ（中身を右・左のどちらから入れるかを決める）
  function renderRibbon(prevTab) {
    var tabs = RIBBON[APP], row = $("rtabs");
    // タブの行は最初に 1 回だけ作り、あとは選んでいるタブと、折りたたみのボタンだけ変える（作り直すと、押したボタンが消えて下線もずれる）
    if (!row.firstChild) {
      row.innerHTML = '<button type="button" class="rtab rtab--file m3-state" data-file>' + esc(t("ファイル")) + "</button>" +
        tabs.map(function (tb, i) { return '<button type="button" class="rtab m3-state" role="tab" aria-selected="false" data-rtab="' + i + '"><span>' + esc(tb.tab) + "</span></button>"; }).join("") +
        '<span class="rtabs__space"></span><button type="button" class="icon-btn m3-state rtabs__fold" data-fold></button>' +
        '<span class="rtabs__ind" id="rtabs-ind" aria-hidden="true"></span>';
    }
    row.querySelectorAll("[data-rtab]").forEach(function (b) { b.setAttribute("aria-selected", String(+b.getAttribute("data-rtab") === ribbonTab)); });
    var fold = row.querySelector("[data-fold]"), foldLabel = ribbonOpen ? t("リボンを折りたたむ") : t("リボンを表示");
    fold.setAttribute("aria-expanded", String(ribbonOpen));
    fold.setAttribute("aria-label", foldLabel);
    fold.title = foldLabel;
    fold.innerHTML = icon(ribbonOpen ? "keyboard_arrow_up" : "keyboard_arrow_down");
    var body = $("rbody");
    var opening = body.hidden && ribbonOpen;
    body.hidden = !ribbonOpen;
    body.classList.remove("is-from-right", "is-from-left", "is-open");
    void body.offsetWidth;
    if (prevTab !== undefined && prevTab !== ribbonTab) body.classList.add(ribbonTab > prevTab ? "is-from-right" : "is-from-left");
    else if (opening) body.classList.add("is-open");
    body.innerHTML = tabs[ribbonTab].groups.map(function (g) {
      return '<div class="rgroup" role="group" aria-label="' + esc(g.label) + '"><div class="rgroup__items">' + g.items.map(ribbonItem).join("") + '</div><div class="rgroup__label">' + esc(g.label) + "</div></div>";
    }).join("");
    syncToggles();
    placeIndicator();
  }
  // 選んだタブの下線を、いまの場所から滑らせて動かす（Office と同じ。下線は 1 本だけで、CSS の transition で動く）。
  // animate: false のときは、すぐその場所に置く（開いたとき・画面の大きさが変わったとき）
  function placeIndicator(animate) {
    var ind = $("rtabs-ind"), tab = document.querySelector('.rtab[aria-selected="true"] > span'), box = $("rtabs");
    if (!ind || !tab || !box) return;
    var r = tab.getBoundingClientRect(), b = box.getBoundingClientRect();
    if (!r.width) return;
    if (animate === false || !ind.style.width) {
      ind.style.transition = "none";
      requestAnimationFrame(function () { requestAnimationFrame(function () { ind.style.transition = ""; }); });
    }
    ind.style.transform = "translateX(" + (r.left - b.left + box.scrollLeft) + "px)";
    ind.style.width = r.width + "px";
  }
  window.addEventListener("resize", function () { placeIndicator(false); });
  function syncToggles() {
    if (APP !== "doc") return;
    var page = $("page");
    var inside = page && document.activeElement === page;
    document.querySelectorAll("[data-toggle]").forEach(function (b) {
      var on = false;
      try { on = inside && document.queryCommandState(b.getAttribute("data-toggle")); } catch (e) { /* 使えないときはオフ */ }
      b.setAttribute("aria-pressed", String(!!on));
    });
  }
  document.addEventListener("selectionchange", function () { if (current && APP === "doc") syncToggles(); });

  function buildEditor() {
    var a = APPS[APP], ed = $("editor");
    if (!a || !ed) return;
    ed.innerHTML =
      '<header class="titlebar">' +
        '<button type="button" class="tbar-logo" data-file aria-label="' + esc(t("ファイル")) + '" title="' + esc(t("ファイル")) + '">' + icon(a.icon, true) + "</button>" +
        '<div class="qat" role="toolbar" aria-label="' + esc(t("クイック アクセス")) + '">' +
          '<button type="button" class="icon-btn m3-state" data-cmd="save" aria-label="' + esc(t("保存")) + '" title="' + esc(t("保存")) + '">' + icon("save") + "</button>" +
          '<button type="button" class="icon-btn m3-state" data-cmd="undo" aria-label="' + esc(t("元に戻す")) + '" title="' + esc(t("元に戻す")) + '">' + icon("undo") + "</button>" +
          '<button type="button" class="icon-btn m3-state" data-cmd="redo" aria-label="' + esc(t("やり直し")) + '" title="' + esc(t("やり直し")) + '">' + icon("redo") + "</button>" +
        "</div>" +
        '<button type="button" class="tbar-name m3-state" id="title-btn" title="' + esc(t("名前を変更")) + '"><span class="tbar-name__text" id="title-text" data-i18n-skip></span><span class="tbar-name__state" id="status"></span>' + icon("expand_more") + "</button>" +
        '<button type="button" class="tbar-search m3-state" data-todo="' + esc(t("検索")) + '">' + icon("search") + "<span>" + esc(t("検索")) + "</span></button>" +
        '<span class="titlebar__space"></span>' +
        '<button type="button" class="icon-btn m3-state hide-sm" data-todo="' + esc(t("コメント")) + '" aria-label="' + esc(t("コメント")) + '" title="' + esc(t("コメント")) + '">' + icon("comment") + "</button>" +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state tbar-share" data-todo="' + esc(t("共有")) + '">' + icon("share") + '<span class="hide-sm">' + esc(t("共有")) + "</span></button>" +
        '<button type="button" class="tbar-avatar" data-avatar-btn data-avatar aria-label="' + esc(t("アカウント")) + '" title="' + esc(t("アカウント")) + '"></button>' +
      "</header>" +
      '<div class="ribbon"><div class="rtabs" id="rtabs" role="tablist" aria-label="' + esc(t("リボン")) + '"></div><div class="rbody" id="rbody" role="toolbar"></div></div>' +
      '<div class="canvas" id="canvas" data-i18n-skip></div>' +
      '<footer class="statusbar"><span class="statusbar__info" id="stat"></span>' +
        '<span class="statusbar__space"></span>' +
        '<span class="views hide-sm">' +
          '<button type="button" class="sb-btn m3-state" data-todo="' + esc(t("閲覧モード")) + '" aria-label="' + esc(t("閲覧モード")) + '" title="' + esc(t("閲覧モード")) + '">' + icon("menu_book") + "</button>" +
          '<button type="button" class="sb-btn is-on m3-state" aria-label="' + esc(t("標準")) + '" title="' + esc(t("標準")) + '">' + icon(APP === "sheet" ? "grid_on" : APP === "slide" ? "view_sidebar" : "article") + "</button>" +
        "</span>" +
        '<span class="zoom"><button type="button" class="sb-btn m3-state" data-cmd="zoomOut" aria-label="' + esc(t("縮小")) + '">' + icon("remove") + "</button>" +
        '<input type="range" id="zoom" min="50" max="200" step="10" value="100" aria-label="' + esc(t("ズーム")) + '">' +
        '<button type="button" class="sb-btn m3-state" data-cmd="zoomIn" aria-label="' + esc(t("拡大")) + '">' + icon("add") + "</button>" +
        '<button type="button" class="sb-zoom m3-state" data-cmd="zoom" id="zoom-v">100%</button></span>' +
      "</footer>";
    $("title-btn").addEventListener("click", renameCurrent);
    $("zoom").addEventListener("input", function (e) { setZoom(+e.target.value); });
    // 文書の中の選んだところが外れないよう、リボンとクイック アクセスを押したときにフォーカスを動かさない
    ed.addEventListener("mousedown", function (e) { if (APP === "doc" && e.target.closest(".ribbon [data-cmd], .qat [data-cmd]")) e.preventDefault(); });
    renderRibbon();
  }
  function setTitle() {
    $("title-text").textContent = current.name;
    document.title = current.name + " — " + APPS[APP].name + " | SK Hub Office";
  }
  function setStatus(s) { var el = $("status"); if (el) el.textContent = s ? " ・ " + s : ""; }
  // 保存（自動保存がオンなら、変えてから少しして保存。オフなら、保存のボタンか Ctrl+S で保存）
  var saveTimer = null, dirty = false;
  function saveNow() {
    clearTimeout(saveTimer); saveTimer = null;
    if (!current) return Promise.resolve(false);
    var f = current;
    dirty = false;
    setStatus(t("保存しています…"));
    return Store.save(f).then(function () {
      if (current === f && !dirty) setStatus(t("保存済み"));
      return true;
    }, function (e) { dirty = true; setStatus(t("保存できませんでした")); failed(e); return false; });
  }
  function saveSoon() {
    dirty = true;
    if (!settings.autosave) { setStatus(t("未保存")); return; }
    setStatus(t("保存しています…"));
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, 600);
  }
  function flush() { if (saveTimer && current) saveNow(); }
  window.addEventListener("pagehide", flush);
  // 保存していない変更があるときは、閉じる前に確かめる
  window.addEventListener("beforeunload", function (e) { if (current && (dirty || saveTimer)) { flush(); e.preventDefault(); e.returnValue = ""; } });
  function renameCurrent() {
    if (!current) return;
    var f = current;
    (saveTimer ? saveNow() : Promise.resolve()).then(function () {
      return renameFile(f, function () {
        if (current !== f) return;
        history.replaceState(null, "", "#f=" + encodeURIComponent(f.id));
        setTitle();
        setStatus(t("保存済み"));
      });
    });
  }
  function closeFile() {
    (dirty ? saveNow() : Promise.resolve()).then(function () {
      current = null; dirty = false;
      history.pushState(null, "", location.pathname);
      route();
    });
  }

  function setZoom(z) {
    zoom = Math.max(50, Math.min(200, Math.round(z / 10) * 10));
    var r = $("zoom"); if (r) r.value = zoom;
    var v = $("zoom-v"); if (v) v.textContent = zoom + "%";
    var target = document.querySelector("#page, #grid-scroll .grid, #stage");
    if (target) target.style.zoom = zoom / 100;
  }

  var COMMANDS = {}, RENDER = {};
  function runCmd(cmd) {
    if (!current) return;
    if (cmd === "save") {
      saveNow().then(function (ok) {
        if (ok) notify(Store.state() === "ready" ? t("フォルダ「{folder}」に保存しました", { folder: Store.folderName() }) : t("このブラウザに保存しました"));
      });
      return;
    }
    if (cmd === "zoomIn") return setZoom(zoom + 10);
    if (cmd === "zoomOut") return setZoom(zoom - 10);
    if (cmd === "zoom100") return setZoom(100);
    if (cmd === "zoom") {
      M3.choose({ title: t("ズーム"), value: String(zoom), options: ["200", "150", "125", "100", "75", "50"].map(function (z) { return { value: z, label: z + "%" }; }) })
        .then(function (v) { if (v) setZoom(+v); });
      return;
    }
    COMMANDS[APP](cmd);
  }

  function openEditor(f) {
    current = f;
    setTitle();
    setStatus(t("保存済み"));
    RENDER[APP]();
    setZoom(zoom);
    placeIndicator(false);
  }

  // ---------- Doc ----------
  RENDER.doc = function () {
    $("canvas").innerHTML = '<div class="doc-scroll"><article class="doc-page" id="page" contenteditable="true" spellcheck="false" aria-label="' + esc(t("文書の本文")) + '"></article></div>';
    var page = $("page");
    page.innerHTML = current.data.html || "<p><br></p>";
    page.addEventListener("input", function () { current.data.html = page.innerHTML; docStat(); saveSoon(); });
    page.addEventListener("focus", syncToggles);
    page.addEventListener("blur", syncToggles);
    docStat();
  };
  function docStat() {
    var page = $("page"), n = page.textContent.replace(/\s/g, "").length;
    var pages = Math.max(1, Math.ceil(page.scrollHeight / 1056));
    $("stat").innerHTML = "<span>" + esc(t("{n} / {m} ページ", { n: 1, m: pages })) + "</span><span>" + esc(t("{n} 文字", { n: n })) + "</span><span class=\"hide-sm\">" + esc(t("日本語")) + "</span>";
  }
  COMMANDS.doc = function (cmd) {
    var page = $("page");
    if (document.activeElement !== page) page.focus();
    if (cmd.indexOf("fmt:") === 0) document.execCommand("formatBlock", false, cmd.slice(4));
    else if (cmd === "table") document.execCommand("insertHTML", false, "<table><tbody>" + [0, 1, 2].map(function () { return "<tr><td><br></td><td><br></td><td><br></td></tr>"; }).join("") + "</tbody></table><p><br></p>");
    else if (cmd === "hr") document.execCommand("insertHorizontalRule");
    else if (cmd === "date") { var d = new Date(); document.execCommand("insertText", false, t("{y}年{m}月{d}日", { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() })); }
    else if (cmd === "count") { notify(t("{n} 文字（空白を除く）", { n: page.textContent.replace(/\s/g, "").length })); return; }
    else if (cmd === "link") {
      var sel = getSelection(), range = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
      M3.prompt({ title: t("リンク"), label: t("アドレス（URL）"), placeholder: "https://", ok: t("挿入") }).then(function (url) {
        url = (url || "").trim();
        if (!url) return;
        if (!/^https?:\/\//i.test(url)) url = "https://" + url;
        page.focus();
        if (range) { sel.removeAllRanges(); sel.addRange(range); }
        if (range && !range.collapsed) document.execCommand("createLink", false, url);
        else document.execCommand("insertHTML", false, '<a href="' + esc(url) + '">' + esc(url) + "</a>");
        current.data.html = page.innerHTML; saveSoon();
      });
      return;
    }
    else document.execCommand(cmd, false, null);
    current.data.html = page.innerHTML;
    docStat();
    syncToggles();
    saveSoon();
  };

  // ---------- Sheet ----------
  var COLS = "ABCDEFGHIJKL".split("");
  var sel = "A1";
  function cellsOf(d) { return d.cells || (d.cells = {}); }
  function num(v) {
    if (typeof v === "number") return v;
    var n = parseFloat(String(v).replace(/[¥,\s]/g, ""));
    return isNaN(n) ? null : n;
  }
  function value(d, ref, depth) {
    var raw = cellsOf(d)[ref];
    if (raw === undefined || raw === "") return "";
    if (raw.charAt(0) !== "=") { var n = num(raw); return n !== null && /^[¥\d,.\s-]+$/.test(raw) ? n : raw; }
    if ((depth || 0) > 20) return "#エラー";
    return formula(d, raw.slice(1).trim().toUpperCase(), (depth || 0) + 1);
  }
  function operand(d, s, depth) {
    if (/^[A-L]\d+$/.test(s)) { var v = num(value(d, s, depth)); return v === null ? 0 : v; }
    return parseFloat(s);
  }
  function formula(d, f, depth) {
    var m = f.match(/^(SUM|AVERAGE|MAX|MIN|COUNT)\(([A-L])(\d+):([A-L])(\d+)\)$/);
    if (m) {
      var vals = [];
      for (var c = Math.min(COLS.indexOf(m[2]), COLS.indexOf(m[4])); c <= Math.max(COLS.indexOf(m[2]), COLS.indexOf(m[4])); c++) {
        for (var r = Math.min(+m[3], +m[5]); r <= Math.max(+m[3], +m[5]); r++) {
          var v = num(value(d, COLS[c] + r, depth));
          if (v !== null) vals.push(v);
        }
      }
      if (m[1] === "COUNT") return vals.length;
      if (!vals.length) return 0;
      if (m[1] === "SUM") return vals.reduce(function (a, b) { return a + b; }, 0);
      if (m[1] === "AVERAGE") return vals.reduce(function (a, b) { return a + b; }, 0) / vals.length;
      return Math[m[1].toLowerCase()].apply(null, vals);
    }
    m = f.match(/^([A-L]\d+|\d+(?:\.\d+)?)\s*([-+*/])\s*([A-L]\d+|\d+(?:\.\d+)?)$/);
    if (m) {
      var a = operand(d, m[1], depth), b = operand(d, m[3], depth);
      if (m[2] === "+") return a + b;
      if (m[2] === "-") return a - b;
      if (m[2] === "*") return a * b;
      return b === 0 ? "#エラー" : a / b;
    }
    if (/^[A-L]\d+$/.test(f)) return value(d, f, depth);
    return "#エラー";
  }
  function shown(d, ref) {
    var v = value(d, ref);
    if (typeof v !== "number") return { text: String(v), n: false, v: null };
    var yen = (d.fmt || {})[ref.charAt(0)] === "yen";
    return { text: (yen ? "¥" : "") + (Math.round(v * 100) / 100).toLocaleString("ja-JP"), n: true, v: v };
  }
  function sheetTable(d, rows, cols, full) {
    var h = '<table class="grid">' + (full ? '<thead><tr><th class="corner"></th>' + COLS.slice(0, cols).map(function (c) { return '<th data-col="' + c + '">' + c + "</th>"; }).join("") + "</tr></thead>" : "") + "<tbody>";
    for (var r = 1; r <= rows; r++) {
      var cls = r === d.head ? "head" : r === d.total ? "total" : "";
      h += '<tr class="' + cls + '">' + (full ? '<th data-row="' + r + '">' + r + "</th>" : "");
      for (var c = 0; c < cols; c++) {
        var ref = COLS[c] + r, s = shown(d, ref);
        h += '<td data-ref="' + ref + '"' + (s.n ? ' class="n"' : "") + ">" + esc(s.text) + "</td>";
      }
      h += "</tr>";
    }
    return h + "</tbody></table>";
  }
  RENDER.sheet = function () {
    var d = current.data, rows = 50;
    Object.keys(cellsOf(d)).forEach(function (k) { rows = Math.max(rows, +k.slice(1) + 20); });
    $("canvas").innerHTML =
      '<div class="sheet">' +
        '<div class="fx-bar"><span class="namebox" id="namebox">A1</span>' +
          '<span class="fx-btns hide-sm"><span class="msr" aria-hidden="true">close</span><span class="msr" aria-hidden="true">check</span></span>' +
          '<span class="fx-ico">fx</span>' +
          '<input class="fx-input" id="fx" autocomplete="off" spellcheck="false" aria-label="' + esc(t("数式バー")) + '"></div>' +
        '<div class="grid-scroll" id="grid-scroll" tabindex="0" aria-label="' + esc(t("ワークシート")) + '">' + sheetTable(d, rows, COLS.length, true) + "</div>" +
        '<div class="sheet-tabs"><span class="sheet-tabs__nav hide-sm"><span class="msr" aria-hidden="true">chevron_left</span><span class="msr" aria-hidden="true">chevron_right</span></span>' +
          '<button type="button" class="stab is-on m3-state">' + esc(t("Sheet1")) + "</button>" +
          '<button type="button" class="icon-btn m3-state" data-todo="' + esc(t("新しいシート")) + '" aria-label="' + esc(t("新しいシート")) + '" title="' + esc(t("新しいシート")) + '">' + icon("add_circle") + "</button></div>" +
      "</div>";
    var gridEl = $("grid-scroll"), fx = $("fx");
    gridEl.addEventListener("click", function (e) { var td = e.target.closest("td[data-ref]"); if (td) select(td.getAttribute("data-ref")); });
    gridEl.addEventListener("dblclick", function () { fx.focus(); fx.select(); });
    gridEl.addEventListener("keydown", function (e) {
      var c = COLS.indexOf(sel.charAt(0)), r = +sel.slice(1);
      var moves = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0], Enter: [0, 1], Tab: [1, 0] };
      if (moves[e.key]) {
        e.preventDefault();
        select(COLS[Math.max(0, Math.min(COLS.length - 1, c + moves[e.key][0]))] + Math.max(1, r + moves[e.key][1]));
      } else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); commit(sel, ""); }
      else if (e.key === "F2") { e.preventDefault(); fx.focus(); }
      else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); fx.value = e.key; fx.focus(); }
    });
    fx.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); commit(sel, fx.value); select(sel.charAt(0) + (+sel.slice(1) + 1)); gridEl.focus({ preventScroll: true }); }
      if (e.key === "Escape") { fx.value = cellsOf(current.data)[sel] || ""; gridEl.focus({ preventScroll: true }); }
    });
    fx.addEventListener("change", function () { commit(sel, fx.value); });
    select("A1");
    gridEl.focus({ preventScroll: true });
  };
  function select(ref) {
    var gridEl = $("grid-scroll");
    if (!gridEl) return;
    var td = gridEl.querySelector('[data-ref="' + ref + '"]');
    if (!td) return;
    gridEl.querySelectorAll(".sel, .hl").forEach(function (x) { x.classList.remove("sel", "hl"); });
    sel = ref;
    td.classList.add("sel");
    var colH = gridEl.querySelector('[data-col="' + ref.charAt(0) + '"]'), rowH = gridEl.querySelector('[data-row="' + ref.slice(1) + '"]');
    if (colH) colH.classList.add("hl");
    if (rowH) rowH.classList.add("hl");
    td.scrollIntoView({ block: "nearest", inline: "nearest" });
    $("namebox").textContent = ref;
    $("fx").value = cellsOf(current.data)[ref] || "";
    var v = shown(current.data, ref);
    $("stat").innerHTML = "<span>" + esc(t("準備完了")) + "</span>" + (v.n ? "<span>" + esc(t("合計: {v}", { v: v.text })) + "</span>" : "");
  }
  function refreshCells() {
    $("grid-scroll").querySelectorAll("td[data-ref]").forEach(function (td) {
      var s = shown(current.data, td.getAttribute("data-ref"));
      td.textContent = s.text;
      td.classList.toggle("n", s.n);
    });
  }
  function commit(ref, raw) {
    var cells = cellsOf(current.data);
    if ((cells[ref] || "") === raw) return;
    if (raw === "") delete cells[ref]; else cells[ref] = raw;
    refreshCells();
    saveSoon();
  }
  COMMANDS.sheet = function (cmd) {
    var d = current.data, col = sel.charAt(0);
    if (cmd === "undo" || cmd === "redo") { notify(t("「{name}」は、まだ使えません", { name: cmd === "undo" ? t("元に戻す") : t("やり直し") })); return; }
    if (cmd === "yen") {
      d.fmt = d.fmt || {};
      if (d.fmt[col] === "yen") delete d.fmt[col]; else d.fmt[col] = "yen";
      refreshCells();
      select(sel);
      saveSoon();
      notify(d.fmt[col] ? t("{col} 列を円で表示します", { col: col }) : t("{col} 列の円の表示をやめました", { col: col }));
    }
    if (cmd.indexOf("fn:") === 0) {
      // 選んだセルの上に続いている数を、その関数で計算する（Excel のオート SUM と同じ）
      var r = +sel.slice(1), top = r - 1;
      while (top >= 1 && num(value(d, col + top)) !== null) top--;
      if (top === r - 1) { notify(t("上に数が入ったセルがありません")); return; }
      commit(sel, "=" + cmd.slice(3) + "(" + col + (top + 1) + ":" + col + (r - 1) + ")");
      select(sel);
    }
  };

  // ---------- Slide ----------
  var cur = 0;
  function slideHtml(s, editable) {
    var ce = editable ? ' contenteditable="true" spellcheck="false"' : "";
    if (s.layout === "blank") return '<div class="sl sl--title"><div class="sl-in"></div></div>';
    var h = '<div class="sl sl--' + (s.layout || "title") + '"><i class="sl-blob" aria-hidden="true"></i><div class="sl-in">';
    h += '<h2 class="sl-title" data-k="title"' + ce + ">" + esc(s.title || "") + "</h2>";
    if (s.layout === "bullets") {
      h += '<ul class="sl-list">' + (s.items && s.items.length ? s.items : [""]).map(function (it, i) { return '<li data-k="items.' + i + '"' + ce + ">" + esc(it) + "</li>"; }).join("") + "</ul>";
    } else {
      h += '<p class="sl-sub" data-k="sub"' + ce + ">" + esc(s.sub || "") + "</p>";
    }
    return h + "</div></div>";
  }
  RENDER.slide = function () {
    $("canvas").innerHTML = '<div class="deck"><div class="deck-thumbs" id="thumbs" role="listbox" aria-label="' + esc(t("スライド")) + '"></div>' +
      '<div class="deck-stage"><div class="deck-slide" id="stage"></div>' +
      '<div class="deck-notes" data-todo="' + esc(t("ノート")) + '">' + esc(t("クリックしてノートを追加")) + "</div></div></div>";
    cur = 0;
    drawThumbs();
    drawStage();
    $("thumbs").addEventListener("click", function (e) {
      var b = e.target.closest("[data-slide]");
      if (b) { cur = +b.getAttribute("data-slide"); drawThumbs(); drawStage(); }
    });
    var stage = $("stage");
    stage.addEventListener("input", function (e) {
      var k = e.target.getAttribute && e.target.getAttribute("data-k");
      if (!k) return;
      var s = current.data.slides[cur];
      if (k.indexOf("items.") === 0) { s.items = s.items || []; s.items[+k.slice(6)] = e.target.textContent; }
      else s[k] = e.target.textContent;
      drawThumbs(true);
      saveSoon();
    });
    // 箇条書き: Enter で次の行を足す・空の行で Backspace で消す
    stage.addEventListener("keydown", function (e) {
      var k = e.target.getAttribute && e.target.getAttribute("data-k");
      if (!k) return;
      var s = current.data.slides[cur];
      if (k.indexOf("items.") === 0) {
        var i = +k.slice(6);
        if (e.key === "Enter") { e.preventDefault(); s.items.splice(i + 1, 0, ""); drawStage(i + 1); saveSoon(); }
        else if (e.key === "Backspace" && !e.target.textContent && s.items.length > 1) { e.preventDefault(); s.items.splice(i, 1); drawStage(Math.max(0, i - 1)); drawThumbs(); saveSoon(); }
      } else if (e.key === "Enter") { e.preventDefault(); }
    });
  };
  function drawThumbs(keepScroll) {
    var slides = current.data.slides;
    $("thumbs").innerHTML = slides.map(function (s, i) {
      return '<button type="button" class="deck-thumb" role="option" aria-selected="' + (i === cur) + '" data-slide="' + i + '"><span class="deck-no">' + (i + 1) + '</span><span class="th th--slide">' + slideHtml(s, false) + "</span></button>";
    }).join("");
    $("stat").innerHTML = "<span>" + esc(t("スライド {i}/{n}", { i: cur + 1, n: slides.length })) + '</span><span class="hide-sm">' + esc(t("日本語")) + "</span>";
    if (!keepScroll) { var on = $("thumbs").querySelector('[aria-selected="true"]'); if (on) on.scrollIntoView({ block: "nearest", inline: "nearest" }); }
  }
  function drawStage(focusItem) {
    $("stage").innerHTML = slideHtml(current.data.slides[cur], true);
    if (focusItem !== undefined) {
      var li = $("stage").querySelector('[data-k="items.' + focusItem + '"]');
      if (li) { li.focus(); var r = document.createRange(); r.selectNodeContents(li); r.collapse(false); var s = getSelection(); s.removeAllRanges(); s.addRange(r); }
    }
  }
  COMMANDS.slide = function (cmd) {
    var slides = current.data.slides;
    if (cmd === "undo" || cmd === "redo") { notify(t("「{name}」は、まだ使えません", { name: cmd === "undo" ? t("元に戻す") : t("やり直し") })); return; }
    if (cmd.indexOf("present:") === 0) { present(cmd === "present:0" ? 0 : cur); return; }
    if (cmd === "add") { slides.splice(cur + 1, 0, { layout: "bullets", title: t("タイトルを入力"), items: [t("テキストを入力")] }); cur++; }
    if (cmd === "dup") { slides.splice(cur + 1, 0, copy(slides[cur])); cur++; }
    if (cmd === "del") {
      if (slides.length < 2) { notify(t("スライドが 1 枚のときは削除できません")); return; }
      M3.confirm({ title: t("このスライドを削除しますか？"), text: t("スライド {i} を削除します。", { i: cur + 1 }), ok: t("削除"), danger: true, icon: "delete" }).then(function (ok) {
        if (!ok) return;
        slides.splice(cur, 1); cur = Math.min(cur, slides.length - 1);
        drawThumbs(); drawStage(); saveSoon();
      });
      return;
    }
    if (cmd === "layout") {
      M3.choose({ title: t("レイアウト"), value: slides[cur].layout, options: [
        { value: "title", label: t("タイトル スライド"), icon: "title" }, { value: "bullets", label: t("タイトルとコンテンツ"), icon: "format_list_bulleted" }, { value: "section", label: t("セクション見出し"), icon: "splitscreen" }
      ] }).then(function (v) {
        if (!v) return;
        var s = slides[cur];
        s.layout = v;
        if (v === "bullets" && !(s.items && s.items.length)) s.items = [s.sub || t("テキストを入力")];
        drawThumbs(); drawStage(); saveSoon();
      });
      return;
    }
    drawThumbs(); drawStage(); saveSoon();
  };
  // スライド ショー（画面いっぱいに出す。← → で前後、Esc で終わる）
  function present(start) {
    flush();
    var slides = current.data.slides, i = start;
    var box = document.createElement("div");
    box.className = "present";
    box.setAttribute("role", "dialog");
    box.setAttribute("aria-label", t("スライド ショー"));
    box.innerHTML = '<div class="present__slide" id="present-slide"></div><div class="present__bar"><span id="present-no"></span><span>' + esc(t("← → で移動・Esc で終わる")) + "</span></div>";
    document.body.appendChild(box);
    function draw() { $("present-slide").innerHTML = slideHtml(slides[i], false); $("present-no").textContent = (i + 1) + " / " + slides.length; }
    function go(n) { if (n >= slides.length || n < 0) { if (n >= slides.length) end(); return; } i = n; draw(); }
    function end() {
      document.removeEventListener("keydown", onKey, true);
      if (document.fullscreenElement) document.exitFullscreen().catch(function () {});
      box.remove();
    }
    function onKey(e) {
      if (e.key === "Escape") { e.preventDefault(); end(); }
      else if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown" || e.key === "Enter") { e.preventDefault(); go(i + 1); }
      else if (e.key === "ArrowLeft" || e.key === "PageUp" || e.key === "Backspace") { e.preventDefault(); go(i - 1); }
    }
    box.addEventListener("click", function () { go(i + 1); });
    document.addEventListener("keydown", onKey, true);
    if (box.requestFullscreen) box.requestFullscreen().catch(function () {});
    document.addEventListener("fullscreenchange", function fs() { if (!document.fullscreenElement && box.isConnected) { document.removeEventListener("fullscreenchange", fs); end(); } });
    draw();
  }

  // ================================================================
  //   クリック・入力・キー
  // ================================================================
  document.addEventListener("click", function (e) {
    var el;
    if ((el = e.target.closest("[data-todo]"))) return notify(t("「{name}」は、まだ使えません", { name: el.getAttribute("data-todo") }));
    if ((el = e.target.closest("[data-tpl]"))) return createFrom(el.getAttribute("data-tpl"));
    if ((el = e.target.closest("[data-open]"))) return openFile(el.getAttribute("data-open"), el.getAttribute("data-open-app"));
    if ((el = e.target.closest("[data-more]"))) return fileActions(el.getAttribute("data-more"), el.getAttribute("data-name"));
    if ((el = e.target.closest("[data-pin]"))) {
      var pid = el.getAttribute("data-pin"), on = !isPinned(pid);
      setPinned(pid, on);
      renderFileTable();
      notify(on ? t("ピン留めしました") : t("ピン留めを外しました"));
      return;
    }
    if (e.target.closest("[data-pick]")) {
      Store.pick().then(function (ok) { if (ok) { notify(t("フォルダ「{name}」に保存します", { name: Store.folderName() })); afterFolder(); } });
      return;
    }
    if (e.target.closest("[data-allow]")) {
      Store.allow().then(function (ok) { if (ok) afterFolder(); else notify(t("許可されなかったので、フォルダを使えません")); });
      return;
    }
    if (e.target.closest("[data-migrate]")) {
      Store.migrate().then(function (n) { notify(t("{n} 件のファイルをフォルダに移しました", { n: n })); renderStart(); }).catch(failed);
      return;
    }
    if (e.target.closest("[data-import]")) return importFile();
    if (e.target.closest("[data-save-as]")) {
      if (current) Store.exportFile(current).then(function (ok) { if (ok) notify(t("コピーを保存しました")); }).catch(failed);
      return;
    }
    if (e.target.closest("[data-close-file]")) return closeFile();
    if (e.target.closest("[data-install]")) {
      if (installEvt) { installEvt.prompt(); installEvt.userChoice.then(function () { installEvt = null; renderStart(); }); }
      return;
    }
    if ((el = e.target.closest("[data-theme-set]"))) { settings.theme = el.getAttribute("data-theme-set"); saveSettings(); applySettings(); renderStart(); return; }
    if (e.target.closest("[data-avatar-btn]")) { flush(); view = "account"; renderStart(); show("home"); return; }
    if ((el = e.target.closest("[data-view]"))) { view = el.getAttribute("data-view"); query = ""; renderStart(); var m = $("start-main"); if (m) m.scrollTop = 0; return; }
    if ((el = e.target.closest("[data-ftab]"))) {
      // スタート画面全体は作り直さず、タブの印とファイルの一覧だけ変える
      fileTab = el.getAttribute("data-ftab");
      el.parentElement.querySelectorAll("[data-ftab]").forEach(function (b) { b.setAttribute("aria-selected", String(b === el)); });
      animateFiles = true;
      renderFileTable();
      return;
    }
    if ((el = e.target.closest("#cat-chips .chip"))) return filterTemplates(el.getAttribute("data-cat"));
    if (e.target.closest("[data-back-edit]")) {
      var ed = $("editor");
      ed.classList.remove("is-back"); show("editor"); void ed.offsetWidth; ed.classList.add("is-back");
      placeIndicator(false);
      return;
    }
    if (e.target.closest("[data-file]")) { flush(); view = "home"; renderStart(); $("home").classList.add("is-backstage"); show("home"); return; }
    if ((el = e.target.closest("[data-rtab]"))) { var prev = ribbonTab; ribbonTab = +el.getAttribute("data-rtab"); ribbonOpen = true; renderRibbon(prev); return; }
    if (e.target.closest("[data-fold]")) { ribbonOpen = !ribbonOpen; renderRibbon(); return; }
    if ((el = e.target.closest("[data-cmd]"))) return runCmd(el.getAttribute("data-cmd"));
  });
  document.addEventListener("input", function (e) {
    if (e.target.id === "file-search") { query = e.target.value.trim(); renderFileTable(); }
  });
  // オプション
  document.addEventListener("change", function (e) {
    var id = e.target.id;
    if (id === "opt-user") { settings.user = e.target.value.trim().slice(0, 40); saveSettings(); fillAvatars(); notify(settings.user ? t("名前を「{name}」にしました", { name: settings.user }) : t("名前を消しました")); }
    if (id === "opt-motion") { settings.motion = e.target.checked ? "reduce" : "auto"; saveSettings(); applySettings(); }
    if (id === "opt-autosave") { settings.autosave = e.target.checked; saveSettings(); if (settings.autosave && dirty) saveNow(); notify(settings.autosave ? t("自動保存をオンにしました") : t("自動保存をオフにしました")); }
  });
  document.addEventListener("keydown", function (e) {
    if (!current || !$("editor") || $("editor").hidden) return;
    var k = e.key.toLowerCase();
    if ((e.ctrlKey || e.metaKey) && k === "s") { e.preventDefault(); runCmd("save"); }
    if (APP === "slide" && e.key === "F5") { e.preventDefault(); present(e.shiftKey ? cur : 0); }
  });

  // ================================================================
  //   画面の切りかえ
  // ================================================================
  function show(v) {
    ["home", "editor"].forEach(function (id) { var el = $(id); if (el) el.hidden = id !== v; });
  }
  function showStart() {
    current = null; dirty = false;
    view = "home";
    document.title = (APPS[APP] ? APPS[APP].name + " — " : "") + "SK Hub Office";
    renderStart();
    show("home");
  }
  function route() {
    flush();
    var m = location.hash.match(/^#f=(.+)$/);
    var id = m && APPS[APP] ? decodeURIComponent(m[1]) : null;
    if (!id || !$("editor")) { showStart(); return Promise.resolve(); }
    // フォルダの許可がまだのときは、スタート画面で「許可する」を押してもらう（押したら、このファイルを開く）
    if (Store.supported && Store.state() !== "ready" && !Store.isOutside(id)) { showStart(); return Promise.resolve(); }
    return Store.get(id).then(function (f) {
      if (f && f.app === APP) { show("editor"); openEditor(f); }
      else { history.replaceState(null, "", location.pathname); notify(t("ファイルが見つかりませんでした")); showStart(); }
    });
  }
  // フォルダを選んだ・許可したあと（開こうとしていたファイルがあれば開く）
  function afterFolder(quiet) {
    if (/^#f=/.test(location.hash) && !current) route();
    else if (!quiet || $("home") && !$("home").hidden) renderStart();
  }
  window.addEventListener("popstate", route);
  // ほかのタブで設定を変えたとき
  window.addEventListener("storage", function (e) {
    if (e.key === SET_KEY) { try { var d = JSON.parse(e.newValue || "{}"); Object.keys(settings).forEach(function (k) { if (k in d) settings[k] = d[k]; }); } catch (x) { /* そのまま */ } applySettings(); }
    if ((e.key === "sk_office_files" || e.key === SET_KEY) && $("home") && !$("home").hidden) renderFileTable();
  });

  applySettings();
  buildEditor();
  fillAvatars();
  Store.init().then(route, route).then(function () { if (window.SKI18N) SKI18N.translatePage(); });
  // オフラインでも開けるように（/office/sw.js）
  if ("serviceWorker" in navigator && /^https?:$/.test(location.protocol)) {
    navigator.serviceWorker.register(BASE + "sw.js", { scope: BASE }).catch(function () { /* なくても使える */ });
  }
  // アプリとして入れたあと、エクスプローラー・Finder で .skdoc などをダブルクリックしたとき（manifest の file_handlers）
  if ("launchQueue" in window && APPS[APP]) {
    window.launchQueue.setConsumer(function (params) {
      if (!params.files || !params.files.length) return;
      Store.init().then(function () { return Store.openHandle(params.files[0]); }).then(function (id) { openFile(id); }).catch(failed);
    });
  }
  // 起動画面は、少し見せてから消す
  var splash = $("splash");
  if (splash) setTimeout(function () {
    splash.classList.add("is-done");
    setTimeout(function () { splash.hidden = true; }, reduced ? 0 : 420);
  }, reduced ? 0 : 900);
})();
