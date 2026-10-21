// SK Hub Office の共通（Doc・Sheet・Slide と入り口。開くのは https://office.sk-hub.jp/。中身は sk-hub.jp の /office/ を _workers/office/ が中継する）
//   画面の形は Microsoft Office（左に色付きのメニューがあるスタート画面・タイトルバー・タブとグループのリボン・ズーム付きのステータスバー）。
//   見た目は SK Material（部品は /toolbox/shared/m3.css、色は office.css）。
//   <html data-app="doc|sheet|slide|hub"> で、どのアプリかを決める。画面の中身はここで組み立てる（各ページは起動画面と入れ物だけ）
//   ・起動画面（#splash）→ スタート画面（ホーム・新規・開く）→ エディタ（#f=<ファイルの ID>）。エディタの「ファイル」でスタート画面に戻る
//     ファイルを開いているときのスタート画面（Office の「バックステージ」）には、情報（版の履歴）・エクスポート・印刷も出る
//   ・アプリの中身（リボン・編集・書き出し）は、アプリごとのファイル（doc.js・sheet.js・slide.js）。このファイルより先に読み込み、
//     window.OfficeApps[アプリ] = function (O) { return { ribbon, render, command, … } } で登録する。O はこのファイルが渡す道具（下の「アプリに渡す道具」）
//   ・ファイルは PC の選んだフォルダに 1 つずつ保存する（フォルダに保存できないブラウザは、ブラウザの中）。中身は /office/shared/store.js
//   ・元に戻す・やり直しは、ファイルの中身（data）をまるごと覚えておく（どのアプリも同じ。O.changed() を呼ぶと覚える）
//   ・アカウント（SK Hub Systems アカウントのログインの状態・製品情報）とオプション（ユーザー名・テーマ・動き・保存先・自動保存）は、スタート画面の左のメニューから開く。
//     ログインは sk-hub.jp の /account/ で行う。office.sk-hub.jp では、ログイン中かのヒントを /assets/hub/bridge.html から受け取る（下の「SK Hub Systems アカウント」）。
//     ファイルをインターネット上に保存するしくみは、まだない
//     オプションは localStorage の sk_office_settings に保存する
//   ・PWA: /office/manifest.json（アプリ全体を 1 つの「SK Hub Office」として入れる。Doc・Sheet・Slide はショートカット）と /office/sw.js（オフライン）。
//     入れたあとは、.skdoc・.sksheet・.skslide をダブルクリックすると、そのアプリで開く（file_handlers・launchQueue）
//   ・まだできない操作のボタンは、押すと「まだ使えません」と出す（data-todo）
//   ・ダイアログは /toolbox/shared/m3.js の M3.prompt・M3.choose・M3.confirm・M3.dialog。ボタンの下に出るメニューは O.menu・色は O.palette
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
  var VERSION = "0.3";
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
  // アカウントのアイコン（ログイン中はアカウントの写真か頭文字と色。そうでなければ名前の 1 文字目。名前がなければ人のアイコン）
  function fillAvatars() {
    var n = userName(), h = accountHint();
    document.querySelectorAll("[data-avatar]").forEach(function (el) {
      el.innerHTML = h && h.photo ? '<img src="' + esc(h.photo) + '" alt="" referrerpolicy="no-referrer">' : n ? esc(Array.from(n)[0].toUpperCase()) : icon("person");
      el.style.background = h && !h.photo && /^#[0-9a-f]{6}$/i.test(h.color || "") ? h.color : "";
      el.style.color = el.style.background ? "#fff" : "";
    });
  }

  // ---------- SK Hub Systems アカウント（ログイン中かどうか） ----------
  //   ログイン・ログアウトは sk-hub.jp の /account/ で行い、終わったら Office に戻ってくる。Office が使うのは、ログイン中かのヒント
  //   （skhub_account: 名前・メールアドレス・写真。assets/hub/account.js が保存する）だけで、ファイルはいままでどおり PC に保存する（オンライン保存はまだ）。
  //   office.sk-hub.jp からは sk-hub.jp の保存場所が見えないので、/assets/hub/bridge.html（見えない iframe）から受け取って、このページの localStorage に写す。
  //   手元（localhost）や sk-hub.jp では同じ保存場所なので、そのまま読む
  var ON_SUB = /^office\.sk-hub\.jp$/.test(location.hostname);
  var bridge = (function () {
    if (!ON_SUB) return null;
    var ORIGIN = "https://sk-hub.jp", waiting = {}, seq = 0, frame = null, loaded = null;
    function open() {
      if (loaded) return loaded;
      loaded = new Promise(function (resolve) {
        window.addEventListener("message", function (e) {
          var m = e.data;
          if (e.origin !== ORIGIN || !m || m.sk !== "bridge") return;
          if (m.ready) { resolve(); return; }
          if (waiting[m.id]) { waiting[m.id](m.value); delete waiting[m.id]; }
        });
        frame = document.createElement("iframe");
        frame.src = ORIGIN + "/assets/hub/bridge.html";
        frame.hidden = true; frame.title = ""; frame.tabIndex = -1;
        frame.setAttribute("aria-hidden", "true");
        document.body.appendChild(frame);
      });
      return loaded;
    }
    return {
      ask: function (op, timeout) {
        return Promise.race([
          open().then(function () { return new Promise(function (resolve) { var id = ++seq; waiting[id] = resolve; frame.contentWindow.postMessage({ sk: "bridge", id: id, op: op }, ORIGIN); }); }),
          new Promise(function (resolve) { setTimeout(function () { resolve(undefined); }, timeout); })
        ]);
      }
    };
  })();
  var accountAt = 0;
  // sk-hub.jp のログインの状態を写す（変わったら、アイコンとアカウントの画面を作り直す）。verify: ログインが切れていないかも確かめる
  function syncAccount(verify) {
    if (!bridge) return Promise.resolve(false);
    accountAt = Date.now();
    return bridge.ask("prefs", 5000).then(function (p) {
      if (!p) return false;
      var before = localStorage.getItem("skhub_account") || null, after = p.skhub_account || null;
      if (after && verify) return bridge.ask("token", 15000).then(function (tk) { if (tk === null) after = null; return apply(before, after); });
      return apply(before, after);
    }).catch(function () { return false; });
    function apply(before, after) {
      if (before === after) return false;
      try { if (after) localStorage.setItem("skhub_account", after); else localStorage.removeItem("skhub_account"); } catch (e) { /* 写せなくても、このページでは使える */ }
      fillAvatars();
      if ($("home") && !$("home").hidden && (view === "account" || view === "home")) renderStart();
      return true;
    }
  }
  // ログインのページから戻ってきたとき・ほかのウィンドウでログインしたあと（1 分に 1 回まで）
  document.addEventListener("visibilitychange", function () { if (!document.hidden && Date.now() - accountAt > 60000) syncAccount(); });
  // ログインのページ（sk-hub.jp/account/）。終わったら、いま開いている Office のページに戻る（sk-hub.jp/office/… は office.sk-hub.jp に移る）
  function loginUrl() {
    var path = (BASE === "/" ? "/office" : "") + location.pathname + location.hash;
    return SITE + "/account/?next=" + encodeURIComponent(path);
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
    if (d.getFullYear() === now.getFullYear()) return t("{m}月{d}日 {time}", { m: d.getMonth() + 1, d: d.getDate(), time: hm });
    return t("{y}年{m}月{d}日", { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() });
  }
  function bytes(n) { return n < 1024 ? n + " B" : n < 1048576 ? (n / 1024).toFixed(1) + " KB" : (n / 1048576).toFixed(1) + " MB"; }
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

  // ================================================================
  //   メニュー（ボタンの下・右クリックしたところに出る）と色えらび
  // ================================================================
  //   menu(場所, items) → 選んだ value（やめたら null）。場所はボタンの要素か { x, y }
  //   items: [{ value, label, icon, sub（右に出る文字。ショートカットなど）, checked, disabled, color（色の見本）} か { sep: 1 } か { head: "見出し" }]
  var menuClose = null;
  function closeMenu() { if (menuClose) menuClose(null); }
  function placeBox(box, at) {
    var x, y, r;
    if (at && at.getBoundingClientRect) { r = at.getBoundingClientRect(); x = r.left; y = r.bottom + 4; }
    else { x = at.x; y = at.y; }
    box.style.left = "0px"; box.style.top = "0px";
    var w = box.offsetWidth, h = box.offsetHeight, vw = window.innerWidth, vh = window.innerHeight;
    if (x + w > vw - 8) x = Math.max(8, vw - w - 8);
    if (y + h > vh - 8) y = r ? Math.max(8, r.top - h - 4) : Math.max(8, vh - h - 8);
    box.style.left = x + "px"; box.style.top = y + "px";
  }
  function popup(at, build) {
    closeMenu();
    return new Promise(function (resolve) {
      var box = document.createElement("div");
      box.className = "menu";
      box.setAttribute("role", "menu");
      document.body.appendChild(box);
      var before = document.activeElement;
      function done(v) {
        if (!box.isConnected) return;
        box.remove(); menuClose = null;
        document.removeEventListener("pointerdown", outside, true);
        document.removeEventListener("keydown", key, true);
        window.removeEventListener("blur", blur);
        if (before && before.focus && document.contains(before)) before.focus({ preventScroll: true });
        resolve(v);
      }
      function outside(e) { if (!box.contains(e.target)) done(null); }
      function blur() { done(null); }
      function key(e) {
        var items = Array.prototype.slice.call(box.querySelectorAll("[data-v]:not([disabled])")), i = items.indexOf(document.activeElement);
        if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); done(null); }
        else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
          e.preventDefault(); e.stopPropagation();
          if (items.length) items[(i + (e.key === "ArrowDown" ? 1 : items.length - 1) + (i < 0 && e.key === "ArrowUp" ? 1 : 0)) % items.length].focus();
        } else if (e.key === "Tab") { e.preventDefault(); done(null); }
      }
      build(box, done);
      // 押したところの選択（文書の中で選んだ文字など）が外れないように
      box.addEventListener("mousedown", function (e) { if (!e.target.closest("input")) e.preventDefault(); });
      placeBox(box, at);
      menuClose = done;
      setTimeout(function () {
        document.addEventListener("pointerdown", outside, true);
        window.addEventListener("blur", blur);
      }, 0);
      document.addEventListener("keydown", key, true);
      var first = box.querySelector('[data-v][aria-checked="true"]') || box.querySelector("[data-v]:not([disabled])");
      if (first && !(at && at.x !== undefined && !at.getBoundingClientRect && at.noFocus)) first.focus({ preventScroll: true });
    });
  }
  function menu(at, items) {
    return popup(at, function (box, done) {
      box.innerHTML = items.map(function (it, i) {
        if (it.sep) return '<div class="menu__sep" role="separator"></div>';
        if (it.head) return '<div class="menu__head">' + esc(it.head) + "</div>";
        return '<button type="button" class="menu__item m3-state" role="menuitem' + (it.checked !== undefined ? "checkbox" : "") + '" data-v="' + i + '"' +
          (it.checked !== undefined ? ' aria-checked="' + !!it.checked + '"' : "") + (it.disabled ? " disabled" : "") + ">" +
          (it.color ? '<span class="menu__swatch" style="background:' + esc(it.color) + '"></span>' : it.icon ? icon(it.icon) : it.checked !== undefined ? icon(it.checked ? "check" : "") : '<span class="menu__pad"></span>') +
          '<span class="menu__label"' + (it.style ? ' style="' + esc(it.style) + '"' : "") + ">" + esc(it.label) + "</span>" +
          (it.sub ? '<span class="menu__sub">' + esc(it.sub) + "</span>" : "") + "</button>";
      }).join("");
      box.addEventListener("click", function (e) {
        var b = e.target.closest("[data-v]");
        if (b && !b.disabled) done(items[+b.getAttribute("data-v")].value);
      });
    });
  }
  // 色えらび。palette(場所, { none: "なし" の文字（出さないときは省く）, value }) → "#rrggbb"・"none"・null（やめた）
  var SWATCH = [
    ["#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#d9d9d9", "#efefef", "#ffffff"],
    ["#ba1a1a", "#ea580c", "#f59e0b", "#16a34a", "#0d9488", "#2563eb", "#7c3aed", "#db2777"],
    ["#fecaca", "#fed7aa", "#fde68a", "#bbf7d0", "#99f6e4", "#bfdbfe", "#ddd6fe", "#fbcfe8"],
    ["#7f1d1d", "#9a3412", "#92400e", "#14532d", "#134e4a", "#1e3a8a", "#4c1d95", "#831843"]
  ];
  function palette(at, opts) {
    opts = opts || {};
    return popup(at, function (box, done) {
      box.classList.add("menu--palette");
      box.innerHTML = (opts.none ? '<button type="button" class="menu__item m3-state" data-v="none">' + icon("format_color_reset") + '<span class="menu__label">' + esc(opts.none) + "</span></button>" : "") +
        '<div class="palette">' + SWATCH.map(function (row) {
          return row.map(function (c) { return '<button type="button" class="palette__c" data-v="' + c + '" style="background:' + c + '" title="' + c + '" aria-label="' + c + '"' + (opts.value === c ? ' aria-checked="true"' : "") + "></button>"; }).join("");
        }).join("") + "</div>" +
        '<label class="menu__item palette__more">' + icon("palette") + '<span class="menu__label">' + esc(t("その他の色")) + '</span><input type="color" value="' + esc(/^#[0-9a-f]{6}$/i.test(opts.value || "") ? opts.value : "#2563eb") + '"></label>';
      box.addEventListener("click", function (e) { var b = e.target.closest("[data-v]"); if (b) done(b.getAttribute("data-v")); });
      box.querySelector('input[type="color"]').addEventListener("change", function (e) { done(e.target.value); });
    });
  }
  // 画像を選んで data: URL にする（ファイルの中に一緒に保存する）。大きな写真は縮める
  function pickFile(accept) {
    return new Promise(function (resolve) {
      var input = document.createElement("input");
      input.type = "file"; input.accept = accept || "";
      input.addEventListener("change", function () { resolve(input.files[0] || null); });
      input.click();
    });
  }
  function imageData(f) {
    return new Promise(function (resolve) {
      if (!f || !/^image\//.test(f.type)) { if (f) notify(t("画像のファイルを選んでください")); return resolve(null); }
      var r = new FileReader();
      r.onload = function () {
        var src = r.result;
        if (f.type === "image/svg+xml" || f.type === "image/gif" || f.size < 400000) return resolve(src);
        // 大きい写真は、長い辺を 1600px にして JPEG（透明があるときは PNG）にする
        var img = new Image();
        img.onload = function () {
          var k = Math.min(1, 1600 / Math.max(img.width, img.height)), cv = document.createElement("canvas");
          cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
          cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
          var out = cv.toDataURL(f.type === "image/png" ? "image/png" : "image/jpeg", 0.85);
          resolve(out.length < src.length ? out : src);
        };
        img.onerror = function () { resolve(src); };
        img.src = src;
      };
      r.readAsDataURL(f);
    });
  }
  function pickImage() { return pickFile("image/*").then(imageData); }
  // ファイルとして書き出す（Chrome・Edge は保存する場所を選ぶ。それ以外はダウンロード）
  function saveBlob(name, blob, desc, mime, ext) {
    if (typeof window.showSaveFilePicker === "function" && ext) {
      var accept = {}; accept[mime || blob.type || "application/octet-stream"] = [ext];
      return window.showSaveFilePicker({ suggestedName: name, types: [{ description: desc || name, accept: accept }] }).then(function (h) {
        return h.createWritable().then(function (w) { return w.write(blob).then(function () { return w.close(); }); }).then(function () { return true; });
      }, function (e) { return e && e.name === "AbortError" ? false : download(); });
    }
    return Promise.resolve(download());
    function download() {
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
      return true;
    }
  }

  // ================================================================
  //   アプリに渡す道具（O）。アプリのファイル（doc.js など）は、これだけを使う
  // ================================================================
  var O = {
    APP: APP, t: t, esc: esc, icon: icon, $: $, copy: copy, notify: notify, when: when,
    file: function () { return current; },
    data: function () { return current ? current.data : null; },
    // 中身を変えたら呼ぶ（保存して、元に戻すために覚える）。typing: 文字を打っているあいだ（まとめて 1 回にする）
    changed: function (typing) { changed(typing); },
    record: function () { record(); },
    stat: function (html) { var el = $("stat"); if (el) el.innerHTML = html; },
    sync: function () { syncRibbon(); },
    menu: menu, palette: palette, closeMenu: closeMenu,
    pickImage: pickImage, pickFile: pickFile, imageData: imageData, saveBlob: saveBlob,
    zoom: function () { return zoom / 100; },
    readMode: function () { return readMode; },
    userName: userName,
    base: BASE,
    fileName: function (ext) { return (current ? current.name : "file").replace(/[\\/:*?"<>|]/g, "_") + ext; }
  };
  var CLIP = { label: t("クリップボード"), items: [
    { big: 1, icon: "content_paste", label: t("貼り付け"), cmd: "paste" },
    { stack: [{ icon: "content_cut", label: t("切り取り"), cmd: "cut" }, { icon: "content_copy", label: t("コピー"), cmd: "copy" }, { icon: "format_paint", label: t("書式のコピー"), cmd: "painter", toggle: "painter" }] }
  ] };
  var ZOOM = { label: t("ズーム"), items: [
    { big: 1, icon: "search", label: t("ズーム"), cmd: "zoom" },
    { stack: [{ icon: "zoom_in", label: t("拡大"), cmd: "zoomIn" }, { icon: "zoom_out", label: t("縮小"), cmd: "zoomOut" }, { icon: "fit_screen", label: "100%", cmd: "zoom100" }] }
  ] };
  var VIEW = { label: t("表示"), items: [
    { big: 1, icon: "print", label: t("印刷"), cmd: "print" },
    { stack: [{ icon: "menu_book", label: t("閲覧モード"), cmd: "read", toggle: "read" }, { icon: "keyboard", label: t("ショートカット"), cmd: "keys" }] }
  ] };
  O.groups = { clip: CLIP, zoom: ZOOM, view: VIEW };
  // フォントの一覧（[表示する名前, CSS の font-family]）。PC に入っていないフォントは、近いもので出る
  O.fonts = [
    ["M PLUS 1", '"M PLUS 1", sans-serif'], ["游ゴシック", '"Yu Gothic", YuGothic, "Hiragino Sans", sans-serif'], ["メイリオ", 'Meiryo, "Hiragino Sans", sans-serif'],
    ["BIZ UDPゴシック", '"BIZ UDPGothic", "Yu Gothic", sans-serif'], ["游明朝", '"Yu Mincho", YuMincho, "Hiragino Mincho ProN", serif'], ["BIZ UD明朝", '"BIZ UDMincho", "Yu Mincho", serif'],
    ["Arial", "Arial, Helvetica, sans-serif"], ["Times New Roman", '"Times New Roman", Times, serif'], ["Georgia", "Georgia, serif"], ["Consolas", "Consolas, Menlo, monospace"]
  ];
  O.fontCss = function (name) {
    for (var i = 0; i < O.fonts.length; i++) if (O.fonts[i][0] === name) return O.fonts[i][1];
    return '"' + String(name).replace(/["\\;{}<>]/g, "") + '", sans-serif';
  };
  O.sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];
  // ライト・ダークのどちらで見ているか（グラフの色に使う）
  O.dark = function () { var th = root.getAttribute("data-theme"); return th ? th === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches; };
  var factory = window.OfficeApps && window.OfficeApps[APP];
  var App = factory ? factory(O) : null;

  // ---------- 元に戻す・やり直し（ファイルの中身をまるごと覚える） ----------
  var hist = { stack: [], i: -1, timer: null, size: 0 };
  function record() {
    clearTimeout(hist.timer); hist.timer = null;
    if (!current) return;
    var s = JSON.stringify(current.data);
    if (hist.stack[hist.i] === s) return;
    hist.stack = hist.stack.slice(0, hist.i + 1);
    hist.stack.push(s);
    // 100 回まで・合わせて 60MB まで（画像の多いファイルで重くならないように）
    hist.size = hist.stack.reduce(function (a, x) { return a + x.length; }, 0);
    while (hist.stack.length > 2 && (hist.stack.length > 100 || hist.size > 60000000)) hist.size -= hist.stack.shift().length;
    hist.i = hist.stack.length - 1;
  }
  function resetHist() { clearTimeout(hist.timer); hist.timer = null; hist.stack = current ? [JSON.stringify(current.data)] : []; hist.i = hist.stack.length - 1; }
  function changed(typing) {
    if (!current) return;
    if (typing) { clearTimeout(hist.timer); hist.timer = setTimeout(record, 700); }
    else record();
    saveSoon();
  }
  function restore(dir) {
    if (hist.timer) record();
    var n = hist.i + dir;
    if (n < 0 || n >= hist.stack.length) { notify(dir < 0 ? t("これ以上、元に戻せません") : t("これ以上、やり直せません")); return; }
    hist.i = n;
    current.data = JSON.parse(hist.stack[n]);
    App.render({ restore: true });
    setZoom(zoom);
    syncRibbon();
    saveSoon();
  }

  // ---------- 小さなプレビュー（テンプレート・最近のファイル） ----------
  function thumb(data, blank) { return App && App.thumb ? App.thumb(data, blank) : ""; }

  // ================================================================
  //   スタート画面（Office の「ホーム・新規・開く・情報・エクスポート・アカウント・オプション」）
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
    if (!current && (view === "info" || view === "export")) view = "home";
    var rail = '<nav class="rail" aria-label="' + esc(t("メニュー")) + '">' +
      (current ? railBtn("data-back-edit", "arrow_back", t("戻る")) : "") +
      '<a class="rail__brand" href="' + BASE + '" title="SK Hub Office"><span class="rail__logo">' + icon(a ? a.icon : "apps", true) + "</span><span>" + (a ? a.name : "Office") + "</span></a>" +
      railItem("home", "home", t("ホーム")) +
      (a ? railItem("new", "note_add", t("新規")) + railItem("open", "folder_open", t("開く")) : "") +
      (current ? railItem("info", "info", t("情報")) + railBtn("data-save-as", "save_as", t("コピーを保存")) + railItem("export", "ios_share", t("エクスポート")) +
        railBtn("data-print", "print", t("印刷")) + railBtn("data-close-file", "close", t("閉じる")) : "") +
      '<span class="rail__space"></span>' +
      (a ? '<a class="rail__item" href="' + BASE + '"><span class="rail__pill">' + icon("apps") + "</span><span>" + esc(t("アプリ")) + "</span></a>"
         : '<a class="rail__item" href="' + SITE + '/products/office/"><span class="rail__pill">' + icon("info") + "</span><span>" + esc(t("紹介")) + "</span></a>") +
      (installEvt && !standalone ? railBtn("data-install", "install_desktop", t("インストール")) : "") +
      railItem("account", "account_circle", t("アカウント")) +
      railItem("options", "settings", t("オプション")) +
      "</nav>";
    var main = view === "new" ? viewNew() : view === "open" ? viewOpen() : view === "account" ? viewAccount() : view === "options" ? viewOptions()
      : view === "info" ? viewInfo() : view === "export" ? viewExport() : viewHome();
    // アプリとして入れたときだけ見える、ウィンドウのいちばん上の帯（is-wco）
    var bar = '<header class="wco-bar"><span class="wco-bar__logo">' + icon(a ? a.icon : "apps", true) + '</span><span class="wco-bar__name">' + (a ? a.name : "SK Hub Office") + "</span>" +
      (a ? '<span class="wco-bar__sub">SK Hub Office</span>' : "") +
      '<span class="wco-bar__space"></span><button type="button" class="wco-search m3-state" data-cmd="find">' + icon("search") + "<span>" + esc(t("検索")) + "</span></button>" +
      '<span class="wco-bar__space"></span><button type="button" class="tbar-avatar" data-view="account" data-avatar aria-label="' + esc(t("アカウント")) + '" title="' + esc(t("アカウント")) + '"></button></header>';
    $("home").innerHTML = bar + rail + '<main class="start__main" id="start-main">' + main + "</main>";
    if (view === "new") filterTemplates("all");
    if (view === "home" || view === "open") renderFileTable();
    if (view === "info") renderVersions();
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
  function importLabel() {
    var x = App && App.imports ? App.imports.label : "";
    return x ? t(".skdoc・.sksheet・.skslide と {more} を開く", { more: x }) : t(".skdoc・.sksheet・.skslide を開く");
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
          '<button type="button" class="place m3-state" role="listitem" data-import>' + icon("upload_file") + "<span><strong>" + esc(t("参照")) + "</strong><small>" + esc(importLabel()) + "</small></span></button>" +
          '<button type="button" class="place m3-state" role="listitem" data-todo="' + esc(t("SK Hub Systems アカウント")) + '">' + icon("cloud") + "<span><strong>" + esc(t("SK Hub Systems アカウント")) + "</strong><small>" + esc(t("インターネット上への保存は準備中です")) + "</small></span></button>" +
        '</div><div class="open__main"><div id="legacy"></div>' + fileSection(false) + "</div>" +
      "</div>";
  }
  // 情報（ファイルのこと・版の履歴）
  function viewInfo() {
    var f = current, a = APPS[APP];
    return '<h1 class="start__h1">' + esc(t("情報")) + "</h1>" +
      '<section class="info-card"><span class="info-card__ico frow__ico" data-app="' + APP + '">' + icon(a.icon, true) + '</span><div class="info-card__text"><strong data-i18n-skip>' + esc(f.name) + "</strong><small>" +
        esc(f.outside ? t("フォルダの外のファイル（直接保存しています）") : Store.state() === "ready" ? t("フォルダ「{name}」", { name: Store.folderName() }) : t("このブラウザ")) + "</small></div>" +
        '<button type="button" class="text-btn m3-state" data-rename>' + esc(t("名前を変更")) + "</button></section>" +
      '<dl class="props"><dt>' + esc(t("更新日時")) + "</dt><dd>" + esc(when(f.updated || Date.now())) + "</dd><dt>" + esc(t("作成日時")) + "</dt><dd>" + esc(when(f.created || Date.now())) + "</dd>" +
        "<dt>" + esc(t("大きさ")) + "</dt><dd>" + esc(bytes(JSON.stringify(f.data).length)) + "</dd>" + (App.props ? App.props().map(function (p) { return "<dt>" + esc(p[0]) + "</dt><dd>" + esc(p[1]) + "</dd>"; }).join("") : "") + "</dl>" +
      '<section class="start__sec"><div class="files-head"><h2 class="start__h2">' + esc(t("版の履歴")) + '</h2><button type="button" class="m3-btn m3-btn--tonal m3-state" data-keep-version>' + icon("bookmark_add") + esc(t("いまの版を残す")) + "</button></div>" +
        '<p class="sec-text">' + esc(t("保存したとき・開いたとき・10 分ごとに、前の版をこのブラウザに残します（20 個まで）。押すと、その版に戻せます。")) + "</p>" +
        '<div class="ftable" id="versions"><div class="files-loading">' + esc(t("読み込んでいます…")) + "</div></div></section>";
  }
  function renderVersions() {
    var box = $("versions");
    if (!box || !current) return;
    Store.versions(current.id).then(function (list) {
      if (!$("versions")) return;
      if (!list.length) { box.innerHTML = '<div class="files-loading">' + esc(t("まだ前の版はありません。")) + "</div>"; return; }
      box.innerHTML = list.map(function (v, i) {
        return '<div class="frow" data-app="' + APP + '"><button type="button" class="frow__open m3-state" data-restore="' + i + '"><span class="frow__ico">' + icon("history") + '</span><span class="frow__text"><span class="frow__name">' + esc(when(v.ts)) + '</span><span class="frow__meta">' + esc(bytes(v.size)) + "</span></span></button>" +
          '<span class="ftable__when"></span><span class="frow__act"><button type="button" class="text-btn m3-state" data-restore="' + i + '">' + esc(t("この版に戻す")) + "</button></span></div>";
      }).join("");
    });
  }
  function restoreVersion(i) {
    Store.versions(current.id).then(function (list) {
      var v = list[i];
      if (!v) return;
      return M3.confirm({ title: t("この版に戻しますか？"), text: t("{when} の版に戻します。いまの中身も、版の履歴に残します。", { when: when(v.ts) }), ok: t("戻す"), icon: "history" }).then(function (ok) {
        if (!ok) return;
        return Store.addVersion(current.id, current.data).then(function () {
          current.data = JSON.parse(v.data);
          if (App.upgrade) App.upgrade(current.data);
          record();
          backToEditor();
          App.render({ restore: true });
          setZoom(zoom);
          saveSoon();
          notify(t("前の版に戻しました"));
        });
      });
    });
  }
  // エクスポート（ほかの形式で書き出す）
  function viewExport() {
    var list = [{ id: "sk", icon: "save_as", label: t("SK Hub Office のファイル"), desc: t("{ext} のファイルとして、好きな場所に保存します", { ext: Store.EXT[APP] }) },
      { id: "pdf", icon: "picture_as_pdf", label: "PDF", desc: t("印刷の画面で「PDF に保存」を選びます") }].concat(App.exports || []);
    return '<h1 class="start__h1">' + esc(t("エクスポート")) + "</h1>" +
      '<div class="export-grid">' + list.map(function (x) {
        return '<button type="button" class="export-card m3-state" data-export="' + esc(x.id) + '">' + icon(x.icon) + "<span><strong>" + esc(x.label) + "</strong><small>" + esc(x.desc || "") + "</small></span></button>";
      }).join("") + "</div>" +
      '<p class="start__note">' + esc(t("Office の形式（.docx・.xlsx・.pptx）は、文字・表・図の主なものを書き出します。細かい書式は変わることがあります。")) + "</p>";
  }
  function exportAs(id) {
    if (!current) return;
    flush();
    if (id === "sk") { Store.exportFile(current).then(function (ok) { if (ok) notify(t("書き出しました")); }).catch(failed); return; }
    if (id === "pdf") { backToEditor(); setTimeout(doPrint, 300); return; }
    Promise.resolve(App.exportAs(id)).then(function (ok) { if (ok) notify(t("書き出しました")); }).catch(function (e) {
      if (window.console) console.error(e);
      notify(t("書き出せませんでした"));
    });
  }
  // アカウント（Microsoft Office の「アカウント」と同じ並び。左: ユーザー情報・プライバシー・テーマ・接続済みサービス、右: 製品情報と大きなボタン）
  function viewAccount() {
    var hint = accountHint(), name = userName(), s = Store.state();
    // アプリとして入れたウィンドウからは、ブラウザでログインのページを開く（戻ってきたら、ログインの状態を読みなおす）
    var target = standalone ? ' target="_blank" rel="noopener"' : "";
    var blank = ' target="_blank" rel="noopener"';
    if (bridge && Date.now() - accountAt > 5000) setTimeout(function () { syncAccount(true); }, 0);
    var themes = [["auto", t("端末に合わせる")], ["light", t("ライト")], ["dark", t("ダーク")]];
    var services = [s === "unsupported" ? ["public", t("このブラウザ"), t("ブラウザの中に保存したファイル")]
      : ["computer", t("この PC"), s === "none" ? t("フォルダを選んでいません") : t("フォルダ「{name}」", { name: Store.folderName() })]];
    if (hint) services.push(["cloud", t("SK Hub Systems アカウント"), hint.email || hint.name || ""]);
    var canInstall = installEvt && !standalone;
    // 右の大きなボタン: [アイコン, ボタンの文字, 見出し, 説明, ボタンの属性（なければ押せない）]
    var tiles = [
      ["system_update_alt", t("今すぐ更新"), t("SK Hub Office の更新"), t("ネットにつながっていれば、開いたときに新しい版が自動で使われます。") + (current ? t("今すぐ更新するときは、ファイルを閉じてから押してください。") : ""), current ? "" : "data-update"],
      ["install_desktop", t("インストール"), t("アプリとしてインストール"), installText(), canInstall ? "data-install" : ""],
      ["help", t("サポート"), t("サポート"), t("使い方・よくある質問・お問い合わせ。"), 'href="https://support.sk-hub.jp/"' + blank],
      ["info", t("バージョン情報"), t("SK Hub Office のバージョン情報"), [t("SK Hub Office、ポリシー、サポートに関する詳しい情報。"), t("バージョン {v}（開発中）", { v: VERSION }), "SK Material（Material Design 3）"], 'href="' + SITE + '/products/office/"' + blank]
    ];
    return '<h1 class="start__h1">' + esc(t("アカウント")) + "</h1>" +
      '<div class="acct">' +
        '<div class="acct__col">' +
          '<section class="acct__sec"><h2 class="acct__h">' + esc(t("ユーザー情報")) + "</h2>" +
            '<div class="acct-user"><span class="acct-avatar" data-avatar aria-hidden="true"></span><div class="acct-user__text"><strong>' + esc(name || t("名前が設定されていません")) + "</strong>" +
              "<small>" + esc(hint && hint.email ? hint.email : t("SK Hub Systems アカウントにログインしていません")) + "</small></div></div>" +
            '<div class="acct-links">' + (hint
              // ログイン中に ?next= 付きで開くと、すぐ戻ってきてしまうので、管理のページはそのまま開く（戻ってきたら、ログインの状態を読みなおす）
              ? '<a href="' + SITE + '/account/"' + blank + ">" + esc(t("ログアウト")) + '</a><a href="' + SITE + '/account/"' + blank + ">" + esc(t("アカウントの管理")) + "</a>"
              : '<a href="' + esc(loginUrl()) + '"' + target + ">" + esc(t("ログイン")) + '</a><a href="' + SITE + '/account/"' + blank + ">" + esc(t("アカウントを作る")) + "</a>") +
              '<button type="button" data-view="options">' + esc(t("名前を変更")) + "</button></div></section>" +
          '<section class="acct__sec"><h2 class="acct__h">' + esc(t("アカウントのプライバシー")) + "</h2>" +
            '<a class="acct-btn m3-state" href="' + SITE + '/policies/"' + blank + ">" + esc(t("プライバシーポリシー")) + "</a></section>" +
          '<section class="acct__sec"><h2 class="acct__h"><label for="acct-theme">' + esc(t("Office テーマ")) + "</label></h2>" +
            '<span class="acct-select"><select id="acct-theme">' + themes.map(function (x) {
              return '<option value="' + x[0] + '"' + (settings.theme === x[0] ? " selected" : "") + ">" + esc(x[1]) + "</option>";
            }).join("") + "</select>" + icon("expand_more") + "</span></section>" +
          '<section class="acct__sec"><h2 class="acct__h">' + esc(t("接続済みサービス")) + "</h2>" +
            '<ul class="acct-services">' + services.map(function (x) {
              return "<li>" + icon(x[0]) + '<span><strong>' + esc(x[1]) + "</strong><small>" + esc(x[2]) + "</small></span></li>";
            }).join("") + "</ul>" +
            '<button type="button" class="acct-btn m3-state" data-todo="' + esc(t("インターネット上への保存")) + '">' + esc(t("サービスの追加")) + icon("expand_more") + "</button>" +
            '<p class="start__note">' + esc(t("ファイルは、ログインしてもしなくても、この PC に保存されます。インターネット上への保存（ほかの端末で開く・共有する）は準備中です。")) + "</p></section>" +
        "</div>" +
        '<div class="acct__col">' +
          '<section class="acct__sec"><h2 class="acct__h">' + esc(t("製品情報")) + "</h2>" +
            '<p class="acct-brand">SK Hub Office</p>' +
            '<p class="acct-product">' + esc(t("SK Hub Office バージョン {v}（開発中）", { v: VERSION })) + "</p>" +
            '<p class="acct-small">' + esc(t("この製品には以下が含まれます。")) + "</p>" +
            '<div class="acct-apps">' + ["doc", "sheet", "slide"].map(function (k) {
              return '<a class="acct-apps__ico" data-app="' + k + '" href="' + APPS[k].path + '" title="' + APPS[k].name + '" aria-label="' + APPS[k].name + '">' + icon(APPS[k].icon, true) + "</a>";
            }).join("") + "</div></section>" +
          '<div class="acct-tiles">' + tiles.map(function (x) {
            var tag = /^href/.test(x[4]) ? "a" : "button", off = !x[4];
            var desc = [].concat(x[3]).map(function (p) { return "<p>" + esc(p) + "</p>"; }).join("");
            return '<div class="acct-tile' + (off ? " is-off" : "") + '"><' + tag + (tag === "button" ? ' type="button"' : "") + ' class="acct-tile__btn m3-state" ' + (off ? "disabled" : x[4]) + ">" + icon(x[0]) + "<span>" + esc(x[1]) + "</span></" + tag + ">" +
              '<div class="acct-tile__text"><strong>' + esc(x[2]) + "</strong>" + desc + "</div></div>";
          }).join("") + "</div>" +
        "</div>" +
      "</div>";
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
        optRow("badge", t("ユーザー名"), t("あいさつ・アカウントのアイコン・コメントに出ます"), '<input class="opt-input" id="opt-user" maxlength="40" autocomplete="nickname" placeholder="' + esc(t("例：SK")) + '" value="' + esc(settings.user) + '" aria-label="' + esc(t("ユーザー名")) + '">') +
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
          '<button type="button" class="tab m3-state" role="tab" data-ftab="pinned" aria-selected="' + (fileTab === "pinned") + '"><span>' + esc(t("ピン留め")) + "</span></button>" +
          '<span class="tabs__ind" aria-hidden="true"></span></div>'
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
      thumb(copy(tp.data), blank) +
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
  function fileActions(id, name, at) {
    menu(at, [
      { value: "open", label: t("開く"), icon: "open_in_new" },
      { value: "rename", label: t("名前を変更"), icon: "edit" },
      { value: "copy", label: t("コピーを作る"), icon: "content_copy" },
      { value: "export", label: t("ファイルとして書き出す"), icon: "download" },
      { sep: 1 },
      { value: "delete", label: t("削除"), icon: "delete" }
    ]).then(function (v) {
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
          .then(function (ok) { if (!ok) return; return Store.remove(id).then(function () { setPinned(id, false); Store.dropVersions(id); renderFileTable(); notify(t("削除しました")); }); });
      }
    }).catch(failed);
  }
  function renameFile(f, done) {
    return M3.prompt({ title: t("名前を変更"), label: t("ファイルの名前"), value: f.name, ok: t("変更") }).then(function (v) {
      v = (v || "").trim();
      if (!v || v === f.name) return;
      var old = f.id;
      return Store.rename(f, v).then(function () {
        if (old !== f.id) { Store.moveVersions(old, f.id); if (isPinned(old)) { setPinned(old, false); setPinned(f.id, true); } }
        if (done) done();
      });
    }).catch(function (e) {
      if (e && e.message === "outside") notify(t("フォルダの外のファイルは、名前を変えられません。「コピーを保存」で、新しい名前で保存できます"));
      else failed(e);
    });
  }
  // ほかの場所のファイルを開く（SK Hub Office のファイル・アプリが読める Office のファイルなど）
  function importFile() {
    ensureFolder().then(function (ok) {
      if (!ok) return;
      var more = App && App.imports ? App.imports.accept : "";
      return pickFile(".skdoc,.sksheet,.skslide" + (more ? "," + more : "")).then(function (file) {
        if (!file) return;
        if (Store.appOf(file.name)) {
          return Store.importFile(file).then(function (f) { notify(t("「{name}」を読み込みました", { name: f.name })); openFile(f.id, f.app); },
            function () { notify(t("SK Hub Office のファイルではないので、開けませんでした")); });
        }
        if (!more) { notify(t("このアプリでは開けないファイルです")); return; }
        notify(t("読み込んでいます…"));
        return Promise.resolve(App.importFile(file)).then(function (data) {
          var name = file.name.replace(/\.[^.]+$/, "") || APPS[APP].untitled;
          return Store.create(APP, name, "import", data).then(function (f) { notify(t("「{name}」を読み込みました", { name: f.name })); openFile(f.id); });
        }, function (e) {
          if (window.console) console.error(e);
          notify(e && e.message === "no-inflate" ? t("このブラウザでは、このファイルを読み込めません") : t("読み込めませんでした。ファイルがこわれているかもしれません"));
        });
      });
    }).catch(failed);
  }

  // ================================================================
  //   エディタ（タイトルバー・リボン・ステータスバー）
  // ================================================================
  // リボン（アプリの App.ribbon）: [{ tab, groups: [{ label, items }] }]
  //   items の 1 つ: { big, icon, label } … 大きなボタン / { stack: [{ icon, label }] } … 文字つきの小さなボタンを縦に
  //                  { rows: [[{ icon, label } か { chip, w, key }]] } … アイコンだけのボタンの行 / { styles: [{ tag, label }] } … スタイルの一覧（Doc）
  //   cmd があれば動く（App.command(cmd, 押したボタン)）。todo: 1 は「まだ使えません」。drop: 1 は下向きの矢印（メニューが出る）
  //   toggle: "名前" は、App.state() の その名前が true のときにオンに見せる。chip の key: "名前" は、App.state() の値を文字として出す
  var ribbonTab = 0, ribbonOpen = true, zoom = 100;

  function btnAttrs(x) {
    return x.todo ? ' data-todo="' + esc(x.label || x.chip) + '"' : ' data-cmd="' + esc(x.cmd) + '"' + (x.toggle ? ' data-toggle="' + x.toggle + '" aria-pressed="false"' : "") + (x.key ? ' data-chip="' + x.key + '"' : "");
  }
  function ribbonItem(x) {
    var drop = x.drop ? '<span class="msr rb-drop" aria-hidden="true">arrow_drop_down</span>' : "";
    if (x.big) return '<button type="button" class="rb-big m3-state"' + btnAttrs(x) + ">" + icon(x.icon) + "<span>" + esc(x.label) + drop + "</span></button>";
    if (x.stack) return '<div class="rb-stack">' + x.stack.map(function (s) { return '<button type="button" class="rb-small m3-state"' + btnAttrs(s) + ">" + icon(s.icon) + "<span>" + esc(s.label) + "</span>" + (s.drop ? '<span class="msr rb-drop" aria-hidden="true">arrow_drop_down</span>' : "") + "</button>"; }).join("") + "</div>";
    if (x.rows) return '<div class="rb-rows">' + x.rows.map(function (row) {
      return '<div class="rb-row">' + row.map(function (b) {
        if (b.chip) return '<button type="button" class="rb-chip m3-state" style="width:' + b.w + 'px" title="' + esc(b.label || b.chip) + '"' + btnAttrs(b) + "><span>" + esc(b.chip) + "</span>" + icon("arrow_drop_down") + "</button>";
        return '<button type="button" class="rb-icon m3-state' + (b.drop ? " rb-icon--drop" : "") + '"' + btnAttrs(b) + ' aria-label="' + esc(b.label) + '" title="' + esc(b.label + (b.sub ? "（" + b.sub + "）" : "")) + '">' + icon(b.icon) + (b.drop ? '<span class="msr rb-drop" aria-hidden="true">arrow_drop_down</span>' : "") + "</button>";
      }).join("") + "</div>";
    }).join("") + "</div>";
    if (x.styles) return '<div class="rb-styles">' + x.styles.map(function (s) {
      return '<button type="button" class="rb-style rb-style--' + s.tag + ' m3-state" data-cmd="fmt:' + s.tag + '" data-toggle="fmt:' + s.tag + '"><span class="rb-style__sample">' + esc(t("あア亜Aa")) + "</span><span>" + esc(s.label) + "</span></button>";
    }).join("") + "</div>";
    return "";
  }
  // prevTab: 切りかえる前のタブ（中身を右・左のどちらから入れるかを決める）
  function ribbonTabs() { return App.ribbon.filter(function (tb) { return !tb.when || tb.when(); }); }
  function renderRibbon(prevTab) {
    var tabs = ribbonTabs(), row = $("rtabs");
    if (ribbonTab >= tabs.length) ribbonTab = 0;
    // タブの行は、タブが同じあいだは作り直さない（作り直すと、押したボタンが消えて下線もずれる）
    var sig = tabs.map(function (tb) { return tb.tab; }).join("|");
    if (row.getAttribute("data-sig") !== sig) {
      row.setAttribute("data-sig", sig);
      row.innerHTML = '<button type="button" class="rtab rtab--file m3-state" data-file>' + icon("folder", true) + esc(t("ファイル")) + "</button>" +
        tabs.map(function (tb, i) { return '<button type="button" class="rtab m3-state' + (tb.ctx ? " rtab--ctx" : "") + '" role="tab" aria-selected="false" data-rtab="' + i + '"><span>' + esc(tb.tab) + "</span></button>"; }).join("") +
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
    syncRibbon();
    placeIndicator();
  }
  // タブの下線（ind）を、tab の下へ滑らせる。下線は 1 本だけで、CSS の transition（left・right）で動く。
  // 両端は同時に動く（伸び縮みせず、幅だけ次のタブに合わせて変わる）。
  // animate: false のときは、すぐその場所に置く（開いたとき・画面の大きさが変わったとき）
  function moveLine(ind, tab, box, animate) {
    var r = tab.getBoundingClientRect(), b = box.getBoundingClientRect();
    if (!r.width) return;
    var left = r.left - b.left + box.scrollLeft;
    if (animate === false || isNaN(parseFloat(ind.style.left))) ind.style.transition = "none";
    ind.style.left = left + "px";
    ind.style.right = (box.clientWidth - left - r.width) + "px";
    if (ind.style.transition === "none") { void ind.offsetWidth; ind.style.transition = ""; }
  }
  function placeIndicator(animate) {
    var ind = $("rtabs-ind"), tab = document.querySelector('.rtab[aria-selected="true"] > span'), box = $("rtabs");
    if (ind && tab && box) moveLine(ind, tab, box, animate);
  }
  window.addEventListener("resize", function () { placeIndicator(false); });
  // フォント（M PLUS 1・アイコン）が読みこまれると、タブの幅が変わるので置き直す
  if (document.fonts) document.fonts.addEventListener("loadingdone", function () { placeIndicator(false); });
  // 選んでいるところに合わせて、リボンのボタンのオン・オフと、フォント名などの文字を変える。
  // 図形を選んだときの「図形の書式」のような、出たり消えたりするタブ（when）も、ここで出し入れする
  var syncTimer = null;
  function syncRibbon() {
    if (!App || !current || !$("rbody")) return;
    var sig = ribbonTabs().map(function (tb) { return tb.tab; }).join("|");
    if ($("rtabs").getAttribute("data-sig") !== sig) {
      var tabs = ribbonTabs(), ctx = -1;
      tabs.forEach(function (tb, i) { if (tb.ctx) ctx = i; });
      var was = App.ribbon.filter(function (tb) { return tb.ctx; }).some(function (tb) { return $("rtabs").getAttribute("data-sig").split("|").indexOf(tb.tab) >= 0; });
      if (ctx >= 0 && !was) ribbonTab = ctx;
      if (ribbonTab >= tabs.length) ribbonTab = 0;
      renderRibbon();
      return;
    }
    var st = App.state ? App.state() : {};
    st.read = readMode;
    st.painter = st.painter || false;
    document.querySelectorAll("#editor [data-toggle]").forEach(function (b) { b.setAttribute("aria-pressed", String(!!st[b.getAttribute("data-toggle")])); });
    document.querySelectorAll("#editor [data-chip]").forEach(function (b) {
      var v = st[b.getAttribute("data-chip")];
      if (v !== undefined && v !== null) b.firstChild.textContent = v;
    });
  }
  function syncSoon() { clearTimeout(syncTimer); syncTimer = setTimeout(syncRibbon, 60); }
  O.syncSoon = syncSoon;
  document.addEventListener("selectionchange", function () { if (current) syncSoon(); });

  function buildEditor() {
    var a = APPS[APP], ed = $("editor");
    if (!a || !ed || !App) return;
    ed.innerHTML =
      '<header class="titlebar">' +
        '<button type="button" class="tbar-logo" data-file aria-label="' + esc(t("ファイル")) + '" title="' + esc(t("ファイル")) + '">' + icon(a.icon, true) + "</button>" +
        '<div class="qat" role="toolbar" aria-label="' + esc(t("クイック アクセス")) + '">' +
          '<button type="button" class="icon-btn m3-state" data-cmd="save" aria-label="' + esc(t("保存")) + '" title="' + esc(t("保存（Ctrl+S）")) + '">' + icon("save") + "</button>" +
          '<button type="button" class="icon-btn m3-state" data-cmd="undo" aria-label="' + esc(t("元に戻す")) + '" title="' + esc(t("元に戻す（Ctrl+Z）")) + '">' + icon("undo") + "</button>" +
          '<button type="button" class="icon-btn m3-state" data-cmd="redo" aria-label="' + esc(t("やり直し")) + '" title="' + esc(t("やり直し（Ctrl+Y）")) + '">' + icon("redo") + "</button>" +
        "</div>" +
        '<button type="button" class="tbar-name m3-state" id="title-btn" title="' + esc(t("名前を変更")) + '"><span class="tbar-name__text" id="title-text" data-i18n-skip></span><span class="tbar-name__state" id="status"></span>' + icon("expand_more") + "</button>" +
        '<button type="button" class="tbar-search m3-state" data-cmd="find">' + icon("search") + "<span>" + esc(t("検索")) + "</span></button>" +
        '<span class="titlebar__space"></span>' +
        '<button type="button" class="icon-btn m3-state hide-sm" ' + (App.comments ? 'data-cmd="comments"' : 'data-todo="' + esc(t("コメント")) + '"') + ' aria-label="' + esc(t("コメント")) + '" title="' + esc(t("コメント")) + '">' + icon("comment") + "</button>" +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state tbar-share" data-todo="' + esc(t("共有")) + '">' + icon("share") + '<span class="hide-sm">' + esc(t("共有")) + "</span></button>" +
        '<button type="button" class="tbar-avatar" data-avatar-btn data-avatar aria-label="' + esc(t("アカウント")) + '" title="' + esc(t("アカウント")) + '"></button>' +
      "</header>" +
      '<div class="ribbon"><div class="rtabs" id="rtabs" role="tablist" aria-label="' + esc(t("リボン")) + '"></div><div class="rbody" id="rbody" role="toolbar"></div></div>' +
      '<div class="findbar" id="findbar" hidden>' + icon("search") +
        '<input id="find-q" type="search" placeholder="' + esc(t("検索")) + '" autocomplete="off" aria-label="' + esc(t("検索する文字")) + '">' +
        '<input id="find-r" type="search" placeholder="' + esc(t("置換後の文字")) + '" autocomplete="off" aria-label="' + esc(t("置換後の文字")) + '">' +
        '<label class="findbar__opt"><input type="checkbox" id="find-case">' + esc(t("大文字と小文字を区別")) + "</label>" +
        '<span class="findbar__info" id="find-info"></span>' +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-cmd="findPrev" aria-label="' + esc(t("前へ")) + '">' + icon("keyboard_arrow_up") + "</button>" +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-cmd="findNext">' + esc(t("次へ")) + "</button>" +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-cmd="replaceOne">' + esc(t("置換")) + "</button>" +
        '<button type="button" class="m3-btn m3-btn--tonal m3-state" data-cmd="replaceAll">' + esc(t("すべて置換")) + "</button>" +
        '<button type="button" class="icon-btn m3-state" data-cmd="findClose" aria-label="' + esc(t("閉じる")) + '">' + icon("close") + "</button></div>" +
      '<div class="canvas" id="canvas" data-i18n-skip></div>' +
      '<footer class="statusbar"><span class="statusbar__info" id="stat"></span>' +
        '<span class="statusbar__space"></span>' +
        '<span class="views hide-sm">' +
          '<button type="button" class="sb-btn m3-state" data-cmd="read" data-toggle="read" aria-label="' + esc(t("閲覧モード")) + '" title="' + esc(t("閲覧モード")) + '">' + icon("menu_book") + "</button>" +
          '<button type="button" class="sb-btn is-on m3-state" data-cmd="normalView" aria-label="' + esc(t("標準")) + '" title="' + esc(t("標準")) + '">' + icon(APP === "sheet" ? "grid_on" : APP === "slide" ? "view_sidebar" : "article") + "</button>" +
          (APP === "slide" ? '<button type="button" class="sb-btn m3-state" data-cmd="present:cur" aria-label="' + esc(t("スライド ショー")) + '" title="' + esc(t("スライド ショー（F5）")) + '">' + icon("slideshow") + "</button>" : "") +
        "</span>" +
        '<span class="zoom"><button type="button" class="sb-btn m3-state" data-cmd="zoomOut" aria-label="' + esc(t("縮小")) + '">' + icon("remove") + "</button>" +
        '<input type="range" id="zoom" min="25" max="300" step="5" value="100" aria-label="' + esc(t("ズーム")) + '">' +
        '<button type="button" class="sb-btn m3-state" data-cmd="zoomIn" aria-label="' + esc(t("拡大")) + '">' + icon("add") + "</button>" +
        '<button type="button" class="sb-zoom m3-state" data-cmd="zoom" id="zoom-v">100%</button></span>' +
      "</footer>";
    $("title-btn").addEventListener("click", renameCurrent);
    $("zoom").addEventListener("input", function (e) { setZoom(+e.target.value); });
    $("find-q").addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); findNext(e.shiftKey ? -1 : 1); } else if (e.key === "Escape") closeFind(); });
    $("find-r").addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); replaceOne(); } else if (e.key === "Escape") closeFind(); });
    $("find-q").addEventListener("input", function () { $("find-info").textContent = ""; });
    // 文書の中の選んだところが外れないよう、リボンとクイック アクセスを押したときにフォーカスを動かさない
    ed.addEventListener("mousedown", function (e) { if (e.target.closest(".ribbon button, .qat button, .statusbar button")) e.preventDefault(); });
    renderRibbon();
  }
  function setTitle() {
    $("title-text").textContent = current.name;
    document.title = current.name + " — " + APPS[APP].name + " | SK Hub Office";
  }
  function setStatus(s) { var el = $("status"); if (el) el.textContent = s ? " ・ " + s : ""; }
  // 保存（自動保存がオンなら、変えてから少しして保存。オフなら、保存のボタンか Ctrl+S で保存）
  var saveTimer = null, dirty = false, lastVersion = 0;
  function saveNow() {
    clearTimeout(saveTimer); saveTimer = null;
    if (!current) return Promise.resolve(false);
    var f = current;
    if (App && App.beforeSave) App.beforeSave();
    dirty = false;
    setStatus(t("保存しています…"));
    return Store.save(f).then(function () {
      if (current === f && !dirty) setStatus(t("保存済み"));
      // 10 分ごとに、版の履歴に残す
      if (Date.now() - lastVersion > 600000) { lastVersion = Date.now(); Store.addVersion(f.id, f.data); }
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
  function flush() { if (hist.timer) record(); if (saveTimer && current) saveNow(); }
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
        if ($("home") && !$("home").hidden && view === "info") renderStart();
      });
    });
  }
  function closeFile() {
    var f = current;
    (dirty ? saveNow() : Promise.resolve()).then(function () {
      if (f) Store.addVersion(f.id, f.data);
      if (App && App.close) App.close();
      current = null; dirty = false;
      history.pushState(null, "", location.pathname);
      route();
    });
  }

  function setZoom(z) {
    zoom = Math.max(25, Math.min(300, Math.round(z / 5) * 5));
    var r = $("zoom"); if (r) r.value = zoom;
    var v = $("zoom-v"); if (v) v.textContent = zoom + "%";
    if (App && App.zoom) App.zoom(zoom / 100);
  }

  // ---------- 検索と置換（中身はアプリの App.find・App.replace） ----------
  function openFind(replace) {
    var bar = $("findbar"); if (!bar) return;
    bar.hidden = false;
    var sel = App.selectedText ? App.selectedText() : "";
    if (sel && sel.length < 100 && sel.indexOf("\n") < 0) $("find-q").value = sel;
    var input = replace && $("find-q").value ? $("find-r") : $("find-q");
    input.focus(); input.select();
  }
  function closeFind() { var bar = $("findbar"); if (bar) bar.hidden = true; if (App.findClose) App.findClose(); }
  function findOpts() { return { q: ($("find-q").value || ""), r: $("find-r").value || "", matchCase: $("find-case").checked }; }
  function findNext(dir) {
    var o = findOpts();
    if (!o.q) return openFind();
    var res = App.find(o.q, dir || 1, o);
    $("find-info").textContent = res === false ? t("見つかりませんでした") : typeof res === "string" ? res : "";
  }
  function replaceOne() {
    var o = findOpts();
    if (!o.q) return openFind(true);
    var n = App.replace(o.q, o.r, false, o);
    if (n) changed();
    findNext(1);
  }
  function replaceAll() {
    var o = findOpts();
    if (!o.q) return openFind(true);
    var n = App.replace(o.q, o.r, true, o);
    if (n) changed();
    $("find-info").textContent = t("{n} か所を置きかえました", { n: n || 0 });
  }

  // ---------- 印刷（PDF もここから。アプリが App.beforePrint で印刷用のものを用意する） ----------
  var printStyle = null;
  function beforePrint() {
    if (!current || !App) return;
    flush();
    root.classList.add("is-printing");
    var css = App.beforePrint ? App.beforePrint() : "";
    if (css) { printStyle = document.createElement("style"); printStyle.textContent = css; document.head.appendChild(printStyle); }
  }
  function afterPrint() {
    root.classList.remove("is-printing");
    if (printStyle) { printStyle.remove(); printStyle = null; }
    if (App && App.afterPrint) App.afterPrint();
  }
  window.addEventListener("beforeprint", beforePrint);
  window.addEventListener("afterprint", afterPrint);
  function doPrint() { if ($("editor").hidden) backToEditor(); window.print(); }

  function showKeys() {
    var list = [["Ctrl+S", t("保存")], ["Ctrl+Z", t("元に戻す")], ["Ctrl+Y", t("やり直し")], ["Ctrl+C / X / V", t("コピー・切り取り・貼り付け")],
      ["Ctrl+F", t("検索")], ["Ctrl+H", t("置換")], ["Ctrl+P", t("印刷")]].concat(App && App.keys ? App.keys : []).concat([["Esc", t("閉じる")]]);
    var body = document.createElement("dl");
    body.className = "keys-list";
    body.innerHTML = list.map(function (x) { return "<dt>" + esc(x[0]) + "</dt><dd>" + esc(x[1]) + "</dd>"; }).join("");
    M3.dialog({ title: t("キーボードショートカット"), icon: "keyboard", body: body, actions: [{ label: t("閉じる"), value: null, primary: true }] });
  }

  function runCmd(cmd, el) {
    if (!current && cmd !== "keys" && cmd !== "find") return;
    if (cmd === "save") {
      saveNow().then(function (ok) {
        if (!ok) return;
        lastVersion = Date.now();
        Store.addVersion(current.id, current.data);
        notify(current.outside ? t("ファイルに保存しました") : Store.state() === "ready" ? t("フォルダ「{folder}」に保存しました", { folder: Store.folderName() }) : t("このブラウザに保存しました"));
      });
      return;
    }
    if (cmd === "undo") { if (App.beforeUndo) App.beforeUndo(); return restore(-1); }
    if (cmd === "redo") { if (App.beforeUndo) App.beforeUndo(); return restore(1); }
    if (cmd === "copy" || cmd === "cut" || cmd === "paste") { if (App.clip) App.clip(cmd); return; }
    if (cmd === "find") {
      if (!current || ($("editor") && $("editor").hidden)) { var fs = $("file-search"); if (fs) fs.focus(); return; }
      return openFind();
    }
    if (cmd === "replace") return openFind(true);
    if (cmd === "findNext") return findNext(1);
    if (cmd === "findPrev") return findNext(-1);
    if (cmd === "findClose") return closeFind();
    if (cmd === "replaceOne") return replaceOne();
    if (cmd === "replaceAll") return replaceAll();
    if (cmd === "print") return doPrint();
    if (cmd === "keys") return showKeys();
    if (cmd === "read") {
      readMode = !readMode;
      root.classList.toggle("is-read", readMode);
      if (App.setRead) App.setRead(readMode);
      syncRibbon();
      notify(readMode ? t("閲覧モードにしました") : t("編集モードにしました"));
      return;
    }
    if (cmd === "normalView") { if (readMode) runCmd("read"); return; }
    if (cmd === "zoomIn") return setZoom(zoom + 10);
    if (cmd === "zoomOut") return setZoom(zoom - 10);
    if (cmd === "zoom100") return setZoom(100);
    if (cmd === "zoom") {
      var opts = ["300", "200", "150", "125", "100", "75", "50"].map(function (z) { return { value: z, label: z + "%", checked: +z === zoom }; });
      if (App.fitZoom) opts.push({ sep: 1 }, { value: "fit", label: t("ウィンドウに合わせる"), icon: "fit_screen" });
      menu(el && el.id !== "zoom-v" ? el : $("zoom-v"), opts).then(function (v) { if (v === "fit") setZoom(App.fitZoom() * 100); else if (v) setZoom(+v); });
      return;
    }
    App.command(cmd, el);
  }

  var readMode = false;
  function openEditor(f) {
    current = f;
    readMode = false;
    root.classList.remove("is-read");
    if (App.upgrade) App.upgrade(f.data);
    setTitle();
    setStatus(t("保存済み"));
    closeFind();
    if (App.open) App.open();
    App.render({});
    resetHist();
    setZoom(App.defaultZoom ? App.defaultZoom() * 100 : zoom);
    ribbonTab = 0;
    renderRibbon();
    placeIndicator(false);
    // 開いたときの版を、版の履歴に残す
    lastVersion = Date.now();
    Store.addVersion(f.id, f.data);
  }

  // ================================================================
  //   クリック・入力・キー
  // ================================================================
  function backToEditor() {
    var ed = $("editor");
    ed.classList.remove("is-back"); show("editor"); void ed.offsetWidth; ed.classList.add("is-back");
    placeIndicator(false);
    if (App && App.focus) App.focus();
  }
  document.addEventListener("click", function (e) {
    var el;
    if ((el = e.target.closest("[data-todo]"))) return notify(t("「{name}」は、まだ使えません", { name: el.getAttribute("data-todo") }));
    if ((el = e.target.closest("[data-tpl]"))) return createFrom(el.getAttribute("data-tpl"));
    if ((el = e.target.closest("[data-open]"))) return openFile(el.getAttribute("data-open"), el.getAttribute("data-open-app"));
    if ((el = e.target.closest("[data-more]"))) return fileActions(el.getAttribute("data-more"), el.getAttribute("data-name"), el);
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
    if (e.target.closest("[data-print]")) { backToEditor(); setTimeout(doPrint, 300); return; }
    if ((el = e.target.closest("[data-export]"))) return exportAs(el.getAttribute("data-export"));
    if ((el = e.target.closest("[data-restore]"))) return restoreVersion(+el.getAttribute("data-restore"));
    if (e.target.closest("[data-keep-version]")) {
      flush();
      Store.addVersion(current.id, current.data).then(function (ok) { notify(ok ? t("いまの版を残しました") : t("いまの版は、もう残っています")); renderVersions(); });
      return;
    }
    if (e.target.closest("[data-rename]")) return renameCurrent();
    if (e.target.closest("[data-close-file]")) return closeFile();
    if (e.target.closest("[data-install]")) {
      if (installEvt) { installEvt.prompt(); installEvt.userChoice.then(function () { installEvt = null; renderStart(); }); }
      return;
    }
    // 今すぐ更新（サービスワーカーに新しい版を聞いてから、開き直す。ファイルを開いているときは押せない）
    if (e.target.closest("[data-update]")) {
      var sw = "serviceWorker" in navigator ? navigator.serviceWorker.getRegistration(BASE).then(function (r) { return r && r.update(); }) : Promise.resolve();
      sw.catch(function () { /* つながらなくても開き直す */ }).then(function () { location.reload(); });
      return;
    }
    if ((el = e.target.closest("[data-theme-set]"))) { settings.theme = el.getAttribute("data-theme-set"); saveSettings(); applySettings(); renderStart(); return; }
    if (e.target.closest("[data-avatar-btn]")) { flush(); view = "account"; renderStart(); show("home"); return; }
    if ((el = e.target.closest("[data-view]"))) { view = el.getAttribute("data-view"); query = ""; renderStart(); var m = $("start-main"); if (m) m.scrollTop = 0; return; }
    if ((el = e.target.closest("[data-ftab]"))) {
      // スタート画面全体は作り直さず、タブの印とファイルの一覧だけ変える
      fileTab = el.getAttribute("data-ftab");
      // 下線は、前のタブの下に置いてから、押したタブへ滑らせる（それまでは CSS の ::after の下線）
      var row = el.parentElement, line = row.querySelector(".tabs__ind"), from = row.querySelector('[aria-selected="true"] > span');
      if (line && from && from.parentElement !== el) { moveLine(line, from, row, false); row.classList.add("has-ind"); moveLine(line, el.firstChild, row); }
      row.querySelectorAll("[data-ftab]").forEach(function (b) { b.setAttribute("aria-selected", String(b === el)); });
      animateFiles = true;
      renderFileTable();
      return;
    }
    if ((el = e.target.closest("#cat-chips .chip"))) return filterTemplates(el.getAttribute("data-cat"));
    if (e.target.closest("[data-back-edit]")) return backToEditor();
    if (e.target.closest("[data-file]")) { flush(); view = "home"; renderStart(); $("home").classList.add("is-backstage"); show("home"); return; }
    if ((el = e.target.closest("[data-rtab]"))) { var prev = ribbonTab; ribbonTab = +el.getAttribute("data-rtab"); ribbonOpen = true; renderRibbon(prev); return; }
    if (e.target.closest("[data-fold]")) { ribbonOpen = !ribbonOpen; renderRibbon(); return; }
    if ((el = e.target.closest("[data-cmd]"))) return runCmd(el.getAttribute("data-cmd"), el);
  });
  document.addEventListener("input", function (e) {
    if (e.target.id === "file-search") { query = e.target.value.trim(); renderFileTable(); }
  });
  // オプション
  document.addEventListener("change", function (e) {
    var id = e.target.id;
    if (id === "opt-user") { settings.user = e.target.value.trim().slice(0, 40); saveSettings(); fillAvatars(); notify(settings.user ? t("名前を「{name}」にしました", { name: settings.user }) : t("名前を消しました")); }
    if (id === "acct-theme") { settings.theme = e.target.value; saveSettings(); applySettings(); }
    if (id === "opt-motion") { settings.motion = e.target.checked ? "reduce" : "auto"; saveSettings(); applySettings(); }
    if (id === "opt-autosave") { settings.autosave = e.target.checked; saveSettings(); if (settings.autosave && dirty) saveNow(); notify(settings.autosave ? t("自動保存をオンにしました") : t("自動保存をオフにしました")); }
  });
  document.addEventListener("keydown", function (e) {
    if (!current || !$("editor") || $("editor").hidden || document.querySelector(".tbs-dialog-scrim, .present")) return;
    var k = e.key.toLowerCase(), mod = e.ctrlKey || e.metaKey;
    // 文字の欄（数式バー・検索など）の中の Ctrl+Z などは、その欄のものを使う（data-office-keys の付いた欄は、アプリのもの）
    var field = e.target.matches && e.target.matches("input, textarea, select") && !e.target.hasAttribute("data-office-keys");
    if (mod && k === "s") { e.preventDefault(); runCmd("save"); return; }
    if (mod && k === "p") { e.preventDefault(); doPrint(); return; }
    if (mod && !e.shiftKey && k === "f") { e.preventDefault(); openFind(); return; }
    if (mod && k === "h") { e.preventDefault(); openFind(true); return; }
    if (!field && mod && !e.altKey && (k === "z" || k === "y")) {
      e.preventDefault();
      runCmd(k === "y" || e.shiftKey ? "redo" : "undo");
      return;
    }
    if (e.key === "Escape" && !$("findbar").hidden && e.target.closest && e.target.closest("#findbar")) { closeFind(); return; }
    if (App.key) App.key(e, field);
  });

  // ================================================================
  //   画面の切りかえ
  // ================================================================
  function show(v) {
    ["home", "editor"].forEach(function (id) { var el = $(id); if (el) el.hidden = id !== v; });
    closeMenu();
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
    if (!id || !$("editor") || !App) { showStart(); return Promise.resolve(); }
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
    // 同じ場所（手元・sk-hub.jp）のほかのタブで、ログイン・ログアウトしたとき
    if (e.key === "skhub_account") { fillAvatars(); if ($("home") && !$("home").hidden && view === "account") renderStart(); }
  });

  applySettings();
  buildEditor();
  fillAvatars();
  // office.sk-hub.jp: 開いたときに、sk-hub.jp のログインの状態を受け取る
  if (bridge) setTimeout(function () { syncAccount(); }, 600);
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
  }, reduced ? 0 : 1000);
})();
