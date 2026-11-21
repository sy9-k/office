// SK Hub Office の招待制（SK Study と同じ。招待された Google アカウントで SK Hub Systems アカウントにログインした人だけが使える）
//   ・招待は sk-hub.jp の管理画面（admin.sk-hub.jp/invites/）で、Firestore の hub_invites に入れる（sk-hub.jp の assets/hub/invite.js）
//   ・office.sk-hub.jp では、sk-hub.jp の /assets/hub/bridge.html（見えない iframe）に「招待されているか」を聞く（op: "invite"）
//   ・デスクトップ版は、ログインのとき（sk-hub.jp/account/?desktop=…）に招待されているかを受け取って覚えておく（office.js の desktopLogin → OfficeGate.setDesktop）
//   ・招待を確かめられた端末には印（sk_office_invited）を残す。印があれば、すぐに開いて、うしろで確かめ直す（オフラインでも開ける）
//     確かめ直して、ログアウトしていた・招待が取り消されていたら、この画面を出す
//   ・手元（localhost）では出さない（?gate を付けると、この画面を確かめられる）
//   ・すべてのページの <head> で、i18n.js のあとに読み込む。この画面のあいだは、body のほかの要素を見えなくする
//   ・ファイルは PC・Google ドライブにあり、SK のサーバーにはないので、これは使いはじめの入り口を閉じるためのもの
(function () {
  "use strict";
  var root = document.documentElement;
  var T = window.__TAURI__;
  var DESKTOP = !!(T && T.core && typeof T.core.invoke === "function");
  var ON_SUB = /^office\.sk-hub\.jp$/.test(location.hostname);
  var PREVIEW = /[?&]gate\b/.test(location.search);
  var SITE = "https://sk-hub.jp";
  var KEY = "sk_office_invited";
  var BASE = (function () { var s = document.currentScript; try { return s ? new URL("..", s.src).pathname : "/"; } catch (e) { return "/"; } })();
  var I18N = window.SKI18N;
  function t(s) { return I18N ? I18N.t(s) : s; }

  var gate = { ok: true, recheck: function () {}, setDesktop: function () {}, reset: function () {} };
  window.OfficeGate = gate;
  if (!ON_SUB && !DESKTOP && !PREVIEW) return;

  function readMark() { try { return JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) { return null; } }
  function mark(on) { try { if (on) localStorage.setItem(KEY, JSON.stringify({ at: Date.now() })); else localStorage.removeItem(KEY); } catch (e) { /* 残せなくても使える */ } }

  // ---------- 見た目 ----------
  var css =
    "html.office-gated,html.office-gated body{overflow:hidden!important}" +
    "html.office-gated body>:not(#office-gate){display:none!important}" +
    "#office-gate{position:fixed;inset:0;z-index:2147483646;display:grid;place-items:center;padding:24px 16px;overflow:auto;box-sizing:border-box;" +
    "background:radial-gradient(640px 360px at 90% 0%,color-mix(in srgb,var(--doc-c,#2563eb) 14%,transparent),transparent 70%)," +
    "radial-gradient(560px 320px at 0% 100%,color-mix(in srgb,var(--sheet-c,#16a34a) 12%,transparent),transparent 70%),var(--md-surface,#f8f9fd);" +
    "color:var(--md-on-surface,#1a1c23);font-family:var(--md-font,system-ui,sans-serif)}" +
    "#office-gate *{box-sizing:border-box}" +
    "#office-gate .og-card{width:100%;max-width:460px;padding:32px 28px;border-radius:28px;background:var(--md-surface-container-low,#f1f2f8);text-align:center}" +
    "#office-gate .og-logo{width:64px;height:64px;display:block;margin:0 auto 12px}" +
    "#office-gate .og-name{margin:0;font-size:14px;font-weight:700;color:var(--md-on-surface-variant,#444853);letter-spacing:.02em}" +
    "#office-gate h1{margin:18px 0 8px;font-size:24px;line-height:1.35;font-weight:700}" +
    "#office-gate p{margin:0 0 12px;font-size:15px;line-height:1.75;color:var(--md-on-surface-variant,#444853)}" +
    "#office-gate .og-actions{display:grid;gap:10px;margin-top:22px}" +
    "#office-gate .og-actions a,#office-gate .og-actions button{width:100%;justify-content:center;text-decoration:none}" +
    "#office-gate .og-apps{display:flex;justify-content:center;gap:14px;margin:22px 0 0;padding:0;list-style:none;font-size:12.5px;color:var(--md-on-surface-variant,#444853)}" +
    "#office-gate .og-apps img{width:28px;height:28px;display:block;margin:0 auto 4px}" +
    "#office-gate .og-spin{width:28px;height:28px;margin:22px auto 4px;border-radius:50%;border:3px solid var(--md-outline-variant,#c4c6d3);border-top-color:var(--md-primary,#35456b);animation:og-spin .9s linear infinite}" +
    "@keyframes og-spin{to{transform:rotate(1turn)}}" +
    "@media (prefers-reduced-motion:reduce){#office-gate .og-spin{animation-duration:3s}}";

  var box = null;
  function el(tag, attrs, text) {
    var n = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) { n.setAttribute(k, attrs[k]); });
    if (text) n.textContent = text;
    return n;
  }
  function whenBody(fn) {
    if (document.body) fn();
    else document.addEventListener("DOMContentLoaded", fn, { once: true });
  }
  function loginUrl() {
    return SITE + "/account/?next=" + encodeURIComponent((BASE === "/" ? "/office" : "") + location.pathname + location.hash);
  }
  function linkBtn(cls, label, href) { var a = el("a", { class: cls, href: href }, label); return a; }
  function btn(cls, label, fn) { var b = el("button", { class: cls, type: "button" }, label); b.addEventListener("click", fn); return b; }
  function loginBtn(cls, label) {
    if (DESKTOP) return btn(cls, label, function () { document.dispatchEvent(new CustomEvent("office-gate-login")); });
    return linkBtn(cls, label, loginUrl());
  }
  var WAITING = SITE + "/waiting/?product=office";

  // state: checking / signed-out / not-invited / offline
  function show(state, reason) {
    gate.ok = false;
    root.classList.add("office-gated");
    whenBody(function () {
      if (!box) {
        var style = el("style");
        style.textContent = css;
        document.head.appendChild(style);
        box = el("div", { id: "office-gate", role: "dialog", "aria-modal": "true", "aria-labelledby": "og-title" });
        document.body.appendChild(box);
      }
      var card = el("div", { class: "og-card" });
      var logo = el("img", { class: "og-logo", src: BASE + "icons/office-logo.png", alt: "", draggable: "false" });
      card.appendChild(logo);
      card.appendChild(el("p", { class: "og-name" }, "SK Hub Office"));
      var actions = el("div", { class: "og-actions" });
      var title, texts = [];
      if (state === "checking") {
        title = t("ログインを確かめています…");
      } else if (state === "signed-out") {
        title = t("ログインしてください");
        texts.push(t("SK Hub Office は、いまは招待制です。招待された Google アカウントで SK Hub Systems アカウントにログインすると使えます。"));
        actions.appendChild(loginBtn("m3-btn", t("ログイン")));
        actions.appendChild(linkBtn("m3-btn m3-btn--tonal", t("順番待ちに登録する"), WAITING));
      } else if (state === "not-invited") {
        title = t("招待制です");
        texts.push(t("SK Hub Office は、いまは招待した方だけが使えます。順番待ちに登録しておくと、順番がきたらメールでお知らせします。"));
        if (reason === "not-google") texts.push(t("招待された Google アカウントでログインしてください。Google 以外のログインでは開けません。"));
        actions.appendChild(linkBtn("m3-btn", t("順番待ちに登録する"), WAITING));
        actions.appendChild(loginBtn("m3-btn m3-btn--tonal", t("別のアカウントでログイン")));
      } else {
        title = t("接続できませんでした");
        texts.push(t("招待されているかを確かめられませんでした。ネットワークを確かめて、もう一度お試しください。"));
        actions.appendChild(btn("m3-btn", t("もう一度"), function () { check(true); }));
      }
      var h1 = el("h1", { id: "og-title" }, title);
      card.appendChild(h1);
      texts.forEach(function (s) { card.appendChild(el("p", null, s)); });
      if (state === "checking") card.appendChild(el("div", { class: "og-spin", "aria-hidden": "true" }));
      else card.appendChild(actions);
      var apps = el("ul", { class: "og-apps", "aria-hidden": "true" });
      [["doc", "Doc"], ["sheet", "Sheet"], ["slide", "Slide"]].forEach(function (a) {
        var li = el("li");
        li.appendChild(el("img", { src: BASE + "icons/" + a[0] + "-logo.png", alt: "" }));
        li.appendChild(document.createTextNode(a[1]));
        apps.appendChild(li);
      });
      card.appendChild(apps);
      box.replaceChildren(card);
    });
  }
  function hide() {
    gate.ok = true;
    root.classList.remove("office-gated");
    if (box) { box.remove(); box = null; }
    document.dispatchEvent(new CustomEvent("office-gate-open"));
  }

  // ---------- 確かめる（office.sk-hub.jp: sk-hub.jp の bridge に聞く） ----------
  var frame = null, loaded = null, seq = 0, waiting = {};
  function bridge() {
    if (loaded) return loaded;
    loaded = new Promise(function (resolve) {
      window.addEventListener("message", function (e) {
        var m = e.data;
        if (e.origin !== SITE || !m || m.sk !== "bridge" || !frame || e.source !== frame.contentWindow) return;
        if (m.ready) { resolve(); return; }
        if (waiting[m.id]) { waiting[m.id](m.value); delete waiting[m.id]; }
      });
      whenBody(function () {
        frame = el("iframe", { src: SITE + "/assets/hub/bridge.html", title: "", tabindex: "-1", "aria-hidden": "true" });
        frame.hidden = true;
        document.body.appendChild(frame);
      });
    });
    return loaded;
  }
  // 招待されているか。{ ok, reason } / null（ログインしていない）/ undefined（確かめられなかった）
  function askInvite() {
    return Promise.race([
      bridge().then(function () {
        return new Promise(function (resolve) {
          var id = "gate" + (++seq);
          waiting[id] = resolve;
          frame.contentWindow.postMessage({ sk: "bridge", id: id, op: "invite", product: "office" }, SITE);
        });
      }),
      new Promise(function (resolve) { setTimeout(function () { resolve(undefined); }, 20000); })
    ]);
  }

  var checking = false, desktopReason = "";
  function check(loud) {
    if (DESKTOP) { if (readMark()) hide(); else if (desktopReason) show("not-invited", desktopReason); else show("signed-out"); return; }
    if (PREVIEW && !ON_SUB) { show(/[?&]gate=(\w[\w-]*)/.test(location.search) ? RegExp.$1 : "signed-out"); return; }
    if (checking) return;
    checking = true;
    var had = !!readMark();
    if (!had || loud) show("checking");
    askInvite().then(function (r) {
      checking = false;
      if (r && r.ok) { mark(true); hide(); return; }
      if (r === null) { mark(false); show("signed-out"); return; }
      if (r && (r.reason === "not-invited" || r.reason === "not-google")) { mark(false); show("not-invited", r.reason); return; }
      // 確かめられなかった（オフラインなど）。前に招待を確かめた端末なら、そのまま使える
      if (had) hide(); else show("offline");
    });
  }

  gate.recheck = function () { check(false); };
  // デスクトップ版: ログインしたときに受け取った招待。invited: 招待されているか / reason: not-invited・not-google（招待されていないとき）
  gate.setDesktop = function (invited, reason) { desktopReason = invited ? "" : reason || "not-invited"; mark(!!invited); check(false); };
  // ログアウトしたとき
  gate.reset = function () { desktopReason = ""; mark(false); check(false); };

  // 印があれば、すぐに開いて、うしろで確かめ直す。なければ、確かめ終わるまで、この画面を出しておく
  if (!readMark() || PREVIEW) root.classList.add("office-gated");
  check(false);
  // ログインのページから戻ってきたとき（この画面が出ているときだけ）
  document.addEventListener("visibilitychange", function () { if (!document.hidden && !gate.ok && !DESKTOP) check(false); });
})();
