// LibreOffice（LOKit）の文書の部品（デスクトップ版の Doc の本文）
//   OfficeDocLok(host, bytes) → 操作の部品。host の中に、描画・カーソル・入力の欄を作る
//   ・bytes: 開く docx（Uint8Array）。表示用のサーバー（desktop の lok_preview）が開く
//   ・ready: 準備ができたら解決する Promise
//   ・save(): いまの内容を docx（Uint8Array）にして返す
//   ・ctrl("b" など): Ctrl＋文字のキーを送る（太字 b・斜体 i・下線 u・元に戻す z・やり直す y）
//   ・focus(), dispose()（サーバーを止める）
//   キーの番号（css::awt::Key）と、描画の地の色の処理は、desktop/lok-spike で確かめたものを使う
(function () {
  "use strict";

  var W = 794, H = 1123;
  var STYLE = [
    ".doc-lok { position: relative; height: 100%; min-height: 0; overflow-y: auto; display: flex; justify-content: center; background: #e6e6e6; }",
    ".doc-lok__paper { position: relative; width: 794px; height: 1123px; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,.2); margin: 24px 0; flex: none; }",
    ".doc-lok__page { display: block; width: 100%; height: 100%; cursor: text; user-select: none; }",
    ".doc-lok__caret { position: absolute; display: none; width: 2px; background: #000; pointer-events: none; animation: doc-lok-blink 1s steps(1) infinite; }",
    ".doc-lok__sel { position: absolute; background: rgba(51, 144, 255, .35); pointer-events: none; }",
    ".doc-lok__ime { position: fixed; left: 0; top: 0; width: 1px; height: 1px; opacity: 0; border: 0; padding: 0; }",
    "@keyframes doc-lok-blink { 50% { opacity: 0; } }"
  ].join("\n");

  function injectStyle() {
    if (document.getElementById("doc-lok-style")) return;
    var s = document.createElement("style");
    s.id = "doc-lok-style";
    s.textContent = STYLE;
    document.head.appendChild(s);
  }

  // 特殊キー → css::awt::Key の番号
  var SPECIAL = {
    Enter: { kc: 1280, cp: 13 }, Backspace: { kc: 1283, cp: 8 }, Delete: { kc: 1286, cp: 0 },
    ArrowLeft: { kc: 1026, cp: 0 }, ArrowRight: { kc: 1027, cp: 0 }, ArrowUp: { kc: 1025, cp: 0 }, ArrowDown: { kc: 1024, cp: 0 },
    Home: { kc: 1028, cp: 0 }, End: { kc: 1029, cp: 0 }
  };
  // Ctrl＋文字: 修飾（Ctrl = 0x2000 = 8192）＋ 文字の番号（A = 512 から順）
  var CTRL = { b: 8705, i: 8712, u: 8724, z: 8729, y: 8728 };

  function hexList(s) {
    return Array.from(s).map(function (c) { return c.codePointAt(0).toString(16).toUpperCase(); }).join(",");
  }

  // LibreOffice は、ページの外側を真っ黒に描く。画面の端からつながる黒だけを、表示の地の色に変える
  function cleanMargin(blob) {
    return createImageBitmap(blob).then(function (bmp) {
      var cv = document.createElement("canvas");
      cv.width = bmp.width; cv.height = bmp.height;
      var g = cv.getContext("2d");
      g.drawImage(bmp, 0, 0);
      var w = cv.width, h = cv.height;
      var id = g.getImageData(0, 0, w, h), d = id.data;
      var seen = new Uint8Array(w * h), stack = [];
      for (var x = 0; x < w; x++) { stack.push(x, (h - 1) * w + x); }
      for (var y = 0; y < h; y++) { stack.push(y * w, y * w + w - 1); }
      while (stack.length) {
        var p = stack.pop();
        if (seen[p]) continue;
        seen[p] = 1;
        var i = p * 4;
        if (d[i] !== 0 || d[i + 1] !== 0 || d[i + 2] !== 0) continue;
        d[i] = d[i + 1] = d[i + 2] = 0xf2;
        var px = p % w;
        if (px > 0) stack.push(p - 1);
        if (px < w - 1) stack.push(p + 1);
        if (p >= w) stack.push(p - w);
        if (p < (h - 1) * w) stack.push(p + w);
      }
      g.putImageData(id, 0, 0);
      return new Promise(function (resolve) { cv.toBlob(resolve, "image/png"); });
    });
  }

  window.OfficeDocLok = function (host, bytes, opts) {
    injectStyle();
    // 中身が変わったことを、Doc の側に知らせる（自動保存のため）
    var changed = (opts && opts.onChange) || function () {};
    host.innerHTML = "";
    var root = document.createElement("div");
    root.className = "doc-lok";
    root.innerHTML = '<div class="doc-lok__paper"><img class="doc-lok__page" alt="" draggable="false"><div class="doc-lok__caret"></div></div>';
    var ime = document.createElement("textarea");
    ime.className = "doc-lok__ime";
    ime.setAttribute("aria-label", "入力");
    host.appendChild(root);
    host.appendChild(ime);

    var paper = root.querySelector(".doc-lok__paper");
    var page = root.querySelector(".doc-lok__page");
    var caret = root.querySelector(".doc-lok__caret");
    var base = "";
    var yPx = 0, docPx = 0, timer = null, renderSeq = 0, objUrl = null;
    var dragging = false, moveBusy = false, disposed = false;

    function post(path) {
      return fetch(base + path, { method: "POST" });
    }

    // 入力・命令のあとに描き直す。スクロール（scrolling）のときは、中身は変わらないので知らせない
    function scheduleRender(delay, scrolling) {
      clearTimeout(timer);
      if (!scrolling) changed();
      timer = setTimeout(function () { render(); setTimeout(render, 600); }, delay === undefined ? 300 : delay);
    }

    function render() {
      if (disposed || !base) return Promise.resolve();
      var seq = ++renderSeq;
      return fetch(base + "/tile?y=" + (yPx * 15) + "&w=" + W + "&h=" + H).then(function (res) {
        if (!res.ok) throw new Error("tile " + res.status);
        return res.blob();
      }).then(cleanMargin).then(function (blob) {
        if (seq !== renderSeq || disposed) return;
        var old = objUrl;
        objUrl = URL.createObjectURL(blob);
        page.src = objUrl;
        if (old) URL.revokeObjectURL(old);
        return updateCaret();
      }).catch(function () { /* 描き直しに失敗しても、次の入力で描き直す */ });
    }

    function drawSelection(sel) {
      root.querySelectorAll(".doc-lok__sel").forEach(function (n) { n.remove(); });
      if (!sel) return;
      sel.split(";").forEach(function (r) {
        var v = r.split(",").map(function (s) { return parseInt(s, 10); });
        if (v.length !== 4 || v.some(isNaN)) return;
        var d = document.createElement("div");
        d.className = "doc-lok__sel";
        d.style.left = (v[0] / 15) + "px";
        d.style.top = (v[1] / 15 - yPx) + "px";
        d.style.width = (v[2] / 15) + "px";
        d.style.height = (v[3] / 15) + "px";
        paper.appendChild(d);
      });
    }

    function updateCaret() {
      return fetch(base + "/cursor").then(function (r) { return r.json(); }).then(function (c) {
        drawSelection(c.sel);
        var top = c.y / 15 - yPx;
        if (!c.valid || top < 0 || top >= H) { caret.style.display = "none"; return; }
        caret.style.display = "block";
        caret.style.left = (c.x / 15) + "px";
        caret.style.top = top + "px";
        caret.style.height = Math.max(12, c.h / 15) + "px";
      });
    }

    function toTwips(e) {
      var rect = page.getBoundingClientRect();
      return { x: Math.round((e.clientX - rect.left) * 15), y: Math.round((e.clientY - rect.top + yPx) * 15) };
    }

    function ctrlKey(k) {
      if (!CTRL[k]) return;
      return post("/key?type=0&cp=0&kc=" + CTRL[k]).then(function () {
        return post("/key?type=1&cp=0&kc=" + CTRL[k]);
      }).then(function () { scheduleRender(); });
    }

    // キー入力（IME の変換中は除く）
    ime.addEventListener("keydown", function (e) {
      if (e.isComposing || e.keyCode === 229 || disposed) return;
      if (e.ctrlKey && !e.altKey && !e.metaKey && CTRL[e.key.toLowerCase()]) {
        e.preventDefault();
        ctrlKey(e.key.toLowerCase());
        return;
      }
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      var sp = SPECIAL[e.key];
      if (!sp && e.key.length !== 1) return;
      e.preventDefault();
      var cp = sp ? sp.cp.toString(16).toUpperCase() : e.key.codePointAt(0).toString(16).toUpperCase();
      var shift = e.shiftKey && /^(Arrow|Home|End)/.test(e.key) ? 4096 : 0;
      var kc = sp ? sp.kc + shift : 0;
      post("/key?type=0&cp=" + cp + "&kc=" + kc).then(function () {
        return post("/key?type=1&cp=" + cp + "&kc=" + kc);
      }).then(function () { scheduleRender(); });
    });

    // 変換中（確定の前の文字）と確定。確定は、確定する文字を変換中として送ってから行う
    ime.addEventListener("compositionupdate", function (e) {
      post("/ime?type=0&text=" + hexList(e.data)).then(function () { scheduleRender(); });
    });
    ime.addEventListener("compositionend", function (e) {
      var list = hexList(e.data);
      post("/ime?type=0&text=" + list).then(function () {
        return post("/ime?type=2&text=" + list);
      }).then(function () { ime.value = ""; scheduleRender(); });
    });
    // keydown を通らずに届いた文字（貼りつけなど）
    ime.addEventListener("input", function (e) {
      if (e.isComposing || !ime.value || disposed) return;
      var text = ime.value;
      ime.value = "";
      var chain = Promise.resolve();
      Array.from(text).forEach(function (ch) {
        var cp = ch.codePointAt(0).toString(16).toUpperCase();
        chain = chain.then(function () { return post("/key?type=0&cp=" + cp); })
          .then(function () { return post("/key?type=1&cp=" + cp); });
      });
      chain.then(function () { scheduleRender(); });
    });

    // クリックとドラッグで位置と選択を決める
    page.addEventListener("mousedown", function (e) {
      e.preventDefault();
      ime.focus({ preventScroll: true });
      dragging = true;
      var p = toTwips(e);
      post("/mouse?type=0&x=" + p.x + "&y=" + p.y);
    });
    window.addEventListener("mousemove", function (e) {
      if (!dragging || moveBusy || disposed) return;
      moveBusy = true;
      var p = toTwips(e);
      post("/mouse?type=2&x=" + p.x + "&y=" + p.y).then(function () { moveBusy = false; }, function () { moveBusy = false; });
    });
    window.addEventListener("mouseup", function (e) {
      if (!dragging) return;
      dragging = false;
      var p = toTwips(e);
      post("/mouse?type=1&x=" + p.x + "&y=" + p.y).then(function () { scheduleRender(200); });
    });

    root.addEventListener("wheel", function (e) {
      e.preventDefault();
      var maxY = Math.max(0, docPx - H);
      yPx = Math.min(maxY, Math.max(0, yPx + Math.sign(e.deltaY) * 120));
      scheduleRender(0, true);
    }, { passive: false });

    var ctl = {
      ready: null,
      ctrl: ctrlKey,
      focus: function () { ime.focus({ preventScroll: true }); },
      save: function () {
        return post("/save").then(function (r) {
          if (!r.ok) throw new Error("save " + r.status);
          return fetch(base + "/saved.docx");
        }).then(function (r) {
          if (!r.ok) throw new Error("saved " + r.status);
          return r.arrayBuffer();
        }).then(function (ab) { return new Uint8Array(ab); });
      },
      dispose: function () {
        disposed = true;
        clearTimeout(timer);
        if (objUrl) URL.revokeObjectURL(objUrl);
        host.innerHTML = "";
        return window.OfficePlatform.lokStop().catch(function () {});
      }
    };

    // 表示用のサーバーを起動して、文書を描く
    ctl.ready = window.OfficePlatform.lokPreview(Array.prototype.slice.call(bytes)).then(function (url) {
      base = String(url).replace(/\/$/, "");
      return fetch(base + "/size").then(function (r) { return r.json(); }).then(function (s) {
        docPx = Math.ceil(s.h / 15);
      });
    }).then(render).then(function () { ime.focus({ preventScroll: true }); });

    return ctl;
  };
})();
