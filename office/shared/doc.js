// SK Hub Office の Doc（文書）。window.OfficeApps.doc に登録し、/office/shared/office.js が呼ぶ（O は office.js が渡す道具）
//   ファイルの中身（data）: { html: 本文, page: 用紙, comments: [コメント] }
//     page: { size: "a4|letter|b5|a3", orient: "portrait|landscape", margin: "normal|narrow|wide", cols: 1〜3, vert: 縦書き, header, footer, num: ページ番号 }
//     comments: [{ id, text, author, ts }]（本文の中では <span class="cmt" data-cmt="id"> で印を付ける）
//   本文は contenteditable の <article id="page">。書式は document.execCommand（styleWithCSS で、色・大きさは span の style になる）
//   画面: ナビゲーション（見出しの一覧）・用紙（ヘッダー・本文・フッター）・コメント。用紙の背景に、印刷したときのページの区切りの線を引く
//   元に戻すは office.js（中身をまるごと覚える）。打っているあいだはまとめて 1 回にする
//   .docx の読み書きは zip.js、グラフは chart.js（図は <figure class="chart" data-spec="…"> に SVG を入れる）
(function () {
  "use strict";
  window.OfficeApps = window.OfficeApps || {};
  window.OfficeApps.doc = function (O) {
    var t = O.t, esc = O.esc, icon = O.icon, $ = O.$, Zip = window.OfficeZip, Chart = window.OfficeChart;

    // ---------- 用紙 ----------
    var SIZES = { a4: { w: 794, h: 1123, label: "A4", tw: 11906, th: 16838, css: "A4" }, b5: { w: 688, h: 972, label: "B5", tw: 10318, th: 14570, css: "182mm 257mm" },
      a3: { w: 1123, h: 1587, label: "A3", tw: 16838, th: 23811, css: "A3" }, letter: { w: 816, h: 1056, label: "Letter", tw: 12240, th: 15840, css: "letter" } };
    var MARGINS = { normal: { x: 96, y: 96, label: t("標準"), mm: "25.4mm" }, narrow: { x: 48, y: 48, label: t("狭い"), mm: "12.7mm" }, wide: { x: 144, y: 96, label: t("広い"), mm: "25.4mm 38.1mm" } };
    function pageOf(d) {
      var p = d.page || (d.page = {});
      if (!SIZES[p.size]) p.size = "a4";
      if (p.orient !== "landscape") p.orient = "portrait";
      if (!MARGINS[p.margin]) p.margin = "normal";
      p.cols = Math.max(1, Math.min(3, +p.cols || 1));
      return p;
    }
    function dims(p) { var s = SIZES[p.size]; return p.orient === "landscape" ? { w: s.h, h: s.w } : { w: s.w, h: s.h }; }
    function upgrade(d) {
      if (typeof d.html !== "string") d.html = "<p><br></p>";
      pageOf(d);
      if (!Array.isArray(d.comments)) d.comments = [];
      return d;
    }

    // ---------- 安全な HTML（貼り付け・読み込み） ----------
    var ALLOW = /^(P|H1|H2|H3|H4|H5|H6|UL|OL|LI|TABLE|THEAD|TBODY|TFOOT|TR|TD|TH|CAPTION|COLGROUP|COL|B|STRONG|I|EM|U|S|STRIKE|DEL|INS|SUB|SUP|SPAN|A|BR|HR|IMG|BLOCKQUOTE|PRE|CODE|DIV|FIGURE|FIGCAPTION|MARK|SMALL|FONT|SECTION|NAV|SVG|G|PATH|RECT|CIRCLE|LINE|POLYLINE|TEXT|TITLE)$/i;
    var STYLE_OK = /^(color|background-color|font-weight|font-style|text-decoration|text-decoration-line|text-align|font-size|font-family|line-height|width|vertical-align|float|margin-left|margin-right)$/i;
    function sanitize(html, keepSvg) {
      var doc = new DOMParser().parseFromString("<body>" + html + "</body>", "text/html");
      (function walk(node) {
        Array.prototype.slice.call(node.childNodes).forEach(function (n) {
          if (n.nodeType === 8) { n.remove(); return; }
          if (n.nodeType !== 1) return;
          var tag = n.tagName;
          if (/^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|LINK|META|NOSCRIPT|TEMPLATE|FORM|INPUT|BUTTON|SELECT|TEXTAREA)$/i.test(tag) || (!keepSvg && /^svg$/i.test(tag))) { n.remove(); return; }
          if (!ALLOW.test(tag)) { walk(n); n.replaceWith.apply(n, Array.prototype.slice.call(n.childNodes)); return; }
          Array.prototype.slice.call(n.attributes).forEach(function (a) {
            var name = a.name.toLowerCase(), v = a.value;
            if (/^on/.test(name)) { n.removeAttribute(a.name); return; }
            if ((name === "href" || name === "src" || name === "xlink:href") && /^\s*(javascript|vbscript|data:text)/i.test(v)) { n.removeAttribute(a.name); return; }
            if (name === "src" && !/^(data:image\/|https?:|blob:)/i.test(v)) { n.removeAttribute(a.name); return; }
            if (name === "style") {
              var keep = v.split(";").map(function (x) { return x.trim(); }).filter(function (x) { var k = x.split(":")[0]; return k && STYLE_OK.test(k.trim()) && !/url\s*\(|expression/i.test(x); });
              if (keep.length) n.setAttribute("style", keep.join(";")); else n.removeAttribute("style");
              return;
            }
            if (/^(href|src|alt|title|colspan|rowspan|class|data-cmt|data-spec|data-fn|id|width|height|viewbox|d|x|y|x1|x2|y1|y2|cx|cy|r|fill|stroke|stroke-width|points|font-size|text-anchor|font-weight|xmlns|rx|transform|fill-opacity|stroke-linejoin|stroke-linecap|pointer-events|font-family|contenteditable)$/i.test(name)) return;
            n.removeAttribute(a.name);
          });
          if (tag === "A") { n.setAttribute("target", "_blank"); n.setAttribute("rel", "noopener"); }
          walk(n);
        });
      })(doc.body);
      return doc.body.innerHTML;
    }

    // ---------- いまの状態 ----------
    var data = null, page = null, saved = null, readOnly = false, imgSel = null, painter = null;
    var showNav = false, showCmts = false, zoomZ = 1;
    function el(id) { return $(id); }
    function sync() { if (!page) return; data.html = page.innerHTML; }
    function changed(typing) { sync(); stat(); O.changed(typing); if (showNav) drawNav(); }
    function inPage(node) { return !!(page && node && page.contains(node)); }
    function keep() { var s = getSelection(); if (s.rangeCount && inPage(s.anchorNode)) saved = s.getRangeAt(0).cloneRange(); }
    function back() {
      if (!page) return;
      if (document.activeElement !== page) page.focus({ preventScroll: true });
      var s = getSelection();
      if (saved && inPage(saved.startContainer)) { s.removeAllRanges(); s.addRange(saved); }
      else if (!s.rangeCount || !inPage(s.anchorNode)) { var r = document.createRange(); r.selectNodeContents(page); r.collapse(false); s.removeAllRanges(); s.addRange(r); }
    }
    function exec(cmd, v) { back(); document.execCommand("styleWithCSS", false, true); document.execCommand(cmd, false, v === undefined ? null : v); keep(); changed(); }
    function selRange() { var s = getSelection(); return s.rangeCount && inPage(s.anchorNode) ? s.getRangeAt(0) : saved; }
    function anchorEl() {
      var r = selRange();
      if (!r) return null;
      var n = r.startContainer;
      return n.nodeType === 1 ? (n.childNodes[r.startOffset] && n.childNodes[r.startOffset].nodeType === 1 ? n.childNodes[r.startOffset] : n) : n.parentNode;
    }
    function closest(sel) { var a = anchorEl(); return a && inPage(a) ? a.closest(sel) : null; }
    function blocks() {
      // 選んだところの段落（p・見出し・箇条書きの 1 行など）
      var r = selRange(), out = [];
      if (!r) return out;
      var BL = "p,h1,h2,h3,h4,h5,h6,li,blockquote,pre,td,th,div:not(.toc)";
      function add(n) { var b = (n.nodeType === 1 ? n : n.parentNode).closest(BL); if (b && inPage(b) && b !== page && out.indexOf(b) < 0) out.push(b); }
      add(r.startContainer);
      var w = document.createTreeWalker(page, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      w.currentNode = r.startContainer;
      while (w.nextNode()) { if (!r.intersectsNode(w.currentNode)) break; add(w.currentNode); }
      add(r.endContainer);
      return out;
    }
    function insertHtml(html) { back(); document.execCommand("insertHTML", false, html); keep(); changed(); }

    // ---------- 画面を作る ----------
    function render(opts) {
      data = O.data();
      upgrade(data);
      var caret = opts && opts.restore ? caretOffset() : null, sy = el("doc-scroll") ? el("doc-scroll").scrollTop : 0;
      imgSel = null; painter = null;
      $("canvas").innerHTML =
        '<div class="doc-wrap" id="doc-wrap">' +
          '<aside class="doc-nav" id="doc-nav"' + (showNav ? "" : " hidden") + ' aria-label="' + esc(t("ナビゲーション")) + '"></aside>' +
          '<div class="doc-scroll" id="doc-scroll">' +
            '<div class="doc-page" id="paper">' +
              '<input class="doc-hf doc-hf--head" id="doc-head" placeholder="' + esc(t("ヘッダー")) + '" aria-label="' + esc(t("ヘッダー")) + '" autocomplete="off">' +
              '<article class="doc-body" id="page" contenteditable="true" spellcheck="false" aria-label="' + esc(t("文書の本文")) + '"></article>' +
              '<input class="doc-hf doc-hf--foot" id="doc-foot" placeholder="' + esc(t("フッター")) + '" aria-label="' + esc(t("フッター")) + '" autocomplete="off">' +
              '<span class="img-rs" id="img-rs" hidden></span>' +
            "</div>" +
          "</div>" +
          '<aside class="doc-cmts" id="doc-cmts"' + (showCmts ? "" : " hidden") + ' aria-label="' + esc(t("コメント")) + '"></aside>' +
        "</div>";
      page = el("page");
      // ほかの人から受け取ったファイルも開くので、ファイルの中の HTML はいつも安全なものだけにしてから入れる
      page.innerHTML = sanitize(data.html || "", true) || "<p><br></p>";
      fixCharts();
      layout();
      wire();
      if (readOnly) page.contentEditable = "false";
      el("doc-scroll").scrollTop = sy;
      if (caret !== null) setCaret(caret);
      stat();
      drawNav(); drawCmts();
      applyZoom();
    }
    function layout() {
      var p = pageOf(data), d = dims(p), m = MARGINS[p.margin], paper = el("paper");
      if (!paper) return;
      paper.style.setProperty("--pw", d.w + "px");
      paper.style.setProperty("--ph", d.h + "px");
      paper.style.setProperty("--mx", m.x + "px");
      paper.style.setProperty("--my", m.y + "px");
      paper.classList.toggle("is-vert", !!p.vert);
      el("doc-scroll").classList.toggle("is-vert", !!p.vert);
      page.style.columnCount = p.cols > 1 ? p.cols : "";
      el("doc-head").value = p.header || "";
      el("doc-foot").value = p.footer || "";
      paper.classList.toggle("has-num", !!p.num);
    }
    function applyZoom() { var p = el("paper"); if (p) p.style.zoom = zoomZ; placeImgHandle(); }
    // グラフ（data-spec の中身から SVG を作りなおす。ダーク・ライトに関係なく白い紙に描く）
    function chartHtml(spec) {
      return '<figure class="chart" contenteditable="false" data-spec="' + esc(JSON.stringify(spec)) + '">' + Chart.svg(spec, { w: 560, h: 320 }) + "</figure>";
    }
    function fixCharts() {
      page.querySelectorAll("figure.chart").forEach(function (f) {
        f.setAttribute("contenteditable", "false");
        try { var spec = JSON.parse(f.getAttribute("data-spec") || "{}"); if (!f.querySelector("svg")) f.innerHTML = Chart.svg(spec, { w: 560, h: 320 }); } catch (e) { /* そのまま */ }
      });
      page.querySelectorAll(".toc").forEach(function (x) { x.setAttribute("contenteditable", "false"); });
    }

    // ---------- カーソルの位置（元に戻したあとに、だいたい同じところへ） ----------
    function caretOffset() {
      var s = getSelection();
      if (!page || !s.rangeCount || !inPage(s.anchorNode)) return null;
      var r = document.createRange();
      r.selectNodeContents(page);
      r.setEnd(s.anchorNode, s.anchorOffset);
      return r.toString().length;
    }
    function setCaret(n) {
      var w = document.createTreeWalker(page, NodeFilter.SHOW_TEXT), left = n, node;
      while ((node = w.nextNode())) {
        if (left <= node.data.length) {
          var r = document.createRange(); r.setStart(node, left); r.collapse(true);
          page.focus({ preventScroll: true });
          var s = getSelection(); s.removeAllRanges(); s.addRange(r); keep();
          return;
        }
        left -= node.data.length;
      }
    }

    // ---------- ステータスバー ----------
    function counts() {
      var text = page ? page.innerText : "", noSp = text.replace(/\s/g, "");
      var words = (text.match(/[A-Za-z0-9’'-]+/g) || []).length + (text.match(/[\u3040-\u30ff\u3400-\u9fff\uf900-\ufaff]/g) || []).length;
      var paras = page ? page.querySelectorAll("p,h1,h2,h3,h4,li,blockquote,pre").length : 0;
      return { chars: noSp.length, withSp: text.replace(/\n/g, "").length, words: words, paras: paras, pages: pages(), lines: text.split("\n").filter(function (x) { return x.trim(); }).length };
    }
    function pages() {
      var paper = el("paper"), p = pageOf(data), d = dims(p);
      if (!paper) return 1;
      return Math.max(1, Math.ceil((p.vert ? paper.scrollWidth : paper.scrollHeight) / ((p.vert ? d.w : d.h) || 1) - 0.02));
    }
    var statTimer = null;
    function stat() {
      clearTimeout(statTimer);
      statTimer = setTimeout(function () {
        if (!page) return;
        var c = counts(), cur = 1, paper = el("paper"), sc = el("doc-scroll"), d = dims(pageOf(data));
        if (paper && sc) cur = Math.min(c.pages, Math.max(1, Math.floor((sc.scrollTop / zoomZ + sc.clientHeight / zoomZ / 3) / d.h) + 1));
        var sel = getSelection(), selN = sel.rangeCount && inPage(sel.anchorNode) ? sel.toString().replace(/\s/g, "").length : 0;
        O.stat("<span>" + esc(t("{n} / {m} ページ", { n: cur, m: c.pages })) + "</span><span>" + esc(selN ? t("{s} / {n} 文字", { s: selN, n: c.chars }) : t("{n} 文字", { n: c.chars })) + "</span>" +
          (readOnly ? "<span>" + esc(t("閲覧モード")) + "</span>" : "") + '<span class="hide-sm">' + esc(t("日本語")) + "</span>");
      }, 120);
    }

    // ---------- ナビゲーション（見出しの一覧） ----------
    function heads() { return page ? Array.prototype.slice.call(page.querySelectorAll("h1,h2,h3")) : []; }
    function drawNav() {
      var box = el("doc-nav");
      if (!box || box.hidden) return;
      var hs = heads();
      box.innerHTML = '<div class="doc-side__head">' + icon("toc") + "<strong>" + esc(t("ナビゲーション")) + '</strong><button type="button" class="icon-btn m3-state" data-cmd="nav" aria-label="' + esc(t("閉じる")) + '">' + icon("close") + "</button></div>" +
        (hs.length ? '<div class="doc-nav__list">' + hs.map(function (h, i) { return '<button type="button" class="doc-nav__item doc-nav__item--' + h.tagName.toLowerCase() + ' m3-state" data-head="' + i + '">' + esc(h.textContent.trim() || t("（空の見出し）")) + "</button>"; }).join("") + "</div>"
          : '<p class="doc-side__empty">' + esc(t("見出しを付けると、ここに一覧が出ます。ホームの「スタイル」から見出しを選んでください。")) + "</p>");
    }

    // ---------- コメント ----------
    function drawCmts() {
      var box = el("doc-cmts");
      if (!box || box.hidden) return;
      var order = Array.prototype.slice.call(page.querySelectorAll(".cmt[data-cmt]")).map(function (x) { return x.getAttribute("data-cmt"); });
      var list = data.comments.slice().sort(function (a, b) { var i = order.indexOf(a.id), j = order.indexOf(b.id); return (i < 0 ? 1e9 : i) - (j < 0 ? 1e9 : j); });
      box.innerHTML = '<div class="doc-side__head">' + icon("comment") + "<strong>" + esc(t("コメント")) + '</strong><button type="button" class="icon-btn m3-state" data-cmd="newComment" aria-label="' + esc(t("新しいコメント")) + '" title="' + esc(t("新しいコメント")) + '">' + icon("add_comment") + '</button><button type="button" class="icon-btn m3-state" data-cmd="comments" aria-label="' + esc(t("閉じる")) + '">' + icon("close") + "</button></div>" +
        (list.length ? list.map(function (c) {
          var mark = page.querySelector('.cmt[data-cmt="' + c.id + '"]');
          return '<div class="cmt-card" data-cmt-card="' + esc(c.id) + '"><div class="cmt-card__top"><span class="cmt-card__who">' + esc(c.author || t("名前なし")) + '</span><span class="cmt-card__when">' + esc(O.when(c.ts)) + "</span>" +
            '<button type="button" class="icon-btn m3-state" data-cmt-del="' + esc(c.id) + '" aria-label="' + esc(t("削除")) + '" title="' + esc(t("削除")) + '">' + icon("delete") + "</button></div>" +
            (mark ? '<div class="cmt-card__quote">' + esc(mark.textContent.slice(0, 60)) + "</div>" : "") + '<div class="cmt-card__text">' + esc(c.text) + "</div></div>";
        }).join("") : '<p class="doc-side__empty">' + esc(t("文字を選んで「新しいコメント」を押すと、コメントを付けられます。")) + "</p>");
    }
    function newComment() {
      if (readOnly) return;
      var r = selRange();
      if (!r || r.collapsed) { O.notify(t("コメントを付ける文字を選んでください")); return; }
      var range = r.cloneRange();
      M3.prompt({ title: t("新しいコメント"), label: t("コメント"), value: "", ok: t("追加"), icon: "add_comment" }).then(function (text) {
        if (!text) return;
        var id = "c" + Date.now().toString(36), span = document.createElement("span");
        span.className = "cmt"; span.setAttribute("data-cmt", id);
        try { range.surroundContents(span); } catch (e) { span.appendChild(range.extractContents()); range.insertNode(span); }
        data.comments.push({ id: id, text: text, author: O.userName() || t("名前なし"), ts: Date.now() });
        showCmts = true;
        el("doc-cmts").hidden = false;
        changed();
        drawCmts();
        O.sync();
      });
    }
    function deleteComment(id) {
      page.querySelectorAll('.cmt[data-cmt="' + id + '"]').forEach(function (s) { s.replaceWith.apply(s, Array.prototype.slice.call(s.childNodes)); });
      data.comments = data.comments.filter(function (c) { return c.id !== id; });
      changed();
      drawCmts();
    }

    // ---------- 表 ----------
    function gridPicker(at) {
      // 表の大きさを、マス目の上でえらぶ（10 × 8 まで。その先はダイアログ）
      return new Promise(function (resolve) {
        O.closeMenu();
        var box = document.createElement("div"), R = 8, C = 10, done = false;
        box.className = "menu";
        box.innerHTML = '<div class="tgrid">' + Array.apply(null, Array(R * C)).map(function (x, i) { return '<button type="button" class="tgrid__c" data-i="' + i + '" aria-label="' + (Math.floor(i / C) + 1) + "×" + (i % C + 1) + '"></button>'; }).join("") + '</div><div class="tgrid__label" id="tgrid-label">' + esc(t("表の大きさを選んでください")) + '</div><div class="menu__sep"></div><button type="button" class="menu__item m3-state" data-more>' + icon("table") + '<span class="menu__label">' + esc(t("表の挿入…")) + "</span></button>";
        document.body.appendChild(box);
        var r = at.getBoundingClientRect();
        box.style.left = Math.min(r.left, window.innerWidth - box.offsetWidth - 8) + "px"; box.style.top = (r.bottom + 4) + "px";
        function finish(v) { if (done) return; done = true; box.remove(); document.removeEventListener("pointerdown", out, true); document.removeEventListener("keydown", key, true); resolve(v); }
        function out(e) { if (!box.contains(e.target)) finish(null); }
        function key(e) { if (e.key === "Escape") { e.preventDefault(); finish(null); } }
        box.addEventListener("mousedown", function (e) { e.preventDefault(); });
        box.addEventListener("pointerover", function (e) {
          var b = e.target.closest("[data-i]"); if (!b) return;
          var i = +b.getAttribute("data-i"), rr = Math.floor(i / C), cc = i % C;
          box.querySelectorAll("[data-i]").forEach(function (x) { var j = +x.getAttribute("data-i"); x.classList.toggle("is-on", Math.floor(j / C) <= rr && j % C <= cc); });
          box.querySelector("#tgrid-label").textContent = t("{r} 行 × {c} 列", { r: rr + 1, c: cc + 1 });
        });
        box.addEventListener("click", function (e) {
          var b = e.target.closest("[data-i]");
          if (b) { var i = +b.getAttribute("data-i"); finish({ r: Math.floor(i / C) + 1, c: i % C + 1 }); return; }
          if (e.target.closest("[data-more]")) {
            finish(null);
            M3.prompt({ title: t("表の挿入"), label: t("行の数 × 列の数（例：5x4）"), value: "3x3", ok: t("挿入") }).then(function (v) {
              var m = /(\d+)\s*[x×*,]\s*(\d+)/i.exec(v || "");
              if (m) resolve({ r: Math.min(200, +m[1]), c: Math.min(30, +m[2]) });
            });
          }
        });
        setTimeout(function () { document.addEventListener("pointerdown", out, true); }, 0);
        document.addEventListener("keydown", key, true);
      });
    }
    function tableHtml(r, c) {
      var h = "<table><tbody>";
      for (var i = 0; i < r; i++) { h += "<tr>"; for (var j = 0; j < c; j++) h += "<td><br></td>"; h += "</tr>"; }
      return h + "</tbody></table><p><br></p>";
    }
    function cellIndex(cell) { return Array.prototype.indexOf.call(cell.parentNode.cells, cell); }
    function tableCmd(cmd, at) {
      var cell = closest("td,th"), tbl = cell && cell.closest("table");
      if (!tbl) { O.notify(t("表の中にカーソルを置いてください")); return; }
      var row = cell.parentNode, ci = cellIndex(cell);
      function blank(tag) { var x = document.createElement(tag || "td"); x.innerHTML = "<br>"; return x; }
      if (cmd === "rowAbove" || cmd === "rowBelow") {
        var nr = document.createElement("tr");
        Array.prototype.forEach.call(row.cells, function (x) { var b = blank(x.tagName === "TH" && row.parentNode.rows[0] !== row ? "th" : "td"); if (x.colSpan > 1) b.colSpan = x.colSpan; nr.appendChild(b); });
        row.parentNode.insertBefore(nr, cmd === "rowAbove" ? row : row.nextSibling);
      } else if (cmd === "colLeft" || cmd === "colRight") {
        Array.prototype.forEach.call(tbl.rows, function (tr) { var ref = tr.cells[Math.min(ci, tr.cells.length - 1)], b = blank(ref && ref.tagName === "TH" ? "th" : "td"); if (!ref) tr.appendChild(b); else tr.insertBefore(b, cmd === "colLeft" ? ref : ref.nextSibling); });
      } else if (cmd === "delRow") {
        if (tbl.rows.length < 2) return tableCmd("delTable");
        row.remove();
      } else if (cmd === "delCol") {
        if (row.cells.length < 2) return tableCmd("delTable");
        Array.prototype.forEach.call(tbl.rows, function (tr) { if (tr.cells[ci]) tr.cells[ci].remove(); });
      } else if (cmd === "delTable") {
        var p = document.createElement("p"); p.innerHTML = "<br>";
        tbl.replaceWith(p);
      } else if (cmd === "headRow") {
        var first = tbl.rows[0], toTh = first.cells[0] && first.cells[0].tagName !== "TH";
        Array.prototype.slice.call(first.cells).forEach(function (x) { var n = document.createElement(toTh ? "th" : "td"); n.innerHTML = x.innerHTML; if (x.colSpan > 1) n.colSpan = x.colSpan; if (x.getAttribute("style")) n.setAttribute("style", x.getAttribute("style")); x.replaceWith(n); });
      } else if (cmd === "shade") {
        O.palette(at, { none: t("色なし"), value: cell.style.backgroundColor }).then(function (v) { if (!v) return; cell.style.backgroundColor = v === "none" ? "" : v; if (!cell.getAttribute("style")) cell.removeAttribute("style"); changed(); });
        return;
      } else if (cmd === "merge") {
        // 右のセルと結合する
        var next = cell.nextElementSibling;
        if (!next) { O.notify(t("右にセルがありません")); return; }
        cell.innerHTML = (cell.textContent.trim() ? cell.innerHTML : "") + (next.textContent.trim() ? " " + next.innerHTML : "") || "<br>";
        cell.colSpan = (cell.colSpan || 1) + (next.colSpan || 1);
        next.remove();
      } else if (cmd === "fit") {
        tbl.style.width = tbl.style.width ? "" : "auto";
      }
      changed();
      O.sync();
    }

    // ---------- 画像 ----------
    function selectImg(img) {
      if (imgSel) imgSel.classList.remove("is-sel");
      imgSel = img;
      if (img) img.classList.add("is-sel");
      placeImgHandle();
      O.sync();
    }
    function placeImgHandle() {
      var h = el("img-rs"), paper = el("paper");
      if (!h) return;
      if (!imgSel || !paper || !paper.contains(imgSel) || readOnly) { h.hidden = true; return; }
      var a = imgSel.getBoundingClientRect(), b = paper.getBoundingClientRect();
      h.hidden = false;
      h.style.left = ((a.right - b.left) / zoomZ - 7) + "px";
      h.style.top = ((a.bottom - b.top) / zoomZ - 7) + "px";
    }
    function imgCmd(cmd, at) {
      var img = imgSel;
      if (!img) return;
      if (cmd.indexOf("imgw:") === 0) { img.style.width = cmd.slice(5) + "%"; img.removeAttribute("width"); img.removeAttribute("height"); }
      else if (cmd === "imgWidth") {
        M3.prompt({ title: t("画像の幅"), label: t("ピクセル"), value: String(Math.round(img.getBoundingClientRect().width / zoomZ)), ok: t("OK") }).then(function (v) {
          var n = Math.round(parseFloat(v)); if (!(n >= 16)) return; img.style.width = n + "px"; changed(); placeImgHandle();
        });
        return;
      } else if (cmd.indexOf("wrap:") === 0) {
        var w = cmd.slice(5);
        img.classList.remove("wrap-l", "wrap-r", "wrap-c");
        if (w !== "inline") img.classList.add("wrap-" + w);
      } else if (cmd === "imgAlt") {
        M3.prompt({ title: t("代替テキスト"), label: t("画像の説明（読み上げに使われます）"), value: img.alt || "", ok: t("OK") }).then(function (v) { if (v === null) return; img.alt = v; changed(); });
        return;
      } else if (cmd === "imgDel") {
        var fig = img.closest("figure.chart") || img;
        fig.remove(); imgSel = null; placeImgHandle(); changed(); O.sync(); return;
      }
      changed();
      placeImgHandle();
    }
    function insertImage(src) {
      if (!src) return;
      insertHtml('<img src="' + esc(src) + '" alt="" style="width:50%">');
    }

    // ---------- グラフ ----------
    function chartDialog(spec) {
      var editing = !!spec;
      spec = spec || { type: "col", title: "", cats: [t("4月"), t("5月"), t("6月"), t("7月")], series: [{ name: t("売上"), values: [120, 150, 135, 180] }] };
      var rows = [[""].concat(spec.series.map(function (s) { return s.name; }))];
      spec.cats.forEach(function (c, i) { rows.push([c].concat(spec.series.map(function (s) { return s.values[i] === null || s.values[i] === undefined ? "" : s.values[i]; }))); });
      var body = document.createElement("div");
      body.innerHTML = '<div class="dlg-grid"><label for="ch-type">' + esc(t("種類")) + '</label><select id="ch-type">' + Chart.TYPES.map(function (x) { return '<option value="' + x.value + '"' + (x.value === spec.type ? " selected" : "") + ">" + esc(t(x.label)) + "</option>"; }).join("") + "</select>" +
        '<label for="ch-title">' + esc(t("タイトル")) + '</label><input id="ch-title" value="' + esc(spec.title || "") + '"></div>' +
        '<p class="dlg-note">' + esc(t("データ（1 行目は系列の名前、1 列目は項目。タブかカンマで区切ります。Sheet からコピーして貼り付けられます）")) + "</p>" +
        '<textarea class="dlg-area" id="ch-data" spellcheck="false">' + esc(rows.map(function (r) { return r.join("\t"); }).join("\n")) + "</textarea>";
      return M3.dialog({ title: editing ? t("グラフのデータを編集") : t("グラフの挿入"), icon: "bar_chart", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: editing ? t("OK") : t("挿入"), primary: true, value: function () {
          var text = body.querySelector("#ch-data").value, sep = text.indexOf("\t") >= 0 ? "\t" : ",";
          var grid = text.split(/\r?\n/).filter(function (l) { return l.trim(); }).map(function (l) { return l.split(sep).map(function (x) { x = x.trim(); var n = parseFloat(x.replace(/[,¥￥%\s]/g, "")); return x !== "" && !isNaN(n) && /^[-+¥￥]?[\d,.]+%?$/.test(x) ? n : x; }); });
          var s = Chart.fromRows(grid);
          s.type = body.querySelector("#ch-type").value; s.title = body.querySelector("#ch-title").value.trim();
          return s;
        } }] });
    }
    function insertChart() {
      keep();
      chartDialog().then(function (spec) { if (spec) insertHtml(chartHtml(spec) + "<p><br></p>"); });
    }
    function editChart(fig) {
      var spec = {};
      try { spec = JSON.parse(fig.getAttribute("data-spec")); } catch (e) { /* 新しく */ }
      chartDialog(spec).then(function (s) { if (!s) return; fig.setAttribute("data-spec", JSON.stringify(s)); fig.innerHTML = Chart.svg(s, { w: 560, h: 320 }); changed(); });
    }

    // ---------- 目次・脚注 ----------
    function tocHtml() {
      var hs = heads().filter(function (h) { return !h.closest(".toc"); });
      hs.forEach(function (h, i) { if (!h.id) h.id = "h-" + Date.now().toString(36) + i; });
      return '<div class="toc" contenteditable="false"><p class="toc__title">' + esc(t("目次")) + "</p>" + (hs.length ? hs.map(function (h) {
        return '<p class="toc__item toc__item--' + h.tagName.toLowerCase() + '"><a href="#' + h.id + '">' + esc(h.textContent.trim()) + "</a></p>";
      }).join("") : '<p class="toc__item">' + esc(t("見出しがありません")) + "</p>") + "</div>";
    }
    function toc() {
      var old = page.querySelector(".toc");
      if (old) { old.outerHTML = tocHtml(); changed(); O.notify(t("目次を更新しました")); return; }
      insertHtml(tocHtml() + "<p><br></p>");
    }
    function footnote() {
      if (readOnly) return;
      keep();
      M3.prompt({ title: t("脚注の挿入"), label: t("脚注の文"), value: "", ok: t("挿入") }).then(function (text) {
        if (!text) return;
        var id = "fn" + Date.now().toString(36);
        back();
        document.execCommand("insertHTML", false, '<sup class="fn-ref" data-fn="' + id + '">*</sup>&#8203;');
        var box = page.querySelector("section.footnotes");
        if (!box) { box = document.createElement("section"); box.className = "footnotes"; box.innerHTML = "<hr><ol></ol>"; page.appendChild(box); }
        var li = document.createElement("li"); li.setAttribute("data-fn", id); li.textContent = text;
        box.querySelector("ol").appendChild(li);
        renumber();
        changed();
      });
    }
    function renumber() {
      var box = page.querySelector("section.footnotes"), ol = box && box.querySelector("ol");
      if (!ol) return;
      var refs = Array.prototype.slice.call(page.querySelectorAll("sup.fn-ref[data-fn]"));
      refs.forEach(function (s, i) { s.textContent = String(i + 1); });
      refs.forEach(function (s) { var li = ol.querySelector('li[data-fn="' + s.getAttribute("data-fn") + '"]'); if (li) ol.appendChild(li); });
      Array.prototype.slice.call(ol.children).forEach(function (li) { if (!page.querySelector('sup.fn-ref[data-fn="' + li.getAttribute("data-fn") + '"]')) li.remove(); });
      if (!ol.children.length) box.remove();
    }

    // ---------- 書式のコピー ----------
    function grabStyle() {
      var a = anchorEl();
      if (!a) return null;
      var cs = getComputedStyle(a.nodeType === 1 ? a : a.parentNode);
      return { b: +cs.fontWeight >= 600, i: cs.fontStyle === "italic", u: /underline/.test(cs.textDecorationLine), s: /line-through/.test(cs.textDecorationLine), color: cs.color, size: cs.fontSize, font: cs.fontFamily, hl: closest("[style*='background-color']") ? closest("[style*='background-color']").style.backgroundColor : "" };
    }
    function applyStyle(st) {
      back();
      document.execCommand("styleWithCSS", false, true);
      ["bold", "italic", "underline", "strikeThrough"].forEach(function (c, i) { var want = [st.b, st.i, st.u, st.s][i]; if (document.queryCommandState(c) !== want) document.execCommand(c); });
      document.execCommand("foreColor", false, st.color);
      document.execCommand("hiliteColor", false, st.hl || "transparent");
      document.execCommand("fontName", false, st.font);
      applySize(parseFloat(st.size) * 0.75);
      changed();
    }
    function applySize(pt) {
      back();
      document.execCommand("styleWithCSS", false, true);
      document.execCommand("fontSize", false, "7");
      page.querySelectorAll('span[style*="xxx-large"], font[size="7"]').forEach(function (s) {
        if (s.tagName === "FONT") { var sp = document.createElement("span"); sp.innerHTML = s.innerHTML; s.replaceWith(sp); s = sp; }
        s.style.fontSize = pt + "pt";
      });
      keep();
    }

    // ---------- 検索と置換 ----------
    function textMap() {
      var w = document.createTreeWalker(page, NodeFilter.SHOW_TEXT), full = "", map = [], n;
      while ((n = w.nextNode())) { map.push({ n: n, s: full.length }); full += n.data; }
      return { full: full, map: map };
    }
    function locate(tm, pos) {
      for (var i = tm.map.length - 1; i >= 0; i--) if (tm.map[i].s <= pos) return { node: tm.map[i].n, off: pos - tm.map[i].s };
      return null;
    }
    function hitsOf(tm, q, o) {
      var hay = o.matchCase ? tm.full : tm.full.toLowerCase(), needle = o.matchCase ? q : q.toLowerCase(), out = [], i = 0;
      while ((i = hay.indexOf(needle, i)) >= 0) { out.push(i); i += needle.length || 1; }
      return out;
    }
    function rangeAt(tm, start, len) {
      var a = locate(tm, start), b = locate(tm, start + len);
      if (!a || !b) return null;
      if (b.off === 0 && start + len > 0) { var p = locate(tm, start + len - 1); b = { node: p.node, off: p.off + 1 }; }
      var r = document.createRange(); r.setStart(a.node, a.off); r.setEnd(b.node, b.off);
      return r;
    }
    function find(q, dir, o) {
      if (!page) return false;
      var tm = textMap(), hits = hitsOf(tm, q, o);
      if (!hits.length) return false;
      var s = getSelection(), cur = 0;
      if (s.rangeCount && inPage(s.anchorNode)) {
        var r0 = document.createRange(); r0.selectNodeContents(page);
        var rr = s.getRangeAt(0);
        r0.setEnd(dir > 0 ? rr.endContainer : rr.startContainer, dir > 0 ? rr.endOffset : rr.startOffset);
        cur = r0.toString().length;
      }
      var i;
      if (dir > 0) { i = hits.findIndex(function (h) { return h >= cur; }); if (i < 0) i = 0; }
      else { i = -1; for (var k = hits.length - 1; k >= 0; k--) if (hits[k] + q.length <= cur && hits[k] < cur) { i = k; break; } if (i < 0) i = hits.length - 1; }
      var range = rangeAt(tm, hits[i], q.length);
      if (!range) return false;
      s.removeAllRanges(); s.addRange(range); keep();
      var rect = range.getBoundingClientRect(), sc = el("doc-scroll"), b = sc.getBoundingClientRect();
      if (rect.top < b.top + 40 || rect.bottom > b.bottom - 40) sc.scrollTop += rect.top - b.top - sc.clientHeight / 3;
      return t("{n} 件中 {i} 件目", { n: hits.length, i: i + 1 });
    }
    function replace(q, rep, all, o) {
      if (!page || readOnly) return 0;
      if (!all) {
        var s = getSelection(), txt = s.rangeCount && inPage(s.anchorNode) ? s.toString() : "";
        if ((o.matchCase ? txt : txt.toLowerCase()) !== (o.matchCase ? q : q.toLowerCase())) { find(q, 1, o); return 0; }
        document.execCommand("insertText", false, rep);
        sync();
        return 1;
      }
      var tm = textMap(), hits = hitsOf(tm, q, o);
      for (var i = hits.length - 1; i >= 0; i--) {
        var r = rangeAt(tm, hits[i], q.length);
        if (!r) continue;
        r.deleteContents();
        if (rep) r.insertNode(document.createTextNode(rep));
      }
      page.normalize();
      sync();
      return hits.length;
    }

    // ---------- クリック・入力 ----------
    function wire() {
      var paper = el("paper"), sc = el("doc-scroll");
      page.addEventListener("input", function () { if (imgSel && !page.contains(imgSel)) selectImg(null); changed(true); });
      page.addEventListener("keyup", keep);
      page.addEventListener("mouseup", function () {
        keep();
        if (painter) { var st = painter; painter = null; var r = selRange(); if (r && !r.collapsed) applyStyle(st); O.sync(); }
      });
      page.addEventListener("focus", function () { O.syncSoon(); });
      page.addEventListener("click", function (e) {
        var img = e.target.closest("img, figure.chart");
        if (img && inPage(img)) { selectImg(img.tagName === "FIGURE" ? img : img); return; }
        if (imgSel) selectImg(null);
        var a = e.target.closest("a[href]");
        if (a && (e.ctrlKey || e.metaKey || readOnly)) {
          var href = a.getAttribute("href");
          if (href.charAt(0) === "#") { var target = page.querySelector(href.replace(/[^#\w-]/g, "")); if (target) target.scrollIntoView({ block: "start", behavior: "smooth" }); }
          else window.open(href, "_blank", "noopener");
          e.preventDefault();
        }
        var sup = e.target.closest("sup.fn-ref");
        if (sup) { var li = page.querySelector('section.footnotes li[data-fn="' + sup.getAttribute("data-fn") + '"]'); if (li) li.scrollIntoView({ block: "center", behavior: "smooth" }); }
        var cm = e.target.closest(".cmt[data-cmt]");
        if (cm && showCmts) { var card = el("doc-cmts").querySelector('[data-cmt-card="' + cm.getAttribute("data-cmt") + '"]'); if (card) { card.scrollIntoView({ block: "nearest" }); card.classList.add("is-on"); setTimeout(function () { card.classList.remove("is-on"); }, 1200); } }
      });
      page.addEventListener("dblclick", function (e) { var f = e.target.closest("figure.chart"); if (f && !readOnly) editChart(f); });
      page.addEventListener("paste", function (e) {
        if (readOnly) return;
        var cd = e.clipboardData, files = cd && Array.prototype.filter.call(cd.files || [], function (f) { return /^image\//.test(f.type); });
        if (files && files.length) { e.preventDefault(); O.imageData(files[0]).then(insertImage); return; }
        var html = cd && cd.getData("text/html");
        if (html) {
          e.preventDefault();
          var body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
          document.execCommand("insertHTML", false, sanitize((body ? body[1] : html).replace(/<!--StartFragment-->|<!--EndFragment-->/g, "")));
          changed();
        }
      });
      page.addEventListener("dragover", function (e) { if (e.dataTransfer && Array.prototype.some.call(e.dataTransfer.items || [], function (i) { return i.kind === "file"; })) e.preventDefault(); });
      page.addEventListener("drop", function (e) {
        var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (!f || !/^image\//.test(f.type) || readOnly) return;
        e.preventDefault();
        var pos = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
        if (pos) { var s = getSelection(); s.removeAllRanges(); s.addRange(pos); keep(); }
        O.imageData(f).then(insertImage);
      });
      page.addEventListener("keydown", function (e) {
        if (e.isComposing) return;
        if (e.key === "Tab") {
          var cell = closest("td,th");
          if (cell) {
            e.preventDefault();
            var cells = Array.prototype.slice.call(cell.closest("table").querySelectorAll("td,th")), i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
            if (i >= cells.length && !e.shiftKey) { tableCmd("rowBelow"); cells = Array.prototype.slice.call(cell.closest("table").querySelectorAll("td,th")); }
            var next = cells[Math.max(0, i)];
            if (next) { var r = document.createRange(); r.selectNodeContents(next); var s = getSelection(); s.removeAllRanges(); s.addRange(r); keep(); }
            return;
          }
          if (closest("li")) { e.preventDefault(); exec(e.shiftKey ? "outdent" : "indent"); return; }
          e.preventDefault();
          exec("insertText", "\u3000");
          return;
        }
        if ((e.key === "Delete" || e.key === "Backspace") && imgSel) { e.preventDefault(); imgCmd("imgDel"); return; }
      });
      paper.addEventListener("click", function (e) { if (e.target === paper) { back(); } });
      el("doc-head").addEventListener("input", function (e) { pageOf(data).header = e.target.value; O.changed(true); });
      el("doc-foot").addEventListener("input", function (e) { pageOf(data).footer = e.target.value; O.changed(true); });
      sc.addEventListener("scroll", function () { placeImgHandle(); stat(); });
      el("img-rs").addEventListener("pointerdown", function (e) {
        if (!imgSel) return;
        e.preventDefault();
        var x0 = e.clientX, w0 = imgSel.getBoundingClientRect().width / zoomZ, img = imgSel;
        function mv(ev) { img.style.width = Math.max(24, Math.round(w0 + (ev.clientX - x0) / zoomZ)) + "px"; img.removeAttribute("width"); img.removeAttribute("height"); placeImgHandle(); }
        function up() { window.removeEventListener("pointermove", mv); window.removeEventListener("pointerup", up); changed(); }
        window.addEventListener("pointermove", mv);
        window.addEventListener("pointerup", up);
      });
      el("doc-wrap").addEventListener("click", function (e) {
        var h = e.target.closest("[data-head]");
        if (h) { var target = heads()[+h.getAttribute("data-head")]; if (target) { target.scrollIntoView({ block: "start", behavior: "smooth" }); var r = document.createRange(); r.selectNodeContents(target); r.collapse(true); page.focus({ preventScroll: true }); var s = getSelection(); s.removeAllRanges(); s.addRange(r); keep(); } return; }
        var del = e.target.closest("[data-cmt-del]");
        if (del) { deleteComment(del.getAttribute("data-cmt-del")); return; }
        var card = e.target.closest("[data-cmt-card]");
        if (card) { var m = page.querySelector('.cmt[data-cmt="' + card.getAttribute("data-cmt-card") + '"]'); if (m) { m.scrollIntoView({ block: "center", behavior: "smooth" }); m.classList.add("is-on"); setTimeout(function () { m.classList.remove("is-on"); }, 1400); } }
      });
    }
    document.addEventListener("selectionchange", function () { var s = getSelection(); if (page && s.rangeCount && inPage(s.anchorNode)) { keep(); stat(); } });

    // ---------- リボンのボタン ----------
    var SYMBOLS = ["※", "・", "〒", "〇", "◎", "△", "□", "■", "◆", "★", "☆", "→", "←", "↑", "↓", "①", "②", "③", "④", "⑤", "℃", "￥", "§", "¶", "©", "®", "™", "±", "×", "÷", "≠", "≦", "≧", "∞", "♪", "々", "〃", "〆"];
    function command(cmd, at) {
      if (!page) return;
      var ro = readOnly && !/^(nav|comments|count|toc:update)$/.test(cmd);
      if (ro) { O.notify(t("閲覧モードでは書きかえられません")); return; }
      if (cmd.indexOf("fmt:") === 0) return exec("formatBlock", cmd.slice(4));
      if (cmd.indexOf("img") === 0 || cmd.indexOf("wrap:") === 0) return imgCmd(cmd, at);
      if (cmd.indexOf("tbl:") === 0) return tableCmd(cmd.slice(4), at);
      var p = pageOf(data);
      switch (cmd) {
        case "bold": case "italic": case "underline": case "strikeThrough": case "subscript": case "superscript":
        case "insertUnorderedList": case "insertOrderedList": case "indent": case "outdent":
        case "justifyLeft": case "justifyCenter": case "justifyRight": case "justifyFull":
          return exec(cmd);
        case "clear": back(); document.execCommand("removeFormat"); document.execCommand("hiliteColor", false, "transparent"); document.execCommand("unlink"); return changed();
        case "fontName":
          keep();
          return O.menu(at, O.fonts.map(function (f) { return { value: f[0], label: f[0], style: "font-family:" + f[1], checked: state().font === f[0] }; })).then(function (v) { if (v) exec("fontName", O.fontCss(v)); });
        case "fontSize":
          keep();
          return O.menu(at, O.sizes.map(function (s) { return { value: String(s), label: String(s), checked: state().size === String(s) }; })).then(function (v) { if (v) { applySize(+v); changed(); } });
        case "grow": case "shrink": {
          var cur = parseFloat(state().size) || 12, list = O.sizes, i = 0;
          if (cmd === "grow") { while (i < list.length && list[i] <= cur) i++; i = Math.min(i, list.length - 1); } else { i = list.length - 1; while (i > 0 && list[i] >= cur) i--; }
          applySize(list[i]); return changed();
        }
        case "color":
          keep();
          return O.palette(at, { none: t("自動"), value: "" }).then(function (v) { if (!v) return; if (v === "none") { back(); document.execCommand("styleWithCSS", false, true); document.execCommand("foreColor", false, "#1a1c20"); changed(); } else exec("foreColor", v); });
        case "highlight":
          keep();
          return O.palette(at, { none: t("色なし"), value: "" }).then(function (v) { if (v) exec("hiliteColor", v === "none" ? "transparent" : v); });
        case "spacing":
          keep();
          return O.menu(at, ["1.0", "1.15", "1.5", "1.85", "2.0", "2.5", "3.0"].map(function (x) { return { value: x, label: x === "1.85" ? x + "（" + t("標準") + "）" : x }; })).then(function (v) {
            if (!v) return;
            back();
            blocks().forEach(function (b) { b.style.lineHeight = v === "1.85" ? "" : v; if (!b.getAttribute("style")) b.removeAttribute("style"); });
            changed();
          });
        case "selectAll": back(); document.execCommand("selectAll"); keep(); return;
        case "pagebreak": return insertHtml('<hr class="pagebreak"><p><br></p>');
        case "table": keep(); return gridPicker(at).then(function (g) { if (g) insertHtml(tableHtml(g.r, g.c)); });
        case "image": keep(); return O.pickImage().then(insertImage);
        case "chart": return insertChart();
        case "link": {
          keep();
          var a = closest("a[href]");
          return M3.prompt({ title: t("リンク"), label: t("アドレス（URL）"), value: a ? a.getAttribute("href") : "", placeholder: "https://", ok: t("OK") }).then(function (url) {
            if (url === null) return;
            url = url.trim();
            if (!url) { if (a) { back(); a.replaceWith.apply(a, Array.prototype.slice.call(a.childNodes)); changed(); } return; }
            if (!/^(https?:|mailto:|#)/i.test(url)) url = "https://" + url;
            if (a) { a.setAttribute("href", url); changed(); return; }
            back();
            var r = selRange();
            if (r && !r.collapsed) exec("createLink", url);
            else insertHtml('<a href="' + esc(url) + '" target="_blank" rel="noopener">' + esc(url) + "</a>&nbsp;");
          });
        }
        case "hr": return exec("insertHorizontalRule");
        case "date": { var d = new Date(); return exec("insertText", t("{y}年{m}月{d}日", { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() })); }
        case "symbol":
          keep();
          return O.menu(at, SYMBOLS.map(function (s) { return { value: s, label: s }; })).then(function (v) { if (v) exec("insertText", v); });
        case "toc": return toc();
        case "footnote": return footnote();
        case "newComment": return newComment();
        case "comments": showCmts = !showCmts; el("doc-cmts").hidden = !showCmts; drawCmts(); O.sync(); return;
        case "delComments":
          if (!data.comments.length) { O.notify(t("コメントはありません")); return; }
          return M3.confirm({ title: t("すべてのコメントを削除しますか？"), text: t("{n} 件のコメントを削除します。", { n: data.comments.length }), ok: t("削除"), danger: true, icon: "delete" }).then(function (ok) {
            if (!ok) return;
            data.comments.slice().forEach(function (c) { deleteComment(c.id); });
          });
        case "nav": showNav = !showNav; el("doc-nav").hidden = !showNav; drawNav(); O.sync(); return;
        case "spell": page.spellcheck = !page.spellcheck; O.notify(page.spellcheck ? t("スペル チェックをオンにしました（ブラウザの機能を使います）") : t("スペル チェックをオフにしました")); O.sync(); return;
        case "count": {
          var c = counts(), body = document.createElement("dl");
          body.className = "dlg-stats";
          body.innerHTML = [[t("ページ数"), c.pages], [t("文字数（スペースを含めない）"), c.chars], [t("文字数（スペースを含める）"), c.withSp], [t("単語数"), c.words], [t("段落数"), c.paras], [t("行数"), c.lines]]
            .map(function (x) { return "<dt>" + esc(x[0]) + "</dt><dd>" + Number(x[1]).toLocaleString("ja-JP") + "</dd>"; }).join("");
          return M3.dialog({ title: t("文字カウント"), icon: "123", body: body, actions: [{ label: t("閉じる"), value: null, primary: true }] });
        }
        case "margin":
          return O.menu(at, Object.keys(MARGINS).map(function (k) { return { value: k, label: MARGINS[k].label, checked: p.margin === k }; })).then(function (v) { if (v) { p.margin = v; layout(); O.changed(); stat(); } });
        case "orient":
          return O.menu(at, [{ value: "portrait", label: t("縦"), icon: "crop_portrait", checked: p.orient === "portrait" }, { value: "landscape", label: t("横"), icon: "crop_landscape", checked: p.orient === "landscape" }]).then(function (v) { if (v) { p.orient = v; layout(); O.changed(); stat(); } });
        case "size":
          return O.menu(at, Object.keys(SIZES).map(function (k) { var s = SIZES[k]; return { value: k, label: s.label, sub: Math.round(s.tw / 56.7) + " × " + Math.round(s.th / 56.7) + " mm", checked: p.size === k }; })).then(function (v) { if (v) { p.size = v; layout(); O.changed(); stat(); } });
        case "cols":
          return O.menu(at, [1, 2, 3].map(function (n) { return { value: String(n), label: t("{n} 段", { n: n }), checked: p.cols === n }; })).then(function (v) { if (v) { p.cols = +v; layout(); O.changed(); } });
        case "vert": p.vert = !p.vert; layout(); O.changed(); O.sync(); stat(); return;
        case "header": el("doc-head").focus(); return;
        case "footer": el("doc-foot").focus(); return;
        case "pageNum": p.num = !p.num; layout(); O.changed(); O.sync(); O.notify(p.num ? t("印刷するときに、ページの下に番号を付けます") : t("ページ番号をやめました")); return;
        case "painter":
          if (painter) { painter = null; O.sync(); return; }
          painter = grabStyle();
          O.sync();
          if (painter) O.notify(t("書式をはり付ける文字を選んでください"));
          return;
      }
    }

    // ---------- リボン ----------
    var R = O.groups;
    var ribbon = [
      { tab: t("ホーム"), groups: [R.clip,
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "M PLUS 1", w: 132, cmd: "fontName", key: "font", label: t("フォント") }, { chip: "12", w: 56, cmd: "fontSize", key: "size", label: t("フォント サイズ") },
           { icon: "text_increase", label: t("フォント サイズの拡大"), cmd: "grow" }, { icon: "text_decrease", label: t("フォント サイズの縮小"), cmd: "shrink" }, { icon: "format_clear", label: t("すべての書式をクリア"), cmd: "clear" }],
          [{ icon: "format_bold", label: t("太字"), sub: "Ctrl+B", cmd: "bold", toggle: "bold" }, { icon: "format_italic", label: t("斜体"), sub: "Ctrl+I", cmd: "italic", toggle: "italic" },
           { icon: "format_underlined", label: t("下線"), sub: "Ctrl+U", cmd: "underline", toggle: "underline" }, { icon: "format_strikethrough", label: t("取り消し線"), cmd: "strikeThrough", toggle: "strikeThrough" },
           { icon: "subscript", label: t("下付き"), cmd: "subscript", toggle: "subscript" }, { icon: "superscript", label: t("上付き"), cmd: "superscript", toggle: "superscript" },
           { icon: "ink_highlighter", label: t("蛍光ペンの色"), cmd: "highlight", drop: 1 }, { icon: "format_color_text", label: t("フォントの色"), cmd: "color", drop: 1 }]
        ] }] },
        { label: t("段落"), items: [{ rows: [
          [{ icon: "format_list_bulleted", label: t("箇条書き"), cmd: "insertUnorderedList", toggle: "insertUnorderedList" }, { icon: "format_list_numbered", label: t("段落番号"), cmd: "insertOrderedList", toggle: "insertOrderedList" },
           { icon: "format_indent_decrease", label: t("インデントを減らす"), cmd: "outdent" }, { icon: "format_indent_increase", label: t("インデントを増やす"), cmd: "indent" }, { icon: "format_line_spacing", label: t("行と段落の間隔"), cmd: "spacing", drop: 1 }],
          [{ icon: "format_align_left", label: t("左揃え"), sub: "Ctrl+L", cmd: "justifyLeft", toggle: "justifyLeft" }, { icon: "format_align_center", label: t("中央揃え"), sub: "Ctrl+E", cmd: "justifyCenter", toggle: "justifyCenter" },
           { icon: "format_align_right", label: t("右揃え"), sub: "Ctrl+R", cmd: "justifyRight", toggle: "justifyRight" }, { icon: "format_align_justify", label: t("両端揃え"), sub: "Ctrl+J", cmd: "justifyFull", toggle: "justifyFull" }]
        ] }] },
        { label: t("スタイル"), items: [{ styles: [{ tag: "p", label: t("標準") }, { tag: "h1", label: t("表題") }, { tag: "h2", label: t("見出し 1") }, { tag: "h3", label: t("見出し 2") }, { tag: "blockquote", label: t("引用") }] }] },
        { label: t("編集"), items: [{ stack: [{ icon: "search", label: t("検索"), cmd: "find" }, { icon: "find_replace", label: t("置換"), cmd: "replace" }, { icon: "select_all", label: t("すべて選択"), cmd: "selectAll" }] }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("ページ"), items: [{ big: 1, icon: "insert_page_break", label: t("改ページ"), cmd: "pagebreak" }] },
        { label: t("表"), items: [{ big: 1, icon: "table", label: t("表"), cmd: "table", drop: 1 }] },
        { label: t("図"), items: [{ big: 1, icon: "image", label: t("画像"), cmd: "image" }, { big: 1, icon: "bar_chart", label: t("グラフ"), cmd: "chart" }] },
        { label: t("リンク"), items: [{ big: 1, icon: "link", label: t("リンク"), sub: "Ctrl+K", cmd: "link" }] },
        { label: t("コメント"), items: [{ big: 1, icon: "add_comment", label: t("コメント"), cmd: "newComment" }] },
        { label: t("ヘッダーとフッター"), items: [{ stack: [{ icon: "vertical_align_top", label: t("ヘッダー"), cmd: "header" }, { icon: "vertical_align_bottom", label: t("フッター"), cmd: "footer" }, { icon: "tag", label: t("ページ番号"), cmd: "pageNum", toggle: "pageNum" }] }] },
        { label: t("テキスト"), items: [{ stack: [{ icon: "horizontal_rule", label: t("横線"), cmd: "hr" }, { icon: "today", label: t("日付"), cmd: "date" }, { icon: "emoji_symbols", label: t("記号"), cmd: "symbol", drop: 1 }] }] }
      ] },
      { tab: t("レイアウト"), groups: [
        { label: t("ページ設定"), items: [{ big: 1, icon: "margin", label: t("余白"), cmd: "margin", drop: 1 }, { big: 1, icon: "crop_rotate", label: t("印刷の向き"), cmd: "orient", drop: 1 },
          { big: 1, icon: "description", label: t("サイズ"), cmd: "size", drop: 1 }, { big: 1, icon: "view_column", label: t("段組み"), cmd: "cols", drop: 1 }, { big: 1, icon: "text_rotate_vertical", label: t("縦書き"), cmd: "vert", toggle: "vert" }] },
        { label: t("段落"), items: [{ stack: [{ icon: "format_line_spacing", label: t("行間"), cmd: "spacing", drop: 1 }, { icon: "format_indent_increase", label: t("インデント"), cmd: "indent" }, { icon: "insert_page_break", label: t("改ページ"), cmd: "pagebreak" }] }] }
      ] },
      { tab: t("参考資料"), groups: [
        { label: t("目次"), items: [{ big: 1, icon: "toc", label: t("目次"), cmd: "toc" }] },
        { label: t("脚注"), items: [{ big: 1, icon: "notes", label: t("脚注の挿入"), cmd: "footnote" }] }
      ] },
      { tab: t("校閲"), groups: [
        { label: t("文章校正"), items: [{ big: 1, icon: "spellcheck", label: t("スペル チェック"), cmd: "spell", toggle: "spell" }, { big: 1, icon: "123", label: t("文字カウント"), cmd: "count" }] },
        { label: t("コメント"), items: [{ big: 1, icon: "add_comment", label: t("新しいコメント"), cmd: "newComment" }, { big: 1, icon: "forum", label: t("コメントの表示"), cmd: "comments", toggle: "cmts" },
          { stack: [{ icon: "delete_sweep", label: t("すべて削除"), cmd: "delComments" }] }] }
      ] },
      { tab: t("表示"), groups: [R.view,
        { label: t("表示"), items: [{ big: 1, icon: "toc", label: t("ナビゲーション"), cmd: "nav", toggle: "nav" }] },
        R.zoom] },
      { tab: t("表のレイアウト"), ctx: 1, when: function () { return !!closest("td,th") && !imgSel; }, groups: [
        { label: t("行と列"), items: [{ big: 1, icon: "add_row_above", label: t("上に挿入"), cmd: "tbl:rowAbove" }, { big: 1, icon: "add_row_below", label: t("下に挿入"), cmd: "tbl:rowBelow" },
          { big: 1, icon: "add_column_left", label: t("左に挿入"), cmd: "tbl:colLeft" }, { big: 1, icon: "add_column_right", label: t("右に挿入"), cmd: "tbl:colRight" }] },
        { label: t("削除"), items: [{ stack: [{ icon: "delete", label: t("行の削除"), cmd: "tbl:delRow" }, { icon: "delete", label: t("列の削除"), cmd: "tbl:delCol" }, { icon: "delete_forever", label: t("表の削除"), cmd: "tbl:delTable" }] }] },
        { label: t("書式"), items: [{ stack: [{ icon: "table_rows", label: t("見出しの行"), cmd: "tbl:headRow" }, { icon: "format_color_fill", label: t("塗りつぶし"), cmd: "tbl:shade", drop: 1 }, { icon: "cell_merge", label: t("右のセルと結合"), cmd: "tbl:merge" }] }] }
      ] },
      { tab: t("図の形式"), ctx: 1, when: function () { return !!imgSel; }, groups: [
        { label: t("サイズ"), items: [{ rows: [[{ chip: "25%", w: 52, cmd: "imgw:25" }, { chip: "50%", w: 52, cmd: "imgw:50" }], [{ chip: "75%", w: 52, cmd: "imgw:75" }, { chip: "100%", w: 52, cmd: "imgw:100" }]] }, { big: 1, icon: "width", label: t("幅を指定"), cmd: "imgWidth" }] },
        { label: t("配置"), items: [{ big: 1, icon: "format_image_left", label: t("左に回り込み"), cmd: "wrap:l" }, { big: 1, icon: "format_image_right", label: t("右に回り込み"), cmd: "wrap:r" },
          { big: 1, icon: "align_horizontal_center", label: t("中央"), cmd: "wrap:c" }, { big: 1, icon: "wrap_text", label: t("行内"), cmd: "wrap:inline" }] },
        { label: t("アクセシビリティ"), items: [{ big: 1, icon: "description", label: t("代替テキスト"), cmd: "imgAlt" }] },
        { label: t("削除"), items: [{ big: 1, icon: "delete", label: t("削除"), cmd: "imgDel" }] }
      ] }
    ];
    function fontLabel(ff) {
      var first = String(ff || "").split(",")[0].replace(/["']/g, "").trim();
      for (var i = 0; i < O.fonts.length; i++) { var f = O.fonts[i][1].split(",")[0].replace(/["']/g, "").trim(); if (f === first) return O.fonts[i][0]; }
      return first || "M PLUS 1";
    }
    function state() {
      if (!page) return {};
      var s = { cmts: showCmts, nav: showNav, spell: page.spellcheck, vert: !!pageOf(data).vert, pageNum: !!pageOf(data).num, painter: !!painter };
      var a = anchorEl();
      ["bold", "italic", "underline", "strikeThrough", "subscript", "superscript", "insertUnorderedList", "insertOrderedList", "justifyLeft", "justifyCenter", "justifyRight", "justifyFull"].forEach(function (c) {
        try { s[c] = !!a && document.queryCommandState(c); } catch (e) { s[c] = false; }
      });
      if (a) {
        var cs = getComputedStyle(a.nodeType === 1 ? a : a.parentNode);
        s.font = fontLabel(cs.fontFamily);
        s.size = String(Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2);
        var b = a.closest && a.closest("h1,h2,h3,blockquote,p");
        ["p", "h1", "h2", "h3", "blockquote"].forEach(function (tag) { s["fmt:" + tag] = !!b && b.tagName.toLowerCase() === tag; });
      }
      return s;
    }

    // ---------- 書き出し（.docx・.html・.md・.txt） ----------
    function b64(u8) { var s = ""; for (var i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); }
    function dataUrlBytes(url) {
      var m = /^data:([^;,]+)(;base64)?,(.*)$/.exec(url);
      if (!m) return null;
      var bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]), u8 = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      return { mime: m[1], bytes: u8 };
    }
    // SVG や読めない形の画像は PNG にする
    function toPng(src, w, h) {
      return new Promise(function (resolve) {
        var img = new Image();
        img.onload = function () {
          var cv = document.createElement("canvas"), k = 2;
          cv.width = Math.max(1, Math.round((w || img.naturalWidth || 400) * k)); cv.height = Math.max(1, Math.round((h || img.naturalHeight || 300) * k));
          var g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0, cv.width, cv.height);
          resolve(cv.toDataURL("image/png"));
        };
        img.onerror = function () { resolve(null); };
        img.src = src;
      });
    }
    var X = function (s) { return Zip.esc(s); };
    function docxFiles() {
      var p = pageOf(data), sz = SIZES[p.size], land = p.orient === "landscape", m = MARGINS[p.margin];
      var holder = document.createElement("div");
      holder.innerHTML = data.html;
      // 画像を集める（グラフ・SVG は PNG にする）
      var jobs = [], media = [];
      holder.querySelectorAll("figure.chart").forEach(function (f) {
        var svg = f.querySelector("svg"), img = document.createElement("img");
        img.setAttribute("data-w", "560"); img.setAttribute("data-h", "320");
        jobs.push(toPng("data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg ? svg.outerHTML : ""), 560, 320).then(function (src) { if (src) img.src = src; }));
        var para = document.createElement("p"); para.appendChild(img);
        f.replaceWith(para);
      });
      var live = page ? Array.prototype.slice.call(page.querySelectorAll("img")) : [];
      holder.querySelectorAll("img").forEach(function (img, i) {
        var lv = live[i], w = lv ? lv.getBoundingClientRect().width / zoomZ : 0, h = lv ? lv.getBoundingClientRect().height / zoomZ : 0;
        if (w && !img.getAttribute("data-w")) { img.setAttribute("data-w", Math.round(w)); img.setAttribute("data-h", Math.round(h)); }
        var src = img.getAttribute("src") || "";
        if (/^data:image\/(png|jpe?g|gif)/i.test(src)) return;
        if (/^data:image\//i.test(src)) jobs.push(toPng(src, +img.getAttribute("data-w") || 0, +img.getAttribute("data-h") || 0).then(function (s) { if (s) img.src = s; }));
      });
      return Promise.all(jobs).then(function () {
        var rels = [], relId = 10, numDefs = [], bodyXml = [], olCount = 0;
        function rel(type, target, ext) { var id = "rId" + (++relId); rels.push('<Relationship Id="' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + type + '" Target="' + X(target) + '"' + (ext ? ' TargetMode="External"' : "") + "/>"); return id; }
        function color(c) {
          if (!c) return null;
          var m2 = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c);
          if (m2) { if (m2[4] !== undefined && +m2[4] === 0) return null; return [m2[1], m2[2], m2[3]].map(function (x) { return ("0" + (+x).toString(16)).slice(-2); }).join("").toUpperCase(); }
          var h = /^#([0-9a-f]{6})$/i.exec(c); if (h) return h[1].toUpperCase();
          h = /^#([0-9a-f]{3})$/i.exec(c); if (h) return h[1].split("").map(function (x) { return x + x; }).join("").toUpperCase();
          return null;
        }
        function runProps(st) {
          var x = "";
          if (st.font) { var f = X(st.font); x += '<w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '" w:eastAsia="' + f + '"/>'; }
          if (st.b) x += "<w:b/>";
          if (st.i) x += "<w:i/>";
          if (st.s) x += "<w:strike/>";
          if (st.code) x += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>';
          if (st.color) x += '<w:color w:val="' + st.color + '"/>';
          if (st.size) x += '<w:sz w:val="' + Math.round(st.size * 2) + '"/><w:szCs w:val="' + Math.round(st.size * 2) + '"/>';
          if (st.hl) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + st.hl + '"/>';
          if (st.u || st.link) x += '<w:u w:val="single"/>';
          if (st.link && !st.color) x += '<w:color w:val="1D4ED8"/>';
          if (st.sub) x += '<w:vertAlign w:val="subscript"/>';
          if (st.sup) x += '<w:vertAlign w:val="superscript"/>';
          return x ? "<w:rPr>" + x + "</w:rPr>" : "";
        }
        function textRun(text, st) {
          return text.split("\t").map(function (part, i) { return (i ? "<w:r>" + runProps(st) + "<w:tab/></w:r>" : "") + (part ? "<w:r>" + runProps(st) + '<w:t xml:space="preserve">' + X(part) + "</w:t></w:r>" : ""); }).join("");
        }
        function imgRun(img) {
          var src = img.getAttribute("src") || "", d = dataUrlBytes(src);
          if (!d) return "";
          var ext = /png/.test(d.mime) ? "png" : /gif/.test(d.mime) ? "gif" : "jpeg";
          var name = "image" + (media.length + 1) + "." + ext;
          media.push({ name: name, bytes: d.bytes });
          var id = rel("image", "media/" + name);
          var maxW = (land ? sz.h : sz.w) - m.x * 2;
          var w = Math.min(+img.getAttribute("data-w") || 300, maxW), h = (+img.getAttribute("data-h") || 200) * (w / (+img.getAttribute("data-w") || 300));
          var cx = Math.round(w * 9525), cy = Math.round(h * 9525), n = media.length;
          return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + n + '" name="Picture ' + n + '" descr="' + X(img.getAttribute("alt") || "") + '"/>' +
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
            '<pic:nvPicPr><pic:cNvPr id="' + n + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + id + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
        }
        function inline(node, st) {
          var out = "";
          Array.prototype.forEach.call(node.childNodes, function (n) {
            if (n.nodeType === 3) { var tx = n.data.replace(/\u200b/g, "").replace(/[\r\n]+/g, " "); if (tx) out += textRun(tx, st); return; }
            if (n.nodeType !== 1) return;
            var tag = n.tagName, s = Object.assign({}, st), css = n.style || {};
            if (tag === "BR") { out += "<w:r><w:br/></w:r>"; return; }
            if (tag === "IMG") { out += imgRun(n); return; }
            if (/^(B|STRONG)$/.test(tag)) s.b = 1;
            if (/^(I|EM)$/.test(tag)) s.i = 1;
            if (tag === "U" || tag === "INS") s.u = 1;
            if (/^(S|STRIKE|DEL)$/.test(tag)) s.s = 1;
            if (tag === "SUB") s.sub = 1;
            if (tag === "SUP") s.sup = 1;
            if (tag === "CODE") s.code = 1;
            if (tag === "MARK") s.hl = "FFF59D";
            if (css.fontWeight && (css.fontWeight === "bold" || +css.fontWeight >= 600)) s.b = 1;
            if (css.fontStyle === "italic") s.i = 1;
            if (/underline/.test(css.textDecoration || css.textDecorationLine || "")) s.u = 1;
            if (/line-through/.test(css.textDecoration || css.textDecorationLine || "")) s.s = 1;
            if (color(css.color)) s.color = color(css.color);
            if (color(css.backgroundColor)) s.hl = color(css.backgroundColor);
            if (css.fontSize) { var fs = parseFloat(css.fontSize); if (/pt$/.test(css.fontSize)) s.size = fs; else if (/px$/.test(css.fontSize)) s.size = fs * 0.75; }
            if (css.fontFamily) s.font = css.fontFamily.split(",")[0].replace(/["']/g, "").trim();
            if (tag === "A" && n.getAttribute("href")) {
              var href = n.getAttribute("href");
              s.link = 1;
              if (href.charAt(0) === "#") { out += inline(n, s); return; }
              out += '<w:hyperlink r:id="' + rel("hyperlink", href, true) + '">' + inline(n, s) + "</w:hyperlink>";
              return;
            }
            out += inline(n, s);
          });
          return out;
        }
        function paraProps(n, extra) {
          var x = extra || "", css = n.style || {}, al = css.textAlign;
          if (al) x += '<w:jc w:val="' + ({ center: "center", right: "right", end: "right", justify: "both", left: "left", start: "left" }[al] || "left") + '"/>';
          if (css.lineHeight && parseFloat(css.lineHeight)) x += '<w:spacing w:line="' + Math.round(parseFloat(css.lineHeight) * 240) + '" w:lineRule="auto"/>';
          return x ? "<w:pPr>" + x + "</w:pPr>" : "";
        }
        function para(n, style, extra) {
          var pp = (style ? '<w:pStyle w:val="' + style + '"/>' : "") + (extra || "");
          return "<w:p>" + paraProps(n, pp) + inline(n, {}) + "</w:p>";
        }
        function list(n, lvl) {
          var ordered = n.tagName === "OL", numId = ordered ? 10 + (++olCount) : 1, out = "";
          if (ordered) numDefs.push('<w:num w:numId="' + numId + '"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>');
          Array.prototype.forEach.call(n.children, function (li) {
            if (li.tagName !== "LI") return;
            var own = li.cloneNode(true);
            own.querySelectorAll("ul,ol").forEach(function (x) { x.remove(); });
            out += para(own, "ListParagraph", '<w:numPr><w:ilvl w:val="' + Math.min(lvl, 8) + '"/><w:numId w:val="' + numId + '"/></w:numPr>');
            Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") out += list(c, lvl + 1); });
          });
          return out;
        }
        function table(n) {
          var rows = Array.prototype.slice.call(n.rows), cols = 0;
          rows.forEach(function (r) { var k = 0; Array.prototype.forEach.call(r.cells, function (c) { k += c.colSpan || 1; }); cols = Math.max(cols, k); });
          var width = ((land ? sz.th : sz.tw) - m.x * 15 * 2), cw = Math.floor(width / Math.max(1, cols));
          var x = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="' + width + '" w:type="dxa"/><w:tblBorders>' +
            ["top", "left", "bottom", "right", "insideH", "insideV"].map(function (b) { return "<w:" + b + ' w:val="single" w:sz="4" w:space="0" w:color="BFC3CC"/>'; }).join("") +
            '</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>' + new Array(cols + 1).join('<w:gridCol w:w="' + cw + '"/>') + "</w:tblGrid>";
          rows.forEach(function (r, ri) {
            x += "<w:tr>" + (ri === 0 && r.cells[0] && r.cells[0].tagName === "TH" ? "<w:trPr><w:tblHeader/></w:trPr>" : "");
            Array.prototype.forEach.call(r.cells, function (c) {
              var span = c.colSpan || 1, bg = color(c.style.backgroundColor) || (c.tagName === "TH" ? "EEF0F6" : null);
              x += '<w:tc><w:tcPr><w:tcW w:w="' + cw * span + '" w:type="dxa"/>' + (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : "") + (bg ? '<w:shd w:val="clear" w:color="auto" w:fill="' + bg + '"/>' : "") + "</w:tcPr>";
              var inner = blocksXml(c, c.tagName === "TH");
              x += (inner || "<w:p/>") + "</w:tc>";
            });
            x += "</w:tr>";
          });
          return x + "</w:tbl><w:p/>";
        }
        // ブロックの並び（段落・見出し・表…）を変える。inlineRoot: 中にブロックがないときは、そのまま 1 段落にする
        function blocksXml(root, bold) {
          var out = "", loose = null;
          function flush() { if (loose && loose.childNodes.length && loose.textContent.replace(/\s/g, "") + loose.querySelectorAll("img,br").length) out += "<w:p>" + inline(loose, bold ? { b: 1 } : {}) + "</w:p>"; loose = null; }
          Array.prototype.forEach.call(root.childNodes, function (n) {
            var tag = n.nodeType === 1 ? n.tagName : "";
            if (!/^(P|H1|H2|H3|H4|H5|H6|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|DIV|SECTION|FIGURE|NAV)$/.test(tag)) { (loose = loose || document.createElement("p")).appendChild(n.cloneNode(true)); return; }
            flush();
            if (tag === "P") out += n.classList.contains("toc__title") ? para(n, "TOCHeading") : para(n, n.classList.contains("toc__item--h2") ? "TOC1" : n.classList.contains("toc__item--h3") ? "TOC2" : null);
            else if (tag === "H1") out += para(n, "Title");
            else if (tag === "H2") out += para(n, "Heading1");
            else if (tag === "H3") out += para(n, "Heading2");
            else if (/^H[4-6]$/.test(tag)) out += para(n, "Heading3");
            else if (tag === "BLOCKQUOTE") out += Array.prototype.some.call(n.children, function (c) { return /^(P|UL|OL)$/.test(c.tagName); }) ? blocksXml(n) : para(n, "Quote");
            else if (tag === "PRE") out += n.textContent.split("\n").map(function (l) { return '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>' + textRun(l, { code: 1 }) + "</w:p>"; }).join("");
            else if (tag === "UL" || tag === "OL") out += list(n, 0);
            else if (tag === "TABLE") out += table(n);
            else if (tag === "HR") out += n.classList.contains("pagebreak") ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="C3C6D7"/></w:pBdr></w:pPr></w:p>';
            else out += blocksXml(n, bold);
          });
          flush();
          return out;
        }
        bodyXml.push(blocksXml(holder));
        var hdrId = p.header ? rel("header", "header1.xml") : null, ftrId = p.footer || p.num ? rel("footer", "footer1.xml") : null;
        var W = land ? sz.th : sz.tw, H = land ? sz.tw : sz.th, mx = m.x * 15, my = m.y * 15;
        var sect = "<w:sectPr>" + (hdrId ? '<w:headerReference w:type="default" r:id="' + hdrId + '"/>' : "") + (ftrId ? '<w:footerReference w:type="default" r:id="' + ftrId + '"/>' : "") +
          '<w:pgSz w:w="' + W + '" w:h="' + H + '"' + (land ? ' w:orient="landscape"' : "") + '/><w:pgMar w:top="' + my + '" w:right="' + mx + '" w:bottom="' + my + '" w:left="' + mx + '" w:header="720" w:footer="720" w:gutter="0"/>' +
          (p.cols > 1 ? '<w:cols w:num="' + p.cols + '" w:space="425"/>' : "") + (p.vert ? '<w:textDirection w:val="tbRl"/>' : "") + '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>';
        var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
        var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        var files = { "[Content_Types].xml": "" };
        files["word/document.xml"] = head + "<w:document " + NS + "><w:body>" + bodyXml.join("") + sect + "</w:body></w:document>";
        if (hdrId) files["word/header1.xml"] = head + "<w:hdr " + NS + '><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + textRun(p.header, { color: "5A5F6B", size: 9 }) + "</w:p></w:hdr>";
        if (ftrId) files["word/footer1.xml"] = head + "<w:ftr " + NS + '><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + (p.footer ? textRun(p.footer + (p.num ? "　" : ""), { color: "5A5F6B", size: 9 }) : "") +
          (p.num ? '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' : "") + "</w:p></w:ftr>";
        var lvls = function (fmt) {
          var out = "";
          for (var i = 0; i < 9; i++) {
            out += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + (fmt === "bullet" ? ["●", "○", "■"][i % 3] : "%" + (i + 1) + ".") + '"/><w:lvlJc w:val="left"/>' +
              '<w:pPr><w:ind w:left="' + (420 * (i + 1)) + '" w:hanging="420"/></w:pPr>' + (fmt === "bullet" ? '<w:rPr><w:rFonts w:ascii="Wingdings" w:hint="default"/><w:sz w:val="16"/></w:rPr>' : "") + "</w:lvl>";
          }
          return out.replace(/<w:rFonts w:ascii="Wingdings" w:hint="default"\/>/g, "");
        };
        files["word/numbering.xml"] = head + "<w:numbering " + NS + '><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' + lvls("bullet") + '</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' + lvls("decimal") + "</w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' + numDefs.join("") + "</w:numbering>";
        var font = '<w:rFonts w:ascii="游ゴシック" w:eastAsia="游ゴシック" w:hAnsi="游ゴシック" w:cs="Times New Roman"/>';
        function sty(id, name, ppr, rpr, extra) { return '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' + (extra || "") + (ppr ? "<w:pPr>" + ppr + "</w:pPr>" : "") + (rpr ? "<w:rPr>" + rpr + "</w:rPr>" : "") + "</w:style>"; }
        files["word/styles.xml"] = head + "<w:styles " + NS + "><w:docDefaults><w:rPrDefault><w:rPr>" + font + '<w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="312" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
          sty("Title", "Title", '<w:spacing w:after="240"/><w:jc w:val="left"/>', '<w:b/><w:sz w:val="44"/><w:szCs w:val="44"/>') +
          sty("Heading1", "heading 1", '<w:keepNext/><w:spacing w:before="360" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/>', '<w:b/><w:color w:val="1D4ED8"/><w:sz w:val="32"/><w:szCs w:val="32"/>') +
          sty("Heading2", "heading 2", '<w:keepNext/><w:spacing w:before="240" w:after="80"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/>', '<w:b/><w:sz w:val="26"/><w:szCs w:val="26"/>') +
          sty("Heading3", "heading 3", '<w:keepNext/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/>', "<w:b/>") +
          sty("Quote", "Quote", '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C3C6D7"/></w:pBdr><w:ind w:left="284"/>', '<w:color w:val="434655"/>') +
          sty("ListParagraph", "List Paragraph", '<w:spacing w:after="40"/><w:ind w:left="840"/>') +
          sty("Code", "Code", '<w:spacing w:after="0"/><w:shd w:val="clear" w:color="auto" w:fill="F1F3F7"/>', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/>') +
          sty("TOCHeading", "TOC Heading", '<w:spacing w:after="120"/>', '<w:b/><w:sz w:val="28"/>') + sty("TOC1", "toc 1", '<w:spacing w:after="40"/>') + sty("TOC2", "toc 2", '<w:spacing w:after="40"/><w:ind w:left="420"/>') +
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';
        files["word/settings.xml"] = head + "<w:settings " + NS + '><w:defaultTabStop w:val="840"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
        media.forEach(function (x) { files["word/media/" + x.name] = x.bytes; });
        files["word/_rels/document.xml.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
          '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' + rels.join("") + "</Relationships>";
        files["_rels/.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
        files["docProps/core.xml"] = head + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
          "<dc:title>" + X(O.file() ? O.file().name : "") + "</dc:title><dc:creator>" + X(O.userName() || "SK Hub Office") + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString().replace(/\.\d+Z$/, "Z") + "</dcterms:created></cp:coreProperties>";
        files["[Content_Types].xml"] = head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
          (hdrId ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : "") +
          (ftrId ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : "") +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
        return files;
      });
    }
    function cleanHtml() {
      var h = document.createElement("div");
      h.innerHTML = data.html;
      h.querySelectorAll("[contenteditable]").forEach(function (x) { x.removeAttribute("contenteditable"); });
      h.querySelectorAll(".is-sel").forEach(function (x) { x.classList.remove("is-sel"); });
      return h.innerHTML;
    }
    function htmlFile() {
      var name = O.file() ? O.file().name : "", p = pageOf(data);
      return '<!DOCTYPE html>\n<html lang="ja">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>' + esc(name) + "</title>\n<style>\n" +
        "body{max-width:760px;margin:40px auto;padding:0 20px;font-family:\"Yu Gothic\",\"Hiragino Sans\",\"Noto Sans JP\",sans-serif;line-height:1.85;color:#1a1c20}" +
        "h1{font-size:30px;line-height:1.35}h2{font-size:21px;color:#1d4ed8;margin-top:28px}h3{font-size:17px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #d5d6dc;padding:8px 12px;text-align:left;vertical-align:top}th{background:#eef0f6}" +
        "blockquote{margin:0 0 12px;padding-left:16px;border-left:3px solid #c3c6d7;color:#434655}img{max-width:100%;height:auto}.wrap-l{float:left;margin:4px 16px 8px 0}.wrap-r{float:right;margin:4px 0 8px 16px}.wrap-c{display:block;margin:8px auto}" +
        "hr.pagebreak{border:0;break-after:page}.meta{color:#5a5f6b;font-size:14px}.right{text-align:right}.toc{padding:12px 16px;background:#f3f4f8;border-radius:8px}.toc a{color:inherit}.toc__item--h3{padding-left:1.5em}" +
        (p.vert ? "body{writing-mode:vertical-rl;max-width:none;height:90vh}" : "") + "\n</style>\n</head>\n<body>\n" + cleanHtml() + "\n</body>\n</html>\n";
    }
    function mdOf(node) {
      var out = "";
      function inl(n) {
        var s = "";
        Array.prototype.forEach.call(n.childNodes, function (c) {
          if (c.nodeType === 3) { s += c.data.replace(/\u200b/g, "").replace(/([*_`\[\]])/g, "\\$1"); return; }
          if (c.nodeType !== 1) return;
          var tag = c.tagName, inner = inl(c);
          if (tag === "BR") s += "  \n";
          else if (/^(B|STRONG)$/.test(tag) || +(c.style && c.style.fontWeight) >= 600) s += inner.trim() ? "**" + inner + "**" : inner;
          else if (/^(I|EM)$/.test(tag)) s += inner.trim() ? "*" + inner + "*" : inner;
          else if (/^(S|DEL|STRIKE)$/.test(tag)) s += "~~" + inner + "~~";
          else if (tag === "CODE") s += "`" + c.textContent + "`";
          else if (tag === "A") s += "[" + inner + "](" + c.getAttribute("href") + ")";
          else if (tag === "IMG") s += "![" + (c.alt || "") + "](" + (/^data:/.test(c.src) ? "image" : c.getAttribute("src")) + ")";
          else s += inner;
        });
        return s;
      }
      function list(n, depth) {
        var i = 0;
        Array.prototype.forEach.call(n.children, function (li) {
          if (li.tagName !== "LI") return;
          i++;
          var own = li.cloneNode(true); own.querySelectorAll("ul,ol").forEach(function (x) { x.remove(); });
          out += new Array(depth + 1).join("  ") + (n.tagName === "OL" ? i + ". " : "- ") + inl(own).trim() + "\n";
          Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") list(c, depth + 1); });
        });
      }
      Array.prototype.forEach.call(node.childNodes, function (n) {
        if (n.nodeType === 3) { if (n.data.trim()) out += n.data.trim() + "\n\n"; return; }
        if (n.nodeType !== 1) return;
        var tag = n.tagName;
        if (/^H([1-6])$/.test(tag)) out += new Array(+tag.charAt(1) + 1).join("#") + " " + inl(n).trim() + "\n\n";
        else if (tag === "P") { var x = inl(n).trim(); out += (x ? x : "") + "\n\n"; }
        else if (tag === "UL" || tag === "OL") { list(n, 0); out += "\n"; }
        else if (tag === "BLOCKQUOTE") out += inl(n).trim().split("\n").map(function (l) { return "> " + l; }).join("\n") + "\n\n";
        else if (tag === "PRE") out += "```\n" + n.textContent + "\n```\n\n";
        else if (tag === "HR") out += "---\n\n";
        else if (tag === "TABLE") {
          var rows = Array.prototype.map.call(n.rows, function (r) { return Array.prototype.map.call(r.cells, function (c) { return inl(c).trim().replace(/\|/g, "\\|").replace(/\n/g, " "); }); });
          if (rows.length) {
            out += "| " + rows[0].join(" | ") + " |\n|" + rows[0].map(function () { return " --- "; }).join("|") + "|\n";
            rows.slice(1).forEach(function (r) { out += "| " + r.join(" | ") + " |\n"; });
            out += "\n";
          }
        } else if (tag === "FIGURE") out += "（" + t("グラフ") + "）\n\n";
        else out += mdOf(n);
      });
      return out;
    }
    function exportAs(id) {
      sync();
      if (id === "docx") {
        return docxFiles().then(Zip.write).then(function (b) {
          var mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          return O.saveBlob(O.fileName(".docx"), new Blob([b], { type: mime }), t("Word 文書"), mime, ".docx");
        });
      }
      if (id === "html") return O.saveBlob(O.fileName(".html"), new Blob([htmlFile()], { type: "text/html" }), "HTML", "text/html", ".html");
      if (id === "md") { var h = document.createElement("div"); h.innerHTML = cleanHtml(); return O.saveBlob(O.fileName(".md"), new Blob([mdOf(h).replace(/\n{3,}/g, "\n\n")], { type: "text/markdown" }), "Markdown", "text/markdown", ".md"); }
      if (id === "txt") { var d = document.createElement("div"); d.innerHTML = cleanHtml(); document.body.appendChild(d); d.style.cssText = "position:fixed;left:-9999px;top:0;width:700px"; var text = d.innerText; d.remove(); return O.saveBlob(O.fileName(".txt"), new Blob(["\ufeff" + text], { type: "text/plain" }), t("テキスト"), "text/plain", ".txt"); }
    }

    // ---------- 読み込み（.docx・.html・.md・.txt） ----------
    function kids(node, name) { return node ? Array.prototype.slice.call(node.getElementsByTagNameNS("*", name)) : []; }
    function kid(node, name) { var k = node && node.firstElementChild; while (k) { if (k.localName === name) return k; k = k.nextElementSibling; } return null; }
    function attr(node, name) { return node ? node.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", name) || node.getAttribute("w:" + name) : null; }
    function readDocx(file) {
      if (!Zip.canRead) return Promise.reject(new Error("no-inflate"));
      return Zip.read(file).then(function (z) {
        var doc = Zip.xml(z["word/document.xml"]);
        if (!doc) throw new Error("not-docx");
        var relDoc = Zip.xml(z["word/_rels/document.xml.rels"]), rels = {};
        kids(relDoc, "Relationship").forEach(function (r) { rels[r.getAttribute("Id")] = { target: r.getAttribute("Target"), type: r.getAttribute("Type") || "", ext: r.getAttribute("TargetMode") === "External" }; });
        var numDoc = Zip.xml(z["word/numbering.xml"]), absFmt = {}, numAbs = {};
        kids(numDoc, "abstractNum").forEach(function (a) { var id = attr(a, "abstractNumId"); absFmt[id] = {}; kids(a, "lvl").forEach(function (l) { absFmt[id][attr(l, "ilvl")] = attr(kid(l, "numFmt"), "val"); }); });
        kids(numDoc, "num").forEach(function (n) { numAbs[attr(n, "numId")] = attr(kid(n, "abstractNumId"), "val"); });
        var styleDoc = Zip.xml(z["word/styles.xml"]), styleName = {};
        kids(styleDoc, "style").forEach(function (s) { var nm = kid(s, "name"); styleName[attr(s, "styleId")] = (attr(nm, "val") || "").toLowerCase(); });
        function blockTag(styleId) {
          var n = (styleName[styleId] || styleId || "").toLowerCase();
          if (/^title$|^表題$/.test(n)) return "h1";
          if (/heading ?1$|見出し ?1$/.test(n)) return "h2";
          if (/heading ?[2-9]$|見出し ?[2-9]$/.test(n)) return "h3";
          if (/quote|引用/.test(n)) return "blockquote";
          return "p";
        }
        function mediaUrl(id) {
          var r = rels[id];
          if (!r) return "";
          var path = r.target.replace(/^\/?word\//, "").replace(/^\//, ""), bytes = z["word/" + path] || z[path];
          if (!bytes) return "";
          var ext = (path.split(".").pop() || "png").toLowerCase(), mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "svg" ? "image/svg+xml" : "image/png";
          if (/emf|wmf|tif/.test(ext)) return "";
          return "data:" + mime + ";base64," + b64(bytes);
        }
        function hex(v) { return v && /^[0-9a-f]{6}$/i.test(v) ? "#" + v.toLowerCase() : null; }
        var HL = { yellow: "#fff59d", green: "#bbf7d0", cyan: "#a5f3fc", magenta: "#f5d0fe", blue: "#bfdbfe", red: "#fecaca", darkBlue: "#93c5fd", lightGray: "#e5e7eb", darkGray: "#9ca3af" };
        function runHtml(r) {
          var pr = kid(r, "rPr"), html = "";
          Array.prototype.forEach.call(r.childNodes, function (c) {
            if (c.nodeType !== 1) return;
            var ln = c.localName;
            if (ln === "t") html += esc(c.textContent);
            else if (ln === "tab") html += "\u3000";
            else if (ln === "br") html += attr(c, "type") === "page" ? '</p><hr class="pagebreak"><p>' : "<br>";
            else if (ln === "drawing" || ln === "pict") {
              var blip = kids(c, "blip")[0], imd = kids(c, "imagedata")[0];
              var id = blip ? blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed") || blip.getAttribute("r:embed") : imd ? imd.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") : null;
              var ext = kids(c, "extent")[0], src = id ? mediaUrl(id) : "";
              if (src) html += '<img src="' + src + '" alt=""' + (ext ? ' style="width:' + Math.round(+ext.getAttribute("cx") / 9525) + 'px"' : "") + ">";
            }
          });
          if (!html || !pr) return html;
          var css = "", open = "", close = "";
          function on(name) { var e = kid(pr, name); return e && attr(e, "val") !== "0" && attr(e, "val") !== "false" && attr(e, "val") !== "none"; }
          if (on("b")) { open += "<b>"; close = "</b>" + close; }
          if (on("i")) { open += "<i>"; close = "</i>" + close; }
          if (on("u")) { open += "<u>"; close = "</u>" + close; }
          if (on("strike") || on("dstrike")) { open += "<s>"; close = "</s>" + close; }
          var va = attr(kid(pr, "vertAlign"), "val");
          if (va === "superscript") { open += "<sup>"; close = "</sup>" + close; }
          if (va === "subscript") { open += "<sub>"; close = "</sub>" + close; }
          var col = hex(attr(kid(pr, "color"), "val")), sz = attr(kid(pr, "sz"), "val"), hl = HL[attr(kid(pr, "highlight"), "val")] || hex(attr(kid(pr, "shd"), "fill"));
          var fonts = kid(pr, "rFonts"), ff = fonts ? attr(fonts, "eastAsia") || attr(fonts, "ascii") : null;
          if (col && col !== "#000000") css += "color:" + col + ";";
          if (sz && +sz !== 21 && +sz !== 22 && +sz !== 24) css += "font-size:" + (+sz / 2) + "pt;";
          if (hl) css += "background-color:" + hl + ";";
          if (ff && !/^(游ゴシック|游明朝|ＭＳ 明朝|ＭＳ ゴシック|Century|Calibri|Aptos|Times New Roman|Yu Gothic|Yu Mincho)/.test(ff)) css += "font-family:" + O.fontCss(ff).replace(/"/g, "'") + ";";
          return open + (css ? '<span style="' + css + '">' + html + "</span>" : html) + close;
        }
        function paraInner(p) {
          var html = "";
          Array.prototype.forEach.call(p.childNodes, function (c) {
            if (c.nodeType !== 1) return;
            if (c.localName === "r") html += runHtml(c);
            else if (c.localName === "hyperlink") {
              var id = c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || c.getAttribute("r:id"), r = id && rels[id], inner = "";
              kids(c, "r").forEach(function (rr) { inner += runHtml(rr); });
              html += r && r.ext ? '<a href="' + esc(r.target) + '">' + inner + "</a>" : inner;
            } else if (c.localName === "ins" || c.localName === "smartTag" || c.localName === "sdt" || c.localName === "fldSimple") {
              var holder = c.localName === "sdt" ? kid(c, "sdtContent") || c : c;
              html += paraInner(holder);
            }
          });
          return html;
        }
        var out = [], listStack = [];
        function closeLists(level) { while (listStack.length > level) out.push("</li></" + listStack.pop() + ">"); }
        function paraHtml(p) {
          var pr = kid(p, "pPr"), st = attr(kid(pr, "pStyle"), "val"), tag = blockTag(st), numPr = kid(pr, "numPr");
          var jc = attr(kid(pr, "jc"), "val"), inner = paraInner(p);
          var style = jc === "center" ? "text-align:center" : jc === "right" || jc === "end" ? "text-align:right" : jc === "both" && tag !== "p" ? "" : "";
          var brk = kid(pr, "pageBreakBefore") ? '<hr class="pagebreak">' : "";
          if (numPr) {
            var lvl = +(attr(kid(numPr, "ilvl"), "val") || 0), nid = attr(kid(numPr, "numId"), "val"), fmt = (absFmt[numAbs[nid]] || {})[lvl] || "bullet", lt = fmt === "bullet" || fmt === "none" ? "ul" : "ol";
            if (listStack.length > lvl + 1) closeLists(lvl + 1);
            if (listStack.length === lvl + 1 && listStack[lvl] !== lt) closeLists(lvl);
            if (listStack.length === lvl + 1) out.push("</li><li>" + inner);
            else { while (listStack.length < lvl + 1) { out.push("<" + lt + "><li>"); listStack.push(lt); } out.push(inner); }
            return;
          }
          closeLists(0);
          out.push(brk + "<" + tag + (style ? ' style="' + style + '"' : "") + ">" + (inner || "<br>") + "</" + tag + ">");
        }
        function tableHtml2(tbl) {
          var h = "<table><tbody>";
          kids(tbl, "tr").filter(function (tr) { return tr.parentNode === tbl; }).forEach(function (tr, ri) {
            var head = kid(kid(tr, "trPr"), "tblHeader");
            h += "<tr>";
            Array.prototype.forEach.call(tr.children, function (tc) {
              if (tc.localName !== "tc") return;
              var pr = kid(tc, "tcPr"), span = +(attr(kid(pr, "gridSpan"), "val") || 1), vm = kid(pr, "vMerge"), fill = hex(attr(kid(pr, "shd"), "fill"));
              if (vm && attr(vm, "val") !== "restart") return;
              var cellTag = head || (ri === 0 && fill) ? "th" : "td", inner = [];
              Array.prototype.forEach.call(tc.children, function (c) { if (c.localName === "p") { var x = paraInner(c); if (x) inner.push(x); } else if (c.localName === "tbl") inner.push(tableHtml2(c)); });
              h += "<" + cellTag + (span > 1 ? ' colspan="' + span + '"' : "") + (fill && cellTag === "td" && fill !== "#ffffff" ? ' style="background-color:' + fill + '"' : "") + ">" + (inner.join("<br>") || "<br>") + "</" + cellTag + ">";
            });
            h += "</tr>";
          });
          return h + "</tbody></table>";
        }
        function walkBody(node) {
          Array.prototype.forEach.call(node.children, function (c) {
            if (c.localName === "p") paraHtml(c);
            else if (c.localName === "tbl") { closeLists(0); out.push(tableHtml2(c)); }
            else if (c.localName === "sdt") { var sc = kid(c, "sdtContent"); if (sc) walkBody(sc); }
          });
        }
        var body = kids(doc, "body")[0];
        walkBody(body);
        closeLists(0);
        var data2 = { html: sanitize(out.join("").replace(/<p><\/p>/g, "")) || "<p><br></p>", page: {}, comments: [] };
        var sect = kids(body, "sectPr").pop(), pg = kids(sect, "pgSz")[0], mar = kids(sect, "pgMar")[0], cols = kids(sect, "cols")[0], td = kids(sect, "textDirection")[0];
        if (pg) {
          var w = +attr(pg, "w"), hh = +attr(pg, "h"), land = attr(pg, "orient") === "landscape" || w > hh, short = Math.min(w, hh);
          data2.page.orient = land ? "landscape" : "portrait";
          data2.page.size = Math.abs(short - 12240) < 200 ? "letter" : Math.abs(short - 10318) < 300 ? "b5" : Math.abs(short - 16838) < 300 ? "a3" : "a4";
        }
        if (mar) { var l = +attr(mar, "left"); data2.page.margin = l < 1000 ? "narrow" : l > 1900 ? "wide" : "normal"; }
        if (cols && +attr(cols, "num") > 1) data2.page.cols = Math.min(3, +attr(cols, "num"));
        if (td && /tbRl/i.test(attr(td, "val") || "")) data2.page.vert = true;
        Object.keys(rels).forEach(function (id) {
          var r = rels[id];
          if (/\/header$/.test(r.type) && !data2.page.header) { var hd = Zip.xml(z["word/" + r.target.replace(/^\/?word\//, "")]); if (hd) data2.page.header = hd.documentElement.textContent.trim().slice(0, 200); }
          if (/\/footer$/.test(r.type) && !data2.page.footer) { var ft = Zip.xml(z["word/" + r.target.replace(/^\/?word\//, "")]); if (ft) { var tx = ft.documentElement.textContent.trim(); if (kids(ft, "fldSimple").length || /PAGE/.test(ft.documentElement.textContent)) data2.page.num = true; data2.page.footer = tx.replace(/^\d+$/, "").slice(0, 200); } }
        });
        return data2;
      });
    }
    function mdToHtml(md) {
      var lines = md.replace(/\r\n?/g, "\n").split("\n"), out = [], i = 0;
      function inl(s) {
        return esc(s).replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, function (m, a, u) { return /^(https?:|data:image\/)/.test(u) ? '<img src="' + u + '" alt="' + a + '">' : a; })
          .replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, function (m, a, u) { return /^(https?:|mailto:|#)/.test(u) ? '<a href="' + u + '">' + a + "</a>" : a; })
          .replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>").replace(/__([^_]+)__/g, "<b>$1</b>")
          .replace(/(^|[^*])\*([^*]+)\*/g, "$1<i>$2</i>").replace(/~~([^~]+)~~/g, "<s>$1</s>");
      }
      while (i < lines.length) {
        var l = lines[i], m;
        if (/^```/.test(l)) { var code = []; i++; while (i < lines.length && !/^```/.test(lines[i])) code.push(lines[i++]); i++; out.push("<pre>" + esc(code.join("\n")) + "</pre>"); continue; }
        if ((m = /^(#{1,6})\s+(.*)$/.exec(l))) { var lv = m[1].length; out.push("<" + (lv === 1 ? "h1" : lv === 2 ? "h2" : "h3") + ">" + inl(m[2]) + "</" + (lv === 1 ? "h1" : lv === 2 ? "h2" : "h3") + ">"); i++; continue; }
        if (/^(-{3,}|\*{3,})\s*$/.test(l)) { out.push("<hr>"); i++; continue; }
        if (/^>\s?/.test(l)) { var q = []; while (i < lines.length && /^>\s?/.test(lines[i])) q.push(lines[i++].replace(/^>\s?/, "")); out.push("<blockquote>" + inl(q.join(" ")) + "</blockquote>"); continue; }
        if (/^\s*([-*+]|\d+[.)])\s+/.test(l)) {
          var ordered = /^\s*\d/.test(l), items = [];
          while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) items.push("<li>" + inl(lines[i++].replace(/^\s*([-*+]|\d+[.)])\s+/, "")) + "</li>");
          out.push((ordered ? "<ol>" : "<ul>") + items.join("") + (ordered ? "</ol>" : "</ul>"));
          continue;
        }
        if (/^\|.*\|\s*$/.test(l)) {
          var rows = [];
          while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) { if (!/^\|[\s:|-]+\|\s*$/.test(lines[i])) rows.push(lines[i].trim().slice(1, -1).split("|").map(function (x) { return inl(x.trim()); })); i++; }
          out.push("<table><tbody>" + rows.map(function (r, ri) { return "<tr>" + r.map(function (c) { return ri ? "<td>" + c + "</td>" : "<th>" + c + "</th>"; }).join("") + "</tr>"; }).join("") + "</tbody></table>");
          continue;
        }
        if (!l.trim()) { i++; continue; }
        var para = [];
        while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|>|\s*([-*+]|\d+[.)])\s+|\|)/.test(lines[i])) para.push(inl(lines[i++]));
        out.push("<p>" + para.join("<br>") + "</p>");
      }
      return out.join("") || "<p><br></p>";
    }
    function importFile(file) {
      if (/\.docx$/i.test(file.name)) return readDocx(file);
      return file.arrayBuffer().then(function (ab) {
        var u8 = new Uint8Array(ab), text;
        try { text = new TextDecoder("utf-8", { fatal: true }).decode(u8); } catch (e) { text = new TextDecoder("shift_jis").decode(u8); }
        text = text.replace(/^\ufeff/, "");
        var html;
        if (/\.html?$/i.test(file.name)) { var b = /<body[^>]*>([\s\S]*)<\/body>/i.exec(text); html = sanitize(b ? b[1] : text); }
        else if (/\.(md|markdown)$/i.test(file.name)) html = sanitize(mdToHtml(text));
        else html = text.split(/\r?\n/).map(function (l) { return "<p>" + (esc(l) || "<br>") + "</p>"; }).join("");
        return { html: html || "<p><br></p>", page: {}, comments: [] };
      });
    }

    // ---------- 印刷 ----------
    function cssText(s) { return '"' + String(s || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/[\r\n]/g, " ") + '"'; }
    function beforePrint() {
      sync();
      if (imgSel) imgSel.classList.remove("is-sel");
      var p = pageOf(data), m = MARGINS[p.margin], sz = SIZES[p.size];
      var boxes = (p.header ? "@top-center { content: " + cssText(p.header) + "; font-size: 9pt; color: #5a5f6b; }" : "") +
        (p.footer || p.num ? "@bottom-center { content: " + (p.footer ? cssText(p.footer + (p.num ? "　" : "")) + " " : "") + (p.num ? "counter(page)" : "") + "; font-size: 9pt; color: #5a5f6b; }" : "");
      return "@page { size: " + sz.css + " " + p.orient + "; margin: " + m.mm + "; " + boxes + " }" +
        "@media print { .doc-wrap, .doc-scroll { display: block !important; height: auto !important; overflow: visible !important; padding: 0 !important; } .doc-page { width: auto !important; min-height: 0 !important; padding: 0 !important; zoom: 1 !important; background: none !important; }" +
        " .doc-hf, .img-rs, .doc-nav, .doc-cmts { display: none !important; } .cmt { background: none !important; } }";
    }
    function afterPrint() { if (imgSel) imgSel.classList.add("is-sel"); }

    // ---------- 小さなプレビュー ----------
    function thumb(d) {
      var html = String(d.html || "").replace(/<img[^>]*>/gi, '<span class="th-img"></span>').replace(/<svg[\s\S]*?<\/svg>/gi, '<span class="th-img"></span>');
      return '<div class="th th--doc" aria-hidden="true"><div class="th-doc">' + sanitize(html) + "</div></div>";
    }

    return {
      ribbon: ribbon,
      upgrade: upgrade,
      thumb: thumb,
      render: render,
      command: command,
      state: state,
      comments: true,
      zoom: function (z) { zoomZ = z; applyZoom(); stat(); },
      fitZoom: function () { var sc = el("doc-scroll"), d = dims(pageOf(data)); return sc ? Math.max(0.25, Math.min(3, (sc.clientWidth - 48) / d.w)) : 1; },
      defaultZoom: function () { var sc = $("canvas"), d = dims(pageOf(O.data() || {})); return sc && sc.clientWidth && sc.clientWidth < d.w + 48 ? Math.max(0.4, Math.floor((sc.clientWidth - 24) / d.w * 20) / 20) : 1; },
      focus: function () { back(); },
      open: function () { showNav = false; showCmts = false; },
      close: function () { page = null; imgSel = null; },
      beforeSave: sync,
      beforeUndo: function () { sync(); },
      setRead: function (on) { readOnly = on; if (page) page.contentEditable = on ? "false" : "true"; ["doc-head", "doc-foot"].forEach(function (id) { if (el(id)) el(id).readOnly = on; }); if (on) selectImg(null); stat(); },
      selectedText: function () { var s = getSelection(); return s.rangeCount && inPage(s.anchorNode) ? s.toString() : ""; },
      find: find,
      replace: replace,
      findClose: function () { back(); },
      clip: function (cmd) {
        back();
        if (cmd === "paste") {
          if (navigator.clipboard && navigator.clipboard.read) {
            navigator.clipboard.read().then(function (items) {
              var it = items[0];
              if (!it) return;
              var img = it.types.filter(function (x) { return /^image\//.test(x); })[0];
              if (img) return it.getType(img).then(function (b) { return O.imageData(new File([b], "paste", { type: img })); }).then(insertImage);
              if (it.types.indexOf("text/html") >= 0) return it.getType("text/html").then(function (b) { return b.text(); }).then(function (h) { insertHtml(sanitize(h)); });
              return it.getType("text/plain").then(function (b) { return b.text(); }).then(function (tx) { back(); document.execCommand("insertText", false, tx); changed(); });
            }).catch(function () { O.notify(t("Ctrl+V で貼り付けてください")); });
          } else O.notify(t("Ctrl+V で貼り付けてください"));
          return;
        }
        if (!document.execCommand(cmd)) O.notify(cmd === "cut" ? t("Ctrl+X で切り取ってください") : t("Ctrl+C でコピーしてください"));
        else if (cmd === "cut") changed();
      },
      key: function (e, field) {
        if (field || !page || e.target !== page && !inPage(e.target)) return;
        var mod = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
        if (!mod || e.altKey) return;
        var map = { e: "justifyCenter", l: "justifyLeft", r: "justifyRight", j: "justifyFull", k: "link" };
        if (map[k] && !e.shiftKey) { e.preventDefault(); command(map[k]); return; }
        if (k === "enter") { e.preventDefault(); command("pagebreak"); return; }
        if (e.shiftKey && (k === ">" || k === ".")) { e.preventDefault(); command("grow"); return; }
        if (e.shiftKey && (k === "<" || k === ",")) { e.preventDefault(); command("shrink"); return; }
      },
      beforePrint: beforePrint,
      afterPrint: afterPrint,
      exports: [
        { id: "docx", icon: "description", label: t("Word 文書（.docx）"), desc: t("Microsoft Word・LibreOffice・Google ドキュメントで開けます") },
        { id: "html", icon: "html", label: t("Web ページ（.html）"), desc: t("ブラウザで見られる 1 つのファイルにします") },
        { id: "md", icon: "markdown", label: "Markdown（.md）", desc: t("見出しや箇条書きを、記号で表した文字にします") },
        { id: "txt", icon: "text_snippet", label: t("テキスト（.txt）"), desc: t("文字だけを書き出します") }
      ],
      exportAs: exportAs,
      imports: { label: ".docx・.html・.md・.txt", accept: ".docx,.html,.htm,.md,.markdown,.txt" },
      importFile: importFile,
      props: function () { var c = page ? counts() : { pages: 1, chars: 0 }; var p = pageOf(data); return [[t("ページ数"), String(c.pages)], [t("文字数"), String(c.chars)], [t("用紙"), SIZES[p.size].label + "・" + (p.orient === "landscape" ? t("横") : t("縦"))]]; },
      keys: [["Ctrl+B / I / U", t("太字・斜体・下線")], ["Ctrl+E / L / R / J", t("中央・左・右・両端揃え")], ["Ctrl+K", t("リンク")], ["Ctrl+Enter", t("改ページ")],
        ["Ctrl+Shift+> / <", t("文字を大きく・小さく")], ["Tab / Shift+Tab", t("箇条書きの段を下げる・上げる（表では次・前のセル）")]]
    };
  };
})();
