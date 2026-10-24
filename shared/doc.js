// SK Hub Office の Doc（文書）。window.OfficeApps.doc に登録し、/office/shared/office.js が呼ぶ（O は office.js が渡す道具）
//   ファイルの中身（data）: { html: 本文, page: 用紙, comments: [コメント] }
//     page: { size: "a4|letter|b5|a3", orient: "portrait|landscape", margin: "normal|narrow|wide", mx, my: 左右・上下の余白（px。あれば margin より優先）, cols: 1〜3, vert: 縦書き, header, footer, num: ページ番号,
//             base: { fs: 本文の文字の大きさ（pt）, lh: 行間, ff: フォント }（.docx の docDefaults。無ければ画面の標準）,
//             design: DESIGNS のキー（見出しの色・フォント・h2 の飾りなど。無ければ標準） }
//     comments: [{ id, text, author, ts }]（本文の中では <span class="cmt" data-cmt="id"> で印を付ける）
//   本文は contenteditable の <article id="page">。書式は document.execCommand（styleWithCSS で、色・大きさは span の style になる）
//   画面: ナビゲーション（見出しの一覧）・用紙（ヘッダー・本文・フッター）・コメント。用紙の背景に、印刷したときのページの区切りの線を引く
//   元に戻すは office.js（中身をまるごと覚える）。打っているあいだはまとめて 1 回にする
//   .docx の読み書きは zip.js、グラフは chart.js（図は <figure class="chart" data-spec="…"> に SVG を入れる）
//   埋め込みフォント（word/fonts/*.odttf）は opentype.js で読み、画面のフォントと行間に使う
(function () {
  "use strict";
  window.OfficeApps = window.OfficeApps || {};
  window.OfficeApps.doc = function (O) {
    var t = O.t, esc = O.esc, icon = O.icon, $ = O.$, Zip = window.OfficeZip, Chart = window.OfficeChart;

    // ---------- 用紙 ----------
    var SIZES = { a4: { w: 794, h: 1123, label: "A4", tw: 11906, th: 16838, css: "A4" }, b5: { w: 688, h: 972, label: "B5", tw: 10318, th: 14570, css: "182mm 257mm" },
      a3: { w: 1123, h: 1587, label: "A3", tw: 16838, th: 23811, css: "A3" }, letter: { w: 816, h: 1056, label: "Letter", tw: 12240, th: 15840, css: "letter" } };
    var MARGINS = { normal: { x: 96, y: 96, label: t("標準"), mm: "25.4mm" }, narrow: { x: 48, y: 48, label: t("狭い"), mm: "12.7mm" }, wide: { x: 144, y: 96, label: t("広い"), mm: "25.4mm 38.1mm" } };
    // ---------- デザイン（見出しの色・フォント・見出しの飾り・表の見出しの色）。page.design に ID を入れる。無いときは「標準」 ----------
    //   h1・h2・h3: 見出しの色（h1 は表題）。ac: 引用の線・リンクの色。th: 表の見出しのセルの色。rule: h2 の飾り（"rule" 下線・"bar" 左の線）。h1c: 表題を中央に。
    //   hf・bf: 画面の見出し・本文のフォント。hface・face: .docx に書くフォント名
    var SERIF = '"Noto Serif JP","Yu Mincho","Hiragino Mincho ProN",serif';
    var DESIGNS = {
      business: { label: t("ビジネス"), group: "biz", h1: "#0b1f3a", h2: "#0b1f3a", h3: "#4a5a73", ac: "#b8862b", th: "#f1ece0", rule: "rule" },
      modern: { label: t("モダン"), group: "biz", h1: "#111827", h2: "#2563eb", h3: "#374151", ac: "#2563eb", th: "#eff6ff", rule: "bar" },
      tech: { label: t("テック"), group: "biz", h1: "#0f172a", h2: "#0d9488", h3: "#334155", ac: "#0d9488", th: "#e6f7f5", rule: "rule" },
      report: { label: t("レポート"), group: "biz", h1: "#1c1c1c", h2: "#1c1c1c", h3: "#444444", ac: "#777777", th: "#f0f0f0", rule: "rule", h1c: 1, hf: SERIF, bf: SERIF, hface: "Yu Mincho", face: "Yu Mincho" },
      mono: { label: t("モノクロ"), group: "biz", h1: "#000000", h2: "#000000", h3: "#444444", ac: "#000000", th: "#eeeeee", rule: "rule" },
      warm: { label: t("あたたかい"), group: "color", h1: "#7c2d12", h2: "#c2410c", h3: "#9a3412", ac: "#c2410c", th: "#fff1e6", rule: "bar" },
      nature: { label: t("自然"), group: "color", h1: "#14532d", h2: "#15803d", h3: "#166534", ac: "#15803d", th: "#e8f5ec", rule: "rule" },
      sakura: { label: t("さくら"), group: "color", h1: "#831843", h2: "#be185d", h3: "#9d174d", ac: "#be185d", th: "#fdeef4", rule: "rule" },
      lavender: { label: t("ラベンダー"), group: "color", h1: "#3b2a73", h2: "#6d4fc8", h3: "#5b3fb0", ac: "#6d4fc8", th: "#f0eaff", rule: "bar" },
      washi: { label: t("和風"), group: "classic", h1: "#2a2118", h2: "#b3261e", h3: "#6a5a47", ac: "#b3261e", th: "#efe6d3", rule: "bar", hf: SERIF, bf: SERIF, hface: "Yu Mincho", face: "Yu Mincho" },
      classic: { label: t("クラシック"), group: "classic", h1: "#3a2c1d", h2: "#6f451f", h3: "#74624b", ac: "#6f451f", th: "#ece0c9", rule: "rule", h1c: 1, hf: SERIF, bf: SERIF, hface: "Yu Mincho", face: "Yu Mincho" }
    };
    var DESIGN_GROUPS = [["", t("SK のデザイン")], ["biz", t("ビジネス")], ["color", t("いろいろな色")], ["classic", t("和・クラシック")]];
    // カスタムスタイル（data.styles）。元のスタイル → .docx のスタイル ID
    var CST_BASE = { p: "Normal", h1: "Title", h2: "Heading1", h3: "Heading2", blockquote: "Quote" };
    var CST_TAG = { p: "p", h1: "h1", h2: "h2", h3: "h3", blockquote: "blockquote" };
    var CST_LABEL = { p: "段落", h1: "表題", h2: "見出し 1", h3: "見出し 2", blockquote: "引用" };
    function pageOf(d) {
      var p = d.page || (d.page = {});
      if (!SIZES[p.size]) p.size = "a4";
      if (p.orient !== "landscape") p.orient = "portrait";
      if (!MARGINS[p.margin]) p.margin = "normal";
      if (!DESIGNS[p.design]) delete p.design;
      ["mx", "my"].forEach(function (k) {
        if (p[k] !== undefined && p[k] !== null && isFinite(+p[k]) && +p[k] >= 0) p[k] = +p[k];
        else delete p[k];
      });
      if (p.base && typeof p.base === "object") {
        var b = {}, bs = p.base;
        if (+bs.fs > 0) b.fs = +bs.fs;
        if (+bs.lh > 0) b.lh = +bs.lh;
        if (bs.ff) b.ff = String(bs.ff);
        if (b.fs || b.lh || b.ff) p.base = b; else delete p.base;
      } else delete p.base;
      p.cols = Math.max(1, Math.min(3, +p.cols || 1));
      return p;
    }
    function dims(p) { var s = SIZES[p.size]; return p.orient === "landscape" ? { w: s.h, h: s.w } : { w: s.w, h: s.h }; }
    function upgrade(d) {
      if (typeof d.html !== "string") d.html = "<p><br></p>";
      pageOf(d);
      if (!Array.isArray(d.comments)) d.comments = [];
      if (!Array.isArray(d.styles)) d.styles = [];
      d.styles = d.styles.filter(function (s) { return s && s.id && s.name; }).map(function (s) {
        return { id: String(s.id), name: String(s.name), base: CST_BASE[s.base] ? s.base : "p", css: String(s.css || "") };
      });
      return d;
    }

    // ---------- 安全な HTML（貼り付け・読み込み） ----------
    var ALLOW = /^(P|H1|H2|H3|H4|H5|H6|UL|OL|LI|TABLE|THEAD|TBODY|TFOOT|TR|TD|TH|CAPTION|COLGROUP|COL|B|STRONG|I|EM|U|S|STRIKE|DEL|INS|SUB|SUP|SPAN|A|BR|HR|IMG|BLOCKQUOTE|PRE|CODE|DIV|FIGURE|FIGCAPTION|MARK|SMALL|FONT|SECTION|NAV|SVG|G|PATH|RECT|CIRCLE|LINE|POLYLINE|TEXT|TITLE)$/i;
    var STYLE_OK = /^(color|background-color|font-weight|font-style|text-decoration|text-decoration-line|text-align|font-size|font-family|line-height|width|height|padding|vertical-align|float|position|left|top|z-index|margin-left|margin-right|margin-top|margin-bottom|border|border-top|border-right|border-bottom|border-left|border-collapse|table-layout)$/i;
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
            if (/^(href|src|alt|title|colspan|rowspan|class|data-cmt|data-spec|data-fn|data-bd|data-cst|data-anc|id|width|height|viewbox|d|x|y|x1|x2|y1|y2|cx|cy|r|fill|stroke|stroke-width|points|font-size|text-anchor|font-weight|xmlns|rx|transform|fill-opacity|stroke-linejoin|stroke-linecap|pointer-events|font-family|contenteditable)$/i.test(name)) return;
            n.removeAttribute(a.name);
          });
          if (tag === "A") { n.setAttribute("target", "_blank"); n.setAttribute("rel", "noopener"); }
          walk(n);
        });
      })(doc.body);
      return doc.body.innerHTML;
    }

    // ---------- 埋め込みフォント（.docx の word/fonts/*.odttf） ----------
    // Word はフォントを同梱できる。opentype.js で読めれば、画面でも同じフォントで描き、
    // フォントの ascent/descent から Word に近い行間を出す（無いときは今までの 1.15 のまま）。
    var fontMetrics = {}, embedFaces = [];
    var REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    function clearEmbeddedFonts() {
      if (document.fonts && embedFaces.length) embedFaces.forEach(function (f) { try { document.fonts.delete(f); } catch (e) { /* そのまま */ } });
      embedFaces = []; fontMetrics = {};
    }
    function wordPath(target) {
      var t = String(target || "").replace(/^\/+/, "");
      if (t.indexOf("word/") === 0) return t;
      var out = [];
      ("word/" + t).split("/").forEach(function (x) { if (x === "..") out.pop(); else if (x && x !== ".") out.push(x); });
      return out.join("/");
    }
    // ECMA-376 の .odttf: 先頭 32 バイトを、fontKey（GUID）を逆順にした 16 バイトで 2 回 XOR する
    function deobfuscateFont(bytes, key) {
      var g = String(key || "").replace(/[{}-]/g, "");
      if (g.length !== 32) return bytes;
      var out = new Uint8Array(bytes), kb = [];
      for (var i = 0; i < 16; i++) kb.push(parseInt(g.substr((15 - i) * 2, 2), 16));
      for (i = 0; i < 32 && i < out.length; i++) out[i] ^= kb[i % 16];
      return out;
    }
    function fontMetricsOf(font) {
      var os2 = (font.tables && font.tables.os2) || {}, hhea = (font.tables && font.tables.hhea) || {};
      var upm = +font.unitsPerEm || 1000, useTypo = (+os2.fsSelection || 0) & 128;
      var asc = useTypo && isFinite(+os2.sTypoAscender) ? +os2.sTypoAscender : (isFinite(+os2.usWinAscent) ? +os2.usWinAscent : (+font.ascender || 0));
      var desc = useTypo && isFinite(+os2.sTypoDescender) ? +os2.sTypoDescender : (isFinite(+os2.usWinDescent) ? -(+os2.usWinDescent) : (+font.descender || 0));
      var gap = useTypo && isFinite(+os2.sTypoLineGap) ? +os2.sTypoLineGap : (+hhea.lineGap || 0);
      var single = (asc - desc + gap) / upm;
      return { single: single > 0.5 && single < 3 ? Math.round(single * 1000) / 1000 : 1.15, ascender: asc, descender: desc, lineGap: gap, unitsPerEm: upm };
    }
    function singleLine(name) {
      var m = fontMetrics[name] || fontMetrics[String(name || "").toLowerCase()];
      return m ? m.single : 1.15;
    }
    function lineCss(ln, name) { return Math.round((+ln || 0) / 240 * singleLine(name) * 100) / 100; }
    function loadEmbeddedFonts(z) {
      clearEmbeddedFonts();
      var ft = Zip.xml(z["word/fontTable.xml"]);
      if (!ft) return;
      var relDoc = Zip.xml(z["word/_rels/fontTable.xml.rels"]), rels = {};
      kids(relDoc, "Relationship").forEach(function (r) { rels[r.getAttribute("Id")] = r.getAttribute("Target"); });
      var op = window.opentype;
      kids(ft, "font").forEach(function (f) {
        var name = attr(f, "name");
        if (!name) return;
        ["embedRegular", "embedBold", "embedItalic", "embedBoldItalic"].forEach(function (tag) {
          var e = kid(f, tag);
          if (!e) return;
          var rid = e.getAttributeNS(REL_NS, "id") || e.getAttribute("r:id"), target = rels[rid];
          if (!target) return;
          var path = wordPath(target), bytes = z[path];
          if (!bytes) return;
          var key = attr(e, "fontKey") || attr(f, "fontKey");
          if (key || /\.odttf$/i.test(path)) bytes = deobfuscateFont(bytes, key);
          var buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
          // メトリクスは regular から。読めないフォントは、今までの 1.15 にまかせる
          if (tag === "embedRegular" && op) {
            try { fontMetrics[name] = fontMetrics[String(name).toLowerCase()] = fontMetricsOf(op.parse(buf)); } catch (err) { /* そのまま */ }
          }
          // 画面表示用。元のファミリー名で登録するので、本文の font-family はそのまま使える
          if (window.FontFace && document.fonts) {
            try {
              var face = new FontFace(String(name).replace(/["\\]/g, ""), buf, { weight: /Bold/i.test(tag) ? "bold" : "normal", style: /Italic/i.test(tag) ? "italic" : "normal" });
              document.fonts.add(face); embedFaces.push(face);
              face.load().then(function () { if (page) layout(); }).catch(function () { /* そのまま */ });
            } catch (err) { /* そのまま */ }
          }
        });
      });
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
      cstSheet();
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
      var mx = p.mx != null ? p.mx : m.x, my = p.my != null ? p.my : m.y;
      paper.style.setProperty("--pw", d.w + "px");
      paper.style.setProperty("--ph", d.h + "px");
      paper.style.setProperty("--mx", mx + "px");
      paper.style.setProperty("--my", my + "px");
      // 読み込んだ .docx の docDefaults（本文の標準の大きさ・行間・フォント）
      var b = p.base || {};
      page.style.fontSize = b.fs ? b.fs + "pt" : "";
      page.style.lineHeight = b.lh ? String(b.lh) : "";
      page.style.fontFamily = b.ff ? O.fontCss(b.ff) : "";
      // デザイン（見出しの色・フォントなどを CSS の変数にして、用紙に付ける。印刷でも同じ）
      var dz = DESIGNS[p.design] || {};
      [["--dd-h1", dz.h1], ["--dd-h2", dz.h2], ["--dd-h3", dz.h3], ["--dd-ac", dz.ac], ["--dd-th", dz.th], ["--dd-hf", dz.hf], ["--dd-bf", dz.bf], ["--dd-h1a", dz.h1c ? "center" : ""]].forEach(function (x) {
        if (x[1]) paper.style.setProperty(x[0], x[1]); else paper.style.removeProperty(x[0]);
      });
      if (dz.rule) paper.setAttribute("data-rule", dz.rule); else paper.removeAttribute("data-rule");
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
      } else if (cmd === "borders") {
        tableBorderDialog(tbl).then(function (v) { if (v) { applyBorders(tbl, v.kind, v.color, v.width); changed(); O.sync(); } });
        return;
      } else if (cmd === "colWidth") {
        tableColDialog(tbl).then(function (w) { if (w) { applyColWidths(tbl, w); changed(); O.sync(); } });
        return;
      }
      changed();
      O.sync();
    }

    // ---------- 表の罫線・列の幅 ----------
    var BD_KINDS = [["none", "罫線なし"], ["outer", "外枠のみ"], ["inner", "内線のみ"], ["all", "外枠と内線"]];
    var BD_WIDTHS = [["0.5", "細（0.5pt）"], ["1", "普通（1pt）"], ["2.5", "太い（2.5pt）"]];
    // 罫線の状態を読み取る。{ none: true } は「ここには線を引かない」
    function borderSide(el, side) {
      var st = el.style["border" + side + "Style"];
      if (st === "none" || st === "hidden") return { none: true };
      var w = el.style["border" + side + "Width"], c = el.style["border" + side + "Color"];
      if (!st || !w || w === "0px") return null;
      return { w: ptOf(w), color: cssColor(c) || "333333", style: st };
    }
    function borderSet(el, side) { return !!el.style["border" + side + "Style"]; }
    // 罫線の種類を、表のすべてのセルに書き込む（外面だけ・内側だけ・全部）
    function applyBorders(tbl, kind, color, wpt) {
      var rows = Array.prototype.slice.call(tbl.rows), cols = colCount(tbl);
      var line = wpt + "pt solid " + color;
      rows.forEach(function (tr, ri) {
        var col = 0;
        Array.prototype.slice.call(tr.cells).forEach(function (c) {
          var span = c.colSpan || 1, first = col === 0, lastCol = col + span >= cols;
          var top = false, right = false, bottom = false, left = false;
          if (kind === "all") { top = right = bottom = left = true; }
          else if (kind === "outer") { top = ri === 0; bottom = ri === rows.length - 1; left = first; right = lastCol; }
          else if (kind === "inner") { bottom = ri < rows.length - 1; right = !lastCol; }
          c.style.borderTop = top ? line : "none";
          c.style.borderRight = right ? line : "none";
          c.style.borderBottom = bottom ? line : "none";
          c.style.borderLeft = left ? line : "none";
          if (!c.getAttribute("style")) c.removeAttribute("style");
          col += span;
        });
      });
      tbl.setAttribute("data-bd", kind);
    }
    // いまの罫線の状態を読み取る（ダイアログの初期値）
    function readBorders(tbl) {
      var out = { kind: "none", color: "#333333", width: "1" }, cells = tbl.querySelectorAll("td,th");
      if (!cells.length) return out;
      var all = true, any = false;
      Array.prototype.forEach.call(cells, function (c) {
        ["Top", "Right", "Bottom", "Left"].forEach(function (s) {
          var b = borderSide(c, s);
          if (b && !b.none) { any = true; out.color = b.color; out.width = String(b.w); } else all = false;
        });
      });
      out.kind = all && any ? "all" : any ? "outer" : "none";
      return out;
    }
    function tableBorderDialog(tbl) {
      var cur = readBorders(tbl), body = document.createElement("div");
      body.innerHTML = '<div class="dlg-grid">' +
        '<label for="tb-kind">' + esc(t("罫線")) + '</label><select id="tb-kind">' +
          BD_KINDS.map(function (k) { return '<option value="' + k[0] + '"' + (k[0] === cur.kind ? " selected" : "") + ">" + esc(t(k[1])) + "</option>"; }).join("") + "</select>" +
        '<label for="tb-color">' + esc(t("色")) + '</label><input id="tb-color" type="color" value="' + cur.color + '">' +
        '<label for="tb-width">' + esc(t("太さ")) + '</label><select id="tb-width">' +
          BD_WIDTHS.map(function (w) { return '<option value="' + w[0] + '"' + (w[0] === cur.width ? " selected" : "") + ">" + esc(t(w[1])) + "</option>"; }).join("") + "</select>" +
        '</div><p class="dlg-note">' + esc(t("表の中のすべてのセルに罫線を引きます。")) + "</p>";
      return M3.dialog({ title: t("罫線の設定"), icon: "border_all", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: t("OK"), primary: true, value: function () {
          return { kind: body.querySelector("#tb-kind").value, color: body.querySelector("#tb-color").value, width: body.querySelector("#tb-width").value };
        } }] });
    }
    // 列の数（結合されたセルは数える）
    function colCount(tbl) {
      var cols = 0;
      Array.prototype.slice.call(tbl.rows).forEach(function (tr) { var k = 0; Array.prototype.slice.call(tr.cells).forEach(function (c) { k += c.colSpan || 1; }); cols = Math.max(cols, k); });
      return cols;
    }
    // 列の幅（ピクセル）
    function colWidths(tbl) {
      var n = colCount(tbl), out = [], cg = tbl.querySelector("colgroup");
      for (var i = 0; i < n; i++) {
        var col = cg && cg.children[i], w = 0;
        if (col && col.style.width) w = Math.round(ptOf(col.style.width) / 0.75);
        out.push(w);
      }
      return out;
    }
    function applyColWidths(tbl, widths) {
      var cg = tbl.querySelector("colgroup");
      if (!cg) { cg = document.createElement("colgroup"); tbl.insertBefore(cg, tbl.firstChild); }
      cg.innerHTML = "";
      var total = 0;
      widths.forEach(function (w) {
        var col = document.createElement("col");
        if (w > 0) { col.setAttribute("style", "width:" + Math.round(w) + "px"); total += w; }
        cg.appendChild(col);
      });
      if (total > 0) { tbl.style.width = Math.round(total) + "px"; tbl.style.tableLayout = "fixed"; }
      else { tbl.style.width = ""; tbl.style.tableLayout = ""; }
      if (!tbl.getAttribute("style")) tbl.removeAttribute("style");
    }
    function tableColDialog(tbl) {
      var ws = colWidths(tbl), body = document.createElement("div");
      body.innerHTML = '<div class="dlg-grid">' + ws.map(function (w, i) {
        return '<label for="cw' + i + '">' + esc(t("{n} 列目", { n: i + 1 })) + '</label><input id="cw' + i + '" type="number" min="0" max="2000" step="1" value="' + (w || "") + '">';
      }).join("") + '</div><p class="dlg-note">' + esc(t("幅をピクセルで指定します。空欄の列は表に合わせて自動で広がります。")) + "</p>";
      return M3.dialog({ title: t("列の幅"), icon: "width", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: t("OK"), primary: true, value: function () {
          return ws.map(function (w, i) { var n = parseFloat(body.querySelector("#cw" + i).value); return isFinite(n) && n > 0 ? Math.min(2000, Math.round(n)) : 0; });
        } }] });
    }

    // ---------- 段落の前後の間隔 ----------
    // CSS の値（"12pt" など）をポイントに直す。px なら 0.75 倍
    function ptOf(v) {
      var m = /^\s*(-?[\d.]+)\s*(pt|px)?\s*$/.exec(String(v === undefined || v === null ? "" : v));
      if (!m) return 0;
      var n = parseFloat(m[1]);
      if (!isFinite(n)) return 0;
      return m[2] === "px" ? n * 0.75 : n;
    }
    // CSS の長さを px に直す（画像の幅・位置は px で持っている）
    function pxOf(v) {
      var m = /^\s*(-?[\d.]+)\s*(pt|px)?\s*$/.exec(String(v === undefined || v === null ? "" : v));
      if (!m) return 0;
      var n = parseFloat(m[1]);
      if (!isFinite(n)) return 0;
      return m[2] === "pt" ? n * 4 / 3 : n;
    }
    // CSS の色（"#abc"・"rgb(1, 2, 3)"）を "RRGGBB" にする。読めないときは null
    function cssColor(c) {
      if (!c) return null;
      var m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c);
      if (m) { if (m[4] !== undefined && +m[4] === 0) return null; return [m[1], m[2], m[3]].map(function (x) { return ("0" + (+x).toString(16)).slice(-2); }).join("").toUpperCase(); }
      var h = /^#([0-9a-f]{6})$/i.exec(c); if (h) return h[1].toUpperCase();
      h = /^#([0-9a-f]{3})$/i.exec(c); if (h) return h[1].split("").map(function (x) { return x + x; }).join("").toUpperCase();
      return null;
    }
    function spaceOf(b) { return { before: ptOf(b.style.marginTop), after: ptOf(b.style.marginBottom) }; }
    function spaceDialog(cur) {
      var body = document.createElement("div");
      body.innerHTML = '<div class="dlg-grid">' +
        '<label for="sp-before">' + esc(t("上の間隔")) + '</label>' +
        '<input id="sp-before" type="number" min="0" max="400" step="1" value="' + cur.before + '">' +
        '<label for="sp-after">' + esc(t("下の間隔")) + '</label>' +
        '<input id="sp-after" type="number" min="0" max="400" step="1" value="' + cur.after + '">' +
        '</div><p class="dlg-note">' + esc(t("ポイント（pt）で指定します。0 で標準に戻ります。")) + "</p>";
      return M3.dialog({ title: t("段落の前後の間隔"), icon: "density_medium", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: t("OK"), primary: true, value: function () {
          function num(id) { var n = parseFloat(body.querySelector(id).value); return isFinite(n) ? Math.max(0, Math.min(400, n)) : 0; }
          return { before: num("#sp-before"), after: num("#sp-after") };
        } }] });
    }
    function paraSpace() {
      var bs = blocks();
      if (!bs.length) { O.notify(t("段落を選んでください")); return; }
      spaceDialog(spaceOf(bs[0])).then(function (v) {
        if (!v) return;
        bs.forEach(function (b) {
          b.style.marginTop = v.before ? v.before + "pt" : "";
          b.style.marginBottom = v.after ? v.after + "pt" : "";
          if (!b.getAttribute("style")) b.removeAttribute("style");
        });
        changed();
      });
    }

    // ---------- カスタムスタイル ----------
    // data.styles: [{ id, name, base: "p|h1|h2|h3|blockquote", css: "color:#…;font-size:14pt" }]
    // 本文には data-cst="id" を付ける。書式は <style> を #paper に入れて当てる（data.html には入れない）
    function cstList(d) { return (d && Array.isArray(d.styles)) ? d.styles : []; }
    function cstById(id) { return cstList(data).filter(function (s) { return s.id === id; })[0] || null; }
    // 入力欄 → css 文字列
    function cstBuild(f) {
      var css = [];
      if (f.size) css.push("font-size:" + f.size + "pt");
      if (f.color) css.push("color:" + f.color);
      if (f.bg) css.push("background-color:" + f.bg);
      if (f.b) css.push("font-weight:bold");
      if (f.i) css.push("font-style:italic");
      if (f.u) css.push("text-decoration:underline");
      if (f.align && f.align !== "left") css.push("text-align:" + f.align);
      return css.join(";");
    }
    // css 文字列 → 入力欄
    function cstFields(css) {
      var f = { size: "", color: "", bg: "", b: 0, i: 0, u: 0, align: "left" };
      String(css || "").split(";").forEach(function (x) {
        var m = /^\s*([\w-]+)\s*:\s*(.+?)\s*$/.exec(x);
        if (!m) return;
        var k = m[1].toLowerCase(), v = m[2], c = cssColor(v);
        if (k === "font-size") { var n = ptOf(v); if (n) f.size = String(Math.round(n * 10) / 10); }
        else if (k === "color") f.color = c ? "#" + c.toLowerCase() : "";
        else if (k === "background-color") f.bg = c ? "#" + c.toLowerCase() : "";
        else if (k === "font-weight") f.b = /^(bold|[6-9]00)$/.test(v) ? 1 : 0;
        else if (k === "font-style") f.i = /italic/.test(v) ? 1 : 0;
        else if (k === "text-decoration" || k === "text-decoration-line") f.u = /underline/.test(v) ? 1 : 0;
        else if (k === "text-align") f.align = v;
      });
      return f;
    }
    // カスタムスタイルの css を <style> に入れて当てる（#paper の中なので data.html には混ざらない）
    function cstSheet() {
      var paper = el("paper");
      if (!paper) return;
      var st = paper.querySelector("style[data-cst]");
      if (!st) { st = document.createElement("style"); st.setAttribute("data-cst", "1"); paper.insertBefore(st, paper.firstChild); }
      st.textContent = cstList(data).filter(function (s) { return s.css; }).map(function (s) {
        return "#page [data-cst=\"" + s.id + "\"] { " + s.css + " }";
      }).join("\n");
    }
    // スタイルを、選んでる段落に適用する
    function cstApply(st) {
      var bs = blocks();
      if (!bs.length) { O.notify(t("段落を選んでください")); return; }
      bs.forEach(function (b) {
        b.setAttribute("data-cst", st.id);
        if (/^(P|H1|H2|H3|BLOCKQUOTE)$/.test(b.tagName)) b.tagName = (CST_TAG[st.base] || "p").toUpperCase();
      });
      changed();
    }
    // いま当てはまっているカスタムスタイル（リボンのオン・オフ用）
    function cstCurrent() {
      var a = anchorEl(), b = a && a.closest && a.closest("[data-cst]");
      return b ? b.getAttribute("data-cst") : "";
    }
    // スタイルを 1 つ作る・直す
    function cstEditor(st) {
      var f = cstFields(st ? st.css : ""), body = document.createElement("div");
      body.innerHTML = '<div class="dlg-grid">' +
        '<label for="cs-name">' + esc(t("名前")) + '</label><input id="cs-name" type="text" maxlength="40" value="' + esc(st ? st.name : "") + '">' +
        '<label for="cs-base">' + esc(t("元のスタイル")) + '</label><select id="cs-base">' +
          Object.keys(CST_LABEL).map(function (k) { return '<option value="' + k + '"' + (k === (st ? st.base : "p") ? " selected" : "") + ">" + esc(t(CST_LABEL[k])) + "</option>"; }).join("") + "</select>" +
        '<label for="cs-size">' + esc(t("文字の大きさ")) + '</label><input id="cs-size" type="number" min="6" max="96" step="0.5" value="' + (f.size || "") + '" placeholder="' + esc(t("標準")) + '">' +
        '<label for="cs-color">' + esc(t("文字の色")) + '</label><input id="cs-color" type="color" value="' + (f.color || "#1a1c20") + '">' +
        '<label for="cs-bg">' + esc(t("背景の色")) + '</label><input id="cs-bg" type="color" value="' + (f.bg || "#ffffff") + '">' +
        "<label>" + esc(t("太字")) + '</label><input id="cs-b" type="checkbox"' + (f.b ? " checked" : "") + ">" +
        "<label>" + esc(t("斜体")) + '</label><input id="cs-i" type="checkbox"' + (f.i ? " checked" : "") + ">" +
        "<label>" + esc(t("下線")) + '</label><input id="cs-u" type="checkbox"' + (f.u ? " checked" : "") + ">" +
        '<label for="cs-align">' + esc(t("配置")) + '</label><select id="cs-align">' +
          [["left", t("左揃え")], ["center", t("中央揃え")], ["right", t("右揃え")], ["justify", t("両端揃え")]].map(function (a) { return '<option value="' + a[0] + '"' + (a[0] === f.align ? " selected" : "") + ">" + esc(a[1]) + "</option>"; }).join("") + "</select>" +
        '</div><p class="dlg-note">' + esc(t("ここで選んだ書式が、スタイルを適用した段落に効きます。背景の色は白のときなしになります。")) + "</p>";
      return M3.dialog({ title: st ? t("スタイルを編集") : t("新しいスタイル"), icon: "style", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: t("保存"), primary: true, value: function () {
          var name = body.querySelector("#cs-name").value.trim();
          if (!name) return null;
          var bg = body.querySelector("#cs-bg").value;
          return { id: st ? st.id : "", name: name, base: body.querySelector("#cs-base").value, css: cstBuild({
            size: parseFloat(body.querySelector("#cs-size").value) || 0,
            color: body.querySelector("#cs-color").value,
            bg: bg === "#ffffff" ? "" : bg,
            b: body.querySelector("#cs-b").checked ? 1 : 0,
            i: body.querySelector("#cs-i").checked ? 1 : 0,
            u: body.querySelector("#cs-u").checked ? 1 : 0,
            align: body.querySelector("#cs-align").value
          }) };
        } }] });
    }
    // スタイルの管理（一覧・適用・編集・削除）
    function cstManager() {
      var body = document.createElement("div"), closeFn = null;
      function draw() {
        var list = cstList(data);
        body.innerHTML = '<p class="dlg-note" style="margin:0 0 8px">' + esc(t("スタイルを選んで、段落に適用します。")) + "</p>" +
          (list.length ? list.map(function (s) {
            return '<div class="cst-row"><span class="cst-row__name">' + esc(s.name) + '</span><span class="cst-row__base">' + esc(t(CST_LABEL[s.base] || "段落")) + "</span>" +
              '<button type="button" class="text-btn m3-state" data-cst-act="apply" data-id="' + esc(s.id) + '">' + esc(t("適用")) + "</button>" +
              '<button type="button" class="text-btn m3-state" data-cst-act="edit" data-id="' + esc(s.id) + '">' + esc(t("編集")) + "</button>" +
              '<button type="button" class="text-btn m3-state" data-cst-act="del" data-id="' + esc(s.id) + '">' + esc(t("削除")) + "</button></div>";
          }).join("") : '<p class="dlg-note">' + esc(t("カスタムスタイルはまだありません。")) + "</p>") +
          '<button type="button" class="text-btn m3-state" data-cst-act="new">' + esc(t("新しいスタイル")) + "</button>";
      }
      draw();
      body.addEventListener("click", function (e) {
        var b = e.target.closest("[data-cst-act]");
        if (!b) return;
        var act = b.getAttribute("data-cst-act"), id = b.getAttribute("data-id");
        if (act === "apply") { var s = cstById(id); if (s) cstApply(s); return; }
        if (act === "del") {
          var idx = -1; cstList(data).forEach(function (x, i) { if (x.id === id) idx = i; });
          if (idx < 0) return;
          data.styles.splice(idx, 1);
          if (page) page.querySelectorAll('[data-cst="' + id + '"]').forEach(function (n) { n.removeAttribute("data-cst"); });
          cstSheet(); changed(); draw();
          return;
        }
        if (closeFn) closeFn({ act: act, id: id });
      });
      return M3.dialog({ title: t("スタイルの管理"), icon: "style", body: body,
        actions: [{ label: t("閉じる"), value: null, primary: true }],
        onReady: function (close) { closeFn = close; } }).then(function (r) {
          if (!r) return;
          if (r.act === "new") return cstEditor(null).then(function (s) { if (s) addCst(s); return cstManager(); });
          if (r.act === "edit") return cstEditor(cstById(r.id)).then(function (s) { if (s) addCst(s); return cstManager(); });
        });
    }
    // スタイルを足す・直す
    function addCst(s) {
      if (!s || !s.name) return;
      if (s.id) {
        var i = -1; cstList(data).forEach(function (x, k) { if (x.id === s.id) i = k; });
        if (i < 0) data.styles.push({ id: s.id, name: s.name, base: s.base, css: s.css });
        else data.styles[i] = { id: s.id, name: s.name, base: s.base, css: s.css };
      } else {
        data.styles.push({ id: "cst" + Date.now().toString(36) + Math.floor(Math.random() * 1000), name: s.name, base: s.base, css: s.css });
      }
      cstSheet(); changed();
    }
    // リボンのスタイル一覧（標準・見出し… ＋ カスタムスタイル）
    function cstRibbon() {
      var out = [{ tag: "p", label: t("標準") }, { tag: "h1", label: t("表題") }, { tag: "h2", label: t("見出し 1") }, { tag: "h3", label: t("見出し 2") }, { tag: "blockquote", label: t("引用") }];
      cstList(data).forEach(function (s) { out.push({ tag: CST_TAG[s.base] || "p", label: s.name, cmd: "cst:" + s.id, toggle: "cst:" + s.id }); });
      return out;
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
      if (cmd === "designMenu") {
        var dp = pageOf(data), items = [{ head: DESIGN_GROUPS[0][1] }, { value: "", label: t("標準"), color: "#1d4ed8", checked: !dp.design }];
        DESIGN_GROUPS.slice(1).forEach(function (g) {
          items.push({ head: g[1] });
          Object.keys(DESIGNS).filter(function (k) { return DESIGNS[k].group === g[0]; }).forEach(function (k) { items.push({ value: k, label: DESIGNS[k].label, color: DESIGNS[k].h2, checked: dp.design === k }); });
        });
        return O.menu(at, items).then(function (v) { if (v !== null && v !== undefined) command("design:" + v); });
      }
      if (cmd.indexOf("design:") === 0) { var dq = pageOf(data), id = cmd.slice(7); if (id) dq.design = id; else delete dq.design; layout(); O.changed(); O.sync(); return; }
      if (cmd.indexOf("fmt:") === 0) return exec("formatBlock", cmd.slice(4));
      if (cmd.indexOf("cst:") === 0) { var st = cstById(cmd.slice(4)); return st ? cstApply(st) : null; }
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
        case "paraSpace": return paraSpace();
        case "cstManager": return cstManager();
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
          return O.menu(at, Object.keys(MARGINS).map(function (k) { return { value: k, label: MARGINS[k].label, checked: p.margin === k && p.mx == null && p.my == null }; })).then(function (v) { if (v) { p.margin = v; delete p.mx; delete p.my; layout(); O.changed(); stat(); } });
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
           { icon: "format_indent_decrease", label: t("インデントを減らす"), cmd: "outdent" }, { icon: "format_indent_increase", label: t("インデントを増やす"), cmd: "indent" }, { icon: "format_line_spacing", label: t("行と段落の間隔"), cmd: "spacing", drop: 1 },
           { icon: "density_medium", label: t("段落の前後の間隔"), cmd: "paraSpace" }],
          [{ icon: "format_align_left", label: t("左揃え"), sub: "Ctrl+L", cmd: "justifyLeft", toggle: "justifyLeft" }, { icon: "format_align_center", label: t("中央揃え"), sub: "Ctrl+E", cmd: "justifyCenter", toggle: "justifyCenter" },
           { icon: "format_align_right", label: t("右揃え"), sub: "Ctrl+R", cmd: "justifyRight", toggle: "justifyRight" }, { icon: "format_align_justify", label: t("両端揃え"), sub: "Ctrl+J", cmd: "justifyFull", toggle: "justifyFull" }]
        ] }] },
        { label: t("スタイル"), items: [{ styles: cstRibbon }, { stack: [{ icon: "style", label: t("カスタムスタイル"), cmd: "cstManager" }] }] },
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
      { tab: t("デザイン"), groups: [
        { label: t("ドキュメントのデザイン"), items: [{ big: 1, icon: "palette", label: t("デザイン"), cmd: "designMenu", drop: 1 }, { big: 1, icon: "format_paint", label: t("標準"), cmd: "design:", toggle: "design:" }]
          .concat(["business", "modern", "report", "warm", "nature", "washi"].map(function (k) { return { big: 1, icon: "palette", label: DESIGNS[k].label, cmd: "design:" + k, toggle: "design:" + k }; })) }
      ] },
      { tab: t("レイアウト"), groups: [
        { label: t("ページ設定"), items: [{ big: 1, icon: "margin", label: t("余白"), cmd: "margin", drop: 1 }, { big: 1, icon: "crop_rotate", label: t("印刷の向き"), cmd: "orient", drop: 1 },
          { big: 1, icon: "description", label: t("サイズ"), cmd: "size", drop: 1 }, { big: 1, icon: "view_column", label: t("段組み"), cmd: "cols", drop: 1 }, { big: 1, icon: "text_rotate_vertical", label: t("縦書き"), cmd: "vert", toggle: "vert" }] },
        { label: t("段落"), items: [{ stack: [{ icon: "format_line_spacing", label: t("行間"), cmd: "spacing", drop: 1 }, { icon: "density_medium", label: t("段落の前後の間隔"), cmd: "paraSpace" }, { icon: "format_indent_increase", label: t("インデント"), cmd: "indent" }, { icon: "insert_page_break", label: t("改ページ"), cmd: "pagebreak" }] }] }
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
        { label: t("書式"), items: [{ stack: [{ icon: "table_rows", label: t("見出しの行"), cmd: "tbl:headRow" }, { icon: "format_color_fill", label: t("塗りつぶし"), cmd: "tbl:shade", drop: 1 }, { icon: "cell_merge", label: t("右のセルと結合"), cmd: "tbl:merge" }] }] },
        { label: t("罫線と幅"), items: [{ stack: [{ icon: "border_all", label: t("罫線の設定"), cmd: "tbl:borders" }, { icon: "width", label: t("列の幅"), cmd: "tbl:colWidth" }] }] }
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
      s["design:"] = !pageOf(data).design;
      Object.keys(DESIGNS).forEach(function (k) { s["design:" + k] = pageOf(data).design === k; });
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
        var cid = b && b.getAttribute && b.getAttribute("data-cst");
        if (cid) s["cst:" + cid] = true;
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
      var pmx = p.mx != null ? p.mx : m.x, pmy = p.my != null ? p.my : m.y;
      // 読み込んだ docx に同梱フォントがあれば、そのメトリクスで行間を戻す（無ければ 1.15）
      var baseSingle = singleLine((p.base || {}).ff);
      var holder = document.createElement("div");
      holder.innerHTML = data.html;
      // カスタムスタイル（data.styles）
      var cstIndex = {};
      (data.styles || []).forEach(function (s) { cstIndex[s.id] = s; });
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
        function color(c) { return cssColor(c); }
        // 文字の書式。Word は要素の順番が決まっている（rStyle → rFonts → b → i → strike → color → sz → u → shd → vertAlign）。順番がちがうと開けない
        function runProps(st, rStyle) {
          var x = rStyle ? '<w:rStyle w:val="' + rStyle + '"/>' : "";
          if (st.code) x += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"' + (st.font ? ' w:eastAsia="' + X(st.font) + '"' : "") + "/>";
          else if (st.font) { var f = X(st.font); x += '<w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '" w:eastAsia="' + f + '"/>'; }
          if (st.b) x += "<w:b/>";
          if (st.i) x += "<w:i/>";
          if (st.s) x += "<w:strike/>";
          if (st.color || st.link) x += '<w:color w:val="' + (st.color || "1D4ED8") + '"/>';
          if (st.size) x += '<w:sz w:val="' + Math.round(st.size * 2) + '"/><w:szCs w:val="' + Math.round(st.size * 2) + '"/>';
          if (st.u || st.link) x += '<w:u w:val="single"/>';
          if (st.hl) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + st.hl + '"/>';
          if (st.sub || st.sup) x += '<w:vertAlign w:val="' + (st.sub ? "subscript" : "superscript") + '"/>';
          return x ? "<w:rPr>" + x + "</w:rPr>" : "";
        }
        // 脚注（本文の <sup class="fn-ref" data-fn> と、最後の <section class="footnotes"> の <li data-fn>）は、Word の脚注（footnotes.xml）にする
        var fnText = {}, fnXml = [], fnNo = 0;
        holder.querySelectorAll("section.footnotes li[data-fn]").forEach(function (li) { fnText[li.getAttribute("data-fn")] = li; });
        // コメント（<span class="cmt" data-cmt> と data.comments）は、Word のコメント（comments.xml）にする。同じコメントの span が分かれていても、範囲は 1 つ
        var cmtById = {}, cmtNo = {}, cmtLeft = {}, cmtXml = [];
        (data.comments || []).forEach(function (c) { cmtById[c.id] = c; });
        holder.querySelectorAll("span.cmt[data-cmt]").forEach(function (s) { var id = s.getAttribute("data-cmt"); if (cmtById[id]) cmtLeft[id] = (cmtLeft[id] || 0) + 1; });
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
          var maxW = (land ? sz.h : sz.w) - pmx * 2;
          var dw = +img.getAttribute("data-w") || Math.round(pxOf(img.style.width)) || 300;
          var dh = +img.getAttribute("data-h") || Math.round(pxOf(img.style.height)) || Math.round(dw * 0.75) || 200;
          var w = Math.min(dw, maxW), h = dh * (w / dw);
          var cx = Math.round(w * 9525), cy = Math.round(h * 9525), n = media.length;
          var graphic = '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
            '<pic:nvPicPr><pic:cNvPr id="' + n + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + id + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
          var docPr = '<wp:docPr id="' + n + '" name="Picture ' + n + '" descr="' + X(img.getAttribute("alt") || "") + '"/>';
          // 自由配置（data-anc）は、Word の「文字列の折り返しなし」のアンカーで書き出す。
          // 画面では段落を基準に left/top で置いているので、column / paragraph 基準の posOffset にする
          if (img.getAttribute("data-anc")) {
            var ax = Math.round(pxOf(img.style.left) * 9525), ay = Math.round(pxOf(img.style.top) * 9525);
            return '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' + (251658240 + n) + '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
              '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>' + ax + '</wp:posOffset></wp:positionH>' +
              '<wp:positionV relativeFrom="paragraph"><wp:posOffset>' + ay + '</wp:posOffset></wp:positionV>' +
              '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' + docPr +
              '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' + graphic + '</wp:anchor></w:drawing></w:r>';
          }
          return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' + docPr +
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
            if (tag === "SUP" && n.classList.contains("fn-ref")) {
              var li = fnText[n.getAttribute("data-fn")];
              if (!li) return;
              var fid = ++fnNo;
              fnXml.push('<w:footnote w:id="' + fid + '"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r>' + runProps({}, "FootnoteReference") + "<w:footnoteRef/></w:r>" +
                '<w:r><w:t xml:space="preserve"> </w:t></w:r>' + inline(li, {}) + "</w:p></w:footnote>");
              out += "<w:r>" + runProps({}, "FootnoteReference") + '<w:footnoteReference w:id="' + fid + '"/></w:r>';
              return;
            }
            if (tag === "SPAN" && n.classList.contains("cmt") && cmtLeft[n.getAttribute("data-cmt")]) {
              var cid = n.getAttribute("data-cmt"), first = !(cid in cmtNo), c = cmtById[cid];
              if (first) {
                cmtNo[cid] = cmtXml.length;
                cmtXml.push('<w:comment w:id="' + cmtNo[cid] + '" w:author="' + X(c.author || "") + '" w:date="' + new Date(c.ts || Date.now()).toISOString().replace(/\.\d+Z$/, "Z") + '" w:initials="' + X(Array.from(c.author || "?")[0]) + '">' +
                  String(c.text || "").split("\n").map(function (l, i) { return "<w:p>" + (i ? "" : "<w:r>" + runProps({}, "CommentReference") + "<w:annotationRef/></w:r>") + textRun(l, {}) + "</w:p>"; }).join("") + "</w:comment>");
              }
              out += (first ? '<w:commentRangeStart w:id="' + cmtNo[cid] + '"/>' : "") + inline(n, s);
              if (!--cmtLeft[cid]) out += '<w:commentRangeEnd w:id="' + cmtNo[cid] + '"/><w:r>' + runProps({}, "CommentReference") + '<w:commentReference w:id="' + cmtNo[cid] + '"/></w:r>';
              return;
            }
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
          // 順番は spacing → jc（Word の決まり）。間隔は 20 分の 1 ポイントで書く
          var sp = "", bf = ptOf(css.marginTop), af = ptOf(css.marginBottom), lh = parseFloat(css.lineHeight);
          if (bf) sp += ' w:before="' + Math.round(bf * 20) + '"';
          if (af) sp += ' w:after="' + Math.round(af * 20) + '"';
          if (lh) sp += ' w:line="' + Math.round(lh * 240 / baseSingle) + '" w:lineRule="auto"';
          if (sp) x += "<w:spacing" + sp + "/>";
          if (al) x += '<w:jc w:val="' + ({ center: "center", right: "right", end: "right", justify: "both", left: "left", start: "left" }[al] || "left") + '"/>';
          return x ? "<w:pPr>" + x + "</w:pPr>" : "";
        }
        function para(n, style, extra) {
          // カスタムスタイルを当てている段落は、そのスタイルを pStyle にする
          var cst = n.getAttribute && n.getAttribute("data-cst"), sid = cst && cstIndex[cst] ? cst : style;
          var pp = (sid ? '<w:pStyle w:val="' + sid + '"/>' : "") + (extra || "");
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
          var full = ((land ? sz.th : sz.tw) - pmx * 15 * 2);
          // 列の幅（<colgroup> に px で書いてある）。指定がなければ幅を等しく分ける
          var cg = n.querySelector("colgroup"), g = [], fixed = false;
          for (var i = 0; i < cols; i++) {
            var col = cg && cg.children[i], w = col ? Math.round(ptOf(col.style.width) * 20) : 0;
            if (w > 0) fixed = true;
            g.push(w);
          }
          if (!fixed) { var eq = Math.floor(full / Math.max(1, cols)); g = new Array(cols).fill(eq); }
          var total = g.reduce(function (a, b) { return a + b; }, 0);
          // 罫線（セルの inline style に書いてあるもの）。あれば表全体の罫線はなしにして、セルごとの罫線を書く
          var has = !!n.getAttribute("data-bd");
          rows.forEach(function (r) { Array.prototype.forEach.call(r.cells, function (c) { if (borderSet(c, "Top") || borderSet(c, "Left") || borderSet(c, "Bottom") || borderSet(c, "Right")) has = true; }); });
          var x = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblBorders>' +
            ["top", "left", "bottom", "right", "insideH", "insideV"].map(function (b) { return "<w:" + b + ' w:val="' + (has ? "none" : "single") + '" w:sz="4" w:space="0" w:color="BFC3CC"/>'; }).join("") +
            "</w:tblBorders>" + (fixed ? '<w:tblLayout w:type="fixed"/>' : "") + '<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>' +
            g.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join("") + "</w:tblGrid>";
          rows.forEach(function (r, ri) {
            x += "<w:tr>" + (ri === 0 && r.cells[0] && r.cells[0].tagName === "TH" ? "<w:trPr><w:tblHeader/></w:trPr>" : "");
            var col = 0;
            Array.prototype.forEach.call(r.cells, function (c) {
              var span = c.colSpan || 1, bg = color(c.style.backgroundColor) || (c.tagName === "TH" ? "EEF0F6" : null);
              var w = 0; for (var k = 0; k < span; k++) w += g[col + k] || 0;
              var bd = tcBorders(c);
              x += '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : "") + (bd ? "<w:tcBorders>" + bd + "</w:tcBorders>" : "") + (bg ? '<w:shd w:val="clear" w:color="auto" w:fill="' + bg + '"/>' : "") + "</w:tcPr>";
              col += span;
              var inner = blocksXml(c, c.tagName === "TH");
              x += (inner || "<w:p/>") + "</w:tc>";
            });
            x += "</w:tr>";
          });
          return x + "</w:tbl><w:p/>";
        }
        // セルの罫線（inline style → w:tcBorders）。順番は top → left → bottom → right
        function tcBorders(c) {
          var x = "";
          [["Top", "top"], ["Left", "left"], ["Bottom", "bottom"], ["Right", "right"]].forEach(function (o) {
            var b = borderSide(c, o[0]);
            if (!b) return;
            if (b.none) { x += '<w:' + o[1] + ' w:val="none" w:sz="0" w:space="0" w:color="auto"/>'; return; }
            var val = b.style === "dashed" ? "dashed" : b.style === "dotted" ? "dotted" : b.style === "double" ? "double" : "single";
            x += '<w:' + o[1] + ' w:val="' + val + '" w:sz="' + Math.max(2, Math.round(b.w * 8)) + '" w:space="0" w:color="' + b.color + '"/>';
          });
          return x;
        }
        // ブロックの並び（段落・見出し・表…）を変える。inlineRoot: 中にブロックがないときは、そのまま 1 段落にする
        function blocksXml(root, bold) {
          var out = "", loose = null;
          // ブロックのあいだの改行や空白だけのところは、段落にしない
          function flush() { if (loose && (loose.textContent.replace(/[\s​]/g, "") || loose.querySelector("img,br"))) out += "<w:p>" + inline(loose, bold ? { b: 1 } : {}) + "</w:p>"; loose = null; }
          Array.prototype.forEach.call(root.childNodes, function (n) {
            var tag = n.nodeType === 1 ? n.tagName : "";
            if (!/^(P|H1|H2|H3|H4|H5|H6|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|DIV|SECTION|FIGURE|NAV)$/.test(tag)) { (loose = loose || document.createElement("p")).appendChild(n.cloneNode(true)); return; }
            flush();
            if (tag === "SECTION" && n.classList.contains("footnotes")) return;
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
        if (fnXml.length) rel("footnotes", "footnotes.xml");
        if (cmtXml.length) rel("comments", "comments.xml");
        var W = land ? sz.th : sz.tw, H = land ? sz.tw : sz.th, mx = pmx * 15, my = pmy * 15;
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
        // 脚注の区切り線（id -1・0）は Word が使う決まりのもの。settings.xml の footnotePr からも指す
        if (fnXml.length) files["word/footnotes.xml"] = head + "<w:footnotes " + NS + '><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
          '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' + fnXml.join("") + "</w:footnotes>";
        if (cmtXml.length) files["word/comments.xml"] = head + "<w:comments " + NS + ">" + cmtXml.join("") + "</w:comments>";
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
        var dz = DESIGNS[p.design] || {}, bp = p.base || {}, bff = bp.ff || dz.face || "游ゴシック", bsz = Math.round((bp.fs || 12) * 2), blh = Math.round((bp.lh || 1.85) * 240 / baseSingle);
        var font = '<w:rFonts w:ascii="' + X(bff) + '" w:eastAsia="' + X(bff) + '" w:hAnsi="' + X(bff) + '" w:cs="Times New Roman"/>';
        function sty(id, name, ppr, rpr, extra) { return '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' + (extra || "") + (ppr ? "<w:pPr>" + ppr + "</w:pPr>" : "") + (rpr ? "<w:rPr>" + rpr + "</w:rPr>" : "") + "</w:style>"; }
        // デザインの見出し（フォント・色・h2 の線）。rPr の中は、rFonts が先頭
        var hRf = dz.hface ? '<w:rFonts w:ascii="' + X(dz.hface) + '" w:eastAsia="' + X(dz.hface) + '" w:hAnsi="' + X(dz.hface) + '"/>' : "";
        function hClr(c) { return c ? '<w:color w:val="' + c.replace("#", "").toUpperCase() + '"/>' : ""; }
        var acHex = dz.ac ? dz.ac.replace("#", "").toUpperCase() : "";
        var h2Bd = dz.rule === "rule" ? '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="2" w:color="' + acHex + '"/></w:pBdr>' : dz.rule === "bar" ? '<w:pBdr><w:left w:val="single" w:sz="36" w:space="6" w:color="' + acHex + '"/></w:pBdr>' : "";
        var cstStyles = (data.styles || []).map(function (s) {
          var f = cstFields(s.css), ppr = "", rpr = "";
          if (f.b) rpr += "<w:b/>";
          if (f.i) rpr += "<w:i/>";
          if (f.color) rpr += '<w:color w:val="' + f.color.replace("#", "").toUpperCase() + '"/>';
          if (f.size) { var sz = Math.round(f.size * 2); rpr += '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>'; }
          if (f.u) rpr += '<w:u w:val="single"/>';
          if (f.bg) rpr += '<w:shd w:val="clear" w:color="auto" w:fill="' + f.bg.replace("#", "").toUpperCase() + '"/>';
          if (f.align && f.align !== "left") ppr += '<w:jc w:val="' + ({ center: "center", right: "right", justify: "both", left: "left" }[f.align] || "left") + '"/>';
          var base = CST_BASE[s.base] || "Normal";
          return '<w:style w:type="paragraph" w:customStyle="1" w:styleId="' + X(s.id) + '"><w:name w:val="' + X(s.name) + '"/><w:basedOn w:val="' + base + '"/><w:next w:val="Normal"/>' + (ppr ? "<w:pPr>" + ppr + "</w:pPr>" : "") + (rpr ? "<w:rPr>" + rpr + "</w:rPr>" : "") + "</w:style>";
        }).join("");
        files["word/styles.xml"] = head + "<w:styles " + NS + "><w:docDefaults><w:rPrDefault><w:rPr>" + font + '<w:kern w:val="2"/><w:sz w:val="' + bsz + '"/><w:szCs w:val="' + bsz + '"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="180" w:line="' + blh + '" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
          sty("Title", "Title", '<w:spacing w:after="240"/><w:jc w:val="' + (dz.h1c ? "center" : "left") + '"/>', hRf + '<w:b/>' + hClr(dz.h1) + '<w:sz w:val="44"/><w:szCs w:val="44"/>') +
          sty("Heading1", "heading 1", '<w:keepNext/>' + h2Bd + '<w:spacing w:before="360" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/>', hRf + '<w:b/>' + hClr(dz.h2 || "#1d4ed8") + '<w:sz w:val="32"/><w:szCs w:val="32"/>') +
          sty("Heading2", "heading 2", '<w:keepNext/><w:spacing w:before="240" w:after="80"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/>', hRf + '<w:b/>' + hClr(dz.h3) + '<w:sz w:val="26"/><w:szCs w:val="26"/>') +
          sty("Heading3", "heading 3", '<w:keepNext/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/>', hRf + "<w:b/>" + hClr(dz.h3)) +
          sty("Quote", "Quote", '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="' + (dz.ac ? dz.ac.replace("#", "").toUpperCase() : "C3C6D7") + '"/></w:pBdr><w:ind w:left="284"/>', '<w:color w:val="434655"/>') +
          sty("ListParagraph", "List Paragraph", '<w:spacing w:after="40"/><w:ind w:left="840"/>') +
          sty("Code", "Code", '<w:shd w:val="clear" w:color="auto" w:fill="F1F3F7"/><w:spacing w:after="0"/>', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/>') +
          sty("FootnoteText", "footnote text", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>', '<w:sz w:val="18"/><w:szCs w:val="18"/>') +
          sty("CommentText", "annotation text", '<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>') +
          '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
          '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>' +
          sty("TOCHeading", "TOC Heading", '<w:spacing w:after="120"/>', '<w:b/><w:sz w:val="28"/>') + sty("TOC1", "toc 1", '<w:spacing w:after="40"/>') + sty("TOC2", "toc 2", '<w:spacing w:after="40"/><w:ind w:left="420"/>') +
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' + cstStyles + "</w:styles>";
        files["word/settings.xml"] = head + "<w:settings " + NS + '><w:defaultTabStop w:val="840"/>' + (fnXml.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : "") + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
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
          (fnXml.length ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : "") +
          (cmtXml.length ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : "") +
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
    // デザインの色・フォントを、書き出す HTML の CSS にする（画面は office.css の .doc-page と同じ見た目）
    function designCss(p) {
      var d = DESIGNS[p.design];
      if (!d) return "";
      var f = function (v) { return String(v).replace(/"/g, "'"); };
      return (d.bf ? "body{font-family:" + f(d.bf) + "}" : "") +
        "h1{color:" + d.h1 + (d.h1c ? ";text-align:center" : "") + (d.hf ? ";font-family:" + f(d.hf) : "") + "}h2{color:" + d.h2 + (d.hf ? ";font-family:" + f(d.hf) : "") +
        (d.rule === "rule" ? ";padding-bottom:4px;border-bottom:2px solid " + d.ac : d.rule === "bar" ? ";padding-left:12px;border-left:6px solid " + d.ac : "") + "}h3{color:" + d.h3 + (d.hf ? ";font-family:" + f(d.hf) : "") + "}" +
        "th{background:" + d.th + "}blockquote{border-left-color:" + d.ac + "}a{color:" + d.ac + "}";
    }
    function htmlFile() {
      var name = O.file() ? O.file().name : "", p = pageOf(data), b = p.base || {};
      var baseCss = (b.fs ? "font-size:" + b.fs + "pt;" : "") + (b.lh ? "line-height:" + b.lh + ";" : "") + (b.ff ? "font-family:" + O.fontCss(b.ff).replace(/"/g, "'") + ";" : "");
      return '<!DOCTYPE html>\n<html lang="ja">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>' + esc(name) + "</title>\n<style>\n" +
        "body{max-width:760px;margin:40px auto;padding:0 20px;font-family:\"Yu Gothic\",\"Hiragino Sans\",\"Noto Sans JP\",sans-serif;line-height:1.85;color:#1a1c20;" + baseCss + "}" +
        "h1{font-size:30px;line-height:1.35}h2{font-size:21px;color:#1d4ed8;margin-top:28px}h3{font-size:17px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #d5d6dc;padding:8px 12px;text-align:left;vertical-align:top}th{background:#eef0f6}" +
        "blockquote{margin:0 0 12px;padding-left:16px;border-left:3px solid #c3c6d7;color:#434655}img{max-width:100%;height:auto}.wrap-l{float:left;margin:4px 16px 8px 0}.wrap-r{float:right;margin:4px 0 8px 16px}.wrap-c{display:block;margin:8px auto}" +
        "hr.pagebreak{border:0;break-after:page}.meta{color:#5a5f6b;font-size:14px}.right{text-align:right}.toc{padding:12px 16px;background:#f3f4f8;border-radius:8px}.toc a{color:inherit}.toc__item--h3{padding-left:1.5em}" +
        "table[data-bd] th,table[data-bd] td{border:0}" +
        designCss(p) +
        cstList(data).filter(function (s) { return s.css; }).map(function (s) { return '[data-cst="' + s.id + '"]{' + s.css + "}"; }).join("") +
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
        // 同梱フォントを先に読む（docDefaults の行間を、フォントの ascent/descent から出すため）
        loadEmbeddedFonts(z);
        var relDoc = Zip.xml(z["word/_rels/document.xml.rels"]), rels = {};
        kids(relDoc, "Relationship").forEach(function (r) { rels[r.getAttribute("Id")] = { target: r.getAttribute("Target"), type: r.getAttribute("Type") || "", ext: r.getAttribute("TargetMode") === "External" }; });
        var numDoc = Zip.xml(z["word/numbering.xml"]), absFmt = {}, numAbs = {};
        kids(numDoc, "abstractNum").forEach(function (a) { var id = attr(a, "abstractNumId"); absFmt[id] = {}; kids(a, "lvl").forEach(function (l) { absFmt[id][attr(l, "ilvl")] = attr(kid(l, "numFmt"), "val"); }); });
        kids(numDoc, "num").forEach(function (n) { numAbs[attr(n, "numId")] = attr(kid(n, "abstractNumId"), "val"); });
        var styleDoc = Zip.xml(z["word/styles.xml"]);
        if (styleDoc && styleDoc.documentElement) styleDoc = styleDoc.documentElement;
        var styleName = {}, cstMap = {};
        kids(styleDoc, "style").forEach(function (s) { var nm = kid(s, "name"); styleName[attr(s, "styleId")] = (attr(nm, "val") || "").toLowerCase(); });
        // 本文の標準（docDefaults）と、組み込みの見出しなどのスタイル。Word の標準は 11pt・行間 1.15 など、画面の標準（12pt・1.85）とちがう
        var dd = kid(styleDoc, "docDefaults"), ddR = kid(kid(dd, "rPrDefault"), "rPr"), ddP = kid(kid(dd, "pPrDefault"), "pPr");
        var baseFs = (+attr(kid(ddR, "sz"), "val") || 24) / 2;
        var baseFf = attr(kid(ddR, "rFonts"), "eastAsia") || attr(kid(ddR, "rFonts"), "ascii") || "";
        var baseLn = attr(kid(ddP, "spacing"), "line"), baseLh = baseLn ? lineCss(baseLn, baseFf) : 0;
        var base = { fs: baseFs };
        if (baseFf) base.ff = baseFf;
        if (baseLh) base.lh = baseLh;
        function onEl(e) { return !!e && attr(e, "val") !== "0" && attr(e, "val") !== "false" && attr(e, "val") !== "none"; }
        var pstyleCss = {};
        kids(styleDoc, "style").forEach(function (s) {
          if (attr(s, "type") !== "paragraph") return;
          var cs = attr(s, "customStyle");
          if (cs === "1" || cs === "true" || cs === "on") return;
          var nm = (attr(kid(s, "name"), "val") || "").toLowerCase();
          if (!/^(title|subtitle|heading ?[1-9]|quote)$/.test(nm)) return;
          var ppr = kid(s, "pPr"), rpr = kid(s, "rPr"), css = [];
          var sp = kid(ppr, "spacing"), bf = attr(sp, "before"), af = attr(sp, "after"), ln = attr(sp, "line"), lr = attr(sp, "lineRule");
          var styFf = attr(kid(rpr, "rFonts"), "eastAsia") || attr(kid(rpr, "rFonts"), "ascii") || baseFf;
          if (bf) css.push("margin-top:" + (+bf / 20) + "pt");
          if (af) css.push("margin-bottom:" + (+af / 20) + "pt");
          if (ln && lr !== "exact" && lr !== "atLeast") css.push("line-height:" + lineCss(ln, styFf));
          var sz = attr(kid(rpr, "sz"), "val");
          if (sz) css.push("font-size:" + (+sz / 2) + "pt");
          // 画面の h1・h2 は太字・色つきなので、Word のスタイルに合わせて打ち消す
          css.push("font-weight:" + (onEl(kid(rpr, "b")) ? "bold" : "normal"));
          css.push("font-style:" + (onEl(kid(rpr, "i")) ? "italic" : "normal"));
          css.push("color:" + (hex(attr(kid(rpr, "color"), "val")) || "#000000"));
          var sid = attr(s, "styleId");
          if (sid) pstyleCss[sid] = css.join(";");
        });
        // カスタムスタイル（w:customStyle="1"）は data.styles にする
        var cstStyles = [];
        kids(styleDoc, "style").forEach(function (s) {
          var cs = attr(s, "customStyle");
          if (cs !== "1" && cs !== "true" && cs !== "on") return;
          var id = attr(s, "styleId"), nm = attr(kid(s, "name"), "val");
          if (!id || !nm) return;
          var base = "p", bo = attr(kid(s, "basedOn"), "val");
          Object.keys(CST_BASE).forEach(function (k) { if (CST_BASE[k] === bo) base = k; });
          var ppr = kid(s, "pPr"), rpr = kid(s, "rPr"), css = [];
          var jc = attr(kid(ppr, "jc"), "val");
          if (jc) css.push("text-align:" + (jc === "both" ? "justify" : jc === "end" ? "right" : jc === "start" ? "left" : jc));
          if (onEl(kid(rpr, "b"))) css.push("font-weight:bold");
          if (onEl(kid(rpr, "i"))) css.push("font-style:italic");
          var col = hex(attr(kid(rpr, "color"), "val"));
          if (col) css.push("color:" + col);
          var sz = attr(kid(rpr, "sz"), "val");
          if (sz) css.push("font-size:" + (+sz / 2) + "pt");
          if (kid(rpr, "u")) css.push("text-decoration:underline");
          var shd = hex(attr(kid(rpr, "shd"), "fill"));
          if (shd) css.push("background-color:" + shd);
          var st = { id: id, name: nm, base: base, css: css.join(";") };
          cstMap[id] = st; cstStyles.push(st);
        });
        function blockTag(styleId) {
          var n = (styleName[styleId] || styleId || "").toLowerCase();
          if (/^title$|^表題$/.test(n)) return "h1";
          if (/heading ?1$|見出し ?1$/.test(n)) return "h2";
          if (/heading ?[2-9]$|見出し ?[2-9]$/.test(n)) return "h3";
          if (/quote|引用/.test(n)) return "blockquote";
          if (cstMap[styleId]) return CST_TAG[cstMap[styleId].base] || "p";
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
        // 脚注（footnotes.xml）は、本文の <sup class="fn-ref"> と、最後の <section class="footnotes"> にする
        var fnById = {}, fnOrder = [];
        kids(Zip.xml(z["word/footnotes.xml"]), "footnote").forEach(function (f) { if (!attr(f, "type") || attr(f, "type") === "normal") fnById[attr(f, "id")] = f; });
        // コメント（comments.xml）は、data.comments と、範囲の <span class="cmt">。範囲が段落をまたぐときは、段落ごとに span を分ける
        var cmtById = {}, openCmts = [], stamp = Date.now().toString(36);
        kids(Zip.xml(z["word/comments.xml"]), "comment").forEach(function (c) {
          var id = attr(c, "id"), ts = Date.parse(attr(c, "date") || "");
          cmtById[id] = { id: "c" + stamp + id, author: attr(c, "author") || "", ts: isNaN(ts) ? Date.now() : ts,
            text: kids(c, "p").map(function (p) { return kids(p, "t").map(function (x) { return x.textContent; }).join(""); }).join("\n").trim() };
        });
        var HL = { yellow: "#fff59d", green: "#bbf7d0", cyan: "#a5f3fc", magenta: "#f5d0fe", blue: "#bfdbfe", red: "#fecaca", darkBlue: "#93c5fd", lightGray: "#e5e7eb", darkGray: "#9ca3af" };
        // スタイル（w:style）を ID から引く。basedOn をたどって、自分→もとのスタイルの順に並べる
        var styleById = {};
        kids(styleDoc, "style").forEach(function (s) { styleById[attr(s, "styleId")] = s; });
        function styleChain(id) {
          var chain = [], seen = {};
          while (id && !seen[id] && styleById[id]) { seen[id] = 1; chain.push(styleById[id]); id = attr(kid(styleById[id], "basedOn"), "val"); }
          return chain;
        }
        function runHtml(r) {
          var pr = kid(r, "rPr"), html = "";
          // 文字スタイル（rStyle）の書式。run に直接書いた書式を先に見て、無ければスタイルを見る
          var rchain = styleChain(attr(kid(pr, "rStyle"), "val"));
          function rk(name) {
            var e = kid(pr, name);
            for (var i = 0; i < rchain.length && !e; i++) e = kid(kid(rchain[i], "rPr"), name);
            return e;
          }
          Array.prototype.forEach.call(r.childNodes, function (c) {
            if (c.nodeType !== 1) return;
            var ln = c.localName;
            if (ln === "footnoteReference" && fnById[attr(c, "id")]) { fnOrder.push(attr(c, "id")); html += '<sup class="fn-ref" data-fn="fn' + stamp + attr(c, "id") + '">' + fnOrder.length + "</sup>"; return; }
            if (ln === "t") html += esc(c.textContent);
            else if (ln === "tab") html += "\u3000";
            else if (ln === "br") html += attr(c, "type") === "page" ? '</p><hr class="pagebreak"><p>' : "<br>";
            else if (ln === "drawing" || ln === "pict") {
              var blip = kids(c, "blip")[0], imd = kids(c, "imagedata")[0];
              var id = blip ? blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed") || blip.getAttribute("r:embed") : imd ? imd.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") : null;
              var ext = kids(c, "extent")[0], src = id ? mediaUrl(id) : "";
              if (!src) return;
              var w = ext ? Math.round(+ext.getAttribute("cx") / 9525) : 0, hh = ext ? Math.round(+ext.getAttribute("cy") / 9525) : 0;
              var st = (w ? "width:" + w + "px;" : "") + (hh ? "height:" + hh + "px;" : ""), cls = "";
              // 図の位置（wp:anchor）。Word の「文字列の折り返しなし」は、段落を基準に重ねて置く
              var an = kids(c, "anchor")[0];
              if (an) {
                var posH = kid(an, "positionH"), rel = posH ? (posH.getAttribute("relativeFrom") || "column") : "";
                var pos = posH && kid(posH, "posOffset"), x = pos ? (+pos.textContent || 0) / 9525 : 0;
                if (/rightMargin|outsideMargin/i.test(rel)) x = contentW - w;
                else if (/^page$/i.test(rel)) x -= pageMx;
                if (kids(c, "wrapNone").length) {
                  var posV = kid(an, "positionV"), posY = posV && kid(posV, "posOffset");
                  var y = posY ? (+posY.textContent || 0) / 9525 : 0;
                  st += "position:absolute;left:" + Math.max(0, Math.round(x)) + "px;top:" + Math.max(0, Math.round(y)) + "px;z-index:1;";
                  html += '<img src="' + src + '" alt="" data-anc="1" style="' + st + '">';
                  return;
                }
                if (kids(c, "wrapTopAndBottom").length) cls = "wrap-c";
                else cls = x + w / 2 < contentW * 0.45 ? "wrap-l" : "wrap-r";
                // 右寄せは、左からの位置が近くなるように右の余白で調整する
                if (cls === "wrap-r") { var gap = Math.round(contentW - (x + w)); if (gap > 0 && gap < contentW) st += "margin-right:" + gap + "px;"; }
                else if (cls === "wrap-l" && x > 0) st += "margin-left:" + Math.round(x) + "px;";
              }
              html += '<img src="' + src + '" alt=""' + (cls ? ' class="' + cls + '"' : "") + (st ? ' style="' + st + '"' : "") + ">";
            }
          });
          if (!html || !pr || /^<sup class="fn-ref"[^>]*>\d+<\/sup>$/.test(html)) return html;
          var css = "", open = "", close = "";
          function on(name) { var e = rk(name); return e && attr(e, "val") !== "0" && attr(e, "val") !== "false" && attr(e, "val") !== "none"; }
          if (on("b")) { open += "<b>"; close = "</b>" + close; }
          if (on("i")) { open += "<i>"; close = "</i>" + close; }
          if (on("u")) { open += "<u>"; close = "</u>" + close; }
          if (on("strike") || on("dstrike")) { open += "<s>"; close = "</s>" + close; }
          var va = attr(rk("vertAlign"), "val");
          if (va === "superscript") { open += "<sup>"; close = "</sup>" + close; }
          if (va === "subscript") { open += "<sub>"; close = "</sub>" + close; }
          var col = hex(attr(rk("color"), "val")), sz = attr(rk("sz"), "val"), hl = HL[attr(rk("highlight"), "val")] || hex(attr(rk("shd"), "fill"));
          var fonts = rk("rFonts"), ff = fonts ? attr(fonts, "eastAsia") || attr(fonts, "ascii") : null;
          if (col && col !== "#000000") css += "color:" + col + ";";
          if (sz && +sz !== Math.round(baseFs * 2)) css += "font-size:" + (+sz / 2) + "pt;";
          if (hl) css += "background-color:" + hl + ";";
          if (ff && ff !== baseFf && !/^(游ゴシック|游明朝|ＭＳ 明朝|ＭＳ ゴシック|Century|Calibri|Aptos|Times New Roman|Yu Gothic|Yu Mincho)/.test(ff)) css += "font-family:" + O.fontCss(ff).replace(/"/g, "'") + ";";
          return open + (css ? '<span style="' + css + '">' + html + "</span>" : html) + close;
        }
        function paraInner(p) {
          var html = "";
          Array.prototype.forEach.call(p.childNodes, function (c) {
            if (c.nodeType !== 1) return;
            if (c.localName === "r") html += runHtml(c);
            else if (c.localName === "commentRangeStart" && cmtById[attr(c, "id")]) { cmtById[attr(c, "id")].used = 1; openCmts.push(cmtById[attr(c, "id")].id); html += '<span class="cmt" data-cmt="' + cmtById[attr(c, "id")].id + '">'; }
            else if (c.localName === "commentRangeEnd" && cmtById[attr(c, "id")]) { var at = openCmts.indexOf(cmtById[attr(c, "id")].id); if (at >= 0) { openCmts.splice(at, 1); html += "</span>"; } }
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
        // 段落 1 つ分。前の段落から続いているコメントの範囲は、開きなおして、終わりで閉じる
        function paraBody(p) {
          var html = openCmts.map(function (id) { return '<span class="cmt" data-cmt="' + id + '">'; }).join("") + paraInner(p);
          return html + openCmts.map(function () { return "</span>"; }).join("");
        }
        var out = [], listStack = [];
        function closeLists(level) { while (listStack.length > level) out.push("</li></" + listStack.pop() + ">"); }
        function paraHtml(p) {
          var pr = kid(p, "pPr"), st = attr(kid(pr, "pStyle"), "val"), tag = blockTag(st), numPr = kid(pr, "numPr");
          var jc = attr(kid(pr, "jc"), "val"), inner = paraBody(p);
          // 段落の書式（余白・行間・揃え）。組み込みのスタイル（見出しなど）の書式を先に置き、段落の指定で上書きする
          var css = [];
          if (pstyleCss[st]) css.push(pstyleCss[st]);
          if (jc === "center") css.push("text-align:center");
          else if (jc === "right" || jc === "end") css.push("text-align:right");
          var spc = kid(pr, "spacing");
          if (spc) {
            var bf = attr(spc, "before"), af = attr(spc, "after"), ln = attr(spc, "line"), lr = attr(spc, "lineRule");
            if (bf) css.push("margin-top:" + (+bf / 20) + "pt");
            if (af) css.push("margin-bottom:" + (+af / 20) + "pt");
            if (ln && lr !== "exact" && lr !== "atLeast") css.push("line-height:" + lineCss(ln, baseFf));
          }
          if (inner.indexOf("data-anc=") >= 0) css.push("position:relative");
          var style = css.join(";");
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
          out.push(brk + "<" + tag + (cstMap[st] ? ' data-cst="' + esc(st) + '"' : "") + (style ? ' style="' + style + '"' : "") + ">" + (inner || "<br>") + "</" + tag + ">");
        }
        function tableHtml2(tbl) {
          var pr = kid(tbl, "tblPr"), grid = kid(tbl, "tblGrid");
          // 列の幅（w:gridCol は twips＝1/20 ポイント。CSS ピクセルにするには 15 で割る）
          // tblGridChange の中にも w:gridCol があるので、tblGrid の直接の子だけを見る
          var widths = grid ? Array.prototype.filter.call(grid.children, function (c) { return c.localName === "gridCol"; }).map(function (c) { return Math.round((+attr(c, "w") || 0) / 15); }) : [];
          var anyW = widths.some(function (w) { return w > 0; });
          var fixed = anyW || attr(kid(pr, "tblLayout"), "type") === "fixed";
          // 表全体の罫線。セルに指定の無い辺は、この罫線（外側・内側）を使う
          // 表のスタイル（tblStyle）。表に直接書いた罫線・余白を先に見て、無ければスタイル（basedOn もたどる）を見る
          var tchain = styleChain(attr(kid(pr, "tblStyle"), "val"));
          function tprop(name) {
            var e = kid(pr, name);
            for (var i = 0; i < tchain.length && !e; i++) e = kid(kid(tchain[i], "tblPr"), name);
            return e;
          }
          // 条件付きの書式（tblStylePr）。種類 → { fill, b, i, color }
          var cond = {};
          tchain.slice().reverse().forEach(function (s) {
            kids(s, "tblStylePr").forEach(function (sp) {
              var ty = attr(sp, "type"), rp = kid(sp, "rPr"), c = cond[ty] || (cond[ty] = {});
              var f = hex(attr(kid(kid(sp, "tcPr"), "shd"), "fill"));
              if (f) c.fill = f;
              if (kid(rp, "b")) c.b = onEl(kid(rp, "b"));
              if (kid(rp, "i")) c.i = onEl(kid(rp, "i"));
              var cc = hex(attr(kid(rp, "color"), "val"));
              if (cc) c.color = cc;
            });
          });
          // tblLook（どの条件を使うか）。無いときは Word の標準（見出し行・最初の列・縞）
          var look = kid(pr, "tblLook"), lookVal = look ? parseInt(attr(look, "val") || "0", 16) : 0x04A0;
          function lookOn(name, bit) { var v = attr(look, name); return v ? v === "1" || v === "true" : !!(lookVal & bit); }
          var useFirstRow = lookOn("firstRow", 0x20), useLastRow = lookOn("lastRow", 0x40), useFirstCol = lookOn("firstColumn", 0x80), useLastCol = lookOn("lastColumn", 0x100), useBand = !lookOn("noHBand", 0x200);
          var bandSize = +(attr(tprop("tblStyleRowBandSize"), "val") || 1);
          function styleOf(ri, ci, span, nrows) {
            var out = {};
            function put(c) { if (c) Object.keys(c).forEach(function (k) { out[k] = c[k]; }); }
            if (useBand) {
              var br = ri - (useFirstRow ? 1 : 0);
              if (br >= 0) put(cond[Math.floor(br / bandSize) % 2 === 0 ? "band1Horz" : "band2Horz"]);
            }
            if (useFirstCol && ci === 0) put(cond.firstCol);
            if (useLastCol && ci + span === lastCol) put(cond.lastCol);
            if (useFirstRow && ri === 0) put(cond.firstRow);
            if (useLastRow && ri === nrows - 1) put(cond.lastRow);
            return out;
          }
          var tb = tprop("tblBorders"), tbm = {};
          ["top", "left", "bottom", "right", "insideH", "insideV"].forEach(function (s) { tbm[s] = kid(tb, s); });
          var bdNone = !!tb && ["top", "left", "bottom", "right", "insideH", "insideV"].every(function (s) { return tbm[s] && attr(tbm[s], "val") === "none"; });
          // このアプリが書き出した「ふつうの表」の罫線（0.5pt・BFC3CC）は、画面の標準の罫線と同じなので、セルには書かない
          var tbPlain = !!tb && !bdNone && ["top", "left", "bottom", "right", "insideH", "insideV"].every(function (s) {
            return tbm[s] && (attr(tbm[s], "val") || "single") === "single" && (+attr(tbm[s], "sz") || 0) === 4 && /^BFC3CC$/i.test(attr(tbm[s], "color") || "");
          });
          var tcm = tprop("tblCellMar");
          // 行・セルを先に集める（罫線の種類を数えるため）
          var rows = kids(tbl, "tr").filter(function (tr) { return tr.parentNode === tbl; });
          var lastCol = 0;
          rows.forEach(function (tr) {
            var k = 0;
            Array.prototype.forEach.call(tr.children, function (tc) { if (tc.localName === "tc") k += +(attr(kid(kid(tc, "tcPr"), "gridSpan"), "val") || 1); });
            lastCol = Math.max(lastCol, k);
          });
          function bdText(e) {
            if (!e) return null;
            var val = attr(e, "val") || "single";
            if (val === "none" || val === "nil") return "none";
            var w = (+attr(e, "sz") || 4) / 8;
            var st = val === "dashed" ? "dashed" : val === "dotted" ? "dotted" : val === "double" ? "double" : "solid";
            return w + "pt " + st + " " + (hex(attr(e, "color")) || "#333333");
          }
          var info = rows.map(function (tr, ri) {
            var col = 0, cells = [];
            Array.prototype.forEach.call(tr.children, function (tc) {
              if (tc.localName !== "tc") return;
              var pr2 = kid(tc, "tcPr"), bd = kid(pr2, "tcBorders"), span = +(attr(kid(pr2, "gridSpan"), "val") || 1), css = [], lines = {};
              [["top", "Top"], ["left", "Left"], ["bottom", "Bottom"], ["right", "Right"]].forEach(function (o) {
                var s = o[0], e = bd ? kid(bd, s) : null;
                if (!e && !tbPlain) e = s === "top" ? (ri === 0 ? tbm.top : tbm.insideH) : s === "bottom" ? (ri === rows.length - 1 ? tbm.bottom : tbm.insideH) : s === "left" ? (col === 0 ? tbm.left : tbm.insideV) : (col + span === lastCol ? tbm.right : tbm.insideV);
                var txt = bdText(e);
                if (!txt) return;
                if (txt === "none") { css.push("border-" + s + ":none"); lines[s] = false; }
                else { css.push("border-" + s + ":" + txt); lines[s] = true; }
              });
              // セルの余白（Word の標準は 100 twips＝約 6.7px）
              var cm = kid(pr2, "tcMar") || tcm;
              if (cm) {
                var mv = function (k) { var e = kid(cm, k); return e ? Math.round((+attr(e, "w") || 0) / 15) : null; };
                var ct = mv("top"), cl = mv("left"), cb = mv("bottom"), cr = mv("right");
                if (ct != null || cl != null || cb != null || cr != null) {
                  if (ct == null) ct = cb == null ? 0 : cb;
                  if (cl == null) cl = cr == null ? 0 : cr;
                  if (cr == null) cr = cl;
                  if (cb == null) cb = ct;
                  css.push("padding:" + ct + "px " + cr + "px " + cb + "px " + cl + "px");
                }
              }
              var va = attr(kid(pr2, "vAlign"), "val");
              if (va === "center") css.push("vertical-align:middle");
              else if (va === "bottom") css.push("vertical-align:bottom");
              cells.push({ tc: tc, pr2: pr2, css: css, lines: lines, span: span });
              col += span;
            });
            return cells;
          });
          // 罫線の種類（外枠のみ・内線のみ・全部）を数える
          var anyLine = false, allFour = true, innerLine = false, outerLine = false;
          info.forEach(function (row, ri) {
            var col = 0;
            row.forEach(function (c) {
              var L = c.lines, topO = ri === 0, botO = ri === info.length - 1, leftO = col === 0, rightO = col + c.span === lastCol;
              var n = 0; ["top", "left", "bottom", "right"].forEach(function (s) { if (L[s]) n++; });
              if (n) anyLine = true;
              if (n < 4) allFour = false;
              if (L.top && !topO || L.bottom && !botO || L.left && !leftO || L.right && !rightO) innerLine = true;
              if (L.top && topO || L.bottom && botO || L.left && leftO || L.right && rightO) outerLine = true;
              col += c.span;
            });
          });
          var kind = allFour && anyLine ? "all" : innerLine ? "inner" : outerLine ? "outer" : "none";
          var tstyle = [];
          if (anyW) tstyle.push("width:" + widths.reduce(function (a, b) { return a + b; }, 0) + "px");
          if (fixed) tstyle.push("table-layout:fixed");
          var h = "<table" + ((anyLine || bdNone) ? ' data-bd="' + kind + '"' : "") + (tstyle.length ? ' style="' + tstyle.join(";") + '"' : "") + ">";
          if (anyW) h += "<colgroup>" + widths.map(function (w) { return '<col style="width:' + w + 'px">'; }).join("") + "</colgroup>";
          h += "<tbody>";
          info.forEach(function (row, ri) {
            // w:tblHeader は「ページをまたぐときに見出しを繰り返す」だけで、val="0" は見出しではない
            var head = onEl(kid(kid(rows[ri], "trPr"), "tblHeader"));
            h += "<tr>";
            var ci = 0;
            row.forEach(function (c) {
              var tc = c.tc, pr2 = c.pr2, vm = kid(pr2, "vMerge"), fill = hex(attr(kid(pr2, "shd"), "fill"));
              var cs = tchain.length ? styleOf(ri, ci, c.span, info.length) : {};
              ci += c.span;
              if (vm && attr(vm, "val") !== "restart") return;
              var cellTag = head || (ri === 0 && fill) ? "th" : "td", inner = [], st = c.css.slice();
              Array.prototype.forEach.call(tc.children, function (x) { if (x.localName === "p") { var y = paraBody(x); if (y) inner.push(y); } else if (x.localName === "tbl") inner.push(tableHtml2(x)); });
              if (fill && cellTag === "td" && fill !== "#ffffff") st.push("background-color:" + fill);
              else if (!fill && cs.fill && cs.fill !== "#ffffff") st.push("background-color:" + cs.fill);
              if (cs.b) st.push("font-weight:bold");
              if (cs.i) st.push("font-style:italic");
              if (cs.color) st.push("color:" + cs.color);
              h += "<" + cellTag + (c.span > 1 ? ' colspan="' + c.span + '"' : "") + (st.length ? ' style="' + st.join(";") + '"' : "") + ">" + (inner.join("<br>") || "<br>") + "</" + cellTag + ">";
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
        // 用紙（sectPr）を先に読む。図の位置（左右）を決めるのに、本文の幅が要る
        var page2 = {}, sect0 = kids(body, "sectPr").pop(), pg0 = kids(sect0, "pgSz")[0], mar0 = kids(sect0, "pgMar")[0], cols0 = kids(sect0, "cols")[0], td0 = kids(sect0, "textDirection")[0];
        if (pg0) {
          var w0 = +attr(pg0, "w"), h0 = +attr(pg0, "h"), land0 = attr(pg0, "orient") === "landscape" || w0 > h0, short0 = Math.min(w0, h0);
          page2.orient = land0 ? "landscape" : "portrait";
          page2.size = Math.abs(short0 - 12240) < 200 ? "letter" : Math.abs(short0 - 10318) < 300 ? "b5" : Math.abs(short0 - 16838) < 300 ? "a3" : "a4";
        }
        if (mar0) {
          var l0 = +attr(mar0, "left") || 0, t0 = +attr(mar0, "top") || 0;
          page2.mx = Math.round(l0 / 15);
          page2.my = Math.round(t0 / 15);
          page2.margin = l0 < 1000 ? "narrow" : l0 > 1900 ? "wide" : "normal";
        }
        if (cols0 && +attr(cols0, "num") > 1) page2.cols = Math.min(3, +attr(cols0, "num"));
        if (td0 && /tbRl/i.test(attr(td0, "val") || "")) page2.vert = true;
        page2.base = base;
        var pageMx = page2.mx != null ? page2.mx : MARGINS[page2.margin || "normal"].x;
        var contentW = (page2.orient === "landscape" ? SIZES[page2.size || "a4"].h : SIZES[page2.size || "a4"].w) - pageMx * 2;
        walkBody(body);
        closeLists(0);
        if (fnOrder.length) {
          out.push('<section class="footnotes"><hr><ol>' + fnOrder.map(function (id) {
            var openBefore = openCmts; openCmts = [];
            var li = '<li data-fn="fn' + stamp + id + '">' + kids(fnById[id], "p").map(paraInner).join("<br>").replace(/^\s+/, "") + "</li>";
            openCmts = openBefore;
            return li;
          }).join("") + "</ol></section>");
        }
        var data2 = { html: sanitize(out.join("").replace(/<p><\/p>/g, "")) || "<p><br></p>", page: page2, comments: [], styles: cstStyles };
        Object.keys(cmtById).forEach(function (id) { var c = cmtById[id]; data2.comments.push({ id: c.id, text: c.text, author: c.author, ts: c.ts }); });
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
      var mx = p.mx != null ? p.mx : m.x, my = p.my != null ? p.my : m.y;
      function mm(v) { return (Math.round(v * 25.4 / 96 * 10) / 10) + "mm"; }
      var boxes = (p.header ? "@top-center { content: " + cssText(p.header) + "; font-size: 9pt; color: #5a5f6b; }" : "") +
        (p.footer || p.num ? "@bottom-center { content: " + (p.footer ? cssText(p.footer + (p.num ? "　" : "")) + " " : "") + (p.num ? "counter(page)" : "") + "; font-size: 9pt; color: #5a5f6b; }" : "");
      return "@page { size: " + sz.css + " " + p.orient + "; margin: " + mm(my) + " " + mm(mx) + "; " + boxes + " }" +
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
      ribbonSig: function () { return cstList(data).length + ":" + cstCurrent(); },
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
