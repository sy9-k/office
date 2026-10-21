// SK Hub Office の Slide（プレゼンテーション）。window.OfficeApps.slide に登録し、/office/shared/office.js が呼ぶ（O は office.js が渡す道具）
//   ファイルの中身（data）: { slides: [スライド], theme: "mikan|ocean|matcha|night|mono|sakura", ratio: "16:9|4:3" }
//     スライド: { layout: "title|bullets|two|section|titleOnly|blank", title, sub, items, items2, notes, bg, trans: "none|fade|push|zoom", objs: [もの] }
//     もの（自由に置けるもの）: { id, type: "text|shape|image|chart", x, y, w, h（スライドの幅・高さに対する %）, rot（度）, html（文字）,
//       shape: "rect|round|ellipse|triangle|diamond|arrow|star|callout|line", fill, line, color, fs（文字の大きさ pt。幅 960px のときの大きさ）, b, i, u, align: "l|c|r", va: "t|m|b",
//       src（画像）, spec（グラフ）, anim: "fade|rise|zoom"（スライド ショーで順に出す） }
//   レイアウトの文字（title・sub・items）は、そのまま書きかえる。ものは、押して選び、ドラッグで動かす・角で大きさを変える・上の丸で回す
//   テーマの色は CSS の変数（office.css の .sl[data-theme]）。.pptx の読み書きは zip.js、グラフは chart.js
(function () {
  "use strict";
  window.OfficeApps = window.OfficeApps || {};
  window.OfficeApps.slide = function (O) {
    var t = O.t, esc = O.esc, icon = O.icon, $ = O.$, Zip = window.OfficeZip, Chart = window.OfficeChart;

    // ---------- テーマ（色とフォントのセット）。pptx に書き出すときも、この色を使う ----------
    var THEMES = {
      mikan: { label: t("みかん"), bg: "#ffffff", ink: "#261813", sub: "#6b4a3d", accent: "#ea580c", soft: "#fff1ec", blob: "#ffe2d4", sec: "#cc4900", secInk: "#fffbff" },
      ocean: { label: t("海"), bg: "#ffffff", ink: "#0f1d33", sub: "#3b4a63", accent: "#2563eb", soft: "#eaf1ff", blob: "#dbe7ff", sec: "#1d4ed8", secInk: "#ffffff" },
      matcha: { label: t("抹茶"), bg: "#fbfdf8", ink: "#16241a", sub: "#3e5443", accent: "#16a34a", soft: "#e9f6ea", blob: "#d6f0d9", sec: "#15803d", secInk: "#ffffff" },
      sakura: { label: t("さくら"), bg: "#fffafb", ink: "#2f1820", sub: "#6b4552", accent: "#db2777", soft: "#fdeef4", blob: "#fbd5e5", sec: "#be185d", secInk: "#ffffff" },
      mono: { label: t("モノクロ"), bg: "#ffffff", ink: "#111111", sub: "#555555", accent: "#111111", soft: "#f2f2f2", blob: "#e8e8e8", sec: "#111111", secInk: "#ffffff" },
      night: { label: t("夜"), bg: "#14161f", ink: "#f1f2f6", sub: "#b6bacb", accent: "#f59e0b", soft: "#232634", blob: "#2a2e3e", sec: "#f59e0b", secInk: "#14161f" }
    };
    var LAYOUTS = [
      { value: "title", label: t("タイトル スライド"), icon: "title" }, { value: "bullets", label: t("タイトルとコンテンツ"), icon: "format_list_bulleted" },
      { value: "two", label: t("2 つのコンテンツ"), icon: "view_column" }, { value: "section", label: t("セクション見出し"), icon: "splitscreen" },
      { value: "titleOnly", label: t("タイトルのみ"), icon: "short_text" }, { value: "blank", label: t("白紙"), icon: "crop_landscape" }
    ];
    var SHAPES = [
      { value: "rect", label: t("四角形"), icon: "rectangle" }, { value: "round", label: t("角丸四角形"), icon: "crop_square" }, { value: "ellipse", label: t("楕円"), icon: "circle" },
      { value: "triangle", label: t("三角形"), icon: "change_history" }, { value: "diamond", label: t("ひし形"), icon: "diamond" }, { value: "arrow", label: t("右矢印"), icon: "arrow_right_alt" },
      { value: "star", label: t("星"), icon: "star" }, { value: "callout", label: t("吹き出し"), icon: "chat_bubble" }, { value: "line", label: t("線"), icon: "horizontal_rule" }
    ];
    var TRANS = [{ value: "none", label: t("なし"), icon: "block" }, { value: "fade", label: t("フェード"), icon: "gradient" }, { value: "push", label: t("プッシュ"), icon: "arrow_back" }, { value: "zoom", label: t("ズーム"), icon: "zoom_in" }];
    var ANIMS = [{ value: "", label: t("なし"), icon: "block" }, { value: "fade", label: t("フェード"), icon: "gradient" }, { value: "rise", label: t("スライドイン"), icon: "north" }, { value: "zoom", label: t("ズーム"), icon: "zoom_in" }];

    function themeOf(d) { return THEMES[d.theme] ? d.theme : "mikan"; }
    function ratio(d) { return d.ratio === "4:3" ? 4 / 3 : 16 / 9; }
    function upgrade(d) {
      if (!Array.isArray(d.slides) || !d.slides.length) d.slides = [{ layout: "title", title: "", sub: "" }];
      d.slides.forEach(function (s) {
        if (!s.layout) s.layout = "title";
        if (!Array.isArray(s.objs)) s.objs = [];
        // 0.1 の画像（右下の写真）は、ものにする
        if (s.image) { s.objs.push({ id: uid(), type: "image", src: s.image, x: 62, y: 52, w: 30, h: 38 }); delete s.image; }
      });
      if (!THEMES[d.theme]) d.theme = "mikan";
      if (d.ratio !== "4:3") d.ratio = "16:9";
      return d;
    }
    var seq = 0;
    function uid() { return "o" + Date.now().toString(36) + (seq++).toString(36); }

    // ---------- 文字（もの）の HTML は、太字などだけ残す ----------
    function cleanText(html) {
      var doc = new DOMParser().parseFromString("<body>" + html + "</body>", "text/html");
      (function walk(n) {
        Array.prototype.slice.call(n.childNodes).forEach(function (c) {
          if (c.nodeType === 8) { c.remove(); return; }
          if (c.nodeType !== 1) return;
          if (!/^(B|STRONG|I|EM|U|S|BR|DIV|P|SPAN|FONT)$/.test(c.tagName)) { walk(c); c.replaceWith.apply(c, Array.prototype.slice.call(c.childNodes)); return; }
          Array.prototype.slice.call(c.attributes).forEach(function (a) {
            if (a.name === "style") { var keep = a.value.split(";").filter(function (x) { return /^\s*(color|font-weight|font-style|text-decoration|font-size)\s*:/.test(x) && !/url\(/i.test(x); }); if (keep.length) c.setAttribute("style", keep.join(";")); else c.removeAttribute("style"); }
            else if (!(c.tagName === "FONT" && a.name === "color")) c.removeAttribute(a.name);
          });
          walk(c);
        });
      })(doc.body);
      return doc.body.innerHTML;
    }
    function plain(html) { var d = document.createElement("div"); d.innerHTML = String(html || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/(div|p)>/gi, "\n"); return d.textContent.replace(/\n+$/, ""); }

    // ---------- スライドの HTML ----------
    function shapeSvg(o) {
      var f = o.fill === "none" ? "none" : (o.fill || "var(--sl-accent)"), l = o.line === "none" || !o.line ? "none" : o.line, sw = o.line && o.line !== "none" ? 2 : 0, d;
      switch (o.shape) {
        case "ellipse": d = '<ellipse cx="50" cy="50" rx="50" ry="50"'; break;
        case "round": d = '<rect x="0" y="0" width="100" height="100" rx="14" ry="14"'; break;
        case "triangle": d = '<path d="M50 0 L100 100 L0 100 Z"'; break;
        case "diamond": d = '<path d="M50 0 L100 50 L50 100 L0 50 Z"'; break;
        case "arrow": d = '<path d="M0 30 L62 30 L62 5 L100 50 L62 95 L62 70 L0 70 Z"'; break;
        case "star": d = '<path d="M50 2 L61.8 36.5 L98 36.8 L69 58.6 L79.4 93.3 L50 72.6 L20.6 93.3 L31 58.6 L2 36.8 L38.2 36.5 Z"'; break;
        case "callout": d = '<path d="M8 0 H92 Q100 0 100 8 V64 Q100 72 92 72 H40 L22 100 L26 72 H8 Q0 72 0 64 V8 Q0 0 8 0 Z"'; break;
        case "line": return '<svg class="sl-shape" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><line x1="0" y1="50" x2="100" y2="50" stroke="' + esc(o.line && o.line !== "none" ? o.line : "var(--sl-ink)") + '" stroke-width="3" vector-effect="non-scaling-stroke"/></svg>';
        default: d = '<rect x="0" y="0" width="100" height="100"';
      }
      return '<svg class="sl-shape" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">' + d + ' fill="' + esc(f) + '" stroke="' + esc(l) + '" stroke-width="' + sw + '" vector-effect="non-scaling-stroke"/></svg>';
    }
    function fsCss(pt) { return "font-size:" + Math.round(pt * 4 / 3 / 9.6 * 1000) / 1000 + "cqw;"; }
    function objHtml(o, mode, r) {
      var st = "left:" + o.x + "%;top:" + o.y + "%;width:" + o.w + "%;height:" + o.h + "%;" + (o.rot ? "transform:rotate(" + o.rot + "deg);" : "");
      var inner = "";
      if (o.type === "image") inner = '<img src="' + esc(o.src || "") + '" alt="' + esc(o.alt || "") + '" draggable="false">';
      else if (o.type === "chart") inner = Chart.svg(o.spec || {}, { w: Math.max(80, o.w * 9.6), h: Math.max(60, o.h * 9.6 / r), transparent: true });
      else {
        if (o.type === "shape") inner = shapeSvg(o);
        var ts = fsCss(o.fs || (o.type === "shape" ? 18 : 20)) + (o.color ? "color:" + o.color + ";" : o.type === "shape" && o.fill !== "none" ? "color:#fff;" : "") + (o.b ? "font-weight:700;" : "") + (o.i ? "font-style:italic;" : "") + (o.u ? "text-decoration:underline;" : "") +
          "text-align:" + ({ l: "left", c: "center", r: "right" }[o.align || (o.type === "shape" ? "c" : "l")]) + ";";
        // 外側（.sl-text）で上下の位置を決め、内側（.sl-tx）に文字を書く
        if (o.shape !== "line") inner += '<div class="sl-text" style="justify-content:' + ({ t: "flex-start", m: "center", b: "flex-end" }[o.va || (o.type === "shape" ? "m" : "t")]) + '"><div class="sl-tx" style="' + ts + '">' + (o.html || "") + "</div></div>";
      }
      return '<div class="sl-obj sl-obj--' + o.type + (o.anim && mode === "show" ? " sl-anim sl-anim--" + o.anim : "") + '" data-obj="' + esc(o.id) + '" style="' + st + '">' + inner + "</div>";
    }
    function listHtml(items, key, ce) {
      var list = items && items.length ? items : [""];
      return '<ul class="sl-list">' + list.map(function (it, i) { return "<li" + (ce ? ' data-k="' + key + "." + i + '"' + ce : "") + ">" + esc(it) + "</li>"; }).join("") + "</ul>";
    }
    function slideHtml(s, mode, d) {
      d = d || data;
      var ce = mode === "edit" ? ' contenteditable="true" spellcheck="false"' : "", lay = s.layout || "title", r = ratio(d);
      var h = '<div class="sl sl--' + lay + '" data-theme="' + themeOf(d) + '"' + (s.bg ? ' style="--sl-bg:' + esc(s.bg) + '"' : "") + ">";
      if (lay !== "blank") h += '<i class="sl-blob" aria-hidden="true"></i>';
      if (lay !== "blank") {
        h += '<div class="sl-in">';
        if (mode === "edit" || s.title) h += '<h2 class="sl-title" data-k="title"' + ce + ">" + esc(s.title || "") + "</h2>";
        if (lay === "bullets") h += listHtml(s.items, "items", ce);
        else if (lay === "two") h += '<div class="sl-cols">' + listHtml(s.items, "items", ce) + listHtml(s.items2, "items2", ce) + "</div>";
        else if (lay === "title" || lay === "section") { if (mode === "edit" || s.sub) h += '<p class="sl-sub" data-k="sub"' + ce + ">" + esc(s.sub || "") + "</p>"; }
        h += "</div>";
      }
      h += (s.objs || []).map(function (o) { return objHtml(o, mode, r); }).join("");
      return h + "</div>";
    }

    // ---------- いまの状態 ----------
    var data = null, cur = 0, sel = [], editingObj = null, readOnly = false, showNotes = true, zoomZ = 1, clipObjs = null, painter = null;
    function el(id) { return $(id); }
    function slide() { return data.slides[cur]; }
    function objById(id) { return (slide().objs || []).filter(function (o) { return o.id === id; })[0]; }
    function selObjs() { return sel.map(objById).filter(Boolean); }
    function changed(typing) { O.changed(typing); }

    // ---------- 画面を作る ----------
    function render(opts) {
      data = O.data();
      upgrade(data);
      if (!(opts && opts.restore)) { cur = 0; sel = []; }
      cur = Math.min(cur, data.slides.length - 1);
      sel = sel.filter(function (id) { return !!objById(id); });
      editingObj = null;
      $("canvas").innerHTML = '<div class="deck"><div class="deck-thumbs" id="thumbs" role="listbox" aria-label="' + esc(t("スライド")) + '" tabindex="0"></div>' +
        '<div class="deck-stage" id="deck-stage"><div class="deck-slide" id="stage"></div>' +
        '<div class="deck-notes" id="notes" contenteditable="true" spellcheck="false" aria-label="' + esc(t("ノート")) + '"' + (showNotes ? "" : " hidden") + "></div></div></div>";
      el("notes").setAttribute("data-ph", t("クリックしてノートを追加"));
      wire();
      drawThumbs();
      drawStage();
      applyZoom();
    }
    function applyRatio() { var r = data.ratio === "4:3" ? "4 / 3" : "16 / 9"; $("canvas").style.setProperty("--sl-ratio", r); }
    function applyZoom() {
      var s = el("stage");
      if (!s) return;
      s.style.width = zoomZ === 1 ? "" : Math.round(1040 * zoomZ) + "px";
      s.style.maxWidth = zoomZ === 1 ? "" : "none";
      placeFrame();
    }
    function drawThumbs(keepScroll) {
      var box = el("thumbs");
      if (!box) return;
      applyRatio();
      box.innerHTML = data.slides.map(function (s, i) {
        return '<button type="button" class="deck-thumb" role="option" aria-selected="' + (i === cur) + '" data-slide="' + i + '" draggable="false"><span class="deck-no">' + (i + 1) + (s.trans && s.trans !== "none" ? '<span class="msr deck-no__fx" aria-hidden="true">auto_awesome</span>' : "") + '</span><span class="th th--slide">' + slideHtml(s, "thumb") + "</span></button>";
      }).join("");
      stat();
      if (!keepScroll) { var on = box.querySelector('[aria-selected="true"]'); if (on) on.scrollIntoView({ block: "nearest", inline: "nearest" }); }
    }
    function refreshThumb(i) {
      var b = el("thumbs") && el("thumbs").querySelector('[data-slide="' + i + '"] .th');
      if (b) b.innerHTML = slideHtml(data.slides[i], "thumb");
    }
    function drawStage(focusItem) {
      var st = el("stage");
      if (!st) return;
      st.innerHTML = slideHtml(slide(), readOnly ? "thumb" : "edit") + '<div class="sel-frame" id="frame" hidden>' +
        ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map(function (h) { return '<span class="sel-frame__h sel-frame__h--' + h + '" data-h="' + h + '"></span>'; }).join("") +
        '<span class="sel-frame__rot" data-h="rot"></span></div><div class="sl-guides" id="guides"></div>';
      el("notes").textContent = slide().notes || "";
      if (focusItem !== undefined) {
        var li = st.querySelector('[data-k="' + focusItem + '"]');
        if (li) { li.focus(); var r = document.createRange(); r.selectNodeContents(li); r.collapse(false); var s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      }
      placeFrame();
      stat();
    }
    function stat() {
      if (!data) return;
      O.stat("<span>" + esc(t("スライド {i}/{n}", { i: cur + 1, n: data.slides.length })) + "</span>" + (sel.length ? "<span>" + esc(t("{n} 個を選んでいます", { n: sel.length })) + "</span>" : "") +
        (readOnly ? "<span>" + esc(t("閲覧モード")) + "</span>" : "") + '<span class="hide-sm">' + esc(THEMES[themeOf(data)].label) + "</span>");
    }
    function go(i) {
      if (i < 0 || i >= data.slides.length) return;
      endObjEdit();
      cur = i; sel = [];
      drawThumbs(); drawStage();
      O.sync();
    }

    // ---------- 選んだものの枠 ----------
    function bounds(list) {
      var x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      list.forEach(function (o) { x1 = Math.min(x1, o.x); y1 = Math.min(y1, o.y); x2 = Math.max(x2, o.x + o.w); y2 = Math.max(y2, o.y + o.h); });
      return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    }
    function placeFrame() {
      var f = el("frame");
      if (!f) return;
      var list = selObjs();
      if (!list.length || readOnly) { f.hidden = true; return; }
      var b = list.length === 1 ? list[0] : bounds(list);
      f.hidden = false;
      f.style.left = b.x + "%"; f.style.top = b.y + "%"; f.style.width = b.w + "%"; f.style.height = b.h + "%";
      f.style.transform = list.length === 1 && list[0].rot ? "rotate(" + list[0].rot + "deg)" : "";
      f.classList.toggle("is-multi", list.length > 1);
      f.classList.toggle("is-line", list.length === 1 && list[0].shape === "line");
      el("stage").querySelectorAll(".sl-obj").forEach(function (x) { x.classList.toggle("is-sel", sel.indexOf(x.getAttribute("data-obj")) >= 0); });
    }
    function select(ids, add) {
      endObjEdit();
      if (add) ids.forEach(function (id) { var i = sel.indexOf(id); if (i >= 0) sel.splice(i, 1); else sel.push(id); });
      else sel = ids.slice();
      placeFrame();
      stat();
      O.sync();
    }
    function redrawObj(o) {
      var node = el("stage").querySelector('[data-obj="' + o.id + '"]');
      if (!node) return;
      var tmp = document.createElement("div");
      tmp.innerHTML = objHtml(o, "edit", ratio(data));
      node.replaceWith(tmp.firstChild);
      placeFrame();
    }
    function redrawSel() { selObjs().forEach(redrawObj); refreshThumb(cur); }

    // ---------- ものを足す ----------
    function addObj(o) {
      if (readOnly) return;
      o.id = uid();
      slide().objs.push(o);
      drawStage();
      select([o.id]);
      refreshThumb(cur);
      changed();
      return o;
    }
    function addText() { var o = addObj({ type: "text", x: 30, y: 40, w: 40, h: 14, html: t("テキストを入力"), fs: 24 }); if (o) startObjEdit(o, true); }
    function addShape(kind) {
      var line = kind === "line";
      addObj({ type: "shape", shape: kind, x: 38, y: line ? 49 : 34, w: 24, h: line ? 2 : 32, html: "", fs: 18, line: line ? THEMES[themeOf(data)].ink : undefined });
    }
    function addImage(src) {
      if (!src) return;
      var img = new Image();
      img.onload = function () {
        var r = ratio(data), aw = img.naturalWidth || 400, ah = img.naturalHeight || 300, w = 46, h = w * r * ah / aw;
        if (h > 70) { h = 70; w = h * aw / ah / r; }
        addObj({ type: "image", src: src, x: (100 - w) / 2, y: (100 - h) / 2, w: w, h: h });
      };
      img.src = src;
    }
    function chartDialog(spec) {
      var editing = !!spec;
      spec = spec || { type: "col", title: "", cats: [t("1月"), t("2月"), t("3月"), t("4月")], series: [{ name: t("系列 1"), values: [42, 58, 51, 72] }] };
      var rows = [[""].concat(spec.series.map(function (s) { return s.name; }))];
      (spec.cats || []).forEach(function (c, i) { rows.push([c].concat(spec.series.map(function (s) { return s.values[i] === null || s.values[i] === undefined ? "" : s.values[i]; }))); });
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
    function addChart() { chartDialog().then(function (spec) { if (spec) addObj({ type: "chart", spec: spec, x: 18, y: 20, w: 64, h: 66 }); }); }

    // ---------- ものの文字を書く ----------
    function startObjEdit(o, selectAll) {
      if (readOnly || !o || o.type === "image" || o.type === "chart" || o.shape === "line") return;
      var node = el("stage").querySelector('[data-obj="' + o.id + '"] .sl-tx');
      if (!node) return;
      editingObj = o.id;
      node.contentEditable = "true";
      node.focus();
      var r = document.createRange(); r.selectNodeContents(node); if (!selectAll) r.collapse(false);
      var s = getSelection(); s.removeAllRanges(); s.addRange(r);
      el("frame").classList.add("is-edit");
      O.sync();
    }
    function endObjEdit() {
      if (!editingObj) return;
      var o = objById(editingObj), node = el("stage") && el("stage").querySelector('[data-obj="' + editingObj + '"] .sl-tx');
      editingObj = null;
      if (el("frame")) el("frame").classList.remove("is-edit");
      if (node) { node.contentEditable = "false"; if (o) { var h = cleanText(node.innerHTML); if (h !== o.html) { o.html = h; refreshThumb(cur); changed(); } } }
    }

    // ---------- ドラッグ（動かす・大きさ・回す）。近くのものや真ん中に合わせる ----------
    function stageRect() { var s = el("stage").querySelector(".sl"); return s.getBoundingClientRect(); }
    function snapLines(skip) {
      var xs = [0, 50, 100], ys = [0, 50, 100];
      slide().objs.forEach(function (o) { if (skip.indexOf(o.id) >= 0) return; xs.push(o.x, o.x + o.w / 2, o.x + o.w); ys.push(o.y, o.y + o.h / 2, o.y + o.h); });
      return { xs: xs, ys: ys };
    }
    function showGuides(gx, gy) {
      var g = el("guides");
      if (!g) return;
      g.innerHTML = gx.map(function (x) { return '<i class="sl-guide sl-guide--v" style="left:' + x + '%"></i>'; }).join("") + gy.map(function (y) { return '<i class="sl-guide sl-guide--h" style="top:' + y + '%"></i>'; }).join("");
    }
    function dragObjs(e, ids) {
      var R = stageRect(), start = ids.map(function (id) { var o = objById(id); return { o: o, x: o.x, y: o.y }; }), x0 = e.clientX, y0 = e.clientY, moved = false;
      var lines = snapLines(ids), b0 = bounds(start.map(function (s) { return s.o; }));
      function mv(ev) {
        var dx = (ev.clientX - x0) / R.width * 100, dy = (ev.clientY - y0) / R.height * 100;
        if (!moved && Math.abs(ev.clientX - x0) + Math.abs(ev.clientY - y0) < 3) return;
        moved = true;
        // 端・真ん中を、近くの線に合わせる（Alt を押しているあいだは合わせない）
        var gx = [], gy = [];
        if (!ev.altKey) {
          var bx = [b0.x + dx, b0.x + dx + b0.w / 2, b0.x + dx + b0.w], by = [b0.y + dy, b0.y + dy + b0.h / 2, b0.y + dy + b0.h], best;
          best = null; bx.forEach(function (v) { lines.xs.forEach(function (l) { var d = l - v; if (Math.abs(d) < 0.8 && (!best || Math.abs(d) < Math.abs(best.d))) best = { d: d, l: l }; }); });
          if (best) { dx += best.d; gx.push(best.l); }
          best = null; by.forEach(function (v) { lines.ys.forEach(function (l) { var d = l - v; if (Math.abs(d) < 1.2 && (!best || Math.abs(d) < Math.abs(best.d))) best = { d: d, l: l }; }); });
          if (best) { dy += best.d; gy.push(best.l); }
        }
        if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
        start.forEach(function (s) {
          s.o.x = Math.round((s.x + dx) * 100) / 100; s.o.y = Math.round((s.y + dy) * 100) / 100;
          var node = el("stage").querySelector('[data-obj="' + s.o.id + '"]');
          if (node) { node.style.left = s.o.x + "%"; node.style.top = s.o.y + "%"; }
        });
        showGuides(gx, gy);
        placeFrame();
      }
      function up() {
        window.removeEventListener("pointermove", mv); window.removeEventListener("pointerup", up);
        showGuides([], []);
        if (moved) { refreshThumb(cur); changed(); }
      }
      window.addEventListener("pointermove", mv); window.addEventListener("pointerup", up);
    }
    function dragHandle(e, h) {
      var o = selObjs()[0];
      if (!o) return;
      var R = stageRect(), a = { x: o.x, y: o.y, w: o.w, h: o.h, rot: o.rot || 0 }, x0 = e.clientX, y0 = e.clientY, keepRatio = o.type === "image";
      var cx = R.left + (a.x + a.w / 2) / 100 * R.width, cy = R.top + (a.y + a.h / 2) / 100 * R.height;
      function mv(ev) {
        if (h === "rot") {
          var ang = Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180 / Math.PI + 90;
          if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
          ang = ((Math.round(ang) % 360) + 360) % 360;
          o.rot = ang === 0 ? undefined : ang;
        } else {
          // 回したものは、回した向きのまま大きさを変える
          var rad = (a.rot || 0) * Math.PI / 180, ddx = ev.clientX - x0, ddy = ev.clientY - y0;
          var lx = (ddx * Math.cos(rad) + ddy * Math.sin(rad)) / R.width * 100, ly = (-ddx * Math.sin(rad) + ddy * Math.cos(rad)) / R.height * 100;
          var x = a.x, y = a.y, w = a.w, hh = a.h;
          if (/e/.test(h)) w = Math.max(1, a.w + lx);
          if (/s/.test(h)) hh = Math.max(o.shape === "line" ? 0.5 : 1, a.h + ly);
          if (/w/.test(h)) { w = Math.max(1, a.w - lx); x = a.x + (a.w - w); }
          if (/n/.test(h)) { hh = Math.max(1, a.h - ly); y = a.y + (a.h - hh); }
          if ((keepRatio || ev.shiftKey) && h.length === 2) {
            var k = Math.max(w / a.w, hh / a.h);
            w = a.w * k; hh = a.h * k;
            if (/w/.test(h)) x = a.x + a.w - w;
            if (/n/.test(h)) y = a.y + a.h - hh;
          }
          o.x = Math.round(x * 100) / 100; o.y = Math.round(y * 100) / 100; o.w = Math.round(w * 100) / 100; o.h = Math.round(hh * 100) / 100;
        }
        redrawObj(o);
      }
      function up() { window.removeEventListener("pointermove", mv); window.removeEventListener("pointerup", up); refreshThumb(cur); changed(); }
      window.addEventListener("pointermove", mv); window.addEventListener("pointerup", up);
    }

    // ---------- クリック・入力・キー ----------
    function wire() {
      var thumbs = el("thumbs"), stage = el("stage"), notes = el("notes");
      thumbs.addEventListener("click", function (e) { var b = e.target.closest("[data-slide]"); if (b) go(+b.getAttribute("data-slide")); });
      thumbs.addEventListener("contextmenu", function (e) {
        var b = e.target.closest("[data-slide]");
        if (!b) return;
        e.preventDefault();
        go(+b.getAttribute("data-slide"));
        slideMenu({ x: e.clientX, y: e.clientY });
      });
      // スライドの順番をドラッグで変える
      thumbs.addEventListener("pointerdown", function (e) {
        var b = e.target.closest("[data-slide]");
        if (!b || e.button !== 0 || readOnly) return;
        var from = +b.getAttribute("data-slide"), y0 = e.clientY, x0 = e.clientX, dragging = false, to = from;
        function mv(ev) {
          if (!dragging && Math.abs(ev.clientY - y0) + Math.abs(ev.clientX - x0) < 8) return;
          dragging = true;
          var items = Array.prototype.slice.call(thumbs.querySelectorAll("[data-slide]"));
          to = items.length - 1;
          for (var i = 0; i < items.length; i++) { var r = items[i].getBoundingClientRect(); if (ev.clientY < r.top + r.height / 2 && (thumbs.scrollWidth <= thumbs.clientWidth || ev.clientX < r.right)) { to = i; break; } }
          items.forEach(function (x, i) { x.classList.toggle("is-drop", i === to && to !== from); });
        }
        function up() {
          window.removeEventListener("pointermove", mv); window.removeEventListener("pointerup", up);
          thumbs.querySelectorAll(".is-drop").forEach(function (x) { x.classList.remove("is-drop"); });
          if (!dragging || to === from) return;
          var s = data.slides.splice(from, 1)[0];
          if (to > from) to--;
          data.slides.splice(to, 0, s);
          cur = to;
          drawThumbs(); drawStage(); changed();
        }
        window.addEventListener("pointermove", mv); window.addEventListener("pointerup", up);
      });
      thumbs.addEventListener("keydown", function (e) {
        if (e.key === "ArrowDown" || e.key === "ArrowRight") { e.preventDefault(); go(cur + 1); }
        else if (e.key === "ArrowUp" || e.key === "ArrowLeft") { e.preventDefault(); go(cur - 1); }
        else if ((e.key === "Delete" || e.key === "Backspace") && !readOnly) { e.preventDefault(); command("del"); }
        else if (e.key === "Enter" && !readOnly) { e.preventDefault(); command("add"); }
      });
      stage.addEventListener("pointerdown", function (e) {
        if (e.button !== 0) return;
        var h = e.target.closest("[data-h]");
        if (h && !readOnly) { e.preventDefault(); dragHandle(e, h.getAttribute("data-h")); return; }
        var node = e.target.closest(".sl-obj");
        if (node && !readOnly) {
          var id = node.getAttribute("data-obj");
          if (editingObj === id) return;
          e.preventDefault();
          if (painter) { applyPainter(objById(id)); return; }
          if (e.shiftKey || e.ctrlKey) { select([id], true); return; }
          if (sel.indexOf(id) < 0) select([id]);
          else endObjEdit();
          stage.focus({ preventScroll: true });
          dragObjs(e, sel.slice());
          return;
        }
        if (e.target.closest("[data-k]")) { if (sel.length) select([]); return; }
        if (e.target.closest("#frame") && sel.length && !readOnly) { e.preventDefault(); dragObjs(e, sel.slice()); return; }
        if (sel.length || editingObj) select([]);
      });
      stage.addEventListener("dblclick", function (e) {
        var node = e.target.closest(".sl-obj") || (e.target.closest("#frame") && sel.length === 1 ? el("stage").querySelector('[data-obj="' + sel[0] + '"]') : null);
        if (!node || readOnly) return;
        var o = objById(node.getAttribute("data-obj"));
        if (o && o.type === "chart") { chartDialog(o.spec).then(function (s) { if (s) { o.spec = s; redrawObj(o); refreshThumb(cur); changed(); } }); return; }
        startObjEdit(o);
      });
      stage.setAttribute("tabindex", "-1");
      stage.addEventListener("input", function (e) {
        if (e.target.closest(".sl-tx")) { refreshThumbSoon(); changed(true); return; }
        var k = e.target.getAttribute && e.target.getAttribute("data-k");
        if (!k) return;
        var s = slide(), m = /^(items2?)\.(\d+)$/.exec(k);
        if (m) { s[m[1]] = s[m[1]] || []; s[m[1]][+m[2]] = e.target.textContent; }
        else s[k] = e.target.textContent;
        refreshThumbSoon();
        changed(true);
      });
      stage.addEventListener("focusout", function (e) { if (editingObj && e.target.closest(".sl-tx") && !stage.contains(e.relatedTarget)) endObjEdit(); });
      // 箇条書き: Enter で次の行・空の行で Backspace で消す。見出しの Enter は使わない
      stage.addEventListener("keydown", function (e) {
        if (e.isComposing) return;
        var k = e.target.getAttribute && e.target.getAttribute("data-k");
        if (k) {
          var s = slide(), m = /^(items2?)\.(\d+)$/.exec(k);
          if (m) {
            var list = s[m[1]] = s[m[1]] || [""], i = +m[2];
            if (e.key === "Enter") { e.preventDefault(); list.splice(i + 1, 0, ""); drawStage(m[1] + "." + (i + 1)); refreshThumb(cur); changed(); }
            else if (e.key === "Backspace" && !e.target.textContent && list.length > 1) { e.preventDefault(); list.splice(i, 1); drawStage(m[1] + "." + Math.max(0, i - 1)); refreshThumb(cur); changed(); }
          } else if (e.key === "Enter") e.preventDefault();
          return;
        }
        if (e.target.closest(".sl-tx")) { if (e.key === "Escape") { e.preventDefault(); var id = editingObj; endObjEdit(); select([id]); } return; }
      });
      notes.addEventListener("input", function () { slide().notes = notes.innerText.replace(/\n$/, ""); changed(true); });
    }
    var thumbTimer = null;
    function refreshThumbSoon() { clearTimeout(thumbTimer); thumbTimer = setTimeout(function () { refreshThumb(cur); }, 250); }
    function onKey(e, field) {
      if (field || !data || editingObj || document.querySelector(".present")) return;
      if (e.target && (e.target.closest && (e.target.closest("[data-k]") || e.target.closest(".deck-notes") || e.target.closest(".sl-tx") || e.target.closest(".deck-thumbs")))) return;
      var mod = e.ctrlKey || e.metaKey, k = e.key;
      if (k === "F5") { e.preventDefault(); present(e.shiftKey ? cur : 0); return; }
      if (!sel.length) {
        if (k === "PageDown") { e.preventDefault(); go(cur + 1); }
        if (k === "PageUp") { e.preventDefault(); go(cur - 1); }
        if (mod && k.toLowerCase() === "v" && clipObjs) { e.preventDefault(); pasteObjs(); }
        if (mod && k.toLowerCase() === "m" && !readOnly) { e.preventDefault(); command("add"); }
        return;
      }
      if (readOnly) return;
      if (k === "Delete" || k === "Backspace") { e.preventDefault(); command("objDel"); return; }
      if (/^Arrow/.test(k)) {
        e.preventDefault();
        var step = e.shiftKey ? 2 : 0.25, d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
        selObjs().forEach(function (o) { o.x = Math.round((o.x + d[0] * step) * 100) / 100; o.y = Math.round((o.y + d[1] * step * ratio(data)) * 100) / 100; redrawObj(o); });
        refreshThumbSoon(); changed(true);
        return;
      }
      if (k === "Enter" || k === "F2") { e.preventDefault(); startObjEdit(selObjs()[0], true); return; }
      if (k === "Escape") { select([]); return; }
      if (mod && k.toLowerCase() === "d") { e.preventDefault(); command("objDup"); return; }
      if (mod && k.toLowerCase() === "c") { e.preventDefault(); clipObjs = O.copy(selObjs()); O.notify(t("コピーしました")); return; }
      if (mod && k.toLowerCase() === "x") { e.preventDefault(); clipObjs = O.copy(selObjs()); command("objDel"); return; }
      if (mod && k.toLowerCase() === "v" && clipObjs) { e.preventDefault(); pasteObjs(); return; }
      if (mod && k.toLowerCase() === "a") { e.preventDefault(); select(slide().objs.map(function (o) { return o.id; })); return; }
      if (mod && !e.shiftKey && /^[biu]$/i.test(k)) { e.preventDefault(); command({ b: "bold", i: "italic", u: "underline" }[k.toLowerCase()]); }
    }
    function pasteObjs() {
      if (!clipObjs || readOnly) return;
      var ids = [];
      clipObjs.forEach(function (o) { var c = O.copy(o); c.id = uid(); c.x += 2; c.y += 2; slide().objs.push(c); ids.push(c.id); });
      clipObjs = clipObjs.map(function (o) { var c = O.copy(o); c.x += 2; c.y += 2; return c; });
      drawStage(); select(ids); refreshThumb(cur); changed();
    }
    function applyPainter(target) {
      var p = painter;
      painter = null;
      if (!target || !p) { O.sync(); return; }
      ["fill", "line", "color", "fs", "b", "i", "u", "align", "va"].forEach(function (k) { if (p[k] === undefined) delete target[k]; else target[k] = p[k]; });
      redrawObj(target); refreshThumb(cur); changed(); O.sync();
    }

    // ---------- スライドのメニュー ----------
    function slideMenu(at) {
      O.menu(at, [
        { value: "add", label: t("新しいスライド"), icon: "add_box", sub: "Ctrl+M" }, { value: "dup", label: t("スライドの複製"), icon: "content_copy" },
        { value: "layout", label: t("レイアウト"), icon: "dashboard" }, { sep: 1 },
        { value: "up", label: t("上へ移動"), icon: "arrow_upward", disabled: cur === 0 }, { value: "down", label: t("下へ移動"), icon: "arrow_downward", disabled: cur === data.slides.length - 1 },
        { sep: 1 }, { value: "del", label: t("スライドの削除"), icon: "delete" }
      ]).then(function (v) { if (v) command(v, at); });
    }

    // ---------- リボンのボタン ----------
    function setObj(patch) {
      var list = selObjs();
      if (!list.length) { O.notify(t("図形やテキスト ボックスを選んでください")); return false; }
      list.forEach(function (o) { Object.keys(patch).forEach(function (k) { if (patch[k] === null) delete o[k]; else o[k] = patch[k]; }); });
      redrawSel(); changed();
      return true;
    }
    function inlineOrObj(cmd, key) {
      // ものの文字を書いているときは選んだ文字だけ、そうでなければ、ものの文字すべて
      var s = getSelection(), a = s.anchorNode && (s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentNode);
      if (a && a.closest && (a.closest(".sl-tx[contenteditable='true']") || a.closest("[data-k]"))) {
        if (a.closest("[data-k]")) { O.notify(t("レイアウトの文字は、まとめて書式を変えられません。テキスト ボックスを使ってください")); return; }
        document.execCommand(cmd);
        changed(true);
        return;
      }
      var o = selObjs()[0];
      if (o) setObj((function () { var p = {}; p[key] = o[key] ? null : 1; return p; })());
      else O.notify(t("図形やテキスト ボックスを選んでください"));
    }
    function align(where) {
      var list = selObjs();
      if (!list.length) return;
      var b = list.length > 1 ? bounds(list) : { x: 0, y: 0, w: 100, h: 100 };
      list.forEach(function (o) {
        if (where === "l") o.x = b.x; if (where === "r") o.x = b.x + b.w - o.w; if (where === "c") o.x = b.x + (b.w - o.w) / 2;
        if (where === "t") o.y = b.y; if (where === "b") o.y = b.y + b.h - o.h; if (where === "m") o.y = b.y + (b.h - o.h) / 2;
        o.x = Math.round(o.x * 100) / 100; o.y = Math.round(o.y * 100) / 100;
      });
      redrawSel(); changed();
    }
    function order(kind) {
      var objs = slide().objs, ids = sel.slice();
      if (!ids.length) return;
      var picked = objs.filter(function (o) { return ids.indexOf(o.id) >= 0; }), rest = objs.filter(function (o) { return ids.indexOf(o.id) < 0; });
      if (kind === "front") slide().objs = rest.concat(picked);
      else if (kind === "back") slide().objs = picked.concat(rest);
      else {
        var i = objs.indexOf(picked[0]), j = kind === "forward" ? Math.min(objs.length - 1, i + 1) : Math.max(0, i - 1);
        objs.splice(i, 1); objs.splice(j, 0, picked[0]);
      }
      drawStage(); placeFrame(); refreshThumb(cur); changed();
    }
    function command(cmd, at) {
      if (!data) return;
      var slides = data.slides, s = slide(), o1 = selObjs()[0];
      if (cmd.indexOf("present:") === 0) { var w = cmd.slice(8); present(w === "0" ? 0 : cur, w === "presenter"); return; }
      if (readOnly && !/^(notes|present)/.test(cmd)) { O.notify(t("閲覧モードでは書きかえられません")); return; }
      if (cmd.indexOf("shape:") === 0) return addShape(cmd.slice(6));
      if (cmd.indexOf("theme:") === 0) { data.theme = cmd.slice(6); drawThumbs(); drawStage(); changed(); O.sync(); return; }
      if (cmd.indexOf("trans:") === 0) { s.trans = cmd.slice(6) === "none" ? undefined : cmd.slice(6); drawThumbs(true); changed(); O.sync(); return; }
      if (cmd.indexOf("anim:") === 0) { if (!o1) { O.notify(t("アニメーションを付けるものを選んでください")); return; } setObj({ anim: cmd.slice(5) || null }); O.sync(); return; }
      if (cmd.indexOf("align:") === 0) return align(cmd.slice(6));
      if (cmd.indexOf("order:") === 0) return order(cmd.slice(6));
      if (cmd.indexOf("tal:") === 0) return setObj({ align: cmd.slice(4) });
      if (cmd.indexOf("tva:") === 0) return setObj({ va: cmd.slice(4) });
      switch (cmd) {
        case "add": case "addMenu": {
          var addIt = function (lay) {
            var ns = { layout: lay, title: t("タイトルを入力"), objs: [] };
            if (lay === "bullets" || lay === "two") ns.items = [t("テキストを入力")];
            if (lay === "two") ns.items2 = [t("テキストを入力")];
            if (lay === "title" || lay === "section") ns.sub = "";
            if (lay === "blank") delete ns.title;
            slides.splice(cur + 1, 0, ns); cur++;
            sel = []; drawThumbs(); drawStage(); changed(); O.sync();
          };
          if (cmd === "addMenu") return O.menu(at, LAYOUTS).then(function (v) { if (v) addIt(v); });
          return addIt(s.layout === "title" ? "bullets" : s.layout === "section" ? "bullets" : s.layout);
        }
        case "dup": slides.splice(cur + 1, 0, O.copy(s)); slides[cur + 1].objs.forEach(function (o) { o.id = uid(); }); cur++; sel = []; drawThumbs(); drawStage(); changed(); return;
        case "del":
          if (slides.length < 2) { O.notify(t("スライドが 1 枚のときは削除できません")); return; }
          slides.splice(cur, 1); cur = Math.min(cur, slides.length - 1); sel = []; drawThumbs(); drawStage(); changed(); O.notify(t("スライドを削除しました。元に戻すで戻せます")); return;
        case "up": case "down": {
          var j = cur + (cmd === "up" ? -1 : 1);
          if (j < 0 || j >= slides.length) return;
          slides.splice(j, 0, slides.splice(cur, 1)[0]); cur = j; drawThumbs(); drawStage(); changed(); return;
        }
        case "layout":
          return O.menu(at || el("stage"), LAYOUTS.map(function (x) { return { value: x.value, label: x.label, icon: x.icon, checked: s.layout === x.value }; })).then(function (v) {
            if (!v) return;
            s.layout = v;
            if ((v === "bullets" || v === "two") && !(s.items && s.items.length)) s.items = [s.sub || t("テキストを入力")];
            if (v === "two" && !(s.items2 && s.items2.length)) s.items2 = [t("テキストを入力")];
            if (v !== "blank" && !s.title) s.title = t("タイトルを入力");
            drawThumbs(); drawStage(); changed();
          });
        case "text": return addText();
        case "shapeMenu": return O.menu(at, SHAPES).then(function (v) { if (v) addShape(v); });
        case "image": return O.pickImage().then(addImage);
        case "chart": return addChart();
        case "bold": return inlineOrObj("bold", "b");
        case "italic": return inlineOrObj("italic", "i");
        case "underline": return inlineOrObj("underline", "u");
        case "fontSize":
          if (!o1) { O.notify(t("テキスト ボックスや図形を選んでください")); return; }
          return O.menu(at, [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 72, 88, 96].map(function (n) { return { value: String(n), label: String(n), checked: (o1.fs || 20) === n }; })).then(function (v) { if (v) setObj({ fs: +v }); });
        case "grow": case "shrink": if (!o1) return; return setObj({ fs: Math.max(8, Math.min(200, Math.round((o1.fs || 20) * (cmd === "grow" ? 1.15 : 0.87)))) });
        case "color":
          return O.palette(at, { none: t("自動"), value: o1 && o1.color }).then(function (v) {
            if (!v) return;
            var a = getSelection().anchorNode, inText = a && (a.nodeType === 1 ? a : a.parentNode).closest(".sl-tx[contenteditable='true']");
            if (inText && v !== "none") { document.execCommand("styleWithCSS", false, true); document.execCommand("foreColor", false, v); changed(true); return; }
            setObj({ color: v === "none" ? null : v });
          });
        case "fill": return O.palette(at, { none: t("塗りつぶしなし"), value: o1 && o1.fill }).then(function (v) { if (v) setObj({ fill: v }); });
        case "line": return O.palette(at, { none: t("枠線なし"), value: o1 && o1.line }).then(function (v) { if (v) setObj({ line: v }); });
        case "arrange":
          return O.menu(at, [{ head: t("順序") }, { value: "order:front", label: t("最前面へ移動"), icon: "flip_to_front" }, { value: "order:forward", label: t("前面へ移動"), icon: "arrow_upward" },
            { value: "order:backward", label: t("背面へ移動"), icon: "arrow_downward" }, { value: "order:back", label: t("最背面へ移動"), icon: "flip_to_back" },
            { head: t("配置") }, { value: "align:l", label: t("左揃え"), icon: "align_horizontal_left" }, { value: "align:c", label: t("左右中央揃え"), icon: "align_horizontal_center" }, { value: "align:r", label: t("右揃え"), icon: "align_horizontal_right" },
            { value: "align:t", label: t("上揃え"), icon: "align_vertical_top" }, { value: "align:m", label: t("上下中央揃え"), icon: "align_vertical_center" }, { value: "align:b", label: t("下揃え"), icon: "align_vertical_bottom" },
            { sep: 1 }, { value: "rot90", label: t("右へ 90 度回転"), icon: "rotate_right" }, { value: "flipNone", label: t("回転をもとに戻す"), icon: "restart_alt" }]).then(function (v) { if (v) command(v); });
        case "rot90": return setObj({ rot: (((o1 && o1.rot) || 0) + 90) % 360 || null });
        case "flipNone": return setObj({ rot: null });
        case "objDel": {
          if (!sel.length) return;
          var ids = sel.slice();
          s.objs = s.objs.filter(function (o) { return ids.indexOf(o.id) < 0; });
          sel = []; drawStage(); refreshThumb(cur); changed(); O.sync(); return;
        }
        case "objDup": { clipObjs = O.copy(selObjs()); pasteObjs(); return; }
        case "chartEdit": if (o1 && o1.type === "chart") chartDialog(o1.spec).then(function (sp) { if (sp) { o1.spec = sp; redrawObj(o1); refreshThumb(cur); changed(); } }); return;
        case "imgReplace": if (o1 && o1.type === "image") O.pickImage().then(function (src) { if (src) { o1.src = src; redrawObj(o1); refreshThumb(cur); changed(); } }); return;
        case "imgAlt": if (o1 && o1.type === "image") M3.prompt({ title: t("代替テキスト"), label: t("画像の説明（読み上げに使われます）"), value: o1.alt || "", ok: t("OK") }).then(function (v) { if (v !== null) { o1.alt = v; changed(); } }); return;
        case "bg": return O.palette(at, { none: t("テーマの背景"), value: s.bg }).then(function (v) { if (!v) return; s.bg = v === "none" ? undefined : v; drawThumbs(true); drawStage(); changed(); });
        case "bgAll": slides.forEach(function (x) { x.bg = s.bg; }); drawThumbs(true); changed(); O.notify(t("すべてのスライドに同じ背景を使います")); return;
        case "transAll": slides.forEach(function (x) { x.trans = s.trans; }); drawThumbs(true); changed(); O.notify(t("すべてのスライドに同じ切りかえを使います")); return;
        case "ratio":
          return O.menu(at, [{ value: "16:9", label: t("ワイド画面（16:9）"), checked: data.ratio !== "4:3" }, { value: "4:3", label: t("標準（4:3）"), checked: data.ratio === "4:3" }]).then(function (v) { if (v) { data.ratio = v; drawThumbs(); drawStage(); changed(); } });
        case "notes": showNotes = !showNotes; el("notes").hidden = !showNotes; O.sync(); return;
        case "painter":
          if (painter) { painter = null; O.sync(); return; }
          if (!o1) { O.notify(t("図形やテキスト ボックスを選んでください")); return; }
          painter = O.copy(o1); O.sync(); O.notify(t("書式をはり付けるものを押してください")); return;
      }
    }

    // ---------- リボン ----------
    var R = O.groups;
    var ribbon = [
      { tab: t("ホーム"), groups: [R.clip,
        { label: t("スライド"), items: [{ big: 1, icon: "add_box", label: t("新しいスライド"), cmd: "addMenu", drop: 1 },
          { stack: [{ icon: "dashboard", label: t("レイアウト"), cmd: "layout", drop: 1 }, { icon: "content_copy", label: t("複製"), cmd: "dup" }, { icon: "delete", label: t("削除"), cmd: "del" }] }] },
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "20", w: 64, cmd: "fontSize", key: "size", label: t("フォント サイズ") }, { icon: "text_increase", label: t("フォント サイズの拡大"), cmd: "grow" }, { icon: "text_decrease", label: t("フォント サイズの縮小"), cmd: "shrink" }],
          [{ icon: "format_bold", label: t("太字"), cmd: "bold", toggle: "b" }, { icon: "format_italic", label: t("斜体"), cmd: "italic", toggle: "i" }, { icon: "format_underlined", label: t("下線"), cmd: "underline", toggle: "u" },
           { icon: "format_color_text", label: t("フォントの色"), cmd: "color", drop: 1 }]
        ] }] },
        { label: t("段落"), items: [{ rows: [
          [{ icon: "format_align_left", label: t("左揃え"), cmd: "tal:l", toggle: "tal:l" }, { icon: "format_align_center", label: t("中央揃え"), cmd: "tal:c", toggle: "tal:c" }, { icon: "format_align_right", label: t("右揃え"), cmd: "tal:r", toggle: "tal:r" }],
          [{ icon: "vertical_align_top", label: t("上揃え"), cmd: "tva:t", toggle: "tva:t" }, { icon: "vertical_align_center", label: t("上下中央揃え"), cmd: "tva:m", toggle: "tva:m" }, { icon: "vertical_align_bottom", label: t("下揃え"), cmd: "tva:b", toggle: "tva:b" }]
        ] }] },
        { label: t("図形描画"), items: [{ big: 1, icon: "category", label: t("図形"), cmd: "shapeMenu", drop: 1 }, { big: 1, icon: "text_fields", label: t("テキスト ボックス"), cmd: "text" },
          { stack: [{ icon: "format_color_fill", label: t("塗りつぶし"), cmd: "fill", drop: 1 }, { icon: "border_color", label: t("枠線"), cmd: "line", drop: 1 }, { icon: "flip_to_front", label: t("配置"), cmd: "arrange", drop: 1 }] }] },
        { label: t("編集"), items: [{ stack: [{ icon: "search", label: t("検索"), cmd: "find" }, { icon: "find_replace", label: t("置換"), cmd: "replace" }] }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("スライド"), items: [{ big: 1, icon: "add_box", label: t("新しいスライド"), cmd: "addMenu", drop: 1 }] },
        { label: t("画像"), items: [{ big: 1, icon: "image", label: t("画像"), cmd: "image" }] },
        { label: t("図"), items: [{ big: 1, icon: "category", label: t("図形"), cmd: "shapeMenu", drop: 1 }, { big: 1, icon: "bar_chart", label: t("グラフ"), cmd: "chart" }] },
        { label: t("テキスト"), items: [{ big: 1, icon: "text_fields", label: t("テキスト ボックス"), cmd: "text" }] }
      ] },
      { tab: t("デザイン"), groups: [
        { label: t("テーマ"), items: Object.keys(THEMES).map(function (k) { return { big: 1, icon: k === "night" ? "dark_mode" : "palette", label: THEMES[k].label, cmd: "theme:" + k, toggle: "theme:" + k }; }) },
        { label: t("ユーザー設定"), items: [{ big: 1, icon: "aspect_ratio", label: t("スライドのサイズ"), cmd: "ratio", drop: 1 }, { big: 1, icon: "format_paint", label: t("背景の書式"), cmd: "bg", drop: 1 }, { stack: [{ icon: "done_all", label: t("背景をすべてに適用"), cmd: "bgAll" }] }] }
      ] },
      { tab: t("画面切り替え"), groups: [
        { label: t("画面切り替え"), items: TRANS.map(function (x) { return { big: 1, icon: x.icon, label: x.label, cmd: "trans:" + x.value, toggle: "trans:" + x.value }; }) },
        { label: t("タイミング"), items: [{ big: 1, icon: "done_all", label: t("すべてに適用"), cmd: "transAll" }] }
      ] },
      { tab: t("アニメーション"), groups: [
        { label: t("アニメーション"), items: ANIMS.map(function (x) { return { big: 1, icon: x.icon, label: x.label, cmd: "anim:" + x.value, toggle: "anim:" + x.value }; }) },
        { label: t("順序"), items: [{ stack: [{ icon: "arrow_upward", label: t("順番を早くする"), cmd: "order:backward" }, { icon: "arrow_downward", label: t("順番を遅くする"), cmd: "order:forward" }] }] }
      ] },
      { tab: t("スライド ショー"), groups: [
        { label: t("スライド ショーの開始"), items: [{ big: 1, icon: "play_arrow", label: t("最初から"), cmd: "present:0" }, { big: 1, icon: "slideshow", label: t("このスライドから"), cmd: "present:cur" }, { big: 1, icon: "co_present", label: t("発表者ビュー"), cmd: "present:presenter" }] }
      ] },
      { tab: t("表示"), groups: [R.view, { label: t("表示"), items: [{ big: 1, icon: "sticky_note_2", label: t("ノート"), cmd: "notes", toggle: "notes" }] }, R.zoom] },
      { tab: t("図形の書式"), ctx: 1, when: function () { return !!(data && selObjs().length); }, groups: [
        { label: t("スタイル"), items: [{ big: 1, icon: "format_color_fill", label: t("塗りつぶし"), cmd: "fill", drop: 1 }, { big: 1, icon: "border_color", label: t("枠線"), cmd: "line", drop: 1 }, { big: 1, icon: "format_color_text", label: t("文字の色"), cmd: "color", drop: 1 }] },
        { label: t("配置"), items: [{ stack: [{ icon: "flip_to_front", label: t("最前面へ"), cmd: "order:front" }, { icon: "flip_to_back", label: t("最背面へ"), cmd: "order:back" }, { icon: "rotate_right", label: t("回転"), cmd: "rot90" }] },
          { rows: [[{ icon: "align_horizontal_left", label: t("左揃え"), cmd: "align:l" }, { icon: "align_horizontal_center", label: t("左右中央揃え"), cmd: "align:c" }, { icon: "align_horizontal_right", label: t("右揃え"), cmd: "align:r" }],
            [{ icon: "align_vertical_top", label: t("上揃え"), cmd: "align:t" }, { icon: "align_vertical_center", label: t("上下中央揃え"), cmd: "align:m" }, { icon: "align_vertical_bottom", label: t("下揃え"), cmd: "align:b" }]] }] },
        { label: t("そのほか"), items: [{ big: 1, icon: "bar_chart", label: t("データの編集"), cmd: "chartEdit" }, { big: 1, icon: "image", label: t("画像の変更"), cmd: "imgReplace" }, { big: 1, icon: "description", label: t("代替テキスト"), cmd: "imgAlt" }, { big: 1, icon: "delete", label: t("削除"), cmd: "objDel" }] }
      ] }
    ];
    function state() {
      if (!data) return {};
      var o = selObjs()[0], s = slide(), st = { notes: showNotes, painter: !!painter, size: o && o.type !== "image" && o.type !== "chart" ? String(o.fs || 20) : "—" };
      st.b = !!(o && o.b); st.i = !!(o && o.i); st.u = !!(o && o.u);
      ["l", "c", "r"].forEach(function (a) { st["tal:" + a] = !!o && (o.align || (o.type === "shape" ? "c" : "l")) === a; });
      ["t", "m", "b"].forEach(function (a) { st["tva:" + a] = !!o && (o.va || (o.type === "shape" ? "m" : "t")) === a; });
      Object.keys(THEMES).forEach(function (k) { st["theme:" + k] = themeOf(data) === k; });
      TRANS.forEach(function (x) { st["trans:" + x.value] = (s.trans || "none") === x.value; });
      ANIMS.forEach(function (x) { st["anim:" + x.value] = !!o && (o.anim || "") === x.value; });
      return st;
    }

    // ---------- スライド ショー（発表者ビューもここ） ----------
    function present(start, presenter) {
      O.closeMenu();
      endObjEdit();
      var slides = data.slides, i = start, step = 0, t0 = Date.now(), timer = null, black = false;
      var box = document.createElement("div");
      box.className = "present" + (presenter ? " is-presenter" : "");
      box.setAttribute("role", "dialog");
      box.setAttribute("aria-label", t("スライド ショー"));
      box.innerHTML = '<div class="present__main"><div class="present__slide" id="present-slide"></div></div>' +
        '<aside class="present__side"><div class="present__next"><span>' + esc(t("次のスライド")) + '</span><div class="present__nextslide" id="present-next"></div></div>' +
        '<div class="present__notes" id="present-notes"></div><div class="present__time"><strong id="present-timer">0:00</strong><span id="present-clock"></span></div></aside>' +
        '<div class="present__bar"><button type="button" data-p="prev" aria-label="' + esc(t("前へ")) + '">' + icon("chevron_left") + '</button><span id="present-no"></span><button type="button" data-p="next" aria-label="' + esc(t("次へ")) + '">' + icon("chevron_right") + "</button>" +
        '<button type="button" data-p="presenter">' + icon("co_present") + "<span>" + esc(t("発表者ビュー（P）")) + '</span></button><button type="button" data-p="end">' + icon("close") + "<span>" + esc(t("終わる（Esc）")) + "</span></button></div>";
      document.body.appendChild(box);
      box.style.setProperty("--sl-ratio", data.ratio === "4:3" ? "4 / 3" : "16 / 9");
      box.style.setProperty("--sl-rn", String(ratio(data)));
      function anims(n) { return (slides[n].objs || []).filter(function (o) { return o.anim; }); }
      function draw(dir) {
        var host = box.querySelector("#present-slide"), s = slides[i];
        host.innerHTML = slideHtml(s, "show");
        host.classList.remove("tr-fade", "tr-push", "tr-zoom", "tr-back");
        void host.offsetWidth;
        if (s.trans && dir) host.classList.add("tr-" + s.trans, dir < 0 ? "tr-back" : "tr-fwd");
        var list = anims(i);
        list.forEach(function (o, k) { var nd = host.querySelector('[data-obj="' + o.id + '"]'); if (nd) nd.classList.toggle("is-in", k < step); });
        box.querySelector("#present-no").textContent = (i + 1) + " / " + slides.length;
        box.querySelector("#present-next").innerHTML = i + 1 < slides.length ? slideHtml(slides[i + 1], "thumb") : '<div class="present__end">' + esc(t("スライド ショーの最後です")) + "</div>";
        box.querySelector("#present-notes").textContent = s.notes || t("ノートはありません");
      }
      function next() {
        var list = anims(i);
        if (step < list.length) {
          var nd = box.querySelector('#present-slide [data-obj="' + list[step].id + '"]');
          if (nd) nd.classList.add("is-in");
          step++;
          return;
        }
        if (i + 1 >= slides.length) { end(); return; }
        i++; step = 0; draw(1);
      }
      function prev() { if (i === 0) return; i--; step = anims(i).length; draw(-1); }
      function end() {
        clearInterval(timer);
        document.removeEventListener("keydown", onKey, true);
        if (document.fullscreenElement) document.exitFullscreen().catch(function () {});
        box.remove();
        if (i !== cur) { cur = i; drawThumbs(); drawStage(); }
      }
      function tick() {
        var sec = Math.floor((Date.now() - t0) / 1000), now = new Date();
        box.querySelector("#present-timer").textContent = Math.floor(sec / 60) + ":" + ("0" + sec % 60).slice(-2);
        box.querySelector("#present-clock").textContent = now.getHours() + ":" + ("0" + now.getMinutes()).slice(-2);
      }
      function onKey(e) {
        var k = e.key;
        if (k === "Escape") { e.preventDefault(); end(); }
        else if (k === "ArrowRight" || k === "ArrowDown" || k === " " || k === "PageDown" || k === "Enter" || k === "n") { e.preventDefault(); next(); }
        else if (k === "ArrowLeft" || k === "ArrowUp" || k === "PageUp" || k === "Backspace" || k === "p" && e.shiftKey) { e.preventDefault(); prev(); }
        else if (k === "Home") { e.preventDefault(); i = 0; step = 0; draw(-1); }
        else if (k === "End") { e.preventDefault(); i = slides.length - 1; step = 0; draw(1); }
        else if (k === "p" || k === "P") { e.preventDefault(); box.classList.toggle("is-presenter"); }
        else if (k === "b" || k === "B" || k === ".") { e.preventDefault(); black = !black; box.classList.toggle("is-black", black); }
        e.stopPropagation();
      }
      box.addEventListener("click", function (e) {
        var b = e.target.closest("[data-p]");
        if (b) { var a = b.getAttribute("data-p"); if (a === "prev") prev(); else if (a === "next") next(); else if (a === "end") end(); else box.classList.toggle("is-presenter"); return; }
        if (e.target.closest(".present__main")) next();
      });
      document.addEventListener("keydown", onKey, true);
      if (box.requestFullscreen) box.requestFullscreen().catch(function () {});
      document.addEventListener("fullscreenchange", function fs() { if (!document.fullscreenElement && box.isConnected) { document.removeEventListener("fullscreenchange", fs); end(); } });
      timer = setInterval(tick, 1000);
      tick();
      draw(0);
    }

    // ---------- 検索と置換（レイアウトの文字・ものの文字・ノート） ----------
    function texts() {
      var out = [];
      data.slides.forEach(function (s, si) {
        if (s.title) out.push({ si: si, get: function () { return s.title; }, set: function (v) { s.title = v; } });
        if (s.sub) out.push({ si: si, get: function () { return s.sub; }, set: function (v) { s.sub = v; } });
        ["items", "items2"].forEach(function (k) { (s[k] || []).forEach(function (x, i) { out.push({ si: si, get: function () { return s[k][i]; }, set: function (v) { s[k][i] = v; } }); }); });
        (s.objs || []).forEach(function (o) { if (o.html) out.push({ si: si, obj: o.id, get: function () { return plain(o.html); }, set: function (v, q, r, re) { o.html = o.html.replace(re, r); } }); });
        if (s.notes) out.push({ si: si, get: function () { return s.notes; }, set: function (v) { s.notes = v; } });
      });
      return out;
    }
    var findPos = -1;
    function find(q, dir, o) {
      var list = texts(), norm = function (x) { return o.matchCase ? x : x.toLowerCase(); }, hits = list.filter(function (x) { return norm(x.get()).indexOf(norm(q)) >= 0; });
      if (!hits.length) return false;
      var after = hits.filter(function (h) { return dir > 0 ? h.si > cur || (h.si === cur && findPos < hits.indexOf(h)) : h.si < cur; });
      var h = dir > 0 ? (hits.filter(function (x, i) { return x.si > cur || (x.si === cur && i > findPos); })[0] || hits[0]) : (hits.filter(function (x, i) { return x.si < cur || (x.si === cur && i < findPos); }).pop() || hits[hits.length - 1]);
      findPos = hits.indexOf(h);
      if (h.si !== cur) go(h.si);
      if (h.obj) select([h.obj]);
      return t("{n} 件中 {i} 件目", { n: hits.length, i: findPos + 1 }) || after;
    }
    function replace(q, rep, all, o) {
      if (readOnly) return 0;
      var re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), o.matchCase ? "g" : "gi"), n = 0;
      texts().forEach(function (x) {
        if (!all && x.si !== cur) return;
        var v = x.get(), m = v.match(re);
        if (!m) return;
        n += m.length;
        x.set(v.replace(re, rep), q, rep, re);
      });
      if (n) { drawThumbs(true); drawStage(); }
      return n;
    }

    // ---------- 書き出し（.pptx）・読み込み ----------
    var X = function (s) { return Zip.esc(s); };
    function hexOf(c, fb) {
      if (!c || c === "none") return fb || null;
      if (/^var\(/.test(c)) return fb || null;
      var m = /^#([0-9a-f]{6})$/i.exec(c);
      if (m) return m[1].toUpperCase();
      m = /^#([0-9a-f]{3})$/i.exec(c);
      if (m) return m[1].split("").map(function (x) { return x + x; }).join("").toUpperCase();
      m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
      if (m) return [m[1], m[2], m[3]].map(function (x) { return ("0" + (+x).toString(16)).slice(-2); }).join("").toUpperCase();
      return fb || null;
    }
    function svgPng(svg, w, h) {
      return new Promise(function (resolve) {
        var img = new Image();
        img.onload = function () { var cv = document.createElement("canvas"); cv.width = w * 2; cv.height = h * 2; var g = cv.getContext("2d"); g.scale(2, 2); g.drawImage(img, 0, 0, w, h); resolve(cv.toDataURL("image/png")); };
        img.onerror = function () { resolve(null); };
        img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
      });
    }
    function bytesOf(url) {
      var m = /^data:([^;,]+)(;base64)?,(.*)$/.exec(url || "");
      if (!m) return null;
      var bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]), u8 = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      return { mime: m[1], bytes: u8 };
    }
    function pptxFiles() {
      var wide = data.ratio !== "4:3", SW = wide ? 12192000 : 9144000, SH = 6858000, th = THEMES[themeOf(data)];
      var jobs = [];
      // グラフは PNG にする
      var pngs = {};
      data.slides.forEach(function (s) {
        (s.objs || []).forEach(function (o) {
          if (o.type === "chart") { var w = Math.round(o.w * 9.6), h = Math.round(o.h * 9.6 / ratio(data)); jobs.push(svgPng(Chart.svg(o.spec || {}, { w: w, h: h }), w, h).then(function (p) { pngs[o.id] = p; })); }
          if (o.type === "image" && o.src && !/^data:image\/(png|jpe?g|gif)/i.test(o.src)) jobs.push(svgPng('<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600"><image href="' + esc(o.src) + '" width="800" height="600" preserveAspectRatio="xMidYMid meet"/></svg>', 800, 600).then(function (p) { pngs[o.id] = p; }));
        });
      });
      return Promise.all(jobs).then(function () {
        var files = { "[Content_Types].xml": "" }, head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', media = 0, mediaExt = {};
        var NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
        function emuX(p) { return Math.round(p / 100 * SW); }
        function emuY(p) { return Math.round(p / 100 * SH); }
        function runs(text, opt) {
          return '<a:r><a:rPr lang="ja-JP" sz="' + Math.round((opt.sz || 18) * 100) + '"' + (opt.b ? ' b="1"' : "") + (opt.i ? ' i="1"' : "") + (opt.u ? ' u="sng"' : "") + ' dirty="0"><a:solidFill><a:srgbClr val="' + (opt.color || hexOf(th.ink)) + '"/></a:solidFill>' +
            '<a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/></a:rPr><a:t>' + X(text) + "</a:t></a:r>";
        }
        function paras(lines, opt) {
          return lines.map(function (l) {
            return "<a:p><a:pPr algn=\"" + ({ l: "l", c: "ctr", r: "r" }[opt.align] || "l") + '"' + (opt.bullet ? ' marL="342900" indent="-342900"' : "") + ">" + (opt.bullet ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : "<a:buNone/>") + "</a:pPr>" +
              (l ? runs(l, opt) : "") + '<a:endParaRPr lang="ja-JP" sz="' + Math.round((opt.sz || 18) * 100) + '" dirty="0"/></a:p>';
          }).join("");
        }
        var nid;
        function sp(name, x, y, w, h, body, o) {
          o = o || {};
          var geom = o.geom || "rect", fill = o.fill ? '<a:solidFill><a:srgbClr val="' + o.fill + '"/></a:solidFill>' : "<a:noFill/>", ln = o.line ? '<a:ln w="19050"><a:solidFill><a:srgbClr val="' + o.line + '"/></a:solidFill></a:ln>' : "<a:ln><a:noFill/></a:ln>";
          nid++;
          return '<p:sp><p:nvSpPr><p:cNvPr id="' + nid + '" name="' + X(name) + " " + nid + '"/><p:cNvSpPr' + (o.box ? ' txBox="1"' : "") + "/><p:nvPr/></p:nvSpPr><p:spPr>" +
            '<a:xfrm' + (o.rot ? ' rot="' + Math.round(o.rot * 60000) + '"' : "") + '><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm><a:prstGeom prst="' + geom + '"><a:avLst/></a:prstGeom>' + fill + ln + "</p:spPr>" +
            '<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="' + ({ t: "t", m: "ctr", b: "b" }[o.va] || "t") + '"><a:normAutofit/></a:bodyPr><a:lstStyle/>' + (body || "<a:p><a:endParaRPr lang=\"ja-JP\"/></a:p>") + "</p:txBody></p:sp>";
        }
        function pic(name, x, y, w, h, rid, rot) {
          nid++;
          return '<p:pic><p:nvPicPr><p:cNvPr id="' + nid + '" name="' + X(name) + " " + nid + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
            '<p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm' + (rot ? ' rot="' + Math.round(rot * 60000) + '"' : "") + '><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
        }
        var GEOM = { rect: "rect", round: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", arrow: "rightArrow", star: "star5", callout: "wedgeRoundRectCallout", line: "line" };
        var slideIds = [], presRels = [], ctypes = [];
        data.slides.forEach(function (s, n) {
          nid = 1;
          var rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'], relN = 1, shapes = [];
          var lay = s.layout || "title", sec = lay === "section", ink = hexOf(sec ? th.secInk : th.ink), sub = hexOf(sec ? th.secInk : th.sub);
          var bg = hexOf(s.bg) || hexOf(sec ? th.sec : th.bg);
          if (lay === "title" || lay === "section") {
            if (s.title) shapes.push(sp("Title", emuX(8), emuY(lay === "section" ? 34 : 28), emuX(84), emuY(20), paras([s.title], { sz: 44, b: 1, color: ink, align: lay === "section" ? "c" : "l" }), { va: "b", box: 1 }));
            if (s.sub) shapes.push(sp("Subtitle", emuX(8), emuY(lay === "section" ? 56 : 52), emuX(84), emuY(14), paras(String(s.sub).split("\n"), { sz: 20, color: sub, align: lay === "section" ? "c" : "l" }), { box: 1 }));
          } else if (lay !== "blank") {
            if (s.title) shapes.push(sp("Title", emuX(8), emuY(6), emuX(84), emuY(14), paras([s.title], { sz: 32, b: 1, color: ink }), { va: "m", box: 1 }));
            if (lay === "bullets") shapes.push(sp("Content", emuX(8), emuY(24), emuX(84), emuY(66), paras((s.items || []).filter(function (x) { return x; }), { sz: 22, color: ink, bullet: 1 }), { box: 1 }));
            if (lay === "two") {
              shapes.push(sp("Content", emuX(8), emuY(24), emuX(40), emuY(66), paras((s.items || []).filter(function (x) { return x; }), { sz: 20, color: ink, bullet: 1 }), { box: 1 }));
              shapes.push(sp("Content", emuX(52), emuY(24), emuX(40), emuY(66), paras((s.items2 || []).filter(function (x) { return x; }), { sz: 20, color: ink, bullet: 1 }), { box: 1 }));
            }
          }
          (s.objs || []).forEach(function (o) {
            var x = emuX(o.x), y = emuY(o.y), w = Math.max(1, emuX(o.w)), h = Math.max(1, emuY(o.h));
            if (o.type === "image" || o.type === "chart") {
              var src = pngs[o.id] || o.src, d = bytesOf(src);
              if (!d) return;
              var ext = /png/.test(d.mime) ? "png" : /gif/.test(d.mime) ? "gif" : "jpeg";
              media++; mediaExt[ext] = 1;
              var mname = "image" + media + "." + ext;
              files["ppt/media/" + mname] = d.bytes;
              relN++;
              rels.push('<Relationship Id="rId' + relN + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + mname + '"/>');
              shapes.push(pic(o.type === "chart" ? "Chart" : "Picture", x, y, w, h, "rId" + relN, o.rot));
              return;
            }
            var lines = plain(o.html).split("\n"), isShape = o.type === "shape";
            var color = hexOf(o.color) || (isShape && o.fill !== "none" ? "FFFFFF" : hexOf(th.ink));
            shapes.push(sp(isShape ? "Shape" : "TextBox", x, y, w, h, o.shape === "line" ? "" : paras(lines, { sz: o.fs || (isShape ? 18 : 20), b: o.b, i: o.i, u: o.u, color: color, align: o.align || (isShape ? "c" : "l") }),
              { geom: isShape ? GEOM[o.shape] || "rect" : "rect", fill: isShape && o.shape !== "line" ? (o.fill === "none" ? null : hexOf(o.fill, hexOf(th.accent))) : null, line: o.shape === "line" ? hexOf(o.line, hexOf(th.ink)) : hexOf(o.line), rot: o.rot, va: o.va || (isShape ? "m" : "t"), box: !isShape }));
          });
          var trans = s.trans === "fade" ? '<p:transition spd="med"><p:fade/></p:transition>' : s.trans === "push" ? '<p:transition spd="med"><p:push dir="l"/></p:transition>' : s.trans === "zoom" ? '<p:transition spd="med"><p:zoom/></p:transition>' : "";
          files["ppt/slides/slide" + (n + 1) + ".xml"] = head + "<p:sld " + NS + "><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"" + bg + '"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' + shapes.join("") + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" + trans + "</p:sld>";
          files["ppt/slides/_rels/slide" + (n + 1) + ".xml.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join("") + "</Relationships>";
          slideIds.push('<p:sldId id="' + (256 + n) + '" r:id="rId' + (n + 10) + '"/>');
          presRels.push('<Relationship Id="rId' + (n + 10) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + (n + 1) + '.xml"/>');
          ctypes.push('<Override PartName="/ppt/slides/slide' + (n + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
        });
        files["ppt/presentation.xml"] = head + "<p:presentation " + NS + ' saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' + slideIds.join("") + "</p:sldIdLst>" +
          '<p:sldSz cx="' + SW + '" cy="' + SH + '"' + (wide ? "" : ' type="screen4x3"') + '/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr></p:defaultTextStyle></p:presentation>';
        files["ppt/_rels/presentation.xml.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' +
          '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' + presRels.join("") + "</Relationships>";
        files["ppt/presProps.xml"] = head + "<p:presentationPr " + NS + "/>";
        files["ppt/viewProps.xml"] = head + "<p:viewPr " + NS + '><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>';
        files["ppt/tableStyles.xml"] = head + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
        var spTreeEmpty = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';
        files["ppt/slideMasters/slideMaster1.xml"] = head + "<p:sldMaster " + NS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' + spTreeEmpty + "</p:cSld>" +
          '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
          '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>';
        files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>';
        files["ppt/slideLayouts/slideLayout1.xml"] = head + "<p:sldLayout " + NS + ' type="blank" preserve="1"><p:cSld name="Blank">' + spTreeEmpty + "</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";
        files["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>';
        var c = function (v) { return '<a:srgbClr val="' + v + '"/>'; };
        var fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
        files["ppt/theme/theme1.xml"] = head + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SK Hub Office"><a:themeElements><a:clrScheme name="SK Hub Office">' +
          "<a:dk1>" + c(hexOf(th.ink)) + "</a:dk1><a:lt1>" + c(hexOf(th.bg)) + "</a:lt1><a:dk2>" + c(hexOf(th.sub)) + "</a:dk2><a:lt2>" + c(hexOf(th.soft)) + "</a:lt2>" +
          "<a:accent1>" + c(hexOf(th.accent)) + "</a:accent1><a:accent2>" + c("2563EB") + "</a:accent2><a:accent3>" + c("16A34A") + "</a:accent3><a:accent4>" + c("F59E0B") + "</a:accent4><a:accent5>" + c("DB2777") + "</a:accent5><a:accent6>" + c("7C3AED") + "</a:accent6>" +
          "<a:hlink>" + c("1D4ED8") + "</a:hlink><a:folHlink>" + c("7C3AED") + "</a:folHlink></a:clrScheme>" +
          '<a:fontScheme name="SK Hub Office"><a:majorFont><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
          '<a:fmtScheme name="SK Hub Office"><a:fillStyleLst>' + fill3 + fill3 + fill3 + '</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
          "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>" + fill3 + fill3 + fill3 + "</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>";
        files["_rels/.rels"] = head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
        files["docProps/core.xml"] = head + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
          "<dc:title>" + X(O.file() ? O.file().name : "") + "</dc:title><dc:creator>" + X(O.userName() || "SK Hub Office") + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString().replace(/\.\d+Z$/, "Z") + "</dcterms:created></cp:coreProperties>";
        files["[Content_Types].xml"] = head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          Object.keys(mediaExt).map(function (e) { return '<Default Extension="' + e + '" ContentType="image/' + e + '"/>'; }).join("") +
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
          '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
          '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
          '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
          '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
          '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' +
          '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ctypes.join("") + "</Types>";
        return files;
      });
    }
    function exportAs(id) {
      endObjEdit();
      if (id === "pptx") {
        return pptxFiles().then(Zip.write).then(function (b) {
          var mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
          return O.saveBlob(O.fileName(".pptx"), new Blob([b], { type: mime }), t("PowerPoint プレゼンテーション"), mime, ".pptx");
        });
      }
      if (id === "outline") {
        var text = data.slides.map(function (s, i) {
          var lines = [(i + 1) + ". " + (s.title || "")];
          if (s.sub) lines.push("   " + s.sub);
          (s.items || []).concat(s.items2 || []).forEach(function (x) { if (x) lines.push("   ・" + x); });
          (s.objs || []).forEach(function (o) { var p = plain(o.html); if (p) lines.push("   " + p.replace(/\n/g, " ")); });
          if (s.notes) lines.push("   " + t("ノート") + ": " + s.notes.replace(/\n/g, " "));
          return lines.join("\r\n");
        }).join("\r\n\r\n");
        return O.saveBlob(O.fileName(".txt"), new Blob(["﻿" + text + "\r\n"], { type: "text/plain" }), t("アウトライン"), "text/plain", ".txt");
      }
    }
    function kids(node, name) { return node ? Array.prototype.slice.call(node.getElementsByTagNameNS("*", name)) : []; }
    function kid(node, name) { var k = node && node.firstElementChild; while (k) { if (k.localName === name) return k; k = k.nextElementSibling; } return null; }
    function b64(u8) { var s = ""; for (var i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); }
    function readPptx(file) {
      if (!Zip.canRead) return Promise.reject(new Error("no-inflate"));
      return Zip.read(file).then(function (z) {
        var pres = Zip.xml(z["ppt/presentation.xml"]);
        if (!pres) throw new Error("not-pptx");
        var RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var prel = {};
        kids(Zip.xml(z["ppt/_rels/presentation.xml.rels"]), "Relationship").forEach(function (r) { prel[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var sz = kids(pres, "sldSz")[0], SW = sz ? +sz.getAttribute("cx") : 12192000, SH = sz ? +sz.getAttribute("cy") : 6858000;
        var out = { slides: [], theme: "mikan", ratio: Math.abs(SW / SH - 4 / 3) < 0.05 ? "4:3" : "16:9" };
        kids(pres, "sldId").forEach(function (sid) {
          var target = prel[sid.getAttributeNS(RNS, "id") || sid.getAttribute("r:id")];
          if (!target) return;
          var path = "ppt/" + target.replace(/^\/?ppt\//, "").replace(/^\//, ""), doc = Zip.xml(z[path]);
          if (!doc) return;
          var relPath = path.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels"), srel = {};
          kids(Zip.xml(z[relPath]), "Relationship").forEach(function (r) { srel[r.getAttribute("Id")] = { target: r.getAttribute("Target"), type: r.getAttribute("Type") || "" }; });
          var s = { layout: "blank", objs: [] }, bodyItems = null;
          function geo(node) {
            var off = kids(node, "off")[0], ext = kids(node, "ext")[0], xf = kids(node, "xfrm")[0];
            if (!off || !ext) return null;
            var rot = xf && xf.getAttribute("rot") ? Math.round(+xf.getAttribute("rot") / 60000) : 0;
            return { x: +off.getAttribute("x") / SW * 100, y: +off.getAttribute("y") / SH * 100, w: +ext.getAttribute("cx") / SW * 100, h: +ext.getAttribute("cy") / SH * 100, rot: rot || undefined };
          }
          function parasOf(sp) { return kids(sp, "p").map(function (p) { return kids(p, "t").map(function (x) { return x.textContent; }).join(""); }); }
          kids(doc, "sp").forEach(function (sp) {
            var ph = kids(sp, "ph")[0], type = ph ? ph.getAttribute("type") || "body" : null, lines = parasOf(sp), text = lines.join("\n").trim();
            if (type === "title" || type === "ctrTitle") { s.title = text; if (type === "ctrTitle") s.layout = "title"; else if (s.layout === "blank") s.layout = "titleOnly"; return; }
            if (type === "subTitle") { s.sub = text; s.layout = "title"; return; }
            if ((type === "body" || type === "obj") && !bodyItems) { bodyItems = lines.filter(function (x) { return x.trim(); }); return; }
            var g = geo(sp);
            if (!g || (!text && !kids(sp, "prstGeom").length)) return;
            var prst = kids(sp, "prstGeom")[0], geom = prst ? prst.getAttribute("prst") : "rect", isText = !!kids(sp, "cNvSpPr").filter(function (x) { return x.getAttribute("txBox") === "1"; }).length || geom === "rect" && !kids(kid(kids(sp, "spPr")[0], "solidFill") || null, "srgbClr").length;
            var rpr = kids(sp, "rPr")[0], fsz = rpr && rpr.getAttribute("sz") ? +rpr.getAttribute("sz") / 100 : null;
            var fillEl = kid(kids(sp, "spPr")[0], "solidFill"), fill = fillEl && kids(fillEl, "srgbClr")[0] ? "#" + kids(fillEl, "srgbClr")[0].getAttribute("val").toLowerCase() : null;
            var colEl = rpr && kids(rpr, "srgbClr")[0], color = colEl ? "#" + colEl.getAttribute("val").toLowerCase() : null;
            var shapeMap = { roundRect: "round", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", rightArrow: "arrow", star5: "star", wedgeRoundRectCallout: "callout", wedgeRectCallout: "callout", line: "line" };
            var o = { id: uid(), type: isText ? "text" : "shape", x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot, html: lines.map(esc).join("<br>"), fs: fsz || undefined, color: color || undefined };
            if (!isText) { o.shape = shapeMap[geom] || "rect"; o.fill = fill || undefined; }
            if (rpr && rpr.getAttribute("b") === "1") o.b = 1;
            s.objs.push(o);
          });
          kids(doc, "pic").forEach(function (p) {
            var blip = kids(p, "blip")[0], id = blip && (blip.getAttributeNS(RNS, "embed") || blip.getAttribute("r:embed")), r = id && srel[id], g = geo(p);
            if (!r || !g) return;
            var mp = "ppt/" + r.target.replace(/^\.\.\//, "").replace(/^\/?ppt\//, ""), bytes = z[mp];
            if (!bytes) return;
            var ext = (mp.split(".").pop() || "png").toLowerCase();
            if (/emf|wmf|tif/.test(ext)) return;
            s.objs.push({ id: uid(), type: "image", src: "data:image/" + (ext === "jpg" ? "jpeg" : ext === "svg" ? "svg+xml" : ext) + ";base64," + b64(bytes), x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot });
          });
          if (bodyItems) { s.items = bodyItems.length ? bodyItems : [""]; s.layout = "bullets"; if (!s.title) s.title = ""; }
          // ノート
          Object.keys(srel).forEach(function (id) {
            if (!/notesSlide$/.test(srel[id].type)) return;
            var nd = Zip.xml(z["ppt/" + srel[id].target.replace(/^\.\.\//, "")]);
            if (!nd) return;
            var body = kids(nd, "sp").filter(function (sp) { var ph = kids(sp, "ph")[0]; return ph && ph.getAttribute("type") === "body"; })[0];
            if (body) s.notes = parasOf(body).join("\n").trim();
          });
          out.slides.push(s);
        });
        if (!out.slides.length) out.slides.push({ layout: "title", title: "", sub: "" });
        return upgrade(out);
      });
    }
    function importFile(file) {
      if (/\.pptx$/i.test(file.name)) return readPptx(file);
      // テキスト（アウトライン）: 空行で区切ったまとまりを 1 枚に。1 行目がタイトル、のこりが箇条書き
      return file.text().then(function (text) {
        var blocks = text.replace(/^﻿/, "").split(/\r?\n\s*\r?\n/).map(function (b) { return b.split(/\r?\n/).map(function (l) { return l.replace(/^\s*(\d+\.|[-・*•])\s*/, "").trim(); }).filter(Boolean); }).filter(function (b) { return b.length; });
        var slides = blocks.map(function (b, i) { return i === 0 && b.length <= 2 ? { layout: "title", title: b[0], sub: b[1] || "" } : { layout: "bullets", title: b[0], items: b.slice(1).length ? b.slice(1) : [""] }; });
        return upgrade({ slides: slides.length ? slides : [{ layout: "title", title: "", sub: "" }] });
      });
    }

    // ---------- 印刷（1 枚を 1 ページに） ----------
    function beforePrint() {
      endObjEdit();
      var box = document.createElement("div");
      box.className = "deck-print";
      box.id = "deck-print";
      box.style.setProperty("--sl-ratio", data.ratio === "4:3" ? "4 / 3" : "16 / 9");
      box.innerHTML = data.slides.map(function (s) { return '<div class="deck-print__slide">' + slideHtml(s, "thumb") + "</div>"; }).join("");
      $("canvas").appendChild(box);
      return "@page { size: " + (data.ratio === "4:3" ? "254mm 190.5mm" : "254mm 142.9mm") + "; margin: 0; }";
    }
    function afterPrint() { var p = el("deck-print"); if (p) p.remove(); }

    return {
      ribbon: ribbon,
      upgrade: upgrade,
      thumb: function (d) { d = upgrade(O.copy(d)); return '<div class="th th--slide" aria-hidden="true" style="--sl-ratio:' + (d.ratio === "4:3" ? "4 / 3" : "16 / 9") + '">' + slideHtml(d.slides[0], "thumb", d) + "</div>"; },
      render: render,
      command: command,
      state: state,
      zoom: function (z) { zoomZ = z; applyZoom(); },
      fitZoom: function () { return 1; },
      focus: function () { if (el("stage")) el("stage").focus({ preventScroll: true }); },
      open: function () { cur = 0; sel = []; },
      close: function () { endObjEdit(); data = null; },
      beforeSave: function () {},
      beforeUndo: function () { endObjEdit(); },
      setRead: function (on) { readOnly = on; sel = []; drawStage(); el("notes").contentEditable = on ? "false" : "true"; O.sync(); },
      selectedText: function () { return getSelection().toString(); },
      find: find,
      replace: replace,
      findClose: function () { findPos = -1; },
      clip: function (cmd) {
        if (sel.length && !editingObj) {
          if (cmd === "copy" || cmd === "cut") { clipObjs = O.copy(selObjs()); if (cmd === "cut") command("objDel"); O.notify(t("コピーしました")); return; }
        }
        if (cmd === "paste" && clipObjs) { pasteObjs(); return; }
        if (cmd === "paste") {
          if (navigator.clipboard && navigator.clipboard.read) {
            navigator.clipboard.read().then(function (items) {
              var it = items[0], img = it && it.types.filter(function (x) { return /^image\//.test(x); })[0];
              if (img) return it.getType(img).then(function (b) { return O.imageData(new File([b], "paste", { type: img })); }).then(addImage);
              return it.getType("text/plain").then(function (b) { return b.text(); }).then(function (tx) { if (tx.trim()) addObj({ type: "text", x: 20, y: 30, w: 60, h: 20, html: esc(tx).replace(/\n/g, "<br>"), fs: 20 }); });
            }).catch(function () { O.notify(t("Ctrl+V で貼り付けてください")); });
          }
          return;
        }
        if (!document.execCommand(cmd)) O.notify(t("文字を選んでください"));
      },
      key: onKey,
      beforePrint: beforePrint,
      afterPrint: afterPrint,
      exports: [
        { id: "pptx", icon: "slideshow", label: t("PowerPoint（.pptx）"), desc: t("Microsoft PowerPoint・LibreOffice・Google スライドで開けます") },
        { id: "outline", icon: "format_list_bulleted", label: t("アウトライン（.txt）"), desc: t("スライドの文字とノートを書き出します") }
      ],
      exportAs: exportAs,
      imports: { label: ".pptx・.txt", accept: ".pptx,.txt" },
      importFile: importFile,
      props: function () { var n = 0; data.slides.forEach(function (s) { n += (s.objs || []).length; }); return [[t("スライド"), String(data.slides.length)], [t("図形・画像など"), String(n)], [t("テーマ"), THEMES[themeOf(data)].label]]; },
      keys: [["F5 / Shift+F5", t("最初から・このスライドからスライド ショー")], ["Ctrl+M", t("新しいスライド")], ["Ctrl+D", t("選んだものの複製")], ["Delete", t("選んだものの削除")],
        ["矢印 / Shift+矢印", t("選んだものを少し・大きく動かす")], ["Enter / F2", t("選んだものの文字を書く")], ["Shift+ドラッグ", t("まっすぐ動かす・縦横の比を保つ")], ["Alt+ドラッグ", t("線に合わせずに動かす")],
        ["P（スライド ショー）", t("発表者ビュー")], ["B（スライド ショー）", t("画面を黒くする")]]
    };
  };
})();
