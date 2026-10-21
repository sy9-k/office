// SK Hub Office の Sheet（表計算）。window.OfficeApps.sheet に登録し、/office/shared/office.js が呼ぶ（O は office.js が渡す道具）
//   ファイルの中身（data）: { v: 2, cur: 開いているシートの番号, sheets: [シート] }
//     シート: { name, cells: { "A1": "入れた文字（= で始まると数式）" }, style: { "A1": 書式 }, cw: { 列の番号: 幅 px }, rh: { 行の番号: 高さ px },
//              merges: ["A1:C1"], freeze: { r: 固定する行の数, c: 固定する列の数 }, head: 見出しの行, total: 合計の行, grid: false（目盛り線を消す）,
//              filter: { r: 見出しの行, c1, c2, hide: { 列の番号: [隠す文字] } }, cf: [条件付き書式], charts: [{ id, type, ref, title, x, y, w, h }] }
//     書式: { b, i, u, s（取り消し線）, ff（フォント）, fs（大きさ pt）, fg（文字の色）, bg（塗りつぶし）, al: "l|c|r", va: "t|m|b", wrap, nf（表示形式）, bd: { t, r, b, l } }
//     条件付き書式: { ref: "B2:B9", type: "gt|lt|eq|text|dup|bar|scale", v, color }
//   0.1 の形（{ cells, fmt: { C: "yen" }, head, total }）は upgrade で 2 の形に直す
//   画面: 数式バー → 表（見えている分だけ作り、下へスクロールすると足す）→ シートのタブ。選んだところの枠・入力の欄・グラフは表の上に重ねる
//   入力の欄（#ced）は、選んだセルの上にいつも置いておき、文字を打つとそのまま入力になる（日本語の入力もそのまま使える）
//   計算は /office/shared/formula.js、グラフは chart.js、.xlsx の読み書きは zip.js
(function () {
  "use strict";
  window.OfficeApps = window.OfficeApps || {};
  window.OfficeApps.sheet = function (O) {
    var t = O.t, esc = O.esc, icon = O.icon, $ = O.$, F = window.OfficeFormula, Chart = window.OfficeChart, Zip = window.OfficeZip;
    var COLW = 96, HEADW = 44, EMPTY = {};

    // ---------- セルの番地 ----------
    function ref(c, r) { return F.colName(c) + r; }
    function parseRef(s) { var m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(s).trim()); return m ? { c: F.colIndex(m[1]), r: +m[2] } : null; }
    function norm(g) { return { c1: Math.min(g.c1, g.c2), r1: Math.min(g.r1, g.r2), c2: Math.max(g.c1, g.c2), r2: Math.max(g.r1, g.r2) }; }
    function parseRange(s) {
      var p = String(s).replace(/^.*!/, "").split(":"), a = parseRef(p[0]), b = p[1] ? parseRef(p[1]) : a;
      return a && b ? norm({ c1: a.c, r1: a.r, c2: b.c, r2: b.r }) : null;
    }
    function rangeText(g) { return ref(g.c1, g.r1) + (g.c1 === g.c2 && g.r1 === g.r2 ? "" : ":" + ref(g.c2, g.r2)); }
    function inRange(g, c, r) { return c >= g.c1 && c <= g.c2 && r >= g.r1 && r <= g.r2; }

    // ---------- 0.1 の形を直す ----------
    function fixSheet(s, i) {
      s.name = s.name || "Sheet" + (i + 1);
      s.cells = s.cells || {}; s.style = s.style || {}; s.cw = s.cw || {}; s.rh = s.rh || {};
      s.merges = s.merges || []; s.charts = s.charts || []; s.cf = s.cf || [];
      return s;
    }
    function upgrade(d) {
      if (!d.sheets || !d.sheets.length) {
        var s = { name: "Sheet1", cells: d.cells || {}, style: d.style || {}, head: d.head, total: d.total };
        // 0.1 の列の書式（円で表示する列）
        Object.keys(d.fmt || {}).forEach(function (col) {
          if (d.fmt[col] !== "yen") return;
          Object.keys(s.cells).forEach(function (k) {
            var p = parseRef(k);
            if (!p || F.colName(p.c) !== col || p.r === d.head) return;
            s.style[k] = s.style[k] || {};
            if (!s.style[k].nf) s.style[k].nf = "¥#,##0";
          });
        });
        ["cells", "fmt", "style", "head", "total"].forEach(function (k) { delete d[k]; });
        d.sheets = [s];
        d.cur = 0;
      }
      d.sheets.forEach(fixSheet);
      if (!(d.cur >= 0 && d.cur < d.sheets.length)) d.cur = 0;
      d.v = 2;
      return d;
    }

    // ---------- 計算（ブック全体。変えたら reset して、見るときに計算しなおす） ----------
    function literal(x, st) {
      if (st && st.nf === "@") return x;
      if (x.charAt(0) === "'") return x.slice(1);
      var u = x.toUpperCase();
      if (u === "TRUE") return true;
      if (u === "FALSE") return false;
      var n = F.parseNum(x);
      return n === null ? x : n;
    }
    function Book(d) {
      var cache = {}, busy = {}, lastCache = {};
      function sheetIndex(name) {
        name = String(name).toLowerCase();
        for (var i = 0; i < d.sheets.length; i++) if (d.sheets[i].name.toLowerCase() === name) return i;
        return -1;
      }
      function val(s, c, r) {
        var k = s + "!" + c + "," + r;
        if (k in cache) return cache[k];
        var sh = d.sheets[s];
        if (!sh) return F.err(F.ERR.REF);
        var key = ref(c, r), x = sh.cells[key], v;
        if (x === undefined || x === "") v = null;
        else if (x.charAt(0) === "=" && x.length > 1) {
          if (busy[k]) return F.err(F.ERR.CIRC);
          busy[k] = 1;
          try { v = F.evaluate(F.parse(x.slice(1)), { get: val, sheet: sheetIndex, here: { s: s, c: c, r: r }, last: last }); }
          catch (e) { v = F.err(F.ERR.VALUE); }
          delete busy[k];
        } else v = literal(x, sh.style[key]);
        cache[k] = v;
        return v;
      }
      function last(s) {
        if (lastCache[s]) return lastCache[s];
        var sh = d.sheets[s], rows = 0, cols = 0;
        if (sh) Object.keys(sh.cells).forEach(function (k) { var p = parseRef(k); if (p) { if (p.r > rows) rows = p.r; if (p.c + 1 > cols) cols = p.c + 1; } });
        return (lastCache[s] = { rows: rows, cols: cols });
      }
      return { val: val, last: last, sheetIndex: sheetIndex, reset: function () { cache = {}; busy = {}; lastCache = {}; } };
    }

    // ---------- 表示（値と表示形式から、出す文字を決める） ----------
    function show(v, nf) {
      if (v === null || v === undefined) return { text: "" };
      if (F.isErr(v)) return { text: v.err, err: true };
      if (typeof v === "boolean") return { text: v ? "TRUE" : "FALSE", bool: true };
      if (typeof v === "number") { var f = F.format(v, nf && nf !== "@" ? nf : null); return { text: f.text, color: f.color, num: true }; }
      return { text: String(v) };
    }
    // 入れた文字から、表示形式を決める（¥1,200 → 通貨、50% → パーセント、2026/10/21 → 日付）
    function autoNf(x) {
      if (x.charAt(0) === "=") {
        if (/^=\s*NOW\(/i.test(x)) return "yyyy/m/d h:mm";
        if (/^=\s*(TODAY|DATE|EDATE|EOMONTH|WORKDAY|DATEVALUE)\(/i.test(x)) return "yyyy/m/d";
        return null;
      }
      if (/^[-+]?[¥￥]/.test(x) && F.parseNum(x) !== null) return /\.\d/.test(x) ? "¥#,##0.00" : "¥#,##0";
      if (/%$/.test(x) && F.parseNum(x) !== null) return /\.\d/.test(x) ? "0.0%" : "0%";
      if (/^\d{4}[\/\-.年]\d{1,2}[\/\-.月]\d{1,2}日?$/.test(x) && F.parseDate(x) !== null) return /年/.test(x) ? "yyyy年m月d日" : "yyyy/m/d";
      if (/^\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}\s+\d/.test(x) && F.parseDate(x) !== null) return "yyyy/m/d h:mm";
      if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(x) && F.parseDate(x) !== null) return /:\d{2}:/.test(x) ? "h:mm:ss" : "h:mm";
      if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(x)) return /\./.test(x) ? "#,##0.00" : "#,##0";
      return null;
    }
    var NF = [
      { v: "", label: t("標準") }, { v: "0", label: t("数値") }, { v: "#,##0", label: t("桁区切り") }, { v: "#,##0.00", label: t("数値（小数 2 けた）") },
      { v: "¥#,##0", label: t("通貨") }, { v: "0%", label: t("パーセント") }, { v: "0.0%", label: t("パーセント（小数 1 けた）") },
      { v: "yyyy/m/d", label: t("日付") }, { v: "yyyy年m月d日", label: t("長い日付") }, { v: "m月d日(aaa)", label: t("曜日つきの日付") },
      { v: "h:mm", label: t("時刻") }, { v: "0.00E+00", label: t("指数") }, { v: "@", label: t("文字列") }
    ];
    function nfName(nf) {
      if (!nf) return t("標準");
      for (var i = 0; i < NF.length; i++) if (NF[i].v === nf) return NF[i].label;
      return /[¥$]/.test(nf) ? t("通貨") : /%/.test(nf) ? t("パーセント") : F.isDateFormat(nf) ? t("日付") : t("ユーザー定義");
    }
    function bdCss(bd) {
      var a = [];
      if (bd.t) a.push("inset 0 " + (bd.t > 1 ? 2 : 1) + "px 0 #1f2328");
      if (bd.b) a.push("inset 0 -" + (bd.b > 1 ? 2 : 1) + "px 0 #1f2328");
      if (bd.l) a.push("inset " + (bd.l > 1 ? 2 : 1) + "px 0 0 #1f2328");
      if (bd.r) a.push("inset -" + (bd.r > 1 ? 2 : 1) + "px 0 0 #1f2328");
      return a.join(",");
    }

    // ---------- いまの状態 ----------
    var data = null, book = null, si = 0, sh = null;
    // act: 選んでいるセル（入力するところ）。anc: 範囲のもう一方の角（Shift+矢印・ドラッグで動く）
    var act = { c: 0, r: 1 }, anc = { c: 0, r: 1 };
    var nRows = 100, nCols = 26;
    var editing = false, editMode = "enter", refIns = null;
    var chartSel = null, clip = null, painter = null, zoomZ = 1;
    var mtop = {}, covered = {}, hiddenRows = {}, fzTop = {}, fzLeft = {}, cfCache = null;
    function el(id) { return $(id); }
    function range() { return norm({ c1: anc.c, r1: anc.r, c2: act.c, r2: act.r }); }
    function isOne(g) { return g.c1 === g.c2 && g.r1 === g.r2; }
    function rawAt(c, r) { return sh.cells[ref(c, r)] || ""; }
    function styleAt(c, r) { return sh.style[ref(c, r)] || EMPTY; }
    function readOnly() { return O.readMode(); }

    // 結合したセル（左上のセルに colspan・rowspan を付け、ほかのセルは作らない）
    function indexMerges() {
      mtop = {}; covered = {};
      sh.merges.forEach(function (m) {
        var g = parseRange(m);
        if (!g) return;
        mtop[ref(g.c1, g.r1)] = { cs: g.c2 - g.c1 + 1, rs: g.r2 - g.r1 + 1, g: g };
        for (var r = g.r1; r <= g.r2; r++) for (var c = g.c1; c <= g.c2; c++) if (c !== g.c1 || r !== g.r1) covered[ref(c, r)] = g;
      });
    }
    function mergeAt(c, r) {
      var k = ref(c, r);
      if (mtop[k]) return mtop[k].g;
      return covered[k] || null;
    }
    // 結合したセルを含むように、範囲を広げる
    function grow(g) {
      var changed = true;
      g = norm(g);
      while (changed) {
        changed = false;
        sh.merges.forEach(function (m) {
          var x = parseRange(m);
          if (!x || x.c2 < g.c1 || x.c1 > g.c2 || x.r2 < g.r1 || x.r1 > g.r2) return;
          if (x.c1 < g.c1 || x.c2 > g.c2 || x.r1 < g.r1 || x.r2 > g.r2) {
            g = { c1: Math.min(g.c1, x.c1), r1: Math.min(g.r1, x.r1), c2: Math.max(g.c2, x.c2), r2: Math.max(g.r2, x.r2) };
            changed = true;
          }
        });
      }
      return g;
    }
    // フィルターで隠す行
    function indexHidden() {
      hiddenRows = {};
      var f = sh.filter;
      if (!f) return;
      var keys = Object.keys(f.hide || {}).filter(function (c) { return f.hide[c] && f.hide[c].length; });
      if (!keys.length) return;
      var end = Math.max(f.r2 || 0, book.last(si).rows);
      for (var r = f.r + 1; r <= end; r++) {
        for (var i = 0; i < keys.length; i++) {
          var c = +keys[i], txt = show(book.val(si, c, r), styleAt(c, r).nf).text;
          if (f.hide[keys[i]].indexOf(txt) >= 0) { hiddenRows[r] = 1; break; }
        }
      }
    }

    // ---------- 条件付き書式 ----------
    function cfIndex() {
      if (cfCache) return cfCache;
      cfCache = sh.cf.map(function (rule) {
        var g = parseRange(rule.ref), x = { rule: rule, g: g, min: Infinity, max: -Infinity, count: {} };
        if (!g) return x;
        for (var r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) for (var c = g.c1; c <= g.c2; c++) {
          var v = book.val(si, c, r);
          if (typeof v === "number") { if (v < x.min) x.min = v; if (v > x.max) x.max = v; }
          if (v !== null && v !== "") { var key = String(v).toLowerCase(); x.count[key] = (x.count[key] || 0) + 1; }
        }
        return x;
      });
      return cfCache;
    }
    function mix(a, b, k) {
      var pa = [1, 3, 5].map(function (i) { return parseInt(a.substr(i, 2), 16); }), pb = [1, 3, 5].map(function (i) { return parseInt(b.substr(i, 2), 16); });
      return "#" + pa.map(function (x, i) { return ("0" + Math.round(x + (pb[i] - x) * k).toString(16)).slice(-2); }).join("");
    }
    function cfStyle(c, r, v) {
      if (!sh.cf.length) return "";
      var css = "";
      cfIndex().forEach(function (x) {
        if (!x.g || !inRange(x.g, c, r)) return;
        var rule = x.rule, hit = false, n = typeof v === "number" ? v : null, color = rule.color || "#fde68a";
        if (rule.type === "bar") {
          if (n === null || x.max <= 0) return;
          var pct = Math.max(0, Math.min(100, (n - Math.min(0, x.min)) / (x.max - Math.min(0, x.min)) * 100));
          css += "background-image:linear-gradient(90deg," + (rule.color || "#93c5fd") + " " + pct + "%,transparent " + pct + "%);";
          return;
        }
        if (rule.type === "scale") {
          if (n === null || x.max === x.min) return;
          var k = (n - x.min) / (x.max - x.min);
          css += "background-color:" + (k < 0.5 ? mix("#f8696b", "#ffeb84", k * 2) : mix("#ffeb84", "#63be7b", (k - 0.5) * 2)) + ";color:#1f2328;";
          return;
        }
        var target = F.parseNum(rule.v);
        if (rule.type === "gt") hit = n !== null && target !== null && n > target;
        else if (rule.type === "lt") hit = n !== null && target !== null && n < target;
        else if (rule.type === "eq") hit = target !== null ? n === target : v !== null && String(v).toLowerCase() === String(rule.v).toLowerCase();
        else if (rule.type === "text") hit = v !== null && String(v).toLowerCase().indexOf(String(rule.v).toLowerCase()) >= 0;
        else if (rule.type === "dup") hit = v !== null && v !== "" && x.count[String(v).toLowerCase()] > 1;
        if (hit) css += "background-color:" + color + ";color:" + (rule.ink || "#1f2328") + ";";
      });
      return css;
    }

    // ---------- 表を作る ----------
    function colW(c) { return sh.cw[c] || COLW; }
    function cellParts(c, r) {
      var k = ref(c, r), st = sh.style[k] || EMPTY, v = book.val(si, c, r), d = show(v, st.nf), cls = [];
      if (d.num) cls.push("n");
      if (d.err) cls.push("is-err");
      if (d.bool) cls.push("is-c");
      if (st.b) cls.push("is-b");
      if (st.i) cls.push("is-i");
      if (st.u && st.s) cls.push("is-us"); else if (st.u) cls.push("is-u"); else if (st.s) cls.push("is-s");
      if (st.al) cls.push("is-" + st.al);
      if (st.va) cls.push("va-" + st.va);
      if (st.wrap) cls.push("is-wrap");
      var css = "";
      if (d.color || st.fg) css += "color:" + (d.color || st.fg) + ";";
      if (st.bg) css += "background-color:" + st.bg + ";";
      if (st.ff) css += "font-family:" + O.fontCss(st.ff) + ";";
      if (st.fs) css += "font-size:" + Math.round(st.fs * 4 / 3 * 10) / 10 + "px;";
      if (st.bd) css += "box-shadow:" + bdCss(st.bd) + ";";
      css += cfStyle(c, r, v);
      return { cls: cls.join(" "), css: css, text: d.text };
    }
    function freezeCss(c, r) {
      var top = fzTop[r], left = fzLeft[c];
      if (top === undefined && left === undefined) return "";
      return "position:sticky;" + (top !== undefined ? "top:" + top + "px;" : "") + (left !== undefined ? "left:" + left + "px;" : "") + "z-index:" + (top !== undefined && left !== undefined ? 4 : 3) + ";";
    }
    // 固定する行・列のセルは、class に fz を足して、その場所に貼りつける
    function partsFor(c, r) {
      var p = cellParts(c, r);
      if (fzTop[r] !== undefined || fzLeft[c] !== undefined) { p.cls = (p.cls + " fz").trim(); p.css += freezeCss(c, r); }
      return p;
    }
    function filterBtn(c, r) {
      var f = sh.filter;
      if (!f || r !== f.r || c < f.c1 || c > f.c2) return "";
      var on = f.hide && f.hide[c] && f.hide[c].length;
      return '<button type="button" class="flt' + (on ? " is-on" : "") + '" data-flt="' + c + '" aria-label="' + esc(t("フィルター")) + '">' + icon(on ? "filter_alt" : "arrow_drop_down") + "</button>";
    }
    function cellHtml(c, r) {
      var k = ref(c, r), m = mtop[k], p = partsFor(c, r);
      return "<td data-c=\"" + c + "\" data-r=\"" + r + "\"" + (m ? (m.cs > 1 ? ' colspan="' + m.cs + '"' : "") + (m.rs > 1 ? ' rowspan="' + m.rs + '"' : "") : "") +
        (p.cls ? ' class="' + p.cls + '"' : "") + (p.css ? ' style="' + p.css + '"' : "") + ">" + esc(p.text) + filterBtn(c, r) + "</td>";
    }
    function rowHtml(r) {
      var cls = (r === sh.head ? "head" : r === sh.total ? "total" : "") + (hiddenRows[r] ? " is-hide" : "");
      var h = '<tr data-r="' + r + '"' + (cls.trim() ? ' class="' + cls.trim() + '"' : "") + (sh.rh[r] ? ' style="height:' + sh.rh[r] + 'px"' : "") + ">" +
        '<th data-row="' + r + '"' + (fzTop[r] !== undefined ? ' style="top:' + fzTop[r] + 'px;z-index:6"' : "") + ">" + r + '<span class="row-rs" data-rrs="' + r + '"></span></th>';
      for (var c = 0; c < nCols; c++) { var k = ref(c, r); if (!covered[k]) h += cellHtml(c, r); }
      return h + "</tr>";
    }
    function measureFreeze() {
      // 固定する行・列の位置（表の中で、ふつうに置いたときの位置。そこに貼りつける）
      fzTop = {}; fzLeft = {};
      var fz = sh.freeze, tbl = el("grid");
      if (!fz || !tbl || (!fz.r && !fz.c)) return false;
      for (var r = 1; r <= (fz.r || 0); r++) { var tr = tbl.querySelector('tr[data-r="' + r + '"]'); if (tr) fzTop[r] = tr.offsetTop; }
      var heads = tbl.tHead ? tbl.tHead.rows[0].cells : [];
      for (var c = 0; c < (fz.c || 0); c++) { if (heads[c + 1]) fzLeft[c] = heads[c + 1].offsetLeft; }
      return true;
    }
    function extent() {
      var last = book.last(si), mr = 0, mc = 0;
      sh.charts.forEach(function (ch) { var g = parseRange(ch.ref); if (g) { mr = Math.max(mr, g.r2); mc = Math.max(mc, g.c2); } });
      sh.merges.forEach(function (m) { var g = parseRange(m); if (g) { mr = Math.max(mr, g.r2); mc = Math.max(mc, g.c2); } });
      Object.keys(sh.style).forEach(function (k) { var p = parseRef(k); if (p) { mr = Math.max(mr, p.r); mc = Math.max(mc, p.c); } });
      nRows = Math.max(nRows, 60, last.rows + 40, mr + 20, act.r + 20);
      nCols = Math.max(nCols, 26, last.cols + 6, mc + 4, act.c + 4);
    }
    function drawGrid(keepScroll) {
      var scroller = el("grid-scroll");
      if (!scroller) return;
      var sx = scroller.scrollLeft, sy = scroller.scrollTop;
      book.reset(); cfCache = null;
      indexMerges(); indexHidden();
      extent();
      fzTop = {}; fzLeft = {};
      var tbl = el("grid"), h = '<colgroup><col style="width:' + HEADW + 'px">';
      var total = HEADW;
      for (var c = 0; c < nCols; c++) { h += '<col style="width:' + colW(c) + 'px">'; total += colW(c); }
      h += '</colgroup><thead><tr><th class="corner" data-corner title="' + esc(t("すべて選択")) + '"></th>';
      for (c = 0; c < nCols; c++) h += '<th data-col="' + c + '">' + F.colName(c) + '<span class="col-rs" data-rs="' + c + '"></span></th>';
      h += "</tr></thead><tbody>";
      for (var r = 1; r <= nRows; r++) h += rowHtml(r);
      tbl.innerHTML = h + "</tbody>";
      tbl.style.width = total + "px";
      tbl.classList.toggle("no-grid", sh.grid === false);
      // 固定する行・列は、置いた位置を測ってから、もう一度その位置に貼りつける
      if (measureFreeze()) refreshCells(true);
      if (keepScroll) { scroller.scrollLeft = sx; scroller.scrollTop = sy; }
      drawCharts();
      drawTabs();
      placeSel();
    }
    // 中身だけを作り直す（行・列の数や結合が変わらないとき）
    function refreshCells(all) {
      var tbl = el("grid");
      if (!tbl) return;
      book.reset(); cfCache = null;
      var hid = hiddenRows;
      indexHidden();
      var rowsChanged = Object.keys(hid).join() !== Object.keys(hiddenRows).join();
      tbl.querySelectorAll("tbody tr").forEach(function (tr) {
        var r = +tr.getAttribute("data-r");
        if (rowsChanged) tr.classList.toggle("is-hide", !!hiddenRows[r]);
        if (all && fzTop[r] !== undefined) { var th = tr.firstChild; th.style.top = fzTop[r] + "px"; th.style.zIndex = 6; }
      });
      tbl.querySelectorAll("td[data-c]").forEach(function (td) {
        var c = +td.getAttribute("data-c"), r = +td.getAttribute("data-r"), p = partsFor(c, r), html = esc(p.text) + filterBtn(c, r);
        if (td.innerHTML !== html) td.innerHTML = html;
        if ((td.getAttribute("class") || "") !== p.cls) { if (p.cls) td.setAttribute("class", p.cls); else td.removeAttribute("class"); }
        if ((td.getAttribute("style") || "") !== p.css) { if (p.css) td.setAttribute("style", p.css); else td.removeAttribute("style"); }
      });
    }
    // 下までスクロールしたら、行を足す
    function moreRows(n) {
      var tb = el("grid") && el("grid").tBodies[0];
      if (!tb) return;
      var h = "", from = nRows + 1;
      nRows += n;
      for (var r = from; r <= nRows; r++) h += rowHtml(r);
      tb.insertAdjacentHTML("beforeend", h);
    }
    function moreCols(n) { nCols += n; drawGrid(true); }

    // ---------- 選んだところの枠・入力の欄 ----------
    function tdAt(c, r) {
      var m = mergeAt(c, r);
      if (m) { c = m.c1; r = m.r1; }
      return el("grid").querySelector('td[data-c="' + c + '"][data-r="' + r + '"]');
    }
    function rectOf(g) {
      var wrap = el("grid-wrap");
      g = grow(g);
      var a = tdAt(g.c1, g.r1), b = tdAt(g.c2, g.r2);
      if (!a || !b || !wrap) return null;
      var w = wrap.getBoundingClientRect(), ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect(), z = zoomZ || 1;
      if (!ra.width && !rb.width) return null;
      return { x: (ra.left - w.left) / z, y: (ra.top - w.top) / z, w: (rb.right - ra.left) / z, h: (rb.bottom - ra.top) / z };
    }
    function placeBox(box, g) {
      var rc = g && rectOf(g);
      if (!box) return;
      if (!rc) { box.style.display = "none"; return; }
      box.style.display = "block";
      box.style.left = rc.x + "px"; box.style.top = rc.y + "px";
      box.style.width = rc.w + "px"; box.style.height = rc.h + "px";
    }
    function placeSel() {
      if (!el("grid")) return;
      var g = range();
      placeBox(el("sel-box"), g);
      el("sel-box").classList.toggle("is-one", isOne(grow(g)));
      placeBox(el("clip-box"), clip && clip.si === si && !clip.used ? clip.g : null);
      placeEditor();
      // 見出しの色（選んだ行と列）
      el("grid").querySelectorAll("th.hl").forEach(function (x) { x.classList.remove("hl"); });
      var gg = grow(g);
      for (var c = gg.c1; c <= Math.min(gg.c2, nCols - 1); c++) { var th = el("grid").querySelector('th[data-col="' + c + '"]'); if (th) th.classList.add("hl"); }
      if (gg.r2 - gg.r1 < 500) for (var r = gg.r1; r <= gg.r2; r++) { var rh = el("grid").querySelector('th[data-row="' + r + '"]'); if (rh) rh.classList.add("hl"); }
      el("namebox").value = isOne(gg) ? ref(act.c, act.r) : rangeText(gg);
      if (!editing) el("fx").value = rawAt(act.c, act.r);
      stat();
      O.syncSoon();
    }
    function placeEditor() {
      var ced = el("ced"), m = mergeAt(act.c, act.r), g = m || { c1: act.c, r1: act.r, c2: act.c, r2: act.r }, rc = rectOf(g);
      if (!ced || !rc) return;
      ced.style.left = rc.x + "px"; ced.style.top = rc.y + "px";
      ced.style.height = rc.h + "px";
      ced.style.minWidth = rc.w + "px";
      if (!editing) ced.style.width = rc.w + "px";
      else fitEditor();
      var st = styleAt(act.c, act.r);
      ced.style.fontWeight = st.b ? "700" : "";
      ced.style.fontStyle = st.i ? "italic" : "";
      ced.style.fontFamily = st.ff ? O.fontCss(st.ff) : "";
      ced.style.fontSize = st.fs ? Math.round(st.fs * 4 / 3) + "px" : "";
    }
    function fitEditor() {
      var ced = el("ced");
      ced.style.width = "0px";
      ced.style.width = Math.min(Math.max(ced.scrollWidth + 16, parseFloat(ced.style.minWidth) || 0), 640) + "px";
    }
    function scrollToCell(c, r) {
      var td = tdAt(c, r), sc = el("grid-scroll");
      if (!td || !sc) return;
      var a = td.getBoundingClientRect(), b = sc.getBoundingClientRect();
      var headH = (el("grid").tHead ? el("grid").tHead.getBoundingClientRect().height : 29), headW = HEADW * zoomZ;
      var fzH = 0, fzW = 0;
      if (sh.freeze && sh.freeze.r && r > sh.freeze.r) { var lr = tdAt(0, sh.freeze.r); if (lr) fzH = lr.getBoundingClientRect().bottom - b.top - headH; }
      if (sh.freeze && sh.freeze.c && c >= sh.freeze.c) fzW = (fzLeft[sh.freeze.c - 1] + colW(sh.freeze.c - 1) - HEADW) * zoomZ;
      if (a.top < b.top + headH + fzH) sc.scrollTop -= b.top + headH + fzH - a.top;
      else if (a.bottom > b.top + sc.clientHeight) sc.scrollTop += a.bottom - (b.top + sc.clientHeight);
      if (a.left < b.left + headW + fzW) sc.scrollLeft -= b.left + headW + fzW - a.left;
      else if (a.right > b.left + sc.clientWidth) sc.scrollLeft += a.right - (b.left + sc.clientWidth);
    }
    function select(c, r, extend) {
      c = Math.max(0, c); r = Math.max(1, r);
      if (r > nRows - 5) moreRows(Math.max(50, r - nRows + 50));
      if (c > nCols - 2) { nCols = c + 10; drawGrid(true); }
      anc = { c: c, r: r };
      if (!extend) act = { c: c, r: r };
      if (chartSel) { chartSel = null; drawCharts(); O.sync(); }
      placeSel();
      scrollToCell(c, r);
    }
    function selectRange(g) {
      if (g.r2 > nRows - 5 && g.r2 - g.r1 < 5000) moreRows(g.r2 - nRows + 20);
      act = { c: g.c1, r: g.r1 }; anc = { c: g.c2, r: g.r2 };
      placeSel();
      scrollToCell(g.c1, g.r1);
    }
    function focusGrid() { var ced = el("ced"); if (ced && document.activeElement !== ced) ced.focus({ preventScroll: true }); }

    // ---------- 入力 ----------
    function editEl() { return document.activeElement === el("fx") ? el("fx") : el("ced"); }
    function startEdit(mode, value, fromFx) {
      var ced = el("ced");
      if (readOnly()) { ced.value = ""; O.notify(t("閲覧モードでは書きかえられません")); return; }
      editing = true; editMode = mode; refIns = null;
      ced.classList.add("is-edit");
      ced.removeAttribute("data-office-keys");
      if (value !== undefined) { ced.value = value; el("fx").value = value; }
      placeEditor();
      // 打ちはじめ（日本語の変換中も）は、カーソルを動かさない
      if (!fromFx && value !== undefined) { ced.focus({ preventScroll: true }); var n = ced.value.length; ced.setSelectionRange(n, n); }
      else if (!fromFx && document.activeElement !== ced) ced.focus({ preventScroll: true });
      stat();
    }
    function endEdit(commit) {
      if (!editing) return false;
      var v = el("ced").value, ced = el("ced");
      editing = false; refIns = null;
      ced.classList.remove("is-edit");
      ced.setAttribute("data-office-keys", "");
      ced.value = "";
      if (commit) commitCell(act.c, act.r, v);
      el("fx").value = rawAt(act.c, act.r);
      placeEditor();
      return true;
    }
    function commitCell(c, r, v) {
      if (readOnly()) return;
      v = String(v).replace(/^＝/, "=");
      // 数式の関数の名前は大文字に
      if (v.charAt(0) === "=") v = "=" + F.tokenize(v.slice(1)).map(function (x) { return x.t === "fn" ? x.s.toUpperCase() : x.s; }).join("");
      var k = ref(c, r), old = sh.cells[k] || "";
      if (v === old) return;
      if (v === "") delete sh.cells[k]; else sh.cells[k] = v;
      if (v && !(sh.style[k] && sh.style[k].nf)) { var nf = autoNf(v); if (nf) { sh.style[k] = Object.assign({}, sh.style[k] || {}, { nf: nf }); } }
      changed();
    }
    // 中身を変えたあと（表を作り直す・保存する・元に戻すために覚える）
    function changed(structural) {
      if (structural) drawGrid(true);
      else { refreshCells(); drawCharts(); placeSel(); }
      O.changed();
    }
    function formulaPoint() {
      if (!editing) return false;
      var e = editEl(), v = e.value, p = e.selectionStart;
      if (v.charAt(0) !== "=") return false;
      if (refIns && refIns.el === e && p === refIns.start + refIns.len) return true;
      return /[=(,+\-*/^&<>:;]$/.test(v.slice(0, p).replace(/\s+$/, ""));
    }
    function insertRef(text) {
      var e = editEl(), v = e.value, start, end;
      if (refIns && refIns.el === e) { start = refIns.start; end = refIns.start + refIns.len; }
      else { start = e.selectionStart; end = e.selectionEnd; }
      e.value = v.slice(0, start) + text + v.slice(end);
      refIns = { el: e, start: start, len: text.length };
      var p = start + text.length;
      e.setSelectionRange(p, p);
      if (e === el("fx")) el("ced").value = e.value; else el("fx").value = e.value;
      if (e === el("ced")) fitEditor();
    }

    // ---------- 書式 ----------
    function eachCell(g, fn) {
      g = grow(g);
      var r2 = Math.min(g.r2, Math.max(nRows, book.last(si).rows)), c2 = Math.min(g.c2, Math.max(nCols, book.last(si).cols));
      for (var r = g.r1; r <= r2; r++) for (var c = g.c1; c <= c2; c++) fn(c, r, ref(c, r));
    }
    function setStyle(g, patch) {
      eachCell(g, function (c, r, k) {
        var st = Object.assign({}, sh.style[k] || {});
        Object.keys(patch).forEach(function (p) { if (patch[p] === null || patch[p] === false || patch[p] === "") delete st[p]; else st[p] = patch[p]; });
        if (Object.keys(st).length) sh.style[k] = st; else delete sh.style[k];
      });
    }
    function toggleStyle(name) {
      if (readOnly()) return;
      var on = !styleAt(act.c, act.r)[name], p = {};
      p[name] = on ? 1 : null;
      setStyle(range(), p);
      changed();
    }
    function borders(kind) {
      var g = grow(range());
      eachCell(g, function (c, r, k) {
        var st = Object.assign({}, sh.style[k] || {}), bd = Object.assign({}, st.bd || {});
        if (kind === "none") bd = {};
        else if (kind === "all") bd = { t: 1, r: 1, b: 1, l: 1 };
        else if (kind === "out" || kind === "thick") {
          var w = kind === "thick" ? 2 : 1;
          if (r === g.r1) bd.t = w; if (r === g.r2) bd.b = w; if (c === g.c1) bd.l = w; if (c === g.c2) bd.r = w;
        }
        else if (kind === "bottom" && r === g.r2) bd.b = 1;
        else if (kind === "double" && r === g.r2) bd.b = 2;
        else if (kind === "top" && r === g.r1) bd.t = 1;
        else if (kind === "left" && c === g.c1) bd.l = 1;
        else if (kind === "right" && c === g.c2) bd.r = 1;
        if (Object.keys(bd).length) st.bd = bd; else delete st.bd;
        if (Object.keys(st).length) sh.style[k] = st; else delete sh.style[k];
      });
      changed();
    }
    function decimals(nf, d) {
      if (nf && F.isDateFormat(nf)) return nf;
      if (!nf) {
        var v = book.val(si, act.c, act.r), dec = typeof v === "number" ? ((String(v).split(".")[1] || "").length) : 0;
        nf = dec ? "0." + new Array(dec + 1).join("0") : "0";
      }
      var m = /\.0+/.exec(nf);
      if (d > 0) return m ? nf.replace(/\.0+/, m[0] + "0") : nf.replace(/([0#])(?![0#,.])/, "$1.0");
      if (!m) return nf;
      return m[0].length > 2 ? nf.replace(/\.0+/, m[0].slice(0, -1)) : nf.replace(/\.0+/, "");
    }

    // ---------- 行と列の挿入・削除（数式の参照も直す） ----------
    function shiftSheet(axis, at, n) {
      var name = sh.name;
      data.sheets.forEach(function (o) {
        Object.keys(o.cells).forEach(function (k) {
          var x = o.cells[k];
          if (x.charAt(0) === "=") o.cells[k] = "=" + F.adjust(x.slice(1), { sheet: name, here: o.name, axis: axis, at: at, n: n });
        });
      });
      function pos(p) { return axis === "r" ? p.r : p.c; }
      function move(map) {
        var out = {};
        Object.keys(map).forEach(function (k) {
          var p = parseRef(k);
          if (!p) return;
          var v = pos(p);
          if (n < 0 && v >= at && v < at - n) return;
          if (v >= at) { if (axis === "r") p.r += n; else p.c += n; }
          if (p.r >= 1 && p.c >= 0) out[ref(p.c, p.r)] = map[k];
        });
        return out;
      }
      sh.cells = move(sh.cells);
      sh.style = move(sh.style);
      function moveNum(map) {
        var out = {};
        Object.keys(map).forEach(function (k) { var v = +k; if (n < 0 && v >= at && v < at - n) return; out[v >= at ? v + n : v] = map[k]; });
        return out;
      }
      if (axis === "c") sh.cw = moveNum(sh.cw); else sh.rh = moveNum(sh.rh);
      function fixRange(s) {
        var x = "=" + s, y = F.adjust(x.slice(1), { sheet: name, here: name, axis: axis, at: at, n: n });
        return /#REF!/.test(y) ? null : y;
      }
      sh.merges = sh.merges.map(fixRange).filter(Boolean);
      sh.cf = sh.cf.map(function (rule) { var r2 = fixRange(rule.ref); if (!r2) return null; rule.ref = r2; return rule; }).filter(Boolean);
      sh.charts.forEach(function (ch) { var r2 = fixRange(ch.ref); if (r2) ch.ref = r2; });
      if (axis === "r") {
        ["head", "total"].forEach(function (k) { if (sh[k] >= at) { if (n < 0 && sh[k] < at - n) delete sh[k]; else sh[k] += n; } });
        if (sh.filter && sh.filter.r >= at) sh.filter.r = Math.max(1, sh.filter.r + n);
      } else if (sh.filter) {
        if (sh.filter.c1 >= at) sh.filter.c1 = Math.max(0, sh.filter.c1 + n);
        if (sh.filter.c2 >= at) sh.filter.c2 = Math.max(sh.filter.c1, sh.filter.c2 + n);
        sh.filter.hide = {};
      }
    }
    function insertRows(where) {
      var g = grow(range()), n = g.r2 - g.r1 + 1;
      shiftSheet("r", where === "below" ? g.r2 + 1 : g.r1, n);
      if (where === "below") { act = { c: g.c1, r: g.r2 + 1 }; anc = { c: g.c2, r: g.r2 + n }; }
      changed(true);
    }
    function insertCols(where) {
      var g = grow(range()), n = g.c2 - g.c1 + 1;
      shiftSheet("c", where === "right" ? g.c2 + 1 : g.c1, n);
      if (where === "right") { act = { c: g.c2 + 1, r: g.r1 }; anc = { c: g.c2 + n, r: g.r2 }; }
      changed(true);
    }
    function deleteRows() { var g = grow(range()); shiftSheet("r", g.r1, -(g.r2 - g.r1 + 1)); act = anc = { c: g.c1, r: g.r1 }; changed(true); }
    function deleteCols() { var g = grow(range()); shiftSheet("c", g.c1, -(g.c2 - g.c1 + 1)); act = anc = { c: g.c1, r: g.r1 }; changed(true); }

    // ---------- まとまった表（選んだセルのまわりの、空でないところ） ----------
    function filled(c, r) { var x = sh.cells[ref(c, r)]; return x !== undefined && x !== ""; }
    function region(c, r) {
      var g = { c1: c, r1: r, c2: c, r2: r }, again = true, guard = 0;
      while (again && guard++ < 400) {
        again = false;
        var i;
        for (i = g.c1 - 1; i <= g.c2 + 1; i++) { if (g.r1 > 1 && i >= 0 && filled(i, g.r1 - 1)) { g.r1--; again = true; break; } }
        for (i = g.c1 - 1; i <= g.c2 + 1; i++) { if (i >= 0 && filled(i, g.r2 + 1)) { g.r2++; again = true; break; } }
        for (i = g.r1 - 1; i <= g.r2 + 1; i++) { if (g.c1 > 0 && i >= 1 && filled(g.c1 - 1, i)) { g.c1--; again = true; break; } }
        for (i = g.r1 - 1; i <= g.r2 + 1; i++) { if (i >= 1 && filled(g.c2 + 1, i)) { g.c2++; again = true; break; } }
      }
      return g;
    }
    function workRange() { var g = range(); return isOne(g) ? region(act.c, act.r) : g; }
    function hasHeader(g) {
      if (sh.head === g.r1 || (sh.filter && sh.filter.r === g.r1)) return true;
      if (g.r2 <= g.r1) return false;
      var topText = true, belowNum = false;
      for (var c = g.c1; c <= g.c2; c++) {
        var v = book.val(si, c, g.r1);
        if (v !== null && typeof v !== "string") topText = false;
        if (typeof book.val(si, c, g.r1 + 1) === "number") belowNum = true;
      }
      return topText && belowNum;
    }

    // ---------- 並べ替え・フィルター・重複の削除 ----------
    function sortRange(desc, keyCol) {
      var g = workRange();
      if (g.r2 <= g.r1) { O.notify(t("並べ替えるデータがありません")); return; }
      var start = hasHeader(g) ? g.r1 + 1 : g.r1, end = g.r2;
      if (sh.total && sh.total > start && sh.total <= end) end = sh.total - 1;
      var key = keyCol === undefined ? (act.c >= g.c1 && act.c <= g.c2 ? act.c : g.c1) : keyCol, rows = [];
      for (var r = start; r <= end; r++) {
        var row = { r: r, key: book.val(si, key, r), cells: {}, styles: {} };
        for (var c = g.c1; c <= g.c2; c++) { var k = ref(c, r); if (k in sh.cells) row.cells[c] = sh.cells[k]; if (k in sh.style) row.styles[c] = sh.style[k]; }
        rows.push(row);
      }
      rows.sort(function (a, b) {
        var ea = a.key === null || a.key === "", eb = b.key === null || b.key === "";
        if (ea || eb) return ea === eb ? a.r - b.r : ea ? 1 : -1;
        var x = F.compare(F.isErr(a.key) ? a.key.err : a.key, F.isErr(b.key) ? b.key.err : b.key);
        return (desc ? -x : x) || a.r - b.r;
      });
      rows.forEach(function (row, i) {
        var R = start + i;
        for (var c = g.c1; c <= g.c2; c++) {
          var k = ref(c, R), raw = row.cells[c];
          if (raw === undefined) delete sh.cells[k];
          else sh.cells[k] = raw.charAt(0) === "=" ? "=" + F.shift(raw.slice(1), 0, R - row.r) : raw;
          if (row.styles[c]) sh.style[k] = row.styles[c]; else delete sh.style[k];
        }
      });
      changed();
      O.notify(desc ? t("降順に並べ替えました") : t("昇順に並べ替えました"));
    }
    function toggleFilter() {
      if (sh.filter) { delete sh.filter; changed(true); O.notify(t("フィルターをやめました")); return; }
      var g = workRange();
      sh.filter = { r: g.r1, c1: g.c1, c2: g.c2, r2: g.r2, hide: {} };
      changed(true);
      O.notify(t("見出しの ▼ を押すと、出す行を選べます"));
    }
    function filterDialog(c) {
      var f = sh.filter, end = Math.max(f.r2 || 0, book.last(si).rows), seen = {}, list = [];
      for (var r = f.r + 1; r <= end; r++) { var txt = show(book.val(si, c, r), styleAt(c, r).nf).text; if (!seen[txt]) { seen[txt] = 1; list.push(txt); } }
      list.sort(function (a, b) { var x = F.parseNum(a), y = F.parseNum(b); return x !== null && y !== null ? x - y : a.localeCompare(b, "ja"); });
      var hidden = (f.hide && f.hide[c]) || [], body = document.createElement("div");
      body.className = "flt-dlg";
      body.innerHTML = '<div class="btn-row"><button type="button" class="m3-btn m3-btn--tonal m3-state" data-sort="asc">' + icon("arrow_upward") + esc(t("昇順")) + '</button><button type="button" class="m3-btn m3-btn--tonal m3-state" data-sort="desc">' + icon("arrow_downward") + esc(t("降順")) + "</button></div>" +
        '<label class="flt-dlg__all"><input type="checkbox" data-all' + (hidden.length ? "" : " checked") + "> " + esc(t("すべて選択")) + "</label>" +
        '<div class="flt-dlg__list">' + list.map(function (x, i) { return '<label><input type="checkbox" data-i="' + i + '"' + (hidden.indexOf(x) >= 0 ? "" : " checked") + "> " + esc(x === "" ? t("（空白）") : x) + "</label>"; }).join("") + "</div>";
      body.addEventListener("change", function (e) {
        if (e.target.hasAttribute("data-all")) body.querySelectorAll("[data-i]").forEach(function (x) { x.checked = e.target.checked; });
      });
      M3.dialog({ title: t("{col} 列のフィルター", { col: F.colName(c) }), icon: "filter_alt", body: body,
        actions: [{ label: t("キャンセル"), value: null }, { label: t("OK"), primary: true, value: function () {
          var hide = [];
          body.querySelectorAll("[data-i]").forEach(function (x) { if (!x.checked) hide.push(list[+x.getAttribute("data-i")]); });
          return { hide: hide };
        } }],
        onReady: function (close) {
          body.querySelectorAll("[data-sort]").forEach(function (b) { b.addEventListener("click", function () { close({ sort: b.getAttribute("data-sort") }); }); });
        }
      }).then(function (res) {
        if (!res) return;
        if (res.sort) { anc = { c: f.c1, r: f.r }; act = { c: f.c2, r: Math.max(f.r2 || 0, book.last(si).rows) }; sortRange(res.sort === "desc", c); act = anc = { c: c, r: f.r }; placeSel(); return; }
        f.hide = f.hide || {};
        if (res.hide.length) f.hide[c] = res.hide; else delete f.hide[c];
        changed(true);
      });
    }
    function dedupe() {
      var g = workRange(), start = hasHeader(g) ? g.r1 + 1 : g.r1, seen = {}, keep = [], removed = 0;
      for (var r = start; r <= g.r2; r++) {
        var key = [];
        for (var c = g.c1; c <= g.c2; c++) key.push(show(book.val(si, c, r)).text);
        key = key.join("\u0001");
        if (seen[key]) { removed++; continue; }
        seen[key] = 1;
        keep.push(r);
      }
      if (!removed) { O.notify(t("重複する行はありませんでした")); return; }
      var rows = keep.map(function (r) { var o = {}; for (var c = g.c1; c <= g.c2; c++) o[c] = [sh.cells[ref(c, r)], sh.style[ref(c, r)], r]; return o; });
      for (var r2 = start; r2 <= g.r2; r2++) for (var c2 = g.c1; c2 <= g.c2; c2++) { delete sh.cells[ref(c2, r2)]; delete sh.style[ref(c2, r2)]; }
      rows.forEach(function (o, i) {
        for (var c = g.c1; c <= g.c2; c++) {
          var k = ref(c, start + i), x = o[c];
          if (x[0] !== undefined) sh.cells[k] = x[0].charAt(0) === "=" ? "=" + F.shift(x[0].slice(1), 0, start + i - x[2]) : x[0];
          if (x[1]) sh.style[k] = x[1];
        }
      });
      changed();
      O.notify(t("重複する {n} 行を削除しました", { n: removed }));
    }

    // ---------- コピー・貼り付け・オートフィル ----------
    function clipData(g, cut) {
      g = grow(g);
      var rows = [], text = [], html = [];
      for (var r = g.r1; r <= g.r2; r++) {
        if (hiddenRows[r]) continue;
        var row = [], tr = [], tdh = [];
        for (var c = g.c1; c <= g.c2; c++) {
          var k = ref(c, r), p = show(book.val(si, c, r), styleAt(c, r).nf).text;
          row.push({ raw: sh.cells[k], st: sh.style[k], c: c, r: r });
          tr.push(/[\t\n"]/.test(p) ? '"' + p.replace(/"/g, '""') + '"' : p);
          tdh.push("<td>" + esc(p) + "</td>");
        }
        rows.push(row); text.push(tr.join("\t")); html.push("<tr>" + tdh.join("") + "</tr>");
      }
      return { si: si, g: g, rows: rows, cut: !!cut, text: text.join("\r\n"), html: '<table><tbody>' + html.join("") + "</tbody></table>" };
    }
    function parseDelimited(text, sep) {
      var rows = [], row = [], cur = "", q = false, i = 0;
      text = text.replace(/^﻿/, "");
      while (i < text.length) {
        var ch = text.charAt(i);
        if (q) {
          if (ch === '"') { if (text.charAt(i + 1) === '"') { cur += '"'; i += 2; continue; } q = false; i++; continue; }
          cur += ch; i++; continue;
        }
        if (ch === '"' && cur === "") { q = true; i++; continue; }
        if (ch === sep) { row.push(cur); cur = ""; i++; continue; }
        if (ch === "\r" || ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; i += ch === "\r" && text.charAt(i + 1) === "\n" ? 2 : 1; continue; }
        cur += ch; i++;
      }
      if (cur !== "" || row.length) { row.push(cur); rows.push(row); }
      return rows;
    }
    function pasteRows(rows) {
      if (readOnly() || !rows.length) return;
      var c0 = act.c, r0 = act.r;
      rows.forEach(function (row, i) { row.forEach(function (v, j) { var k = ref(c0 + j, r0 + i); if (v === "") delete sh.cells[k]; else sh.cells[k] = v; }); });
      act = { c: c0, r: r0 }; anc = { c: c0 + Math.max.apply(null, rows.map(function (x) { return x.length; })) - 1, r: r0 + rows.length - 1 };
      changed(true);
    }
    function pasteClip(cb, onlyValues) {
      if (readOnly()) return;
      var src = cb.rows, h = src.length, w = h ? src[0].length : 0;
      if (!h) return;
      var g = range(), H = Math.max(h, g.r2 - g.r1 + 1), W = Math.max(w, g.c2 - g.c1 + 1);
      if (H % h) H = h;
      if (W % w) W = w;
      var target = data.sheets[cb.si];
      // 切り取りのときは、先に元のセルを消す（同じところに重なってもよいように）
      if (cb.cut) src.forEach(function (row) { row.forEach(function (x) { var k = ref(x.c, x.r); delete target.cells[k]; delete target.style[k]; }); });
      for (var i = 0; i < H; i++) for (var j = 0; j < W; j++) {
        var x = src[i % h][j % w], c = g.c1 + j, r = g.r1 + i, k = ref(c, r), raw = x.raw;
        if (raw !== undefined && raw.charAt(0) === "=" && !cb.cut) raw = onlyValues ? valueRaw(cb.si, x.c, x.r) : "=" + F.shift(raw.slice(1), c - x.c, r - x.r);
        if (raw === undefined || raw === "") delete sh.cells[k]; else sh.cells[k] = raw;
        if (!onlyValues) { if (x.st) sh.style[k] = x.st; else delete sh.style[k]; }
      }
      act = { c: g.c1, r: g.r1 }; anc = { c: g.c1 + W - 1, r: g.r1 + H - 1 };
      if (cb.cut) clip = null;
      changed(true);
    }
    function valueRaw(s, c, r) {
      var v = book.val(s, c, r);
      if (v === null) return undefined;
      if (F.isErr(v)) return v.err;
      if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
      if (typeof v === "number") return F.str(v);
      return /^[=+\-']/.test(v) || F.parseNum(v) !== null ? "'" + v : v;
    }
    function onCopy(e, cut) {
      if (editing) return;
      e.preventDefault();
      clip = clipData(range(), cut);
      e.clipboardData.setData("text/plain", clip.text);
      e.clipboardData.setData("text/html", clip.html);
      placeSel();
      O.notify(cut ? t("切り取りました。貼り付けるところを選んで Ctrl+V を押してください") : t("コピーしました"));
    }
    function onPaste(e) {
      if (editing) return;
      e.preventDefault();
      var text = e.clipboardData.getData("text/plain"), html = e.clipboardData.getData("text/html");
      pasteText(text, html);
    }
    function pasteText(text, html) {
      if (clip && text.replace(/\r\n/g, "\n") === clip.text.replace(/\r\n/g, "\n")) return pasteClip(clip);
      if (html && /<table/i.test(html)) {
        var doc = new DOMParser().parseFromString(html, "text/html"), rows = [];
        doc.querySelectorAll("tr").forEach(function (tr) { var row = []; tr.querySelectorAll("td,th").forEach(function (td) { row.push(td.textContent.replace(/ /g, " ").trim()); }); rows.push(row); });
        if (rows.length) return pasteRows(rows);
      }
      if (text) pasteRows(parseDelimited(text, text.indexOf("\t") >= 0 ? "\t" : "\u0000"));
    }
    // オートフィル（数のならび・「第 1 回」のような文字・日付は続きにする。数式は参照をずらす）
    function fillFrom(src, dst) {
      var down = dst.r1 !== src.r1 || dst.r2 !== src.r2, axis = down ? "r" : "c";
      var len = axis === "r" ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      var lines = axis === "r" ? [src.c1, src.c2] : [src.r1, src.r2];
      for (var ln = lines[0]; ln <= lines[1]; ln++) {
        var seq = [];
        for (var i = 0; i < len; i++) { var c = axis === "r" ? ln : src.c1 + i, r = axis === "r" ? src.r1 + i : ln; seq.push({ raw: sh.cells[ref(c, r)], st: sh.style[ref(c, r)], c: c, r: r }); }
        var nums = seq.map(function (x) { return x.raw !== undefined && x.raw.charAt(0) !== "=" ? F.parseNum(x.raw) : null; });
        var allNum = nums.every(function (n) { return n !== null; }), step = allNum && len > 1 ? (nums[len - 1] - nums[0]) / (len - 1) : 0;
        var isDate = allNum && seq[0].raw && F.parseDate(seq[0].raw) !== null && /[\/年-]/.test(seq[0].raw);
        if (isDate && len === 1) step = 1;
        var from = axis === "r" ? dst.r1 : dst.c1, to = axis === "r" ? dst.r2 : dst.c2, base = axis === "r" ? src.r1 : src.c1;
        for (var p = from; p <= to; p++) {
          if (p >= base && p < base + len) continue;
          var off = p - base, idx = ((off % len) + len) % len, x = seq[idx], raw = x.raw;
          var tc = axis === "r" ? ln : p, tr = axis === "r" ? p : ln, k = ref(tc, tr);
          if (raw !== undefined && raw.charAt(0) === "=") raw = "=" + F.shift(raw.slice(1), tc - x.c, tr - x.r);
          else if (allNum && (len > 1 || isDate)) {
            var v = nums[0] + step * off;
            raw = isDate ? F.format(v, "yyyy/m/d").text : F.str(Math.round(v * 1e10) / 1e10);
          } else if (raw !== undefined && len === 1 && /\d/.test(raw) && F.parseNum(raw) === null && raw.charAt(0) !== "=") {
            // 「第 1 回」「No.3」のような文字は、いちばん後ろの数を増やす
            var m = /^(.*?)(\d+)(\D*)$/.exec(raw);
            if (m) raw = m[1] + Math.max(0, +m[2] + off) + m[3];
          }
          if (raw === undefined) delete sh.cells[k]; else sh.cells[k] = raw;
          if (x.st) sh.style[k] = x.st; else delete sh.style[k];
        }
      }
      act = { c: dst.c1, r: dst.r1 }; anc = { c: dst.c2, r: dst.r2 };
      changed();
    }
    function fillDir(dir) {
      var g = grow(range());
      if (dir === "down") { if (g.r1 === g.r2) { if (g.r1 < 2) return; g.r1--; } fillFrom({ c1: g.c1, r1: g.r1, c2: g.c2, r2: g.r1 }, g); }
      else if (dir === "right") { if (g.c1 === g.c2) { if (g.c1 < 1) return; g.c1--; } fillFrom({ c1: g.c1, r1: g.r1, c2: g.c1, r2: g.r2 }, g); }
      else if (dir === "up") { if (g.r1 === g.r2) g.r2++; fillFrom({ c1: g.c1, r1: g.r2, c2: g.c2, r2: g.r2 }, g); }
      else if (dir === "left") { if (g.c1 === g.c2) g.c2++; fillFrom({ c1: g.c2, r1: g.r1, c2: g.c2, r2: g.r2 }, g); }
    }
    function clearRange(what) {
      if (readOnly()) return;
      eachCell(range(), function (c, r, k) {
        if (what !== "format") delete sh.cells[k];
        if (what !== "content") delete sh.style[k];
      });
      changed();
    }

    // ---------- オート SUM・関数 ----------
    function autoFn(fn) {
      if (readOnly()) return;
      var g = range();
      if (!isOne(g)) {
        // 範囲を選んでいるときは、その下に列ごとの計算を入れる
        var r = g.r2 + 1;
        for (var c = g.c1; c <= g.c2; c++) commitQuiet(c, r, "=" + fn + "(" + ref(c, g.r1) + ":" + ref(c, g.r2) + ")");
        act = anc = { c: g.c1, r: r };
        changed();
        return;
      }
      var top = act.r - 1, col = act.c;
      while (top >= 1 && typeof book.val(si, col, top) === "number") top--;
      if (top < act.r - 1) { commitCell(col, act.r, "=" + fn + "(" + ref(col, top + 1) + ":" + ref(col, act.r - 1) + ")"); placeSel(); return; }
      var left = col - 1;
      while (left >= 0 && typeof book.val(si, left, act.r) === "number") left--;
      if (left < col - 1) { commitCell(col, act.r, "=" + fn + "(" + ref(left + 1, act.r) + ":" + ref(col - 1, act.r) + ")"); placeSel(); return; }
      startEdit("enter", "=" + fn + "(");
    }
    function commitQuiet(c, r, v) { var k = ref(c, r); if (v === "") delete sh.cells[k]; else sh.cells[k] = v; }
    var FN_CATS = [
      { id: "logic", icon: "rule", label: t("論理"), list: "IF IFS AND OR NOT XOR IFERROR IFNA SWITCH" },
      { id: "text", icon: "text_fields", label: t("文字列操作"), list: "CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN FIND SEARCH SUBSTITUTE REPLACE TRIM UPPER LOWER PROPER TEXT VALUE REPT EXACT CHAR CODE ASC JIS" },
      { id: "date", icon: "calendar_month", label: t("日付/時刻"), list: "TODAY NOW DATE YEAR MONTH DAY HOUR MINUTE SECOND TIME WEEKDAY WEEKNUM DATEDIF DAYS EDATE EOMONTH NETWORKDAYS WORKDAY DATEVALUE" },
      { id: "lookup", icon: "manage_search", label: t("検索/行列"), list: "VLOOKUP HLOOKUP XLOOKUP LOOKUP INDEX MATCH XMATCH CHOOSE OFFSET INDIRECT ROW ROWS COLUMN COLUMNS FILTER SORT UNIQUE TRANSPOSE" },
      { id: "math", icon: "calculate", label: t("数学/三角"), list: "SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SQRT POWER PI RAND RANDBETWEEN CEILING FLOOR SIGN EXP LN LOG LOG10 SIN COS TAN" },
      { id: "stat", icon: "query_stats", label: t("統計"), list: "AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXIFS MIN MINIFS MEDIAN MODE RANK LARGE SMALL STDEV STDEV.P STDEV.S VAR VAR.P PERCENTILE QUARTILE" }
    ];
    function fnList(cat) { return cat.list.split(" ").filter(function (n) { return F.names.indexOf(n) >= 0; }); }
    function insertFn(name) {
      if (editing && el("ced").value.charAt(0) === "=") {
        var e = editEl(), p = e.selectionStart, v = e.value;
        e.value = v.slice(0, p) + name + "(" + v.slice(e.selectionEnd);
        e.setSelectionRange(p + name.length + 1, p + name.length + 1);
        if (e === el("ced")) { el("fx").value = e.value; fitEditor(); } else el("ced").value = e.value;
        e.focus();
        return;
      }
      startEdit("edit", "=" + name + "(");
    }
    function fnMenu(at, cat) {
      if (cat) {
        O.menu(at, fnList(cat).map(function (n) { return { value: n, label: n }; })).then(function (v) { if (v) insertFn(v); });
        return;
      }
      O.menu(at, FN_CATS.map(function (c) { return { value: c.id, label: c.label, icon: c.icon, sub: "›" }; }).concat([{ sep: 1 }, { value: "all", label: t("すべての関数…"), icon: "functions" }])).then(function (v) {
        if (!v) return;
        if (v === "all") {
          M3.choose({ title: t("関数の挿入"), icon: "functions", options: F.names.map(function (n) { return { value: n, label: n }; }) }).then(function (n) { if (n) insertFn(n); });
          return;
        }
        fnMenu(at, FN_CATS.filter(function (c) { return c.id === v; })[0]);
      });
    }

    // ---------- 結合・固定・幅 ----------
    function toggleMerge() {
      if (readOnly()) return;
      var g = grow(range()), key = rangeText(g), i = sh.merges.indexOf(key);
      if (i >= 0) { sh.merges.splice(i, 1); changed(true); return; }
      if (isOne(g)) { O.notify(t("結合するセルを 2 つ以上選んでください")); return; }
      sh.merges = sh.merges.filter(function (m) { var x = parseRange(m); return x && (x.c2 < g.c1 || x.c1 > g.c2 || x.r2 < g.r1 || x.r1 > g.r2); });
      sh.merges.push(key);
      var k = ref(g.c1, g.r1);
      sh.style[k] = Object.assign({}, sh.style[k] || {}, { al: "c" });
      act = anc = { c: g.c1, r: g.r1 };
      changed(true);
    }
    function freeze(kind) {
      if (kind === "none") delete sh.freeze;
      else if (kind === "row") sh.freeze = { r: 1, c: 0 };
      else if (kind === "col") sh.freeze = { r: 0, c: 1 };
      else sh.freeze = { r: Math.max(0, act.r - 1), c: Math.max(0, act.c) };
      if (sh.freeze && !sh.freeze.r && !sh.freeze.c) delete sh.freeze;
      changed(true);
    }
    var measureCv = null;
    function textWidth(s, st) {
      measureCv = measureCv || document.createElement("canvas");
      var g = measureCv.getContext("2d");
      g.font = (st.b ? "700 " : "400 ") + (st.fs ? st.fs * 4 / 3 : 14) + "px " + (st.ff ? O.fontCss(st.ff) : '"M PLUS 1", sans-serif');
      return g.measureText(s).width;
    }
    function autoWidth(c) {
      var w = 40;
      Object.keys(sh.cells).forEach(function (k) {
        var p = parseRef(k);
        if (!p || p.c !== c || mergeAt(p.c, p.r)) return;
        w = Math.max(w, textWidth(show(book.val(si, p.c, p.r), styleAt(p.c, p.r).nf).text, styleAt(p.c, p.r)) + 22);
      });
      sh.cw[c] = Math.min(600, Math.ceil(w));
    }
    function askWidth(rows) {
      var g = grow(range()), cur = rows ? (sh.rh[g.r1] || 28) : colW(g.c1);
      M3.prompt({ title: rows ? t("行の高さ") : t("列の幅"), label: t("ピクセル"), value: String(cur), ok: t("OK") }).then(function (v) {
        var n = Math.round(parseFloat(v));
        if (!(n >= 8 && n <= 1000)) return;
        for (var i = rows ? g.r1 : g.c1; i <= (rows ? g.r2 : g.c2); i++) { if (rows) sh.rh[i] = n; else sh.cw[i] = n; }
        changed(true);
      });
    }

    // ---------- シート ----------
    function uniqueName(base) {
      var n = 1, name;
      do { name = base + n; n++; } while (book.sheetIndex(name) >= 0);
      return name;
    }
    function switchSheet(i) {
      if (i < 0 || i >= data.sheets.length || i === si) return;
      endEdit(true);
      data.cur = si = i;
      sh = data.sheets[si];
      act = anc = { c: 0, r: 1 };
      nRows = 100; nCols = 26;
      chartSel = null;
      el("grid-scroll").scrollTop = 0; el("grid-scroll").scrollLeft = 0;
      drawGrid();
      O.sync();
      O.changed();
    }
    function addSheet() {
      data.sheets.splice(si + 1, 0, fixSheet({ name: uniqueName("Sheet") }, 0));
      O.changed();
      switchSheet(si + 1);
    }
    function renameSheet(i) {
      var s = data.sheets[i];
      M3.prompt({ title: t("シートの名前を変更"), label: t("シートの名前"), value: s.name, ok: t("変更") }).then(function (v) {
        v = (v || "").trim().replace(/[\\\/?*\[\]:]/g, "").slice(0, 31);
        if (!v || v === s.name) return;
        if (book.sheetIndex(v) >= 0) { O.notify(t("同じ名前のシートがあります")); return; }
        var from = s.name;
        data.sheets.forEach(function (o) { Object.keys(o.cells).forEach(function (k) { var x = o.cells[k]; if (x.charAt(0) === "=") o.cells[k] = "=" + F.renameSheet(x.slice(1), from, v); }); });
        s.name = v;
        changed(true);
      });
    }
    function deleteSheet(i) {
      if (data.sheets.length < 2) { O.notify(t("シートが 1 枚のときは削除できません")); return; }
      M3.confirm({ title: t("「{name}」を削除しますか？", { name: data.sheets[i].name }), text: t("このシートの中身が消えます。元に戻すで戻せます。"), ok: t("削除"), danger: true, icon: "delete" }).then(function (ok) {
        if (!ok) return;
        data.sheets.splice(i, 1);
        data.cur = si = Math.min(si, data.sheets.length - 1);
        sh = data.sheets[si];
        act = anc = { c: 0, r: 1 };
        changed(true);
      });
    }
    function moveSheet(i, d) {
      var j = i + d;
      if (j < 0 || j >= data.sheets.length) return;
      var s = data.sheets.splice(i, 1)[0];
      data.sheets.splice(j, 0, s);
      data.cur = si = j;
      sh = data.sheets[si];
      changed(true);
    }
    function dupSheet(i) {
      var s = O.copy(data.sheets[i]), base = data.sheets[i].name.replace(/\s*\(\d+\)$/, ""), n = 2;
      while (book.sheetIndex(base + " (" + n + ")") >= 0) n++;
      s.name = base + " (" + n + ")";
      data.sheets.splice(i + 1, 0, s);
      O.changed();
      switchSheet(i + 1);
    }
    function sheetMenu(i, at) {
      O.menu(at, [
        { value: "rename", label: t("名前の変更"), icon: "edit" },
        { value: "dup", label: t("コピーを作る"), icon: "content_copy" },
        { value: "left", label: t("左へ移動"), icon: "arrow_back", disabled: i === 0 },
        { value: "right", label: t("右へ移動"), icon: "arrow_forward", disabled: i === data.sheets.length - 1 },
        { sep: 1 },
        { value: "add", label: t("新しいシート"), icon: "add" },
        { value: "del", label: t("削除"), icon: "delete", disabled: data.sheets.length < 2 }
      ]).then(function (v) {
        if (v === "rename") renameSheet(i);
        else if (v === "dup") dupSheet(i);
        else if (v === "left") moveSheet(i, -1);
        else if (v === "right") moveSheet(i, 1);
        else if (v === "add") addSheet();
        else if (v === "del") deleteSheet(i);
      });
    }
    function drawTabs() {
      var box = el("sheet-tabs");
      if (!box) return;
      box.innerHTML = '<span class="sheet-tabs__nav hide-sm"><button type="button" class="icon-btn m3-state" data-cmd="sheetPrev" aria-label="' + esc(t("前のシート")) + '">' + icon("chevron_left") + '</button><button type="button" class="icon-btn m3-state" data-cmd="sheetNext" aria-label="' + esc(t("次のシート")) + '">' + icon("chevron_right") + "</button></span>" +
        '<span class="sheet-tabs__list" role="tablist">' + data.sheets.map(function (s, i) {
          return '<button type="button" class="stab m3-state' + (i === si ? " is-on" : "") + '" role="tab" aria-selected="' + (i === si) + '" data-sheet="' + i + '">' + esc(s.name) + "</button>";
        }).join("") + "</span>" +
        '<button type="button" class="icon-btn m3-state" data-cmd="sheetAdd" aria-label="' + esc(t("新しいシート")) + '" title="' + esc(t("新しいシート")) + '">' + icon("add_circle") + "</button>";
    }

    // ---------- グラフ ----------
    function chartRows(ch) {
      var g = parseRange(ch.ref), rows = [];
      if (!g) return rows;
      for (var r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) {
        if (hiddenRows[r]) continue;
        var row = [];
        for (var c = g.c1; c <= Math.min(g.c2, g.c1 + 30); c++) { var v = book.val(si, c, r); row.push(v === null || F.isErr(v) ? null : typeof v === "boolean" ? String(v) : v); }
        rows.push(row);
      }
      return rows;
    }
    function chartSvg(ch, light) {
      var spec = Chart.fromRows(chartRows(ch));
      spec.type = ch.type; spec.title = ch.title;
      return Chart.svg(spec, { w: ch.w, h: ch.h, dark: !light && O.dark() });
    }
    function drawCharts() {
      var box = el("charts");
      if (!box) return;
      box.innerHTML = sh.charts.map(function (ch) {
        return '<div class="s-chart' + (chartSel === ch.id ? " is-sel" : "") + '" data-chart="' + esc(ch.id) + '" style="left:' + ch.x + "px;top:" + ch.y + "px;width:" + ch.w + "px;height:" + ch.h + 'px">' + chartSvg(ch) + '<span class="s-chart__rs" data-chart-rs></span></div>';
      }).join("");
    }
    function chartById(id) { return sh.charts.filter(function (c) { return c.id === id; })[0]; }
    function addChart(type) {
      if (readOnly()) return;
      var g = workRange();
      if (isOne(g) && !filled(g.c1, g.r1)) { O.notify(t("グラフにするデータを選んでください")); return; }
      var rc = rectOf({ c1: g.c2 + 1, r1: g.r1, c2: g.c2 + 1, r2: g.r1 }) || { x: 200, y: 40 };
      var ch = { id: "c" + Date.now().toString(36), type: type, ref: rangeText(g), title: "", x: Math.round(rc.x + 16), y: Math.round(rc.y), w: 480, h: 300 };
      var head = hasHeader(g) && g.c2 > g.c1 ? show(book.val(si, g.c1 + 1, g.r1)).text : "";
      if (g.c2 - g.c1 === 1 && head) ch.title = head;
      sh.charts.push(ch);
      chartSel = ch.id;
      changed();
      O.sync();
    }
    function chartCmd(cmd, at) {
      var ch = chartById(chartSel);
      if (!ch) return;
      if (cmd.indexOf("ctype:") === 0) { ch.type = cmd.slice(6); changed(); return; }
      if (cmd === "chartTitle") {
        M3.prompt({ title: t("グラフのタイトル"), label: t("タイトル"), value: ch.title || "", ok: t("OK") }).then(function (v) { if (v === null) return; ch.title = v; changed(); });
        return;
      }
      if (cmd === "chartRange") {
        M3.prompt({ title: t("データの範囲"), label: t("範囲（例：A1:C8）"), value: ch.ref, ok: t("OK") }).then(function (v) { var g = v && parseRange(v); if (!g) return; ch.ref = rangeText(g); changed(); });
        return;
      }
      if (cmd === "chartDelete") { sh.charts = sh.charts.filter(function (c) { return c !== ch; }); chartSel = null; changed(); O.sync(); return; }
      if (cmd === "chartPng") {
        var svg = chartSvg(ch, true), img = new Image();
        img.onload = function () {
          var cv = document.createElement("canvas"); cv.width = ch.w * 2; cv.height = ch.h * 2;
          var g2 = cv.getContext("2d"); g2.scale(2, 2); g2.drawImage(img, 0, 0);
          cv.toBlob(function (b) { O.saveBlob(O.fileName(".png"), b, "PNG", "image/png", ".png"); });
        };
        img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
      }
    }

    // ---------- 条件付き書式 ----------
    function cfMenu(at) {
      O.menu(at, [
        { head: t("セルの強調表示") },
        { value: "gt", label: t("指定の値より大きい…"), icon: "keyboard_double_arrow_up" },
        { value: "lt", label: t("指定の値より小さい…"), icon: "keyboard_double_arrow_down" },
        { value: "eq", label: t("指定の値に等しい…"), icon: "drag_handle" },
        { value: "text", label: t("文字列を含む…"), icon: "match_case" },
        { value: "dup", label: t("重複する値"), icon: "content_copy" },
        { sep: 1 },
        { value: "bar", label: t("データ バー"), icon: "data_usage" },
        { value: "scale", label: t("カラー スケール"), icon: "gradient" },
        { sep: 1 },
        { value: "clear", label: t("選んだセルのルールをクリア"), icon: "ink_eraser" },
        { value: "clearAll", label: t("シート全体のルールをクリア"), icon: "delete_sweep" }
      ]).then(function (v) {
        if (!v || readOnly()) return;
        var g = grow(range()), key = rangeText(isOne(g) ? workRange() : g);
        if (v === "clear") { sh.cf = sh.cf.filter(function (x) { var r = parseRange(x.ref); return r && (r.c2 < g.c1 || r.c1 > g.c2 || r.r2 < g.r1 || r.r1 > g.r2); }); changed(); return; }
        if (v === "clearAll") { sh.cf = []; changed(); return; }
        if (v === "dup" || v === "bar" || v === "scale") { sh.cf.push({ ref: key, type: v, color: v === "dup" ? "#fecaca" : undefined }); changed(); return; }
        var label = { gt: t("次の値より大きいセルに色を付けます"), lt: t("次の値より小さいセルに色を付けます"), eq: t("次の値と同じセルに色を付けます"), text: t("次の文字を含むセルに色を付けます") }[v];
        M3.prompt({ title: t("条件付き書式"), label: label, value: "", ok: t("OK") }).then(function (x) {
          if (x === null) return;
          sh.cf.push({ ref: key, type: v, v: x, color: v === "lt" ? "#fecaca" : v === "gt" ? "#bbf7d0" : "#fde68a" });
          changed();
        });
      });
    }

    // ---------- ステータスバー ----------
    function stat() {
      var g = grow(range()), h = "";
      if (editing) h = "<span>" + esc(editMode === "edit" ? t("編集") : t("入力")) + "</span>";
      else h = "<span>" + esc(readOnly() ? t("閲覧モード") : t("準備完了")) + "</span>";
      if (!isOne(g) && (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) <= 200000) {
        var sum = 0, n = 0, cnt = 0;
        var r2 = Math.min(g.r2, book.last(si).rows), c2 = Math.min(g.c2, book.last(si).cols - 1);
        for (var r = g.r1; r <= r2; r++) for (var c = g.c1; c <= c2; c++) {
          if (hiddenRows[r]) continue;
          var v = book.val(si, c, r);
          if (v !== null && v !== "") cnt++;
          if (typeof v === "number") { sum += v; n++; }
        }
        var nf = styleAt(act.c, act.r).nf, fmt = function (x) { return F.format(x, nf && !F.isDateFormat(nf) ? nf : "#,##0.##").text; };
        if (n) h += "<span>" + esc(t("平均: {v}", { v: fmt(sum / n) })) + "</span>";
        h += "<span>" + esc(t("データの個数: {n}", { n: cnt })) + "</span>";
        if (n) h += "<span>" + esc(t("合計: {v}", { v: fmt(sum) })) + "</span>";
      }
      O.stat(h);
    }

    // ---------- 画面を作る ----------
    function render(opts) {
      data = O.data();
      upgrade(data);
      book = Book(data);
      si = data.cur; sh = data.sheets[si];
      if (!(opts && opts.restore)) { act = anc = { c: 0, r: 1 }; nRows = 100; nCols = 26; chartSel = null; clip = null; }
      if (chartSel && !chartById(chartSel)) chartSel = null;
      editing = false;
      var cv = $("canvas"), sx = 0, sy = 0;
      if (opts && opts.restore && el("grid-scroll")) { sx = el("grid-scroll").scrollLeft; sy = el("grid-scroll").scrollTop; }
      cv.innerHTML =
        '<div class="sheet">' +
          '<div class="fx-bar"><input class="namebox" id="namebox" autocomplete="off" spellcheck="false" aria-label="' + esc(t("名前ボックス")) + '" title="' + esc(t("セルの番地を入れると、そこへ移ります")) + '">' +
            '<span class="fx-btns hide-sm"><button type="button" class="icon-btn m3-state" id="fx-cancel" aria-label="' + esc(t("キャンセル")) + '">' + icon("close") + '</button><button type="button" class="icon-btn m3-state" id="fx-ok" aria-label="' + esc(t("入力")) + '">' + icon("check") + "</button></span>" +
            '<button type="button" class="fx-ico m3-state" data-cmd="fnMenu" title="' + esc(t("関数の挿入")) + '">fx</button>' +
            '<input class="fx-input" id="fx" autocomplete="off" spellcheck="false" aria-label="' + esc(t("数式バー")) + '"></div>' +
          '<div class="grid-scroll" id="grid-scroll"><div class="grid-wrap" id="grid-wrap">' +
            '<table class="grid" id="grid" aria-label="' + esc(t("ワークシート")) + '"></table>' +
            '<div id="charts"></div><div class="clip-box" id="clip-box"></div><div class="fill-box" id="fill-box"></div>' +
            '<div class="sel-box" id="sel-box"><span class="sel-fill" id="sel-fill"></span></div>' +
            '<input class="cell-editor" id="ced" data-office-keys autocomplete="off" spellcheck="false" aria-label="' + esc(t("セル")) + '">' +
          "</div></div>" +
          '<div class="sheet-tabs" id="sheet-tabs"></div>' +
        "</div>";
      wire();
      drawGrid();
      if (opts && opts.restore) { el("grid-scroll").scrollLeft = sx; el("grid-scroll").scrollTop = sy; placeSel(); }
      applyZoom();
      focusGrid();
    }
    function applyZoom() {
      var w = el("grid-wrap");
      if (w) w.style.zoom = zoomZ;
      placeSel();
    }

    // ---------- クリック・ドラッグ・キー ----------
    function cellFrom(ev) {
      var x = document.elementFromPoint(ev.clientX, ev.clientY), td = x && x.closest && x.closest("td[data-c]");
      if (td) return { c: +td.getAttribute("data-c"), r: +td.getAttribute("data-r"), td: td };
      var th = x && x.closest && x.closest("th[data-col], th[data-row]");
      if (th) return th.hasAttribute("data-col") ? { c: +th.getAttribute("data-col"), r: Math.max(1, act.r) } : { c: Math.max(0, act.c), r: +th.getAttribute("data-row") };
      return null;
    }
    function drag(e, move, up) {
      var sc = el("grid-scroll"), timer = null, last = e;
      function auto() {
        var b = sc.getBoundingClientRect(), dx = 0, dy = 0;
        if (last.clientY > b.bottom - 8) dy = 24; else if (last.clientY < b.top + 36) dy = -24;
        if (last.clientX > b.right - 8) dx = 32; else if (last.clientX < b.left + 48) dx = -32;
        if (dx || dy) { sc.scrollTop += dy; sc.scrollLeft += dx; move(last); }
      }
      function mv(ev) { last = ev; move(ev); }
      function end(ev) {
        clearInterval(timer);
        window.removeEventListener("pointermove", mv);
        window.removeEventListener("pointerup", end);
        window.removeEventListener("pointercancel", end);
        if (up) up(ev);
      }
      timer = setInterval(auto, 60);
      window.addEventListener("pointermove", mv);
      window.addEventListener("pointerup", end);
      window.addEventListener("pointercancel", end);
    }
    function wire() {
      var wrap = el("grid-wrap"), sc = el("grid-scroll"), ced = el("ced"), fx = el("fx"), nb = el("namebox");
      sc.addEventListener("scroll", function () {
        if (sc.scrollTop + sc.clientHeight > sc.scrollHeight - 300) moreRows(100);
        if (sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 200 && nCols < 16384) moreCols(10);
      });
      wrap.addEventListener("pointerdown", function (e) {
        if (e.button !== 0) return;
        var tg = e.target;
        if (tg === ced && editing) return;
        var flt = tg.closest("[data-flt]");
        if (flt) { e.preventDefault(); filterDialog(+flt.getAttribute("data-flt")); return; }
        var chEl = tg.closest("[data-chart]");
        if (chEl) { e.preventDefault(); chartDrag(e, chEl, !!tg.closest("[data-chart-rs]")); return; }
        var rs = tg.closest("[data-rs]");
        if (rs) { e.preventDefault(); colResize(e, +rs.getAttribute("data-rs")); return; }
        var rrs = tg.closest("[data-rrs]");
        if (rrs) { e.preventDefault(); rowResize(e, +rrs.getAttribute("data-rrs")); return; }
        if (tg.closest("#sel-fill")) { e.preventDefault(); fillDrag(e); return; }
        if (tg.closest("[data-corner]")) { e.preventDefault(); endEdit(true); selectRange({ c1: 0, r1: 1, c2: nCols - 1, r2: nRows }); focusGrid(); return; }
        var colTh = tg.closest("th[data-col]"), rowTh = tg.closest("th[data-row]");
        if (colTh || rowTh) {
          e.preventDefault();
          endEdit(true);
          // 列（行）全体を選ぶ。選んだセル（act）は、押した列のいちばん上（押した行の左はし）
          var isCol = !!colTh, i0 = +(isCol ? colTh.getAttribute("data-col") : rowTh.getAttribute("data-row"));
          var a0 = e.shiftKey ? (isCol ? act.c : act.r) : i0;
          var pick = function (i) {
            if (isCol) { act = { c: a0, r: 1 }; anc = { c: i, r: nRows }; }
            else { act = { c: 0, r: a0 }; anc = { c: nCols - 1, r: i }; }
            placeSel();
          };
          pick(i0);
          focusGrid();
          drag(e, function (ev) { var p = cellFrom(ev); if (p) pick(isCol ? p.c : p.r); });
          return;
        }
        var p = cellFrom(e);
        if (!p) return;
        e.preventDefault();
        // 数式を書いているときは、押したセルの番地を入れる
        if (editing && formulaPoint()) {
          var start = { c: p.c, r: p.r };
          insertRef(ref(p.c, p.r));
          drag(e, function (ev) { var q = cellFrom(ev); if (q) insertRef(rangeText(norm({ c1: start.c, r1: start.r, c2: q.c, r2: q.r }))); });
          return;
        }
        endEdit(true);
        if (chartSel) { chartSel = null; drawCharts(); }
        if (e.shiftKey) { anc = { c: p.c, r: p.r }; }
        else { anc = act = { c: p.c, r: p.r }; }
        placeSel();
        focusGrid();
        O.sync();
        drag(e, function (ev) { var q = cellFrom(ev); if (q && (q.c !== anc.c || q.r !== anc.r)) { anc = { c: q.c, r: q.r }; placeSel(); } }, function () {
          if (painter) { applyPainter(); }
        });
      });
      wrap.addEventListener("dblclick", function (e) {
        if (e.target.closest("[data-rs]")) { var c = +e.target.closest("[data-rs]").getAttribute("data-rs"); autoWidth(c); changed(true); return; }
        if (e.target.closest("[data-chart]")) { chartCmd("chartTitle"); return; }
        var p = cellFrom(e);
        if (p && p.td) { act = anc = { c: p.c, r: p.r }; placeSel(); startEdit("edit", rawAt(p.c, p.r)); }
      });
      wrap.addEventListener("contextmenu", function (e) {
        var p = cellFrom(e);
        if (!p) return;
        e.preventDefault();
        if (!inRange(grow(range()), p.c, p.r)) { endEdit(true); act = anc = { c: p.c, r: p.r }; placeSel(); }
        cellMenu({ x: e.clientX, y: e.clientY });
      });
      ced.addEventListener("keydown", onKey);
      ced.addEventListener("input", function () {
        refIns = null;
        if (!editing) { var v = ced.value; startEdit("enter"); ced.value = v; }
        fx.value = ced.value;
        fitEditor();
      });
      ced.addEventListener("compositionstart", function () { if (!editing) startEdit("enter"); });
      ced.addEventListener("copy", function (e) { onCopy(e, false); });
      ced.addEventListener("cut", function (e) { if (readOnly()) return; onCopy(e, true); });
      ced.addEventListener("paste", onPaste);
      fx.addEventListener("focus", function () { if (!editing) { startEdit("edit", rawAt(act.c, act.r), true); } });
      fx.addEventListener("input", function () { refIns = null; ced.value = fx.value; if (!editing) startEdit("edit", undefined, true); });
      fx.addEventListener("keydown", function (e) {
        if (e.isComposing || e.keyCode === 229) return;
        if (e.key === "Enter") { e.preventDefault(); ced.value = fx.value; endEdit(true); select(act.c, act.r + 1); focusGrid(); }
        else if (e.key === "Tab") { e.preventDefault(); ced.value = fx.value; endEdit(true); select(act.c + 1, act.r); focusGrid(); }
        else if (e.key === "Escape") { e.preventDefault(); endEdit(false); focusGrid(); }
      });
      el("fx-ok").addEventListener("click", function () { if (editing) { ced.value = editEl() === fx ? fx.value : ced.value; endEdit(true); } focusGrid(); });
      el("fx-cancel").addEventListener("click", function () { endEdit(false); focusGrid(); });
      nb.addEventListener("focus", function () { nb.select(); });
      nb.addEventListener("keydown", function (e) {
        if (e.key === "Escape") { nb.value = ref(act.c, act.r); focusGrid(); return; }
        if (e.key !== "Enter" || e.isComposing) return;
        e.preventDefault();
        var v = nb.value.trim(), m = /^(.*)!(.+)$/.exec(v);
        if (m) { var i = book.sheetIndex(m[1].replace(/^'|'$/g, "")); if (i >= 0) switchSheet(i); v = m[2]; }
        var g = parseRange(v);
        if (!g) { O.notify(t("「{ref}」というセルはありません", { ref: v })); return; }
        endEdit(true);
        selectRange(g);
        focusGrid();
      });
      el("sheet-tabs").addEventListener("click", function (e) {
        var b = e.target.closest("[data-sheet]");
        if (!b) return;
        var i = +b.getAttribute("data-sheet");
        if (i === si) sheetMenu(i, b); else switchSheet(i);
      });
      el("sheet-tabs").addEventListener("dblclick", function (e) { var b = e.target.closest("[data-sheet]"); if (b) renameSheet(+b.getAttribute("data-sheet")); });
      el("sheet-tabs").addEventListener("contextmenu", function (e) { var b = e.target.closest("[data-sheet]"); if (!b) return; e.preventDefault(); var i = +b.getAttribute("data-sheet"); if (i !== si) switchSheet(i); sheetMenu(si, { x: e.clientX, y: e.clientY }); });
    }
    function colResize(e, c) {
      var x0 = e.clientX, w0 = colW(c);
      drag(e, function (ev) {
        var w = Math.max(16, Math.round(w0 + (ev.clientX - x0) / zoomZ));
        var col = el("grid").querySelectorAll("col")[c + 1];
        if (col) col.style.width = w + "px";
        sh.cw[c] = w;
        var total = HEADW; for (var i = 0; i < nCols; i++) total += colW(i);
        el("grid").style.width = total + "px";
        placeSel();
      }, function () { changed(true); });
    }
    function rowResize(e, r) {
      var y0 = e.clientY, h0 = sh.rh[r] || 28;
      drag(e, function (ev) {
        var h = Math.max(12, Math.round(h0 + (ev.clientY - y0) / zoomZ));
        sh.rh[r] = h;
        var tr = el("grid").querySelector('tr[data-r="' + r + '"]');
        if (tr) tr.style.height = h + "px";
        placeSel();
      }, function () { changed(true); });
    }
    function fillDrag(e) {
      if (readOnly()) return;
      var src = grow(range()), dst = src, box = el("fill-box");
      drag(e, function (ev) {
        var p = cellFrom(ev);
        if (!p) return;
        var dr = p.r > src.r2 ? p.r - src.r2 : p.r < src.r1 ? p.r - src.r1 : 0, dc = p.c > src.c2 ? p.c - src.c2 : p.c < src.c1 ? p.c - src.c1 : 0;
        if (Math.abs(dr) >= Math.abs(dc) && dr) dst = { c1: src.c1, c2: src.c2, r1: Math.min(src.r1, p.r), r2: Math.max(src.r2, p.r) };
        else if (dc) dst = { r1: src.r1, r2: src.r2, c1: Math.min(src.c1, p.c), c2: Math.max(src.c2, p.c) };
        else dst = src;
        placeBox(box, dst === src ? null : dst);
      }, function () {
        box.style.display = "none";
        if (dst !== src) fillFrom(src, dst);
        focusGrid();
      });
    }
    function chartDrag(e, chEl, resize) {
      var ch = chartById(chEl.getAttribute("data-chart"));
      if (!ch) return;
      endEdit(true);
      chartSel = ch.id;
      chEl.classList.add("is-sel");
      el("grid-wrap").querySelectorAll(".s-chart").forEach(function (x) { if (x !== chEl) x.classList.remove("is-sel"); });
      O.sync();
      focusGrid();
      var x0 = e.clientX, y0 = e.clientY, a = { x: ch.x, y: ch.y, w: ch.w, h: ch.h }, moved = false;
      drag(e, function (ev) {
        var dx = (ev.clientX - x0) / zoomZ, dy = (ev.clientY - y0) / zoomZ;
        if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
        if (resize) { ch.w = Math.max(200, Math.round(a.w + dx)); ch.h = Math.max(140, Math.round(a.h + dy)); chEl.style.width = ch.w + "px"; chEl.style.height = ch.h + "px"; }
        else { ch.x = Math.max(0, Math.round(a.x + dx)); ch.y = Math.max(0, Math.round(a.y + dy)); chEl.style.left = ch.x + "px"; chEl.style.top = ch.y + "px"; }
      }, function () { if (moved) { if (resize) drawCharts(); O.changed(); } });
    }
    function applyPainter() {
      var p = painter;
      painter = null;
      var g = grow(range()), h = p.rows.length, w = p.rows[0].length;
      eachCell(g, function (c, r, k) { var st = p.rows[(r - g.r1) % h][(c - g.c1) % w]; if (st) sh.style[k] = st; else delete sh.style[k]; });
      changed();
      O.sync();
    }
    function moveBy(dc, dr, extend) {
      var from = extend ? anc : act, m = mergeAt(from.c, from.r);
      var c = from.c, r = from.r;
      if (m && !extend) { if (dc > 0) c = m.c2; if (dr > 0) r = m.r2; if (dc < 0) c = m.c1; if (dr < 0) r = m.r1; }
      c += dc; r += dr;
      // 隠れた行はとばす
      while (dr && hiddenRows[r] && r > 1) r += dr > 0 ? 1 : -1;
      var mm = mergeAt(c, r);
      if (mm && !extend) { c = mm.c1; r = mm.r1; }
      select(c, r, extend);
    }
    function jump(dc, dr, from) {
      // Ctrl+矢印: データの端へ
      var c = from.c, r = from.r, last = book.last(si);
      var cur = filled(c, r), next = filled(c + dc, r + dr);
      if (cur && next) { while (c + dc >= 0 && r + dr >= 1 && filled(c + dc, r + dr)) { c += dc; r += dr; } }
      else {
        c += dc; r += dr;
        while (c >= 0 && r >= 1 && !filled(c, r) && c <= last.cols && r <= last.rows) { c += dc; r += dr; }
        if (c > last.cols || r > last.rows || c < 0 || r < 1) { c = dc > 0 ? Math.max(last.cols - 1, act.c) : dc < 0 ? 0 : c; r = dr > 0 ? Math.max(last.rows, act.r) : dr < 0 ? 1 : r; }
      }
      return { c: Math.max(0, c), r: Math.max(1, r) };
    }
    function onKey(e) {
      if (e.isComposing || e.keyCode === 229) return;
      var k = e.key, mod = e.ctrlKey || e.metaKey;
      if (editing) {
        if (k === "Enter" && e.altKey) {
          e.preventDefault();
          var ced = el("ced"), p = ced.selectionStart;
          ced.value = ced.value.slice(0, p) + "\n" + ced.value.slice(ced.selectionEnd);
          ced.setSelectionRange(p + 1, p + 1);
          el("fx").value = ced.value;
          return;
        }
        if (k === "Enter" && mod) { e.preventDefault(); var v = el("ced").value; endEdit(false); fillValue(v); return; }
        if (k === "Enter") { e.preventDefault(); endEdit(true); moveBy(0, e.shiftKey ? -1 : 1); return; }
        if (k === "Tab") { e.preventDefault(); endEdit(true); moveBy(e.shiftKey ? -1 : 1, 0); return; }
        if (k === "Escape") { e.preventDefault(); endEdit(false); return; }
        if (k === "F2") { e.preventDefault(); editMode = editMode === "edit" ? "enter" : "edit"; stat(); return; }
        if (editMode === "enter" && /^Arrow/.test(k) && !formulaPoint()) { e.preventDefault(); endEdit(true); arrowKey(k, false); return; }
        if (editMode === "enter" && /^Arrow/.test(k) && formulaPoint()) {
          // 数式の中で矢印: 番地を動かして入れる
          e.preventDefault();
          var base = refIns ? parseRange(el("ced").value.substr(refIns.start, refIns.len)) : null, pos = base ? { c: base.c1, r: base.r1 } : { c: act.c, r: act.r };
          var d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
          insertRef(ref(Math.max(0, pos.c + d[0]), Math.max(1, pos.r + d[1])));
          return;
        }
        return;
      }
      if (chartSel && (k === "Delete" || k === "Backspace")) { e.preventDefault(); chartCmd("chartDelete"); return; }
      if (k === "Escape") { if (clip) { clip = null; placeSel(); } if (chartSel) { chartSel = null; drawCharts(); O.sync(); } painter = null; O.sync(); return; }
      if (/^Arrow/.test(k)) { e.preventDefault(); arrowKey(k, e.shiftKey, mod); return; }
      if (k === "Enter") { e.preventDefault(); moveBy(0, e.shiftKey ? -1 : 1); return; }
      if (k === "Tab") { e.preventDefault(); moveBy(e.shiftKey ? -1 : 1, 0); return; }
      if (k === "Home") { e.preventDefault(); if (mod) select(0, 1, e.shiftKey); else select(0, act.r, e.shiftKey); return; }
      if (k === "End" && mod) { e.preventDefault(); var l = book.last(si); select(Math.max(0, l.cols - 1), Math.max(1, l.rows), e.shiftKey); return; }
      if (k === "PageDown" || k === "PageUp") {
        e.preventDefault();
        if (mod) { switchSheet(si + (k === "PageDown" ? 1 : -1)); return; }
        moveBy(0, k === "PageDown" ? 20 : -20, e.shiftKey);
        return;
      }
      if (k === "Delete" || k === "Backspace") { e.preventDefault(); if (readOnly()) return; clearRange("content"); return; }
      if (k === "F2") { e.preventDefault(); startEdit("edit", rawAt(act.c, act.r)); return; }
      if (mod && !e.altKey) {
        var kk = k.toLowerCase();
        if (kk === "b") { e.preventDefault(); toggleStyle("b"); return; }
        if (kk === "i") { e.preventDefault(); toggleStyle("i"); return; }
        if (kk === "u") { e.preventDefault(); toggleStyle("u"); return; }
        if (kk === "5") { e.preventDefault(); toggleStyle("s"); return; }
        if (kk === "a") { e.preventDefault(); var rg = region(act.c, act.r); if (isOne(range()) && !isOne(rg)) selectRange(rg); else selectRange({ c1: 0, r1: 1, c2: nCols - 1, r2: nRows }); return; }
        if (kk === "d") { e.preventDefault(); fillDir("down"); return; }
        if (kk === "r") { e.preventDefault(); fillDir("right"); return; }
        if (k === ";" || k === "+") { e.preventDefault(); var dd = new Date(); commitCell(act.c, act.r, dd.getFullYear() + "/" + (dd.getMonth() + 1) + "/" + dd.getDate()); return; }
        if (k === ":" || k === "*") { e.preventDefault(); var tn = new Date(); commitCell(act.c, act.r, tn.getHours() + ":" + ("0" + tn.getMinutes()).slice(-2)); return; }
      }
      if (e.altKey && (k === "=" || k === "-") && !mod) { e.preventDefault(); autoFn("SUM"); return; }
    }
    function arrowKey(k, extend, mod) {
      var d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
      if (mod) { var p = jump(d[0], d[1], extend ? anc : act); select(p.c, p.r, extend); return; }
      moveBy(d[0], d[1], extend);
    }
    function fillValue(v) {
      if (readOnly()) return;
      var g = grow(range());
      eachCell(g, function (c, r, k) {
        var x = v.charAt(0) === "=" ? "=" + F.shift(v.slice(1), c - act.c, r - act.r) : v;
        if (x === "") delete sh.cells[k]; else sh.cells[k] = x;
      });
      changed();
    }
    function cellMenu(at) {
      O.menu(at, [
        { value: "cut", label: t("切り取り"), icon: "content_cut", sub: "Ctrl+X" },
        { value: "copy", label: t("コピー"), icon: "content_copy", sub: "Ctrl+C" },
        { value: "paste", label: t("貼り付け"), icon: "content_paste", sub: "Ctrl+V", disabled: !clip },
        { value: "pasteValues", label: t("値だけ貼り付け"), icon: "content_paste_go", disabled: !clip },
        { sep: 1 },
        { value: "rowAbove", label: t("上に行を挿入"), icon: "add_row_above" },
        { value: "colLeft", label: t("左に列を挿入"), icon: "add_column_left" },
        { value: "delRow", label: t("行を削除"), icon: "delete" },
        { value: "delCol", label: t("列を削除"), icon: "delete" },
        { sep: 1 },
        { value: "clear", label: t("数式と値のクリア"), icon: "ink_eraser", sub: "Delete" },
        { value: "sortAsc", label: t("昇順で並べ替え"), icon: "arrow_upward" },
        { value: "sortDesc", label: t("降順で並べ替え"), icon: "arrow_downward" },
        { value: "merge", label: t("セルを結合して中央揃え"), icon: "cell_merge" }
      ]).then(function (v) {
        if (!v) return;
        focusGrid();
        if (v === "cut" || v === "copy") { clip = clipData(range(), v === "cut"); placeSel(); copyText(clip.text); return; }
        if (v === "paste") return clip && pasteClip(clip);
        if (v === "pasteValues") return clip && pasteClip(clip, true);
        var map = { rowAbove: function () { insertRows("above"); }, colLeft: function () { insertCols("left"); }, delRow: deleteRows, delCol: deleteCols, clear: function () { clearRange("content"); },
          sortAsc: function () { sortRange(false); }, sortDesc: function () { sortRange(true); }, merge: toggleMerge };
        if (readOnly()) { O.notify(t("閲覧モードでは書きかえられません")); return; }
        map[v]();
      });
    }
    function copyText(text) {
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).catch(function () {});
    }

    // ---------- リボンのボタン ----------
    function command(cmd, at) {
      if (cmd.indexOf("chart:") === 0) return addChart(cmd.slice(6));
      if (cmd.indexOf("ctype:") === 0 || /^chart(Title|Range|Delete|Png)$/.test(cmd)) return chartCmd(cmd, at);
      var g = range(), st = styleAt(act.c, act.r);
      var ro = readOnly() && !/^(sheetPrev|sheetNext|freeze|gridlines|fnMenu|fncat:|recalc)/.test(cmd);
      if (ro) { O.notify(t("閲覧モードでは書きかえられません")); return; }
      // 関数を選ぶときは、書いている数式をそのままにする
      if (!/^(fnMenu|fncat:)/.test(cmd)) endEdit(true);
      switch (cmd) {
        case "bold": return toggleStyle("b");
        case "italic": return toggleStyle("i");
        case "underline": return toggleStyle("u");
        case "strike": return toggleStyle("s");
        case "wrap": return toggleStyle("wrap");
        case "al:l": case "al:c": case "al:r": setStyle(g, { al: st.al === cmd.slice(3) ? null : cmd.slice(3) }); return changed();
        case "va:t": case "va:m": case "va:b": setStyle(g, { va: st.va === cmd.slice(3) ? null : cmd.slice(3) }); return changed();
        case "merge": return toggleMerge();
        case "grow": case "shrink": {
          var cur = st.fs || 11, list = O.sizes, i = 0;
          if (cmd === "grow") { while (i < list.length && list[i] <= cur) i++; setStyle(g, { fs: list[Math.min(i, list.length - 1)] }); }
          else { i = list.length - 1; while (i > 0 && list[i] >= cur) i--; setStyle(g, { fs: list[i] }); }
          return changed();
        }
        case "fontName":
          return O.menu(at, O.fonts.map(function (f) { return { value: f[0], label: f[0], style: "font-family:" + f[1], checked: (st.ff || "M PLUS 1") === f[0] }; })).then(function (v) { if (v) { setStyle(g, { ff: v === "M PLUS 1" ? null : v }); changed(); } });
        case "fontSize":
          return O.menu(at, O.sizes.map(function (s) { return { value: String(s), label: String(s), checked: (st.fs || 11) === s }; })).then(function (v) { if (v) { setStyle(g, { fs: +v === 11 ? null : +v }); changed(); } });
        case "color":
          return O.palette(at, { none: t("自動"), value: st.fg }).then(function (v) { if (v) { setStyle(g, { fg: v === "none" ? null : v }); changed(); } });
        case "fill":
          return O.palette(at, { none: t("塗りつぶしなし"), value: st.bg }).then(function (v) { if (v) { setStyle(g, { bg: v === "none" ? null : v }); changed(); } });
        case "border":
          return O.menu(at, [
            { value: "bottom", label: t("下罫線"), icon: "border_bottom" }, { value: "top", label: t("上罫線"), icon: "border_top" },
            { value: "left", label: t("左罫線"), icon: "border_left" }, { value: "right", label: t("右罫線"), icon: "border_right" },
            { sep: 1 }, { value: "all", label: t("格子"), icon: "border_all" }, { value: "out", label: t("外枠"), icon: "border_outer" },
            { value: "thick", label: t("太い外枠"), icon: "border_style" }, { value: "double", label: t("下太罫線"), icon: "border_bottom" },
            { sep: 1 }, { value: "none", label: t("枠なし"), icon: "border_clear" }
          ]).then(function (v) { if (v) borders(v); });
        case "nf":
          return O.menu(at, NF.map(function (x) {
            var v = book.val(si, act.c, act.r), sample = typeof v === "number" ? F.format(v, x.v || null).text : "";
            return { value: x.v || "general", label: x.label, sub: sample, checked: (st.nf || "") === x.v };
          }).concat([{ sep: 1 }, { value: "custom", label: t("その他の表示形式…"), icon: "tune" }])).then(function (v) {
            if (!v) return;
            if (v === "custom") {
              return M3.prompt({ title: t("表示形式"), label: t("形式（例：#,##0「円」は #,##0\"円\"）"), value: st.nf || "", ok: t("OK") }).then(function (x) { if (x === null) return; setStyle(g, { nf: x || null }); changed(); });
            }
            setStyle(g, { nf: v === "general" ? null : v });
            changed();
          });
        case "yen": setStyle(g, { nf: st.nf === "¥#,##0" ? null : "¥#,##0" }); return changed();
        case "pct": setStyle(g, { nf: st.nf === "0%" ? null : "0%" }); return changed();
        case "comma": setStyle(g, { nf: st.nf === "#,##0" ? null : "#,##0" }); return changed();
        case "decUp": case "decDown": setStyle(g, { nf: decimals(st.nf, cmd === "decUp" ? 1 : -1) || null }); return changed();
        case "insertMenu":
          return O.menu(at, [
            { value: "above", label: t("上に行を挿入"), icon: "add_row_above" }, { value: "below", label: t("下に行を挿入"), icon: "add_row_below" },
            { value: "left", label: t("左に列を挿入"), icon: "add_column_left" }, { value: "right", label: t("右に列を挿入"), icon: "add_column_right" },
            { sep: 1 }, { value: "sheet", label: t("シートを挿入"), icon: "tab" }
          ]).then(function (v) { if (v === "above" || v === "below") insertRows(v); else if (v === "left" || v === "right") insertCols(v); else if (v === "sheet") addSheet(); });
        case "deleteMenu":
          return O.menu(at, [{ value: "row", label: t("行を削除"), icon: "delete" }, { value: "col", label: t("列を削除"), icon: "delete" }, { sep: 1 }, { value: "sheet", label: t("シートを削除"), icon: "tab_close" }])
            .then(function (v) { if (v === "row") deleteRows(); else if (v === "col") deleteCols(); else if (v === "sheet") deleteSheet(si); });
        case "sizeMenu":
          return O.menu(at, [
            { value: "colw", label: t("列の幅…"), icon: "width" }, { value: "auto", label: t("列の幅を自動調整"), icon: "fit_width" }, { value: "rowh", label: t("行の高さ…"), icon: "height" },
            { sep: 1 }, { value: "rename", label: t("シートの名前を変更"), icon: "edit" }
          ]).then(function (v) {
            if (v === "colw") askWidth(false); else if (v === "rowh") askWidth(true); else if (v === "rename") renameSheet(si);
            else if (v === "auto") { var gg = grow(range()); for (var c = gg.c1; c <= gg.c2; c++) autoWidth(c); changed(true); }
          });
        case "fillMenu":
          return O.menu(at, [{ value: "down", label: t("下方向へコピー"), icon: "arrow_downward", sub: "Ctrl+D" }, { value: "right", label: t("右方向へコピー"), icon: "arrow_forward", sub: "Ctrl+R" },
            { value: "up", label: t("上方向へコピー"), icon: "arrow_upward" }, { value: "left", label: t("左方向へコピー"), icon: "arrow_back" }]).then(function (v) { if (v) fillDir(v); });
        case "clearMenu":
          return O.menu(at, [{ value: "all", label: t("すべてクリア"), icon: "delete_sweep" }, { value: "format", label: t("書式のクリア"), icon: "format_clear" }, { value: "content", label: t("数式と値のクリア"), icon: "ink_eraser" }])
            .then(function (v) { if (v) clearRange(v); });
        case "sortMenu":
          return O.menu(at, [{ value: "asc", label: t("昇順"), icon: "arrow_upward" }, { value: "desc", label: t("降順"), icon: "arrow_downward" }, { sep: 1 }, { value: "filter", label: t("フィルター"), icon: "filter_alt", checked: !!sh.filter }])
            .then(function (v) { if (v === "asc" || v === "desc") sortRange(v === "desc"); else if (v === "filter") toggleFilter(); });
        case "sort:asc": return sortRange(false);
        case "sort:desc": return sortRange(true);
        case "filter": return toggleFilter();
        case "filterClear": if (sh.filter) { sh.filter.hide = {}; changed(true); } return;
        case "dedupe": return dedupe();
        case "cf": return cfMenu(at);
        case "painter":
          if (painter) { painter = null; O.sync(); return; }
          var src = grow(range()), rows = [];
          for (var r = src.r1; r <= src.r2; r++) { var row = []; for (var c2 = src.c1; c2 <= src.c2; c2++) row.push(sh.style[ref(c2, r)] || null); rows.push(row); }
          painter = { rows: rows };
          O.sync();
          O.notify(t("書式をはり付けるセルを選んでください"));
          return;
        case "fn:SUM": case "fn:AVERAGE": case "fn:COUNT": case "fn:MAX": case "fn:MIN": return autoFn(cmd.slice(3));
        case "autoSumMenu":
          return O.menu(at, [{ value: "SUM", label: t("合計"), icon: "functions" }, { value: "AVERAGE", label: t("平均"), icon: "avg_pace" }, { value: "COUNT", label: t("数値の個数"), icon: "tag" },
            { value: "MAX", label: t("最大値"), icon: "arrow_upward" }, { value: "MIN", label: t("最小値"), icon: "arrow_downward" }]).then(function (v) { if (v) autoFn(v); });
        case "fnMenu": return fnMenu(at);
        case "recalc": book.reset(); refreshCells(); drawCharts(); O.notify(t("再計算しました")); return;
        case "today": var d0 = new Date(); return commitCell(act.c, act.r, d0.getFullYear() + "/" + (d0.getMonth() + 1) + "/" + d0.getDate());
        case "now": var n0 = new Date(); return commitCell(act.c, act.r, n0.getHours() + ":" + ("0" + n0.getMinutes()).slice(-2));
        case "freeze":
          return O.menu(at, [{ value: "here", label: t("ウィンドウ枠の固定（選んだセルの上と左）"), icon: "grid_view" }, { value: "row", label: t("先頭行の固定"), icon: "table_rows" },
            { value: "col", label: t("先頭列の固定"), icon: "view_column" }, { sep: 1 }, { value: "none", label: t("固定の解除"), icon: "close", disabled: !sh.freeze }]).then(function (v) { if (v) freeze(v); });
        case "gridlines": sh.grid = sh.grid === false ? undefined : false; if (sh.grid === undefined) delete sh.grid; return changed(true);
        case "headRow": if (sh.head === act.r) delete sh.head; else sh.head = act.r; return changed(true);
        case "sheetAdd": return addSheet();
        case "sheetPrev": return switchSheet(si - 1);
        case "sheetNext": return switchSheet(si + 1);
      }
      var cat = FN_CATS.filter(function (c) { return "fncat:" + c.id === cmd; })[0];
      if (cat) return fnMenu(at, cat);
    }

    // ---------- リボン ----------
    var R = O.groups;
    var ribbon = [
      { tab: t("ホーム"), groups: [R.clip,
        { label: t("フォント"), items: [{ rows: [
          [{ chip: "M PLUS 1", w: 128, cmd: "fontName", key: "font", label: t("フォント") }, { chip: "11", w: 56, cmd: "fontSize", key: "size", label: t("フォント サイズ") },
           { icon: "text_increase", label: t("フォント サイズの拡大"), cmd: "grow" }, { icon: "text_decrease", label: t("フォント サイズの縮小"), cmd: "shrink" }],
          [{ icon: "format_bold", label: t("太字"), sub: "Ctrl+B", cmd: "bold", toggle: "b" }, { icon: "format_italic", label: t("斜体"), sub: "Ctrl+I", cmd: "italic", toggle: "i" },
           { icon: "format_underlined", label: t("下線"), sub: "Ctrl+U", cmd: "underline", toggle: "u" }, { icon: "format_strikethrough", label: t("取り消し線"), cmd: "strike", toggle: "s" },
           { icon: "border_all", label: t("罫線"), cmd: "border", drop: 1 }, { icon: "format_color_fill", label: t("塗りつぶしの色"), cmd: "fill", drop: 1 }, { icon: "format_color_text", label: t("フォントの色"), cmd: "color", drop: 1 }]
        ] }] },
        { label: t("配置"), items: [{ rows: [
          [{ icon: "vertical_align_top", label: t("上揃え"), cmd: "va:t", toggle: "va:t" }, { icon: "vertical_align_center", label: t("上下中央揃え"), cmd: "va:m", toggle: "va:m" }, { icon: "vertical_align_bottom", label: t("下揃え"), cmd: "va:b", toggle: "va:b" },
           { icon: "wrap_text", label: t("折り返して全体を表示する"), cmd: "wrap", toggle: "wrap" }],
          [{ icon: "format_align_left", label: t("左揃え"), cmd: "al:l", toggle: "al:l" }, { icon: "format_align_center", label: t("中央揃え"), cmd: "al:c", toggle: "al:c" }, { icon: "format_align_right", label: t("右揃え"), cmd: "al:r", toggle: "al:r" },
           { icon: "cell_merge", label: t("セルを結合して中央揃え"), cmd: "merge", toggle: "merge" }]
        ] }] },
        { label: t("数値"), items: [{ rows: [
          [{ chip: t("標準"), w: 132, cmd: "nf", key: "nf", label: t("表示形式") }],
          [{ icon: "currency_yen", label: t("通貨表示形式"), cmd: "yen" }, { icon: "percent", label: t("パーセント スタイル"), cmd: "pct" }, { icon: "format_list_numbered_rtl", label: t("桁区切りスタイル"), cmd: "comma" },
           { icon: "decimal_increase", label: t("小数点以下の表示桁数を増やす"), cmd: "decUp" }, { icon: "decimal_decrease", label: t("小数点以下の表示桁数を減らす"), cmd: "decDown" }]
        ] }] },
        { label: t("スタイル"), items: [{ big: 1, icon: "format_color_fill", label: t("条件付き書式"), cmd: "cf", drop: 1 }, { stack: [{ icon: "table_rows", label: t("見出しの行"), cmd: "headRow", toggle: "headRow" }] }] },
        { label: t("セル"), items: [{ stack: [{ icon: "add_box", label: t("挿入"), cmd: "insertMenu", drop: 1 }, { icon: "indeterminate_check_box", label: t("削除"), cmd: "deleteMenu", drop: 1 }, { icon: "width", label: t("書式"), cmd: "sizeMenu", drop: 1 }] }] },
        { label: t("編集"), items: [{ big: 1, icon: "functions", label: t("オート SUM"), cmd: "fn:SUM" },
          { stack: [{ icon: "arrow_downward", label: t("フィル"), cmd: "fillMenu", drop: 1 }, { icon: "ink_eraser", label: t("クリア"), cmd: "clearMenu", drop: 1 }, { icon: "sort", label: t("並べ替えとフィルター"), cmd: "sortMenu", drop: 1 }] },
          { stack: [{ icon: "search", label: t("検索"), cmd: "find" }, { icon: "find_replace", label: t("置換"), cmd: "replace" }] }] }
      ] },
      { tab: t("挿入"), groups: [
        { label: t("グラフ"), items: [{ big: 1, icon: "bar_chart", label: t("縦棒"), cmd: "chart:col" }, { big: 1, icon: "show_chart", label: t("折れ線"), cmd: "chart:line" }, { big: 1, icon: "pie_chart", label: t("円"), cmd: "chart:pie" },
          { stack: [{ icon: "align_horizontal_left", label: t("横棒"), cmd: "chart:bar" }, { icon: "area_chart", label: t("面"), cmd: "chart:area" }, { icon: "scatter_plot", label: t("散布図"), cmd: "chart:scatter" }] }] },
        { label: t("関数"), items: [{ big: 1, icon: "functions", label: t("関数の挿入"), cmd: "fnMenu" }] },
        { label: t("テキスト"), items: [{ stack: [{ icon: "today", label: t("今日の日付"), sub: "Ctrl+;", cmd: "today" }, { icon: "schedule", label: t("いまの時刻"), cmd: "now" }] }] },
        { label: t("セル"), items: [{ big: 1, icon: "add_box", label: t("挿入"), cmd: "insertMenu", drop: 1 }] }
      ] },
      { tab: t("数式"), groups: [
        { label: t("関数ライブラリ"), items: [{ big: 1, icon: "functions", label: t("関数の挿入"), cmd: "fnMenu" }, { big: 1, icon: "add_circle", label: t("オート SUM"), cmd: "autoSumMenu", drop: 1 }]
          .concat(FN_CATS.map(function (c) { return { big: 1, icon: c.icon, label: c.label, cmd: "fncat:" + c.id, drop: 1 }; })) },
        { label: t("計算"), items: [{ big: 1, icon: "calculate", label: t("再計算"), cmd: "recalc" }] }
      ] },
      { tab: t("データ"), groups: [
        { label: t("並べ替えとフィルター"), items: [{ big: 1, icon: "arrow_upward", label: t("昇順"), cmd: "sort:asc" }, { big: 1, icon: "arrow_downward", label: t("降順"), cmd: "sort:desc" },
          { big: 1, icon: "filter_alt", label: t("フィルター"), cmd: "filter", toggle: "filter" }, { stack: [{ icon: "filter_alt_off", label: t("クリア"), cmd: "filterClear" }] }] },
        { label: t("データ ツール"), items: [{ big: 1, icon: "layers_clear", label: t("重複の削除"), cmd: "dedupe" }, { big: 1, icon: "format_color_fill", label: t("条件付き書式"), cmd: "cf", drop: 1 }] }
      ] },
      { tab: t("表示"), groups: [R.view,
        { label: t("表示"), items: [{ stack: [{ icon: "grid_on", label: t("目盛線"), cmd: "gridlines", toggle: "grid" }] }] },
        { label: t("ウィンドウ"), items: [{ big: 1, icon: "grid_view", label: t("ウィンドウ枠の固定"), cmd: "freeze", drop: 1, toggle: "frozen" }] },
        R.zoom] },
      { tab: t("グラフのデザイン"), ctx: 1, when: function () { return !!chartSel; }, groups: [
        { label: t("種類"), items: Chart.TYPES.map(function (x) { return { big: 1, icon: x.icon, label: t(x.label), cmd: "ctype:" + x.value, toggle: "ctype:" + x.value }; }) },
        { label: t("データ"), items: [{ big: 1, icon: "title", label: t("タイトル"), cmd: "chartTitle" }, { big: 1, icon: "data_array", label: t("データの範囲"), cmd: "chartRange" }] },
        { label: t("グラフ"), items: [{ big: 1, icon: "image", label: t("画像で保存"), cmd: "chartPng" }, { big: 1, icon: "delete", label: t("削除"), cmd: "chartDelete" }] }
      ] }
    ];
    function state() {
      if (!sh) return {};
      var st = styleAt(act.c, act.r), g = grow(range()), ch = chartSel && chartById(chartSel), s = {
        b: !!st.b, i: !!st.i, u: !!st.u, s: !!st.s, wrap: !!st.wrap,
        "al:l": st.al === "l", "al:c": st.al === "c", "al:r": st.al === "r", "va:t": st.va === "t", "va:m": st.va === "m", "va:b": st.va === "b",
        merge: sh.merges.indexOf(rangeText(g)) >= 0, font: st.ff || "M PLUS 1", size: String(st.fs || 11), nf: nfName(st.nf),
        filter: !!sh.filter, grid: sh.grid !== false, frozen: !!sh.freeze, headRow: sh.head === act.r, painter: !!painter
      };
      Chart.TYPES.forEach(function (x) { s["ctype:" + x.value] = !!ch && ch.type === x.value; });
      return s;
    }

    // ---------- 検索と置換 ----------
    function find(q, dir, o) {
      var keys = Object.keys(sh.cells).map(parseRef).filter(Boolean).sort(function (a, b) { return a.r - b.r || a.c - b.c; });
      var norm2 = function (s) { return o.matchCase ? s : s.toLowerCase(); }, needle = norm2(q);
      var hits = keys.filter(function (p) { var txt = show(book.val(si, p.c, p.r), styleAt(p.c, p.r).nf).text; return norm2(txt).indexOf(needle) >= 0 || norm2(sh.cells[ref(p.c, p.r)]).indexOf(needle) >= 0; });
      if (!hits.length) return false;
      var i = -1;
      if (dir > 0) { for (var j = 0; j < hits.length; j++) if (hits[j].r > act.r || (hits[j].r === act.r && hits[j].c > act.c)) { i = j; break; } if (i < 0) i = 0; }
      else { for (var k = hits.length - 1; k >= 0; k--) if (hits[k].r < act.r || (hits[k].r === act.r && hits[k].c < act.c)) { i = k; break; } if (i < 0) i = hits.length - 1; }
      select(hits[i].c, hits[i].r);
      return t("{n} 件中 {i} 件目", { n: hits.length, i: i + 1 });
    }
    function replace(q, rep, all, o) {
      if (readOnly()) return 0;
      var flags = o.matchCase ? "g" : "gi", re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), flags), n = 0;
      var targets = all ? Object.keys(sh.cells) : [ref(act.c, act.r)];
      targets.forEach(function (k) {
        var x = sh.cells[k];
        if (x === undefined || !re.test(x)) return;
        re.lastIndex = 0;
        sh.cells[k] = x.replace(re, function () { n++; return rep; });
        re.lastIndex = 0;
      });
      if (n) { refreshCells(); drawCharts(); placeSel(); }
      return n;
    }

    // ---------- 書き出し・読み込み ----------
    function usedRange(s) {
      var b = Book(data), si2 = data.sheets.indexOf(s), l = b.last(si2), mc = 0, mr = 0;
      s.merges.forEach(function (m) { var g = parseRange(m); if (g) { mr = Math.max(mr, g.r2); mc = Math.max(mc, g.c2); } });
      return { c1: 0, r1: 1, c2: Math.max(l.cols - 1, mc), r2: Math.max(l.rows, mr) };
    }
    function csvOf(s) {
      var b = Book(data), idx = data.sheets.indexOf(s), u = usedRange(s), lines = [];
      for (var r = 1; r <= u.r2; r++) {
        var row = [];
        for (var c = 0; c <= u.c2; c++) {
          var x = show(b.val(idx, c, r), (s.style[ref(c, r)] || EMPTY).nf).text;
          row.push(/[",\r\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x);
        }
        lines.push(row.join(","));
      }
      return "﻿" + lines.join("\r\n") + "\r\n";
    }
    var X = Zip ? Zip.esc : esc;
    function hex6(c) { var m = /^#?([0-9a-f]{6})$/i.exec(c || ""); return m ? "FF" + m[1].toUpperCase() : null; }
    function xlsxFiles() {
      var fonts = ['<font><sz val="11"/><name val="游ゴシック"/><family val="3"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
      var bordersX = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], numFmts = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
      var fontIdx = {}, fillIdx = {}, bdIdx = {}, nfIdx = {}, xfIdx = {};
      var BUILTIN = { "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "0.00E+00": 11, "@": 49 };
      function idx(map, list, key, xml) { if (!(key in map)) { map[key] = list.length; list.push(xml); } return map[key]; }
      function xfOf(st) {
        if (!st || !Object.keys(st).length) return 0;
        var fKey = [st.b, st.i, st.u, st.s, st.fs, st.fg, st.ff].join("|"), f = 0;
        if (st.b || st.i || st.u || st.s || st.fs || st.fg || st.ff) {
          f = idx(fontIdx, fonts, fKey, "<font>" + (st.b ? "<b/>" : "") + (st.i ? "<i/>" : "") + (st.s ? "<strike/>" : "") + (st.u ? "<u/>" : "") + '<sz val="' + (st.fs || 11) + '"/>' +
            (hex6(st.fg) ? '<color rgb="' + hex6(st.fg) + '"/>' : "") + '<name val="' + X(st.ff || "游ゴシック") + '"/><family val="3"/></font>');
        }
        var fl = st.bg && hex6(st.bg) ? idx(fillIdx, fills, st.bg, '<fill><patternFill patternType="solid"><fgColor rgb="' + hex6(st.bg) + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
        var bd = 0;
        if (st.bd && Object.keys(st.bd).length) {
          var side = function (n, w) { return w ? "<" + n + ' style="' + (w > 1 ? "medium" : "thin") + '"><color auto="1"/></' + n + ">" : "<" + n + "/>"; };
          bd = idx(bdIdx, bordersX, JSON.stringify(st.bd), "<border>" + side("left", st.bd.l) + side("right", st.bd.r) + side("top", st.bd.t) + side("bottom", st.bd.b) + "<diagonal/></border>");
        }
        var nf = 0;
        if (st.nf) nf = BUILTIN[st.nf] || (nfIdx[st.nf] || (nfIdx[st.nf] = 164 + numFmts.push('<numFmt numFmtId="' + (164 + numFmts.length) + '" formatCode="' + X(st.nf) + '"/>') - 1));
        var al = st.al || st.va || st.wrap ? '<alignment' + (st.al ? ' horizontal="' + { l: "left", c: "center", r: "right" }[st.al] + '"' : "") + (st.va ? ' vertical="' + { t: "top", m: "center", b: "bottom" }[st.va] + '"' : "") + (st.wrap ? ' wrapText="1"' : "") + "/>" : "";
        var key = [nf, f, fl, bd, al].join("|");
        return idx(xfIdx, xfs, key, '<xf numFmtId="' + nf + '" fontId="' + f + '" fillId="' + fl + '" borderId="' + bd + '" xfId="0"' + (nf ? ' applyNumberFormat="1"' : "") + (f ? ' applyFont="1"' : "") + (fl ? ' applyFill="1"' : "") + (bd ? ' applyBorder="1"' : "") + (al ? ' applyAlignment="1">' + al + "</xf>" : "/>"));
      }
      // [Content_Types].xml を ZIP のいちばん先に入れる（先にないと読めないソフトがある）
      var b = Book(data), files = { "[Content_Types].xml": "" }, sheetsXml = [], rels = [], ctypes = [];
      data.sheets.forEach(function (s, n) {
        var u = usedRange(s), rows = [];
        var byRow = {};
        Object.keys(s.cells).concat(Object.keys(s.style)).forEach(function (k) { var p = parseRef(k); if (p) (byRow[p.r] = byRow[p.r] || {})[p.c] = 1; });
        Object.keys(byRow).map(Number).sort(function (a, c) { return a - c; }).forEach(function (r) {
          var cs = Object.keys(byRow[r]).map(Number).sort(function (a, c) { return a - c; }), cells = [];
          cs.forEach(function (c) {
            var k = ref(c, r), raw = s.cells[k], st = s.style[k];
            // 見出しの行・合計の行は、セルの書式にして書き出す
            if (r === s.head) st = Object.assign({ b: 1, bg: "#d7efd9" }, st || {});
            else if (r === s.total) st = Object.assign({ b: 1, bd: { t: 1 } }, st || {});
            var sx = xfOf(st), sAttr = sx ? ' s="' + sx + '"' : "";
            if (raw === undefined || raw === "") { if (sx) cells.push('<c r="' + k + '"' + sAttr + "/>"); return; }
            var v = b.val(n, c, r);
            if (raw.charAt(0) === "=") {
              var f = "<f>" + X(raw.slice(1)) + "</f>";
              if (typeof v === "number") cells.push('<c r="' + k + '"' + sAttr + ">" + f + "<v>" + v + "</v></c>");
              else if (typeof v === "boolean") cells.push('<c r="' + k + '"' + sAttr + ' t="b">' + f + "<v>" + (v ? 1 : 0) + "</v></c>");
              else if (F.isErr(v)) cells.push('<c r="' + k + '"' + sAttr + ' t="e">' + f + "<v>" + X(v.err) + "</v></c>");
              else cells.push('<c r="' + k + '"' + sAttr + ' t="str">' + f + "<v>" + X(v === null ? "" : String(v)) + "</v></c>");
            } else if (typeof v === "number") cells.push('<c r="' + k + '"' + sAttr + "><v>" + v + "</v></c>");
            else if (typeof v === "boolean") cells.push('<c r="' + k + '"' + sAttr + ' t="b"><v>' + (v ? 1 : 0) + "</v></c>");
            else cells.push('<c r="' + k + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + X(String(v)) + "</t></is></c>");
          });
          rows.push('<row r="' + r + '"' + (s.rh[r] ? ' ht="' + Math.round(s.rh[r] * 0.75 * 10) / 10 + '" customHeight="1"' : "") + ">" + cells.join("") + "</row>");
        });
        var cols = Object.keys(s.cw).map(Number).sort(function (a, c) { return a - c; }).map(function (c) { return '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + Math.max(1, Math.round((s.cw[c] - 5) / 7 * 100) / 100) + '" customWidth="1"/>'; });
        var fz = s.freeze, pane = fz && (fz.r || fz.c) ? '<pane' + (fz.c ? ' xSplit="' + fz.c + '"' : "") + (fz.r ? ' ySplit="' + fz.r + '"' : "") + ' topLeftCell="' + ref(fz.c || 0, (fz.r || 0) + 1) + '" activePane="' + (fz.r && fz.c ? "bottomRight" : fz.r ? "bottomLeft" : "topRight") + '" state="frozen"/>' : "";
        var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<dimension ref="A1:' + ref(Math.max(0, u.c2), Math.max(1, u.r2)) + '"/>' +
          '<sheetViews><sheetView' + (s.grid === false ? ' showGridLines="0"' : "") + (n === data.cur ? ' tabSelected="1"' : "") + ' workbookViewId="0">' + pane + "</sheetView></sheetViews>" +
          '<sheetFormatPr defaultRowHeight="18"/>' + (cols.length ? "<cols>" + cols.join("") + "</cols>" : "") +
          "<sheetData>" + rows.join("") + "</sheetData>" +
          (s.filter ? '<autoFilter ref="' + ref(s.filter.c1, s.filter.r) + ":" + ref(s.filter.c2, Math.max(s.filter.r2 || 0, u.r2)) + '"/>' : "") +
          (s.merges.length ? '<mergeCells count="' + s.merges.length + '">' + s.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join("") + "</mergeCells>" : "") +
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
        files["xl/worksheets/sheet" + (n + 1) + ".xml"] = xml;
        sheetsXml.push('<sheet name="' + X(s.name) + '" sheetId="' + (n + 1) + '" r:id="rId' + (n + 1) + '"/>');
        rels.push('<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (n + 1) + '.xml"/>');
        ctypes.push('<Override PartName="/xl/worksheets/sheet' + (n + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      });
      var N = data.sheets.length;
      files["xl/styles.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join("") + "</numFmts>" : "") +
        '<fonts count="' + fonts.length + '">' + fonts.join("") + '</fonts><fills count="' + fills.length + '">' + fills.join("") + '</fills><borders count="' + bordersX.length + '">' + bordersX.join("") + "</borders>" +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>" +
        '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
      files["xl/workbook.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView activeTab="' + data.cur + '"/></bookViews><sheets>' + sheetsXml.join("") + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
      files["xl/_rels/workbook.xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join("") +
        '<Relationship Id="rId' + (N + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
      files["_rels/.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
      files["docProps/core.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        "<dc:title>" + X(O.file() ? O.file().name : "") + "</dc:title><dc:creator>" + X(O.userName() || "SK Hub Office") + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString().replace(/\.\d+Z$/, "Z") + "</dcterms:created></cp:coreProperties>";
      files["[Content_Types].xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ctypes.join("") + "</Types>";
      return files;
    }
    function exportAs(id) {
      if (id === "csv") {
        var blob = new Blob([csvOf(sh)], { type: "text/csv" });
        return O.saveBlob(O.fileName(data.sheets.length > 1 ? "_" + sh.name + ".csv" : ".csv"), blob, "CSV", "text/csv", ".csv");
      }
      if (id === "xlsx") {
        return Zip.write(xlsxFiles()).then(function (b) {
          return O.saveBlob(O.fileName(".xlsx"), new Blob([b], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), t("Excel ブック"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
        });
      }
    }
    // .xlsx を読む（値・数式・主な書式・列の幅・結合・固定）
    function kids(node, name) { return node ? Array.prototype.slice.call(node.getElementsByTagNameNS("*", name)) : []; }
    function kid(node, name) { return kids(node, name)[0] || null; }
    function textOf(node) {
      // ふりがな（rPh）は入れない
      var s = "";
      kids(node, "t").forEach(function (x) { var p = x.parentNode, skip = false; while (p && p !== node) { if (p.localName === "rPh") { skip = true; break; } p = p.parentNode; } if (!skip) s += x.textContent; });
      return s;
    }
    var BUILTIN_NF = { 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 5: "¥#,##0", 6: "¥#,##0", 7: "¥#,##0.00", 8: "¥#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 14: "yyyy/m/d", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "yyyy/m/d h:mm", 31: "yyyy年m月d日", 32: "h時mm分", 33: "h時mm分ss秒", 37: "#,##0", 38: "#,##0", 39: "#,##0.00", 40: "#,##0.00", 49: "@", 55: "yyyy年m月d日", 56: "m月d日", 57: "yyyy年m月d日", 58: "m月d日" };
    function argb(node) { var c = node && node.getAttribute("rgb"); return c && /^[0-9a-f]{8}$/i.test(c) ? "#" + c.slice(2).toLowerCase() : null; }
    function readXlsx(file) {
      if (!Zip.canRead) return Promise.reject(new Error("no-inflate"));
      return Zip.read(file).then(function (z) {
        var wb = Zip.xml(z["xl/workbook.xml"]);
        if (!wb) throw new Error("not-xlsx");
        var relsDoc = Zip.xml(z["xl/_rels/workbook.xml.rels"]), relMap = {};
        kids(relsDoc, "Relationship").forEach(function (r) { relMap[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var shared = kids(Zip.xml(z["xl/sharedStrings.xml"]), "si").map(textOf);
        var styles = Zip.xml(z["xl/styles.xml"]), nfMap = {}, fonts = [], fillsL = [], bds = [], xfs = [];
        if (styles) {
          kids(kid(styles, "numFmts"), "numFmt").forEach(function (n) { nfMap[n.getAttribute("numFmtId")] = n.getAttribute("formatCode"); });
          kids(kid(styles, "fonts"), "font").forEach(function (f) {
            var o = {};
            if (kid(f, "b") && kid(f, "b").getAttribute("val") !== "0") o.b = 1;
            if (kid(f, "i") && kid(f, "i").getAttribute("val") !== "0") o.i = 1;
            if (kid(f, "u")) o.u = 1;
            if (kid(f, "strike") && kid(f, "strike").getAttribute("val") !== "0") o.s = 1;
            var sz = kid(f, "sz"), col = argb(kid(f, "color")), nm = kid(f, "name");
            if (sz && +sz.getAttribute("val") !== 11) o.fs = +sz.getAttribute("val");
            if (col && col !== "#000000") o.fg = col;
            if (nm && !/^(游ゴシック|Yu Gothic|ＭＳ Ｐゴシック|MS PGothic|Calibri|Aptos Narrow|Aptos)$/i.test(nm.getAttribute("val"))) o.ff = nm.getAttribute("val");
            fonts.push(o);
          });
          kids(kid(styles, "fills"), "fill").forEach(function (f) { var p = kid(f, "patternFill"); fillsL.push(p && p.getAttribute("patternType") === "solid" ? argb(kid(p, "fgColor")) : null); });
          kids(kid(styles, "borders"), "border").forEach(function (b) {
            var o = {};
            [["left", "l"], ["right", "r"], ["top", "t"], ["bottom", "b"]].forEach(function (x) { var n = kid(b, x[0]), s = n && n.getAttribute("style"); if (s) o[x[1]] = /medium|thick|double/.test(s) ? 2 : 1; });
            bds.push(Object.keys(o).length ? o : null);
          });
          kids(kid(styles, "cellXfs"), "xf").forEach(function (x) {
            var o = Object.assign({}, fonts[+x.getAttribute("fontId")] || {}), id = +x.getAttribute("numFmtId"), nf = nfMap[id] || BUILTIN_NF[id];
            if (nf && nf !== "General") o.nf = nf.replace(/\[\$-[0-9A-F]+\]/gi, "").replace(/\\/g, "");
            var bg = fillsL[+x.getAttribute("fillId")];
            if (bg) o.bg = bg;
            var bd = bds[+x.getAttribute("borderId")];
            if (bd) o.bd = bd;
            var al = kid(x, "alignment");
            if (al) {
              var h = al.getAttribute("horizontal"), v = al.getAttribute("vertical");
              if (h === "left" || h === "center" || h === "right") o.al = h.charAt(0);
              if (h === "centerContinuous") o.al = "c";
              if (v === "top" || v === "center") o.va = v === "top" ? "t" : "m";
              if (al.getAttribute("wrapText") === "1") o.wrap = 1;
            }
            xfs.push(o);
          });
        }
        var out = { v: 2, cur: 0, sheets: [] };
        kids(wb, "sheet").forEach(function (sn, n) {
          var rid = sn.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sn.getAttribute("r:id");
          var target = (relMap[rid] || "worksheets/sheet" + (n + 1) + ".xml").replace(/^\/?xl\//, "").replace(/^\//, "");
          var doc = Zip.xml(z["xl/" + target]);
          var s = fixSheet({ name: sn.getAttribute("name") || "Sheet" + (n + 1) }, n), sharedF = {};
          if (!doc) { out.sheets.push(s); return; }
          kids(doc, "c").forEach(function (c) {
            var k = c.getAttribute("r"), p = parseRef(k || "");
            if (!p) return;
            var tp = c.getAttribute("t"), vEl = kid(c, "v"), fEl = kid(c, "f"), st = xfs[+c.getAttribute("s") || 0], raw;
            if (st && Object.keys(st).length) s.style[k] = O.copy(st);
            if (fEl) {
              var ftxt = fEl.textContent, si2 = fEl.getAttribute("si");
              if (fEl.getAttribute("t") === "shared" && si2 !== null) {
                if (ftxt) sharedF[si2] = { f: ftxt, c: p.c, r: p.r };
                else if (sharedF[si2]) ftxt = F.shift(sharedF[si2].f, p.c - sharedF[si2].c, p.r - sharedF[si2].r);
              }
              if (ftxt) { s.cells[k] = "=" + ftxt.replace(/_xlfn\.(_xlws\.)?/g, ""); return; }
            }
            if (tp === "s") raw = shared[+(vEl ? vEl.textContent : -1)];
            else if (tp === "inlineStr") raw = textOf(kid(c, "is"));
            else if (tp === "b") raw = vEl && vEl.textContent === "1" ? "TRUE" : "FALSE";
            else if (tp === "str" || tp === "e") raw = vEl ? vEl.textContent : "";
            else if (vEl) {
              var num = +vEl.textContent, nf = st && st.nf;
              raw = nf && F.isDateFormat(nf) && num >= 1 && num < 2958466 && Math.floor(num) === num ? F.format(num, "yyyy/m/d").text : F.str(num);
            }
            if (raw === undefined || raw === null || raw === "") return;
            if (tp === "s" || tp === "inlineStr" || tp === "str") { if (/^[=+\-']/.test(raw) || F.parseNum(raw) !== null || /^(TRUE|FALSE)$/i.test(raw)) raw = "'" + raw; }
            s.cells[k] = raw;
          });
          kids(doc, "col").forEach(function (c) {
            var w = +c.getAttribute("width"), a = +c.getAttribute("min") - 1, b = Math.min(+c.getAttribute("max") - 1, a + 200);
            if (!w || c.getAttribute("customWidth") === "0") return;
            for (var i = a; i <= b; i++) s.cw[i] = Math.max(16, Math.round(w * 7 + 5));
          });
          kids(doc, "row").forEach(function (r) { var ht = +r.getAttribute("ht"); if (ht && r.getAttribute("customHeight") === "1") s.rh[+r.getAttribute("r")] = Math.round(ht * 4 / 3); });
          kids(doc, "mergeCell").forEach(function (m) { var g = parseRange(m.getAttribute("ref")); if (g) s.merges.push(rangeText(g)); });
          var pane = kid(doc, "pane");
          if (pane && pane.getAttribute("state") === "frozen") s.freeze = { r: +pane.getAttribute("ySplit") || 0, c: +pane.getAttribute("xSplit") || 0 };
          var sv = kid(doc, "sheetView");
          if (sv && sv.getAttribute("showGridLines") === "0") s.grid = false;
          out.sheets.push(s);
        });
        if (!out.sheets.length) out.sheets.push(fixSheet({ name: "Sheet1" }, 0));
        var view = kid(wb, "workbookView");
        if (view && view.getAttribute("activeTab")) out.cur = Math.min(out.sheets.length - 1, +view.getAttribute("activeTab"));
        return out;
      });
    }
    function importFile(file) {
      if (/\.xlsx$/i.test(file.name)) return readXlsx(file);
      return file.arrayBuffer().then(function (ab) {
        var u8 = new Uint8Array(ab), text;
        try { text = new TextDecoder("utf-8", { fatal: true }).decode(u8); } catch (e) { text = new TextDecoder("shift_jis").decode(u8); }
        var sep = /\.tsv$/i.test(file.name) || (text.indexOf("\t") >= 0 && text.split("\n")[0].indexOf(",") < 0) ? "\t" : ",";
        var rows = parseDelimited(text, sep), s = fixSheet({ name: file.name.replace(/\.[^.]+$/, "").slice(0, 31) || "Sheet1" }, 0);
        rows.forEach(function (row, r) { row.forEach(function (v, c) { if (v !== "") s.cells[ref(c, r + 1)] = v; }); });
        Object.keys(s.cells).forEach(function (k) { var nf = autoNf(s.cells[k]); if (nf) s.style[k] = { nf: nf }; });
        return { v: 2, cur: 0, sheets: [s] };
      });
    }

    // ---------- 印刷（使っているところだけを、見出しなしの表にする） ----------
    function beforePrint() {
      endEdit(true);
      var u = usedRange(sh), box = document.createElement("div");
      box.className = "sheet-print";
      box.id = "sheet-print";
      var save = { fzTop: fzTop, fzLeft: fzLeft };
      fzTop = {}; fzLeft = {};
      var h = '<table class="grid' + (sh.grid === false ? " no-grid" : "") + '"><colgroup>';
      for (var c = 0; c <= u.c2; c++) h += '<col style="width:' + colW(c) + 'px">';
      h += "</colgroup><tbody>";
      for (var r = 1; r <= u.r2; r++) {
        if (hiddenRows[r]) continue;
        h += '<tr class="' + (r === sh.head ? "head" : r === sh.total ? "total" : "") + '"' + (sh.rh[r] ? ' style="height:' + sh.rh[r] + 'px"' : "") + ">";
        for (c = 0; c <= u.c2; c++) {
          var k = ref(c, r), m = mtop[k], p = cellParts(c, r);
          if (covered[k]) continue;
          h += "<td" + (m ? (m.cs > 1 ? ' colspan="' + m.cs + '"' : "") + (m.rs > 1 ? ' rowspan="' + m.rs + '"' : "") : "") + (p.cls ? ' class="' + p.cls + '"' : "") + (p.css ? ' style="' + p.css + '"' : "") + ">" + esc(p.text) + "</td>";
        }
        h += "</tr>";
      }
      h += "</tbody></table>";
      fzTop = save.fzTop; fzLeft = save.fzLeft;
      box.innerHTML = h + sh.charts.map(function (ch) { return '<div class="sheet-print__chart">' + chartSvg(ch, true) + "</div>"; }).join("");
      $("canvas").appendChild(box);
      var wide = (u.c2 + 1) > 7;
      return "@page { size: A4 " + (wide ? "landscape" : "portrait") + "; margin: 12mm; }";
    }
    function afterPrint() { var p = el("sheet-print"); if (p) p.remove(); }

    // ---------- 小さなプレビュー ----------
    function thumb(d0, blank) {
      var d = upgrade(O.copy(d0)), b = Book(d), s = d.sheets[d.cur], h = '<div class="th th--sheet" aria-hidden="true"><table class="grid"><tbody>';
      for (var r = 1; r <= 8; r++) {
        h += '<tr class="' + (r === s.head ? "head" : r === s.total ? "total" : "") + '">';
        for (var c = 0; c < 5; c++) {
          var v = b.val(d.cur, c, r), x = show(v, (s.style[ref(c, r)] || EMPTY).nf);
          h += "<td" + (x.num ? ' class="n"' : "") + ">" + esc(x.text) + "</td>";
        }
        h += "</tr>";
      }
      return h + "</tbody></table></div>";
    }

    return {
      ribbon: ribbon,
      upgrade: upgrade,
      thumb: thumb,
      render: render,
      command: command,
      state: state,
      zoom: function (z) { zoomZ = z; applyZoom(); },
      fitZoom: function () {
        var sc = el("grid-scroll"), u = usedRange(sh), w = HEADW;
        for (var c = 0; c <= Math.max(u.c2, 5); c++) w += colW(c);
        return sc ? Math.max(0.25, Math.min(3, (sc.clientWidth - 24) / w)) : 1;
      },
      focus: focusGrid,
      open: function () { nRows = 100; nCols = 26; },
      close: function () { endEdit(false); clip = null; painter = null; },
      beforeSave: function () {},
      beforeUndo: function () { if (editing) endEdit(false); },
      setRead: function () { endEdit(false); stat(); },
      selectedText: function () { return show(book.val(si, act.c, act.r), styleAt(act.c, act.r).nf).text; },
      find: find,
      replace: replace,
      findClose: function () { focusGrid(); },
      clip: function (cmd) {
        focusGrid();
        if (cmd === "paste") {
          if (navigator.clipboard && navigator.clipboard.readText) {
            navigator.clipboard.readText().then(function (text) { pasteText(text, ""); }, function () { if (clip) pasteClip(clip); else O.notify(t("Ctrl+V で貼り付けてください")); });
          } else if (clip) pasteClip(clip);
          return;
        }
        clip = clipData(range(), cmd === "cut");
        copyText(clip.text);
        placeSel();
        O.notify(cmd === "cut" ? t("切り取りました。貼り付けるところを選んでください") : t("コピーしました"));
      },
      key: function () {},
      beforePrint: beforePrint,
      afterPrint: afterPrint,
      exports: [
        { id: "xlsx", icon: "table_view", label: t("Excel ブック（.xlsx）"), desc: t("Microsoft Excel・LibreOffice・Google スプレッドシートで開けます") },
        { id: "csv", icon: "csv", label: "CSV（.csv）", desc: t("いまのシートの値を、カンマで区切った文字にします") }
      ],
      exportAs: exportAs,
      imports: { label: ".xlsx・.csv", accept: ".xlsx,.csv,.tsv,.txt" },
      importFile: importFile,
      props: function () {
        var n = 0;
        data.sheets.forEach(function (s) { n += Object.keys(s.cells).length; });
        return [[t("シート"), String(data.sheets.length)], [t("入力したセル"), String(n)]];
      },
      keys: [["F2", t("セルを編集")], ["Enter / Tab", t("入力して下・右へ")], ["Alt+Enter", t("セルの中で改行")], ["Ctrl+Enter", t("選んだセルすべてに入力")],
        ["Ctrl+D / Ctrl+R", t("下・右へコピー")], ["Ctrl+B / I / U", t("太字・斜体・下線")], ["Ctrl+;", t("今日の日付")], ["Alt+=", t("オート SUM")],
        ["Ctrl+矢印", t("データの端へ移動")], ["Shift+矢印", t("選ぶ範囲を広げる")], ["Ctrl+A", t("表全体を選ぶ")], ["Ctrl+PageUp / PageDown", t("シートの切りかえ")]]
    };
  };
})();
