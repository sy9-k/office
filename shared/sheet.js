// SK Hub Office の Sheet（表計算）。window.OfficeApps.sheet に登録し、/office/shared/office.js が呼ぶ（O は office.js が渡す道具）
//   ファイルの中身（data）: { v: 2, cur: 開いているシートの番号, sheets: [シート] }
//     シート: { name, cells: { "A1": "入れた文字（= で始まると数式）" }, style: { "A1": 書式 }, cw: { 列の番号: 幅 px }, rh: { 行の番号: 高さ px },
//              merges: ["A1:C1"], freeze: { r: 固定する行の数, c: 固定する列の数 }, head: 見出しの行, total: 合計の行, grid: false（目盛り線を消す）,
//              filter: { r: 見出しの行, c1, c2, hide: { 列の番号: [隠す文字] } }, cf: [条件付き書式], charts: [{ id, type, ref, title, x, y, w, h }],
//              dv: [入力規則], notes: { "A1": { text, author, ts }（セルのコメント） } }
//     書式: { b, i, u, s（取り消し線）, ff（フォント）, fs（大きさ pt）, fg（文字の色）, bg（塗りつぶし）, al: "l|c|r", va: "t|m|b", wrap, nf（表示形式）, bd: { t, r, b, l } }
//     条件付き書式: { ref: "B2:B9", type: "gt|lt|eq|text|dup|bar|scale", v, color }
//     入力規則: { ref: "B2:B9", type: "list|whole|decimal|date|textLength|custom", op: "between|notBetween|equal|notEqual|greaterThan|lessThan|greaterThanOrEqual|lessThanOrEqual",
//                 f1（値・最小値・リストの元の値）, f2（最大値）, allowBlank, promptTitle, prompt, errTitle, err, style: "stop|warning|information" }
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
      s.dv = s.dv || []; s.notes = s.notes || {};
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

    // ---------- 入力規則（data.dv） ----------
    //   { ref, type: "list|whole|decimal|date|textLength|custom", op, f1, f2, allowBlank,
    //     promptTitle, prompt, errTitle, err, style: "stop|warning|information" }
    var DV_TYPES = [
      { v: "list", label: t("リスト") }, { v: "whole", label: t("整数") }, { v: "decimal", label: t("小数") },
      { v: "date", label: t("日付") }, { v: "textLength", label: t("文字の長さ") }, { v: "custom", label: t("ユーザー設定（数式）") }
    ];
    var DV_OPS = [
      { v: "between", label: t("次の値の間") }, { v: "notBetween", label: t("次の値の間以外") },
      { v: "equal", label: t("次の値に等しい") }, { v: "notEqual", label: t("次の値に等しくない") },
      { v: "greaterThan", label: t("次の値より大きい") }, { v: "lessThan", label: t("次の値より小さい") },
      { v: "greaterThanOrEqual", label: t("次の値以上") }, { v: "lessThanOrEqual", label: t("次の値以下") }
    ];
    var DV_STYLES = [{ v: "stop", label: t("停止") }, { v: "warning", label: t("警告") }, { v: "information", label: t("情報") }];
    function dvEmpty() { return { ref: "", type: "list", op: "between", f1: "", f2: "", allowBlank: true, promptTitle: "", prompt: "", errTitle: "", err: "", style: "stop" }; }
    function dvRuleAt(c, r) {
      var list = (sh && sh.dv) || [];
      for (var i = 0; i < list.length; i++) { var g = parseRange(list[i].ref); if (g && inRange(g, c, r)) return list[i]; }
      return null;
    }
    // リストの候補（カンマ区切り、または A1:A9 のような範囲）
    function dvListValues(rule) {
      var src = String(rule.f1 || "").trim().replace(/^=/, ""), g = parseRange(src), out = [];
      if (g) {
        for (var r = g.r1; r <= g.r2; r++) for (var c = g.c1; c <= g.c2; c++) {
          var x = show(book.val(si, c, r), styleAt(c, r).nf).text;
          if (x !== "") out.push(x);
        }
        return out;
      }
      return src.split(/[,、;]/).map(function (s) { return s.trim(); }).filter(function (s) { return s !== ""; });
    }
    // 規則に合わない値なら true。数式は検証しない（計算結果を信じる）
    function dvFail(rule, c, r, raw) {
      if (!rule) return false;
      raw = raw === undefined ? rawAt(c, r) : String(raw);
      if (raw === "") return rule.allowBlank === false;
      if (raw.charAt(0) === "=") return false;
      if (rule.type === "list") {
        var low = raw.toLowerCase();
        return dvListValues(rule).map(function (x) { return String(x).toLowerCase(); }).indexOf(low) < 0;
      }
      if (rule.type === "custom") {
        var g = parseRange(rule.ref) || { c1: c, r1: r };
        try {
          var src = F.shift(String(rule.f1 || "").trim().replace(/^=/, ""), c - g.c1, r - g.r1);
          if (!src) return false;
          var v = F.evaluate(F.parse(src), { get: book.val, sheet: book.sheetIndex, here: { s: si, c: c, r: r }, last: book.last });
          return !v || F.isErr(v);
        } catch (e) { return false; }
      }
      var val;
      if (rule.type === "date") { val = F.parseDate(raw); if (val === null) val = F.parseNum(raw); }
      else if (rule.type === "textLength") val = raw.length;
      else val = F.parseNum(raw);
      if (val === null || (typeof val === "number" && !isFinite(val))) return true;
      function lim(key) {
        var s = String(rule[key] === undefined || rule[key] === null ? "" : rule[key]).trim().replace(/^=/, "");
        if (!s) return null;
        if (rule.type === "date") { var d = F.parseDate(s); return d === null ? F.parseNum(s) : d; }
        return F.parseNum(s);
      }
      var a = lim("f1"), b = lim("f2"), hit;
      switch (rule.op) {
        case "between": if (a === null || b === null) return false; hit = val >= a && val <= b; break;
        case "notBetween": if (a === null || b === null) return false; hit = val >= a && val <= b; break;
        case "equal": if (a === null) return false; hit = val === a; break;
        case "notEqual": if (a === null) return false; hit = val === a; break;
        case "greaterThan": if (a === null) return false; hit = val > a; break;
        case "lessThan": if (a === null) return false; hit = val < a; break;
        case "greaterThanOrEqual": if (a === null) return false; hit = val >= a; break;
        case "lessThanOrEqual": if (a === null) return false; hit = val <= a; break;
        default: return false;
      }
      if (rule.op === "notBetween" || rule.op === "notEqual") return hit;
      return !hit;
    }

    // ---------- セルのコメント（data.notes） ----------
    function noteAt(c, r) { return (sh && sh.notes && sh.notes[ref(c, r)]) || null; }
    function noteMark(c, r) {
      var n = noteAt(c, r);
      if (!n || !n.text) return "";
      var tip = (n.author ? n.author + "：" : "") + n.text;
      return '<span class="cmt-mark" title="' + esc(tip) + '" aria-label="' + esc(t("コメント")) + '"></span>';
    }

    // ---------- いまの状態 ----------
    var data = null, book = null, si = 0, sh = null;
    // act: 選んでいるセル（入力するところ）。anc: 範囲のもう一方の角（Shift+矢印・ドラッグで動く）
    var act = { c: 0, r: 1 }, anc = { c: 0, r: 1 };
    var nRows = 100, nCols = 26;
    var editing = false, editMode = "enter", refIns = null;
    var chartSel = null, clip = null, painter = null, zoomZ = 1;
    var mtop = {}, covered = {}, hiddenRows = {}, fzTop = {}, fzLeft = {}, cfCache = null;
    // 見えている行だけを作る（win: 作っている行。rowH: 作ったときに測った高さ。defH: ふつうの行の高さ。posCache: 各行の上の位置）
    var win = { r0: 1, r1: 0 }, rowH = {}, defH = 0, posCache = null, BUF = 40;
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
      if (dvFail(dvRuleAt(c, r), c, r, rawAt(c, r))) cls.push("is-invalid");
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
        (p.cls ? ' class="' + p.cls + '"' : "") + (p.css ? ' style="' + p.css + '"' : "") + ">" + esc(p.text) + noteMark(c, r) + filterBtn(c, r) + "</td>";
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
      Object.keys(sh.notes || {}).forEach(function (k) { var p = parseRef(k); if (p) { mr = Math.max(mr, p.r); mc = Math.max(mc, p.c); } });
      (sh.dv || []).forEach(function (rule) { var g = parseRange(rule.ref); if (g) { mr = Math.max(mr, g.r2); mc = Math.max(mc, g.c2); } });
      nRows = Math.max(nRows, 60, last.rows + 40, mr + 20, act.r + 20);
      nCols = Math.max(nCols, 26, last.cols + 6, mc + 4, act.c + 4);
    }

    // ---------- 見えている行だけを作る ----------
    //   何万行もあると、全部の行を作ると重いので、見えているところと上下 BUF 行だけを作る。その上と下は、高さだけの空の行（tr.vpad）で場所をとる。
    //   行の位置は positions()（隠した行は 0、作ったことのある行は測った高さ、ほかは sh.rh かふつうの高さ）。固定する行（1〜frozenRows() 行目）はいつも作る
    function frozenRows() { return sh.freeze && sh.freeze.r ? Math.min(sh.freeze.r, nRows) : 0; }
    function hOf(r) { return hiddenRows[r] ? 0 : rowH[r] || sh.rh[r] || defH || 29; }
    function positions() {
      if (posCache && posCache.length === nRows + 2) return posCache;
      var p = new Float64Array(nRows + 2);
      for (var r = 1; r <= nRows; r++) p[r + 1] = p[r] + hOf(r);
      return (posCache = p);
    }
    function rowAt(y) {
      var p = positions(), lo = 1, hi = nRows;
      if (y <= 0) return 1;
      while (lo < hi) { var mid = (lo + hi + 1) >> 1; if (p[mid] <= y) lo = mid; else hi = mid - 1; }
      return lo;
    }
    function padHtml(kind, h) { return '<tr class="vpad" data-pad="' + kind + '" style="height:' + Math.max(0, h) + 'px"><td colspan="' + (nCols + 1) + '"></td></tr>'; }
    function updatePads() {
      var tbl = el("grid"), p = positions(), fr = frozenRows();
      if (!tbl) return;
      var top = tbl.querySelector('tr[data-pad="top"]'), bot = tbl.querySelector('tr[data-pad="bot"]');
      if (top) top.style.height = Math.max(0, p[Math.max(win.r0, fr + 1)] - p[fr + 1]) + "px";
      if (bot) bot.style.height = Math.max(0, p[nRows + 1] - p[Math.min(nRows, Math.max(win.r1, fr)) + 1]) + "px";
    }
    // 作った行の高さを測る（折り返しで高くなった行など）。変わっていたら、位置と空の行の高さを直す
    function measureRows(force) {
      var tb = el("grid") && el("grid").tBodies[0], diff = !!force, minH = 0;
      if (!tb) return;
      Array.prototype.forEach.call(tb.rows, function (tr) {
        if (tr.classList.contains("vpad")) return;
        var r = +tr.getAttribute("data-r"), h = tr.offsetHeight;
        if (!h || hiddenRows[r]) return;
        if (!sh.rh[r] && (!minH || h < minH)) minH = h;
        if (h !== hOf(r)) { rowH[r] = h; diff = true; }
      });
      if (!defH && minH) { defH = minH; diff = true; }
      if (diff) { posCache = null; updatePads(); }
    }
    // 行 r の上の位置（.grid-wrap の中。ズームする前の px）
    function rowY(r) { var tbl = el("grid"); return tbl.offsetTop + tbl.tBodies[0].offsetTop + positions()[Math.max(1, Math.min(r, nRows + 1))]; }
    function renderWindow(force) {
      var tbl = el("grid"), sc = el("grid-scroll"), tb = tbl && tbl.tBodies[0];
      if (!tb || !sc) return;
      var z = zoomZ || 1, fr = frozenRows(), p = positions(), base = tbl.offsetTop + tb.offsetTop;
      var y0 = sc.scrollTop / z - base, y1 = y0 + sc.clientHeight / z;
      var a = Math.max(fr + 1, rowAt(y0)), b = Math.min(nRows, rowAt(y1));
      if (!force && win.r1 >= win.r0 && a >= win.r0 && b <= win.r1) return;
      var r0 = Math.max(fr + 1, a - BUF), r1 = Math.min(nRows, Math.max(b, a) + BUF);
      // 結合したセルが窓の端にかかるときは、結合が全部入るように広げる（左上のセルがないと、結合が描けない）
      sh.merges.forEach(function (m) { var g = parseRange(m); if (g && g.r1 > fr && g.r2 >= r0 && g.r1 <= r1) { r0 = Math.min(r0, g.r1); r1 = Math.max(r1, g.r2); } });
      win = { r0: r0, r1: r1 };
      Array.prototype.slice.call(tb.rows).forEach(function (tr) { if (tr.classList.contains("vpad") || +tr.getAttribute("data-r") > fr) tr.remove(); });
      var h = padHtml("top", p[r0] - p[fr + 1]);
      for (var r = r0; r <= r1; r++) h += rowHtml(r);
      tb.insertAdjacentHTML("beforeend", h + padHtml("bot", p[nRows + 1] - p[Math.max(r1, fr) + 1]));
      measureRows();
      placeSel();
    }
    // 行 r を作って見えるところに出す（縦だけ。横は scrollToCell）
    function revealRow(r) {
      var sc = el("grid-scroll"), tbl = el("grid");
      if (!sc || !tbl) return;
      var z = zoomZ || 1, fr = frozenRows(), p = positions();
      var headH = tbl.tHead ? tbl.tHead.offsetHeight : 29, top = rowY(r) * z, bottom = rowY(r + 1) * z, topLimit = sc.scrollTop + (headH + (r > fr ? p[fr + 1] : 0)) * z;
      if (top < topLimit) sc.scrollTop -= topLimit - top;
      else if (bottom > sc.scrollTop + sc.clientHeight) sc.scrollTop = bottom - sc.clientHeight;
      renderWindow();
    }
    function drawGrid(keepScroll) {
      var scroller = el("grid-scroll");
      if (!scroller) return;
      var sx = scroller.scrollLeft, sy = scroller.scrollTop;
      book.reset(); cfCache = null;
      indexMerges(); indexHidden();
      extent();
      fzTop = {}; fzLeft = {}; rowH = {}; posCache = null; win = { r0: 1, r1: 0 };
      var tbl = el("grid"), h = '<colgroup><col style="width:' + HEADW + 'px">';
      var total = HEADW;
      for (var c = 0; c < nCols; c++) { h += '<col style="width:' + colW(c) + 'px">'; total += colW(c); }
      h += '</colgroup><thead><tr><th class="corner" data-corner title="' + esc(t("すべて選択")) + '"></th>';
      for (c = 0; c < nCols; c++) h += '<th data-col="' + c + '">' + F.colName(c) + '<span class="col-rs" data-rs="' + c + '"></span></th>';
      h += "</tr></thead><tbody>";
      // ここでは固定する行だけを作る。ほかの行は renderWindow が、見えているところだけ作る
      for (var r = 1; r <= frozenRows(); r++) h += rowHtml(r);
      tbl.innerHTML = h + padHtml("top", 0) + padHtml("bot", 0) + "</tbody>";
      tbl.style.width = total + "px";
      tbl.classList.toggle("no-grid", sh.grid === false);
      // 固定する行・列は、置いた位置を測ってから、もう一度その位置に貼りつける
      if (measureFreeze()) refreshCells(true);
      measureRows(true);
      if (keepScroll) { scroller.scrollLeft = sx; scroller.scrollTop = sy; }
      renderWindow(true);
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
        var c = +td.getAttribute("data-c"), r = +td.getAttribute("data-r"), p = partsFor(c, r), html = esc(p.text) + noteMark(c, r) + filterBtn(c, r);
        if (td.innerHTML !== html) td.innerHTML = html;
        if ((td.getAttribute("class") || "") !== p.cls) { if (p.cls) td.setAttribute("class", p.cls); else td.removeAttribute("class"); }
        if ((td.getAttribute("style") || "") !== p.css) { if (p.css) td.setAttribute("style", p.css); else td.removeAttribute("style"); }
      });
      // 中身が変わると、折り返しで行の高さが変わることがある。フィルターで隠す行が変わったら、位置を計算しなおす
      if (rowsChanged) posCache = null;
      measureRows(rowsChanged);
    }
    // 下までスクロールしたら、行を足す（行の数を増やして、下の空の行を高くするだけ）
    function moreRows(n) {
      if (!el("grid")) return;
      nRows += n;
      posCache = null;
      updatePads();
      renderWindow();
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
      if (!wrap) return null;
      var a = tdAt(g.c1, g.r1), b = tdAt(g.c2, g.r2);
      // 範囲の端の行をまだ作っていないときは、行の位置（positions）と列の見出しから計算する
      if (!a || !b) {
        var tbl = el("grid"), heads = tbl.tHead ? tbl.tHead.rows[0].cells : [], ca = heads[g.c1 + 1], cb = heads[Math.min(g.c2, nCols - 1) + 1];
        if (!ca || !cb) return null;
        var y = rowY(g.r1);
        return { x: tbl.offsetLeft + ca.offsetLeft, y: y, w: cb.offsetLeft + cb.offsetWidth - ca.offsetLeft, h: rowY(g.r2 + 1) - y };
      }
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
      if (!td && sc) { revealRow(mergeAt(c, r) ? mergeAt(c, r).r1 : r); td = tdAt(c, r); }
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
      if (commit) {
        if (commitCell(act.c, act.r, v) === false) {
          // 入力規則で止めたときは、入れた値を数式バーに残す（直しやすいように）
          el("fx").value = v;
          placeEditor();
          return true;
        }
      }
      el("fx").value = rawAt(act.c, act.r);
      placeEditor();
      return true;
    }
    function commitCell(c, r, v) {
      if (readOnly()) return false;
      v = String(v).replace(/^＝/, "=");
      // 数式の関数の名前は大文字に
      if (v.charAt(0) === "=") v = "=" + F.tokenize(v.slice(1)).map(function (x) { return x.t === "fn" ? x.s.toUpperCase() : x.s; }).join("");
      var k = ref(c, r), old = sh.cells[k] || "";
      if (v === old) return true;
      // 入力規則。停止なら入れない（警告・情報なら入れて知らせる）
      var rule = dvRuleAt(c, r);
      if (v !== "" && dvFail(rule, c, r, v)) {
        if ((rule.style || "stop") === "stop") {
          O.notify(rule.err || t("このセルには、決めた条件の値だけを入れられます"));
          return false;
        }
        O.notify(rule.err || t("入力規則に合っていない値を入れました"));
      }
      if (v === "") delete sh.cells[k]; else sh.cells[k] = v;
      if (v && !(sh.style[k] && sh.style[k].nf)) { var nf = autoNf(v); if (nf) { sh.style[k] = Object.assign({}, sh.style[k] || {}, { nf: nf }); } }
      changed();
      return true;
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
    // ---------- デザイン: テーブルのスタイル（見出しの行・縞・罫線）とセルのスタイル。どちらもセルの書式にするので、.xlsx にもそのまま書き出される ----------
    //   head・fg: 見出しの行の色と文字の色。band: 1 行おきの縞の色。bd: "all" 全部の罫線・"h" 横線だけ・"" なし
    var TSTYLES = [
      { id: "plain", label: t("シンプル"), group: "light", head: "#eceff1", fg: "#111111", bd: "h" },
      { id: "slate", label: t("スレート"), group: "light", head: "#475569", fg: "#ffffff", band: "#e2e8f0", bd: "" },
      { id: "mono", label: t("モノクロ"), group: "light", head: "#111111", fg: "#ffffff", band: "#f2f2f2", bd: "all" },
      { id: "matcha", label: t("抹茶"), group: "sk", head: "#16a34a", fg: "#ffffff", band: "#e9f6ea", bd: "" },
      { id: "ocean", label: t("海"), group: "sk", head: "#2563eb", fg: "#ffffff", band: "#eaf1ff", bd: "" },
      { id: "mikan", label: t("みかん"), group: "sk", head: "#ea580c", fg: "#ffffff", band: "#fff1ec", bd: "" },
      { id: "sakura", label: t("さくら"), group: "sk", head: "#db2777", fg: "#ffffff", band: "#fdeef4", bd: "" },
      { id: "navy", label: t("ネイビー"), group: "biz", head: "#0b1f3a", fg: "#ffffff", band: "#eef2f8", bd: "h" },
      { id: "gold", label: t("ゴールド"), group: "biz", head: "#b8862b", fg: "#ffffff", band: "#f6f1e4", bd: "h" },
      { id: "indigo", label: t("インディゴ"), group: "biz", head: "#4338ca", fg: "#ffffff", band: "#e8e8ff", bd: "" },
      { id: "dark", label: t("ダーク"), group: "biz", head: "#1f2937", fg: "#ffffff", band: "#f3f4f6", bd: "all" },
      { id: "forest", label: t("森"), group: "color", head: "#2f6f4e", fg: "#ffffff", band: "#e5f0e8", bd: "" },
      { id: "mint", label: t("ミント"), group: "color", head: "#0d9488", fg: "#ffffff", band: "#dcf5f1", bd: "" },
      { id: "sunset", label: t("夕焼け"), group: "color", head: "#e5484d", fg: "#ffffff", band: "#ffece6", bd: "" },
      { id: "lavender", label: t("ラベンダー"), group: "color", head: "#7c5cd6", fg: "#ffffff", band: "#f0eaff", bd: "" },
      { id: "pastel", label: t("パステル"), group: "color", head: "#e46a9a", fg: "#ffffff", band: "#ffeef4", bd: "" },
      { id: "washi", label: t("和紙"), group: "color", head: "#b3261e", fg: "#ffffff", band: "#efe6d3", bd: "h" }
    ];
    var TSTYLE_GROUPS = [["light", t("淡い・シンプル")], ["sk", t("SK のデザイン")], ["biz", t("ビジネス")], ["color", t("いろいろな色")]];
    var CSTYLES = [
      // 見出し系は、文字色を決めない（ダークモードでも、テーマの文字色で読める）
      { id: "title", label: t("タイトル"), group: "head", st: { b: 1, fs: 18 } },
      { id: "h1", label: t("見出し 1"), group: "head", st: { b: 1, fs: 15, bd: { b: 2 } } },
      { id: "h2", label: t("見出し 2"), group: "head", st: { b: 1, fs: 13, bd: { b: 1 } } },
      { id: "h3", label: t("見出し 3"), group: "head", st: { b: 1, i: 1 } },
      { id: "good", label: t("良い"), group: "status", st: { bg: "#c6efce", fg: "#006100" } },
      { id: "bad", label: t("悪い"), group: "status", st: { bg: "#ffc7ce", fg: "#9c0006" } },
      { id: "neutral", label: t("どちらでもない"), group: "status", st: { bg: "#ffeb9c", fg: "#9c5700" } },
      { id: "warn", label: t("警告"), group: "status", st: { fg: "#ff0000", b: 1 } },
      { id: "note", label: t("メモ"), group: "status", st: { bg: "#fff2cc", fg: "#333333", bd: { t: 1, r: 1, b: 1, l: 1 } } },
      { id: "input", label: t("入力"), group: "data", st: { bg: "#ffcc99", fg: "#3f3f76", bd: { t: 1, r: 1, b: 1, l: 1 } } },
      { id: "output", label: t("出力"), group: "data", st: { bg: "#f2f2f2", fg: "#3f3f3f", b: 1, bd: { t: 1, r: 1, b: 1, l: 1 } } },
      { id: "calc", label: t("計算"), group: "data", st: { bg: "#f2f2f2", fg: "#fa7d00", b: 1 } },
      { id: "total", label: t("集計"), group: "data", st: { b: 1, bd: { t: 1, b: 2 } } },
      { id: "key", label: t("強調"), group: "data", st: { bg: "#dbe7ff", fg: "#1d4ed8", b: 1 } }
    ];
    var CSTYLE_GROUPS = [["head", t("タイトルと見出し")], ["status", t("良い・悪い・メモ")], ["data", t("データと計算")]];
    var STYLE_KEYS = ["b", "i", "u", "s", "fs", "fg", "bg", "bd"];
    function applyTableStyle(id) {
      var g = workRange(), ts = TSTYLES.filter(function (x) { return x.id === id; })[0];
      if (!ts) { setStyle(g, { bg: null, fg: null, bd: null }); changed(); return; }
      if (isOne(g) && !filled(g.c1, g.r1)) { O.notify(t("スタイルを付けるデータを選んでください")); return; }
      var bd = ts.bd === "all" ? { t: 1, r: 1, b: 1, l: 1 } : ts.bd === "h" ? { t: 1, b: 1 } : null;
      for (var r = g.r1; r <= g.r2; r++) {
        var row = { c1: g.c1, r1: r, c2: g.c2, r2: r };
        if (r === g.r1) setStyle(row, { b: 1, bg: ts.head, fg: ts.fg, bd: bd });
        else {
          // 縞の行は薄い色なので、ダークモードでも読めるよう、濃い文字色も付ける
          var banded = ts.band && (r - g.r1) % 2 === 0;
          setStyle(row, { bg: banded ? ts.band : null, fg: banded ? "#1f2328" : null, bd: bd });
        }
      }
      changed();
      O.notify(t("「{name}」を付けました", { name: ts.label }));
    }
    function applyCellStyle(id) {
      var cs = CSTYLES.filter(function (x) { return x.id === id; })[0], patch = {};
      STYLE_KEYS.forEach(function (k) { patch[k] = null; });
      if (cs) Object.keys(cs.st).forEach(function (k) { patch[k] = cs.st[k]; });
      setStyle(range(), patch);
      changed();
    }
    function styleMenu(kind, at) {
      var items = [];
      if (kind === "table") {
        items.push({ value: "", label: t("スタイルなし（色と罫線をクリア）"), icon: "format_color_reset" });
        TSTYLE_GROUPS.forEach(function (gr) {
          items.push({ head: gr[1] });
          TSTYLES.filter(function (x) { return x.group === gr[0]; }).forEach(function (x) { items.push({ value: x.id, label: x.label, color: x.head }); });
        });
      } else {
        items.push({ value: "", label: t("標準（書式をクリア）"), icon: "format_clear" });
        CSTYLE_GROUPS.forEach(function (gr) {
          items.push({ head: gr[1] });
          CSTYLES.filter(function (x) { return x.group === gr[0]; }).forEach(function (x) { items.push({ value: x.id, label: x.label, color: x.st.bg || x.st.fg || "#6b7280" }); });
        });
      }
      return O.menu(at, items).then(function (v) { if (v === null || v === undefined) return; if (kind === "table") applyTableStyle(v); else applyCellStyle(v); });
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
      sh.notes = move(sh.notes);
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
      sh.dv = sh.dv.map(function (rule) {
        var r2 = fixRange(rule.ref); if (!r2) return null;
        rule.ref = r2;
        if (rule.type === "list" && parseRange(String(rule.f1 || "").replace(/^=/, ""))) { var f2 = fixRange(String(rule.f1).replace(/^=/, "")); if (f2) rule.f1 = f2; }
        return rule;
      }).filter(Boolean);
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

    // ---------- 入力規則 ----------
    function dvDialog(at) {
      var cur = dvRuleAt(act.c, act.r), r = Object.assign(dvEmpty(), cur || { ref: rangeText(range()) });
      var body = document.createElement("div"), list = function (a) { return a.map(function (x) { return '<option value="' + x.v + '">' + esc(x.label) + "</option>"; }).join(""); };
      body.innerHTML = '<div class="dlg-grid">' +
        '<label for="dv-ref">' + esc(t("範囲")) + '</label><input id="dv-ref" type="text" value="' + esc(r.ref) + '">' +
        '<label for="dv-type">' + esc(t("種類")) + '</label><select id="dv-type">' + list(DV_TYPES) + "</select>" +
        '<label for="dv-op" data-dv="op">' + esc(t("条件")) + '</label><select id="dv-op" data-dv="op">' + list(DV_OPS) + "</select>" +
        '<label for="dv-f1" data-dv="f1">' + esc(t("値")) + '</label><input id="dv-f1" type="text" value="' + esc(r.f1 || "") + '">' +
        '<label for="dv-f2" data-dv="f2">' + esc(t("最大値")) + '</label><input id="dv-f2" type="text" value="' + esc(r.f2 || "") + '">' +
        "<label>" + esc(t("空白を許可")) + '</label><input id="dv-blank" type="checkbox"' + (r.allowBlank !== false ? " checked" : "") + ">" +
        '<label for="dv-prompt-title">' + esc(t("入力時の見出し")) + '</label><input id="dv-prompt-title" type="text" maxlength="32" value="' + esc(r.promptTitle || "") + '">' +
        '<label for="dv-prompt">' + esc(t("入力時のメッセージ")) + '</label><input id="dv-prompt" type="text" maxlength="255" value="' + esc(r.prompt || "") + '">' +
        '<label for="dv-err-title">' + esc(t("エラーの見出し")) + '</label><input id="dv-err-title" type="text" maxlength="32" value="' + esc(r.errTitle || "") + '">' +
        '<label for="dv-err">' + esc(t("エラー メッセージ")) + '</label><input id="dv-err" type="text" maxlength="255" value="' + esc(r.err || "") + '">' +
        '<label for="dv-style">' + esc(t("エラーの種類")) + '</label><select id="dv-style">' + list(DV_STYLES) + "</select>" +
        '</div><p class="dlg-note">' + esc(t("決めた条件に合わない値を入れたときの動きを選びます。停止は入れられません。警告と情報は入れて知らせます。")) + "</p>";
      var typeEl = body.querySelector("#dv-type"), opEl = body.querySelector("#dv-op");
      typeEl.value = r.type; opEl.value = r.op; body.querySelector("#dv-style").value = r.style;
      function show(id, on) {
        var n = body.querySelector("#" + id), lb = body.querySelector('label[for="' + id + '"]');
        if (n) n.style.display = on ? "" : "none";
        if (lb) lb.style.display = on ? "" : "none";
      }
      function sync() {
        var type = typeEl.value, op = opEl.value, cmp = type !== "list" && type !== "custom";
        show("dv-op", cmp);
        show("dv-f2", cmp && (op === "between" || op === "notBetween"));
        var l1 = body.querySelector('label[for="dv-f1"]'), l2 = body.querySelector('label[for="dv-f2"]');
        l1.textContent = type === "list" ? t("リスト（カンマ区切り、または A1:A9）") : type === "custom" ? t("数式（例：LEN(A1)>3）") : type === "textLength" ? t("文字数") : type === "date" ? t("日付") : t("最小値");
        l2.textContent = type === "date" ? t("日付") : t("最大値");
      }
      sync();
      typeEl.addEventListener("change", sync);
      opEl.addEventListener("change", sync);
      return M3.dialog({ title: cur ? t("入力規則を編集") : t("入力規則"), icon: "rule", body: body, actions: [
        { label: t("キャンセル"), value: null },
        { label: t("保存"), primary: true, value: function () {
          var g = parseRange(body.querySelector("#dv-ref").value);
          if (!g) return null;
          return { ref: rangeText(g), type: typeEl.value, op: opEl.value, f1: body.querySelector("#dv-f1").value.trim(), f2: body.querySelector("#dv-f2").value.trim(),
            allowBlank: body.querySelector("#dv-blank").checked, promptTitle: body.querySelector("#dv-prompt-title").value.trim(), prompt: body.querySelector("#dv-prompt").value.trim(),
            errTitle: body.querySelector("#dv-err-title").value.trim(), err: body.querySelector("#dv-err").value.trim(), style: body.querySelector("#dv-style").value };
        } }
      ] }).then(function (v) {
        if (!v) return;
        var idx = cur ? sh.dv.indexOf(cur) : -1;
        if (idx < 0) { sh.dv.forEach(function (x, i) { if (x.ref === v.ref) idx = i; }); }
        if (idx >= 0) sh.dv[idx] = v; else sh.dv.push(v);
        changed();
      });
    }
    function dvClear() {
      var g = range(), n = sh.dv.length;
      sh.dv = sh.dv.filter(function (rule) { var r = parseRange(rule.ref); return !r || r.c2 < g.c1 || r.c1 > g.c2 || r.r2 < g.r1 || r.r1 > g.r2; });
      if (sh.dv.length === n) O.notify(t("このセルには入力規則がありません"));
      else changed();
    }

    // ---------- セルのコメント ----------
    function noteEdit(c, r) {
      if (readOnly()) { O.notify(t("閲覧モードでは書きかえられません")); return; }
      var cur = noteAt(c, r), body = document.createElement("div"), me = O.userName() || t("自分");
      body.innerHTML = '<div class="dlg-grid"><label for="nt-author">' + esc(t("名前")) + '</label><input id="nt-author" type="text" maxlength="40" value="' + esc(cur ? cur.author || me : me) + '"></div>' +
        '<textarea id="nt-text" class="dlg-area dlg-area--note" placeholder="' + esc(t("コメントを入力")) + '">' + esc(cur ? cur.text : "") + "</textarea>" +
        '<p class="dlg-note">' + esc(t("セルの右上に赤い印が付きます。カーソルを合わせるとコメントが読めます。")) + "</p>";
      var actions = [{ label: t("キャンセル"), value: null }];
      if (cur) actions.push({ label: t("削除"), danger: true, value: { del: true } });
      actions.push({ label: t("保存"), primary: true, value: function () {
        var text = body.querySelector("#nt-text").value.trim();
        if (!text) return { del: true };
        return { text: text, author: body.querySelector("#nt-author").value.trim() || me, ts: Date.now() };
      } });
      return M3.dialog({ title: cur ? t("コメントを編集") : t("コメントを挿入"), icon: "sticky_note_2", body: body, actions: actions }).then(function (v) {
        if (v === null || v === undefined) return;
        if (v.del) delete sh.notes[ref(c, r)];
        else sh.notes[ref(c, r)] = { text: v.text, author: v.author, ts: v.ts };
        changed();
      });
    }
    function noteDelete() {
      if (readOnly()) { O.notify(t("閲覧モードでは書きかえられません")); return; }
      if (!noteAt(act.c, act.r)) { O.notify(t("このセルにはコメントがありません")); return; }
      delete sh.notes[ref(act.c, act.r)];
      changed();
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
      if (opts && opts.restore) { el("grid-scroll").scrollLeft = sx; el("grid-scroll").scrollTop = sy; renderWindow(); placeSel(); }
      applyZoom();
      focusGrid();
    }
    function applyZoom() {
      var w = el("grid-wrap");
      if (w) w.style.zoom = zoomZ;
      renderWindow();
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
      var winFrame = 0;
      sc.addEventListener("scroll", function () {
        if (!winFrame) winFrame = requestAnimationFrame(function () { winFrame = 0; renderWindow(); });
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
        sh.rh[r] = h; rowH[r] = h; posCache = null;
        var tr = el("grid").querySelector('tr[data-r="' + r + '"]');
        if (tr) tr.style.height = h + "px";
        measureRows(true);
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
      // Ctrl+Home は、いちばん上・左はしまで戻す（A1 が固定した行にあると、選ぶだけでは表が動かないため）
      if (k === "Home") { e.preventDefault(); if (mod) { var hs = el("grid-scroll"); hs.scrollTop = 0; hs.scrollLeft = 0; renderWindow(); select(0, 1, e.shiftKey); } else select(0, act.r, e.shiftKey); return; }
      if (k === "End" && mod) { e.preventDefault(); var l = book.last(si); select(Math.max(0, l.cols - 1), Math.max(1, l.rows), e.shiftKey); return; }
      if (k === "PageDown" || k === "PageUp") {
        e.preventDefault();
        if (mod) { switchSheet(si + (k === "PageDown" ? 1 : -1)); return; }
        moveBy(0, k === "PageDown" ? 20 : -20, e.shiftKey);
        return;
      }
      if (k === "Delete" || k === "Backspace") { e.preventDefault(); if (readOnly()) return; clearRange("content"); return; }
      if (k === "F2" && e.shiftKey) { e.preventDefault(); noteEdit(act.c, act.r); return; }
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
        { value: "note", label: t(noteAt(act.c, act.r) ? "コメントを編集" : "コメントを挿入"), icon: "sticky_note_2", sub: "Shift+F2" },
        { value: "noteDel", label: t("コメントを削除"), icon: "delete", disabled: !noteAt(act.c, act.r) },
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
          sortAsc: function () { sortRange(false); }, sortDesc: function () { sortRange(true); }, merge: toggleMerge,
          note: function () { noteEdit(act.c, act.r); }, noteDel: noteDelete };
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
      if (cmd === "tableStyleMenu") return styleMenu("table", at);
      if (cmd === "cellStyleMenu") return styleMenu("cell", at);
      if (cmd.indexOf("tstyle:") === 0) return applyTableStyle(cmd.slice(7));
      if (cmd.indexOf("cstyle:") === 0) return applyCellStyle(cmd.slice(7));
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
        case "dv": return dvDialog(at);
        case "dvClear": return dvClear();
        case "note": return noteEdit(act.c, act.r);
        case "noteDelete": return noteDelete();
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
        { label: t("コメント"), items: [{ big: 1, icon: "sticky_note_2", label: t("コメント"), sub: "Shift+F2", cmd: "note" },
          { stack: [{ icon: "delete", label: t("コメントを削除"), cmd: "noteDelete" }] }] },
        { label: t("セル"), items: [{ big: 1, icon: "add_box", label: t("挿入"), cmd: "insertMenu", drop: 1 }] }
      ] },
      { tab: t("デザイン"), groups: [
        { label: t("テーブルのスタイル"), items: [{ big: 1, icon: "table_chart", label: t("スタイル"), cmd: "tableStyleMenu", drop: 1 }]
          .concat(["plain", "matcha", "ocean", "mikan", "navy", "washi"].map(function (id) { var x = TSTYLES.filter(function (y) { return y.id === id; })[0]; return { big: 1, icon: "table_rows", label: x.label, cmd: "tstyle:" + id }; })) },
        { label: t("セルのスタイル"), items: [{ big: 1, icon: "style", label: t("セルのスタイル"), cmd: "cellStyleMenu", drop: 1 }] }
      ] },
      { tab: t("数式"), groups: [
        { label: t("関数ライブラリ"), items: [{ big: 1, icon: "functions", label: t("関数の挿入"), cmd: "fnMenu" }, { big: 1, icon: "add_circle", label: t("オート SUM"), cmd: "autoSumMenu", drop: 1 }]
          .concat(FN_CATS.map(function (c) { return { big: 1, icon: c.icon, label: c.label, cmd: "fncat:" + c.id, drop: 1 }; })) },
        { label: t("計算"), items: [{ big: 1, icon: "calculate", label: t("再計算"), cmd: "recalc" }] }
      ] },
      { tab: t("データ"), groups: [
        { label: t("並べ替えとフィルター"), items: [{ big: 1, icon: "arrow_upward", label: t("昇順"), cmd: "sort:asc" }, { big: 1, icon: "arrow_downward", label: t("降順"), cmd: "sort:desc" },
          { big: 1, icon: "filter_alt", label: t("フィルター"), cmd: "filter", toggle: "filter" }, { stack: [{ icon: "filter_alt_off", label: t("クリア"), cmd: "filterClear" }] }] },
        { label: t("データ ツール"), items: [{ big: 1, icon: "layers_clear", label: t("重複の削除"), cmd: "dedupe" }, { big: 1, icon: "format_color_fill", label: t("条件付き書式"), cmd: "cf", drop: 1 },
          { big: 1, icon: "rule", label: t("入力規則"), cmd: "dv" }, { stack: [{ icon: "rule_folder", label: t("入力規則のクリア"), cmd: "dvClear" }] }] }
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
    // Excel 2010 より後の関数は、ファイルの中では _xlfn. を付ける（付けないと Excel で #NAME? になる）。引数の区切りは , にそろえる
    var XLFN = /^(CONCAT|TEXTJOIN|IFS|MAXIFS|MINIFS|SWITCH|XLOOKUP|IFNA|XOR|DAYS|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|RANK\.EQ)$/;
    // セルのコメント（comments.xml）。書き手ごとに <author> を作る
    function commentsXml(s, noteKeys) {
      var authors = [], aid = {};
      noteKeys.forEach(function (k) { var a = s.notes[k].author || O.userName() || "SK Hub Office"; if (!(a in aid)) { aid[a] = authors.length; authors.push(a); } });
      var cms = noteKeys.map(function (k) {
        var nt = s.notes[k], a = nt.author || O.userName() || "SK Hub Office";
        return '<comment ref="' + k + '" authorId="' + aid[a] + '"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">' + X(nt.text) + "</t></r></text></comment>";
      });
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>' +
        authors.map(function (a) { return "<author>" + X(a) + "</author>"; }).join("") + "</authors><commentList>" + cms.join("") + "</commentList></comments>";
    }
    // ---------- Excel のグラフ（c:chartSpace）。グラフの範囲のセルを指すので、Excel でデータを直すとグラフも変わる ----------
    //   種類は 縦棒・横棒・折れ線・面・円・散布図（アプリのグラフと同じ）。系列の色は chart.js と同じ 8 色
    function absRef(c, r) { return ref(c, r).replace(/^([A-Z]+)(\d+)$/, function (m, a, b) { return "$" + a + "$" + b; }); }
    // グラフの元の値。範囲の中に空の行があるときは、セルを指さずに値をそのまま書く（exact でない）
    function chartModel(b, n, ch) {
      var g = parseRange(ch.ref);
      if (!g) return null;
      var r2 = Math.min(g.r2, g.r1 + 500), c2 = Math.min(g.c2, g.c1 + 30), rows = [];
      for (var r = g.r1; r <= r2; r++) {
        var row = [];
        for (var c = g.c1; c <= c2; c++) { var v = b.val(n, c, r); row.push(v === null || F.isErr(v) ? null : typeof v === "boolean" ? String(v) : v); }
        rows.push(row);
      }
      var spec = Chart.fromRows(rows);
      if (!spec.series.length) return null;
      spec.type = ch.type; spec.title = ch.title;
      return { g: g, spec: spec, exact: !spec.dropped };
    }
    function chartXml(m, ch, sheetName) {
      var g = m.g, spec = m.spec, type = ch.type, pal = (Chart.colors || []).map(function (c) { return c.replace("#", "").toUpperCase(); });
      var sn = "'" + String(sheetName).replace(/'/g, "''") + "'!", cats = spec.cats || [], head = !!spec.head;
      var rowA = g.r1 + (head ? 1 : 0), rowB = rowA + cats.length - 1;
      function rng(c) { return X(sn + absRef(c, rowA) + ":" + absRef(c, rowB)); }
      function strPts(list) { return '<c:ptCount val="' + list.length + '"/>' + list.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + X(v) + "</c:v></c:pt>"; }).join(""); }
      function numPts(list) { return '<c:ptCount val="' + list.length + '"/>' + list.map(function (v, i) { return typeof v === "number" && isFinite(v) ? '<c:pt idx="' + i + '"><c:v>' + v + "</c:v></c:pt>" : ""; }).join(""); }
      function fill(c) { return '<a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill>'; }
      var isPie = type === "pie", isXY = type === "scatter", series = (isPie ? spec.series.slice(0, 1) : spec.series).slice(0, 8);
      var numericCats = isXY && cats.length && cats.every(function (c) { return c !== "" && isFinite(+c); });
      var catXml = !cats.length ? "" : "<c:cat>" + (m.exact && spec.catCol ? "<c:strRef><c:f>" + rng(g.c1) + "</c:f><c:strCache>" + strPts(cats) + "</c:strCache></c:strRef>" : "<c:strLit>" + strPts(cats) + "</c:strLit>") + "</c:cat>";
      var xXml = !isXY ? "" : "<c:xVal>" + (numericCats ? (m.exact && spec.catCol ? "<c:numRef><c:f>" + rng(g.c1) + "</c:f><c:numCache><c:formatCode>General</c:formatCode>" + numPts(cats.map(Number)) + "</c:numCache></c:numRef>" : "<c:numLit><c:formatCode>General</c:formatCode>" + numPts(cats.map(Number)) + "</c:numLit>")
        : "<c:numLit><c:formatCode>General</c:formatCode>" + numPts(cats.map(function (c, i) { return i + 1; })) + "</c:numLit>") + "</c:xVal>";
      var sers = series.map(function (s, i) {
        var col = pal[i % pal.length], c = g.c1 + s.col;
        var tx = "<c:tx>" + (m.exact && head ? "<c:strRef><c:f>" + X(sn + absRef(c, g.r1)) + "</c:f><c:strCache>" + strPts([s.name]) + "</c:strCache></c:strRef>" : "<c:v>" + X(s.name) + "</c:v>") + "</c:tx>";
        var vals = m.exact ? "<c:numRef><c:f>" + rng(c) + "</c:f><c:numCache><c:formatCode>General</c:formatCode>" + numPts(s.values) + "</c:numCache></c:numRef>" : "<c:numLit><c:formatCode>General</c:formatCode>" + numPts(s.values) + "</c:numLit>";
        var sp, extra = "", tail = "";
        if (type === "line" || isXY) {
          sp = "<c:spPr><a:ln w=\"28575\" cap=\"rnd\">" + (isXY ? "<a:noFill/>" : fill(col)) + "<a:round/></a:ln></c:spPr>";
          extra = '<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>' + fill(col) + "<a:ln><a:noFill/></a:ln></c:spPr></c:marker>";
          tail = '<c:smooth val="0"/>';
        } else sp = "<c:spPr>" + fill(col) + "</c:spPr>";
        if (type === "col" || type === "bar") extra = '<c:invertIfNegative val="0"/>';
        if (isPie) {
          extra = cats.map(function (x, k) { return '<c:dPt><c:idx val="' + k + '"/><c:bubble3D val="0"/><c:spPr>' + fill(pal[k % pal.length]) + "</c:spPr></c:dPt>"; }).join("") +
            '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>';
        }
        return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/>' + tx + sp + extra + (isXY ? xXml + "<c:yVal>" + vals + "</c:yVal>" : catXml + "<c:val>" + vals + "</c:val>") + tail + "</c:ser>";
      }).join("");
      var ax = '<c:axId val="111111"/><c:axId val="222222"/>', plot;
      if (type === "col" || type === "bar") plot = '<c:barChart><c:barDir val="' + (type === "bar" ? "bar" : "col") + '"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + sers + '<c:gapWidth val="80"/>' + ax + "</c:barChart>";
      else if (type === "line") plot = '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + sers + '<c:marker val="1"/>' + ax + "</c:lineChart>";
      else if (type === "area") plot = '<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>' + sers + ax + "</c:areaChart>";
      else if (isPie) plot = '<c:pieChart><c:varyColors val="1"/>' + sers + '<c:firstSliceAng val="0"/></c:pieChart>';
      else plot = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' + sers + ax + "</c:scatterChart>";
      var grid = '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E1E0D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
      var axes = "";
      if (!isPie) {
        var horiz = type === "bar", common = '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
        if (isXY) {
          axes = '<c:valAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' + common + '<c:crossAx val="222222"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>' +
            '<c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>' + grid + common + '<c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>';
        } else {
          axes = '<c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? "l" : "b") + '"/>' + common + '<c:crossAx val="222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
            '<c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? "b" : "l") + '"/>' + grid + common + '<c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="' + (type === "area" ? "midCat" : "between") + '"/></c:valAx>';
        }
      }
      var title = ch.title ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="ja-JP" sz="1400" b="1"/><a:t>' + X(ch.title) + "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" : "";
      var legend = spec.legend !== false && (isPie ? cats.length > 0 : series.length > 1) ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : "";
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<c:roundedCorners val="0"/><c:chart>' + title + '<c:autoTitleDeleted val="' + (title ? 0 : 1) + '"/><c:plotArea><c:layout/>' + plot + axes + "</c:plotArea>" + legend + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
        '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="E1E0D9"/></a:solidFill></a:ln></c:spPr>' +
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr></c:chartSpace>';
    }
    // グラフの場所（画面の px）を、セルの番号とずれ（EMU）にする。画面は、行の見出し（44px）・列の見出し（28px）の分だけ内側から始まる
    function chartAnchor(s, ch) {
      var x = Math.max(0, ch.x - HEADW), y = Math.max(0, ch.y - 28), c = 0, r = 1;
      while (x >= (s.cw[c] || COLW) && c < 16383) { x -= s.cw[c] || COLW; c++; }
      while (y >= (s.rh[r] || 28) && r < 1048575) { y -= s.rh[r] || 28; r++; }
      return { col: c, colOff: Math.round(x * 9525), row: r - 1, rowOff: Math.round(y * 9525) };
    }
    // シート（<worksheet>）の直下に、<name r:id="rid"/> を入れる。before に並べた要素のどれかの前（なければいちばん後ろ）。同じ名前があれば先に消す。
    //   文字列の検索ではなく XML として扱う（dataBar の cfRule の中にも extLst があるため）
    function sheetChild(xml, name, rid, before) {
      var RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", doc = new DOMParser().parseFromString(xml, "application/xml"), root = doc.documentElement, at = null;
      Array.prototype.slice.call(root.children).forEach(function (c) { if (c.localName === name) root.removeChild(c); });
      Array.prototype.some.call(root.children, function (c) { if (before.indexOf(c.localName) >= 0) { at = c; return true; } });
      var node = doc.createElementNS(root.namespaceURI, name);
      node.setAttributeNS(RNS, "r:id", rid);
      root.insertBefore(node, at);
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
    }
    // どのシートも使える形で、グラフ・ドローイング・つながりを足す。get(道) → 文字・set(道, 文字)。Content_Types に足す Override を返す
    function nativeCharts(b, names, get, set) {
      var overrides = [], k = 0, d = 0;
      var R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      data.sheets.forEach(function (s, n) {
        var items = [];
        (s.charts || []).forEach(function (ch) {
          var m = chartModel(b, n, ch);
          if (!m) return;
          k++;
          set("xl/charts/chart" + k + ".xml", chartXml(m, ch, names[n]));
          overrides.push('<Override PartName="/xl/charts/chart' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
          items.push({ k: k, ch: ch });
        });
        if (!items.length) return;
        d++;
        set("xl/drawings/drawing" + d + ".xml", head + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + R + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
          items.map(function (it, i) {
            var a = chartAnchor(s, it.ch);
            return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + a.col + "</xdr:col><xdr:colOff>" + a.colOff + "</xdr:colOff><xdr:row>" + a.row + "</xdr:row><xdr:rowOff>" + a.rowOff + '</xdr:rowOff></xdr:from><xdr:ext cx="' + Math.round(it.ch.w * 9525) + '" cy="' + Math.round(it.ch.h * 9525) + '"/>' +
              '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Chart ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
              '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId' + (i + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>';
          }).join("") + "</xdr:wsDr>");
        set("xl/drawings/_rels/drawing" + d + ".xml.rels", head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          items.map(function (it, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="' + R + '/chart" Target="../charts/chart' + it.k + '.xml"/>'; }).join("") + "</Relationships>");
        overrides.push('<Override PartName="/xl/drawings/drawing' + d + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
        // シートからドローイングへのつながり。ほかのつながり（コメントなど）の番号とぶつからない rId にする
        var rp = "xl/worksheets/_rels/sheet" + (n + 1) + ".xml.rels", rels = get(rp) || head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        var next = 1; (rels.match(/Id="rId(\d+)"/g) || []).forEach(function (x) { next = Math.max(next, +x.match(/\d+/)[0] + 1); });
        set(rp, rels.replace("</Relationships>", '<Relationship Id="rId' + next + '" Type="' + R + '/drawing" Target="../drawings/drawing' + d + '.xml"/></Relationships>'));
        var sp = "xl/worksheets/sheet" + (n + 1) + ".xml";
        set(sp, sheetChild(get(sp), "drawing", "rId" + next, ["legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]));
      });
      return overrides;
    }
    function fileFormula(src) { return F.tokenize(src).map(function (x) { return x.t === "fn" && XLFN.test(x.v) ? "_xlfn." + x.s : x.t === "," ? "," : x.s; }).join(""); }
    // Excel にあるエラーだけを書く（#CIRC! などは #VALUE! にする。開いたときに Excel が計算しなおす）
    function fileErr(e) { return /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/.test(e) ? e : "#VALUE!"; }
    // 条件付き書式（Sheet の { type, v, color }）を Excel の cfRule にする。dxf(色) は塗りつぶしの書式の番号を返す
    function cfRuleXml(rule, g, prio, dxf) {
      var tl = ref(g.c1, g.r1), target = F.parseNum(rule.v), fill = hex6(rule.color || (rule.type === "bar" ? "#93c5fd" : "#fde68a")) || "FFFDE68A";
      var head = '<cfRule priority="' + prio + '" ';
      if (rule.type === "gt" || rule.type === "lt" || rule.type === "eq") {
        if (target === null && rule.type !== "eq") return "";
        var val = target !== null ? String(target) : '"' + String(rule.v || "").replace(/"/g, '""') + '"';
        return head + 'type="cellIs" dxfId="' + dxf(fill) + '" operator="' + { gt: "greaterThan", lt: "lessThan", eq: "equal" }[rule.type] + '"><formula>' + X(val) + "</formula></cfRule>";
      }
      if (rule.type === "text") {
        var q = String(rule.v || "").replace(/"/g, '""');
        return head + 'type="containsText" dxfId="' + dxf(fill) + '" operator="containsText" text="' + X(rule.v || "") + '"><formula>' + X('NOT(ISERROR(SEARCH("' + q + '",' + tl + ")))") + "</formula></cfRule>";
      }
      if (rule.type === "dup") return head + 'type="duplicateValues" dxfId="' + dxf(fill) + '"/>';
      if (rule.type === "bar") return head + 'type="dataBar"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="' + fill + '"/></dataBar></cfRule>';
      if (rule.type === "scale") return head + 'type="colorScale"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>';
      return "";
    }
    // 入力規則（Sheet の dv）を Excel の dataValidation にする
    function dvRuleXml(rule) {
      var g = parseRange(rule.ref), type = rule.type || "list";
      if (!g || ["list", "whole", "decimal", "date", "textLength", "custom"].indexOf(type) < 0) return "";
      function form(v, isList) {
        var s = String(v === undefined || v === null ? "" : v).trim().replace(/^=/, "");
        if (!s) return "";
        if (isList && !parseRange(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
      }
      var f1 = form(rule.f1, type === "list"), f2 = form(rule.f2, false);
      if (!f1) return "";
      var at = ' type="' + type + '"';
      if (type !== "list" && type !== "custom") at += ' operator="' + (rule.op || "between") + '"';
      at += ' allowBlank="' + (rule.allowBlank !== false ? 1 : 0) + '"';
      if (rule.prompt || rule.promptTitle) {
        at += ' showInputMessage="1"';
        if (rule.promptTitle) at += ' promptTitle="' + X(rule.promptTitle) + '"';
        if (rule.prompt) at += ' prompt="' + X(rule.prompt) + '"';
      }
      at += ' showErrorMessage="1" errorStyle="' + (["stop", "warning", "information"].indexOf(rule.style) >= 0 ? rule.style : "stop") + '"';
      if (rule.errTitle) at += ' errorTitle="' + X(rule.errTitle) + '"';
      if (rule.err) at += ' error="' + X(rule.err) + '"';
      return "<dataValidation" + at + ' sqref="' + rangeText(g) + '"><formula1>' + X(f1) + "</formula1>" + (f2 ? "<formula2>" + X(f2) + "</formula2>" : "") + "</dataValidation>";
    }
    function xlsxFiles() {
      var fonts = ['<font><sz val="11"/><name val="游ゴシック"/><family val="3"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
      var bordersX = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], numFmts = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
      var fontIdx = {}, fillIdx = {}, bdIdx = {}, nfIdx = {}, xfIdx = {}, dxfs = [], dxfIdx = {};
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
        // セルのない行も、高さを変えていれば書く
        Object.keys(s.rh).forEach(function (r) { if (+r >= 1) byRow[r] = byRow[r] || {}; });
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
              var f = "<f>" + X(fileFormula(raw.slice(1))) + "</f>";
              if (typeof v === "number" && isFinite(v)) cells.push('<c r="' + k + '"' + sAttr + ">" + f + "<v>" + v + "</v></c>");
              else if (typeof v === "boolean") cells.push('<c r="' + k + '"' + sAttr + ' t="b">' + f + "<v>" + (v ? 1 : 0) + "</v></c>");
              else if (F.isErr(v) || typeof v === "number") cells.push('<c r="' + k + '"' + sAttr + ' t="e">' + f + "<v>" + X(fileErr(typeof v === "number" ? "#NUM!" : v.err)) + "</v></c>");
              else cells.push('<c r="' + k + '"' + sAttr + ' t="str">' + f + "<v>" + X(v === null ? "" : String(v)) + "</v></c>");
            } else if (typeof v === "number" && isFinite(v)) cells.push('<c r="' + k + '"' + sAttr + "><v>" + v + "</v></c>");
            else if (typeof v === "boolean") cells.push('<c r="' + k + '"' + sAttr + ' t="b"><v>' + (v ? 1 : 0) + "</v></c>");
            else cells.push('<c r="' + k + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + X(String(v)) + "</t></is></c>");
          });
          rows.push('<row r="' + r + '"' + (s.rh[r] ? ' ht="' + Math.round(s.rh[r] * 0.75 * 10) / 10 + '" customHeight="1"' : "") + ">" + cells.join("") + "</row>");
        });
        var cols = Object.keys(s.cw).map(Number).sort(function (a, c) { return a - c; }).map(function (c) { return '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + Math.max(1, Math.round((s.cw[c] - 5) / 7 * 100) / 100) + '" customWidth="1"/>'; });
        var cfs = [];
        (s.cf || []).forEach(function (rule) {
          var g = parseRange(rule.ref), x = g && cfRuleXml(rule, g, cfs.length + 1, function (fill) {
            return idx(dxfIdx, dxfs, fill, '<dxf><font><color rgb="FF1F2328"/></font><fill><patternFill patternType="solid"><bgColor rgb="' + fill + '"/></patternFill></fill></dxf>');
          });
          if (x) cfs.push('<conditionalFormatting sqref="' + rangeText(g) + '">' + x + "</conditionalFormatting>");
        });
        var dvs = (s.dv || []).map(dvRuleXml).filter(Boolean);
        var dvStr = dvs.length ? '<dataValidations count="' + dvs.length + '">' + dvs.join("") + "</dataValidations>" : "";
        // セルのコメント（旧形式の comments + VML。Excel・LibreOffice・openpyxl が読める形）
        var noteKeys = Object.keys(s.notes || {}).filter(function (k) { return s.notes[k] && s.notes[k].text && parseRef(k); })
          .sort(function (a, b) { var p = parseRef(a), q = parseRef(b); return p.r - q.r || p.c - q.c; });
        var legacy = "";
        if (noteKeys.length) {
          files["xl/comments" + (n + 1) + ".xml"] = commentsXml(s, noteKeys);
          var shapes = noteKeys.map(function (k, i) {
            var p = parseRef(k), row = p.r - 1, col = p.c;
            var anchor = col + ", 15, " + row + ", 2, " + Math.min(col + 2, 16383) + ", 15, " + Math.min(row + 4, 1048575) + ", 4";
            return '<v:shape id="_x0000_s' + (1025 + i) + '" type="#_x0000_t202" style="position:absolute;margin-left:60pt;margin-top:2pt;width:108pt;height:59pt;z-index:' + (i + 1) + ';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">' +
              '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
              '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
              '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>' + anchor + "</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>" + row + "</x:Row><x:Column>" + col + "</x:Column></x:ClientData></v:shape>";
          });
          files["xl/drawings/vmlDrawing" + (n + 1) + ".vml"] = '<?xml version="1.0" encoding="UTF-8"?>\n<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + (n + 1) + '"/></o:shapelayout>' +
            '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
            shapes.join("") + "</xml>";
          files["xl/worksheets/_rels/sheet" + (n + 1) + ".xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' + (n + 1) + '.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' + (n + 1) + '.vml"/></Relationships>';
          ctypes.push('<Override PartName="/xl/comments' + (n + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
          legacy = '<legacyDrawing r:id="rId2"/>';
        }
        var fz = s.freeze, pane = fz && (fz.r || fz.c) ? '<pane' + (fz.c ? ' xSplit="' + fz.c + '"' : "") + (fz.r ? ' ySplit="' + fz.r + '"' : "") + ' topLeftCell="' + ref(fz.c || 0, (fz.r || 0) + 1) + '" activePane="' + (fz.r && fz.c ? "bottomRight" : fz.r ? "bottomLeft" : "topRight") + '" state="frozen"/>' : "";
        var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<dimension ref="A1:' + ref(Math.max(0, u.c2), Math.max(1, u.r2)) + '"/>' +
          '<sheetViews><sheetView' + (s.grid === false ? ' showGridLines="0"' : "") + (n === data.cur ? ' tabSelected="1"' : "") + ' workbookViewId="0">' + pane + "</sheetView></sheetViews>" +
          '<sheetFormatPr defaultRowHeight="18"/>' + (cols.length ? "<cols>" + cols.join("") + "</cols>" : "") +
          "<sheetData>" + rows.join("") + "</sheetData>" +
          (s.filter ? '<autoFilter ref="' + ref(s.filter.c1, s.filter.r) + ":" + ref(s.filter.c2, Math.max(s.filter.r2 || 0, u.r2)) + '"/>' : "") +
          (s.merges.length ? '<mergeCells count="' + s.merges.length + '">' + s.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join("") + "</mergeCells>" : "") + cfs.join("") + dvStr +
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' + legacy + "</worksheet>";
        files["xl/worksheets/sheet" + (n + 1) + ".xml"] = xml;
        sheetsXml.push('<sheet name="' + X(s.name) + '" sheetId="' + (n + 1) + '" r:id="rId' + (n + 1) + '"/>');
        rels.push('<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (n + 1) + '.xml"/>');
        ctypes.push('<Override PartName="/xl/worksheets/sheet' + (n + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      });
      // グラフは、Excel のグラフ（c:chart）として書く
      ctypes = ctypes.concat(nativeCharts(b, data.sheets.map(function (s) { return s.name; }), function (p) { return files[p] || ""; }, function (p, v) { files[p] = v; }));
      var N = data.sheets.length;
      files["xl/styles.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join("") + "</numFmts>" : "") +
        '<fonts count="' + fonts.length + '">' + fonts.join("") + '</fonts><fills count="' + fills.length + '">' + fills.join("") + '</fills><borders count="' + bordersX.length + '">' + bordersX.join("") + "</borders>" +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>" +
        '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>' + (dxfs.length ? '<dxfs count="' + dxfs.length + '">' + dxfs.join("") + "</dxfs>" : "") + "</styleSheet>";
      files["xl/workbook.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView activeTab="' + data.cur + '"/></bookViews><sheets>' + sheetsXml.join("") + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
      files["xl/_rels/workbook.xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join("") +
        '<Relationship Id="rId' + (N + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
      files["_rels/.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
      files["docProps/core.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        "<dc:title>" + X(O.file() ? O.file().name : "") + "</dc:title><dc:creator>" + X(O.userName() || "SK Hub Office") + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString().replace(/\.\d+Z$/, "Z") + "</dcterms:created></cp:coreProperties>";
      files["[Content_Types].xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ctypes.join("") + "</Types>";
      return files;
    }
    // ---------- 大きな部品（使うときだけ読み込む） ----------
    function loadVendor(globalName, file) {
      if (window[globalName]) return Promise.resolve(window[globalName]);
      return O.loadVendor(file).then(function () {
        if (!window[globalName]) throw new Error("vendor-missing: " + file);
        return window[globalName];
      });
    }
    function loadExcelJS() { return loadVendor("ExcelJS", "exceljs.min.js"); }
    function loadSheetJS() { return loadVendor("XLSX", "xlsx.full.min.js"); }

    // ---------- ExcelJS で .xlsx を書き出す（Excel のスタイルの書き方は ExcelJS まかせ） ----------
    //   いまの軽い書き出し（xlsxFiles）と同じ中身。グラフは、どちらも nativeCharts で Excel のグラフとして足す。
    function excelSheetName(name, i, used) {
      var n = String(name || "Sheet" + (i + 1)).replace(/[\[\]:*?\/\\]/g, "_").replace(/^'+|'+$/g, "").slice(0, 31) || "Sheet" + (i + 1);
      var base = n, k = 2;
      while (used[n.toLowerCase()]) { n = (base.slice(0, 28) + " " + k++).slice(0, 31); }
      used[n.toLowerCase()] = 1;
      return n;
    }
    function excelFont(st) {
      if (!st) return null;
      var f = {}, any = false;
      if (st.b) { f.bold = true; any = true; }
      if (st.i) { f.italic = true; any = true; }
      if (st.u) { f.underline = true; any = true; }
      if (st.s) { f.strike = true; any = true; }
      if (st.fs) { f.size = +st.fs; any = true; }
      if (st.ff) { f.name = st.ff; any = true; }
      var c = hex6(st.fg); if (c) { f.color = { argb: c }; any = true; }
      return any ? f : null;
    }
    function excelFill(st) {
      var c = st && hex6(st.bg);
      return c ? { type: "pattern", pattern: "solid", fgColor: { argb: c } } : null;
    }
    function excelBorder(st) {
      if (!st || !st.bd) return null;
      var out = {}, any = false;
      [["t", "top"], ["r", "right"], ["b", "bottom"], ["l", "left"]].forEach(function (x) {
        var w = st.bd[x[0]];
        if (!w) return;
        out[x[1]] = { style: w > 1 ? "medium" : "thin", color: { argb: "FF1F2328" } }; any = true;
      });
      return any ? out : null;
    }
    function excelAlign(st) {
      if (!st) return null;
      var a = {}, any = false;
      if (st.al) { a.horizontal = { l: "left", c: "center", r: "right" }[st.al]; any = true; }
      if (st.va) { a.vertical = { t: "top", m: "middle", b: "bottom" }[st.va]; any = true; }
      if (st.wrap) { a.wrapText = true; any = true; }
      return any ? a : null;
    }
    function excelApply(cell, st) {
      if (!st) return;
      var f = excelFont(st); if (f) cell.font = f;
      var fl = excelFill(st); if (fl) cell.fill = fl;
      var bd = excelBorder(st); if (bd) cell.border = bd;
      var al = excelAlign(st); if (al) cell.alignment = al;
      if (st.nf) cell.numFmt = st.nf;
    }
    function excelCfStyle(color) {
      var c = hex6(color || "#fde68a");
      return c ? { fill: { type: "pattern", pattern: "solid", bgColor: { argb: c } } } : undefined;
    }
    function excelFormula(src) {
      var f = fileFormula(String(src).replace(/^=/, ""));
      return f;
    }
    function excelJsFiles() {
      return loadExcelJS().then(function (ExcelJS) {
        var wb = new ExcelJS.Workbook(), used = {}, names = [];
        wb.creator = O.userName() || "SK Hub Office";
        wb.created = new Date();
        // グラフは、書き出したあとに Excel のグラフ（nativeCharts）として足す
        return Promise.resolve().then(function () {
          data.sheets.forEach(function (s, n) {
            names.push(excelSheetName(s.name, n, used));
            var ws = wb.addWorksheet(names[n], { properties: { defaultRowHeight: 18 } });
            var b = Book(data), maxC = 0, maxR = 0;
            Object.keys(s.cells).forEach(function (k) {
              var p = parseRef(k); if (!p) return;
              maxC = Math.max(maxC, p.c + 1); maxR = Math.max(maxR, p.r);
              var raw = s.cells[k], cell = ws.getCell(k.replace(/\$/g, "")), st = s.style[k] || EMPTY, v;
              if (typeof raw === "string" && raw.charAt(0) === "=" && raw.length > 1) {
                v = b.val(n, p.c, p.r);
                var result;
                if (F.isErr(v)) result = { error: fileErr(v.err) };
                else if (v === null || v === undefined) result = "";
                else if (typeof v === "number" || typeof v === "boolean" || typeof v === "string") result = v;
                else result = String(v);
                cell.value = { formula: excelFormula(raw), result: result };
              } else {
                v = b.val(n, p.c, p.r);
                if (v === null || v === undefined) v = "";
                else if (F.isErr(v)) v = fileErr(v.err);
                else if (typeof v === "string" && raw.charAt(0) === "'") v = raw.slice(1);
                else if (typeof v === "number" || typeof v === "boolean" || typeof v === "string") v = v;
                else v = String(v);
                cell.value = v;
              }
              excelApply(cell, st);
            });
            // 行の高さ・列の幅（アプリの px を、Excel の pt・文字数に直す）
            Object.keys(s.cw || {}).forEach(function (c) { ws.getColumn(+c + 1).width = Math.max(1, Math.round((s.cw[c] - 5) / 7 * 100) / 100); });
            Object.keys(s.rh || {}).forEach(function (r) { ws.getRow(+r).height = Math.max(1, Math.round(s.rh[r] * 0.75 * 100) / 100); });
            // コメント
            Object.keys(s.notes || {}).forEach(function (k) {
              var nt = s.notes[k], p = parseRef(k);
              if (!p || !nt || !nt.text) return;
              ws.getCell(k).note = { texts: [{ text: String(nt.text) }], author: nt.author || O.userName() || "SK Hub Office" };
            });
            // 入力規則
            (s.dv || []).forEach(function (rule) {
              var g = parseRange(rule.ref); if (!g) return;
              var f1 = String(rule.f1 === undefined || rule.f1 === null ? "" : rule.f1).trim().replace(/^=/, ""), f2 = String(rule.f2 === undefined || rule.f2 === null ? "" : rule.f2).trim().replace(/^=/, "");
              if (!f1) return;
              if (rule.type === "list" && !parseRange(f1)) f1 = '"' + f1.replace(/"/g, '""') + '"';
              var dv = {
                type: rule.type || "list", allowBlank: rule.allowBlank !== false,
                formulae: f2 ? [f1, f2] : [f1], showErrorMessage: true,
                errorStyle: /^(stop|warning|information)$/.test(rule.style) ? rule.style : "stop"
              };
              if (rule.type !== "list" && rule.type !== "custom") dv.operator = rule.op || "between";
              if (rule.prompt || rule.promptTitle) { dv.showInputMessage = true; if (rule.promptTitle) dv.promptTitle = rule.promptTitle; if (rule.prompt) dv.prompt = rule.prompt; }
              if (rule.errTitle) dv.errorTitle = rule.errTitle;
              if (rule.err) dv.error = rule.err;
              ws.dataValidations.add(rangeText(g), dv);
            });
            // 条件付き書式
            (s.cf || []).forEach(function (rule) {
              var g = parseRange(rule.ref); if (!g) return;
              var style = excelCfStyle(rule.color), target = F.parseNum(rule.v), first = ref(g.c1, g.r1), rng = rangeText(g), model;
              if (rule.type === "gt" || rule.type === "lt" || rule.type === "eq") {
                if (target === null && rule.type !== "eq") return;
                var f = target !== null ? String(target) : '"' + String(rule.v || "").replace(/"/g, '""') + '"';
                model = { type: "cellIs", operator: { gt: "greaterThan", lt: "lessThan", eq: "equal" }[rule.type], formulae: [f], style: style };
              } else if (rule.type === "text") {
                model = { type: "containsText", operator: "containsText", text: String(rule.v || ""), formulae: ['NOT(ISERROR(SEARCH("' + String(rule.v || "").replace(/"/g, '""') + '",' + first + ")))"], style: style };
              } else if (rule.type === "dup") {
                // ExcelJS は duplicateValues をまだ書けないので、同じ意味の数式にする（読みこむときに dup に戻す）
                model = { type: "expression", formulae: ["COUNTIF(" + rng + "," + first + ")>1"], style: style };
              } else if (rule.type === "bar") {
                model = { type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: hex6(rule.color || "#93c5fd") || "FF93C5FD" } };
              } else if (rule.type === "scale") {
                model = { type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }] };
              }
              if (model) ws.addConditionalFormatting({ ref: rangeText(g), rules: [model] });
            });
            // 結合・固定・フィルター・目盛線
            (s.merges || []).forEach(function (m) { var g = parseRange(m); if (g) ws.mergeCells(g.r1, g.c1 + 1, g.r2, g.c2 + 1); });
            if (s.filter) {
              var fg = parseRange(ref(s.filter.c1, s.filter.r) + ":" + ref(s.filter.c2, Math.max(s.filter.r2 || 0, maxR)));
              if (fg) ws.autoFilter = rangeText(fg);
            }
            var top = s.freeze ? s.freeze.r : 0, left = s.freeze ? s.freeze.c : 0;
            var view = { showGridLines: s.grid !== false, tabSelected: n === data.cur };
            if (top || left) {
              view.state = "frozen"; view.xSplit = left; view.ySplit = top; view.topLeftCell = ref(left, top + 1);
              view.activePane = top && left ? "bottomRight" : top ? "bottomLeft" : "topRight";
            }
            ws.views = [view];
            ws.pageSetup.margins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
          });
          return wb.xlsx.writeBuffer().then(function (buf) { return excelJsFix(buf, names); });
        });
      });
    }
    // ExcelJS の出力の直し。書き手が "Author" 固定になる・<legacyDrawing> が <drawing> より前に来る（Excel は順番がちがうと開けない）
    function excelJsFix(buf, names) {
      if (!Zip.canRead) return buf;
      return Zip.read(new Blob([buf])).then(function (z) {
        var out = { "[Content_Types].xml": z["[Content_Types].xml"] };
        Object.keys(z).forEach(function (nm) { out[nm] = z[nm]; });
        data.sheets.forEach(function (s, n) {
          var sp = "xl/worksheets/sheet" + (n + 1) + ".xml", rp = "xl/worksheets/_rels/sheet" + (n + 1) + ".xml.rels";
          if (!out[sp]) return;
          // 順番は drawing → legacyDrawing → tableParts → extLst。グラフ（drawing）は、あとで legacyDrawing の前に足す
          var xml = Zip.text(out[sp]), m = /<legacyDrawing\b[^>]*r:id="([^"]+)"/.exec(xml);
          if (m) out[sp] = new TextEncoder().encode(sheetChild(xml, "legacyDrawing", m[1], ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]));
          var noteKeys = Object.keys(s.notes || {}).filter(function (k) { return s.notes[k] && s.notes[k].text && parseRef(k); });
          var rel = /<Relationship\b[^>]*relationships\/comments"[^>]*>/.exec(Zip.text(out[rp])), tg = rel && /Target="([^"]+)"/.exec(rel[0]);
          if (noteKeys.length && tg) {
            var cp = resolvePart("xl/worksheets/", tg[1]);
            if (out[cp]) out[cp] = new TextEncoder().encode(commentsXml(s, noteKeys));
          }
        });
        // グラフ（Excel のグラフ）を足す。[Content_Types].xml にも足す
        var enc = new TextEncoder(), ov = nativeCharts(Book(data), names, function (p) { return out[p] ? Zip.text(out[p]) : ""; }, function (p, v) { out[p] = enc.encode(v); });
        if (ov.length) out["[Content_Types].xml"] = enc.encode(Zip.text(out["[Content_Types].xml"]).replace("</Types>", ov.join("") + "</Types>"));
        return Zip.write(out);
      });
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
      if (id === "xlsx-charts") {
        return excelJsFiles().then(function (buf) {
          return O.saveBlob(O.fileName(".xlsx"), new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), t("Excel ブック（グラフ付き）"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
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
    // 相対のつながり（../comments1.xml など）を、ZIP の中の道に直す
    function resolvePart(base, target) {
      if (/^\//.test(target)) return target.replace(/^\/+/, "");
      var parts = (base + target).split("/"), out = [];
      parts.forEach(function (x) { if (x === "..") out.pop(); else if (x !== "." && x !== "") out.push(x); });
      return out.join("/");
    }
    // 日付の入力規則は、ファイルの中では 1899/12/30 からの数。読むときに yyyy/m/d に戻す
    function dvNumText(type, v) {
      if (type !== "date" || v === "" || !/^-?\d+(\.\d+)?$/.test(String(v).trim())) return v;
      var n = +v;
      return n >= 1 && n < 2958466 ? F.format(n, "yyyy/m/d").text : v;
    }
    // ---------- Excel のグラフ（c:chart）を読む ----------
    //   このアプリのグラフは「データの範囲（ref）1 つ」で決まる（1 列目が項目・1 行目が系列の名前）。
    //   Excel のグラフは系列ごとにセルを指すので、系列が同じ行の並んだ列で、項目の列がすぐ左にあるものだけ、範囲にして読む。
    //   ほかのシートを指すもの・飛び飛びの列・値を直接書いたものは、読まずに飛ばす。種類は 縦棒・横棒・折れ線・面・円・散布図
    var CHART_KIND = { barChart: "col", bar3DChart: "col", lineChart: "line", line3DChart: "line", areaChart: "area", area3DChart: "area", pieChart: "pie", pie3DChart: "pie", doughnutChart: "pie", scatterChart: "scatter" };
    function dkid(node, name) { var k = node && node.firstElementChild; while (k) { if (k.localName === name) return k; k = k.nextElementSibling; } return null; }
    function dkids(node, name) { var out = [], k = node && node.firstElementChild; while (k) { if (k.localName === name) out.push(k); k = k.nextElementSibling; } return out; }
    // 'シート名'!$B$2:$B$3 → { sheet, g }（シート名がなければ sheet は ""）
    function chartFormula(f) {
      var m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(String(f || "").trim()), sheet = "", range = String(f || "").trim();
      if (m) { sheet = (m[1] || m[2]).replace(/''/g, "'"); range = m[3]; }
      var g = parseRange(range.replace(/\$/g, ""));
      return g ? { sheet: sheet, g: g } : null;
    }
    function chartFromXml(cdoc, sheetName) {
      var plot = kid(cdoc, "plotArea"), group = null, kind = null;
      Array.prototype.some.call(plot ? plot.children : [], function (c) { if (CHART_KIND[c.localName]) { group = c; kind = CHART_KIND[c.localName]; return true; } });
      if (!group) return null;
      if (kind === "col" && (dkid(group, "barDir") || { getAttribute: function () { return ""; } }).getAttribute("val") === "bar") kind = "bar";
      var sers = dkids(group, "ser"), cols = [], cat = null, first = null, headRow = null;
      if (!sers.length) return null;
      if (kind === "pie") sers = sers.slice(0, 1);
      function refOf(wrap) { var f = wrap && kid(wrap, "f"); return f ? chartFormula(f.textContent) : null; }
      for (var i = 0; i < sers.length; i++) {
        var v = refOf(dkid(sers[i], kind === "scatter" ? "yVal" : "val")), nm = refOf(dkid(sers[i], "tx")), ct = refOf(dkid(sers[i], kind === "scatter" ? "xVal" : "cat"));
        if (!v || v.g.c1 !== v.g.c2 || (v.sheet && v.sheet !== sheetName)) return null;
        if (first && (v.g.r1 !== first.g.r1 || v.g.r2 !== first.g.r2 || v.g.c1 !== cols[cols.length - 1] + 1)) return null;
        if (!first) first = v;
        if (nm) {
          if (nm.g.c1 !== v.g.c1 || nm.g.r1 !== nm.g.r2 || nm.g.r1 !== v.g.r1 - 1) return null;
          headRow = v.g.r1 - 1;
        } else if (headRow !== null) return null;
        if (ct && !cat) cat = ct;
        cols.push(v.g.c1);
      }
      var c1 = cols[0], r1 = headRow !== null ? headRow : first.g.r1;
      if (cat) {
        if (cat.g.c1 !== cat.g.c2 || cat.g.c1 !== c1 - 1 || cat.g.r1 !== first.g.r1 || cat.g.r2 !== first.g.r2 || (cat.sheet && cat.sheet !== sheetName)) return null;
        c1 = c1 - 1;
      }
      var title = "", tEl = dkid(kid(cdoc, "chart"), "title"), atd = dkid(kid(cdoc, "chart"), "autoTitleDeleted");
      if (tEl && !(atd && atd.getAttribute("val") === "1")) title = kids(tEl, "t").map(function (x) { return x.textContent; }).join("").trim();
      return { type: kind, ref: rangeText({ c1: c1, r1: r1, c2: cols[cols.length - 1], r2: first.g.r2 }), title: title };
    }
    // シートの drawing にあるグラフを、s.charts に足す。dir は、シートのファイルのあるフォルダ（xl/ から）
    function readXlsxCharts(z, wrel, dir, s) {
      var RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", drawPart = null;
      if (wrel) kids(wrel, "Relationship").forEach(function (r) { if (/\/drawing$/.test(r.getAttribute("Type") || "")) drawPart = resolvePart("xl/" + dir, r.getAttribute("Target") || ""); });
      var ddoc = drawPart && Zip.xml(z[drawPart]);
      if (!ddoc) return;
      var dd = drawPart.replace(/[^\/]+$/, ""), drel = Zip.xml(z[dd + "_rels/" + drawPart.slice(dd.length) + ".rels"]), parts = {};
      if (drel) kids(drel, "Relationship").forEach(function (r) { parts[r.getAttribute("Id")] = resolvePart(dd, r.getAttribute("Target") || ""); });
      function colX(c) { var x = HEADW; for (var i = 0; i < c; i++) x += s.cw[i] || COLW; return x; }
      function rowY(r) { var y = 28; for (var i = 1; i <= r; i++) y += s.rh[i] || 28; return y; }
      function pos(from) {
        var n = function (k) { var e = kid(from, k); return e ? +e.textContent || 0 : 0; };
        return { x: colX(n("col")) + n("colOff") / 9525, y: rowY(n("row")) + n("rowOff") / 9525 };
      }
      Array.prototype.forEach.call(ddoc.documentElement.children, function (a) {
        if (!/Anchor$/.test(a.localName)) return;
        var cEl = kids(a, "chart")[0], rid = cEl && (cEl.getAttributeNS(RNS, "id") || cEl.getAttribute("r:id")), part = rid && parts[rid], cdoc = part && Zip.xml(z[part]);
        if (!cdoc) return;
        var ch = chartFromXml(cdoc, s.name);
        if (!ch) return;
        var from = dkid(a, "from"), to = dkid(a, "to"), ext = dkid(a, "ext"), ap = dkid(a, "pos"), x, y, w, h;
        if (from) { var p = pos(from); x = p.x; y = p.y; } else if (ap) { x = HEADW + (+ap.getAttribute("x") || 0) / 9525; y = 28 + (+ap.getAttribute("y") || 0) / 9525; } else return;
        if (ext) { w = (+ext.getAttribute("cx") || 0) / 9525; h = (+ext.getAttribute("cy") || 0) / 9525; }
        else if (to) { var q = pos(to); w = q.x - x; h = q.y - y; }
        ch.id = "x" + Date.now().toString(36) + s.charts.length;
        ch.x = Math.round(x); ch.y = Math.round(y); ch.w = Math.max(160, Math.min(1600, Math.round(w || 480))); ch.h = Math.max(120, Math.min(1200, Math.round(h || 300)));
        s.charts.push(ch);
      });
    }
    function readXlsx(file) {
      if (!Zip.canRead) return Promise.reject(new Error("no-inflate"));
      return Zip.read(file).then(function (z) {
        var wb = Zip.xml(z["xl/workbook.xml"]);
        if (!wb) throw new Error("not-xlsx");
        var relsDoc = Zip.xml(z["xl/_rels/workbook.xml.rels"]), relMap = {};
        kids(relsDoc, "Relationship").forEach(function (r) { relMap[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var shared = kids(Zip.xml(z["xl/sharedStrings.xml"]), "si").map(textOf);
        var styles = Zip.xml(z["xl/styles.xml"]), nfMap = {}, fonts = [], fillsL = [], bds = [], xfs = [], dxfFill = [];
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
          // 条件付き書式の色（dxf の塗りつぶし）
          kids(kid(styles, "dxfs"), "dxf").forEach(function (d) { var p = kid(d, "patternFill"); dxfFill.push(p ? argb(kid(p, "bgColor")) || argb(kid(p, "fgColor")) : null); });
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
          // 入力規則
          kids(doc, "dataValidation").forEach(function (dv) {
            var sq = (dv.getAttribute("sqref") || "").split(/\s+/).filter(Boolean)[0], g = parseRange(sq);
            var type = dv.getAttribute("type") || "any";
            if (!g || ["list", "whole", "decimal", "date", "textLength", "custom"].indexOf(type) < 0) return;
            var f1El = kid(dv, "formula1"), f2El = kid(dv, "formula2"), f1 = f1El ? f1El.textContent : "", f2 = f2El ? f2El.textContent : "";
            if (type === "list" && /^".*"$/.test(f1)) f1 = f1.slice(1, -1).replace(/""/g, '"');
            s.dv.push({ ref: rangeText(g), type: type, op: dv.getAttribute("operator") || "between", f1: dvNumText(type, f1), f2: dvNumText(type, f2),
              allowBlank: dv.getAttribute("allowBlank") !== "0",
              promptTitle: dv.getAttribute("promptTitle") || "", prompt: dv.getAttribute("prompt") || "",
              errTitle: dv.getAttribute("errorTitle") || "", err: dv.getAttribute("error") || "",
              style: /^(stop|warning|information)$/.test(dv.getAttribute("errorStyle")) ? dv.getAttribute("errorStyle") : "stop" });
          });
          // セルのコメント（ワークシートの rels → comments）。VML は読まなくてよい
          var dir = target.replace(/[^\/]+$/, ""), baseName = target.slice(dir.length);
          var wrel = Zip.xml(z["xl/" + dir + "_rels/" + baseName + ".rels"]), commentsPart = null;
          if (wrel) kids(wrel, "Relationship").forEach(function (r) {
            if (/\/comments$/.test(r.getAttribute("Type") || "")) commentsPart = resolvePart("xl/" + dir, r.getAttribute("Target") || "");
          });
          var cdoc = commentsPart && Zip.xml(z[commentsPart]);
          if (cdoc) {
            var authors = kids(cdoc, "author").map(function (a) { return a.textContent; });
            kids(cdoc, "comment").forEach(function (cm) {
              var rf = cm.getAttribute("ref"), tx = textOf(kid(cm, "text")), aid = +cm.getAttribute("authorId");
              if (rf && parseRef(rf) && tx) s.notes[rf] = { text: tx, author: authors[aid] || "", ts: Date.now() };
            });
          }
          readXlsxCharts(z, wrel, dir, s);
          var pane = kid(doc, "pane");
          if (pane && pane.getAttribute("state") === "frozen") s.freeze = { r: +pane.getAttribute("ySplit") || 0, c: +pane.getAttribute("xSplit") || 0 };
          var sv = kid(doc, "sheetView");
          if (sv && sv.getAttribute("showGridLines") === "0") s.grid = false;
          // 条件付き書式（Sheet で描けるものだけ）
          kids(doc, "conditionalFormatting").forEach(function (cf) {
            (cf.getAttribute("sqref") || "").split(/\s+/).forEach(function (sq) {
              var g = parseRange(sq);
              if (!g) return;
              kids(cf, "cfRule").forEach(function (rule) {
                var tp = rule.getAttribute("type"), op = rule.getAttribute("operator"), fe = kid(rule, "formula"), fx = fe ? fe.textContent : "";
                var color = dxfFill[+rule.getAttribute("dxfId")] || undefined, o = null;
                if (tp === "cellIs" && /^(greaterThan|lessThan|equal)$/.test(op)) o = { type: { greaterThan: "gt", lessThan: "lt", equal: "eq" }[op], v: /^".*"$/.test(fx) ? fx.slice(1, -1).replace(/""/g, '"') : fx };
                else if (tp === "containsText") o = { type: "text", v: rule.getAttribute("text") || "" };
                else if (tp === "duplicateValues") o = { type: "dup" };
                else if (tp === "expression" && /COUNTIF\s*\(/i.test(fx) && /\)\s*>\s*1\s*$/.test(fx)) o = { type: "dup" };
                else if (tp === "dataBar") { o = { type: "bar" }; color = argb(kids(rule, "color")[0]) || undefined; }
                else if (tp === "colorScale") { o = { type: "scale" }; color = undefined; }
                if (!o) return;
                o.ref = rangeText(g);
                if (color) o.color = color;
                s.cf.push(o);
              });
            });
          });
          out.sheets.push(s);
        });
        if (!out.sheets.length) out.sheets.push(fixSheet({ name: "Sheet1" }, 0));
        var view = kid(wb, "workbookView");
        if (view && view.getAttribute("activeTab")) out.cur = Math.min(out.sheets.length - 1, +view.getAttribute("activeTab"));
        return out;
      });
    }
    // .xls・.xlsb・.ods・.fods（と、読めなかった .xlsx）を SheetJS で読む。
    //   値・数式・結合・列の幅・行の高さ・コメント・入力規則・フィルターを読む。
    //   SheetJS CE は細かい書式（フォント・塗り）をあまり返さないので、取れたものだけ使う。
    function jsColor(c) {
      if (!c) return null;
      var s = c.rgb ? String(c.rgb) : "";
      if (/^[0-9a-f]{8}$/i.test(s)) return "#" + s.slice(2).toLowerCase();
      if (/^[0-9a-f]{6}$/i.test(s)) return "#" + s.toLowerCase();
      return null;
    }
    function jsStyle(cell) {
      var cs = cell.s || {}, st = {}, font = cs.font || {}, fill = cs.fill || {}, al = cs.alignment || {}, bd = cs.border || {};
      if (font.bold) st.b = 1;
      if (font.italic) st.i = 1;
      if (font.underline) st.u = 1;
      if (font.strike) st.s = 1;
      if (font.sz) st.fs = +font.sz;
      if (font.name) st.ff = font.name;
      var fg = jsColor(font.color); if (fg) st.fg = fg;
      if (fill.patternType === "solid") { var bg = jsColor(fill.fgColor) || jsColor(fill.bgColor); if (bg) st.bg = bg; }
      [["top", "t"], ["right", "r"], ["bottom", "b"], ["left", "l"]].forEach(function (x) {
        var side = bd[x[0]];
        if (side && side.style) { st.bd = st.bd || {}; st.bd[x[1]] = /medium|thick|double/.test(side.style) ? 2 : 1; }
      });
      if (al.horizontal === "left" || al.horizontal === "center" || al.horizontal === "right") st.al = al.horizontal.charAt(0);
      if (al.vertical === "top" || al.vertical === "center") st.va = al.vertical === "top" ? "t" : "m";
      if (al.wrapText) st.wrap = 1;
      var nf = cell.z && cell.z !== "General" ? String(cell.z).replace(/\[\$-[0-9A-F]+\]/gi, "").replace(/\\/g, "") : "";
      if (nf) st.nf = nf;
      return Object.keys(st).length ? st : null;
    }
    function jsDateText(v, nf) {
      var d = v instanceof Date ? v : new Date(v);
      if (!isFinite(d.getTime())) return "";
      var text = d.getFullYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate();
      if (/[hs]/i.test(nf || "")) text += " " + ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2) + (/(?:ss|s\b)/i.test(nf || "") ? ":" + ("0" + d.getSeconds()).slice(-2) : "");
      return text;
    }
    function readSheetJs(file) {
      return loadSheetJS().then(function (XLSX) {
        return file.arrayBuffer().then(function (ab) {
          var wb = XLSX.read(new Uint8Array(ab), { type: "array", cellStyles: true, cellDates: true, cellNF: true, cellFormula: true, sheetStubs: true });
          var out = { v: 2, cur: 0, sheets: [] };
          (wb.SheetNames || []).forEach(function (name, n) {
            var ws = wb.Sheets[name] || {}, s = fixSheet({ name: name || "Sheet" + (n + 1) }, n);
            Object.keys(ws).forEach(function (k) {
              if (k.charAt(0) === "!") return;
              var p = parseRef(k), cell = ws[k];
              if (!p || !cell) return;
              var raw;
              if (cell.f) raw = "=" + String(cell.f).replace(/_xlfn\.(_xlws\.)?/g, "");
              else if (cell.t === "s" || cell.t === "str") raw = String(cell.v === undefined || cell.v === null ? "" : cell.v);
              else if (cell.t === "b") raw = cell.v ? "TRUE" : "FALSE";
              else if (cell.t === "e") raw = String(cell.w || cell.v || "");
              else if (cell.t === "d") raw = jsDateText(cell.v, cell.z);
              else if (cell.t === "n") {
                var num = +cell.v;
                raw = cell.z && F.isDateFormat(cell.z) && num >= 1 && num < 2958466 && Math.floor(num) === num ? F.format(num, "yyyy/m/d").text : F.str(num);
              } else raw = cell.v === undefined || cell.v === null ? "" : String(cell.v);
              if (raw === "") return;
              if (!cell.f && (cell.t === "s" || cell.t === "str") && (/^[=+\-']/.test(raw) || F.parseNum(raw) !== null || /^(TRUE|FALSE)$/i.test(raw))) raw = "'" + raw;
              s.cells[k] = raw;
              var st = jsStyle(cell); if (st) s.style[k] = st;
              if (cell.c) {
                var cs = Array.isArray(cell.c) ? cell.c : [cell.c];
                cs.forEach(function (cm) {
                  if (cm && (cm.t || cm.h)) s.notes[k] = { text: String(cm.t || cm.h || ""), author: String(cm.a || ""), ts: Date.now() };
                });
              }
            });
            if (ws["!merges"]) ws["!merges"].forEach(function (m) {
              if (m && m.s && m.e) s.merges.push(rangeText({ c1: m.s.c, r1: m.s.r + 1, c2: m.e.c, r2: m.e.r + 1 }));
            });
            (ws["!cols"] || []).forEach(function (c, i) {
              if (!c) return;
              var w = c.wpx ? c.wpx : c.wch ? Math.round(c.wch * 7 + 5) : c.width ? Math.round(c.width * 7 + 5) : 0;
              if (w) s.cw[i] = Math.max(16, w);
            });
            (ws["!rows"] || []).forEach(function (r, i) {
              if (!r) return;
              var h = r.hpx ? r.hpx : r.hpt ? Math.round(r.hpt * 4 / 3) : 0;
              if (h) s.rh[i] = h;
            });
            if (ws["!autofilter"] && ws["!autofilter"].ref) {
              var fg = parseRange(ws["!autofilter"].ref);
              if (fg) s.filter = { r: fg.r1, c1: fg.c1, c2: fg.c2, r2: fg.r2 };
            }
            var fz = ws["!freeze"];
            if (fz) {
              if (typeof fz === "string") { var fzr = parseRange(fz); if (fzr) s.freeze = { r: fzr.r1 - 1, c: fzr.c1 }; }
              else s.freeze = { r: +fz.ySplit || 0, c: +fz.xSplit || 0 };
            }
            out.sheets.push(s);
          });
          if (!out.sheets.length) out.sheets.push(fixSheet({ name: "Sheet1" }, 0));
          var view = wb.Workbook && wb.Workbook.WBView && wb.Workbook.WBView[0];
          if (view && view.activeTab !== undefined) out.cur = Math.min(out.sheets.length - 1, +view.activeTab || 0);
          return out;
        });
      });
    }
    function importFile(file) {
      if (/\.(xlsx|xlsm)$/i.test(file.name)) {
        if (!Zip.canRead) return readSheetJs(file);
        return readXlsx(file).catch(function () { return readSheetJs(file); });
      }
      if (/\.(xls|xlsb|ods|fods)$/i.test(file.name)) return readSheetJs(file);
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
        { id: "xlsx-charts", icon: "insert_chart", label: t("Excel ブック（グラフ付き .xlsx）"), desc: t("グラフを Excel のグラフとして入れます。少し時間がかかります") },
        { id: "csv", icon: "csv", label: "CSV（.csv）", desc: t("いまのシートの値を、カンマで区切った文字にします") }
      ],
      exportAs: exportAs,
      imports: { label: ".xlsx・.xlsm・.xls・.xlsb・.ods・.csv", accept: ".xlsx,.xlsm,.xls,.xlsb,.ods,.fods,.csv,.tsv,.txt" },
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
