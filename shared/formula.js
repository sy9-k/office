// SK Hub Office の数式（window.OfficeFormula）。Sheet の「=」で始まるセルを計算する（Excel と同じ書き方）
//   ・読む（parse）→ 計算する（evaluate）。参照は A1・$A$1・A1:B5・A:A・1:3・Sheet2!A1・'名前 の シート'!A1:B2
//   ・演算子: + - * / ^ & % = <> < > <= >=（範囲どうしの計算もできる。SUMPRODUCT(A1:A3*B1:B3) など）
//   ・関数は FN（SUM・IF・VLOOKUP・XLOOKUP・COUNTIF・TEXT・DATE など 100 ほど）
//   ・コピー・オートフィルで参照をずらす（shift）、行・列の挿入と削除で参照を直す（adjust）、シートの名前を変える（renameSheet）
//   ・数値の表示形式（format: "#,##0"・"¥#,##0"・"0.0%"・"yyyy/m/d" など）と、日付の読み取り（parseDate）
//   値: 数・文字列・true/false・null（空のセル）・{ err: "#DIV/0!" }・範囲 { rows: [[値…]…] }
//   evaluate の ctx: { get(シート番号, 列, 行) → 値, sheet(名前) → シート番号（ないときは -1）, here: { s, c, r }, last(シート番号) → { rows, cols } }
(function () {
  "use strict";
  var ERR = { DIV0: "#DIV/0!", VALUE: "#VALUE!", REF: "#REF!", NAME: "#NAME?", NA: "#N/A", NUM: "#NUM!", CIRC: "#CIRC!" };
  function err(code) { return { err: code }; }
  function isErr(v) { return v !== null && typeof v === "object" && typeof v.err === "string"; }
  function isRange(v) { return v !== null && typeof v === "object" && Array.isArray(v.rows); }

  function colIndex(s) { var n = 0; s = s.toUpperCase(); for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; }
  function colName(i) { var s = ""; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  var MAX_COL = 16383, MAX_ROW = 1048576;

  // ---------- 字句（トークン）。どのトークンも元の文字（s）を持つ（shift・adjust で書き直すため） ----------
  var RE = {
    ws: /^\s+/,
    str: /^"((?:[^"]|"")*)"/,
    num: /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/,
    err: /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|CIRC!|NULL!)/i,
    qsheet: /^'((?:[^']|'')+)'!/,
    sheet: /^([A-Za-z_\p{L}][\w.\p{L}]*)!/u,
    cell: /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?![\w(])/,
    cellRange: /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6}):(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?![\w(])/,
    colRange: /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/,
    rowRange: /^(\$?)([1-9]\d{0,6}):(\$?)([1-9]\d{0,6})(?![\w(.])/,
    fn: /^([A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()/,
    name: /^[A-Za-z_\p{L}][\w.\p{L}]*/u,
    op: /^(?:<=|>=|<>|[-+*/^&=<>%])/
  };
  function tokenize(src) {
    var out = [], i = 0, m, rest;
    while (i < src.length) {
      rest = src.slice(i);
      if ((m = RE.ws.exec(rest))) { out.push({ t: "ws", s: m[0] }); i += m[0].length; continue; }
      if ((m = RE.str.exec(rest))) { out.push({ t: "str", s: m[0], v: m[1].replace(/""/g, '"') }); i += m[0].length; continue; }
      // シートの名前つきの参照
      var sheet = null, pre = "";
      if ((m = RE.qsheet.exec(rest)) || (m = RE.sheet.exec(rest))) {
        var after = rest.slice(m[0].length);
        if (RE.cellRange.test(after) || RE.cell.test(after) || RE.colRange.test(after) || RE.rowRange.test(after) || /^#REF!/i.test(after)) {
          sheet = m[1].replace(/''/g, "'"); pre = m[0]; rest = after;
        }
      }
      if ((m = RE.cellRange.exec(rest))) {
        out.push({ t: "range", s: pre + m[0], sheet: sheet, a: { ac: !!m[1], c: colIndex(m[2]), ar: !!m[3], r: +m[4] }, b: { ac: !!m[5], c: colIndex(m[6]), ar: !!m[7], r: +m[8] } });
        i += pre.length + m[0].length; continue;
      }
      if ((m = RE.cell.exec(rest)) && !/^(TRUE|FALSE)$/i.test(m[2] + m[4])) {
        out.push({ t: "ref", s: pre + m[0], sheet: sheet, a: { ac: !!m[1], c: colIndex(m[2]), ar: !!m[3], r: +m[4] } });
        i += pre.length + m[0].length; continue;
      }
      if ((m = RE.colRange.exec(rest)) && (sheet || out.length === 0 || /^[(,:=+\-*/^&<>; ]$/.test(lastChar(out)))) {
        out.push({ t: "range", s: pre + m[0], sheet: sheet, a: { ac: !!m[1], c: colIndex(m[2]), ar: false, r: 0 }, b: { ac: !!m[3], c: colIndex(m[4]), ar: false, r: 0 }, whole: "col" });
        i += pre.length + m[0].length; continue;
      }
      if ((m = RE.rowRange.exec(rest)) && (sheet || out.length === 0 || /^[(,:=+\-*/^&<>; ]$/.test(lastChar(out)))) {
        out.push({ t: "range", s: pre + m[0], sheet: sheet, a: { ac: false, c: 0, ar: !!m[1], r: +m[2] }, b: { ac: false, c: 0, ar: !!m[3], r: +m[4] }, whole: "row" });
        i += pre.length + m[0].length; continue;
      }
      if (pre) { out.push({ t: "err", s: pre + "#REF!", v: ERR.REF }); i += pre.length + 5; continue; }
      if ((m = RE.num.exec(rest))) { out.push({ t: "num", s: m[0], v: parseFloat(m[0]) }); i += m[0].length; continue; }
      if ((m = RE.err.exec(rest))) { out.push({ t: "err", s: m[0], v: m[0].toUpperCase() }); i += m[0].length; continue; }
      if ((m = RE.fn.exec(rest))) { out.push({ t: "fn", s: m[0], v: m[0].toUpperCase() }); i += m[0].length; continue; }
      if ((m = RE.name.exec(rest))) {
        var up = m[0].toUpperCase();
        out.push(up === "TRUE" || up === "FALSE" ? { t: "bool", s: m[0], v: up === "TRUE" } : { t: "name", s: m[0], v: m[0] });
        i += m[0].length; continue;
      }
      if ((m = RE.op.exec(rest))) { out.push({ t: "op", s: m[0], v: m[0] }); i += m[0].length; continue; }
      var ch = src.charAt(i);
      out.push({ t: ch === "(" || ch === ")" ? ch : ch === "," || ch === ";" ? "," : "bad", s: ch });
      i++;
    }
    return out;
  }
  function lastChar(toks) {
    for (var i = toks.length - 1; i >= 0; i--) { if (toks[i].t !== "ws") return toks[i].s.slice(-1); }
    return "";
  }

  // ---------- 構文（木にする） ----------
  var cache = {}, cacheN = 0;
  function parse(src) {
    if (cache[src]) return cache[src];
    var toks = tokenize(src).filter(function (x) { return x.t !== "ws"; }), p = 0, ast;
    function peek() { return toks[p]; }
    function isOp(v) { var x = toks[p]; return x && x.t === "op" && x.v === v; }
    function fail() { throw new Error("syntax"); }
    function cmp() {
      var a = concat();
      while (peek() && peek().t === "op" && /^(=|<>|<|>|<=|>=)$/.test(peek().v)) { var op = toks[p++].v; a = { k: "bin", op: op, a: a, b: concat() }; }
      return a;
    }
    function concat() { var a = add(); while (isOp("&")) { p++; a = { k: "bin", op: "&", a: a, b: add() }; } return a; }
    function add() { var a = mul(); while (isOp("+") || isOp("-")) { var op = toks[p++].v; a = { k: "bin", op: op, a: a, b: mul() }; } return a; }
    function mul() { var a = pow(); while (isOp("*") || isOp("/")) { var op = toks[p++].v; a = { k: "bin", op: op, a: a, b: pow() }; } return a; }
    function pow() { var a = unary(); while (isOp("^")) { p++; a = { k: "bin", op: "^", a: a, b: unary() }; } return a; }
    function unary() {
      if (isOp("-")) { p++; return { k: "neg", a: unary() }; }
      if (isOp("+")) { p++; return unary(); }
      return postfix();
    }
    function postfix() { var a = primary(); while (isOp("%")) { p++; a = { k: "pct", a: a }; } return a; }
    function primary() {
      var x = toks[p++];
      if (!x) fail();
      if (x.t === "num") return { k: "num", v: x.v };
      if (x.t === "str") return { k: "str", v: x.v };
      if (x.t === "bool") return { k: "bool", v: x.v };
      if (x.t === "err") return { k: "err", v: x.v };
      if (x.t === "name") return { k: "name", v: x.v };
      if (x.t === "ref") return { k: "ref", sheet: x.sheet, c: x.a.c, r: x.a.r };
      if (x.t === "range") {
        return { k: "range", sheet: x.sheet, c1: Math.min(x.a.c, x.b.c), r1: Math.min(x.a.r, x.b.r), c2: Math.max(x.a.c, x.b.c), r2: Math.max(x.a.r, x.b.r), whole: x.whole };
      }
      if (x.t === "(") { var e = cmp(); if (!toks[p] || toks[p].t !== ")") fail(); p++; return e; }
      if (x.t === "fn") {
        p++; // (
        var args = [];
        if (toks[p] && toks[p].t === ")") { p++; return { k: "fn", name: x.v, args: args }; }
        for (;;) {
          if (toks[p] && (toks[p].t === "," || toks[p].t === ")")) args.push({ k: "empty" });
          else args.push(cmp());
          var y = toks[p++];
          if (!y) fail();
          if (y.t === ")") break;
          if (y.t !== ",") fail();
        }
        return { k: "fn", name: x.v, args: args };
      }
      fail();
    }
    try {
      ast = toks.length ? cmp() : { k: "str", v: "" };
      if (p < toks.length) ast = { k: "err", v: ERR.NAME, bad: true };
    } catch (e) { ast = { k: "err", v: ERR.NAME, bad: true }; }
    if (++cacheN > 4000) { cache = {}; cacheN = 0; }
    cache[src] = ast;
    return ast;
  }
  function valid(src) { return !parse(src).bad; }

  // ---------- 型を変える ----------
  function num(v) {
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (v === null || v === undefined) return 0;
    if (isErr(v)) return v;
    if (typeof v === "string") {
      var n = parseNum(v);
      return n === null ? err(ERR.VALUE) : n;
    }
    return err(ERR.VALUE);
  }
  // 文字を数として読む（"1,200"・"¥300"・"50%"・"2026/10/21" も）。読めなければ null
  function parseNum(s) {
    s = String(s).trim();
    if (!s) return null;
    var m = /^([-+]?)[¥$￥]?\s*([\d,]*\.?\d+(?:[eE][+-]?\d+)?)\s*(%?)$/.exec(s.replace(/^[¥$￥]([-+])/, "$1"));
    if (m && (m[2].indexOf(",") < 0 || /^\d{1,3}(,\d{3})*(\.\d+)?$/.test(m[2]))) {
      var n = parseFloat(m[2].replace(/,/g, ""));
      if (m[1] === "-") n = -n;
      return m[3] ? n / 100 : n;
    }
    var d = parseDate(s);
    return d === null ? null : d;
  }
  function str(v) {
    if (typeof v === "string") return v;
    if (typeof v === "number") return numText(v);
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v === null || v === undefined) return "";
    if (isErr(v)) return v.err;
    return "";
  }
  function numText(v) {
    if (!isFinite(v)) return ERR.NUM;
    if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
    var s = String(parseFloat(v.toPrecision(15)));
    if (/e/.test(s) && Math.abs(v) >= 1e-9 && Math.abs(v) < 1e15) s = v.toFixed(10).replace(/\.?0+$/, "");
    return s;
  }
  function bool(v) {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (v === null || v === undefined) return false;
    if (isErr(v)) return v;
    if (typeof v === "string") { var u = v.toUpperCase(); if (u === "TRUE") return true; if (u === "FALSE") return false; }
    return err(ERR.VALUE);
  }
  // 値の比べ方（Excel と同じ: 数 < 文字 < TRUE/FALSE。文字は大文字と小文字を区別しない）
  function rank(v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 0; }
  function compare(a, b) {
    if (a === null || a === undefined) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
    if (b === null || b === undefined) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
    var ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (typeof a === "string") { a = a.toLowerCase(); b = b.toLowerCase(); }
    if (typeof a === "boolean") { a = a ? 1 : 0; b = b ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // ---------- 計算 ----------
  function evaluate(ast, ctx) {
    var v = ev(ast, ctx);
    if (isRange(v)) v = intersect(v, ctx);
    if (v === null || v === undefined) return 0;
    if (typeof v === "number" && !isFinite(v)) return err(ERR.NUM);
    return v;
  }
  // 範囲をひとつの値にする（同じ行・列のセル。1 つだけのときはそのセル）
  function intersect(rg, ctx) {
    if (rg.rows.length === 1 && rg.rows[0].length === 1) return rg.rows[0][0];
    if (rg.c1 !== undefined && ctx && ctx.here) {
      if (rg.c1 === rg.c2 && ctx.here.r >= rg.r1 && ctx.here.r <= rg.r2) return rg.rows[ctx.here.r - rg.r1][0];
      if (rg.r1 === rg.r2 && ctx.here.c >= rg.c1 && ctx.here.c <= rg.c2) return rg.rows[0][ctx.here.c - rg.c1];
    }
    return rg.rows.length && rg.rows[0].length ? rg.rows[0][0] : err(ERR.VALUE);
  }
  function scalar(v, ctx) { return isRange(v) ? intersect(v, ctx) : v; }
  function sheetOf(node, ctx) {
    if (!node.sheet) return ctx.here ? ctx.here.s : 0;
    return ctx.sheet(node.sheet);
  }
  function ev(n, ctx) {
    switch (n.k) {
      case "num": case "str": case "bool": return n.v;
      case "err": return err(n.v);
      case "empty": return null;
      case "name": return err(ERR.NAME);
      case "ref": {
        var s = sheetOf(n, ctx);
        if (s < 0) return err(ERR.REF);
        var v = ctx.get(s, n.c, n.r);
        return v === undefined ? null : v;
      }
      case "range": {
        var si = sheetOf(n, ctx);
        if (si < 0) return err(ERR.REF);
        var c1 = n.c1, c2 = n.c2, r1 = n.r1, r2 = n.r2;
        if (n.whole) {
          var last = ctx.last ? ctx.last(si) : { rows: 100, cols: 26 };
          if (n.whole === "col") { r1 = 1; r2 = Math.max(1, last.rows); }
          else { c1 = 0; c2 = Math.max(0, last.cols - 1); }
        }
        var rows = [];
        for (var r = r1; r <= r2; r++) {
          var row = [];
          for (var c = c1; c <= c2; c++) { var x = ctx.get(si, c, r); row.push(x === undefined ? null : x); }
          rows.push(row);
        }
        return { rows: rows, s: si, c1: c1, r1: r1, c2: c2, r2: r2 };
      }
      case "neg": return map1(ev(n.a, ctx), function (x) { x = num(x); return isErr(x) ? x : -x; });
      case "pct": return map1(ev(n.a, ctx), function (x) { x = num(x); return isErr(x) ? x : x / 100; });
      case "bin": return bin(n.op, ev(n.a, ctx), ev(n.b, ctx));
      case "fn": return call(n, ctx);
    }
    return err(ERR.VALUE);
  }
  function map1(v, f) {
    if (isRange(v)) return { rows: v.rows.map(function (row) { return row.map(f); }) };
    return f(v);
  }
  function bin(op, a, b) {
    if (isRange(a) || isRange(b)) {
      var A = isRange(a) ? a.rows : [[a]], B = isRange(b) ? b.rows : [[b]];
      var h = Math.max(A.length, B.length), w = Math.max(A[0] ? A[0].length : 0, B[0] ? B[0].length : 0), rows = [];
      for (var r = 0; r < h; r++) {
        var row = [];
        for (var c = 0; c < w; c++) {
          var x = (A[A.length === 1 ? 0 : r] || [])[A[0] && A[0].length === 1 ? 0 : c];
          var y = (B[B.length === 1 ? 0 : r] || [])[B[0] && B[0].length === 1 ? 0 : c];
          row.push(x === undefined || y === undefined ? err(ERR.NA) : bin1(op, x, y));
        }
        rows.push(row);
      }
      return { rows: rows };
    }
    return bin1(op, a, b);
  }
  function bin1(op, a, b) {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (op === "&") return str(a) + str(b);
    if (/^(=|<>|<|>|<=|>=)$/.test(op)) {
      var c = compare(a, b);
      return op === "=" ? c === 0 : op === "<>" ? c !== 0 : op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
    }
    var x = num(a), y = num(b);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    if (op === "+") return x + y;
    if (op === "-") return x - y;
    if (op === "*") return x * y;
    if (op === "/") return y === 0 ? err(ERR.DIV0) : x / y;
    if (op === "^") { var p = Math.pow(x, y); return isFinite(p) && !isNaN(p) ? p : err(ERR.NUM); }
    return err(ERR.VALUE);
  }

  // ---------- 関数 ----------
  // 範囲の中の数だけを集める（文字・空のセルは数えない）。直接書いた値は数にする。エラーがあればエラー
  function nums(args, opts) {
    var out = [];
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (isRange(a)) {
        for (var r = 0; r < a.rows.length; r++) for (var c = 0; c < a.rows[r].length; c++) {
          var v = a.rows[r][c];
          if (isErr(v)) return v;
          if (typeof v === "number") out.push(v);
          else if (opts && opts.all && typeof v === "boolean") out.push(v ? 1 : 0);
        }
      } else if (a !== null || (opts && opts.empty)) {
        var n = num(a);
        if (isErr(n)) return n;
        out.push(n);
      }
    }
    return out;
  }
  function flat(args) {
    var out = [];
    args.forEach(function (a) {
      if (isRange(a)) a.rows.forEach(function (row) { row.forEach(function (v) { out.push(v); }); });
      else out.push(a);
    });
    return out;
  }
  function sum(a) { var s = 0; for (var i = 0; i < a.length; i++) s += a[i]; return s; }
  function rows(v) { return isRange(v) ? v.rows : [[v]]; }
  function vec(v) { var r = rows(v); if (r.length === 1) return r[0].slice(); return r.map(function (x) { return x[0]; }); }
  // 条件（"10"・">=10"・"<>りんご"・"り*"）を関数にする
  function criteria(c) {
    if (isErr(c)) return function () { return false; };
    if (typeof c === "number" || typeof c === "boolean") return function (v) { return compare(v, c) === 0 && v !== null && typeof v === typeof c; };
    var s = str(c), m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s), op = m[1] || "=", val = m[2];
    var n = parseNum(val);
    if (val === "") {
      if (op === "=") return function (v) { return v === null || v === ""; };
      if (op === "<>") return function (v) { return v !== null && v !== ""; };
    }
    if (n !== null && !/[*?]/.test(val)) {
      return function (v) {
        if (typeof v !== "number") { if (op === "<>") return true; var pv = typeof v === "string" ? parseNum(v) : null; if (pv === null) return false; v = pv; }
        return op === "=" ? v === n : op === "<>" ? v !== n : op === "<" ? v < n : op === ">" ? v > n : op === "<=" ? v <= n : v >= n;
      };
    }
    if (op === "=" || op === "<>") {
      var re = wild(val);
      return function (v) { var ok = v !== null && re.test(str(v)); return op === "=" ? ok : !ok; };
    }
    return function (v) { if (typeof v !== "string") return false; var c2 = compare(v, val); return op === "<" ? c2 < 0 : op === ">" ? c2 > 0 : op === "<=" ? c2 <= 0 : c2 >= 0; };
  }
  function wild(p) {
    var s = "";
    for (var i = 0; i < p.length; i++) {
      var ch = p.charAt(i);
      if (ch === "~" && i + 1 < p.length) { s += p.charAt(++i).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
      s += ch === "*" ? "[\\s\\S]*" : ch === "?" ? "[\\s\\S]" : ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    return new RegExp("^" + s + "$", "i");
  }
  function ifs(args, start) {
    // [合計する範囲]、条件の範囲、条件、条件の範囲、条件…
    var pairs = [];
    for (var i = start; i + 1 < args.length; i += 2) pairs.push([rows(args[i]), criteria(scalar(args[i + 1]))]);
    return pairs;
  }
  function matchAll(pairs, r, c) {
    for (var i = 0; i < pairs.length; i++) { var row = pairs[i][0][r]; if (!row || !pairs[i][1](row[c] === undefined ? null : row[c])) return false; }
    return true;
  }
  function round(x, d, mode) {
    var f = Math.pow(10, d), y = x * f;
    y = parseFloat(y.toPrecision(15));
    var r = mode === "up" ? (y < 0 ? -Math.ceil(-y) : Math.ceil(y)) : mode === "down" ? (y < 0 ? -Math.floor(-y) : Math.floor(y)) : (y < 0 ? -Math.round(-y) : Math.round(y));
    return r / f;
  }
  function lookupIndex(v, list, mode) {
    // mode: 0 完全一致・1 以下で最大（並んでいる）・-1 以上で最小
    var i, found = -1;
    if (mode === 0) {
      var isWild = typeof v === "string" && /[*?~]/.test(v), re = isWild ? wild(v) : null;
      for (i = 0; i < list.length; i++) {
        if (re ? list[i] !== null && re.test(str(list[i])) : list[i] !== null && rank(list[i]) === rank(v) && compare(list[i], v) === 0) return i;
      }
      return -1;
    }
    for (i = 0; i < list.length; i++) {
      if (list[i] === null || rank(list[i]) !== rank(v)) continue;
      var c = compare(list[i], v);
      if (mode === 1) { if (c <= 0) found = i; else break; }
      else { if (c >= 0) found = i; else break; }
    }
    return found;
  }
  // 日付（Excel と同じ。1899/12/30 からの日数）
  var DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
  function serial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY); }
  function fromSerial(n) { return new Date(EPOCH + Math.floor(n) * DAY + Math.round((n - Math.floor(n)) * DAY)); }
  function parseDate(s) {
    s = String(s).trim();
    var m = /^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    if (m) {
      var y = +m[1], mo = +m[2], d = +m[3];
      if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
      var v = serial(y, mo, d);
      if (fromSerial(v).getUTCDate() !== d) return null;
      if (m[4] !== undefined) v += (+m[4] * 3600 + +m[5] * 60 + (+m[6] || 0)) / 86400;
      return v;
    }
    m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
    if (m && +m[1] < 48 && +m[2] < 60) return (+m[1] * 3600 + +m[2] * 60 + (+m[3] || 0)) / 86400;
    return null;
  }
  function today() { var d = new Date(); return serial(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function dateParts(v) { v = num(v); if (isErr(v)) return v; var d = fromSerial(v); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds() }; }

  function N(f) { return function (args) { var x = num(scalar(args[0])); return isErr(x) ? x : f(x, args); }; }
  function agg(f, opts) { return function (args) { var a = nums(args, opts); return isErr(a) ? a : f(a); }; }
  var FN = {
    SUM: agg(sum),
    AVERAGE: agg(function (a) { return a.length ? sum(a) / a.length : err(ERR.DIV0); }),
    AVERAGEA: agg(function (a) { return a.length ? sum(a) / a.length : err(ERR.DIV0); }, { all: true }),
    MIN: agg(function (a) { return a.length ? Math.min.apply(null, a) : 0; }),
    MAX: agg(function (a) { return a.length ? Math.max.apply(null, a) : 0; }),
    PRODUCT: agg(function (a) { return a.length ? a.reduce(function (x, y) { return x * y; }, 1) : 0; }),
    COUNT: function (args) { return flat(args).filter(function (v, i) { return typeof v === "number" || (!isRange(args[i]) && typeof v === "string" && parseNum(v) !== null); }).length; },
    COUNTA: function (args) { return flat(args).filter(function (v) { return v !== null && v !== ""; }).length; },
    COUNTBLANK: function (args) { return flat(args).filter(function (v) { return v === null || v === ""; }).length; },
    MEDIAN: agg(function (a) { if (!a.length) return err(ERR.NUM); a.sort(function (x, y) { return x - y; }); var m = a.length >> 1; return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2; }),
    MODE: agg(function (a) { var cnt = {}, best = null, bc = 1; a.forEach(function (x) { cnt[x] = (cnt[x] || 0) + 1; if (cnt[x] > bc) { bc = cnt[x]; best = x; } }); return best === null ? err(ERR.NA) : best; }),
    STDEV: agg(function (a) { return a.length < 2 ? err(ERR.DIV0) : Math.sqrt(variance(a, 1)); }),
    "STDEV.S": agg(function (a) { return a.length < 2 ? err(ERR.DIV0) : Math.sqrt(variance(a, 1)); }),
    STDEVP: agg(function (a) { return a.length < 1 ? err(ERR.DIV0) : Math.sqrt(variance(a, 0)); }),
    "STDEV.P": agg(function (a) { return a.length < 1 ? err(ERR.DIV0) : Math.sqrt(variance(a, 0)); }),
    VAR: agg(function (a) { return a.length < 2 ? err(ERR.DIV0) : variance(a, 1); }),
    "VAR.S": agg(function (a) { return a.length < 2 ? err(ERR.DIV0) : variance(a, 1); }),
    "VAR.P": agg(function (a) { return a.length < 1 ? err(ERR.DIV0) : variance(a, 0); }),
    LARGE: function (args) { var a = nums([args[0]]), k = num(scalar(args[1])); if (isErr(a)) return a; if (isErr(k)) return k; a.sort(function (x, y) { return y - x; }); return k >= 1 && k <= a.length ? a[Math.ceil(k) - 1] : err(ERR.NUM); },
    SMALL: function (args) { var a = nums([args[0]]), k = num(scalar(args[1])); if (isErr(a)) return a; if (isErr(k)) return k; a.sort(function (x, y) { return x - y; }); return k >= 1 && k <= a.length ? a[Math.ceil(k) - 1] : err(ERR.NUM); },
    RANK: rankFn, "RANK.EQ": rankFn,
    SUMPRODUCT: function (args) {
      var arrs = args.map(rows), h = arrs[0].length, w = arrs[0][0] ? arrs[0][0].length : 0, s = 0;
      for (var i = 1; i < arrs.length; i++) if (arrs[i].length !== h || (arrs[i][0] || []).length !== w) return err(ERR.VALUE);
      for (var r = 0; r < h; r++) for (var c = 0; c < w; c++) {
        var p = 1;
        for (var k = 0; k < arrs.length; k++) { var v = arrs[k][r][c]; if (isErr(v)) return v; p *= typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0; }
        s += p;
      }
      return s;
    },
    ROUND: function (args) { var x = num(scalar(args[0])), d = args.length > 1 ? num(scalar(args[1])) : 0; if (isErr(x)) return x; if (isErr(d)) return d; return round(x, Math.trunc(d)); },
    ROUNDUP: function (args) { var x = num(scalar(args[0])), d = args.length > 1 ? num(scalar(args[1])) : 0; if (isErr(x)) return x; if (isErr(d)) return d; return round(x, Math.trunc(d), "up"); },
    ROUNDDOWN: function (args) { var x = num(scalar(args[0])), d = args.length > 1 ? num(scalar(args[1])) : 0; if (isErr(x)) return x; if (isErr(d)) return d; return round(x, Math.trunc(d), "down"); },
    TRUNC: function (args) { var x = num(scalar(args[0])), d = args.length > 1 ? num(scalar(args[1])) : 0; if (isErr(x)) return x; if (isErr(d)) return d; return round(x, Math.trunc(d), "down"); },
    INT: N(Math.floor), ABS: N(Math.abs), SIGN: N(Math.sign), EXP: N(Math.exp),
    SQRT: N(function (x) { return x < 0 ? err(ERR.NUM) : Math.sqrt(x); }),
    LN: N(function (x) { return x <= 0 ? err(ERR.NUM) : Math.log(x); }),
    LOG10: N(function (x) { return x <= 0 ? err(ERR.NUM) : Math.log10(x); }),
    LOG: function (args) { var x = num(scalar(args[0])), b = args.length > 1 ? num(scalar(args[1])) : 10; if (isErr(x)) return x; if (isErr(b)) return b; return x <= 0 || b <= 0 || b === 1 ? err(ERR.NUM) : Math.log(x) / Math.log(b); },
    MOD: function (args) { var x = num(scalar(args[0])), y = num(scalar(args[1])); if (isErr(x)) return x; if (isErr(y)) return y; if (y === 0) return err(ERR.DIV0); return x - y * Math.floor(x / y); },
    POWER: function (args) { return bin1("^", scalar(args[0]), scalar(args[1])); },
    QUOTIENT: function (args) { var x = num(scalar(args[0])), y = num(scalar(args[1])); if (isErr(x)) return x; if (isErr(y)) return y; return y === 0 ? err(ERR.DIV0) : Math.trunc(x / y); },
    PI: function () { return Math.PI; },
    RAND: function () { return Math.random(); },
    RANDBETWEEN: function (args) { var a = num(scalar(args[0])), b = num(scalar(args[1])); if (isErr(a)) return a; if (isErr(b)) return b; a = Math.ceil(a); b = Math.floor(b); return a + Math.floor(Math.random() * (b - a + 1)); },
    CEILING: function (args) { var x = num(scalar(args[0])), s = args.length > 1 ? num(scalar(args[1])) : 1; if (isErr(x)) return x; if (isErr(s)) return s; return s === 0 ? 0 : Math.ceil(parseFloat((x / s).toPrecision(15))) * s; },
    FLOOR: function (args) { var x = num(scalar(args[0])), s = args.length > 1 ? num(scalar(args[1])) : 1; if (isErr(x)) return x; if (isErr(s)) return s; return s === 0 ? err(ERR.DIV0) : Math.floor(parseFloat((x / s).toPrecision(15))) * s; },
    EVEN: N(function (x) { var y = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -y : y; }),
    ODD: N(function (x) { var y = Math.ceil(Math.abs(x)); if (y % 2 === 0) y++; return x < 0 ? -y : y; }),
    // 論理
    IF: { lazy: true, f: function (a, ctx) {
      var c = bool(scalar(ev(a[0], ctx), ctx));
      if (isErr(c)) return c;
      return c ? (a.length > 1 ? ev(a[1], ctx) : true) : (a.length > 2 ? ev(a[2], ctx) : false);
    } },
    IFS: { lazy: true, f: function (a, ctx) {
      for (var i = 0; i + 1 < a.length; i += 2) { var c = bool(scalar(ev(a[i], ctx), ctx)); if (isErr(c)) return c; if (c) return ev(a[i + 1], ctx); }
      return err(ERR.NA);
    } },
    IFERROR: { lazy: true, f: function (a, ctx) { var v = ev(a[0], ctx); return isErr(scalar(v, ctx)) ? ev(a[1], ctx) : v; } },
    IFNA: { lazy: true, f: function (a, ctx) { var v = ev(a[0], ctx), s = scalar(v, ctx); return isErr(s) && s.err === ERR.NA ? ev(a[1], ctx) : v; } },
    SWITCH: { lazy: true, f: function (a, ctx) {
      var v = scalar(ev(a[0], ctx), ctx);
      if (isErr(v)) return v;
      for (var i = 1; i + 1 < a.length; i += 2) { if (compare(scalar(ev(a[i], ctx), ctx), v) === 0) return ev(a[i + 1], ctx); }
      return a.length % 2 === 0 ? ev(a[a.length - 1], ctx) : err(ERR.NA);
    } },
    CHOOSE: { lazy: true, f: function (a, ctx) { var i = num(scalar(ev(a[0], ctx), ctx)); if (isErr(i)) return i; i = Math.trunc(i); return i >= 1 && i < a.length ? ev(a[i], ctx) : err(ERR.VALUE); } },
    AND: function (args) { var vs = flat(args), any = false; for (var i = 0; i < vs.length; i++) { var v = vs[i]; if (v === null || typeof v === "string") continue; var b = bool(v); if (isErr(b)) return b; any = true; if (!b) return false; } return any ? true : err(ERR.VALUE); },
    OR: function (args) { var vs = flat(args), any = false; for (var i = 0; i < vs.length; i++) { var v = vs[i]; if (v === null || typeof v === "string") continue; var b = bool(v); if (isErr(b)) return b; any = true; if (b) return true; } return any ? false : err(ERR.VALUE); },
    XOR: function (args) { var n = 0; flat(args).forEach(function (v) { if (v !== null && typeof v !== "string" && bool(v) === true) n++; }); return n % 2 === 1; },
    NOT: function (args) { var b = bool(scalar(args[0])); return isErr(b) ? b : !b; },
    TRUE: function () { return true; }, FALSE: function () { return false; },
    NA: function () { return err(ERR.NA); },
    // 条件で数える・足す
    COUNTIF: function (args) { var f = criteria(scalar(args[1])); return flat([args[0]]).filter(function (v) { return f(v); }).length; },
    COUNTIFS: function (args) {
      var pairs = ifs(args, 0); if (!pairs.length) return 0;
      var h = pairs[0][0].length, w = pairs[0][0][0].length, n = 0;
      for (var r = 0; r < h; r++) for (var c = 0; c < w; c++) if (matchAll(pairs, r, c)) n++;
      return n;
    },
    SUMIF: function (args) {
      var f = criteria(scalar(args[1])), test = rows(args[0]), val = args.length > 2 ? rows(args[2]) : test, s = 0;
      for (var r = 0; r < test.length; r++) for (var c = 0; c < test[r].length; c++) {
        if (f(test[r][c])) { var v = (val[r] || [])[c]; if (isErr(v)) return v; if (typeof v === "number") s += v; }
      }
      return s;
    },
    SUMIFS: function (args) { return ifsAgg(args, function (a) { return sum(a); }); },
    AVERAGEIF: function (args) {
      var f = criteria(scalar(args[1])), test = rows(args[0]), val = args.length > 2 ? rows(args[2]) : test, a = [];
      for (var r = 0; r < test.length; r++) for (var c = 0; c < test[r].length; c++) {
        if (f(test[r][c])) { var v = (val[r] || [])[c]; if (typeof v === "number") a.push(v); }
      }
      return a.length ? sum(a) / a.length : err(ERR.DIV0);
    },
    AVERAGEIFS: function (args) { return ifsAgg(args, function (a) { return a.length ? sum(a) / a.length : err(ERR.DIV0); }); },
    MAXIFS: function (args) { return ifsAgg(args, function (a) { return a.length ? Math.max.apply(null, a) : 0; }); },
    MINIFS: function (args) { return ifsAgg(args, function (a) { return a.length ? Math.min.apply(null, a) : 0; }); },
    // 探す
    VLOOKUP: function (args) {
      var v = scalar(args[0]), t = rows(args[1]), col = num(scalar(args[2])), approx = args.length < 4 || args[3] === null ? true : bool(scalar(args[3]));
      if (isErr(v)) return v; if (isErr(col)) return col;
      col = Math.trunc(col);
      if (col < 1 || col > (t[0] || []).length) return err(ERR.REF);
      var i = lookupIndex(v, t.map(function (r) { return r[0]; }), approx ? 1 : 0);
      return i < 0 ? err(ERR.NA) : t[i][col - 1];
    },
    HLOOKUP: function (args) {
      var v = scalar(args[0]), t = rows(args[1]), row = num(scalar(args[2])), approx = args.length < 4 || args[3] === null ? true : bool(scalar(args[3]));
      if (isErr(v)) return v; if (isErr(row)) return row;
      row = Math.trunc(row);
      if (row < 1 || row > t.length) return err(ERR.REF);
      var i = lookupIndex(v, t[0], approx ? 1 : 0);
      return i < 0 ? err(ERR.NA) : t[row - 1][i];
    },
    XLOOKUP: function (args) {
      var v = scalar(args[0]), look = vec(args[1]), ret = rows(args[2]), nf = args.length > 3 && args[3] !== null ? args[3] : err(ERR.NA);
      var mode = args.length > 4 && args[4] !== null ? num(scalar(args[4])) : 0, back = args.length > 5 && args[5] !== null && num(scalar(args[5])) < 0;
      if (isErr(v)) return v;
      var idx = -1, i;
      var order = look.map(function (_, k) { return k; });
      if (back) order.reverse();
      if (mode === 2 || mode === 0) {
        var re = mode === 2 && typeof v === "string" ? wild(v) : null;
        for (i = 0; i < order.length; i++) { var x = look[order[i]]; if (x !== null && (re ? re.test(str(x)) : rank(x) === rank(v) && compare(x, v) === 0)) { idx = order[i]; break; } }
      } else {
        var best = null;
        for (i = 0; i < look.length; i++) {
          var y = look[i]; if (y === null || rank(y) !== rank(v)) continue;
          var c = compare(y, v);
          if (c === 0) { idx = i; break; }
          if (mode < 0 && c < 0 && (best === null || compare(y, look[best]) > 0)) best = i;
          if (mode > 0 && c > 0 && (best === null || compare(y, look[best]) < 0)) best = i;
        }
        if (idx < 0 && best !== null) idx = best;
      }
      if (idx < 0) return nf;
      if (ret.length === look.length && ret.length > 1) return ret[idx].length === 1 ? ret[idx][0] : { rows: [ret[idx]] };
      if (ret.length === 1) return ret[0][idx];
      return ret[idx] ? ret[idx][0] : err(ERR.REF);
    },
    MATCH: function (args) {
      var v = scalar(args[0]), list = vec(args[1]), t = args.length > 2 && args[2] !== null ? num(scalar(args[2])) : 1;
      if (isErr(v)) return v;
      var i = lookupIndex(v, list, t === 0 ? 0 : t > 0 ? 1 : -1);
      return i < 0 ? err(ERR.NA) : i + 1;
    },
    INDEX: function (args) {
      var t = rows(args[0]), r = args.length > 1 && args[1] !== null ? num(scalar(args[1])) : 0, c = args.length > 2 && args[2] !== null ? num(scalar(args[2])) : 0;
      if (isErr(r)) return r; if (isErr(c)) return c;
      r = Math.trunc(r); c = Math.trunc(c);
      if (t.length === 1 && args.length === 2) { c = r; r = 1; }
      if (r < 0 || c < 0 || r > t.length || c > (t[0] || []).length) return err(ERR.REF);
      if (r === 0 && c === 0) return { rows: t };
      if (r === 0) return { rows: t.map(function (x) { return [x[c - 1]]; }) };
      if (c === 0) return t[0].length === 1 ? t[r - 1][0] : { rows: [t[r - 1]] };
      return t[r - 1][c - 1];
    },
    ROW: { ctx: true, f: function (args, ctx) { return args.length && isRange(args[0]) && args[0].r1 ? args[0].r1 : ctx.here ? ctx.here.r : 1; } },
    COLUMN: { ctx: true, f: function (args, ctx) { return args.length && isRange(args[0]) && args[0].c1 !== undefined ? args[0].c1 + 1 : ctx.here ? ctx.here.c + 1 : 1; } },
    ROWS: function (args) { return rows(args[0]).length; },
    COLUMNS: function (args) { return (rows(args[0])[0] || []).length; },
    TRANSPOSE: function (args) { var t = rows(args[0]); return { rows: (t[0] || []).map(function (_, c) { return t.map(function (row) { return row[c]; }); }) }; },
    // 文字
    LEN: function (args) { var v = scalar(args[0]); return isErr(v) ? v : Array.from(str(v)).length; },
    LEFT: function (args) { var v = scalar(args[0]), n = args.length > 1 ? num(scalar(args[1])) : 1; if (isErr(v)) return v; if (isErr(n)) return n; return Array.from(str(v)).slice(0, Math.max(0, n)).join(""); },
    RIGHT: function (args) { var v = scalar(args[0]), n = args.length > 1 ? num(scalar(args[1])) : 1; if (isErr(v)) return v; if (isErr(n)) return n; var a = Array.from(str(v)); return n <= 0 ? "" : a.slice(-n).join(""); },
    MID: function (args) { var v = scalar(args[0]), s = num(scalar(args[1])), n = num(scalar(args[2])); if (isErr(v)) return v; if (isErr(s)) return s; if (isErr(n)) return n; if (s < 1) return err(ERR.VALUE); return Array.from(str(v)).slice(s - 1, s - 1 + n).join(""); },
    UPPER: function (args) { var v = scalar(args[0]); return isErr(v) ? v : str(v).toUpperCase(); },
    LOWER: function (args) { var v = scalar(args[0]); return isErr(v) ? v : str(v).toLowerCase(); },
    PROPER: function (args) { var v = scalar(args[0]); return isErr(v) ? v : str(v).toLowerCase().replace(/(^|[^a-z])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); }); },
    TRIM: function (args) { var v = scalar(args[0]); return isErr(v) ? v : str(v).replace(/[ 　]+/g, " ").trim(); },
    CONCAT: function (args) { var vs = flat(args), s = ""; for (var i = 0; i < vs.length; i++) { if (isErr(vs[i])) return vs[i]; s += str(vs[i]); } return s; },
    CONCATENATE: function (args) { var s = ""; for (var i = 0; i < args.length; i++) { var v = scalar(args[i]); if (isErr(v)) return v; s += str(v); } return s; },
    TEXTJOIN: function (args) {
      var d = str(scalar(args[0])), skip = bool(scalar(args[1])), vs = flat(args.slice(2)), out = [];
      for (var i = 0; i < vs.length; i++) { if (isErr(vs[i])) return vs[i]; if (skip && (vs[i] === null || vs[i] === "")) continue; out.push(str(vs[i])); }
      return out.join(d);
    },
    TEXT: function (args) { var v = scalar(args[0]), f = str(scalar(args[1])); if (isErr(v)) return v; var n = typeof v === "number" ? v : parseNum(str(v)); return n === null ? str(v) : format(n, f).text; },
    VALUE: function (args) { var v = scalar(args[0]); if (isErr(v) || typeof v === "number") return v; var n = parseNum(str(v)); return n === null ? err(ERR.VALUE) : n; },
    FIND: function (args) { var f = str(scalar(args[0])), s = str(scalar(args[1])), st = args.length > 2 ? num(scalar(args[2])) : 1; var i = s.indexOf(f, st - 1); return i < 0 ? err(ERR.VALUE) : i + 1; },
    SEARCH: function (args) {
      var f = str(scalar(args[0])), s = str(scalar(args[1])), st = args.length > 2 ? num(scalar(args[2])) : 1;
      var re = new RegExp(wild(f).source.slice(1, -1), "i"), m = re.exec(s.slice(st - 1));
      return m ? m.index + st : err(ERR.VALUE);
    },
    SUBSTITUTE: function (args) {
      var s = str(scalar(args[0])), o = str(scalar(args[1])), n = str(scalar(args[2])), k = args.length > 3 ? num(scalar(args[3])) : 0;
      if (!o) return s;
      if (!k) return s.split(o).join(n);
      var i = -1; for (var j = 0; j < k; j++) { i = s.indexOf(o, i + 1); if (i < 0) return s; }
      return s.slice(0, i) + n + s.slice(i + o.length);
    },
    REPLACE: function (args) { var s = Array.from(str(scalar(args[0]))), st = num(scalar(args[1])), n = num(scalar(args[2])), r = str(scalar(args[3])); return s.slice(0, st - 1).join("") + r + s.slice(st - 1 + n).join(""); },
    REPT: function (args) { var s = str(scalar(args[0])), n = num(scalar(args[1])); if (isErr(n)) return n; return n < 0 ? err(ERR.VALUE) : new Array(Math.floor(n) + 1).join(s); },
    EXACT: function (args) { return str(scalar(args[0])) === str(scalar(args[1])); },
    CHAR: N(function (x) { return String.fromCharCode(x); }),
    CODE: function (args) { var s = str(scalar(args[0])); return s ? s.charCodeAt(0) : err(ERR.VALUE); },
    T: function (args) { var v = scalar(args[0]); return typeof v === "string" ? v : ""; },
    N: function (args) { var v = scalar(args[0]); return typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0; },
    // 日付と時間
    TODAY: function () { return today(); },
    NOW: function () { var d = new Date(); return serial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
    DATE: function (args) {
      var y = num(scalar(args[0])), m = num(scalar(args[1])), d = num(scalar(args[2]));
      if (isErr(y)) return y; if (isErr(m)) return m; if (isErr(d)) return d;
      if (y < 100) y += 1900;
      return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY);
    },
    TIME: function (args) { var h = num(scalar(args[0])), m = num(scalar(args[1])), s = num(scalar(args[2])); if (isErr(h)) return h; if (isErr(m)) return m; if (isErr(s)) return s; return ((h * 3600 + m * 60 + s) / 86400) % 1; },
    YEAR: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.y; },
    MONTH: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.m; },
    DAY: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.d; },
    HOUR: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.h; },
    MINUTE: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.mi; },
    SECOND: function (args) { var p = dateParts(scalar(args[0])); return isErr(p) ? p : p.s; },
    WEEKDAY: function (args) {
      var p = dateParts(scalar(args[0])), t = args.length > 1 ? num(scalar(args[1])) : 1;
      if (isErr(p)) return p;
      return t === 2 ? (p.wd + 6) % 7 + 1 : t === 3 ? (p.wd + 6) % 7 : p.wd + 1;
    },
    DAYS: function (args) { var a = num(scalar(args[0])), b = num(scalar(args[1])); if (isErr(a)) return a; if (isErr(b)) return b; return Math.floor(a) - Math.floor(b); },
    EDATE: function (args) { var p = dateParts(scalar(args[0])), n = num(scalar(args[1])); if (isErr(p)) return p; if (isErr(n)) return n; var d = new Date(Date.UTC(p.y, p.m - 1 + Math.trunc(n), 1)); var last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate(); return serial(d.getUTCFullYear(), d.getUTCMonth() + 1, Math.min(p.d, last)); },
    EOMONTH: function (args) { var p = dateParts(scalar(args[0])), n = num(scalar(args[1])); if (isErr(p)) return p; if (isErr(n)) return n; return Math.round((Date.UTC(p.y, p.m + Math.trunc(n), 0) - EPOCH) / DAY); },
    DATEDIF: function (args) {
      var a = dateParts(scalar(args[0])), b = dateParts(scalar(args[1])), u = str(scalar(args[2])).toUpperCase();
      if (isErr(a)) return a; if (isErr(b)) return b;
      var sa = serial(a.y, a.m, a.d), sb = serial(b.y, b.m, b.d);
      if (sa > sb) return err(ERR.NUM);
      var months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
      if (u === "D") return sb - sa;
      if (u === "M") return months;
      if (u === "Y") return Math.floor(months / 12);
      if (u === "YM") return months % 12;
      if (u === "MD") return b.d >= a.d ? b.d - a.d : sb - serial(b.y, b.m - 1, a.d);
      return err(ERR.NUM);
    },
    NETWORKDAYS: function (args) {
      var a = num(scalar(args[0])), b = num(scalar(args[1])); if (isErr(a)) return a; if (isErr(b)) return b;
      var hol = args.length > 2 ? nums([args[2]]) : []; if (isErr(hol)) hol = [];
      var s = Math.min(a, b), e = Math.max(a, b), n = 0;
      for (var d = Math.floor(s); d <= e; d++) { var wd = fromSerial(d).getUTCDay(); if (wd !== 0 && wd !== 6 && hol.indexOf(d) < 0) n++; }
      return a <= b ? n : -n;
    },
    WEEKNUM: function (args) { var p = dateParts(scalar(args[0])); if (isErr(p)) return p; var jan1 = serial(p.y, 1, 1), wd = fromSerial(jan1).getUTCDay(); return Math.floor((serial(p.y, p.m, p.d) - jan1 + wd) / 7) + 1; },
    // しらべる
    ISBLANK: function (args) { var v = scalar(args[0]); return v === null; },
    ISNUMBER: function (args) { return typeof scalar(args[0]) === "number"; },
    ISTEXT: function (args) { return typeof scalar(args[0]) === "string"; },
    ISNONTEXT: function (args) { return typeof scalar(args[0]) !== "string"; },
    ISLOGICAL: function (args) { return typeof scalar(args[0]) === "boolean"; },
    ISERROR: function (args) { return isErr(scalar(args[0])); },
    ISERR: function (args) { var v = scalar(args[0]); return isErr(v) && v.err !== ERR.NA; },
    ISNA: function (args) { var v = scalar(args[0]); return isErr(v) && v.err === ERR.NA; },
    ISEVEN: N(function (x) { return Math.floor(Math.abs(x)) % 2 === 0; }),
    ISODD: N(function (x) { return Math.floor(Math.abs(x)) % 2 === 1; })
  };
  function variance(a, ddof) { var m = sum(a) / a.length, s = 0; a.forEach(function (x) { s += (x - m) * (x - m); }); return s / (a.length - ddof); }
  function rankFn(args) {
    var v = num(scalar(args[0])), list = nums([args[1]]), asc = args.length > 2 && args[2] !== null && num(scalar(args[2])) !== 0;
    if (isErr(v)) return v; if (isErr(list)) return list;
    if (list.indexOf(v) < 0) return err(ERR.NA);
    return 1 + list.filter(function (x) { return asc ? x < v : x > v; }).length;
  }
  function ifsAgg(args, f) {
    var val = rows(args[0]), pairs = ifs(args, 1), a = [];
    for (var r = 0; r < val.length; r++) for (var c = 0; c < val[r].length; c++) {
      if (matchAll(pairs, r, c)) { var v = val[r][c]; if (isErr(v)) return v; if (typeof v === "number") a.push(v); }
    }
    return f(a);
  }
  function call(n, ctx) {
    var def = FN[n.name];
    if (!def) {
      // _xlfn.XLOOKUP など（Excel のファイルの新しい関数の書き方）
      var alt = n.name.replace(/^_XLFN\.(_XLWS\.)?/, "");
      def = FN[alt];
      if (!def) return err(ERR.NAME);
    }
    if (def.lazy) return def.f(n.args, ctx);
    var args = n.args.map(function (a) { return ev(a, ctx); });
    if (def.ctx) return def.f(args, ctx);
    return def(args, ctx);
  }

  // ---------- 参照を書き直す（コピー・オートフィル・行や列の挿入と削除・シートの名前） ----------
  function refText(a, ac, ar) { return (ac ? "$" : "") + colName(a.c) + (ar ? "$" : "") + a.r; }
  function sheetText(name) {
    if (!name) return "";
    return (/^[A-Za-z_\p{L}][\w\p{L}]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'") + "!";
  }
  function rewrite(src, fn) {
    return tokenize(src).map(function (x) {
      if (x.t !== "ref" && x.t !== "range") return x.s;
      var r = fn(x);
      return r === undefined ? x.s : r;
    }).join("");
  }
  function printToken(x) {
    var pre = sheetText(x.sheet);
    if (x.t === "ref") return pre + refText(x.a, x.a.ac, x.a.ar);
    if (x.whole === "col") return pre + (x.a.ac ? "$" : "") + colName(x.a.c) + ":" + (x.b.ac ? "$" : "") + colName(x.b.c);
    if (x.whole === "row") return pre + (x.a.ar ? "$" : "") + x.a.r + ":" + (x.b.ar ? "$" : "") + x.b.r;
    return pre + refText(x.a, x.a.ac, x.a.ar) + ":" + refText(x.b, x.b.ac, x.b.ar);
  }
  // コピーしたときに、$ の付いていない参照を dc 列・dr 行ずらす
  function shift(src, dc, dr) {
    if (!dc && !dr) return src;
    return rewrite(src, function (x) {
      var pts = x.t === "ref" ? [x.a] : [x.a, x.b], bad = false;
      pts.forEach(function (p) {
        if (x.whole !== "row" && !p.ac) p.c += dc;
        if (x.whole !== "col" && !p.ar) p.r += dr;
        if (p.c < 0 || p.c > MAX_COL || (x.whole !== "col" && (p.r < 1 || p.r > MAX_ROW))) bad = true;
      });
      return bad ? sheetText(x.sheet) + "#REF!" : printToken(x);
    });
  }
  // 行・列を入れた・消したときに参照を直す。info: { sheet: そのシートの名前, here: 数式があるシートの名前, axis: "r"|"c", at, n }
  //   at は入れた（消した）最初の行の番号（1 から）か列の番号（0 から）。n は入れた数（消したときは −消した数）
  function adjust(src, info) {
    return rewrite(src, function (x) {
      var target = x.sheet ? x.sheet : info.here;
      if (target !== info.sheet) return undefined;
      var key = info.axis === "r" ? "r" : "c";
      if ((key === "r" && x.whole === "col") || (key === "c" && x.whole === "row")) return undefined;
      var at = info.at, n = info.n;
      if (x.t === "ref") {
        var v = x.a[key];
        if (n > 0) { if (v >= at) x.a[key] = v + n; }
        else { var end = at - n - 1; if (v >= at && v <= end) return sheetText(x.sheet) + "#REF!"; if (v > end) x.a[key] = v + n; }
        return printToken(x);
      }
      var a = Math.min(x.a[key], x.b[key]), b = Math.max(x.a[key], x.b[key]);
      if (n > 0) { if (a >= at) a += n; if (b >= at) b += n; }
      else {
        var e = at - n - 1;
        if (a >= at && b <= e) return sheetText(x.sheet) + "#REF!";
        if (a > e) a += n; else if (a >= at) a = at;
        if (b > e) b += n; else if (b >= at) b = at - 1;
      }
      x.a[key] = a; x.b[key] = b;
      return printToken(x);
    });
  }
  function renameSheet(src, from, to) {
    return rewrite(src, function (x) {
      if (!x.sheet || x.sheet.toLowerCase() !== from.toLowerCase()) return undefined;
      x.sheet = to;
      return printToken(x);
    });
  }
  // 数式の中の参照（色分け・範囲の枠に使う）: [{ sheet, c1, r1, c2, r2, start, end }]
  function refs(src) {
    var out = [], pos = 0;
    tokenize(src).forEach(function (x) {
      if (x.t === "ref") out.push({ sheet: x.sheet, c1: x.a.c, r1: x.a.r, c2: x.a.c, r2: x.a.r, start: pos, end: pos + x.s.length });
      else if (x.t === "range" && !x.whole) out.push({ sheet: x.sheet, c1: Math.min(x.a.c, x.b.c), r1: Math.min(x.a.r, x.b.r), c2: Math.max(x.a.c, x.b.c), r2: Math.max(x.a.r, x.b.r), start: pos, end: pos + x.s.length });
      pos += x.s.length;
    });
    return out;
  }

  // ---------- 表示形式 ----------
  //   format(数, 形式) → { text, color }。形式は Excel と同じ書き方（General・0・0.00・#,##0・¥#,##0・0%・yyyy/m/d・h:mm・"文字"・[Red] など）
  var WEEK_JA = ["日", "月", "火", "水", "木", "金", "土"], WEEK_EN = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MON_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var fmtCache = {};
  function general(v) {
    if (!isFinite(v)) return ERR.NUM;
    var a = Math.abs(v);
    if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, "E").replace("e", "E");
    return numText(parseFloat(v.toPrecision(11)));
  }
  function splitSections(code) {
    var out = [], cur = "", q = false;
    for (var i = 0; i < code.length; i++) {
      var ch = code.charAt(i);
      if (ch === '"') q = !q;
      if (ch === "\\" && !q) { cur += ch + code.charAt(++i); continue; }
      if (ch === ";" && !q) { out.push(cur); cur = ""; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  function compile(sec) {
    var color = null, cond = null;
    sec = sec.replace(/\[([^\]]*)\]/g, function (m, x) {
      if (/^(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan|赤|青|緑|黒|白|黄)$/i.test(x)) { color = { red: "#ba1a1a", "赤": "#ba1a1a", blue: "#1d4ed8", "青": "#1d4ed8", green: "#006b2c", "緑": "#006b2c", black: "#191b23", "黒": "#191b23", white: "#ffffff", "白": "#ffffff", yellow: "#a16207", "黄": "#a16207", magenta: "#a21caf", cyan: "#0e7490" }[x.toLowerCase()] || null; return ""; }
      var cm = /^(<=|>=|<>|<|>|=)(-?[\d.]+)$/.exec(x);
      if (cm) { cond = { op: cm[1], v: +cm[2] }; return ""; }
      if (/^\$/.test(x)) { var cur = /^\$([^-]*)/.exec(x); return cur ? '"' + cur[1] + '"' : ""; }
      if (/^(h|m|s)+$/i.test(x)) return x; // 経過時間 [h]
      return "";
    });
    // 文字（"…"・\x）とそれ以外に分ける
    var parts = [], i = 0, lit = "";
    function flushLit() { if (lit) { parts.push({ lit: lit }); lit = ""; } }
    while (i < sec.length) {
      var ch = sec.charAt(i);
      if (ch === '"') { var j = sec.indexOf('"', i + 1); if (j < 0) j = sec.length; lit += sec.slice(i + 1, j); i = j + 1; continue; }
      if (ch === "\\") { lit += sec.charAt(i + 1); i += 2; continue; }
      if (ch === "_") { lit += " "; i += 2; continue; }
      if (ch === "*") { i += 2; continue; }
      var m = /^(yyyy|yy|e|g{1,3}|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s|AM\/PM|am\/pm|A\/P)/i.exec(sec.slice(i));
      if (m) { flushLit(); parts.push({ dt: m[0] }); i += m[0].length; continue; }
      m = /^[0#?.,][0#?,.%]*/.exec(sec.slice(i));
      if (m && /[0#?]/.test(m[0])) { flushLit(); parts.push({ num: m[0] }); i += m[0].length; continue; }
      if (ch === "@") { flushLit(); parts.push({ at: true }); i++; continue; }
      if (ch === "%") { flushLit(); parts.push({ pct: true }); i++; continue; }
      lit += ch; i++;
    }
    flushLit();
    var isDate = parts.some(function (p) { return p.dt; });
    // m は、h の後か s の前なら「分」
    if (isDate) {
      parts.forEach(function (p, k) {
        if (!p.dt || !/^m{1,2}$/i.test(p.dt)) return;
        var prev = null, next = null, a;
        for (a = k - 1; a >= 0; a--) if (parts[a].dt) { prev = parts[a].dt; break; }
        for (a = k + 1; a < parts.length; a++) if (parts[a].dt) { next = parts[a].dt; break; }
        if ((prev && /^h/i.test(prev)) || (next && /^s/i.test(next))) p.min = true;
      });
    }
    var pct = parts.some(function (p) { return p.pct || (p.num && /%/.test(p.num)); });
    return { parts: parts, isDate: isDate, pct: pct, color: color, cond: cond, ampm: parts.some(function (p) { return p.dt && /\//.test(p.dt); }) };
  }
  function fmtNum(v, pat) {
    var p = pat.replace(/%/g, ""), dot = p.indexOf("."), ip = dot < 0 ? p : p.slice(0, dot), dp = dot < 0 ? "" : p.slice(dot + 1);
    var comma = /,/.test(ip) && /[0#?],[0#?]/.test(ip);
    var trail = (/(,+)$/.exec(ip) || ["", ""])[1].length; // 末尾の , は 1000 で割る
    if (trail) v = v / Math.pow(1000, trail);
    var dmin = (dp.match(/0/g) || []).length, dmax = (dp.match(/[0#?]/g) || []).length, imin = (ip.match(/0/g) || []).length;
    var s = v.toFixed(dmax), parts = s.split(".");
    var ints = parts[0], decs = parts[1] || "";
    while (decs.length > dmin && decs.slice(-1) === "0") decs = decs.slice(0, -1);
    while (ints.length < imin) ints = "0" + ints;
    if (imin === 0 && ints === "0") ints = "";
    if (comma) ints = ints.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return ints + (decs || (dmax && dmin) ? "." + decs : "");
  }
  function fmtDate(v, c) {
    var d = fromSerial(v), H = d.getUTCHours(), out = "";
    c.parts.forEach(function (p) {
      if (p.lit !== undefined) { out += p.lit; return; }
      if (p.num) { out += p.num; return; }
      if (!p.dt) return;
      var t = p.dt;
      if (p.min) { out += t.length > 1 ? ("0" + d.getUTCMinutes()).slice(-2) : d.getUTCMinutes(); return; }
      switch (t.toLowerCase()) {
        case "yyyy": out += d.getUTCFullYear(); break;
        case "yy": out += ("0" + d.getUTCFullYear() % 100).slice(-2); break;
        case "e": out += d.getUTCFullYear() - 2018; break; // 令和
        case "g": case "gg": case "ggg": out += "令和"; break;
        case "mmmm": out += MON_EN[d.getUTCMonth()]; break;
        case "mmm": out += MON_EN[d.getUTCMonth()]; break;
        case "mm": out += ("0" + (d.getUTCMonth() + 1)).slice(-2); break;
        case "m": out += d.getUTCMonth() + 1; break;
        case "dddd": case "ddd": out += WEEK_EN[d.getUTCDay()]; break;
        case "dd": out += ("0" + d.getUTCDate()).slice(-2); break;
        case "d": out += d.getUTCDate(); break;
        case "aaaa": out += WEEK_JA[d.getUTCDay()] + "曜日"; break;
        case "aaa": out += WEEK_JA[d.getUTCDay()]; break;
        case "hh": out += ("0" + (c.ampm ? (H % 12 || 12) : H)).slice(-2); break;
        case "h": out += c.ampm ? (H % 12 || 12) : H; break;
        case "ss": out += ("0" + d.getUTCSeconds()).slice(-2); break;
        case "s": out += d.getUTCSeconds(); break;
        default: out += /^a/i.test(t) ? (H < 12 ? (t.length === 3 ? "A" : "AM") : (t.length === 3 ? "P" : "PM")) : "";
      }
    });
    return out;
  }
  function format(v, code) {
    if (typeof v !== "number") return { text: str(v) };
    if (!code || /^general$/i.test(code) || code === "G/標準") return { text: general(v) };
    var secs = fmtCache[code] || (fmtCache[code] = splitSections(code).map(compile));
    var c = secs[0], val = v, neg = false;
    if (secs.length > 1 && !secs[0].cond) {
      if (v < 0) { c = secs[1]; val = -v; }
      else if (v === 0 && secs.length > 2) c = secs[2];
    } else if (secs[0].cond) {
      var ok = function (s) { var q = s.cond; return !q || (q.op === ">" ? v > q.v : q.op === "<" ? v < q.v : q.op === ">=" ? v >= q.v : q.op === "<=" ? v <= q.v : q.op === "=" ? v === q.v : v !== q.v); };
      c = secs.filter(ok)[0] || secs[secs.length - 1];
      if (v < 0 && !c.cond) neg = true;
      val = Math.abs(v);
    } else if (v < 0) { neg = true; val = -v; }
    if (c.isDate) return { text: fmtDate(val, c), color: c.color };
    if (c.pct) val *= 100;
    var out = "", used = false;
    c.parts.forEach(function (p) {
      if (p.lit !== undefined) out += p.lit;
      else if (p.num) { out += fmtNum(val, p.num); used = true; if (/%/.test(p.num)) out += "%"; }
      else if (p.pct) out += "%";
      else if (p.at) out += general(v);
    });
    if (!used && !c.parts.some(function (p) { return p.at; }) && !c.parts.length) out = general(val);
    if (neg && /[1-9]/.test(out)) out = "-" + out;
    return { text: out, color: c.color };
  }
  function isDateFormat(code) {
    if (!code || /^general$/i.test(code)) return false;
    var secs = fmtCache[code] || (fmtCache[code] = splitSections(code).map(compile));
    return secs[0].isDate;
  }

  window.OfficeFormula = {
    ERR: ERR, err: err, isErr: isErr, isRange: isRange,
    parse: parse, valid: valid, evaluate: evaluate, tokenize: tokenize,
    shift: shift, adjust: adjust, renameSheet: renameSheet, refs: refs, sheetText: sheetText,
    colName: colName, colIndex: colIndex,
    num: num, str: str, parseNum: parseNum, parseDate: parseDate, compare: compare,
    format: format, isDateFormat: isDateFormat, serial: serial, fromSerial: fromSerial, today: today,
    names: Object.keys(FN).sort()
  };
})();
