// SK Hub Office のグラフ（window.OfficeChart）。Sheet・Slide・Doc で同じものを使う
//   OfficeChart.svg(spec, { w, h, dark }) → SVG の文字列
//     spec: { type: "col|bar|line|area|pie|scatter", title, cats: [項目], series: [{ name, values: [数] }], legend: true|false }
//   OfficeChart.fromRows(2 次元の値) → { cats, series }（1 列目が文字なら項目、1 行目が文字なら系列の名前）
//   色の順番はいつも同じ（系列の 1 番目は青、2 番目は橙…）。色だけで見分けないよう、系列が 2 つ以上なら凡例を出す
//   線は細く・棒の先は 4px の角丸・棒と棒のあいだは 2px あける・目盛りの線は薄く
(function () {
  "use strict";
  // 色覚の多様性に配慮して確かめた 8 色（この順番で使う。9 番目からは「その他」にまとめる）
  var LIGHT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"];
  var DARK = ["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181", "#008300", "#9085e9", "#e66767"];
  var INK = { light: { text: "#0b0b0b", sub: "#52514e", muted: "#898781", grid: "#e1e0d9", axis: "#c3c2b7", surface: "#ffffff" },
              dark: { text: "#f0efec", sub: "#c3c2b7", muted: "#898781", grid: "#2c2c2a", axis: "#383835", surface: "#1a1a19" } };
  var FONT = '"M PLUS 1", "Hiragino Sans", "Noto Sans JP", "Yu Gothic UI", system-ui, sans-serif';
  function esc(s) { return String(s === undefined || s === null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function fmt(v) {
    if (typeof v !== "number" || !isFinite(v)) return "";
    var a = Math.abs(v);
    if (a >= 1e8) return +(v / 1e8).toFixed(1) + "億";
    if (a >= 1e4 && Math.round(v) === v && a % 1000 === 0) return +(v / 1e4).toFixed(1) + "万";
    return (Math.round(v * 100) / 100).toLocaleString("ja-JP");
  }
  // きりのよい目盛り
  function ticks(min, max, n) {
    if (min === max) { max = min + 1; if (min > 0) min = 0; }
    var span = max - min, step = Math.pow(10, Math.floor(Math.log10(span / n))), err = span / n / step;
    if (err >= 7.5) step *= 10; else if (err >= 3.5) step *= 5; else if (err >= 1.5) step *= 2;
    var lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step, out = [];
    for (var v = lo; v <= hi + step / 2; v += step) out.push(Math.round(v / step) * step);
    return out;
  }
  function textW(s, size) { var w = 0; Array.from(String(s)).forEach(function (ch) { w += /[\u0000-ÿ]/.test(ch) ? size * 0.58 : size; }); return w; }
  function clip(s, size, max) { s = String(s); if (textW(s, size) <= max) return s; var a = Array.from(s); while (a.length > 1 && textW(a.join("") + "…", size) > max) a.pop(); return a.join("") + "…"; }

  function fromRows(rows) {
    rows = (rows || []).filter(function (r) { return r.some(function (v) { return v !== null && v !== ""; }); });
    if (!rows.length) return { cats: [], series: [] };
    var isNum = function (v) { return typeof v === "number"; };
    var headRow = rows[0].slice(1).some(function (v) { return typeof v === "string" && v !== ""; }) || (rows.length > 1 && !rows[0].some(isNum));
    var catCol = rows[0].length > 1 && (rows.slice(headRow ? 1 : 0).some(function (r) { return typeof r[0] === "string"; }) || (headRow && (rows[0][0] === null || rows[0][0] === "")));
    var body = rows.slice(headRow ? 1 : 0), start = catCol ? 1 : 0, w = rows[0].length;
    var cats = body.map(function (r, i) { return catCol ? (r[0] === null ? "" : String(r[0])) : String(i + 1); });
    var series = [];
    for (var c = start; c < w; c++) {
      series.push({ name: headRow ? String(rows[0][c] === null ? "" : rows[0][c]) : "系列 " + (c - start + 1), values: body.map(function (r) { return isNum(r[c]) ? r[c] : null; }) });
    }
    series = series.filter(function (s) { return s.values.some(function (v) { return v !== null; }); });
    return { cats: cats, series: series };
  }

  function svg(spec, o) {
    o = o || {};
    var W = o.w || 640, H = o.h || 400, ink = o.dark ? INK.dark : INK.light, pal = o.dark ? DARK : LIGHT;
    var type = spec.type || "col", series = (spec.series || []).slice(0, 8), cats = spec.cats || [];
    var fs = Math.max(10, Math.min(14, W / 46)), out = [];
    var top = 12, title = spec.title ? String(spec.title) : "";
    out.push('<rect width="' + W + '" height="' + H + '" fill="' + (o.transparent ? "none" : ink.surface) + '"/>');
    if (title) { out.push('<text x="' + W / 2 + '" y="' + (top + fs * 1.2) + '" text-anchor="middle" font-size="' + fs * 1.25 + '" font-weight="700" fill="' + ink.text + '">' + esc(clip(title, fs * 1.25, W - 24)) + "</text>"); top += fs * 2; }
    var showLegend = spec.legend !== false && (type === "pie" ? cats.length > 0 : series.length > 1);
    var legendItems = type === "pie" ? pieSlices(series[0], cats).map(function (s) { return s.name; }) : series.map(function (s) { return s.name; });
    var legendH = 0;
    if (showLegend) {
      var lx = 0, rowsN = 1;
      legendItems.forEach(function (n) { var w = 18 + textW(n, fs) + 16; if (lx + w > W - 24 && lx > 0) { rowsN++; lx = 0; } lx += w; });
      legendH = rowsN * fs * 1.6 + 6;
    }
    var box = { x: 12, y: top + 4, w: W - 24, h: H - top - 12 - legendH };
    if (!series.length || !cats.length) {
      out.push('<text x="' + W / 2 + '" y="' + H / 2 + '" text-anchor="middle" font-size="' + fs + '" fill="' + ink.muted + '">データがありません</text>');
      return wrap(out, W, H);
    }
    if (type === "pie") drawPie(out, series[0], cats, box, pal, ink, fs);
    else drawXY(out, type, series, cats, box, pal, ink, fs);
    if (showLegend) {
      var x = 12, y = H - legendH + fs * 0.6, lines = [[]], widths = [0];
      legendItems.forEach(function (n, i) {
        var w = 18 + textW(clip(n, fs, W / 2), fs) + 16;
        if (widths[widths.length - 1] + w > W - 24 && lines[lines.length - 1].length) { lines.push([]); widths.push(0); }
        lines[lines.length - 1].push({ n: clip(n, fs, W / 2), i: i, w: w });
        widths[widths.length - 1] += w;
      });
      lines.forEach(function (ln, li) {
        x = (W - widths[li]) / 2 + 8;
        ln.forEach(function (it) {
          var cy = y + li * fs * 1.6;
          out.push('<rect x="' + x + '" y="' + (cy - fs * 0.35) + '" width="' + fs * 0.8 + '" height="' + fs * 0.8 + '" rx="2" fill="' + (it.i >= 8 ? ink.muted : pal[it.i]) + '"/>');
          out.push('<text x="' + (x + fs * 0.8 + 6) + '" y="' + (cy + fs * 0.35) + '" font-size="' + fs + '" fill="' + ink.sub + '">' + esc(it.n) + "</text>");
          x += it.w;
        });
      });
    }
    return wrap(out, W, H);
  }
  function wrap(out, W, H) {
    return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + W + " " + H + '" width="' + W + '" height="' + H + '" font-family=\'' + FONT + "'>" + out.join("") + "</svg>";
  }
  function pieSlices(s, cats) {
    var list = cats.map(function (c, i) { return { name: c, v: s && typeof s.values[i] === "number" && s.values[i] > 0 ? s.values[i] : 0 }; }).filter(function (x) { return x.v > 0; });
    if (list.length > 8) {
      var rest = list.slice(7).reduce(function (a, x) { return a + x.v; }, 0);
      list = list.slice(0, 7).concat([{ name: "その他", v: rest, other: true }]);
    }
    return list;
  }
  function drawPie(out, s, cats, box, pal, ink, fs) {
    var list = pieSlices(s, cats), total = list.reduce(function (a, x) { return a + x.v; }, 0);
    if (!total) return;
    var r = Math.min(box.w, box.h) / 2 - fs, cx = box.x + box.w / 2, cy = box.y + box.h / 2, a0 = -Math.PI / 2;
    list.forEach(function (x, i) {
      var a1 = a0 + x.v / total * Math.PI * 2, large = a1 - a0 > Math.PI ? 1 : 0;
      var p = list.length === 1 ? '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '"' :
        '<path d="M' + cx + "," + cy + "L" + (cx + r * Math.cos(a0)) + "," + (cy + r * Math.sin(a0)) + "A" + r + "," + r + " 0 " + large + " 1 " + (cx + r * Math.cos(a1)) + "," + (cy + r * Math.sin(a1)) + 'Z"';
      out.push(p + ' fill="' + (x.other ? ink.muted : pal[i]) + '" stroke="' + ink.surface + '" stroke-width="2"><title>' + esc(x.name + ": " + fmt(x.v) + "（" + Math.round(x.v / total * 100) + "%）") + "</title>" + (list.length === 1 ? "</circle>" : "</path>"));
      // 大きいところだけ、割合を書く
      if (x.v / total >= 0.06) {
        var am = (a0 + a1) / 2, tx = cx + r * 0.64 * Math.cos(am), ty = cy + r * 0.64 * Math.sin(am);
        out.push('<text x="' + tx + '" y="' + (ty + fs * 0.35) + '" text-anchor="middle" font-size="' + fs + '" font-weight="700" fill="#ffffff" pointer-events="none">' + Math.round(x.v / total * 100) + "%</text>");
      }
      a0 = a1;
    });
  }
  function drawXY(out, type, series, cats, box, pal, ink, fs) {
    var all = [];
    series.forEach(function (s) { s.values.forEach(function (v) { if (typeof v === "number") all.push(v); }); });
    if (!all.length) return;
    var min = Math.min(0, Math.min.apply(null, all)), max = Math.max(0, Math.max.apply(null, all));
    if (type === "line" || type === "scatter") { min = Math.min.apply(null, all); max = Math.max.apply(null, all); if (min > 0 && min < max * 0.5) min = 0; }
    var tk = ticks(min, max, Math.max(3, Math.min(6, Math.floor((type === "bar" ? box.w : box.h) / 60))));
    min = tk[0]; max = tk[tk.length - 1];
    var labW = Math.max.apply(null, tk.map(function (v) { return textW(fmt(v), fs * 0.9); })) + 10;
    var horiz = type === "bar";
    var catW = horiz ? Math.min(box.w * 0.3, Math.max.apply(null, cats.map(function (c) { return textW(c, fs * 0.9); })) + 10) : 0;
    var px = box.x + (horiz ? catW : labW), py = box.y + 4, pw = box.w - (horiz ? catW : labW) - 8, ph = box.h - (horiz ? fs * 1.6 : fs * 2) - 4;
    var n = cats.length, step = (horiz ? ph : pw) / n;
    function val(v) { return horiz ? px + (v - min) / (max - min) * pw : py + ph - (v - min) / (max - min) * ph; }
    // 目盛りの線（薄く）と数
    tk.forEach(function (v) {
      if (horiz) {
        var x = val(v);
        out.push('<line x1="' + x + '" y1="' + py + '" x2="' + x + '" y2="' + (py + ph) + '" stroke="' + (v === 0 ? ink.axis : ink.grid) + '" stroke-width="1"/>');
        out.push('<text x="' + x + '" y="' + (py + ph + fs * 1.3) + '" text-anchor="middle" font-size="' + fs * 0.9 + '" fill="' + ink.muted + '">' + esc(fmt(v)) + "</text>");
      } else {
        var y = val(v);
        out.push('<line x1="' + px + '" y1="' + y + '" x2="' + (px + pw) + '" y2="' + y + '" stroke="' + (v === 0 ? ink.axis : ink.grid) + '" stroke-width="1"/>');
        out.push('<text x="' + (px - 6) + '" y="' + (y + fs * 0.32) + '" text-anchor="end" font-size="' + fs * 0.9 + '" fill="' + ink.muted + '">' + esc(fmt(v)) + "</text>");
      }
    });
    // 項目の名前（多いときは間引く）
    var every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? ph : pw) / (horiz ? fs * 1.5 : Math.max(40, fs * 4))))));
    cats.forEach(function (c, i) {
      if (i % every) return;
      if (horiz) out.push('<text x="' + (px - 6) + '" y="' + (py + step * (i + 0.5) + fs * 0.32) + '" text-anchor="end" font-size="' + fs * 0.9 + '" fill="' + ink.sub + '">' + esc(clip(c, fs * 0.9, catW - 8)) + "</text>");
      else out.push('<text x="' + (px + step * (i + 0.5)) + '" y="' + (py + ph + fs * 1.4) + '" text-anchor="middle" font-size="' + fs * 0.9 + '" fill="' + ink.sub + '">' + esc(clip(c, fs * 0.9, step * every - 4)) + "</text>");
    });
    var zero = val(Math.max(min, Math.min(max, 0)));
    if (type === "col" || type === "bar") {
      var k = series.length, band = step * 0.72, bw = Math.max(1, (band - (k - 1) * 2) / k);
      series.forEach(function (s, si) {
        s.values.forEach(function (v, i) {
          if (typeof v !== "number") return;
          var start = step * i + (step - band) / 2 + si * (bw + 2), tip = val(v), r = Math.min(4, bw / 2);
          var d = horiz ? barPath(zero, py + start, tip, bw, r, true) : barPath(zero, px + start, tip, bw, r, false);
          out.push('<path d="' + d + '" fill="' + pal[si] + '"><title>' + esc((k > 1 ? s.name + " ・ " : "") + cats[i] + ": " + fmt(v)) + "</title></path>");
        });
      });
    } else {
      series.forEach(function (s, si) {
        var pts = [];
        s.values.forEach(function (v, i) { if (typeof v === "number") pts.push([px + step * (i + 0.5), val(v), v, i]); });
        if (!pts.length) return;
        if (type === "area") {
          out.push('<path d="M' + pts[0][0] + "," + zero + pts.map(function (p) { return "L" + p[0] + "," + p[1]; }).join("") + "L" + pts[pts.length - 1][0] + "," + zero + 'Z" fill="' + pal[si] + '" fill-opacity="' + (series.length > 1 ? 0.22 : 0.3) + '"/>');
        }
        if (type !== "scatter") out.push('<polyline points="' + pts.map(function (p) { return p[0] + "," + p[1]; }).join(" ") + '" fill="none" stroke="' + pal[si] + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>');
        var dot = type === "scatter" ? 5 : pts.length <= 24 ? 4 : 0;
        pts.forEach(function (p) {
          out.push('<circle cx="' + p[0] + '" cy="' + p[1] + '" r="' + Math.max(dot, 0.01) + '" fill="' + pal[si] + '" stroke="' + ink.surface + '" stroke-width="' + (dot ? 2 : 0) + '"><title>' + esc((series.length > 1 ? s.name + " ・ " : "") + cats[p[3]] + ": " + fmt(p[2])) + "</title></circle>");
        });
      });
    }
  }
  // 棒（先だけ角を丸くする）
  function barPath(base, pos, tip, w, r, horiz) {
    var len = tip - base, sgn = len < 0 ? -1 : 1, a = Math.abs(len);
    r = Math.min(r, a);
    if (horiz) {
      var y0 = pos, y1 = pos + w, x0 = base, x1 = tip;
      return "M" + x0 + "," + y0 + "L" + (x1 - sgn * r) + "," + y0 + "Q" + x1 + "," + y0 + " " + x1 + "," + (y0 + r) + "L" + x1 + "," + (y1 - r) + "Q" + x1 + "," + y1 + " " + (x1 - sgn * r) + "," + y1 + "L" + x0 + "," + y1 + "Z";
    }
    var X0 = pos, X1 = pos + w, Y0 = base, Y1 = tip;
    return "M" + X0 + "," + Y0 + "L" + X0 + "," + (Y1 - sgn * r) + "Q" + X0 + "," + Y1 + " " + (X0 + r) + "," + Y1 + "L" + (X1 - r) + "," + Y1 + "Q" + X1 + "," + Y1 + " " + X1 + "," + (Y1 - sgn * r) + "L" + X1 + "," + Y0 + "Z";
  }
  var TYPES = [
    { value: "col", label: "縦棒", icon: "bar_chart" }, { value: "bar", label: "横棒", icon: "align_horizontal_left" },
    { value: "line", label: "折れ線", icon: "show_chart" }, { value: "area", label: "面", icon: "area_chart" },
    { value: "pie", label: "円", icon: "pie_chart" }, { value: "scatter", label: "散布図", icon: "scatter_plot" }
  ];
  window.OfficeChart = { svg: svg, fromRows: fromRows, TYPES: TYPES, colors: LIGHT };
})();
