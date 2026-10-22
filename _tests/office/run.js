// SK Hub Office のテスト（ブラウザの中で動かす）
//   テンプレートと fixtures.js のファイルを 1 つずつ、アプリ（doc.js・sheet.js・slide.js）に開かせて、
//   ① 書き出す（.docx・.xlsx・.pptx・.csv など）→ ② Office のファイルのしくみ（ZIP・XML・rels・[Content_Types].xml）を確かめる →
//   ③ 書き出したファイルを読みこみなおして、中身（文字・セルの値・スライドの枚数・ノート）が残っているかを確かめる
//   office.js の代わりに、小さな道具（O）をここで作って渡す。保存の画面は出さず、書き出したファイルをここで受け取る
//   テストだけの第三者エンジン（docx・docx-preview）でも、書き出した .docx を開けるか確かめる
//   ?post=1 のときは、結果と書き出したファイルを run.py（テスト用のサーバー）に送る（Python でも確かめるため）
(function () {
  "use strict";
  var Zip = window.OfficeZip, results = [], files = [];
  var POST = /[?&]post=1/.test(location.search);
  var logEl = document.getElementById("log");

  function log(ok, app, fx, step, msg) {
    results.push({ ok: ok, app: app, fixture: fx, step: step, msg: msg || "" });
    var line = document.createElement("div");
    line.className = ok ? "ok" : "ng";
    line.textContent = (ok ? "OK  " : "NG  ") + app + " / " + fx + " / " + step + (msg ? " — " + msg : "");
    logEl.appendChild(line);
  }
  function copy(o) { return JSON.parse(JSON.stringify(o)); }
  function esc(s) { return String(s === undefined || s === null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function t(s, p) { return s.replace(/\{(\w+)\}/g, function (m, k) { return p && k in p ? p[k] : m; }); }

  // office.js の代わりの道具（アプリが使うものだけ）
  function makeO(app, file) {
    var O = {
      APP: app, t: t, esc: esc, copy: copy, icon: function () { return ""; }, $: function (id) { return document.getElementById(id); },
      notify: function () {}, when: function (ts) { return new Date(ts).toISOString(); },
      file: function () { return file; }, data: function () { return file.data; },
      changed: function () {}, record: function () {}, stat: function () {}, sync: function () {}, syncSoon: function () {},
      menu: function () { return Promise.resolve(null); }, palette: function () { return Promise.resolve(null); }, closeMenu: function () {},
      pickImage: function () { return Promise.resolve(null); }, pickFile: function () { return Promise.resolve(null); }, imageData: function () { return Promise.resolve(null); },
      saveBlob: function (name, blob) { O.saved = { name: name, blob: blob }; return Promise.resolve(true); },
      loadScript: function (src) {
        var hit = Array.prototype.filter.call(document.scripts, function (s) { return s.src === new URL(src, location.href).href; })[0];
        if (hit) return Promise.resolve(hit);
        return new Promise(function (resolve, reject) {
          var s = document.createElement("script"); s.src = src; s.async = true;
          s.onload = function () { resolve(s); }; s.onerror = function () { reject(new Error("load-failed: " + src)); };
          document.head.appendChild(s);
        });
      },
      loadVendor: function (name) { return O.loadScript(O.base + "shared/vendor/" + name); },
      zoom: function () { return 1; }, readMode: function () { return false; }, userName: function () { return "テストの人"; },
      base: "../../office/", fileName: function (ext) { return file.name + ext; },
      groups: { clip: { label: "", items: [] }, zoom: { label: "", items: [] }, view: { label: "", items: [] } },
      fonts: [["游ゴシック", "sans-serif"]], fontCss: function () { return "sans-serif"; }, sizes: [10.5, 11, 12], dark: function () { return false; }
    };
    return O;
  }
  function open(app, name, data) {
    var file = { id: "test", app: app, name: name, data: copy(data) };
    var O = makeO(app, file), App = window.OfficeApps[app](O);
    document.getElementById("canvas").innerHTML = "";
    App.render();
    return { O: O, App: App, file: file };
  }
  function exportBlob(ctx, id) {
    ctx.O.saved = null;
    return Promise.resolve(ctx.App.exportAs(id)).then(function () {
      if (!ctx.O.saved) throw new Error("書き出したファイルを受け取れませんでした");
      return ctx.O.saved;
    });
  }

  // ---------- Office のファイルのしくみを確かめる ----------
  function partDir(p) { return p.indexOf("/") < 0 ? "" : p.slice(0, p.lastIndexOf("/") + 1); }
  function resolve(base, target) {
    if (target.charAt(0) === "/") return target.slice(1);
    var parts = (base + target).split("/"), out = [];
    parts.forEach(function (x) { if (x === "..") out.pop(); else if (x !== ".") out.push(x); });
    return out.join("/");
  }
  function firstName(u8) { var dv = new DataView(u8.buffer, u8.byteOffset); return new TextDecoder().decode(u8.subarray(30, 30 + dv.getUint16(26, true))); }
  function checkPackage(blob) {
    var problems = [];
    return blob.arrayBuffer().then(function (ab) {
      if (firstName(new Uint8Array(ab)) !== "[Content_Types].xml") problems.push("ZIP のいちばん先が [Content_Types].xml ではありません");
      return Zip.read(ab);
    }).then(function (z) {
      var names = Object.keys(z), docs = {};
      names.forEach(function (n) {
        if (!/\.(xml|rels)$/.test(n)) return;
        var d = Zip.xml(z[n]);
        if (!d || d.getElementsByTagName("parsererror").length) { problems.push(n + " の XML がこわれています: " + (d ? d.getElementsByTagName("parsererror")[0].textContent.slice(0, 160) : "")); return; }
        docs[n] = d;
      });
      var ct = docs["[Content_Types].xml"];
      if (!ct) { problems.push("[Content_Types].xml がありません"); return problems; }
      var defaults = {}, overrides = {};
      Array.prototype.forEach.call(ct.getElementsByTagNameNS("*", "Default"), function (d) { defaults[d.getAttribute("Extension").toLowerCase()] = d.getAttribute("ContentType"); });
      Array.prototype.forEach.call(ct.getElementsByTagNameNS("*", "Override"), function (d) { overrides[d.getAttribute("PartName").replace(/^\//, "")] = d.getAttribute("ContentType"); });
      Object.keys(overrides).forEach(function (p) { if (!(p in z)) problems.push("[Content_Types].xml にあるのに、ファイルがありません: " + p); });
      names.forEach(function (n) {
        if (n === "[Content_Types].xml") return;
        if (!(n in overrides) && !(n.split(".").pop().toLowerCase() in defaults)) problems.push("種類（Content Type）が決まっていません: " + n);
      });
      if (!z["_rels/.rels"]) problems.push("_rels/.rels がありません");
      names.filter(function (n) { return /\.rels$/.test(n); }).forEach(function (n) {
        var owner = n.replace(/_rels\/([^/]+)\.rels$/, "$1"), base = partDir(owner), ids = {};
        if (n !== "_rels/.rels" && !(owner in z)) problems.push("つながり（rels）の持ち主のファイルがありません: " + n);
        Array.prototype.forEach.call(docs[n] ? docs[n].getElementsByTagNameNS("*", "Relationship") : [], function (r) {
          var id = r.getAttribute("Id");
          if (ids[id]) problems.push(n + " で同じ Id が 2 回使われています: " + id);
          ids[id] = 1;
          if (r.getAttribute("TargetMode") === "External") return;
          var target = resolve(n === "_rels/.rels" ? "" : base, r.getAttribute("Target"));
          if (!(target in z)) problems.push(n + " の " + id + " の先がありません: " + target);
        });
        // 持ち主の XML の中で使っている r:id が、rels にあるか
        var od = docs[owner];
        if (od) {
          var all = od.getElementsByTagName("*");
          for (var i = 0; i < all.length; i++) {
            for (var a = 0; a < all[i].attributes.length; a++) {
              var at = all[i].attributes[a];
              if (at.namespaceURI === "http://schemas.openxmlformats.org/officeDocument/2006/relationships" && !ids[at.value]) problems.push(owner + " の r:" + at.localName + "=\"" + at.value + "\" が rels にありません");
            }
          }
        }
      });
      return problems;
    });
  }

  // ---------- 第三者エンジン（docx-preview）で、書き出した .docx を開けるか ----------
  // 本物のアプリには入れない（テストのときだけ読み込む）。画面の外に置いて、本文の先頭の文字だけ見る
  function docxOracle(blob, expect) {
    if (!window.docx || !window.docx.renderAsync || !window.JSZip) return Promise.resolve();
    var box = document.createElement("div");
    box.style.cssText = "position:fixed;left:-10000px;top:0;width:800px;height:600px;overflow:hidden";
    document.body.appendChild(box);
    return window.docx.renderAsync(blob, box, box, { inWrapper: false, ignoreWidth: true, ignoreHeight: true }).then(function () {
      // docx-preview は、同じ入れ物に <style> も入れる。文字だけを見るため、先に取る
      Array.prototype.forEach.call(box.querySelectorAll("style"), function (s) { s.remove(); });
      var text = box.textContent || "", sample = ((expect && expect.text) || [])[0], bad = [];
      if (sample && text.indexOf(sample) < 0) bad.push("本文に「" + sample + "」がありません");
      box.remove();
      log(!bad.length, "doc", "oracle", "docx-preview", bad.join(" / ") || (sample ? "「" + sample + "」を確認" : "本文なし"));
    }).catch(function (e) {
      box.remove();
      log(false, "doc", "oracle", "docx-preview", (e && e.message) || String(e));
    });
  }

  // ---------- 読みこみなおして、中身を確かめる ----------
  function textOfHtml(html) { var d = document.createElement("div"); d.innerHTML = html; return d.textContent; }
  // CSS の宣言（並べる順番は関係ない）を足して並べる
  function normCss(css) { return String(css || "").split(";").map(function (x) { return x.trim().replace(/\s+/g, " "); }).filter(Boolean).sort().join(";"); }
  function parseCsv(text) {
    var rows = [[]], cur = "", q = false;
    text = text.replace(/^﻿/, "");
    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (q) { if (ch === '"') { if (text.charAt(i + 1) === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
      if (ch === '"') q = true;
      else if (ch === ",") { rows[rows.length - 1].push(cur); cur = ""; }
      else if (ch === "\r") { /* 次の \n で区切る */ }
      else if (ch === "\n") { rows[rows.length - 1].push(cur); cur = ""; rows.push([]); }
      else cur += ch;
    }
    return rows;
  }
  function colIndex(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n - 1; }
  // Sheet: シートごとに CSV に書き出して、見える文字を読む（{ "シート名!A1": "文字" }）
  function sheetValues(data, wanted) {
    var out = {}, chain = Promise.resolve();
    data.sheets.forEach(function (s, i) {
      chain = chain.then(function () {
        var d = copy(data); d.cur = i;
        var ctx = open("sheet", "csv", d);
        return exportBlob(ctx, "csv").then(function (r) { return r.blob.text(); }).then(function (text) {
          var rows = parseCsv(text);
          Object.keys(wanted).forEach(function (k) {
            var m = /^(.*)!([A-Z]+)(\d+)$/.exec(k);
            if (m[1] !== s.name) return;
            var row = rows[+m[3] - 1] || [];
            out[k] = row[colIndex(m[2])] === undefined ? "" : row[colIndex(m[2])];
          });
        });
      });
    });
    return chain.then(function () { return out; });
  }
  function checkContent(app, data, expect) {
    var bad = [];
    if (!expect) return Promise.resolve(bad);
    if (app === "doc") {
      var text = textOfHtml(data.html);
      (expect.text || []).forEach(function (x) { if (text.indexOf(x) < 0) bad.push("文字がありません: " + x); });
      var holder = document.createElement("div"); holder.innerHTML = data.html;
      (expect.footnotes || []).forEach(function (x) {
        var li = Array.prototype.filter.call(holder.querySelectorAll("section.footnotes li[data-fn]"), function (l) { return l.textContent.indexOf(x) >= 0; })[0];
        if (!li) bad.push("脚注がありません: " + x);
        else if (!holder.querySelector('sup.fn-ref[data-fn="' + li.getAttribute("data-fn") + '"]')) bad.push("脚注の印が本文にありません: " + x);
      });
      (expect.comments || []).forEach(function (x) {
        var c = (data.comments || []).filter(function (c) { return c.text === x[1]; })[0];
        if (!c) { bad.push("コメントがありません: " + x[1]); return; }
        var span = holder.querySelector('span.cmt[data-cmt="' + c.id + '"]');
        if (!span || span.textContent !== x[0]) bad.push("コメントの範囲が「" + x[0] + "」ではありません: " + (span ? span.textContent : "なし"));
        if (c.author !== x[2]) bad.push("コメントを書いた人が " + c.author);
      });
      (expect.style || []).forEach(function (x) {
        var e = holder.querySelector(x.sel);
        if (!e) { bad.push("要素がありません: " + x.sel); return; }
        Object.keys(x.css || {}).forEach(function (k) {
          var got = e.style.getPropertyValue(k);
          if (got !== x.css[k]) bad.push(x.sel + " の " + k + " が " + JSON.stringify(got) + "（" + JSON.stringify(x.css[k]) + " のはず）");
        });
      });
      (expect.attr || []).forEach(function (x) {
        var e = holder.querySelector(x.sel);
        if (!e) { bad.push("要素がありません: " + x.sel); return; }
        var got = e.getAttribute(x.name);
        if (got !== x.value) bad.push(x.sel + " の " + x.name + " が " + JSON.stringify(got) + "（" + JSON.stringify(x.value) + " のはず）");
      });
      (expect.styles || []).forEach(function (x) {
        var hit = (data.styles || []).filter(function (s) { return s.id === x.id; })[0];
        if (!hit) { bad.push("カスタムスタイルがありません: " + x.id); return; }
        if (x.name !== undefined && hit.name !== x.name) bad.push("スタイル " + x.id + " の名前: " + hit.name + "（" + x.name + " のはず）");
        if (x.base !== undefined && hit.base !== x.base) bad.push("スタイル " + x.id + " の元のスタイル: " + hit.base + "（" + x.base + " のはず）");
        if (x.css !== undefined && normCss(hit.css) !== normCss(x.css)) bad.push("スタイル " + x.id + " の書式: " + hit.css + "（" + x.css + " のはず）");
      });
      return Promise.resolve(bad);
    }
    if (app === "slide") {
      if (expect.slides !== undefined && data.slides.length !== expect.slides) bad.push("スライドの枚数: " + data.slides.length + "（" + expect.slides + " のはず）");
      var flat = [];
      (function walk(list) {
        (list || []).forEach(function (o) { flat.push(o); if (o.type === "group") walk(o.children); });
      })(data.slides.reduce(function (a, s) { return a.concat(s.objs || []); }, []));
      var all = textOfHtml(data.slides.map(function (s) {
        return [s.title, s.sub, (s.items || []).join(" "), (s.items2 || []).join(" ")].map(esc).join(" ");
      }).join(" ") + " " + flat.map(function (o) { return o.html || ""; }).join(" "));
      (expect.text || []).forEach(function (x) { if (all.indexOf(x) < 0) bad.push("文字がありません: " + x); });
      if (expect.groups !== undefined) {
        var groups = flat.filter(function (o) { return o.type === "group"; });
        if (groups.length !== expect.groups) bad.push("グループの数: " + groups.length + "（" + expect.groups + " のはず）");
        if (expect.groupChildren !== undefined) {
          var kids = groups.reduce(function (a, g) { return a + (g.children || []).length; }, 0);
          if (kids !== expect.groupChildren) bad.push("グループの中のもの: " + kids + "（" + expect.groupChildren + " のはず）");
        }
      }
      (expect.lw || []).forEach(function (w) {
        if (!flat.some(function (o) { return Math.abs(+o.lw - w) < 0.01; })) bad.push("線の太さがありません: " + w + " pt");
      });
      var notes = data.slides.map(function (s) { return s.notes || ""; }).join("\n");
      (expect.notes || []).forEach(function (x) { if (notes.indexOf(x) < 0) bad.push("ノートがありません: " + x); });
      return Promise.resolve(bad);
    }
    if (expect.sheets !== undefined && data.sheets.length !== expect.sheets) bad.push("シートの数: " + data.sheets.length + "（" + expect.sheets + " のはず）");
    Object.keys(expect.cf || {}).forEach(function (name) {
      var s = data.sheets.filter(function (x) { return x.name === name; })[0], got = s ? (s.cf || []).map(function (r) { return r.type; }).sort().join(",") : "";
      if (got !== expect.cf[name]) bad.push(name + " の条件付き書式: " + got + "（" + expect.cf[name] + " のはず）");
    });
    function dvKey(r) { return [r.ref, r.type, r.op, r.f1, r.f2, r.style, r.promptTitle, r.prompt, r.errTitle, r.err].join("|"); }
    Object.keys(expect.dv || {}).forEach(function (name) {
      var s = data.sheets.filter(function (x) { return x.name === name; })[0];
      var got = s ? (s.dv || []).map(dvKey).sort().join(",") : "", want = expect.dv[name].map(dvKey).sort().join(",");
      if (got !== want) bad.push(name + " の入力規則が " + (got || "なし") + "（" + want + " のはず）");
    });
    Object.keys(expect.notes || {}).forEach(function (k) {
      var m = /^(.*)!([A-Z]+[0-9]+)$/.exec(k);
      if (!m) return;
      var s = data.sheets.filter(function (x) { return x.name === m[1]; })[0], nt = s && s.notes && s.notes[m[2]], want = expect.notes[k];
      if (!nt) { bad.push(k + " のコメントがありません"); return; }
      if (nt.text !== want.text) bad.push(k + " のコメントが " + JSON.stringify(nt.text) + "（" + JSON.stringify(want.text) + " のはず）");
      if (want.author !== undefined && nt.author !== want.author) bad.push(k + " のコメントの書き手が " + JSON.stringify(nt.author) + "（" + JSON.stringify(want.author) + " のはず）");
    });
    return sheetValues(data, expect.cells || {}).then(function (vals) {
      Object.keys(expect.cells || {}).forEach(function (k) { if (vals[k] !== expect.cells[k]) bad.push(k + " が " + JSON.stringify(vals[k]) + "（" + JSON.stringify(expect.cells[k]) + " のはず）"); });
      return bad;
    });
  }

  var MAIN = { doc: "docx", sheet: "xlsx", slide: "pptx" };
  function outExt(id) {
    if (id === "outline") return "txt";
    if (id === "xlsx-charts") return "xlsx";
    if (id === "pptx-native" || id === "pptx-preview") return "pptx";
    return id;
  }
  // PptxGenJS の .pptx は、グループを作れない・アプリの部品とは形がちがうので、
  // 文字とノートと、画面切り替え・本物のグラフが入っているかだけを見る
  function nativeCheck(blob, expect) {
    var problems = [];
    expect = expect || {};
    return blob.arrayBuffer().then(Zip.read).then(function (z) {
      var slides = Object.keys(z).filter(function (n) { return /^ppt\/slides\/slide\d+\.xml$/.test(n); });
      var allText = "", allNotes = "", charts = 0, trans = 0;
      slides.forEach(function (n) {
        var doc = Zip.xml(z[n]), xml = Zip.text(z[n]);
        if (!doc) { problems.push(n + " を読めません"); return; }
        Array.prototype.forEach.call(doc.getElementsByTagNameNS("*", "t"), function (t) { allText += t.textContent + "\n"; });
        if (/<c:chart\b/.test(xml)) charts++;
        if (/<p:transition\b/.test(xml)) trans++;
      });
      Object.keys(z).filter(function (n) { return /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n); }).forEach(function (n) {
        var doc = Zip.xml(z[n]);
        if (doc) Array.prototype.forEach.call(doc.getElementsByTagNameNS("*", "t"), function (t) { allNotes += t.textContent + "\n"; });
      });
      if (expect.slides !== undefined && slides.length !== expect.slides) problems.push("スライドの数: " + slides.length + "（" + expect.slides + " のはず）");
      (expect.text || []).forEach(function (x) { if (allText.indexOf(x) < 0) problems.push("文字がありません: " + x); });
      (expect.notes || []).forEach(function (x) { if (allNotes.indexOf(x) < 0) problems.push("ノートがありません: " + x); });
      if (expect.chart && charts < 1) problems.push("本物のグラフ（c:chart）がありません");
      if (expect.transitions && trans < expect.transitions) problems.push("画面切り替えが " + trans + " 枚（" + expect.transitions + " 枚のはず）");
      return problems;
    });
  }
  function post(path, body) {
    if (!POST) return Promise.resolve();
    return fetch(path, { method: "POST", body: body }).catch(function () {});
  }
  function runOne(app, fx) {
    var label = fx.id, ctx;
    try { ctx = open(app, fx.name, fx.data); } catch (e) { log(false, app, label, "開く", e.message); return Promise.resolve(); }
    log(true, app, label, "開く");
    var chain = Promise.resolve();
    // 書き出す前に、期待している中身が、もとのファイルで合っているか（期待のほうのまちがいを見つける）
    chain = chain.then(function () { return checkContent(app, ctx.file.data, fx.expect); }).then(function (bad) { log(!bad.length, app, label, "もとの中身", bad.join(" / ")); });
    (ctx.App.exports || []).forEach(function (ex) {
      chain = chain.then(function () {
        if (ex.preview) {
          log(ex.preview === true, app, label, "プレビュー（保存しない）", "書き出しません");
          return;
        }
        ctx = open(app, fx.name, fx.data);
        return exportBlob(ctx, ex.id).then(function (r) {
          log(r.blob.size > 0, app, label, "書き出し ." + ex.id, r.blob.size + " バイト");
          var ext = outExt(ex.id), base = app + "-" + label + (ex.id === MAIN[app] ? "" : "-" + ex.id.replace(/[^a-z0-9-]/gi, "_"));
          var outName = base + "." + ext;
          files.push(outName);
          var sent = post("/__save?name=" + encodeURIComponent(outName), r.blob);
          if (!/\.(docx|xlsx|pptx)$/i.test(outName)) return sent;
          return sent.then(function () { return checkPackage(r.blob); }).then(function (problems) {
            log(!problems.length, app, label, "ファイルのしくみ ." + ex.id, problems.join(" / "));
            return app === "doc" ? docxOracle(r.blob, fx.expect) : Promise.resolve();
          }).then(function () {
            if (ex.id === "pptx-native") return nativeCheck(r.blob, fx.expect).then(function (bad) { log(!bad.length, app, label, "本物のグラフ・ノート・切り替え", bad.join(" / ")); });
            if (ex.id !== MAIN[app] && ex.id !== "xlsx-charts") return;
            var back = open(app, "back", app === "doc" ? { html: "" } : app === "sheet" ? { v: 2, cur: 0, sheets: [{ name: "Sheet1", cells: {} }] } : { slides: [] });
            return back.App.importFile(new File([r.blob], outName)).then(function (data) {
              return checkContent(app, data, fx.expect).then(function (bad) { log(!bad.length, app, label, "読みこみなおし ." + ex.id, bad.join(" / ")); });
            });
          });
        }).catch(function (e) { log(false, app, label, "書き出し ." + ex.id, (e && e.stack) || String(e)); });
      });
    });
    return chain;
  }

  // Sheet: 何万行あっても、見えている行だけを作る（開く速さ・作ったセルの数・遠くのセルへ移る・スクロール）
  function virtualTest() {
    var cells = {}, N = 20000;
    for (var r = 1; r <= N; r++) { cells["A" + r] = String(r); cells["B" + r] = "=A" + r + "*2"; }
    cells.D7 = "折り返す\n長い\n文字";
    var t0 = performance.now(), ctx = open("sheet", "virtual", { v: 2, cur: 0, sheets: [{ name: "S", cells: cells, style: { D7: { wrap: 1 } }, merges: ["E30:F60"], freeze: { r: 1, c: 0 } }] });
    var ms = Math.round(performance.now() - t0), tds = document.querySelectorAll("#grid td[data-c]").length;
    log(ms < 3000 && tds < 6000, "sheet", "virtual", "2 万行を開く", ms + " ms・セル " + tds + " 個");
    var nb = document.getElementById("namebox"), sc = document.getElementById("grid-scroll");
    function go(v) { nb.value = v; nb.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); }
    function frame() { return new Promise(function (res) { requestAnimationFrame(function () { requestAnimationFrame(res); }); }); }
    go("B15000");
    var td = document.querySelector('#grid td[data-c="1"][data-r="15000"]'), box = document.getElementById("sel-box");
    var ok = !!td && box.style.display !== "none" && Math.abs(td.getBoundingClientRect().top - box.getBoundingClientRect().top) < 2 && td.textContent === "30000";
    log(ok, "sheet", "virtual", "B15000 へ移る", td ? "セル " + td.textContent + "・枠のずれ " + Math.round(td.getBoundingClientRect().top - box.getBoundingClientRect().top) + "px" : "セルがありません");
    sc.scrollTop = 0;
    return frame().then(function () {
      var r1 = document.querySelector('#grid td[data-c="0"][data-r="2"]'), fz = document.querySelector('#grid td[data-c="0"][data-r="1"]');
      log(!!r1 && !!fz && !document.querySelector('#grid td[data-r="15000"]'), "sheet", "virtual", "上へ戻る", r1 ? "" : "2 行目がありません");
      // 結合したセル（E30:F60）の途中から見えるときも、左上のセルを作る
      sc.scrollTop = (document.querySelector('#grid tr[data-r="2"]').offsetTop + 20 * 29) * 1;
      return frame();
    }).then(function () {
      var m = document.querySelector('#grid td[data-c="4"][data-r="30"]');
      log(!!m && m.rowSpan === 31, "sheet", "virtual", "結合したセル", m ? "rowspan " + m.rowSpan : "左上のセルがありません");
      // 折り返した行は、測った高さで位置を計算する（下の行の枠がずれない）
      go("A9");
      var a9 = document.querySelector('#grid td[data-c="0"][data-r="9"]');
      log(!!a9 && Math.abs(a9.getBoundingClientRect().top - box.getBoundingClientRect().top) < 2, "sheet", "virtual", "折り返した行の下", "");
      // 何行目まであるか（下の空の行の高さ）が、行の数と合っている
      var p = document.querySelector('#grid tr[data-pad="bot"]'), rows = document.querySelectorAll("#grid tbody tr[data-r]"), last = rows[rows.length - 1];
      var left = p && last ? parseFloat(p.style.height) / last.offsetHeight : 0, want = N - +last.getAttribute("data-r");
      log(left >= want && left <= want + 200, "sheet", "virtual", "下の空の行", "残り " + Math.round(left) + " 行分（" + want + " 行より少し多いはず）");
    });
  }

  // Google ドライブに保存する（drive-mock.js の、テスト用のドライブで）: つなぐ・作る・一覧・読む・続けて保存・名前の変更・コピー・ごみ箱・ログインが切れる
  function driveTest() {
    var S = window.OfficeStore, D = window.OfficeDrive, M = window.DRIVE_MOCK, data = window.OFFICE_FIXTURES.sheet[0].data, f, c;
    function step(ok, name, msg) { log(!!ok, "drive", "mock", name, msg || ""); }
    return D.disconnect().then(function () {
      step(D.state() === "off" && S.place() === "local", "つなぐ前", D.state() + "・" + S.place());
      return D.connect("test@example.com");
    }).then(function (ok) {
      S.setPlace("drive");
      step(ok && D.state() === "ready" && S.place() === "drive" && D.account().email === "test@example.com", "つなぐ", D.state() + "・" + D.account().email);
      return S.create("sheet", "売上/表", "tpl-sales", copy(data));
    }).then(function (x) {
      f = x;
      var raw = M.files[f.id.slice(2)], folder = raw && M.files[raw.parents[0]];
      step(/^g:/.test(f.id) && f.place === "drive" && raw.name === "売上_表.sksheet" && raw.appProperties.skApp === "sheet" && folder && folder.name === "SK Hub Office", "作る", raw ? raw.name + "（" + (folder && folder.name) + "）" : "ファイルがありません");
      return S.list();
    }).then(function (all) {
      var hit = all.filter(function (x) { return x.id === f.id; })[0];
      step(hit && hit.place === "drive" && hit.name === "売上_表" && hit.app === "sheet", "一覧", hit ? hit.name : "一覧にありません");
      return S.get(f.id);
    }).then(function (g) {
      step(g && g.data.sheets[0].cells.A1 === "品名" && g.tpl === "tpl-sales", "読む", g ? g.name : "読めません");
      // 自動保存で続けて保存しても、送るのは「いま送っている分」と「いちばん新しい分」だけ
      M.saves = 0;
      var ps = [];
      for (var i = 0; i < 6; i++) { g.data.sheets[0].cells.Z1 = String(i); ps.push(S.save(JSON.parse(JSON.stringify(g)))); }
      return Promise.all(ps).then(function () {
        var last = JSON.parse(M.files[g.id.slice(2)].text).data.sheets[0].cells.Z1;
        step(M.saves <= 2 && last === "5", "続けて保存", M.saves + " 回送った・最後の中身 " + last);
        return S.rename(g, "新しい名前");
      }).then(function () {
        step(M.files[g.id.slice(2)].name === "新しい名前.sksheet" && g.name === "新しい名前", "名前の変更", M.files[g.id.slice(2)].name);
        return S.copy(g, "コピー");
      }).then(function (x) {
        c = x;
        step(/^g:/.test(c.id) && c.id !== g.id && M.files[c.id.slice(2)].name === "コピー.sksheet", "コピー", M.files[c.id.slice(2)] && M.files[c.id.slice(2)].name);
        return S.remove(g.id);
      }).then(function () {
        step(M.files[g.id.slice(2)].trashed, "ごみ箱", "");
        // ログインが切れる（ほかのトークンになった）→ Error("auth")・つなぎなおしが要る状態 → つなぎなおすと読める
        M.token = "other";
        return S.get(c.id).then(function () { step(false, "ログインが切れる", "読めてしまいました"); }, function (e) { step(e && e.message === "auth" && D.state() === "needs-auth" && S.place() === "drive", "ログインが切れる", D.state()); });
      }).then(function () {
        M.nextToken = "tok-2";
        return D.connect();
      }).then(function () { return S.get(c.id); }).then(function (x) { step(x && x.name === "コピー" && D.state() === "ready", "つなぎなおす", x ? x.name : "読めません"); });
    }).catch(function (e) { step(false, "エラー", (e && e.stack) || String(e)); }).then(function () {
      S.setPlace("local");
      return D.disconnect();
    }).then(function () { step(D.state() === "off" && S.place() === "local", "外す", D.state()); });
  }

  // docx（dolanmiu）で作った、本物の wp:anchor（wrapNone）を読みこめるか。
  // アプリの書き出しは、いまの軽い実装のまま。ここでは「別の作り手のファイル」を読む力を確かめる
  function thirdPartyAnchorTest() {
    if (!window.SK_DOCX_GEN || !window.SK_DOCX_GEN.Packer) return Promise.resolve();
    var D = window.SK_DOCX_GEN, b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
    var bin = atob(b64), u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    var img = new D.ImageRun({
      type: "png", data: u8, transformation: { width: 20, height: 20 },
      floating: {
        horizontalPosition: { relative: D.HorizontalPositionRelativeFrom.COLUMN, offset: 30 * 9525 },
        verticalPosition: { relative: D.VerticalPositionRelativeFrom.PARAGRAPH, offset: 12 * 9525 },
        wrap: { type: D.TextWrappingType.NONE }
      }
    });
    var doc = new D.Document({ sections: [{ children: [new D.Paragraph({ children: [new D.TextRun("third party"), img] })] }] });
    return D.Packer.toBlob(doc).then(function (blob) {
      var back = open("doc", "third-party", { html: "" });
      return back.App.importFile(new File([blob], "third-party.docx")).then(function (data) {
        var holder = document.createElement("div"); holder.innerHTML = data.html;
        var image = holder.querySelector("img[data-anc]"), text = holder.textContent || "", bad = [];
        if (!image) bad.push("data-anc の画像がありません");
        else {
          if (image.style.position !== "absolute") bad.push("position が " + image.style.position);
          if (Math.abs(parseFloat(image.style.left) - 30) > 1) bad.push("left が " + image.style.left);
          if (Math.abs(parseFloat(image.style.top) - 12) > 1) bad.push("top が " + image.style.top);
        }
        if (text.indexOf("third party") < 0) bad.push("文字がありません");
        log(!bad.length, "doc", "third-party", "docx の anchor を読みこむ", bad.join(" / ") || "left 30px・top 12px");
      });
    }).catch(function (e) { log(false, "doc", "third-party", "docx の anchor を読みこむ", (e && e.stack) || String(e)); });
  }

  var list = [];
  ["doc", "sheet", "slide"].forEach(function (app) {
    (window.OFFICE_TEMPLATES[app] || []).forEach(function (tp) { list.push([app, { id: "tpl-" + tp.id, name: tp.name, data: tp.data }]); });
    (window.OFFICE_FIXTURES[app] || []).forEach(function (fx) { list.push([app, fx]); });
  });
  var only = /[?&]only=(\w+)/.exec(location.search);
  if (only) list = list.filter(function (x) { return x[0] === only[1]; });
  window.addEventListener("error", function (e) { log(false, "-", "-", "エラー", e.message + " " + (e.filename || "") + ":" + (e.lineno || "")); });
  list.reduce(function (p, x) { return p.then(function () { return runOne(x[0], x[1]); }); }, Promise.resolve()).then(function () {
    if (!only || only[1] === "sheet") return virtualTest();
  }).then(function () {
    if (!only || only[1] === "drive") return driveTest();
  }).then(function () {
    if (!only || only[1] === "doc") return thirdPartyAnchorTest();
  }).catch(function (e) { log(false, "-", "-", "エラー", (e && e.stack) || String(e)); }).then(function () {
    var ng = results.filter(function (r) { return !r.ok; }).length;
    var sum = document.getElementById("sum");
    sum.className = ng ? "ng" : "ok";
    sum.textContent = ng ? results.length + " 件中 " + ng + " 件が NG です" : "すべて OK（" + results.length + " 件）";
    window.TEST_DONE = { results: results, files: files };
    return post("/__done", JSON.stringify(window.TEST_DONE));
  });
})();
