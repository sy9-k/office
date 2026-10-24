// テストに使うファイルの中身（window.OFFICE_FIXTURES[アプリ] = [{ id, name, data, expect }]）
//   テンプレート（office/shared/templates.js）に加えて、できることを 1 つのファイルにできるだけ詰めこんだもの。
//   expect … 書き出して読みこみなおしたときに残っていてほしいもの（run.js が確かめる）
//     Doc:   { text: [本文に入っている文字] }
//     Sheet: { cells: { "シート名!A1": "見える文字" }, sheets: シートの数 }
//     Slide: { slides: 枚数, text: [どこかに入っている文字], notes: [ノートの文字] }
(function () {
  "use strict";
  // 1×1 の赤い点（PNG）
  var PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
  var SVG = "data:image/svg+xml;charset=utf-8," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#16a34a"/></svg>');
  var spec = { type: "col", title: "売上", cats: ["4月", "5月"], series: [{ name: "A", values: [3, 5] }, { name: "B", values: [2, 4] }] };
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  var chartFig = '<figure class="chart" contenteditable="false" data-spec="' + esc(JSON.stringify(spec)) + '">' + window.OfficeChart.svg(spec, { w: 560, h: 320 }) + "</figure>";
  // 罫線と列の幅を指定した表（data-bd・セルの inline style・colgroup）
  var BD = "border-top:1pt solid #333333;border-right:1pt solid #333333;border-bottom:1pt solid #333333;border-left:1pt solid #333333";
  function bdCell(text, tag) { return "<" + (tag || "td") + ' style="' + BD + '">' + text + "</" + (tag || "td") + ">"; }
  var bdTable = '<table data-bd="all" style="width:200px;table-layout:fixed">' +
    '<colgroup><col style="width:120px"><col style="width:80px"></colgroup><tbody>' +
    "<tr>" + bdCell("罫線 A", "th") + bdCell("罫線 B", "th") + "</tr>" +
    "<tr>" + bdCell("セルの A") + bdCell("セルの B") + "</tr>" +
    "</tbody></table>\n";

  window.OFFICE_FIXTURES = {
    doc: [
      { id: "rich", name: "Doc のいろいろ", data: {
        html:
          "<h1>機能のテスト</h1>" +
          '<p style="text-align:center;line-height:2">中央<b>太字</b><i>斜体</i><u>下線</u><s>取り消し</s>H<sub>2</sub>O x<sup>2</sup>' +
          '<span style="color:#d32f2f;font-size:14pt;font-family:游明朝">色と大きさ</span><span style="background-color:rgb(255, 245, 157)">蛍光ペン</span>' +
          '<b><span style="color:rgb(29, 78, 216);font-family:Consolas"><code>太字のコード</code></span></b>' +
          '<a href="https://sk-hub.jp/?a=1&amp;b=2">リンク</a></p>' +
          '<p>脚注の付いた<sup class="fn-ref" data-fn="fn1">1</sup>文。<span class="cmt" data-cmt="c1">コメントの範囲</span>です。</p>' +
          '<p class="sp-test" style="margin-top:12pt;margin-bottom:6pt">段落の前後の間隔</p>' +
          '<p data-cst="cst1">カスタムスタイルの段落</p>' +
          "<h2>見出し 1</h2><h3>見出し 2</h3><h4>見出し 3</h4><blockquote>引用の文</blockquote>" +
          "<ul><li>箇条書き A<ul><li>A の下</li></ul></li><li>箇条書き B</li></ul><ol><li>番号 一</li><li>番号 二</li></ol>" +
          '<table><tbody><tr><th>見出しのセル</th><th>値</th></tr><tr><td colspan="2" style="background-color:#e3f2fd">結合したセル</td></tr><tr><td><p>中の段落</p><ul><li>中の箇条書き</li></ul></td><td><img src="' + PNG + '" alt="表の中の画像"></td></tr></tbody></table>\n' +
          bdTable +
          '<p><img src="' + PNG + '" alt="赤い点" class="wrap-l" style="width:40px"> 画像のとなり <img src="' + SVG + '" alt="SVG の画像"></p>\n' +
          chartFig +
          "<pre>1 行目\n\t2 行目（タブ）</pre>" +
          '<hr><hr class="pagebreak">' +
          "<p>改ページのあと &lt;特殊&gt; &amp; \"引用符\" 'アポストロフィ'</p>" +
          '<section class="footnotes"><hr><ol><li data-fn="fn1">脚注の文</li></ol></section>',
        page: { size: "a4", orient: "landscape", margin: "narrow", cols: 2, header: "ヘッダーの文", footer: "フッターの文", num: true },
        comments: [{ id: "c1", text: "コメントの本文", author: "テストの人", ts: 1760000000000 }],
        styles: [{ id: "cst1", name: "重要な行", base: "p", css: "font-weight:bold;color:#b91c1c;font-size:14pt" }]
      }, expect: { text: ["機能のテスト", "太字", "色と大きさ", "コメントの範囲", "見出し 1", "A の下", "番号 二", "結合したセル", "中の箇条書き", "2 行目", "<特殊> & \"引用符\"", "段落の前後の間隔", "カスタムスタイルの段落", "罫線 A", "セルの B"],
        footnotes: ["脚注の文"], comments: [["コメントの範囲", "コメントの本文", "テストの人"]],
        style: [{ sel: "p[style*='margin-top']", css: { "margin-top": "12pt", "margin-bottom": "6pt" } },
          { sel: "table[data-bd]", css: { width: "200px", "table-layout": "fixed" } },
          { sel: "table[data-bd] col", css: { width: "120px" } },
          { sel: "table[data-bd] td", css: { "border-top-width": "1pt", "border-top-style": "solid", "border-top-color": "rgb(51, 51, 51)", "border-left-width": "1pt" } }],
        attr: [{ sel: "p[data-cst]", name: "data-cst", value: "cst1" }],
        styles: [{ id: "cst1", name: "重要な行", base: "p", css: "font-weight:bold;color:#b91c1c;font-size:14pt" }] } },
      { id: "anchor", name: "Doc anchored image", data: { html: '<p style="position:relative"><img src="' + PNG + '" alt="floating" data-anc="1" style="position:absolute;left:30px;top:12px;width:20px;height:20px;z-index:1"></p><p>floating image</p>' },
        expect: { text: ["floating image"],
          attr: [{ sel: "img[data-anc]", name: "data-anc", value: "1" }],
          style: [{ sel: "img[data-anc]", css: { position: "absolute", left: "30px", top: "12px" } }] } },
      { id: "vert", name: "縦書き", data: { html: "<h1>縦書きの文書</h1><p>吾輩は猫である。名前はまだ無い。</p>", page: { size: "b5", vert: true } },
        expect: { text: ["縦書きの文書", "吾輩は猫である"] } },
      // デザイン: 見出しの色・フォント・h2 の線が、.docx の styles.xml に書かれる
      { id: "design", name: "デザインのテスト", data: { html: "<h1>表題</h1><h2>見出し 1</h2><h3>見出し 2</h3><blockquote>引用</blockquote><p>本文</p>", page: { design: "washi" } },
        expect: { text: ["表題", "見出し 1", "見出し 2", "引用", "本文"] } },
      { id: "empty", name: "からっぽ", data: { html: "" }, expect: { text: [] } }
    ],

    sheet: [
      { id: "rich", name: "Sheet のいろいろ", data: { v: 2, cur: 0, sheets: [
        { name: "売上", cells: {
            A1: "品名", B1: "数", C1: "単価", D1: "金額", E1: "日付", F1: "新しい関数",
            A2: "りんご", B2: "3", C2: "120", D2: "=B2*C2", E2: "2026/10/22", F2: '=XLOOKUP("みかん",A2:A3,D2:D3)',
            A3: "みかん", B3: "5", C3: "80", D3: "=B3*C3", E3: "2026/10/23", F3: '=TEXTJOIN("・",TRUE,A2:A3)',
            A4: "合計", B4: "=SUM(B2:B3)", D4: "=SUM(D2:D3)", F4: '=IFS(B2>2,"多い",TRUE,"少ない")',
            F5: "=CONCAT(A2,A3)", F6: "=MAXIFS(D2:D3,B2:B3,\">1\")", F7: "=STDEV.S(B2:B3)", F8: "=IFNA(1,0)", F9: "=1/0", F10: "=F10+1",
            G2: "TRUE", G3: "'0123", G4: "=データ!A1&\"です\"", G5: "=_xlfn.XLOOKUP(5,B2:B3,A2:A3,\"なし\")", G6: "50%", G7: "¥1,200",
            A6: "<&>\"'特殊な文字\u0001", A7: "1行目\n2行目"
          },
          style: {
            A1: { b: 1, bg: "#ffeb3b", al: "c", va: "m", bd: { t: 1, b: 2, l: 1, r: 1 } }, D2: { nf: "#,##0" }, D3: { nf: "#,##0" }, D4: { nf: "¥#,##0", b: 1 },
            E2: { nf: "yyyy/m/d" }, E3: { nf: "yyyy\"年\"m\"月\"d\"日\"" }, A6: { i: 1, u: 1, s: 1, fs: 14, fg: "#1d4ed8", ff: "游明朝" }, A7: { wrap: 1 }, H9: { bg: "#e3f2fd" }
          },
          cw: { 0: 140, 5: 160 }, rh: { 1: 30, 7: 44, 12: 40 }, merges: ["A6:C6"], freeze: { r: 1, c: 1 }, head: 1, total: 4,
          filter: { r: 1, c1: 0, c2: 4 },
          cf: [{ ref: "B2:B3", type: "gt", v: "3", color: "#ffcdd2" }, { ref: "C2:C3", type: "lt", v: "100", color: "#c8e6c9" }, { ref: "B2:B3", type: "eq", v: "5", color: "#bbdefb" },
               { ref: "A2:A3", type: "text", v: "り", color: "#fff9c4" }, { ref: "A2:A4", type: "dup", color: "#f8bbd0" }, { ref: "D2:D3", type: "bar", color: "#2a78d6" }, { ref: "C2:C3", type: "scale" }],
          charts: [{ id: "ch1", type: "col", ref: "A1:B3", title: "数", x: 400, y: 40, w: 360, h: 220 }, { id: "ch2", type: "pie", ref: "A1:B3", title: "円", x: 400, y: 280, w: 300, h: 220 },
            { id: "ch3", type: "bar", ref: "A1:C3", title: "", x: 40, y: 520, w: 320, h: 200 }, { id: "ch4", type: "line", ref: "A1:C3", title: "折れ線", x: 400, y: 520, w: 320, h: 200 },
            { id: "ch5", type: "area", ref: "A1:B3", title: "面", x: 40, y: 740, w: 320, h: 200 }, { id: "ch6", type: "scatter", ref: "B1:C3", title: "散布", x: 400, y: 740, w: 320, h: 200 }],
          dv: [
            { ref: "B2:B3", type: "whole", op: "between", f1: "1", f2: "10", allowBlank: true, promptTitle: "数を入れてください", prompt: "1〜10 の整数", errTitle: "入力を確認", err: "1〜10 の整数を入れてください", style: "stop" },
            { ref: "C2:C3", type: "decimal", op: "greaterThan", f1: "0", f2: "", allowBlank: true, promptTitle: "", prompt: "", errTitle: "", err: "", style: "warning" },
            { ref: "G2:G3", type: "list", op: "between", f1: "TRUE,FALSE", f2: "", allowBlank: true, promptTitle: "", prompt: "", errTitle: "", err: "", style: "information" }
          ],
          notes: { A2: { text: "りんごのメモ", author: "テストの人", ts: 1760000000000 }, D2: { text: "金額は数式です", author: "テストの人", ts: 1760000000000 } } },
        { name: "データ", cells: { A1: "データ", B1: "=売上!B2*2" }, grid: false },
        { name: "空白 のシート", cells: {} }
      ] }, expect: { sheets: 3, cf: { "売上": "bar,dup,eq,gt,lt,scale,text" },
        charts: { "売上": [
          { type: "col", ref: "A1:B3", title: "数", x: 400, y: 40, w: 360, h: 220 }, { type: "pie", ref: "A1:B3", title: "円", x: 400, y: 280, w: 300, h: 220 },
          { type: "bar", ref: "A1:C3", title: "", x: 40, y: 520, w: 320, h: 200 }, { type: "line", ref: "A1:C3", title: "折れ線", x: 400, y: 520, w: 320, h: 200 },
          { type: "area", ref: "A1:B3", title: "面", x: 40, y: 740, w: 320, h: 200 }, { type: "scatter", ref: "B1:C3", title: "散布", x: 400, y: 740, w: 320, h: 200 }] },
        cells: { "売上!A1": "品名", "売上!D2": "360", "売上!D4": "¥760", "売上!E2": "2026/10/22", "売上!F2": "400", "売上!F3": "りんご・みかん", "売上!F4": "多い",
        "売上!F5": "りんごみかん", "売上!G3": "0123", "売上!G4": "データです", "売上!G5": "みかん", "データ!B1": "6", "売上!A7": "1行目\n2行目" },
        dv: { "売上": [
          { ref: "B2:B3", type: "whole", op: "between", f1: "1", f2: "10", style: "stop", promptTitle: "数を入れてください", prompt: "1〜10 の整数", errTitle: "入力を確認", err: "1〜10 の整数を入れてください" },
          { ref: "C2:C3", type: "decimal", op: "greaterThan", f1: "0", f2: "", style: "warning", promptTitle: "", prompt: "", errTitle: "", err: "" },
          { ref: "G2:G3", type: "list", op: "between", f1: "TRUE,FALSE", f2: "", style: "information", promptTitle: "", prompt: "", errTitle: "", err: "" }
        ] },
        notes: { "売上!A2": { text: "りんごのメモ", author: "テストの人" }, "売上!D2": { text: "金額は数式です", author: "テストの人" } } } },
      { id: "big", name: "大きな表", data: (function () {
        var cells = {};
        for (var r = 1; r <= 3000; r++) { cells["A" + r] = String(r); cells["B" + r] = "=A" + r + "*2"; cells["C" + r] = "名前" + r; }
        return { v: 2, cur: 0, sheets: [{ name: "Sheet1", cells: cells }] };
      })(), expect: { sheets: 1, cells: { "Sheet1!B3000": "6000", "Sheet1!C2999": "名前2999" } } }
    ],

    slide: [
      { id: "rich", name: "Slide のいろいろ", data: { theme: "ocean", ratio: "16:9", slides: [
        { layout: "title", title: "発表のタイトル", sub: "サブタイトル\n2 行目", notes: "最初のノート\n2 行目のノート", trans: "fade" },
        { layout: "bullets", title: "箇条書き", items: ["一つ目", "二つ目 <&>", ""], notes: "箇条書きのノート", trans: "push", bg: "#fff7ed" },
        { layout: "two", title: "2 つ", items: ["左"], items2: ["右"], trans: "zoom" },
        { layout: "section", title: "セクション" },
        { layout: "titleOnly", title: "ものを置く", objs: [
          { id: "t1", type: "text", x: 5, y: 20, w: 30, h: 10, html: "<b>テキスト</b> ボックス<br>2 行目", fs: 24, b: 1, color: "#b91c1c", align: "c" },
          { id: "s1", type: "shape", shape: "round", x: 40, y: 20, w: 20, h: 15, html: "図形の文字", fill: "#2563eb", line: "#111111", lw: 2.25, rot: 15 },
          { id: "s2", type: "shape", shape: "line", x: 5, y: 40, w: 50, h: 1, line: "#ea580c", lw: 6 },
          { id: "s3", type: "shape", shape: "star", x: 70, y: 20, w: 15, h: 15, fill: "none", anim: "fade" },
          { id: "i1", type: "image", x: 5, y: 50, w: 10, h: 10, src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==", rot: 30, anim: "rise" },
          { id: "i2", type: "image", x: 20, y: 50, w: 10, h: 10, src: SVG },
          { id: "c1", type: "chart", x: 40, y: 50, w: 40, h: 40, spec: spec, anim: "zoom" },
          { id: "g1", type: "group", x: 86, y: 4, w: 11, h: 38, rot: 5, children: [
            { id: "g1a", type: "shape", shape: "rect", x: 0, y: 0, w: 100, h: 45, html: "グループの図形", fill: "#16a34a", line: "#111111", lw: 4.5 },
            { id: "g1b", type: "text", x: 0, y: 55, w: 100, h: 45, html: "グループの文字", fs: 12 }
          ] },
          { id: "tb1", type: "table", x: 5, y: 68, w: 30, h: 22, head: 1, fs: 14, rows: [["品名", "数"], ["りんご <&>", "3"], ["みかん", "5\n(複数行)"]] }
        ], notes: "ものの<ノート> & \"引用\"" },
        { layout: "blank" }
      ] }, expect: { slides: 6, text: ["発表のタイトル", "サブタイトル", "一つ目", "二つ目 <&>", "左", "右", "セクション", "テキスト", "図形の文字", "グループの図形", "グループの文字"], groups: 1, groupChildren: 2, lw: [2.25, 6, 4.5], notes: ["最初のノート", "箇条書きのノート", "ものの<ノート> & \"引用\""], chart: true, transitions: 3, tables: [{ rows: [["品名", "数"], ["りんご <&>", "3"], ["みかん", "5\n(複数行)"]], head: 1 }] } },
      { id: "washi", name: "和紙のテーマ", data: { theme: "washi", slides: [{ layout: "title", title: "和紙のタイトル", sub: "明朝の見た目" }, { layout: "section", title: "セクション" }] }, expect: { slides: 2, text: ["和紙のタイトル", "明朝の見た目"] } },
      { id: "four", name: "4:3 と夜のテーマ", data: { theme: "night", ratio: "4:3", slides: [{ layout: "title", title: "4:3" }] }, expect: { slides: 1, text: ["4:3"] } }
    ]
  };
})();
