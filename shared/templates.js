// SK Hub Office のテンプレート（Doc・Sheet・Slide）
//   window.OFFICE_TEMPLATES[app] = [{ id, name, cat, desc, data }]。いちばん上は白紙（cat: "blank"）
//   cat … "blank"（白紙）・"school"（学校）・"work"（仕事）・"life"（くらし）。テンプレートの一覧のしぼりこみに使う
//   data … ファイルの中身の最初の形（/office/shared/office.js が、選んだときにそのまま写してファイルにする）
//     Doc:   { html, page?: { size, orient, margin } }  … 文書の本文（h1・h2・p・ul・ol・table・img）と用紙
//     Sheet: { cells, fmt, style, head, total, sheets?, cur? } … セル（= で始まると数式）・列の書式・セルの見た目・複数シート
//     Slide: { slides: [{ layout: "title|bullets|section", title, sub, items, notes, image }] }
(function () {
  "use strict";

  // Sheet: 2 次元の配列から cells を作る（1 行目が rows[0]）
  function grid(rows) {
    var cells = {};
    rows.forEach(function (row, r) {
      row.forEach(function (v, c) {
        if (v === "" || v === null || v === undefined) return;
        cells[String.fromCharCode(65 + c) + (r + 1)] = String(v);
      });
    });
    return cells;
  }

  window.OFFICE_TEMPLATES = {
    doc: [
      { id: "blank", name: "白紙の文書", cat: "blank", desc: "何も書いていない文書", data: { html: "<h1>無題の文書</h1><p><br></p>" } },
      { id: "report", name: "レポート", cat: "school", desc: "調べたことをまとめる", data: { html:
        "<h1>○○についてのレポート</h1>" +
        "<p class=\"meta\">2 年 1 組　名前：　　　　　　　提出日：　月　日</p>" +
        "<h2>1. はじめに</h2><p>このテーマを選んだ理由と、調べようと思ったきっかけを書きます。</p>" +
        "<h2>2. 調べたこと</h2><p>本やインターネットで調べたこと、実験や観察でわかったことを書きます。</p>" +
        "<ul><li>わかったこと 1</li><li>わかったこと 2</li><li>わかったこと 3</li></ul>" +
        "<h2>3. まとめ</h2><p>調べてわかったことと、これから調べてみたいことを書きます。</p>" +
        "<h2>参考にしたもの</h2><ol><li>本の名前・作った人・出版社</li><li>Web サイトの名前・URL・見た日</li></ol>" } },
      { id: "notice", name: "お知らせ", cat: "life", desc: "行事やイベントを知らせる", data: { html:
        "<h1>秋の地域清掃のお知らせ</h1>" +
        "<p>いつもご協力ありがとうございます。今年も、地域の清掃を行います。お近くの方は、ぜひご参加ください。</p>" +
        "<table><tbody><tr><th>日時</th><td>11 月 9 日（日）9:00〜10:30（雨の日はお休み）</td></tr>" +
        "<tr><th>集合場所</th><td>中央公園の東口</td></tr><tr><th>持ち物</th><td>軍手・飲み物・タオル</td></tr></tbody></table>" +
        "<h2>お問い合わせ</h2><p>自治会　清掃係（○○）まで</p>" } },
      { id: "minutes", name: "議事録", cat: "work", desc: "話し合いで決まったことを残す", data: { html:
        "<h1>○○の打ち合わせ　議事録</h1>" +
        "<table><tbody><tr><th>日時</th><td>10 月 21 日（火）16:00〜17:00</td></tr><tr><th>場所</th><td>会議室 A・オンライン</td></tr><tr><th>出席</th><td>○○・○○・○○</td></tr></tbody></table>" +
        "<h2>話し合ったこと</h2><ul><li>○○の進み具合</li><li>次の発表の準備</li></ul>" +
        "<h2>決まったこと</h2><ul><li>○○は来週の金曜日までに作る</li><li>発表の練習は 10 月 28 日に行う</li></ul>" +
        "<h2>次にやること</h2><ol><li>資料を作る（担当：○○）</li><li>会場を予約する（担当：○○）</li></ol>" } },
      { id: "proposal", name: "企画書", cat: "work", desc: "やりたいことを提案する", data: { html:
        "<h1>○○の企画書</h1><p class=\"meta\">2026 年 10 月 21 日　提案：○○</p>" +
        "<h2>やりたいこと</h2><p>何を、だれのために、どうしてやりたいのかを、短く書きます。</p>" +
        "<h2>進め方</h2><ol><li>準備（11 月）</li><li>試してみる（12 月）</li><li>みんなに広める（1 月）</li></ol>" +
        "<h2>必要なもの</h2><table><tbody><tr><th>もの</th><th>数</th><th>費用</th></tr><tr><td>○○</td><td>2</td><td>3,000 円</td></tr><tr><td>○○</td><td>1</td><td>1,500 円</td></tr></tbody></table>" +
        "<h2>期待できること</h2><p>この企画で、どんなよいことがあるかを書きます。</p>" } },
      { id: "letter", name: "お礼の手紙", cat: "life", desc: "お世話になった人へ", data: { html:
        "<p class=\"right\">2026 年 10 月 21 日</p><p>○○様</p>" +
        "<p>拝啓　さわやかな秋晴れの日が続いていますが、いかがお過ごしでしょうか。</p>" +
        "<p>先日は、○○をしていただき、本当にありがとうございました。おかげで、○○ができました。</p>" +
        "<p>寒くなってきましたので、どうぞお体に気をつけてお過ごしください。</p>" +
        "<p class=\"right\">敬具</p><p class=\"right\">○○より</p>" } }
    ],

    sheet: [
      { id: "blank", name: "空白のブック", cat: "blank", desc: "何も入っていない表", data: { cells: {}, fmt: {} } },
      { id: "sales", name: "売上の記録", cat: "school", desc: "模擬店などの売上を数える", data: {
        head: 1, total: 12, fmt: { C: "yen", E: "yen" },
        cells: (function () {
          var rows = [["日付", "品目", "単価", "数量", "売上", "メモ"]];
          [["10/18", "焼きそば", 300, 42], ["10/18", "フランクフルト", 200, 35], ["10/18", "ラムネ", 150, 58], ["10/18", "わたあめ", 200, 27],
           ["10/19", "焼きそば", 300, 61], ["10/19", "フランクフルト", 200, 44], ["10/19", "ラムネ", 150, 73], ["10/19", "わたあめ", 200, 39],
           ["10/19", "チョコバナナ", 250, 30], ["10/19", "チョコバナナ", 250, 26]].forEach(function (r, i) {
            rows.push([r[0], r[1], r[2], r[3], "=C" + (i + 2) + "*D" + (i + 2), i === 4 ? "午前で完売" : ""]);
          });
          rows.push(["合計", "", "", "=SUM(D2:D11)", "=SUM(E2:E11)", ""]);
          return grid(rows);
        })() } },
      { id: "budget", name: "おこづかい帳", cat: "life", desc: "入ったお金と使ったお金", data: {
        head: 1, total: 9, fmt: { C: "yen", D: "yen" },
        cells: grid([
          ["日付", "内容", "入ったお金", "使ったお金", "メモ"],
          ["10/1", "おこづかい", 3000, "", ""], ["10/3", "ノート", "", 240, "2 冊"], ["10/6", "お菓子", "", 180, ""],
          ["10/12", "本", "", 880, "まんが"], ["10/15", "お手伝い", 500, "", "そうじ"], ["10/18", "文化祭", "", 650, ""],
          ["10/20", "電車", "", 420, "往復"],
          ["合計", "", "=SUM(C2:C8)", "=SUM(D2:D8)", ""],
          ["残り", "", "=C9-D9", "", ""]
        ]) } },
      { id: "timetable", name: "時間割", cat: "school", desc: "月曜日から金曜日まで", data: {
        head: 1, fmt: {},
        cells: grid([
          ["", "月", "火", "水", "木", "金"],
          ["1", "国語", "数学", "英語", "理科", "社会"], ["2", "数学", "英語", "国語", "数学", "理科"],
          ["3", "英語", "体育", "理科", "社会", "国語"], ["4", "社会", "国語", "数学", "英語", "体育"],
          ["5", "理科", "音楽", "美術", "体育", "数学"], ["6", "体育", "総合", "", "学活", ""]
        ]) } },
      { id: "roster", name: "名簿", cat: "work", desc: "名前と連絡先をまとめる", data: {
        head: 1, fmt: {},
        cells: grid([
          ["番号", "名前", "ふりがな", "係", "連絡先"],
          ["1", "青木 さくら", "あおき さくら", "図書", "—"], ["2", "井上 はると", "いのうえ はると", "体育", "—"],
          ["3", "上田 ゆい", "うえだ ゆい", "保健", "—"], ["4", "江藤 そうた", "えとう そうた", "放送", "—"],
          ["5", "小川 ひな", "おがわ ひな", "美化", "—"]
        ]) } },
      { id: "shopping", name: "買い物リスト", cat: "life", desc: "買うものと値段", data: {
        head: 1, total: 7, fmt: { C: "yen", D: "yen" },
        cells: grid([
          ["品物", "数", "値段", "小計", "買った"],
          ["牛乳", 2, 198, "=B2*C2", "○"], ["たまご", 1, 248, "=B3*C3", "○"], ["食パン", 1, 168, "=B4*C4", ""],
          ["バナナ", 1, 128, "=B5*C5", ""], ["トマト", 3, 98, "=B6*C6", ""],
          ["合計", "=SUM(B2:B6)", "", "=SUM(D2:D6)", ""]
        ]) } }
    ],

    slide: [
      { id: "blank", name: "空白のプレゼンテーション", cat: "blank", desc: "表紙だけのスライド", data: { slides: [
        { layout: "title", title: "タイトルを入力", sub: "サブタイトルを入力" }
      ] } },
      { id: "talk", name: "発表", cat: "school", desc: "表紙・もくじ・本文・まとめ", data: { slides: [
        { layout: "title", title: "○○についての発表", sub: "2 年 1 組　○○" },
        { layout: "bullets", title: "もくじ", items: ["調べたきっかけ", "調べたこと", "わかったこと", "まとめ"] },
        { layout: "bullets", title: "調べたきっかけ", items: ["身の回りで気になったこと", "みんなに知ってほしいと思ったこと"] },
        { layout: "bullets", title: "わかったこと", items: ["わかったこと 1", "わかったこと 2", "わかったこと 3"] },
        { layout: "section", title: "まとめ", sub: "ご清聴ありがとうございました" }
      ] } },
      { id: "intro", name: "自己紹介", cat: "life", desc: "好きなことや目標を伝える", data: { slides: [
        { layout: "title", title: "はじめまして、○○です", sub: "自己紹介" },
        { layout: "bullets", title: "好きなこと", items: ["休みの日にしていること", "最近はまっていること"] },
        { layout: "bullets", title: "今年の目標", items: ["がんばりたいこと", "できるようになりたいこと"] },
        { layout: "section", title: "よろしくお願いします" }
      ] } },
      { id: "pitch", name: "提案", cat: "work", desc: "企画を短く伝える", data: { slides: [
        { layout: "title", title: "○○の提案", sub: "2026 年 10 月　○○" },
        { layout: "bullets", title: "いま困っていること", items: ["○○に時間がかかる", "○○がわかりにくい"] },
        { layout: "bullets", title: "やりたいこと", items: ["○○を使って、○○を短くする", "まず 1 か月、試してみる"] },
        { layout: "bullets", title: "進め方", items: ["11 月：準備", "12 月：試す", "1 月：ふり返り"] },
        { layout: "section", title: "ご意見をください" }
      ] } },
      { id: "event", name: "イベントの案内", cat: "life", desc: "日時と場所を大きく", data: { slides: [
        { layout: "title", title: "秋のお楽しみ会", sub: "11 月 15 日（土）13:00〜　中央公民館" },
        { layout: "bullets", title: "やること", items: ["ゲーム大会", "クイズ", "おやつの時間"] },
        { layout: "bullets", title: "持ち物", items: ["上ばき", "飲み物"] },
        { layout: "section", title: "みんなで来てね" }
      ] } }
    ]
  };
})();
