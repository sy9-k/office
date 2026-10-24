# SK Hub Office — いまの作りと、これからの計画

文書の **Doc**・表計算の **Sheet**・スライドの **Slide** をまとめた、SK のオフィスアプリです。開くのは **https://office.sk-hub.jp/**（開発中・検索には載せない）。紹介ページは https://sk-hub.jp/products/office/ 。

目標は「LibreOffice くらいの機能をそろえて、Google Workspace のように、クラウドに保存・共有・同時に編集できる」こと（下の「決めたこと」）。画面の形は Microsoft Office、見た目は SK Material（Google の Material Design 3 がもと）。

---

## 続きを始めるとき（人も AI も、まずここ）

1. このファイルを読む
2. `office/shared/office.js` のいちばん上の説明（画面の流れ・保存・アカウント・PWA）
3. `office/shared/store.js` のいちばん上の説明（ファイルの保存）
4. `_workers/office/README.md`（office.sk-hub.jp のしくみ）
5. 作るアプリのファイル（`doc.js`・`sheet.js`・`slide.js`）のいちばん上の説明。全部は読まなくてよい

### 手元で動かす

```
python -m http.server 8765
```

リポジトリのいちばん上で動かして、http://localhost:8765/office/ を開く（`.claude/launch.json` の `site`。Mac は `site-mac`）。build はいらない（`office/` は build で作らないページ）。紹介ページ（`_build/pages/office.html`）を直したときだけ `python _build/build.py`。

- Chrome・Edge では、最初に保存するフォルダを選ぶ。自動で試すとき（AI のブラウザなど）は、フォルダを選ぶ画面を操作できないので、ブラウザの中の場所（OPFS）を「選んだフォルダ」にすると試せる:
  `navigator.storage.getDirectory()` で作ったフォルダを、IndexedDB `sk_office` の `kv` に `folder` という名前で入れて、ページを開き直す
- 書き出したファイル（.docx・.xlsx・.pptx）は、ZIP の中の XML がこわれていないか・つながり（rels）が合っているかを確かめる。.pptx は python-pptx で開けるかも確かめた

### テスト（書き出し・読み込み・Sheet の表示）

```
python _tests/office/run.py            （Doc・Sheet・Slide を全部。--only sheet で 1 つだけ）
```

Edge か Chrome を見えないまま開いて、テンプレートと `_tests/office/fixtures.js`（できることを詰めこんだファイル）を 1 つずつ書き出し、

1. ZIP・XML・つながり（rels）・`[Content_Types].xml` を確かめる（`run.js`。ブラウザの中）
2. 書き出したファイルを読みこみなおして、文字・セルの値・脚注・コメント・ノート・条件付き書式が残っているかを確かめる（`run.js`）
3. 要素の順番（Word・Excel・PowerPoint は順番がちがうと開けない）・Excel にないエラー・`_xlfn.` を確かめ、.pptx は python-pptx で開く（`check.py`）
4. Sheet の 2 万行のファイルで、見えている行だけを作っているか・遠くのセルへ移れるか（`run.js` の virtualTest）
5. `.docx` は、テストのときだけ入れる別の実装（`docx-preview`）でも開けるか、`docx`（dolanmiu）が作った本物の `wp:anchor` を読みこめるかを確かめる（`run.js` の docxOracle・thirdPartyAnchorTest）

書き出したファイルは `_tests/office/out/` に残る（Word・LibreOffice で開いて見た目も確かめられる）。NG があると終了コード 1。書き出しを直したら、まずこれを動かす。新しい機能を足したら `fixtures.js` にも足す。`http://localhost:8765/_tests/office/` を開いても動く（`_` で始まるので公開されない）。`out` を Word で開いたままなどで消せないときは、`OFFICE_TEST_OUT` に別のフォルダを入れて動かす（`$env:OFFICE_TEST_OUT="_tests/office/out-tmp"`）

### 気をつけること

- **ページの中の読み込み・リンクは相対パス**（`../shared/office.js`・`../doc/`）。office.sk-hub.jp では `/`、手元では `/office/` が Office のいちばん上になるため。サイトと共通の部品だけ `/assets/…`・`/toolbox/shared/…` と書く（`_workers/office/src/worker.js` の `SHARED`）。
- **sk-hub.jp/office/ を開くと office.sk-hub.jp に移る**（各ページの `<head>` の最初のスクリプト）。手元（localhost）では移らない。
- **サービスワーカー**（`office/sw.js`）: 画面の部品を足したら `SHELL` に入れ、大きく変えたら `VERSION` を上げる（いまは `sk-office-v18`）。sk-hub.jp のものはネットを先に使うので、ふだんはそのままで新しいものが出る。
- **ファイルの中身の形を変えるとき**は、`store.js` の `FILE_VERSION`（いまは 2）を上げて、古い形も読めるようにする（各アプリの `upgrade`。使う人の PC に古いファイルが残っている）。
- **アプリの中身は、アプリのファイルに書く**。`office.js` は画面の枠（スタート画面・タイトルバー・リボン・ステータスバー・元に戻す・保存・検索の欄・印刷）だけ。アプリには道具 `O` を渡す（`office.js` の「アプリに渡す道具」）。
- **まだできないボタン**は `todo: 1`（リボン）か `data-todo="名前"` を付けておくと、押したときに「まだ使えません」と出る。
- 文は「です・ます」で、やさしく短く（SK's Brand の「ことば」）。JavaScript の中の文は `t("…")` で書く（多言語。訳は `_build/i18n/`）。
- この Windows の PC では、PowerShell に `git` がない。GitHub Desktop の中の git（`%LOCALAPPDATA%\GitHubDesktop\app-*\resources\app\git\cmd\git.exe`）は使えるが、GitHub へのプッシュと PR は GitHub Desktop で行う。

---

## いまの状態（バージョン 0.3・2026-10）

### 共通

- **起動画面** → **スタート画面**（ホーム・新規・開く・アカウント・オプション）→ **エディタ**。ファイルを開いているときのスタート画面（バックステージ）には、情報（版の履歴）・コピーを保存・エクスポート・印刷・閉じる
- **テンプレート**: Doc・Sheet・Slide それぞれ白紙と 5 種類（`templates.js`）
- **ファイルの保存**: PC の選んだフォルダに 1 つずつ（`名前.skdoc`・`.sksheet`・`.skslide`。Chrome・Edge のパソコン）。それ以外のブラウザは、ブラウザの中に保存して、書き出し・読み込みで代わりにする。最近使ったアイテム・ピン留め・検索・名前の変更・コピー・削除
- **版の履歴**: 開いたとき・保存したとき・10 分ごとに、前の版をこのブラウザ（IndexedDB）に残す（20 個まで）。戻せる
- **元に戻す・やり直し**（どのアプリも。中身をまるごと覚える。100 回・60MB まで）
- **検索と置換**（Ctrl+F・Ctrl+H）、**印刷・PDF**（Ctrl+P。アプリごとに印刷用の形にする）、**閲覧モード**、**ズーム**（25〜300%）、**キーボードショートカットの一覧**
- **リボン**: タブとグループ・下向きの矢印のメニュー・色えらび。選んだものに合わせて出るタブ（表のレイアウト・図の形式・グラフのデザイン・図形の書式）
- **アカウント**（下の「SK Hub Systems アカウント」）・**オプション**（ユーザー名・テーマ・動きを減らす・保存先・自動保存・インストール）
- **PWA**: Doc・Sheet・Slide は別々のアプリとしてインストールできる（`doc/manifest.json` など。アイコンと色もアプリごと。入り口の `manifest.json` は「SK Hub Office」として入れる）・オフライン（`sw.js` は 3 つとも同じ）・`.skdoc` などをダブルクリックで開く（アプリごとの `file_handlers`）・ウィンドウのいちばん上の帯を Office のタイトルバーにする（Window Controls Overlay）
- **アプリ**（スタート画面の左のメニュー）: SK Hub のアプリ（Office・Nagi・SK Study・MALU・Y-FILTER.・SK's Toolbox のアプリ）を全部出す（`office.js` の `skApps`。SK のアプリを足したら、ここにも足す）

### Doc（`doc.js`）

- フォント・大きさ・太字・斜体・下線・取り消し線・上付き・下付き・文字の色・蛍光ペン・書式のクリア・書式のコピー
- 箇条書き・段落番号・インデント・揃え・行間、段落の前後の間隔、スタイル（標準・表題・見出し 1・見出し 2・引用）と自作のスタイル（もとにするスタイルと書式を決めて作る・変える・消す）
- 表（マス目で大きさを選ぶ。行・列の挿入と削除・見出しの行・塗りつぶし・右のセルと結合・罫線・列の幅。Tab で次のセル）
- 画像（ファイル・貼り付け・ドラッグ。大きさ・回り込み・代替テキスト）、グラフ（データを貼って作る・ダブルクリックで直す）、リンク、横線、日付、記号、改ページ
- 用紙（A4・B5・A3・Letter、縦・横、余白、段組み）、**縦書き**、ヘッダー・フッター・ページ番号（印刷したとき）、ページの区切りの線
- デザイン 11 種類（「デザイン」タブ。ビジネス・モダン・テック・レポート・モノクロ・あたたかい・自然・さくら・ラベンダー・和風・クラシック。見出しの色とフォント・h2 の下線／左の線・表の見出しの色・引用の線。`page.design` に入れる。`.docx` の styles.xml にも書く）
- 目次（見出しから作る・更新）、脚注、コメント（右の欄）、ナビゲーション（見出しの一覧）、スペル チェック（ブラウザの機能）、文字カウント
- 読み込み: `.docx`・`.html`・`.md`・`.txt` / 書き出し: `.docx`・`.html`・`.md`・`.txt`・PDF
- `.docx` の脚注・コメントは、Word の脚注（`footnotes.xml`）・コメント（`comments.xml`）として書き出し・読み込みする
- `.docx` にフォントが埋め込まれている（`word/fontTable.xml`・`word/fonts/*.odttf`）ときは、`opentype.js` でフォントを読み、画面でも同じフォントで出す。行間はフォントの ascent/descent/lineGap から Word に近い値にする（読めないときは今までどおり 1.15）
- 図の自由配置（`data-anc`）は、書き出すときも Word の `wp:anchor`＋`wp:wrapNone`（column / paragraph 基準）で残す。表のセルの中の画像は今までどおり行内

### Sheet（`sheet.js`。計算は `formula.js`）

- セルにそのまま入力（日本語の変換もそのまま）・数式バー・名前ボックス（番地で移動）・F2 で編集・Alt+Enter で改行・数式の中で押したセルの番地を入れる
- 関数 100 ほど（SUM・IF・VLOOKUP・XLOOKUP・COUNTIF・SUMIFS・TEXT・DATE など）、`$A$1`、ほかのシートの参照、循環参照の検出
- コピー・切り取り・貼り付け（参照をずらす・値だけ・ほかのアプリとの貼り付け）、オートフィル（数・日付・「第 1 回」の続き）、Ctrl+D・Ctrl+R
- 書式（フォント・大きさ・太字など・文字と塗りの色・罫線・揃え・折り返し・結合）、表示形式（桁区切り・通貨・%・日付・時刻・ユーザー定義・小数の桁）、列の幅（ドラッグ・ダブルクリックで合わせる）・行の高さ
- 行・列の挿入と削除（数式の参照も直す）、複数のシート（追加・名前の変更・複製・移動・削除）、ウィンドウ枠の固定、目盛線
- 並べ替え・フィルター・重複の削除・条件付き書式（大きい・小さい・等しい・文字を含む・重複・データ バー・カラー スケール）
- 入力規則（リスト・数値・日付・文字の長さ。止める・警告・情報から選ぶ）、セルのコメント（赤い三角・Shift+F2）
- グラフ（縦棒・横棒・折れ線・面・円・散布図。表の上に置いて動かす・大きさを変える・PNG で保存）
- デザイン（「デザイン」タブ）: テーブルのスタイル 17 種類（見出しの行の色・縞・罫線）と、セルのスタイル 14 種類（タイトル・見出し・良い・悪い・メモ・入力・出力・集計など）。どちらもセルの書式にするので、.xlsx にもそのまま書き出される
- ステータスバーに平均・個数・合計
- 読み込み: `.xlsx`（値・数式・主な書式・列の幅・結合・固定・条件付き書式・入力規則・コメント・グラフ）・`.csv` / 書き出し: `.xlsx`（条件付き書式・入力規則・コメント・グラフ（Excel のグラフ。縦棒・横棒・折れ線・面・円・散布図）も。新しい関数には `_xlfn.` を付ける）・`.csv`・PDF
- 表は、見えている行と上下 40 行だけを作る（`renderWindow`。上と下は高さだけの空の行）。2 万行でも開くのは一瞬、作るセルは 2,000 ほど

### Slide（`slide.js`）

- レイアウト（タイトル・タイトルとコンテンツ・2 つのコンテンツ・セクション見出し・タイトルのみ・白紙）の文字をそのまま書きかえる
- テキスト ボックス・図形（9 種類）・画像・グラフ・表を自由に置く。ドラッグで動かす・角で大きさを変える・上の丸で回す・ほかのものや真ん中に合わせる線・重なり順・整列・複製・コピーと貼り付け・矢印で少しずつ動かす
- グループ化・グループ解除（Ctrl+G / Ctrl+Shift+G。グループのまま動かす・大きさを変える・回す）、図形の線の太さ
- テーマ 21 種類（SK の 6 色のほか、コーポレート・インディゴ・スレート・森・ミント・夕焼け・ラベンダー・パステル・ポップ・和紙・セピア・ネイビー・エメラルド・ゴールド・ネオン。明朝のもの・暗い背景のものも。「デザイン」タブのメニューから選ぶ。色は `slide.js` の `THEMES` が唯一の定義で、CSS はそこから作る。.pptx には、テーマの色と明朝などのフォントも書く）、背景の色、スライドのサイズ（16:9・4:3）
- 画面切り替え（フェード・プッシュ・ズーム）、アニメーション（フェード・スライドイン・ズーム。順に出す）
- スライドの並べかえ（ドラッグ）・複製・削除・右クリックのメニュー、ノート
- スライド ショー（F5）と**発表者ビュー**（P。いまのスライド・次のスライド・ノート・時間）、B で画面を黒く
- 読み込み: `.pptx`（文字・図形・画像・ノート・グループ・線の太さ）・`.txt`（アウトライン） / 書き出し: `.pptx`（ノート・グループ・線の太さも）・アウトライン（`.txt`）・PDF

### Google ドライブ（保存先の 1 つ。`drive.js`）

- アカウントの「Google ドライブをつなぐ」か、開く・オプションの「Google ドライブ」で、自分の Google ドライブにつなぐ。新しいファイルの保存先を「この PC」と「Google ドライブ」で切りかえられる（`Store.place()`）
- ドライブの「SK Hub Office」フォルダに、PC と同じ形のファイル（`名前.sksheet` など）を 1 つずつ保存する。権限は `drive.file` だけ（SK Hub Office で作った・開いたファイルしか見えない）。ファイルは SK のサーバーを通らない
- ファイルの ID が `g:` で始まるものがドライブのファイル（`store.js` が ID で行き先を決める）。一覧は PC とドライブの両方を出す
- 自動保存で続けて保存しても、送るのは「いま送っている分」と「いちばん新しい分」だけ。削除はドライブのごみ箱に入れる（30 日は戻せる）
- ログイン（Google Identity Services のトークン）は 1 時間で切れる。切れたら「つなぎなおしますか？」と聞き、押したらつないで保存を続ける
- **まだ動かない**: `drive.js` の `GOOGLE_CLIENT_ID` が空のあいだは「準備中」と出る。下の「Google ドライブを使えるようにする」をしてから入れる。テストは `drive-mock.js`（本物の代わりのドライブ）で行っていて、本物の Google ドライブではまだ試していない

#### Google ドライブを使えるようにする（Google Cloud の設定。アカウントを持っている人が行う）

1. [Google Cloud コンソール](https://console.cloud.google.com/) で、SK Hub Systems アカウントの Firebase と同じプロジェクトを選ぶ
2. 「API とサービス」→「ライブラリ」で **Google Drive API** を有効にする
3. 「Google Auth Platform」（前の名前は「OAuth 同意画面」）の「ブランディング」に、アプリの名前（SK Hub Office）・サポートのメール・ホームページ（https://sk-hub.jp/products/office/）・プライバシーポリシー（https://sk-hub.jp/policies/）・承認済みドメイン（sk-hub.jp）を入れる
4. 「データアクセス」で、スコープ `https://www.googleapis.com/auth/drive.file` を足す（機密でないスコープなので、審査はブランディングの確認だけ）
5. 「クライアント」で「ウェブ アプリケーション」のクライアントを作り、「承認済みの JavaScript 生成元」に `https://office.sk-hub.jp` と `http://localhost:8765` を入れる（リダイレクト URI はいらない）
6. できたクライアント ID を `office/shared/drive.js` の `GOOGLE_CLIENT_ID` に入れる。入れる前に試すときは、ブラウザのコンソールで `localStorage.setItem("sk_office_google_client", "クライアント ID")` をして開きなおす
7. 「対象」が「テスト中」のあいだは、テストユーザーに足した Google アカウントだけが使える。みんなに使ってもらうときは「アプリを公開」する

### SK Hub Systems アカウント（ログインだけ）

- Office の「アカウント」から **ログイン** を押すと、sk-hub.jp の `/account/` でログインし、終わったら Office に戻る
- Office が使うのは **ログイン中かのヒント**（`skhub_account`: 名前・メールアドレス・写真）だけ。アイコンと名前が出る。ファイルはいままでどおり PC に保存する
- office.sk-hub.jp からは sk-hub.jp の保存場所が見えないので、sk-hub.jp の `/assets/hub/bridge.html`（見えない iframe）から受け取る（`assets/hub/bridge.js` の `ORIGINS` に `https://office.sk-hub.jp` を足した）。手元・sk-hub.jp では、そのまま読む
- ログアウト・名前の変更は、アカウントのページ（新しいタブ）で行う。Office に戻ってくると、ログインの状態を読みなおす
- まだ Firebase は Office では読み込まない（Firestore に何も保存しない）。`SERVICES` に `office` も足していない

### まだのこと（弱いところ）

- **office.sk-hub.jp でのログインの受け取りは、公開してから確かめる**（手元では同じ場所なので、bridge を通らない）
- **Microsoft Office・LibreOffice で、書き出したファイルを開いて確かめていない**。`_tests/office/` のテストで、ZIP・XML・つながり・要素の順番（書き出しでよく使う要素）・読みこみなおしは確かめている。`.docx` はテストのときだけ `docx-preview` でも開き、`docx`（dolanmiu）が作った本物のアンカー画像も読みこめる。.pptx は python-pptx で開ける。この PC に LibreOffice を入れれば、`soffice --headless --convert-to pdf` で開けるかもテストに足せる
- 印刷は、印刷用の形を作るところまで確かめた。実際の印刷・PDF の見た目はまだ確かめていない
- Doc は `contenteditable` と `document.execCommand`（古いしくみ）。変更の記録（校閲）・数式（KaTeX）・図形はない。.docx の細かい書式（タブ・均等割り付けなど）は変わることがある。ページの区切りは、画面に線を引くだけ（ページごとのヘッダー・ページの下の脚注は、印刷のときだけ）
- Doc の `.docx` 読み込みは、Word の見た目に近づけた（用紙の余白 `w:pgMar`、本文の既定のフォント・大きさ・行間 `docDefaults`、表題・見出し 1〜9・引用の組み込みスタイル、表の罫線 `tblBorders`・セルの余白 `tcMar`・列幅 `tblGrid`、図の重ね置き `wrapNone`）。埋め込みフォントは `opentype.js` で読み、画面のフォントと行間に使う。表のスタイル（`tblStyle`。罫線・セルの余白・見出し行・縞の塗り。basedOn もたどる）と文字スタイル（`rStyle`）も読みこむ。まだ合わないのは、表のスタイルの罫線の条件（見出し行だけ下線など）、段落スタイルの文字書式の継承、均等割り付け・タブ位置の細かい指定。図の自由配置は `wp:anchor` で書き出すが、`wrapSquare`・`wrapTight` の回り込みの細かい位置はまだ残らない
- Sheet の計算は、変えるたびにブック全体を計算しなおす（見えているセルだけ計算するので、表示は速い）。ピボットテーブル・テーブルの .xlsx への書き出しはない。.xlsx のグラフは、読み書きできる範囲に限りがある（飛び飛びの列・ほかのシートを指すもの・値を直接書いたもの・積み上げ・軸の設定・組み合わせグラフは、読み込むとき飛ばすか、ふつうのグラフになる）
- Slide の表は、セルの結合・セルごとの色・列の幅の変更がまだない（.pptx の表の読み込みは、結合したセルの文字だけ残る）。スライドマスター・.pptx のアニメーションの書き出しはない
- `.odt`・`.ods`・`.odp`（LibreOffice の形式）
- 多言語の訳（Office の文はまだ日本語だけのものが多い。今回足した文は訳した。訳のない一覧は `_build/i18n/missing/`）

---

## 作り

### しくみ

- **公開**: ページは sk-hub.jp（GitHub Pages）の `office/` に置き、Cloudflare Worker（`_workers/office/`）が office.sk-hub.jp で中継する。直すときはプッシュするだけ（Worker の deploy は、Worker を直したときだけ）。
- **画面**: 各ページ（`office/doc/index.html` など）は起動画面と入れ物だけ。スタート画面・エディタの枠は `office.js` が組み立て、エディタの中身はアプリのファイル（`doc.js` など）が作る。`<html data-app="doc|sheet|slide|hub">` でアプリを決める。
- **アプリのファイル**: `window.OfficeApps[アプリ] = function (O) { return { ribbon, render, command, state, … } }`。`office.js` より先に読み込む。返すもの（`upgrade`・`thumb`・`find`・`replace`・`clip`・`key`・`beforePrint`・`exports`・`exportAs`・`imports`・`importFile` など）は `office.js` の中の使い方を見る。
- **画面と編集は自作**（決めた）: 画面を置きかえる大きな部品（TipTap・HyperFormula・Konva・Univer など）は使わない。オフラインで動き、SK らしい画面にしやすいため。Office のファイルの読み書きも基本は自作（`zip.js` は ZIP、XML はブラウザの DOMParser）。ただし、ファイルの忠実さを上げる小さな headless の部品は使う: Doc の埋め込みフォントに `opentype.js`（MIT、`shared/vendor/`）。テストだけに `docx`（MIT）・`docx-preview`（Apache-2.0）・`JSZip`（MIT/GPL のどちらか）を入れる（`_tests/office/vendor/`、配信しない）。
- **色**: `office.css` の `:root[data-app="…"]` に M3 の色（`--md-*`）。Doc＝SK's Blue・Sheet＝SK's Matcha・Slide＝SK's Mikan。スライドのテーマは `.sl[data-theme]` の `--sl-*`（`slide.js` の `THEMES` と同じ値）。
- **部品**: `/toolbox/shared/m3.css`（ボタン・スイッチ・ダイアログなど）と `/toolbox/shared/m3.js`（波紋、`M3.prompt`・`M3.choose`・`M3.confirm`・`M3.dialog`）。文字は M PLUS 1、アイコンは Material Symbols Rounded。

### データ

| もの | 場所 | 形 |
|---|---|---|
| ファイル（フォルダ） | 使う人が選んだ PC のフォルダ | `{ format: "sk-hub-office", version: 2, app, tpl, created, data }` の JSON（version 1 も読める） |
| ファイル（ブラウザの中） | localStorage `sk_office_files` | `[{ id, app, name, tpl, created, updated, data }]` |
| 選んだフォルダ・版の履歴 | IndexedDB `sk_office`（`kv` の `folder`・`v:ファイルの ID`） | FileSystemDirectoryHandle・`[{ ts, size, data }]` |
| オプション | localStorage `sk_office_settings` | `{ theme, motion, autosave, user, pinned }` |
| ファイル（Google ドライブ） | 使う人の Google ドライブの「SK Hub Office」フォルダ | フォルダのファイルと同じ JSON。`appProperties.skApp` にアプリ |
| Google ドライブのログイン | localStorage `sk_office_drive` | `{ token, exp, email, name, folder, on }`（トークンは 1 時間で切れる） |
| 新しいファイルの保存先 | localStorage `sk_office_place` | `"local"` か `"drive"` |
| ログイン中かのヒント | localStorage `skhub_account`（sk-hub.jp のものを写す） | `{ name, email, photo, letter, color, … }` |

`data` の中身（くわしくは各アプリのファイルのいちばん上）:

- Doc: `{ html, page: { size, orient, margin, mx, my, cols, vert, header, footer, num, base }, comments: [{ id, text, author, ts }], styles: [{ id, name, base, css }] }`（`styles` は自作のスタイル。段落は `data-cst="ID"` で使う。`mx`・`my` は余白の px で、Word から読んだときの細かい値。`base` は本文の既定で `{ fs: pt, lh: CSS の行間, ff: フォント }`）
- Sheet: `{ v: 2, cur, sheets: [{ name, cells: { "A1": "…" }, style, cw, rh, merges, freeze, head, total, filter, cf, dv, notes, charts, grid }] }`（`dv` は入力規則 `[{ ref, type, op, f1, f2, allowBlank, promptTitle, prompt, errTitle, err, style }]`、`notes` はコメント `{ "A1": { text, author, ts } }`。0.1 の `{ cells, fmt, head, total }` は `upgrade` で直す）
- Slide: `{ slides: [{ layout, title, sub, items, items2, notes, bg, trans, objs: [{ id, type, x, y, w, h, rot, … }] }], theme, ratio }`（`type` が `group` のものは `children: [同じ形の図形]` を持ち、子の `x`・`y`・`w`・`h` はグループの中での %。図形の線の太さは `lw`（pt））

### ファイルの場所

| ファイル | 中身 |
|---|---|
| `office/index.html` | 入り口（3 つのアプリを選ぶ・すべてのファイル） |
| `office/doc/`・`sheet/`・`slide/index.html` | 各アプリのページ（中身はほぼ同じ。`data-app`・アプリ名・読み込むファイルだけちがう） |
| `office/shared/office.js` | 画面の枠（スタート画面・リボン・エディタの枠・元に戻す・検索・印刷・アカウント・設定・PWA） |
| `office/shared/doc.js`・`sheet.js`・`slide.js` | 各アプリの中身（リボン・編集・読み込み・書き出し・印刷） |
| `office/shared/vendor/opentype.min.js` | 埋め込みフォントの解析（opentype.js 2.0.0・MIT。ライセンスは `opentype.LICENSE`） |
| `office/shared/formula.js` | 数式（字句・構文・計算・関数・参照のずらし・表示形式） |
| `office/shared/chart.js` | グラフ（SVG。Sheet・Doc・Slide で同じものを使う） |
| `office/shared/zip.js` | ZIP の読み書き（Office のファイル用） |
| `office/shared/store.js` | ファイルの保存（フォルダ・ブラウザの中・Google ドライブ・ダブルクリックで開いたファイル・版の履歴） |
| `office/shared/drive.js` | Google ドライブ（ログイン・Drive API。store.js より先に読み込む） |
| `office/shared/templates.js` | テンプレート |
| `office/shared/office.css` | 見た目（色・スタート画面・リボン・メニュー・各エディタ・スライド ショー・タイトルバー・印刷） |
| `office/doc/`・`sheet/`・`slide/manifest.json` | アプリごとの PWA（インストール・ファイルの関連付け） |
| `office/manifest.json`・`office/sw.js` | 入り口の PWA・オフライン（サービスワーカーは全アプリ共通） |
| `office/icons/` | アイコン（`python _build/office/icons.py` で作り直す） |
| `assets/hub/bridge.js` | sk-hub.jp からログイン中かのヒントを受け取る（office.sk-hub.jp も使う） |
| `_workers/office/` | office.sk-hub.jp の Worker |
| `_build/pages/office.html`・`products/office/` | 紹介ページ（build で作る） |
| `_mockups/office/` | 最初の試作（配信しない） |
| `_tests/office/` | テスト（`run.py`・`run.js`・`fixtures.js`・`check.py`。配信しない） |
| `_tests/office/vendor/` | テストだけの第三者エンジン（`docx`・`docx-preview`・`JSZip` とライセンス。配信しない） |

---

## これからの計画（LibreOffice くらいまで）

### 段階 0: 中身のしくみ — 終わり

- [x] どれを使うか決める → 自作（上の「しくみ」）
- [x] ファイルの `version` を 2 にして、0.1 のファイルを読みこめるようにする

### 段階 1: どのアプリにも要る土台

- [x] 元に戻す・やり直し
- [x] コピー・切り取り・貼り付け（リボンのクリップボード）・書式のコピー
- [x] 検索と置換
- [x] 画像を入れる（ファイルの中に一緒に保存する）
- [x] 印刷と PDF の書き出し（実際に印刷して見た目を確かめる）
- [x] 版の履歴
- [x] キーボードショートカットの一覧
- [x] 書き出し・読み込みの自動テスト（`_tests/office/`）。見つけたずれを直した（.docx の書式の要素の順番・`rFonts` が 2 回・空の段落、.xlsx の `_xlfn.`・`#CIRC!`、.pptx のノート）
- [ ] 書き出したファイルを Microsoft Office・LibreOffice で開いて確かめ、ずれを直す

### 段階 2: Doc（LibreOffice Writer 相当）

- [x] フォント・大きさ・色・蛍光ペン、行の高さ
- [x] 段落の前後の間隔、スタイルを自分で作る・変える
- [x] 用紙のサイズ・向き・余白、ページの区切りの表示
- [x] ヘッダー・フッター・ページ番号、段組み
- [x] 表（結合・塗りつぶし）、画像の配置と文字の回り込み
- [x] 表の罫線・列の幅
- [ ] 図形
- [x] 脚注・目次の自動作成、コメント、スペルチェック
- [ ] 文末脚注、数式（KaTeX）、変更の記録
- [x] 縦書き
- [x] `.docx` の読み込み・書き出し

### 段階 3: Sheet（LibreOffice Calc 相当）

- [x] セルの書式・表示形式・罫線・配置・列の幅・行の高さ・結合
- [x] よく使う関数、オートフィル、相対参照と絶対参照、ほかのシートの参照
- [x] 複数のシート、行と列の固定
- [x] 並べ替え・フィルター、条件付き書式
- [x] 入力規則（リストから選ぶ）
- [x] セルのコメント
- [x] グラフ（自作の `chart.js`）
- [x] 10 万行でも重くならない表示（見えている行だけ作る）
- [ ] ピボットテーブル、変えたセルにかかわるところだけ計算しなおす（いまはブック全体）
- [x] `.xlsx`・`.csv` の読み込み・書き出し
- [x] 条件付き書式の `.xlsx` への書き出し・読み込み
- [x] グラフの `.xlsx` への書き出し（Excel のグラフとして書く。範囲のセルを指すので、Excel でデータを直すとグラフも変わる）
- [x] `.xlsx` のグラフの読み込み（縦棒・横棒・折れ線・面・円・散布図のうち、系列が同じ行の並んだ列で、項目の列がすぐ左にあるものだけ）

### 段階 4: Slide（LibreOffice Impress 相当）

- [x] テキストボックス・図形・画像を自由に置く・大きさを変える・回す
- [x] 整列・重なり順
- [x] グループ化、図形の線の太さ
- [x] テーマ（色のセット）、背景
- [ ] スライドマスター、フォントのセット
- [x] グラフ（Sheet と共通の `chart.js`）
- [x] 表（挿入・セルに直接入力・行と列の追加と削除・見出しの行。.pptx は表のまま書き出し・読み込み）
- [x] スライドの切りかえ、アニメーション
- [x] 発表者ビュー、発表者メモ
- [x] `.pptx` の書き出し・読み込み、PDF

### 段階 5: SK らしさ

- [x] SK Hub Systems アカウントでログイン（名前とアイコンを出すだけ）
- [x] 保存先を切りかえられるようにする（この PC・Google ドライブ。`store.js` が ID で行き先を決める）
- [ ] Google ドライブへの保存を本物で試す（Google Cloud の設定のあと）。オフラインのときに、あとで送る
- [ ] 共有（SK Hub Systems アカウントどうしだけ。相手のアカウントにつないだ Google アカウントに、ドライブの共有をかける）・「共有アイテム」
- [ ] みんなで同時に編集（Yjs。中継は Cloudflare Durable Objects で、中身は残さない。ログインは Worker で Firebase のトークンを確かめる）。Sheet・Slide から
- [ ] Doc の作り直し（同時に編集できる文書のしくみ）→ Doc の同時編集
- [ ] コメントの @メンション・通知・提案モード
- [ ] SK Study・SK's Toolbox とのつながり（例: Sheet の名簿をクラスに配る）
- [ ] `.odt`・`.ods`・`.odp`（LibreOffice の形式）の読み書き
- [ ] Microsoft Store への公開（PWABuilder）。必要なら Tauri で Windows・Mac のインストーラー

### 決めたこと（2026-10-22）

目標は **Google Workspace に近づける**（クラウドに保存・共有・同時に編集）。

- **保存先は、使う人の Google ドライブ**（`drive.file`）。クラウドに保存・共有できるのは、SK Hub Systems アカウントに Google をつないだ人だけ。ほかの人は、いままでどおり PC のフォルダ。あとでほかの保存先も足せるよう、保存先は切りかえられる形にした
- **同時に編集するしくみは Yjs**（MIT ライセンス）。「外の部品は使わない」の例外にする（CRDT を自作すると、それだけで数か月かかり、まちがいも出やすいため）
- **共有は SK Hub Systems アカウントどうしだけ**（リンクを知っていればだれでも、にはしない。学校・子どもの利用を考えて）
- 作る順番: 保存先の切りかえと Google ドライブ → 共有 → Sheet・Slide の同時編集 → Doc の作り直しと同時編集 → コメント・通知

### 決めていないこと

- （決めた）~~インターネット上への保存（オンライン保存）の置き場所~~ → 上の「決めたこと」。そのときの候補:

  | 置き場所 | よいところ | 気になるところ |
  |---|---|---|
  | Firestore（いまのアカウントと同じ） | もう使っている。権限（ルール）もアカウントと同じしくみ | 1 件 1MB まで（画像の入ったファイルは分ける必要がある）。読み書きの回数で料金 |
  | Cloud Storage for Firebase | 大きなファイル・画像をそのまま置ける。Firebase のログインで権限を決められる | 料金がかかる（Blaze プラン）。保存した量と転送量で料金 |
  | Cloudflare R2（office.sk-hub.jp の Worker と同じところ） | 10GB まで無料・転送料なし。Worker からそのまま読み書きできる | ログインの確かめ（Firebase の ID トークン）を Worker で書く必要がある |
  | 使う人の Google ドライブ（Drive API） | SK のサーバーにファイルを置かない。容量は使う人のもの | Google のアカウントだけ。審査（OAuth の確認）がいる。共有のしくみは Google まかせ |

  決めるときに考えること: ファイルの大きさ（画像を入れると数 MB）、使う人の数と料金、共有・同時編集までするか、子ども（学校）の利用でのプライバシー
- Office をアカウントの `SERVICES` に足すか（接続の確認・`accounts/{UID}` に保存するなら、y-filter リポジトリの `systems/firestore.rules` も直す）
- 公開する時期と、紹介ページの「まだ使えません」をいつ外すか
- 多言語の訳
