# SK Hub Office — いまの作りと、これからの計画

文書の **Doc**・表計算の **Sheet**・スライドの **Slide** をまとめた、SK のオフィスアプリです。開くのは **https://office.sk-hub.jp/**（開発中・検索には載せない）。紹介ページは https://sk-hub.jp/products/office/ 。

目標は「Microsoft Office ほどではないけれど、LibreOffice くらいの機能はそろえる」こと。画面の形は Microsoft Office、見た目は SK Material（Google の Material Design 3 がもと）。

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

### 気をつけること

- **ページの中の読み込み・リンクは相対パス**（`../shared/office.js`・`../doc/`）。office.sk-hub.jp では `/`、手元では `/office/` が Office のいちばん上になるため。サイトと共通の部品だけ `/assets/…`・`/toolbox/shared/…` と書く（`_workers/office/src/worker.js` の `SHARED`）。
- **sk-hub.jp/office/ を開くと office.sk-hub.jp に移る**（各ページの `<head>` の最初のスクリプト）。手元（localhost）では移らない。
- **サービスワーカー**（`office/sw.js`）: 画面の部品を足したら `SHELL` に入れ、大きく変えたら `VERSION` を上げる（いまは `sk-office-v6`）。sk-hub.jp のものはネットを先に使うので、ふだんはそのままで新しいものが出る。
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
- **PWA**: インストール・オフライン・`.skdoc` などをダブルクリックで開く・ウィンドウのいちばん上の帯を Office のタイトルバーにする（Window Controls Overlay）

### Doc（`doc.js`）

- フォント・大きさ・太字・斜体・下線・取り消し線・上付き・下付き・文字の色・蛍光ペン・書式のクリア・書式のコピー
- 箇条書き・段落番号・インデント・揃え・行間、スタイル（標準・表題・見出し 1・見出し 2・引用）
- 表（マス目で大きさを選ぶ。行・列の挿入と削除・見出しの行・塗りつぶし・右のセルと結合。Tab で次のセル）
- 画像（ファイル・貼り付け・ドラッグ。大きさ・回り込み・代替テキスト）、グラフ（データを貼って作る・ダブルクリックで直す）、リンク、横線、日付、記号、改ページ
- 用紙（A4・B5・A3・Letter、縦・横、余白、段組み）、**縦書き**、ヘッダー・フッター・ページ番号（印刷したとき）、ページの区切りの線
- 目次（見出しから作る・更新）、脚注、コメント（右の欄）、ナビゲーション（見出しの一覧）、スペル チェック（ブラウザの機能）、文字カウント
- 読み込み: `.docx`・`.html`・`.md`・`.txt` / 書き出し: `.docx`・`.html`・`.md`・`.txt`・PDF

### Sheet（`sheet.js`。計算は `formula.js`）

- セルにそのまま入力（日本語の変換もそのまま）・数式バー・名前ボックス（番地で移動）・F2 で編集・Alt+Enter で改行・数式の中で押したセルの番地を入れる
- 関数 100 ほど（SUM・IF・VLOOKUP・XLOOKUP・COUNTIF・SUMIFS・TEXT・DATE など）、`$A$1`、ほかのシートの参照、循環参照の検出
- コピー・切り取り・貼り付け（参照をずらす・値だけ・ほかのアプリとの貼り付け）、オートフィル（数・日付・「第 1 回」の続き）、Ctrl+D・Ctrl+R
- 書式（フォント・大きさ・太字など・文字と塗りの色・罫線・揃え・折り返し・結合）、表示形式（桁区切り・通貨・%・日付・時刻・ユーザー定義・小数の桁）、列の幅（ドラッグ・ダブルクリックで合わせる）・行の高さ
- 行・列の挿入と削除（数式の参照も直す）、複数のシート（追加・名前の変更・複製・移動・削除）、ウィンドウ枠の固定、目盛線
- 並べ替え・フィルター・重複の削除・条件付き書式（大きい・小さい・等しい・文字を含む・重複・データ バー・カラー スケール）
- グラフ（縦棒・横棒・折れ線・面・円・散布図。表の上に置いて動かす・大きさを変える・PNG で保存）
- ステータスバーに平均・個数・合計
- 読み込み: `.xlsx`（値・数式・主な書式・列の幅・結合・固定）・`.csv` / 書き出し: `.xlsx`・`.csv`・PDF

### Slide（`slide.js`）

- レイアウト（タイトル・タイトルとコンテンツ・2 つのコンテンツ・セクション見出し・タイトルのみ・白紙）の文字をそのまま書きかえる
- テキスト ボックス・図形（9 種類）・画像・グラフを自由に置く。ドラッグで動かす・角で大きさを変える・上の丸で回す・ほかのものや真ん中に合わせる線・重なり順・整列・複製・コピーと貼り付け・矢印で少しずつ動かす
- テーマ（みかん・海・抹茶・さくら・モノクロ・夜）、背景の色、スライドのサイズ（16:9・4:3）
- 画面切り替え（フェード・プッシュ・ズーム）、アニメーション（フェード・スライドイン・ズーム。順に出す）
- スライドの並べかえ（ドラッグ）・複製・削除・右クリックのメニュー、ノート
- スライド ショー（F5）と**発表者ビュー**（P。いまのスライド・次のスライド・ノート・時間）、B で画面を黒く
- 読み込み: `.pptx`（文字・図形・画像・ノート）・`.txt`（アウトライン） / 書き出し: `.pptx`・アウトライン（`.txt`）・PDF

### SK Hub Systems アカウント（ログインだけ）

- Office の「アカウント」から **ログイン** を押すと、sk-hub.jp の `/account/` でログインし、終わったら Office に戻る
- Office が使うのは **ログイン中かのヒント**（`skhub_account`: 名前・メールアドレス・写真）だけ。アイコンと名前が出る。ファイルはいままでどおり PC に保存する
- office.sk-hub.jp からは sk-hub.jp の保存場所が見えないので、sk-hub.jp の `/assets/hub/bridge.html`（見えない iframe）から受け取る（`assets/hub/bridge.js` の `ORIGINS` に `https://office.sk-hub.jp` を足した）。手元・sk-hub.jp では、そのまま読む
- ログアウト・名前の変更は、アカウントのページ（新しいタブ）で行う。Office に戻ってくると、ログインの状態を読みなおす
- まだ Firebase は Office では読み込まない（Firestore に何も保存しない）。`SERVICES` に `office` も足していない

### まだのこと（弱いところ）

- **office.sk-hub.jp でのログインの受け取りは、公開してから確かめる**（手元では同じ場所なので、bridge を通らない）
- **Microsoft Office・LibreOffice で、書き出したファイルを開いて確かめていない**（XML とつながりは確かめた。.pptx は python-pptx で開けた）
- 印刷は、印刷用の形を作るところまで確かめた。実際の印刷・PDF の見た目はまだ確かめていない
- Doc は `contenteditable` と `document.execCommand`（古いしくみ）。変更の記録（校閲）・数式（KaTeX）・図形はない。.docx の細かい書式（段落の間隔・表の幅など）は変わることがある
- Sheet は、作った行を全部画面に置く（下へスクロールすると足す）。何万行もあると重い。入力規則（リストから選ぶ）・ピボットテーブル・テーブル・セルのコメント・グラフの .xlsx への書き出しはない
- Slide のグループ化・表・スライドマスター・図形の線の太さ・.pptx のノートとアニメーションの書き出しはない
- `.odt`・`.ods`・`.odp`（LibreOffice の形式）
- 多言語の訳（新しく足した文は日本語だけ。訳のない文は `_build/i18n/missing/`）

---

## 作り

### しくみ

- **公開**: ページは sk-hub.jp（GitHub Pages）の `office/` に置き、Cloudflare Worker（`_workers/office/`）が office.sk-hub.jp で中継する。直すときはプッシュするだけ（Worker の deploy は、Worker を直したときだけ）。
- **画面**: 各ページ（`office/doc/index.html` など）は起動画面と入れ物だけ。スタート画面・エディタの枠は `office.js` が組み立て、エディタの中身はアプリのファイル（`doc.js` など）が作る。`<html data-app="doc|sheet|slide|hub">` でアプリを決める。
- **アプリのファイル**: `window.OfficeApps[アプリ] = function (O) { return { ribbon, render, command, state, … } }`。`office.js` より先に読み込む。返すもの（`upgrade`・`thumb`・`find`・`replace`・`clip`・`key`・`beforePrint`・`exports`・`exportAs`・`imports`・`importFile` など）は `office.js` の中の使い方を見る。
- **中身は自作**（決めた）: 外の部品（TipTap・HyperFormula・Konva など）は使わない。オフラインで動き、ライセンスを気にしなくてよく、SK らしい画面にしやすいため。Office のファイルの読み書きも自作（`zip.js` は ZIP、XML はブラウザの DOMParser）。
- **色**: `office.css` の `:root[data-app="…"]` に M3 の色（`--md-*`）。Doc＝SK's Blue・Sheet＝SK's Matcha・Slide＝SK's Mikan。スライドのテーマは `.sl[data-theme]` の `--sl-*`（`slide.js` の `THEMES` と同じ値）。
- **部品**: `/toolbox/shared/m3.css`（ボタン・スイッチ・ダイアログなど）と `/toolbox/shared/m3.js`（波紋、`M3.prompt`・`M3.choose`・`M3.confirm`・`M3.dialog`）。文字は M PLUS 1、アイコンは Material Symbols Rounded。

### データ

| もの | 場所 | 形 |
|---|---|---|
| ファイル（フォルダ） | 使う人が選んだ PC のフォルダ | `{ format: "sk-hub-office", version: 2, app, tpl, created, data }` の JSON（version 1 も読める） |
| ファイル（ブラウザの中） | localStorage `sk_office_files` | `[{ id, app, name, tpl, created, updated, data }]` |
| 選んだフォルダ・版の履歴 | IndexedDB `sk_office`（`kv` の `folder`・`v:ファイルの ID`） | FileSystemDirectoryHandle・`[{ ts, size, data }]` |
| オプション | localStorage `sk_office_settings` | `{ theme, motion, autosave, user, pinned }` |
| ログイン中かのヒント | localStorage `skhub_account`（sk-hub.jp のものを写す） | `{ name, email, photo, letter, color, … }` |

`data` の中身（くわしくは各アプリのファイルのいちばん上）:

- Doc: `{ html, page: { size, orient, margin, cols, vert, header, footer, num }, comments: [{ id, text, author, ts }] }`
- Sheet: `{ v: 2, cur, sheets: [{ name, cells: { "A1": "…" }, style, cw, rh, merges, freeze, head, total, filter, cf, charts, grid }] }`（0.1 の `{ cells, fmt, head, total }` は `upgrade` で直す）
- Slide: `{ slides: [{ layout, title, sub, items, items2, notes, bg, trans, objs: [{ id, type, x, y, w, h, rot, … }] }], theme, ratio }`

### ファイルの場所

| ファイル | 中身 |
|---|---|
| `office/index.html` | 入り口（3 つのアプリを選ぶ・すべてのファイル） |
| `office/doc/`・`sheet/`・`slide/index.html` | 各アプリのページ（中身はほぼ同じ。`data-app`・アプリ名・読み込むファイルだけちがう） |
| `office/shared/office.js` | 画面の枠（スタート画面・リボン・エディタの枠・元に戻す・検索・印刷・アカウント・設定・PWA） |
| `office/shared/doc.js`・`sheet.js`・`slide.js` | 各アプリの中身（リボン・編集・読み込み・書き出し・印刷） |
| `office/shared/formula.js` | 数式（字句・構文・計算・関数・参照のずらし・表示形式） |
| `office/shared/chart.js` | グラフ（SVG。Sheet・Doc・Slide で同じものを使う） |
| `office/shared/zip.js` | ZIP の読み書き（Office のファイル用） |
| `office/shared/store.js` | ファイルの保存（フォルダ・ブラウザの中・ダブルクリックで開いたファイル・版の履歴） |
| `office/shared/templates.js` | テンプレート |
| `office/shared/office.css` | 見た目（色・スタート画面・リボン・メニュー・各エディタ・スライド ショー・タイトルバー・印刷） |
| `office/manifest.json`・`office/sw.js` | PWA（インストール・オフライン・ファイルの関連付け） |
| `office/icons/` | アイコン（`python _build/office/icons.py` で作り直す） |
| `assets/hub/bridge.js` | sk-hub.jp からログイン中かのヒントを受け取る（office.sk-hub.jp も使う） |
| `_workers/office/` | office.sk-hub.jp の Worker |
| `_build/pages/office.html`・`products/office/` | 紹介ページ（build で作る） |
| `_mockups/office/` | 最初の試作（配信しない） |

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
- [ ] 書き出したファイルを Microsoft Office・LibreOffice で開いて確かめ、ずれを直す

### 段階 2: Doc（LibreOffice Writer 相当）

- [x] フォント・大きさ・色・蛍光ペン、行の高さ
- [ ] 段落の前後の間隔、スタイルを自分で作る・変える
- [x] 用紙のサイズ・向き・余白、ページの区切りの表示
- [x] ヘッダー・フッター・ページ番号、段組み
- [x] 表（結合・塗りつぶし）、画像の配置と文字の回り込み
- [ ] 表の罫線・列の幅、図形
- [x] 脚注・目次の自動作成、コメント、スペルチェック
- [ ] 文末脚注、数式（KaTeX）、変更の記録
- [x] 縦書き
- [x] `.docx` の読み込み・書き出し

### 段階 3: Sheet（LibreOffice Calc 相当）

- [x] セルの書式・表示形式・罫線・配置・列の幅・行の高さ・結合
- [x] よく使う関数、オートフィル、相対参照と絶対参照、ほかのシートの参照
- [x] 複数のシート、行と列の固定
- [x] 並べ替え・フィルター、条件付き書式
- [ ] 入力規則（リストから選ぶ）
- [x] グラフ（自作の `chart.js`）
- [ ] ピボットテーブル、10 万行でも重くならない表示（見えている行だけ作る）
- [x] `.xlsx`・`.csv` の読み込み・書き出し
- [ ] グラフ・条件付き書式の `.xlsx` への書き出し

### 段階 4: Slide（LibreOffice Impress 相当）

- [x] テキストボックス・図形・画像を自由に置く・大きさを変える・回す
- [x] 整列・重なり順
- [ ] グループ化
- [x] テーマ（色のセット）、背景
- [ ] スライドマスター、フォントのセット
- [x] グラフ（Sheet と共通の `chart.js`）
- [ ] 表
- [x] スライドの切りかえ、アニメーション
- [x] 発表者ビュー、発表者メモ
- [x] `.pptx` の書き出し・読み込み、PDF

### 段階 5: SK らしさ

- [x] SK Hub Systems アカウントでログイン（名前とアイコンを出すだけ）
- [ ] インターネット上への保存・同期（下の「決めていないこと」）
- [ ] みんなで同時に編集（Yjs など）
- [ ] SK Study・SK's Toolbox とのつながり（例: Sheet の名簿をクラスに配る）
- [ ] `.odt`・`.ods`・`.odp`（LibreOffice の形式）の読み書き
- [ ] Microsoft Store への公開（PWABuilder）。必要なら Tauri で Windows・Mac のインストーラー

### 決めていないこと

- **インターネット上への保存（オンライン保存）の置き場所**。いまは PC のフォルダだけ。候補:

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
