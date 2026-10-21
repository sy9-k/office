# SK Hub Office — いまの作りと、これからの計画

文書の **Doc**・表計算の **Sheet**・スライドの **Slide** をまとめた、SK のオフィスアプリです。開くのは **https://office.sk-hub.jp/**（開発中・検索には載せない）。紹介ページは https://sk-hub.jp/products/office/ 。

目標は「Microsoft Office ほどではないけれど、LibreOffice くらいの機能はそろえる」こと。画面の形は Microsoft Office、見た目は SK Material（Google の Material Design 3 がもと）。

---

## 続きを始めるとき（人も AI も、まずここ）

1. このファイルを読む
2. `office/shared/office.js` のいちばん上の説明（画面の流れ・保存・PWA）
3. `office/shared/store.js` のいちばん上の説明（ファイルの保存）
4. `_workers/office/README.md`（office.sk-hub.jp のしくみ）
5. 作る機能のところ（下の「ファイルの場所」）だけ読む。全部は読まなくてよい

### 手元で動かす

```
python -m http.server 8765
```

リポジトリのいちばん上で動かして、http://localhost:8765/office/ を開く（`.claude/launch.json` の `site`。Mac は `site-mac`）。build はいらない（`office/` は build で作らないページ）。紹介ページ（`_build/pages/office.html`）を直したときだけ `python _build/build.py`。

### 気をつけること

- **ページの中の読み込み・リンクは相対パス**（`../shared/office.js`・`../doc/`）。office.sk-hub.jp では `/`、手元では `/office/` が Office のいちばん上になるため。サイトと共通の部品だけ `/assets/…`・`/toolbox/shared/…` と書く（`_workers/office/src/worker.js` の `SHARED`）。
- **sk-hub.jp/office/ を開くと office.sk-hub.jp に移る**（各ページの `<head>` の最初のスクリプト）。手元（localhost）では移らない。
- **サービスワーカー**（`office/sw.js`）: 画面の部品を大きく変えたら `VERSION` を上げる。sk-hub.jp のものはネットを先に使うので、ふだんはそのままで新しいものが出る。
- **ファイルの中身の形を変えるとき**は、`store.js` の `version`（いまは 1）を上げて、古い形も読めるようにする（使う人の PC に古いファイルが残っている）。
- **まだできないボタン**は `data-todo="名前"` を付けておくと、押したときに「まだ使えません」と出る。作ったら `cmd` に変えて `COMMANDS` に足す。
- 文は「です・ます」で、やさしく短く（SK's Brand の「ことば」）。JavaScript の中の文は `t("…")` で書く（多言語。訳は `_build/i18n/`）。
- この Windows の PC では、PowerShell に `git` がない。GitHub Desktop の中の git（`%LOCALAPPDATA%\GitHubDesktop\app-*\resources\app\git\cmd\git.exe`）は使えるが、GitHub へのプッシュと PR は GitHub Desktop で行う。

---

## いまの状態（バージョン 0.1・2026-10）

### できること

- **起動画面** → **スタート画面**（左のメニュー: ホーム・新規・開く・アカウント・オプション）→ **エディタ**
- **テンプレート**: Doc・Sheet・Slide それぞれ白紙と 5 種類（`office/shared/templates.js`）
- **ファイルの保存**: PC の選んだフォルダに 1 つずつ（`名前.skdoc`・`.sksheet`・`.skslide`。Chrome・Edge のパソコン）。それ以外のブラウザは、ブラウザの中に保存して、書き出し・読み込みで代わりにする。最近使ったアイテム・ピン留め・検索・名前の変更・コピー・削除
- **エディタの画面**: タイトルバー（保存・元に戻す・やり直し・ファイル名）・リボン（タブとグループ）・ステータスバー（ズーム 50〜200%）。自動保存（オプションで切れる）・Ctrl+S
- **Doc**: 太字・斜体・下線・取り消し線・箇条書き・段落番号・インデント・揃え・スタイル（標準・表題・見出し・引用）・表（3×3）・リンク・横線・日付・文字数
- **Sheet**: セルに入力（数式バー・矢印キー・Enter・Delete）・数式（`SUM`・`AVERAGE`・`MAX`・`MIN`・`COUNT` の範囲、`A1*B1` などの 2 項の計算）・オート SUM などのボタン・列の円の表示
- **Slide**: 3 つのレイアウト（表紙・箇条書き・区切り）・文字をその場で書きかえる・スライドの追加・複製・削除・スライド ショー（F5）
- **アカウント**（ログインの状態・製品情報）・**オプション**（ユーザー名・テーマ・動きを減らす・保存先・自動保存・インストール）
- **PWA**: インストール・オフライン・`.skdoc` などをダブルクリックで開く・ウィンドウのいちばん上の帯を Office のタイトルバーにする（Window Controls Overlay）

### まだのこと（いまの作りの弱いところ）

- Doc は `contenteditable` と `document.execCommand`（古いしくみ）。ページの区切り・ヘッダー・画像などが作りにくい
- Sheet は自作の小さな数式だけ。元に戻すもない。行と列は 50 行・12 列ほど（`COLS`）
- Slide は文字だけ。図形・画像を自由に置けない。元に戻すもない
- Office のファイル（`.docx`・`.xlsx`・`.pptx`）の読み書き・PDF・印刷はまだ
- SK Hub Systems アカウントとはまだつながっていない（office.sk-hub.jp では「ログインしていません」と出る）

---

## 作り

### しくみ

- **公開**: ページは sk-hub.jp（GitHub Pages）の `office/` に置き、Cloudflare Worker（`_workers/office/`）が office.sk-hub.jp で中継する。直すときはプッシュするだけ（Worker の deploy は、Worker を直したときだけ）。
- **画面**: 各ページ（`office/doc/index.html` など）は起動画面と入れ物だけ。スタート画面・エディタは `office.js` が組み立てる。`<html data-app="doc|sheet|slide|hub">` でアプリを決める。
- **色**: `office.css` の `:root[data-app="…"]` に M3 の色（`--md-*`）。Doc＝SK's Blue・Sheet＝SK's Matcha・Slide＝SK's Mikan（SK's Toolbox のテーマカラーと同じ値。Toolbox のテーマカラーを変えても変わらない）。
- **部品**: `/toolbox/shared/m3.css`（ボタン・スイッチ・ダイアログなど）と `/toolbox/shared/m3.js`（波紋、`M3.prompt`・`M3.choose`・`M3.confirm`）。文字は M PLUS 1、アイコンは Material Symbols Rounded。

### データ

| もの | 場所 | 形 |
|---|---|---|
| ファイル（フォルダ） | 使う人が選んだ PC のフォルダ | `{ format: "sk-hub-office", version: 1, app, tpl, created, data }` の JSON |
| ファイル（ブラウザの中） | localStorage `sk_office_files` | `[{ id, app, name, tpl, created, updated, data }]` |
| 選んだフォルダ | IndexedDB `sk_office`（`kv` の `folder`） | FileSystemDirectoryHandle |
| オプション | localStorage `sk_office_settings` | `{ theme, motion, autosave, user, pinned }` |

`data` の中身（`templates.js` の説明と同じ）:

- Doc: `{ html }`
- Sheet: `{ cells: { "A1": "…" }, fmt: { C: "yen" }, head: 1, total: 12 }`（`=` で始まると数式）
- Slide: `{ slides: [{ layout: "title|bullets|section", title, sub, items }] }`

### ファイルの場所

| ファイル | 中身 |
|---|---|
| `office/index.html` | 入り口（3 つのアプリを選ぶ・すべてのファイル） |
| `office/doc/`・`sheet/`・`slide/index.html` | 各アプリのページ（中身はほぼ同じ。`data-app` とアプリ名・色だけちがう） |
| `office/shared/office.js` | 画面のすべて（スタート画面・リボン・エディタ・設定・PWA） |
| `office/shared/store.js` | ファイルの保存（フォルダ・ブラウザの中・ダブルクリックで開いたファイル） |
| `office/shared/templates.js` | テンプレート |
| `office/shared/office.css` | 見た目（色・スタート画面・リボン・各エディタ・タイトルバー） |
| `office/manifest.json`・`office/sw.js` | PWA（インストール・オフライン・ファイルの関連付け） |
| `office/icons/` | アイコン（`python _build/office/icons.py` で作り直す） |
| `_workers/office/` | office.sk-hub.jp の Worker |
| `_build/pages/office.html`・`products/office/` | 紹介ページ（build で作る） |
| `_mockups/office/` | 最初の試作（配信しない） |

`office.js` の中の目じるし（`// ====` の見出し）: 保存と設定 → タイトルバー（Window Controls Overlay）→ アプリとして入れる → スタート画面 → エディタ（リボンの `RIBBON`・`COMMANDS`・`RENDER`）→ Doc → Sheet → Slide → クリック・入力・キー → 画面の切りかえ。

---

## これからの計画（LibreOffice くらいまで）

全部を自分で作ると何年もかかるので、**中身のしくみは実績のあるオープンソースの部品を使い、画面と SK らしさは自分で作る**。

### 段階 0: 中身のしくみを決めて入れかえる（いちばん先）

| アプリ | 候補 | ライセンス | Office のファイル |
|---|---|---|---|
| Doc | TipTap / ProseMirror | MIT | 書き出し `docx`（MIT）・読み込み `mammoth`（崩れることがある） |
| Sheet | 表示は自作（見える分だけ描く）＋数式は HyperFormula | GPLv3 か有料（要確認） | ExcelJS（MIT。書式も読み書きできる） |
| Slide | Konva | MIT | 書き出し PptxGenJS（MIT）・読み込みは難しい |
| まとめて | Univer | Apache 2.0（Office ファイルの読み書きは有料版かもしれない。要確認） | — |

- [ ] どれを使うか決める（ライセンスを確かめる）
- [ ] ファイルの `version` を 2 にして、0.1 のファイルを読みこめるようにする

### 段階 1: どのアプリにも要る土台

- [ ] 元に戻す・やり直し（Sheet・Slide にない）
- [ ] コピー・切り取り・貼り付け（リボンのクリップボード）
- [ ] 検索と置換
- [ ] 画像を入れる（ファイルの中に一緒に保存する）
- [ ] 印刷と PDF の書き出し
- [ ] 版の履歴（保存のたびに前の版を残す）
- [ ] キーボードショートカットの一覧

### 段階 2: Doc（LibreOffice Writer 相当）

- [ ] フォント・大きさ・色・蛍光ペン、段落の間隔・行の高さ
- [ ] スタイルを自分で作る・変える
- [ ] 用紙のサイズ・向き・余白、ページの区切りの表示
- [ ] ヘッダー・フッター・ページ番号、段組み
- [ ] 表（結合・罫線・幅）、画像の配置と文字の回り込み、図形
- [ ] 脚注・文末脚注、目次の自動作成、数式（KaTeX）
- [ ] コメント、変更の記録、スペルチェック
- [ ] 縦書き（日本語ならではで、あると強い）
- [ ] `.docx` の読み込み・書き出し

### 段階 3: Sheet（LibreOffice Calc 相当）

- [ ] セルの書式（数値・日付・%・桁区切り）、太字・色・罫線・配置、列の幅・行の高さ、セルの結合
- [ ] よく使う関数（IF・VLOOKUP・XLOOKUP・COUNTIF・SUMIF・日付・文字列など）
- [ ] オートフィル、相対参照と絶対参照（`$A$1`）、ほかのシートの参照
- [ ] 複数のシート、行と列の固定
- [ ] 並べ替え・フィルター、条件付き書式、入力規則（リストから選ぶ）
- [ ] グラフ（棒・折れ線・円・散布図。ECharts か Chart.js）
- [ ] ピボットテーブル、10 万行でも重くならない表示
- [ ] `.xlsx`・`.csv` の読み込み・書き出し

### 段階 4: Slide（LibreOffice Impress 相当）

- [ ] テキストボックス・図形・画像を自由に置く・大きさを変える・回す
- [ ] 整列・重なり順・グループ化
- [ ] テーマ（色とフォントのセット）、スライドマスター、背景
- [ ] 表・グラフ（Sheet と共通にする）
- [ ] スライドの切りかえ、アニメーション
- [ ] 発表者ビュー（メモ・次のスライド・時間）、発表者メモを書けるようにする
- [ ] `.pptx` の書き出し（読み込みはあとで）、PDF

### 段階 5: SK らしさ

- [ ] SK Hub Systems アカウントでの同期（Firebase の承認済みドメインに `office.sk-hub.jp`）
- [ ] みんなで同時に編集（Yjs など）
- [ ] SK Study・SK's Toolbox とのつながり（例: Sheet の名簿をクラスに配る）
- [ ] `.odt`・`.ods`・`.odp`（LibreOffice の形式）の読み書き
- [ ] Microsoft Store への公開（PWABuilder）。必要なら Tauri で Windows・Mac のインストーラー

### 決めていないこと

- 段階 0 の部品（ライセンスを含めて）
- 公開する時期と、紹介ページの「まだ使えません」をいつ外すか
- 多言語の訳（いまは日本語だけ。訳のない文は `_build/i18n/missing/`）
