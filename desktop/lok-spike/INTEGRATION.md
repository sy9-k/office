# Doc の編集画面を LibreOffice（LOKit）に差し替える計画

対象: `shared/doc.js`（2303 行）。Web 版は今のまま残し、デスクトップ版（`desktop/`）だけ差し替える。

## 今の作り（調査の結果）
- 本文は `contenteditable` の `<article id="page">`。文書の中身は HTML（`data.html`）として持つ
- 書式は `document.execCommand`（太字・色・フォント・見出しなど）で変える
- 入力・選択・貼りつけ・ドラッグは、ブラウザの DOM のイベントで処理している（`input`・`keydown`・`paste`・`selectionchange` など）
- 独自の機能: 図表（`figure.chart`、SVG）、脚注（`sup.fn-ref`）、目次（`.toc`）、ページ設定（サイズ・余白・段組・縦書き・ページ番号）、デザインのテーマ（`design:`）、画像のつまみでの大きさ変更、見出しの一覧（ナビ）、ヘッダー・フッター
- 保存・読み込み: `.docx`（`docxFiles()` で XML を組み立てる）、HTML・Markdown・テキストの書き出し

## 差し替えの対応表

| 今の部分 | 扱い | 差し替え先（LOKit） |
|---|---|---|
| 本文の描画（`contenteditable`） | **置きかえ** | タイルの描画（`paintTile` → `/tile`） |
| 文字入力・IME（`keydown`・`composition*`・`input`） | **置きかえ** | `postKeyEvent`・`postWindowExtTextInputEvent`（試作で確認済み） |
| 選択・カーソル（`selectionchange`・`keyup`・`mouseup`） | **置きかえ** | `setTextSelection`・カーソルのコールバック（未確認） |
| 書式コマンド（`exec("bold")` など） | **置きかえ** | `postUnoCommand`（`.uno:Bold` など） |
| 貼りつけ・ドラッグ（`paste`・`drop`） | **置きかえ** | LOKit のクリップボード（未確認） |
| ページ設定（サイズ・余白・縦横） | **置きかえ** | 文書の設定を UNO で変更（未確認） |
| 脚注・目次・見出し | **LibreOffice の機能に任せる** | 脚注・目次の挿入コマンド |
| 図表（独自の `figure.chart`） | **要判断** | LibreOffice のグラフに変えるか、今の SVG を画像として残すか |
| デザインのテーマ（`design:`） | **要判断** | LibreOffice のスタイルに変えるか、対応をやめるか |
| 画像のつまみ・ナビ・ヘッダーとフッター | **作り直し** | 選択のコールバックで位置を取り、今の UI に重ねる |
| メニュー・ツールバー・スタート画面・Hub | **そのまま** | 今の UI。押した命令だけ LOKit に渡す |
| 保存（`.docx`） | **置きかえ** | `saveAs`（`docx` 形式）。拡張子は要判断（前の質問） |
| HTML・Markdown・テキストの書き出し | **要判断** | `saveAs` の各形式、または今の処理を残す |
| 埋め込みフォントの読み込み | **LibreOffice に任せる** | 文書を開くときに LOKit が読む |

## 進め方（段階）
1. **表示だけ**: デスクトップ版で、docx を LOKit で開いて描画する（今の Doc の代わり。編集はしない）
2. **入力**: 文字入力と IME を、表示の画面に送る（試作の範囲）
3. **書式**: ツールバーの太字・斜体・色などを `postUnoCommand` に対応づける
4. **保存**: `saveAs` で docx を書き、今の保存先（PC・Google ドライブ）に渡す
5. **独自機能**: 図表・デザイン・ページ設定の扱いを決めて、順に移す

## 決めてほしいこと
- 拡張子（`.docx` のまま か、`.skdoc` を残すか）
- 図表のグラフ・デザインのテーマを、LibreOffice に合わせて作り直すか、対応をやめるか
- 段階 1 の範囲（表示だけの試作を、デスクトップ版の Doc の画面の中に入れるか、別のウィンドウにするか）
