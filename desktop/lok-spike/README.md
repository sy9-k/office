# LOKit の試作（Doc・案 A の確認）

LibreOffice の中核（LOKit）を Windows で読み込み、docx を開いて描画・入力できるかを確かめる試作です。製品のコードではありません。

## 中身
- `lok_spike.c`: `mergedlo.dll` を読み込み、`libreofficekit_hook_2` で LOKit を起動。docx を開いて 1 ページを BMP に描く。引数を渡すと文字も入れて再描画する
- `include/`: LibreOffice 26.8.1.1 の公式ソースから取った LOKit のヘッダー（`LibreOfficeKit.h`・`LibreOfficeKitInit.h`・`LibreOfficeKitTypes.h`）
- `build.bat`: MSVC（VS 2022 BuildTools）でビルドする
- `dllcheck.c`: DLL を 1 つずつ読み込んで確かめる補助（今回は原因の特定には使わなかった）
- `profile/`: LibreOffice のユーザー設定の置き場（試作用。コミットしない）

## 実行方法
```bash
lok_spike.exe sample.docx out.bmp "C:\Program Files\LibreOffice\program" "D:\GitHub\office\desktop\lok-spike\profile" "41,3042,3044"
```
- 引数のパスは必ず絶対パスと引用符で渡す（空白を含むパスで失敗する）
- 5 番目の引数は、16 進のコードポイントを「,」区切り（`41`＝A、`3042`＝あ）

## 結果（2026-10-28）
- LibreOffice 26.8.1.1 を winget で導入。`mergedlo.dll` に LOKit の入り口がある
- 初期化 約 1.4 秒、docx（`sample.docx`、既存の `doc-rich.docx`）の読み込み 約 2.6〜4.3 秒
- 1 ページの描画 約 100〜200 ms。見出し・色・表・グラフ・脚注まで LibreOffice の描画どおりに出た（画素は BGRA）
- キー入力（コードポイント指定）は反映される。1 文字あたりの再描画は 約 17〜25 ms
- **注意: 描画は非同期。** 入力の直後に描き直すと、最後の文字が出ないことがあった（300 ms 待つと出た）。製品では、固定の待ち時間ではなく LOK のコールバック（状態が変わった通知）で描き直す必要がある
- **日本語 IME（確認済み・シミュレーション）:** `postWindowExtTextInputEvent` で変換中の文字を送れる。
  - 変換中（種類 0）: 「か」→「かん」→「かんじ」と送ると、そのたびに下線付きの変換中の文字が描かれる
  - 候補を選ぶ（種類 0）: 「漢字」を変換中として送る
  - 確定（種類 2）: 同じ「漢字」を付けて送ると、下線が消えて本文の文字になる
  - **確定の前に、確定する文字列を変換中として送る必要がある。** 確定だけを送っても、変換中の「かんじ」が残る
- **未確認:** Tauri の WebView からのキー入力と IME イベントの受け渡し（WebView の compositionstart などを LOK の呼び出しに変換する）、変換候補ウィンドウの位置（カーソル位置の取得）、Windows の IME（TSF）の実機での動き、タイル単位の描画と差分更新、ページ送り・スクロール

## 失敗の原因（メモ）
- `Start-Process` は空白を含む引数を引用符なしで渡すので、パスが分かれてしまう（引用符で囲む）
- `URE_BOOTSTRAP` と `UNO_PATH` を設定する（ini の場所を指す）
- 文書とプロファイルのパスは絶対パスにする（相対パスは `file:///` の URL にできない）
- LOK の描画関数は `LOK_USE_UNSTABLE_API` を定義しないと見えない

## 次にやること
1. 変換中の IME（未確定文字）が LOK に渡せるかを確かめる
2. LOK のコールバックで描き直す仕組みを作る
3. 同じ試作を Tauri の中から呼べるか（ネイティブのプロセスとして動かし、WebView と通信）を確かめる

## WebView での確認（2026-10-28）
- `lok_server.c`: LOKit を読み込み、127.0.0.1 で待ち受ける。`/tile`（BMP）、`/key`、`/ime`、`/size` を返す。LOK の呼び出しは 1 本のスレッドで行う
- `web/index.html`: 描画した画像を表示し、キー入力と IME のイベント（compositionupdate・compositionend）を `/key` と `/ime` に送る
- 起動: `lok_server.exe sample.docx "C:\Program Files\LibreOffice\program" "...\profile" "...\web" 8765`
- 結果（ブラウザのペインで確認）: 英字「Hi」の入力が反映された。IME は DOM のイベントを同じ順番で発生させて確認した（変換中「か」「かん」「かんじ」→ 確定「漢字」）。確定後の描画も正しい
- 未確認: 実機の Windows IME（TSF）の動き、変換候補ウィンドウの位置（カーソル位置を LOKit から取って WebView に伝える必要がある）、Tauri の殻（WebView2）への組み込み。今回のペインは Chromium 系だが、Tauri の WebView2 とは別物

## 書式（太字・斜体）と読み取り専用の原因（2026-10-29 調査）
- **読み取り専用になる原因**: `.~lock.<名前>#` のロックファイルが残っていると、LibreOffice はその文書を読み取り専用（`EditDoc=false`）で開き、文字入力が効かない。テストのたびにサーバーを強制終了していたため、ロックが残った。起動の前に、開くファイルの古いロックを消すように直した。
- **書式の命令（`postUnoCommand`）は効かない**: `.uno:Bold` などを送っても、通知も文書の変化も起きない。全選択（`.uno:SelectAll`）でも通知が出ないことから、命令そのものが実行されていない。
- **Ctrl+キーなら効く**: キーの番号は「Ctrl（0x2000 = 8192）＋ 文字の番号」（`css::awt::Key`。A = 512 から順。B = 513、I = 520、U = 532、Y = 536、Z = 537）。Ctrl+B = 8705、Ctrl+I = 8712、Ctrl+U = 8724、Ctrl+Y = 8728、Ctrl+Z = 8729。パネルの B・I・U・元に戻す・やり直すのボタンは、このキーを送る形に変えた。
- 確かめた結果: 見出しの「abc」を選んで B を押して保存すると、`<w:b w:val="false"/>`（見出しは太字のスタイルなので、太字が外れる）。I を押すと `<w:i/>` が入る。
