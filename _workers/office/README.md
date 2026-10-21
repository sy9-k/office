# SK Hub Office の Worker（office.sk-hub.jp）

`_` で始まるフォルダなので、GitHub Pages では公開されません。

## しくみ

office.sk-hub.jp にはページを置いていません。この Worker が、来たアクセスごとに sk-hub.jp（GitHub Pages）の `/office/` を取ってきて返します（admin.sk-hub.jp・support.sk-hub.jp と同じしくみ。`_workers/admin/`・`_workers/support/`）。

| office.sk-hub.jp | 取ってくるところ |
|---|---|
| `/` | `https://sk-hub.jp/office/`（入り口） |
| `/doc/`・`/sheet/`・`/slide/` | `https://sk-hub.jp/office/doc/` など |
| `/shared/…` | `https://sk-hub.jp/office/shared/…`（Office の共通の部分） |
| `/assets/…`・`/toolbox/shared/…`・`/frameworks/…`（`SHARED`） | `https://sk-hub.jp/` の同じパス |
| `/office/…` | `/…` に移す |
| `/robots.txt` | `Disallow: /`（開発中は検索に出さない。ページにも `X-Robots-Tag: noindex`） |

- **Office を直すときは、sk-hub.jp のリポジトリの `office/` にプッシュするだけ**です（この Worker の deploy はいりません）。Cloudflare に 60 秒だけ置いておくので、出るまで最大 1 分ほどかかります。
- Office のページの中の読み込みとリンクは、相対パス（`../shared/office.js`・`../doc/`）で書きます。office.sk-hub.jp でも、手元のサーバー（`localhost:8765/office/`）でも同じように動きます。サイトと共通の部品だけは `/assets/…`・`/toolbox/shared/…` のように書きます（`SHARED`）。
- sk-hub.jp の `/office/` を直接開くと、ページの最初のスクリプトで office.sk-hub.jp の同じページに移ります（ファイルが 2 つの場所に分かれないように）。
- サイトと共通の部品を Office で新しく使うときは、`src/worker.js` の `SHARED` に足して deploy します。
- ファイルは、使う人が選んだ PC のフォルダに保存します（`office/shared/store.js`）。選んだフォルダは、そのブラウザの office.sk-hub.jp の IndexedDB に覚えます。フォルダに保存できないブラウザでは、ブラウザの中（localStorage の `sk_office_files`）に保存します。
- office.sk-hub.jp は sk-hub.jp と別の場所なので、ブラウザの中のデータ（オプションの `sk_office_settings`・選んだフォルダ・SK's Toolbox の設定・SK Hub Systems アカウントのログインの状態）はサイトと分かれます。ライト・ダークは Office のオプションか端末の設定に合わせます。SK Hub Systems アカウントでログインできるようにするときは、Firebase の承認済みドメインに `office.sk-hub.jp` が要ります。

## Worker を直したとき・はじめて公開するとき（deploy）

```
cd _workers/office
npx wrangler deploy
```

wrangler は sakueinakai2023@gmail.com の Cloudflare アカウントにログインして使います（sk-admin・sk-support・sk-status・sk-api と同じ）。初めての deploy で、office.sk-hub.jp の DNS とカスタムドメインが自動で作られます。
