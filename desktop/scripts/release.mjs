// 配布の準備: 作ったインストーラー（dist-installer/）に署名して、自動更新が読む latest.json を作る
//   使い方: npm run release -- [--base <インストーラーを置く URL>] [--notes "更新の内容"]
//   署名の鍵: %USERPROFILE%\.tauri\sk-hub-office.key（tauri signer generate で作った。リポジトリには入れない。なくすと、更新ができなくなる）
//     別の場所なら、環境変数 TAURI_SIGNING_PRIVATE_KEY_PATH（パスワードがあれば TAURI_SIGNING_PRIVATE_KEY_PASSWORD）
//   できるもの（dist-installer/）:
//     SK-Hub-Office-Setup-<版>.exe       … インストーラー（初めて入れる人にも、この 1 つを渡す）
//     SK-Hub-Office-Setup-<版>.exe.sig   … 署名
//     latest.json                        … 自動更新が読む（tauri.conf.json の plugins.updater.endpoints の場所に置く。既定は https://office.sk-hub.jp/desktop/latest.json）
//   置き方（どちらも office.sk-hub.jp に公開）: latest.json → /desktop/latest.json ・ インストーラー → <base>/（既定 https://office.sk-hub.jp/download/）
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import os from "node:os";
import path from "node:path";

const desktop = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(desktop, "dist-installer");
const arg = (name, fallback) => { const i = process.argv.indexOf(name); return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback; };
const version = JSON.parse(readFileSync(path.join(desktop, "src-tauri", "tauri.conf.json"), "utf8")).version;
// インストーラーは GitHub の Releases に置く（git の履歴に大きなファイルを入れないため）。タグは v<版>
const base = arg("--base", `https://github.com/sy9-k/office/releases/download/v${version}/`).replace(/\/?$/, "/");
// 更新の内容は、RELEASE_NOTES.md（または --notes-file）から読む。コマンドの引数の日本語は、環境によって文字化けするので使わない
const notesFile = arg("--notes-file", path.join(desktop, "RELEASE_NOTES.md"));
const notes = existsSync(notesFile) ? readFileSync(notesFile, "utf8").replace(/^﻿/, "").trim() : "";
const name = `SK-Hub-Office-Setup-${version}.exe`;
const installer = path.join(dist, name);
if (!existsSync(installer)) { console.error(`${installer} がありません。先に npm run build:suite を実行してください`); process.exit(1); }

const keyPath = process.env.TAURI_SIGNING_PRIVATE_KEY_PATH || path.join(os.homedir(), ".tauri", "sk-hub-office.key");
if (!existsSync(keyPath)) { console.error(`署名の鍵がありません: ${keyPath}`); process.exit(1); }
const r = spawnSync("npx", ["tauri", "signer", "sign", "-f", keyPath, ...(process.env.TAURI_SIGNING_PRIVATE_KEY_PASSWORD ? ["-p", process.env.TAURI_SIGNING_PRIVATE_KEY_PASSWORD] : []), installer], { stdio: "inherit", shell: process.platform === "win32", cwd: desktop });
if (r.status !== 0) { console.error("署名に失敗しました"); process.exit(r.status || 1); }
const sig = readFileSync(installer + ".sig", "utf8").trim();

const latest = {
  version,
  notes,
  pub_date: new Date().toISOString(),
  platforms: { "windows-x86_64": { signature: sig, url: base + name } },
};
const json = JSON.stringify(latest, null, 2) + "\n";
writeFileSync(path.join(dist, "latest.json"), json, "utf8");
// 自動更新が読む場所（office.sk-hub.jp/desktop/latest.json）は、このリポジトリの desktop/latest.json（GitHub Pages）
writeFileSync(path.join(desktop, "latest.json"), json, "utf8");
console.log(readdirSync(dist).map((f) => "  " + f).join("\n"));
console.log(`\n公開の手順（この順に）:
  1. GitHub の Releases で、タグ v${version} のリリースを作り、${name} を添付する（${base}${name} で取れるようにする）
  2. desktop/latest.json（いま更新した）をコミットして、プッシュする → office.sk-hub.jp/desktop/latest.json が新しくなり、起動中のアプリが更新に気づく
  （1 より先に 2 をすると、更新のダウンロードが失敗します）`);
