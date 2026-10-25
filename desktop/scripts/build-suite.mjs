// 3 つのアプリ（Doc・Sheet・Slide）を作って、1 つのインストーラー（dist-installer/SK-Hub-Office-Setup-<版>.exe）にまとめる
//   アプリごとに  tauri build --no-bundle --config apps/<名前>.json  で実行ファイルを作り、stage/ に Doc.exe・Sheet.exe・Slide.exe として集める
//   インストーラーは installer/suite.nsi を、Tauri が取ってきた makensis（%LOCALAPPDATA%\tauri\NSIS）で作る
//   使い方: npm run build:suite   （--skip-build で、前に作った stage/ だけをまとめなおす）
import { spawnSync } from "node:child_process";
import { copyFileSync, cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const desktop = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriDir = path.join(desktop, "src-tauri");
const stage = path.join(desktop, "stage");
const outDir = path.join(desktop, "dist-installer");
const APPS = [["doc", "Doc"], ["sheet", "Sheet"], ["slide", "Slide"]];
const version = JSON.parse(readFileSync(path.join(tauriDir, "tauri.conf.json"), "utf8")).version;
// 版は tauri.conf.json が元（インストーラー・自動更新・アプリの表示）。Cargo.toml と package.json がちがっていたら、作らない
const cargoVersion = /^version\s*=\s*"([^"]+)"/m.exec(readFileSync(path.join(tauriDir, "Cargo.toml"), "utf8"))?.[1];
const npmVersion = JSON.parse(readFileSync(path.join(desktop, "package.json"), "utf8")).version;
if (cargoVersion !== version || npmVersion !== version) {
  console.error(`版がそろっていません: tauri.conf.json=${version} Cargo.toml=${cargoVersion} package.json=${npmVersion}`);
  process.exit(1);
}
const run = (cmd, args, opts = {}) => {
  const r = spawnSync(cmd, args, { stdio: "inherit", shell: process.platform === "win32", cwd: desktop, ...opts });
  if (r.status !== 0) { console.error(`失敗: ${cmd} ${args.join(" ")}`); process.exit(r.status || 1); }
};

mkdirSync(stage, { recursive: true });
if (!process.argv.includes("--skip-build")) {
  run("node", ["scripts/prepare-web.mjs"]);
  for (const [key, name] of APPS) {
    run("npx", ["tauri", "build", "--no-bundle", "--config", `apps/${key}.json`]);
    copyFileSync(path.join(tauriDir, "target", "release", "sk-hub-office.exe"), path.join(stage, `${name}.exe`));
  }
}
// ファイルの種類のアイコン（installer/icons/*-file.ico。icons/src/sk-*-file.png から tauri icon で作ったもの）
cpSync(path.join(desktop, "installer", "icons"), path.join(stage, "icons"), { recursive: true });
copyFileSync(path.join(tauriDir, "icons", "icon.ico"), path.join(stage, "icons", "suite.ico")); // 「アプリと機能」に出す SK Hub Office のアイコン
for (const [, name] of APPS) if (!existsSync(path.join(stage, `${name}.exe`))) { console.error(`stage/${name}.exe がありません（--skip-build を外して作ってください）`); process.exit(1); }

const makensis = path.join(process.env.LOCALAPPDATA || "", "tauri", "NSIS", "makensis.exe");
if (!existsSync(makensis)) { console.error("makensis が見つかりません。一度 npm run build を実行すると、Tauri が取ってきます: " + makensis); process.exit(1); }

// NSIS は UTF-8 の BOM がないと日本語を読めないので、BOM つきの写しを作ってから渡す
const src = readFileSync(path.join(desktop, "installer", "suite.nsi"), "utf8").replace(/^﻿/, "");
const nsi = path.join(stage, "suite.nsi");
writeFileSync(nsi, "﻿" + src, "utf8");
mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, `SK-Hub-Office-Setup-${version}.exe`);
// 前の版のインストーラーが開いている（実行中）こともあるので、フォルダごとは消さず、この版のものだけを作りなおす
for (const f of [outFile, outFile + ".sig", path.join(outDir, "latest.json")]) rmSync(f, { force: true });
run(makensis, ["/V2", `/DVERSION=${version}`, `/DSTAGE=${stage}`, `/DOUTFILE=${outFile}`, `/DICON=${path.join(tauriDir, "icons", "icon.ico")}`, nsi]);
console.log("できました: " + outFile);
