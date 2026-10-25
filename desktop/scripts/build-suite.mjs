// 3 つのアプリ（Doc・Sheet・Slide）を作って、1 つのインストーラー（dist-installer/SK-Hub-Office-Setup-<版>.exe）にまとめる
//   アプリごとに  tauri build --no-bundle --config apps/<名前>.json  で実行ファイルを作り、stage/ に Doc.exe・Sheet.exe・Slide.exe として集める
//   インストーラーは installer/suite.nsi を、Tauri が取ってきた makensis（%LOCALAPPDATA%\tauri\NSIS）で作る
//   使い方: npm run build:suite   （--skip-build で、前に作った stage/ だけをまとめなおす）
import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const desktop = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriDir = path.join(desktop, "src-tauri");
const stage = path.join(desktop, "stage");
const outDir = path.join(desktop, "dist-installer");
const APPS = [["doc", "Doc"], ["sheet", "Sheet"], ["slide", "Slide"]];
const version = JSON.parse(readFileSync(path.join(tauriDir, "tauri.conf.json"), "utf8")).version;
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
for (const [, name] of APPS) if (!existsSync(path.join(stage, `${name}.exe`))) { console.error(`stage/${name}.exe がありません（--skip-build を外して作ってください）`); process.exit(1); }

const makensis = path.join(process.env.LOCALAPPDATA || "", "tauri", "NSIS", "makensis.exe");
if (!existsSync(makensis)) { console.error("makensis が見つかりません。一度 npm run build を実行すると、Tauri が取ってきます: " + makensis); process.exit(1); }

// NSIS は UTF-8 の BOM がないと日本語を読めないので、BOM つきの写しを作ってから渡す
const src = readFileSync(path.join(desktop, "installer", "suite.nsi"), "utf8").replace(/^﻿/, "");
const nsi = path.join(stage, "suite.nsi");
writeFileSync(nsi, "﻿" + src, "utf8");
rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, `SK-Hub-Office-Setup-${version}.exe`);
run(makensis, ["/V2", `/DVERSION=${version}`, `/DSTAGE=${stage}`, `/DOUTFILE=${outFile}`, `/DICON=${path.join(tauriDir, "icons", "icon.ico")}`, nsi]);
console.log("できました: " + outFile);
