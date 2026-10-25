// Web 版（リポジトリの一番上）から、配信に要るものだけを desktop/dist にコピーする（Tauri はここを同梱する）
// Web 版のソースは 1 本のまま。デスクトップ専用の画面はコピーせず、shared/platform.js の機能フラグで分ける
import { cp, rm, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const dist = path.resolve(root, "desktop", "dist");

const ITEMS = ["index.html", "manifest.json", "about", "doc", "sheet", "slide", "shared", "icons"];
const SKIP = [/[\\/]icons[\\/]src([\\/]|$)/]; // 元絵（大きい PNG）は同梱しない

await rm(dist, { recursive: true, force: true });
await mkdir(dist, { recursive: true });
for (const item of ITEMS) {
  await cp(path.join(root, item), path.join(dist, item), {
    recursive: true,
    filter: (src) => !SKIP.some((re) => re.test(src)),
  });
}
console.log("desktop/dist ← " + ITEMS.join(", "));
