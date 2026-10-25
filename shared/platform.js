// SK Hub Office が動いている場所（window.OfficePlatform）
//   ・"web": ブラウザ（PWA を含む）。 ・"desktop": Tauri のデスクトップ版（desktop/）
//   画面と保存のしくみは 1 本のまま、この機能のフラグ（features）で分ける。デスクトップ版にしかできないことは、ここに足していく
//   Web 版では何も足さず、ほかの部品は features を見るだけ（desktop 以外では invoke を呼ばない）
//   Rust 側の命令は desktop/src-tauri/src/lib.rs
(function () {
  "use strict";
  var T = window.__TAURI__;
  var desktop = !!(T && T.core && typeof T.core.invoke === "function");
  var Platform = {
    kind: desktop ? "desktop" : "web",
    isDesktop: desktop,
    // デスクトップ版にだけある機能（Web 版との違い。about の比較表と合わせる）
    features: {
      nativeFs: desktop,      // フォルダ・ファイルをパスで直接読み書きする（毎回の許可なし。Safari・Firefox 相当の環境でも同じ）
      launchFiles: desktop    // 起動したときの引数のファイル（.skdoc などのダブルクリック）を受け取る
    },
    invoke: function (cmd, args) {
      if (!desktop) return Promise.reject(new Error("not-desktop"));
      return T.core.invoke(cmd, args || {});
    }
  };
  window.OfficePlatform = Platform;
  document.documentElement.setAttribute("data-platform", Platform.kind);
})();
