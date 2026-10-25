// SK Hub Office デスクトップ版の入り口。画面は Web 版（desktop/dist）で、ここは OS とのつなぎだけを持つ
// ファイル・フォント・印刷などのコマンドは、機能ごとにここへ足していく（shared/platform.js から呼ぶ）

#[tauri::command]
fn app_info() -> serde_json::Value {
    serde_json::json!({
        "kind": "desktop",
        "version": env!("CARGO_PKG_VERSION"),
        "os": std::env::consts::OS,
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![app_info])
        .run(tauri::generate_context!())
        .expect("SK Hub Office の起動に失敗しました");
}
