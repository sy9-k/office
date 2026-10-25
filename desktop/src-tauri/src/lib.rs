// SK Hub Office デスクトップ版の入り口。画面は Web 版（desktop/dist）で、ここは OS とのつなぎだけを持つ
// shared/platform.js の OfficePlatform.invoke("命令名", {...}) から呼ばれる
//
// ファイルの命令は、SK Hub Office のファイル（.skdoc・.sksheet・.skslide）だけを読み書きする
//   ・名前に区切り文字（/ \ : など）や .. を含められない（フォルダの外に出られない）
//   ・書き込みは、同じフォルダの一時ファイルに書いてから置きかえる（途中で止まっても、前のファイルが壊れない）
use serde::Serialize;
use std::path::{Path, PathBuf};
use tauri::Manager;
use tauri_plugin_dialog::DialogExt;

const EXTS: [&str; 3] = [".skdoc", ".sksheet", ".skslide"];

fn has_ext(name: &str) -> bool {
    let lower = name.to_lowercase();
    EXTS.iter().any(|e| lower.ends_with(e))
}

fn safe_name(name: &str) -> Result<&str, String> {
    let bad = name.is_empty()
        || name.contains(['/', '\\', ':', '*', '?', '"', '<', '>', '|', '\0'])
        || name.contains("..")
        || name.starts_with('.');
    if bad || !has_ext(name) {
        return Err(format!("bad-name: {name}"));
    }
    Ok(name)
}

fn in_dir(dir: &str, name: &str) -> Result<PathBuf, String> {
    let dir = Path::new(dir);
    if !dir.is_absolute() {
        return Err("bad-dir".into());
    }
    Ok(dir.join(safe_name(name)?))
}

fn abs_office_path(path: &str) -> Result<PathBuf, String> {
    let p = Path::new(path);
    let name = p.file_name().and_then(|n| n.to_str()).ok_or("bad-path")?;
    if !p.is_absolute() || !has_ext(name) {
        return Err("bad-path".into());
    }
    Ok(p.to_path_buf())
}

fn modified_ms(meta: &std::fs::Metadata) -> u64 {
    meta.modified()
        .ok()
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

fn write_atomic(path: &Path, text: &str) -> Result<(), String> {
    let name = path.file_name().and_then(|n| n.to_str()).ok_or("bad-path")?;
    let tmp = path.with_file_name(format!(".{name}.tmp"));
    std::fs::write(&tmp, text).map_err(|e| e.to_string())?;
    std::fs::rename(&tmp, path).map_err(|e| {
        let _ = std::fs::remove_file(&tmp);
        e.to_string()
    })
}

#[derive(Serialize)]
struct Entry {
    name: String,
    size: u64,
    modified: u64,
}

#[derive(Serialize)]
struct Content {
    text: String,
    modified: u64,
}

fn read_file(path: &Path) -> Result<Content, String> {
    let meta = std::fs::metadata(path).map_err(|e| e.to_string())?;
    let text = std::fs::read_to_string(path).map_err(|e| e.to_string())?;
    Ok(Content { text, modified: modified_ms(&meta) })
}

#[tauri::command]
fn app_info() -> serde_json::Value {
    serde_json::json!({
        "kind": "desktop",
        "version": env!("CARGO_PKG_VERSION"),
        "os": std::env::consts::OS,
    })
}

// 最初に使う保存先（ドキュメント\SK Hub Office。なければ作る）
#[tauri::command]
async fn default_dir(app: tauri::AppHandle) -> Result<String, String> {
    let dir = app.path().document_dir().map_err(|e| e.to_string())?.join("SK Hub Office");
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.to_string_lossy().into_owned())
}

// 保存先のフォルダを選ぶ（やめたら null）
#[tauri::command]
async fn pick_folder(app: tauri::AppHandle, start: Option<String>) -> Result<Option<String>, String> {
    let mut dialog = app.dialog().file().set_title("SK Hub Office の保存先フォルダ");
    if let Some(s) = start.filter(|s| Path::new(s).is_dir()) {
        dialog = dialog.set_directory(s);
    }
    match dialog.blocking_pick_folder() {
        Some(p) => Ok(Some(p.into_path().map_err(|e| e.to_string())?.to_string_lossy().into_owned())),
        None => Ok(None),
    }
}

#[tauri::command]
async fn dir_exists(dir: String) -> bool {
    Path::new(&dir).is_dir()
}

#[tauri::command]
async fn fs_list(dir: String) -> Result<Vec<Entry>, String> {
    let mut out = Vec::new();
    for item in std::fs::read_dir(&dir).map_err(|e| e.to_string())? {
        let item = item.map_err(|e| e.to_string())?;
        let name = item.file_name().to_string_lossy().into_owned();
        if name.starts_with('.') || !has_ext(&name) {
            continue;
        }
        if let Ok(meta) = item.metadata() {
            if meta.is_file() {
                out.push(Entry { name, size: meta.len(), modified: modified_ms(&meta) });
            }
        }
    }
    Ok(out)
}

#[tauri::command]
async fn fs_exists(dir: String, name: String) -> Result<bool, String> {
    Ok(in_dir(&dir, &name)?.is_file())
}

#[tauri::command]
async fn fs_read(dir: String, name: String) -> Result<Content, String> {
    read_file(&in_dir(&dir, &name)?)
}

#[tauri::command]
async fn fs_write(dir: String, name: String, text: String) -> Result<(), String> {
    write_atomic(&in_dir(&dir, &name)?, &text)
}

#[tauri::command]
async fn fs_remove(dir: String, name: String) -> Result<(), String> {
    std::fs::remove_file(in_dir(&dir, &name)?).map_err(|e| e.to_string())
}

// フォルダの外のファイル（ダブルクリックで開いたもの）を、パスで直接読み書きする
#[tauri::command]
async fn fs_read_path(path: String) -> Result<Content, String> {
    read_file(&abs_office_path(&path)?)
}

#[tauri::command]
async fn fs_write_path(path: String, text: String) -> Result<(), String> {
    write_atomic(&abs_office_path(&path)?, &text)
}

// 起動したときの引数にある、SK Hub Office のファイル（エクスプローラーでダブルクリックして開いたもの）
#[tauri::command]
fn launch_files() -> Vec<String> {
    std::env::args()
        .skip(1)
        .filter(|a| {
            let p = Path::new(a);
            p.is_absolute() && p.is_file() && p.file_name().and_then(|n| n.to_str()).map(has_ext).unwrap_or(false)
        })
        .collect()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            app_info,
            default_dir,
            pick_folder,
            dir_exists,
            fs_list,
            fs_exists,
            fs_read,
            fs_write,
            fs_remove,
            fs_read_path,
            fs_write_path,
            launch_files,
        ])
        .run(tauri::generate_context!())
        .expect("SK Hub Office の起動に失敗しました");
}
