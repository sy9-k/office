// SK Hub Office デスクトップ版の入り口。画面は Web 版（desktop/dist）で、ここは OS とのつなぎだけを持つ
// shared/platform.js の OfficePlatform.invoke("命令名", {...}) から呼ばれる
//
// ファイルの命令は、SK Hub Office のファイル（.skdoc・.sksheet・.skslide）だけを読み書きする
//   ・名前に区切り文字（/ \ : など）や .. を含められない（フォルダの外に出られない）
//   ・書き込みは、同じフォルダの一時ファイルに書いてから置きかえる（途中で止まっても、前のファイルが壊れない）
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
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

// ---------- ウィンドウ ----------
// アプリ（Doc・Sheet・Slide・入り口）ごとに別のウィンドウで開く。起動の引数のファイル（ダブルクリック）も、ここで開く

static WINDOW_SEQ: AtomicUsize = AtomicUsize::new(0);

// このアプリが Doc・Sheet・Slide のどれか。識別子の最後（jp.sk-hub.office.doc の doc）で決まる（アプリごとの設定は desktop/apps/）。
// どれでもない（開発用のすべて入り）なら "hub"
fn app_kind(app: &tauri::AppHandle) -> &'static str {
    match app.config().identifier.rsplit('.').next() {
        Some("doc") => "doc",
        Some("sheet") => "sheet",
        Some("slide") => "slide",
        _ => "hub",
    }
}

fn home_page(kind: &str) -> &'static str {
    match kind {
        "doc" => "doc/index.html",
        "sheet" => "sheet/index.html",
        "slide" => "slide/index.html",
        _ => "index.html",
    }
}

fn app_title(kind: &str) -> &'static str {
    match kind {
        "doc" => "SK Hub Office Doc",
        "sheet" => "SK Hub Office Sheet",
        "slide" => "SK Hub Office Slide",
        _ => "SK Hub Office",
    }
}

// ---------- アプリのあいだで共有する設定 ----------
// 識別子がちがうと、画面（WebView）の保存場所（localStorage）がアプリごとに分かれる。ログイン・Google ドライブのつなぎ・保存先のフォルダ・設定は、
// どのアプリでも同じにしたいので、決まったキーだけを共有のファイル（%APPDATA%\SK Hub Office\shared.json）に写す
//   画面側（shared/platform.js）が、ウィンドウを開くときにここから localStorage に戻し、変えたときにここへ書く
const SHARED_KEYS: [&str; 5] = ["skhub_account", "sk_office_drive", "sk_office_place", "sk_office_native_dir", "sk_office_settings"];
static SHARED_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn shared_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().data_dir().map_err(|e| e.to_string())?.join("SK Hub Office");
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.join("shared.json"))
}

fn read_shared(app: &tauri::AppHandle) -> std::collections::BTreeMap<String, String> {
    shared_path(app)
        .ok()
        .and_then(|p| std::fs::read_to_string(p).ok())
        .and_then(|t| serde_json::from_str(&t).ok())
        .unwrap_or_default()
}

#[tauri::command]
async fn shared_all(app: tauri::AppHandle) -> std::collections::BTreeMap<String, String> {
    read_shared(&app)
}

// value が null なら消す。決まったキーだけ
#[tauri::command]
async fn shared_set(app: tauri::AppHandle, key: String, value: Option<String>) -> Result<(), String> {
    if !SHARED_KEYS.contains(&key.as_str()) {
        return Err("bad-key".into());
    }
    let _guard = SHARED_LOCK.lock().map_err(|e| e.to_string())?;
    let mut all = read_shared(&app);
    match value {
        Some(v) => all.insert(key, v),
        None => all.remove(&key),
    };
    write_atomic(&shared_path(&app)?, &serde_json::to_string(&all).map_err(|e| e.to_string())?)
}

fn page_of_file(path: &str) -> &'static str {
    let lower = path.to_lowercase();
    if lower.ends_with(".skdoc") {
        "doc/index.html"
    } else if lower.ends_with(".sksheet") {
        "sheet/index.html"
    } else if lower.ends_with(".skslide") {
        "slide/index.html"
    } else {
        "index.html"
    }
}

// 引数から、実在する SK Hub Office のファイルのパスだけを取り出す
fn office_files<I: IntoIterator<Item = String>>(args: I) -> Vec<String> {
    args.into_iter()
        .filter(|a| {
            let p = Path::new(a);
            p.is_absolute() && p.is_file() && p.file_name().and_then(|n| n.to_str()).map(has_ext).unwrap_or(false)
        })
        .collect()
}

// page: 開くページ（許可したものだけ）。hash: 開くファイルの指定（"f=…"）。open: フォルダの外のファイルのパス
fn open_window(app: &tauri::AppHandle, page: &str, hash: Option<&str>, open: Option<&str>) -> Result<(), String> {
    const PAGES: [&str; 4] = ["index.html", "doc/index.html", "sheet/index.html", "slide/index.html"];
    if !PAGES.contains(&page) {
        return Err(format!("bad-page: {page}"));
    }
    let label = format!("w{}", WINDOW_SEQ.fetch_add(1, Ordering::Relaxed));
    let kind = app_kind(app);
    let script = format!(
        "(function(){{try{{var h={};if(h&&!location.hash)history.replaceState(null,'','#'+h);window.__SK_OPEN__={};window.__SK_APP__={};window.__SK_SHARED__={};}}catch(e){{}}}})();",
        serde_json::to_string(&hash).map_err(|e| e.to_string())?,
        serde_json::to_string(&open.map(|o| vec![o]).unwrap_or_default()).map_err(|e| e.to_string())?,
        serde_json::to_string(kind).map_err(|e| e.to_string())?,
        serde_json::to_string(&read_shared(app)).map_err(|e| e.to_string())?,
    );
    tauri::WebviewWindowBuilder::new(app, label, tauri::WebviewUrl::App(page.into()))
        .title(app_title(kind))
        .inner_size(1280.0, 800.0)
        .min_inner_size(800.0, 560.0)
        // Windows の標準の枠・タイトルバーは使わず、Office のタイトルバーに「－ □ ×」を入れる（shared/platform.js・office.css の .win-ctrl）
        .decorations(false)
        .initialization_script(script)
        .build()
        .map_err(|e| e.to_string())?;
    Ok(())
}

fn open_files_or_hub(app: &tauri::AppHandle, files: Vec<String>) {
    if files.is_empty() {
        // すでに開いているウィンドウがあれば前に出し、なければ入り口を開く
        if let Some(w) = app.webview_windows().values().next() {
            let _ = w.unminimize();
            let _ = w.set_focus();
        } else {
            let _ = open_window(app, home_page(app_kind(app)), None, None);
        }
        return;
    }
    for f in files {
        let _ = open_window(app, page_of_file(&f), None, Some(&f));
    }
}

#[tauri::command]
async fn new_window(app: tauri::AppHandle, page: String, hash: Option<String>) -> Result<(), String> {
    open_window(&app, &page, hash.as_deref(), None)
}

// ---------- 外部ブラウザでのログイン ----------
// 画面（WebView）のポップアップは使わず、いつものブラウザでログインしてもらう（RFC 8252 のループバック方式）
//   1. 127.0.0.1 の空いているポートで待ちうけ、URL の {port}・{redirect} をそれに置きかえてブラウザで開く
//   2. ログインが終わると、ブラウザが http://127.0.0.1:ポート/callback?… に戻ってくる。その検索文字列（code・state・hint など）を画面に返す
// 開けるのは、SK Hub Systems（sk-hub.jp）と Google のログインのページだけ

static LOGIN_GEN: AtomicUsize = AtomicUsize::new(0);

fn login_host_ok(url: &url::Url) -> bool {
    url.scheme() == "https"
        && url
            .host_str()
            .map(|h| h == "accounts.google.com" || h == "sk-hub.jp" || h.ends_with(".sk-hub.jp"))
            .unwrap_or(false)
}

const DONE_PAGE: &str = "<!doctype html><html lang=\"ja\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>SK Hub Office</title><body style=\"font-family:system-ui,'Yu Gothic',sans-serif;display:grid;place-items:center;min-height:100vh;margin:0;background:#faf8ff;color:#1a1c20\"><div style=\"text-align:center;padding:24px\"><h1 style=\"font-size:20px\">ログインしました</h1><p>このタブを閉じて、SK Hub Office に戻ってください。</p></div></body></html>";

#[tauri::command]
async fn external_login(
    app: tauri::AppHandle,
    url_template: String,
    timeout_secs: u64,
) -> Result<std::collections::HashMap<String, String>, String> {
    use std::io::{Read, Write};
    let mine = LOGIN_GEN.fetch_add(1, Ordering::SeqCst) + 1; // 前の待ちうけは、ここで終わる
    tauri::async_runtime::spawn_blocking(move || {
        let listener = std::net::TcpListener::bind("127.0.0.1:0").map_err(|e| e.to_string())?;
        listener.set_nonblocking(true).map_err(|e| e.to_string())?;
        let port = listener.local_addr().map_err(|e| e.to_string())?.port();
        let redirect = format!("http://127.0.0.1:{port}/callback");
        let encoded: String = url::form_urlencoded::byte_serialize(redirect.as_bytes()).collect();
        let target = url::Url::parse(&url_template.replace("{port}", &port.to_string()).replace("{redirect}", &encoded))
            .map_err(|e| e.to_string())?;
        if !login_host_ok(&target) {
            return Err("bad-login-url".into());
        }
        tauri_plugin_opener::OpenerExt::opener(&app)
            .open_url(target.as_str(), None::<&str>)
            .map_err(|e| e.to_string())?;
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(timeout_secs.clamp(10, 900));
        loop {
            if LOGIN_GEN.load(Ordering::SeqCst) != mine {
                return Err("cancelled".into());
            }
            if std::time::Instant::now() > deadline {
                return Err("timeout".into());
            }
            let (mut stream, _) = match listener.accept() {
                Ok(s) => s,
                Err(e) if e.kind() == std::io::ErrorKind::WouldBlock => {
                    std::thread::sleep(std::time::Duration::from_millis(120));
                    continue;
                }
                Err(e) => return Err(e.to_string()),
            };
            let _ = stream.set_nonblocking(false);
            let _ = stream.set_read_timeout(Some(std::time::Duration::from_secs(5)));
            let mut buf = [0u8; 8192];
            let n = stream.read(&mut buf).unwrap_or(0);
            let head = String::from_utf8_lossy(&buf[..n]);
            let path = head.lines().next().and_then(|l| l.split_whitespace().nth(1)).unwrap_or("/");
            let parsed = url::Url::parse(&format!("http://127.0.0.1{path}")).ok();
            let params: std::collections::HashMap<String, String> = parsed
                .as_ref()
                .filter(|u| u.path() == "/callback")
                .map(|u| u.query_pairs().map(|(k, v)| (k.into_owned(), v.into_owned())).collect())
                .unwrap_or_default();
            let (status, body) = if params.is_empty() { ("404 Not Found", "") } else { ("200 OK", DONE_PAGE) };
            let _ = write!(
                stream,
                "HTTP/1.1 {status}\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\nConnection: close\r\nCache-Control: no-store\r\n\r\n{body}",
                body.len()
            );
            if !params.is_empty() {
                let mut params = params;
                params.insert("_redirect".into(), redirect.clone()); // トークンを取るときに、同じ redirect_uri を送る
                return Ok(params);
            }
        }
    })
    .await
    .map_err(|e| e.to_string())?
}

// 画面の中のリンク（SK Hub のアカウントの管理・ポリシーなど）を、いつものブラウザで開く。開けるのは sk-hub.jp だけ
#[tauri::command]
fn open_external(app: tauri::AppHandle, url: String) -> Result<(), String> {
    let u = url::Url::parse(&url).map_err(|e| e.to_string())?;
    let ok = u.scheme() == "https" && u.host_str().map(|h| h == "sk-hub.jp" || h.ends_with(".sk-hub.jp")).unwrap_or(false);
    if !ok {
        return Err("bad-url".into());
    }
    tauri_plugin_opener::OpenerExt::opener(&app).open_url(u.as_str(), None::<&str>).map_err(|e| e.to_string())
}

#[tauri::command]
fn external_login_cancel() {
    LOGIN_GEN.fetch_add(1, Ordering::SeqCst);
}

// ---------- Google のトークン ----------
// リフレッシュトークンは Windows の資格情報マネージャー（OS の保管庫）に入れ、画面（JS）には渡さない
const KEYRING_SERVICE: &str = "jp.sk-hub.office";
const GOOGLE_TOKEN_URL: &str = "https://oauth2.googleapis.com/token";

fn refresh_entry() -> Result<keyring::Entry, String> {
    keyring::Entry::new(KEYRING_SERVICE, "google-refresh-token").map_err(|e| e.to_string())
}

fn http_agent() -> ureq::Agent {
    ureq::Agent::config_builder()
        .http_status_as_error(false)
        .timeout_global(Some(std::time::Duration::from_secs(30)))
        .build()
        .into()
}

// code を access_token に替える・保管したリフレッシュトークンで access_token を取りなおす。返すのは access_token・expires_in・error だけ
#[tauri::command]
async fn google_token(params: std::collections::HashMap<String, String>, use_stored_refresh: bool) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        const KEYS: [&str; 6] = ["grant_type", "code", "code_verifier", "redirect_uri", "client_id", "client_secret"];
        let mut form: Vec<(String, String)> = params.into_iter().filter(|(k, _)| KEYS.contains(&k.as_str())).collect();
        if use_stored_refresh {
            let rt = refresh_entry()?.get_password().map_err(|_| "no-refresh-token".to_string())?;
            form.push(("refresh_token".into(), rt));
        }
        let pairs: Vec<(&str, &str)> = form.iter().map(|(k, v)| (k.as_str(), v.as_str())).collect();
        let mut res = http_agent().post(GOOGLE_TOKEN_URL).send_form(pairs).map_err(|e| e.to_string())?;
        let status = res.status().as_u16();
        let body: serde_json::Value = res.body_mut().read_json().map_err(|e| e.to_string())?;
        if let Some(rt) = body.get("refresh_token").and_then(|v| v.as_str()) {
            refresh_entry()?.set_password(rt).map_err(|e| e.to_string())?;
        }
        if status >= 400 {
            // リフレッシュトークンが取り消された・切れた: 保管したものを消して、つなぎなおしてもらう
            if use_stored_refresh && body.get("error").and_then(|v| v.as_str()) == Some("invalid_grant") {
                let _ = refresh_entry()?.delete_credential();
            }
            return Ok(serde_json::json!({ "error": body.get("error").cloned().unwrap_or(serde_json::json!("token-error")) }));
        }
        Ok(serde_json::json!({ "access_token": body.get("access_token"), "expires_in": body.get("expires_in") }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn google_has_refresh() -> bool {
    refresh_entry().and_then(|e| e.get_password().map_err(|e| e.to_string())).is_ok()
}

// つなぎを外す（保管したリフレッシュトークンを取り消して消す）
#[tauri::command]
async fn google_signout() -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(|| {
        let entry = refresh_entry()?;
        if let Ok(rt) = entry.get_password() {
            let _ = http_agent().post("https://oauth2.googleapis.com/revoke").send_form([("token", rt.as_str())]);
        }
        let _ = entry.delete_credential();
        Ok(())
    })
    .await
    .map_err(|e| e.to_string())?
}

// ---------- システムフォント ----------
// PC に入っているフォントの名前の一覧（重複なし・五十音と ABC 順）
#[tauri::command]
async fn system_fonts() -> Result<Vec<String>, String> {
    tauri::async_runtime::spawn_blocking(|| {
        let mut db = fontdb::Database::new();
        db.load_system_fonts();
        let names: std::collections::BTreeSet<String> = db
            .faces()
            .filter_map(|f| f.families.first().map(|(n, _)| n.clone()))
            .filter(|n| !n.starts_with('.') && !n.starts_with('@'))
            .collect();
        names.into_iter().collect()
    })
    .await
    .map_err(|e| e.to_string())
}

// ---------- PDF（印刷のダイアログなしで書き出す） ----------
// 保存する場所を選んでもらう（やめたら null）
#[tauri::command]
async fn pick_pdf_path(app: tauri::AppHandle, name: String) -> Result<Option<String>, String> {
    let file = app
        .dialog()
        .file()
        .add_filter("PDF", &["pdf"])
        .set_file_name(format!("{}.pdf", name.trim_end_matches(".pdf")))
        .blocking_save_file();
    match file {
        Some(p) => Ok(Some(p.into_path().map_err(|e| e.to_string())?.to_string_lossy().into_owned())),
        None => Ok(None),
    }
}

// 呼んだウィンドウの、いまの画面（印刷用の見た目に切りかえてあるもの）を PDF にする。用紙の大きさは mm
#[cfg(windows)]
#[tauri::command]
async fn print_pdf(window: tauri::WebviewWindow, path: String, width_mm: f64, height_mm: f64) -> Result<(), String> {
    use webview2_com::Microsoft::Web::WebView2::Win32::{
        ICoreWebView2_2, ICoreWebView2_7, ICoreWebView2Environment6, COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT,
    };
    use windows::core::{Interface, HSTRING};

    if !path.to_lowercase().ends_with(".pdf") {
        return Err("bad-path".into());
    }
    let (tx, rx) = std::sync::mpsc::channel::<Result<(), String>>();
    window
        .with_webview(move |wv| {
            let run = || -> Result<(), String> {
                unsafe {
                    let core = wv.controller().CoreWebView2().map_err(|e| e.to_string())?;
                    let core2: ICoreWebView2_2 = core.cast().map_err(|e| e.to_string())?;
                    let core7: ICoreWebView2_7 = core.cast().map_err(|e| e.to_string())?;
                    let env: ICoreWebView2Environment6 = core2
                        .Environment()
                        .map_err(|e| e.to_string())?
                        .cast()
                        .map_err(|e| e.to_string())?;
                    let settings = env.CreatePrintSettings().map_err(|e| e.to_string())?;
                    let _ = settings.SetOrientation(COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT);
                    let _ = settings.SetPageWidth(width_mm / 25.4);
                    let _ = settings.SetPageHeight(height_mm / 25.4);
                    let _ = settings.SetMarginTop(0.0);
                    let _ = settings.SetMarginBottom(0.0);
                    let _ = settings.SetMarginLeft(0.0);
                    let _ = settings.SetMarginRight(0.0);
                    let _ = settings.SetShouldPrintBackgrounds(true);
                    let _ = settings.SetShouldPrintHeaderAndFooter(false);
                    let path = HSTRING::from(path);
                    webview2_com::PrintToPdfCompletedHandler::wait_for_async_operation(
                        Box::new(move |handler| {
                            core7.PrintToPdf(&path, &settings, &handler).map_err(webview2_com::Error::WindowsError)
                        }),
                        Box::new(|error_code, ok| {
                            error_code?;
                            if ok {
                                Ok(())
                            } else {
                                Err(windows::core::Error::from(windows::Win32::Foundation::E_FAIL))
                            }
                        }),
                    )
                    .map_err(|e| format!("{e:?}"))
                }
            };
            let _ = tx.send(run());
        })
        .map_err(|e| e.to_string())?;
    tauri::async_runtime::spawn_blocking(move || rx.recv().map_err(|e| e.to_string())?)
        .await
        .map_err(|e| e.to_string())?
}

#[cfg(not(windows))]
#[tauri::command]
async fn print_pdf(_window: tauri::WebviewWindow, _path: String, _width_mm: f64, _height_mm: f64) -> Result<(), String> {
    Err("unsupported".into())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        // すでに起動しているときに、また起動された（別のファイルをダブルクリックした）ら、いまの SK Hub Office で開く
        .plugin(tauri_plugin_single_instance::init(|app, argv, _cwd| {
            open_files_or_hub(app, office_files(argv.into_iter().skip(1)));
        }))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            open_files_or_hub(app.handle(), office_files(std::env::args().skip(1)));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            new_window,
            shared_all,
            shared_set,
            external_login,
            external_login_cancel,
            open_external,
            google_token,
            google_has_refresh,
            google_signout,
            system_fonts,
            pick_pdf_path,
            print_pdf,
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
        ])
        .run(tauri::generate_context!())
        .expect("SK Hub Office の起動に失敗しました");
}
