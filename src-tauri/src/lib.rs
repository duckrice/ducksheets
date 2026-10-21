use std::{
    fs,
    path::{Path, PathBuf},
};

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{Emitter, Manager};

/// The files among a program's arguments, as absolute paths.
fn files_in<I: IntoIterator<Item = std::ffi::OsString>>(args: I, cwd: &Path) -> Vec<String> {
    args.into_iter()
        .map(PathBuf::from)
        .map(|path| if path.is_absolute() { path } else { cwd.join(path) })
        .filter(|path| path.is_file())
        .map(|path| path.to_string_lossy().into_owned())
        .collect()
}

/// The workbooks the app was launched with ("Open with" or a double-click).
#[tauri::command]
fn startup_files() -> Vec<String> {
    let cwd = std::env::current_dir().unwrap_or_default();
    files_in(std::env::args_os().skip(1), &cwd)
}

/// Opens a link from a workbook in the browser (or mail program). Only web and e-mail
/// addresses: a workbook from anyone must not start programs or open files.
#[tauri::command]
fn open_url(url: String) -> Result<(), String> {
    let lower = url.to_ascii_lowercase();
    if !(lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:")) {
        return Err("Only web and e-mail addresses can be opened.".into());
    }
    // Test runs check the request reaches here; they must not open the user's browser.
    if std::env::var_os("DUCKSHEETS_TEST_OFFSCREEN").is_some() {
        return Ok(());
    }
    // Explorer hands the address to the default browser; no shell, so nothing in it is run.
    std::process::Command::new("explorer").arg(&url).spawn().map(|_| ()).map_err(|e| e.to_string())
}

/// Which of the given files still exist, for the list of recent files.
#[tauri::command]
fn existing_files(paths: Vec<String>) -> Vec<bool> {
    paths.iter().map(|path| Path::new(path).is_file()).collect()
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    fs::read(&path)
        .map(Response::new)
        .map_err(|err| format!("Could not read {path}: {err}"))
}

/// Saves the request body to the path given in the `x-path` header (percent-encoded,
/// since header values must be ASCII). The bytes travel as a raw body so large
/// files aren't serialised as JSON.
#[tauri::command]
fn write_file(request: Request<'_>) -> Result<(), String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("write_file expects a binary body".into());
    };
    let encoded = request
        .headers()
        .get("x-path")
        .and_then(|value| value.to_str().ok())
        .ok_or("write_file is missing the x-path header")?;
    let path = PathBuf::from(
        percent_decode_str(encoded)
            .decode_utf8()
            .map_err(|err| err.to_string())?
            .into_owned(),
    );
    write_atomically(&path, bytes).map_err(|err| format!("Could not save {}: {err}", path.display()))
}

/// Writes next to the target and renames over it, so a failed save can't leave
/// a half-written workbook behind.
fn write_atomically(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let mut name = path.file_name().unwrap_or_default().to_os_string();
    name.push(format!(".{}.tmp", std::process::id()));
    let tmp = path.with_file_name(name);
    fs::write(&tmp, bytes)?;
    fs::rename(&tmp, path).inspect_err(|_| {
        let _ = fs::remove_file(&tmp);
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let mut builder = tauri::Builder::default();
    // Opening a workbook while DuckSheets is running hands it to the open window,
    // which shows it in a new tab. Tests set DUCKSHEETS_MULTI_INSTANCE so they
    // can run beside a copy that is in use.
    if std::env::var_os("DUCKSHEETS_MULTI_INSTANCE").is_none() {
        builder = builder.plugin(tauri_plugin_single_instance::init(|app, args, cwd| {
            let files = files_in(args.into_iter().skip(1).map(Into::into), Path::new(&cwd));
            let _ = app.emit("open-files", files);
            if let Some(window) = app.get_webview_window("main") {
                let _ = window.unminimize();
                let _ = window.set_focus();
            }
        }));
    }
    builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .setup(|app| {
            let config = &app.config().app.windows[0];
            let builder = tauri::WebviewWindowBuilder::from_config(app.handle(), config)?;
            // Tests open the window off-screen and in the background, so they
            // can run without taking over the screen. A window can't be created
            // off-screen, so it's created hidden, moved, then shown without focus.
            if std::env::var_os("DUCKSHEETS_TEST_OFFSCREEN").is_some() {
                let window = builder.visible(false).focused(false).skip_taskbar(true).build()?;
                window.set_position(tauri::PhysicalPosition::new(-20000, -20000))?;
                window.show()?;
            } else {
                builder.build()?;
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            startup_files,
            existing_files,
            read_file,
            write_file,
            open_url
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
