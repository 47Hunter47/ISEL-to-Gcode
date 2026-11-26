// The thin IPC skin over the conversion core. The state machine that lived in
// app.rs (worker threads + channels drained by the UI pass) is replaced by
// events emitted from the convert worker and by the web front end.
//
// Every blocking call (rfd dialogs, file IO, the conversion itself) runs in an
// async command: tauri runs async commands on its thread pool, never on the
// window thread. That keeps the rule app.rs enforced with its spawned threads:
// a shell dialog on a network share may take seconds, but the window never
// freezes.

use crate::convert::convert_file;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::Mutex;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, webview::WebviewWindowBuilder};

pub const PREVIEW_LABEL: &str = "preview";

/// Seeds the input field when the exe was launched with a file argument
/// (Windows "Open with" / drag onto the exe), like app.rs's `prefill`.
pub struct Prefill(pub Option<String>);

/// preview_view rebuilds the scene on every pan/zoom event; re-parsing the
/// G-code each time would stutter on big programs, so the parse result is
/// cached per path. `convert` drops its output path from the cache, so a fresh
/// conversion is never previewed stale.
pub struct GcodeCache(pub Mutex<HashMap<String, crate::gcode::GcodeParsed>>);

#[derive(Clone, Serialize)]
struct LogMsg {
    line: String,
}

#[derive(Clone, Serialize)]
struct ProgressMsg {
    percent: f64,
    text: String,
}

#[derive(Clone, Serialize)]
struct DoneMsg {
    minutes: f64,
    output: String,
}

#[derive(Clone, Serialize)]
struct FailedMsg {
    error: String,
}

#[derive(Clone, Serialize)]
struct LoadPath {
    path: String,
}

/// The validation `start_convert` in app.rs ran before asking where to save.
fn check_input_plan(path: &str) -> Result<(), String> {
    if path.is_empty() {
        return Err("No input file selected".into());
    }
    if !std::path::Path::new(path).is_file() {
        return Err(format!("File not found: {}", path));
    }
    Ok(())
}

#[tauri::command]
pub async fn check_input(path: String) -> Result<(), String> {
    check_input_plan(&path)
}

/// `rfd`'s Windows backend runs the modal IFileDialog on the calling thread;
/// being async, this runs on tauri's thread pool instead.
#[tauri::command]
pub async fn pick_file() -> Option<String> {
    rfd::FileDialog::new()
        .pick_file()
        .map(|p| p.to_string_lossy().into_owned())
}

#[tauri::command]
pub async fn save_file(default_name: String) -> Option<String> {
    rfd::FileDialog::new()
        .set_file_name(default_name)
        .add_filter("G-code Files", &["ngc"])
        .add_filter("All Files", &["*"])
        .save_file()
        .map(|p| p.to_string_lossy().into_owned())
}

/// Runs `convert_file` on a blocking task and forwards its log/progress
/// closures as window events. `convert_file` itself is untouched.
#[tauri::command]
pub async fn convert(
    app: AppHandle,
    input: String,
    output: String,
    arcs: bool,
    corners: bool,
) -> Result<(), String> {
    let (a_log, a_prog, a_end) = (app.clone(), app.clone(), app.clone());
    let out_for_cache = output.clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut log = |m: &str| {
            let _ = a_log.emit("convert-log", LogMsg { line: m.to_string() });
        };
        let mut progress: Option<Box<dyn FnMut(f64, &str)>> = Some(Box::new(
            move |p: f64, t: &str| {
                let _ = a_prog.emit("convert-progress", ProgressMsg { percent: p, text: t.to_string() });
            },
        ));
        match convert_file(&input, &output, arcs, corners, &mut log, &mut progress) {
            Ok(mins) => {
                let _ = a_end.emit("convert-done", DoneMsg { minutes: mins, output });
            }
            Err(e) => {
                let _ = a_end.emit("convert-failed", FailedMsg { error: e });
            }
        }
    })
    .await
    .map_err(|e| format!("Conversion task failed: {e}"))?;
    if let Some(cache) = app.try_state::<GcodeCache>() {
        cache.0.lock().unwrap_or_else(|e| e.into_inner()).remove(&out_for_cache);
    }
    Ok(())
}

#[tauri::command]
pub fn get_prefill(state: tauri::State<'_, Prefill>) -> Option<String> {
    state.0.clone()
}

/// The footer shows this; a command avoids the app-info capability the JS
/// API would need.
#[tauri::command]
pub fn get_version() -> &'static str {
    crate::app::APP_VERSION
}

#[tauri::command]
pub async fn parse_gcode(path: String) -> Result<crate::gcode::GcodeParsed, String> {
    crate::gcode::parse_gcode(&path)
}

#[derive(Deserialize)]
pub struct PreviewReq {
    pub path: String,
    pub width: f32,
    pub height: f32,
    pub iso: bool,
    /// None asks the server to fit the view (window opened or resized).
    pub camera: Option<crate::preview::Camera>,
}

#[tauri::command]
pub async fn preview_view(
    req: PreviewReq,
    cache: tauri::State<'_, GcodeCache>,
) -> Result<crate::preview::PreviewScene, String> {
    let mut map = cache.0.lock().unwrap_or_else(|e| e.into_inner());
    if !map.contains_key(&req.path) {
        map.insert(req.path.clone(), crate::gcode::parse_gcode(&req.path)?);
    }
    let data = map.get(&req.path).unwrap();
    Ok(crate::preview::scene(data, req.width, req.height, req.iso, req.camera.as_ref()))
}

/// Mirror of app.rs `open_preview` decision making: a missing output is a
/// silent no-op, an empty program is a message instead of a window.
fn open_preview_plan(path: &str) -> Result<Option<String>, String> {
    if !std::path::Path::new(path).is_file() {
        return Ok(None);
    }
    match crate::gcode::parse_gcode(path) {
        Ok(data) if data.points.is_empty() => Ok(Some("No moves found in G-code file.".into())),
        Ok(_) => Ok(None),
        Err(e) => Err(format!("Could not read G-code:\n{}", e)),
    }
}

/// The preview lives in its own OS window (like app.rs's separate viewport).
/// Returns a message the front end must show instead of opening (empty file).
pub fn open_preview(app: &AppHandle, path: String) -> Result<Option<String>, String> {
    if let Some(msg) = open_preview_plan(&path)? {
        return Ok(Some(msg));
    }
    let title = format!(
        "Path Preview \u{2014} {}",
        std::path::Path::new(&path)
            .file_name()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default()
    );
    if let Some(w) = app.get_webview_window(PREVIEW_LABEL) {
        // a stale preview window: reuse it and tell it to load the new path
        let _ = w.set_title(&title);
        let _ = w.emit("load-path", LoadPath { path: path.clone() });
        let _ = w.set_focus();
        return Ok(None);
    }
    let url = format!("preview.html?p={}", percent_encode(&path.replace('\\', "/")));
    WebviewWindowBuilder::new(app, PREVIEW_LABEL, WebviewUrl::App(url.into()))
        .title(&title)
        .inner_size(960.0, 720.0)
        .resizable(true)
        .build()
        .map_err(|e| format!("Could not open the preview window: {e}"))?;
    Ok(None)
}

#[tauri::command]
pub async fn open_preview_window(app: AppHandle, path: String) -> Result<Option<String>, String> {
    open_preview(&app, path)
}

/// Percent-encode a path for the preview window's query string. `#` and the
/// space are legal in Windows file names and must not leak raw into a URL.
fn percent_encode(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for b in s.as_bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' | b'/' | b':' => {
                out.push(*b as char)
            }
            _ => out.push_str(&format!("%{:02X}", b)),
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_input_is_rejected() {
        assert_eq!(check_input_plan("").unwrap_err(), "No input file selected");
    }

    #[test]
    fn missing_input_file_is_rejected() {
        let e = check_input_plan("C:/definitely/not/here.isel").unwrap_err();
        assert!(e.contains("File not found"), "{e}");
    }

    #[test]
    fn open_preview_reports_missing_file_without_a_window() {
        // silent no-op, exactly like app.rs's open_preview (its test of the
        // same name): missing output never opens anything, never talks.
        assert_eq!(open_preview_plan("C:/definitely/not/here.ngc").unwrap(), None);
    }

    #[test]
    fn empty_gcode_yields_message_not_window() {
        let dir = std::env::temp_dir();
        let path = dir.join(format!("isel_empty_{}.ngc", std::process::id()));
        std::fs::write(&path, "M30\n").unwrap();
        let r = open_preview_plan(&path.to_string_lossy()).unwrap();
        assert_eq!(r.as_deref(), Some("No moves found in G-code file."));
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn open_preview_on_real_golden_opens_window() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/sample_expected.arcs.ngc");
        assert_eq!(open_preview_plan(path).unwrap(), None);
    }

    #[test]
    fn junk_file_is_no_moves_not_an_error() {
        // parse_gcode never errors on content — unknown lines and malformed
        // numbers are skipped, exactly as the egui preview did; only file
        // access can fail. Cargo.toml is used as a real junk file.
        let r = open_preview_plan(concat!(env!("CARGO_MANIFEST_DIR"), "/Cargo.toml")).unwrap();
        assert_eq!(r.as_deref(), Some("No moves found in G-code file."));
    }

    #[test]
    fn percent_encode_escapes_url_breakers() {
        assert_eq!(percent_encode("C:/parts/x y#z.isel"), "C:/parts/x%20y%23z.isel");
        assert_eq!(percent_encode("a/b:c-d.ngc"), "a/b:c-d.ngc");
    }
}
