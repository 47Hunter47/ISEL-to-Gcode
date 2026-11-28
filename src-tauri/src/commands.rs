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
use std::sync::Mutex;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, webview::WebviewWindowBuilder};

pub const PREVIEW_LABEL: &str = "preview";

/// Seeds the input field when the exe was launched with a file argument
/// (Windows "Open with" / drag onto the exe), like app.rs's `prefill`.
pub struct Prefill(pub Option<String>);

/// Preview_view rebuilds the scene on every fit/resize/view request; re-parsing
/// the G-code each time would stutter on big programs, so the parse result is
/// kept for the file currently previewed. One entry only: switching files
/// replaces it, so memory never grows with the number of previewed programs.
/// `convert` drops its output path from the cache, so a fresh conversion is
/// never previewed stale.
pub struct GcodeCache(pub Mutex<Option<(String, crate::gcode::GcodeParsed)>>);

impl GcodeCache {
    /// Run `f` with the parse of `path`, parsing (and replacing the cached
    /// entry) on a miss. Parsing inside the lock serialises concurrent opens
    /// of the same file: exactly one of them parses.
    fn with_parsed<T>(&self, path: &str, f: impl FnOnce(&crate::gcode::GcodeParsed) -> Result<T, String>) -> Result<T, String> {
        let mut map = self.0.lock().unwrap_or_else(|e| e.into_inner());
        if map.as_ref().map(|(p, _)| p != path).unwrap_or(true) {
            let parsed = crate::gcode::parse_gcode(path)?;
            *map = Some((path.to_string(), parsed));
        }
        f(&map.as_ref().unwrap().1)
    }
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
    params: Option<crate::corner::CornerParams>,
) -> Result<(), String> {
    // The toggle is the bool; the edited parameters ride along when it is on.
    let corner = corners.then(|| params.unwrap_or_default());
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
        match convert_file(&input, &output, arcs, corner.as_ref(), &mut log, &mut progress) {
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
        let mut map = cache.0.lock().unwrap_or_else(|e| e.into_inner());
        // the output was just rewritten: a cached parse of that path is stale
        if map.as_ref().map(|(p, _)| p == &out_for_cache).unwrap_or(false) {
            *map = None;
        }
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

/// Seeds the corner-parameter editor on the About page. The defaults live in
/// `corner.rs`; a second copy in the front end would drift.
#[tauri::command]
pub fn get_corner_defaults() -> crate::corner::CornerParams {
    crate::corner::CornerParams::default()
}

/// gcode::Counts + gcode::Bounds: everything the preview window's stats bar
/// shows, without serializing the geometry to the front end twice.
#[derive(Serialize)]
pub struct GcodeSummary {
    pub counts: crate::gcode::Counts,
    pub bounds: crate::gcode::Bounds,
}

#[tauri::command]
pub async fn preview_stats(
    path: String,
    cache: tauri::State<'_, GcodeCache>,
) -> Result<GcodeSummary, String> {
    cache.with_parsed(&path, |d| {
        Ok(GcodeSummary { counts: d.counts, bounds: d.bounds })
    })
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
    cache.with_parsed(&req.path, |d| {
        Ok(crate::preview::scene(d, req.width, req.height, req.iso, req.camera.as_ref()))
    })
}

/// What opening a preview decided, with the parse it needed already done:
/// a missing output is a silent no-op, an empty program is a message instead
/// of a window, and a parse with moves comes back here so the preview window
/// never parses the same file a second time.
enum OpenPlan {
    Missing,
    NoMoves,
    Parsed(crate::gcode::GcodeParsed),
}

const NO_MOVES_MSG: &str = "No moves found in G-code file.";

fn open_preview_plan(path: &str) -> Result<OpenPlan, String> {
    if !std::path::Path::new(path).is_file() {
        return Ok(OpenPlan::Missing);
    }
    match crate::gcode::parse_gcode(path) {
        Ok(data) if data.has_moves() => Ok(OpenPlan::Parsed(data)),
        Ok(_) => Ok(OpenPlan::NoMoves),
        Err(e) => Err(format!("Could not read G-code:\n{}", e)),
    }
}

/// The preview lives in its own OS window (like app.rs's separate viewport).
/// Returns a message the front end must show instead of opening (empty file).
pub fn open_preview(app: &AppHandle, path: String) -> Result<Option<String>, String> {
    let parsed = match open_preview_plan(&path)? {
        OpenPlan::Missing => return Ok(None),
        OpenPlan::NoMoves => return Ok(Some(NO_MOVES_MSG.to_string())),
        OpenPlan::Parsed(data) => data,
    };
    // seed the preview cache with the parse this decision already paid for
    if let Some(cache) = app.try_state::<GcodeCache>() {
        *cache.0.lock().unwrap_or_else(|e| e.into_inner()) = Some((path.clone(), parsed));
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
        assert!(matches!(open_preview_plan("C:/definitely/not/here.ngc"), Ok(OpenPlan::Missing)));
    }

    #[test]
    fn empty_gcode_yields_message_not_window() {
        let dir = std::env::temp_dir();
        let path = dir.join(format!("isel_empty_{}.ngc", std::process::id()));
        std::fs::write(&path, "M30\n").unwrap();
        let r = open_preview_plan(&path.to_string_lossy()).unwrap();
        assert!(matches!(r, OpenPlan::NoMoves));
        // the python port's wording, pinned verbatim
        assert_eq!(NO_MOVES_MSG, "No moves found in G-code file.");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn open_preview_on_real_golden_opens_window() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/sample_expected.arcs.ngc");
        assert!(matches!(open_preview_plan(path), Ok(OpenPlan::Parsed(_))));
    }

    #[test]
    fn junk_file_is_no_moves_not_an_error() {
        // parse_gcode never errors on content — unknown lines and malformed
        // numbers are skipped, exactly as the egui preview did; only file
        // access can fail. Cargo.toml is used as a real junk file.
        let r = open_preview_plan(concat!(env!("CARGO_MANIFEST_DIR"), "/Cargo.toml")).unwrap();
        assert!(matches!(r, OpenPlan::NoMoves));
    }

    #[test]
    fn percent_encode_escapes_url_breakers() {
        assert_eq!(percent_encode("C:/parts/x y#z.isel"), "C:/parts/x%20y%23z.isel");
        assert_eq!(percent_encode("a/b:c-d.ngc"), "a/b:c-d.ngc");
    }

    #[test]
    fn cache_serves_hits_and_keeps_one_entry() {
        let dir = std::env::temp_dir();
        let a = dir.join(format!("isel_cache_a_{}.ngc", std::process::id()));
        let b = dir.join(format!("isel_cache_b_{}.ngc", std::process::id()));
        std::fs::write(&a, "N1 G1 X10 Y10\n").unwrap();
        std::fs::write(&b, "N1 G1 X20 Y30\n").unwrap();
        let cache = GcodeCache(Mutex::new(None));
        let ca = cache.with_parsed(&a.to_string_lossy(), |d| Ok(d.counts.cut)).unwrap();
        let ca2 = cache.with_parsed(&a.to_string_lossy(), |d| Ok(d.counts.cut)).unwrap();
        assert_eq!((ca, ca2), (2, 2));
        let cb = cache.with_parsed(&b.to_string_lossy(), |d| Ok(d.counts.cut)).unwrap();
        assert_eq!(cb, 2);
        // single active entry: switching files evicted A, memory cannot grow
        let g = cache.0.lock().unwrap();
        assert_eq!(g.as_ref().unwrap().0, b.to_string_lossy());
        drop(g);
        // file replaced by path B again but content rewritten (what a fresh
        // conversion does): a path hit on a dropped entry re-parses, and
        // `convert` drops the entry for exactly that reason.
        let mut g = cache.0.lock().unwrap();
        g.take();
        drop(g);
        std::fs::write(&b, "N1 G1 X1 Y1\nN2 G1 X5 Y5\n").unwrap();
        let cb2 = cache.with_parsed(&b.to_string_lossy(), |d| Ok(d.counts.cut)).unwrap();
        assert_eq!(cb2, 4);
        let _ = std::fs::remove_file(&a);
        let _ = std::fs::remove_file(&b);
    }

    #[test]
    fn reopen_after_rewrite_sees_the_new_program() {
        // open_preview_plan never answers from a cache: a re-converted output
        // is re-parsed by the next open, whatever the cache held.
        let dir = std::env::temp_dir();
        let p = dir.join(format!("isel_stale_{}.ngc", std::process::id()));
        std::fs::write(&p, "N1 G1 X10 Y10\n").unwrap();
        let s = &p.to_string_lossy();
        assert!(matches!(open_preview_plan(s).unwrap(), OpenPlan::Parsed(ref d) if d.counts.cut == 2));
        std::fs::write(&p, "N1 G1 X10 Y10\nN2 G1 X20 Y20\n").unwrap();
        assert!(matches!(open_preview_plan(s).unwrap(), OpenPlan::Parsed(ref d) if d.counts.cut == 4));
        let _ = std::fs::remove_file(&p);
    }
}
