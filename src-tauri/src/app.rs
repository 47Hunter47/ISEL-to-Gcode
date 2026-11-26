// Window creation and startup diagnostics: what `run_with` in the old egui
// app.rs did for eframe, now for tauri. The startup log, the panic hook and
// the error box are kept verbatim — on a machine where the window never
// appears they are the only evidence.

use crate::commands::{self, GcodeCache, Prefill};
use std::sync::Mutex;
use tauri::Manager;

pub const APP_VERSION: &str = env!("CARGO_PKG_VERSION");

/// First thing `run_with` does. Lets us tell "the exe never ran" (SmartScreen /
/// AV quarantine) apart from "it ran and failed to open a window".
fn startup_log(line: &str) {
    use std::io::Write;
    let path = std::env::temp_dir().join("ISEL_to_Gcode_startup.log");
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
    {
        let _ = writeln!(
            f,
            "{} {}",
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.as_secs())
                .unwrap_or(0),
            line
        );
    }
}

// A GUI-subsystem exe has no console: without this, a panic in release leaves
// no trace at all.
fn show_error(text: &str) {
    rfd::MessageDialog::new()
        .set_title("ISEL to G-code")
        .set_description(text)
        .show();
}

pub fn run_with(prefill: Option<String>) {
    startup_log(&format!(
        "started v{} exe={:?} mtime={} prefill={:?}",
        APP_VERSION,
        std::env::current_exe(),
        // build fingerprint: which exe he actually ran
        std::env::current_exe()
            .ok()
            .and_then(|p| std::fs::metadata(&p).ok())
            .and_then(|m| m.modified().ok())
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| d.as_secs())
            .unwrap_or(0),
        prefill
    ));
    std::panic::set_hook(Box::new(|info| {
        let msg = info.to_string();
        startup_log(&format!("PANIC {msg}"));
        show_error(&format!(
            "A problem occurred:\n\n{msg}\n\nDetails: %TEMP%\\ISEL_to_Gcode_startup.log"
        ));
    }));

    let app = match tauri::Builder::default()
        .manage(Prefill(prefill.clone()))
        .manage(GcodeCache(Mutex::new(None)))
        .invoke_handler(tauri::generate_handler![
            commands::check_input,
            commands::pick_file,
            commands::save_file,
            commands::convert,
            commands::get_prefill,
            commands::get_version,
            commands::preview_stats,
            commands::preview_view,
            commands::open_preview_window
        ])
        .setup(move |app| {
            startup_log("setup ran");
            let handle = app.handle();
            #[cfg(windows)]
            if let Some(w) = handle.get_webview_window("main") {
                // Mica first (built into Win11), acrylic as the Win10 fallback.
                // If both fail (older GPU, remote session) the opaque CSS
                // background still paints, so the window is never black.
                if window_vibrancy::apply_mica(&w, Some(true)).is_err() {
                    let _ = window_vibrancy::apply_acrylic(&w, Some((18, 18, 20, 255)));
                }
                let _ = w.set_title(&format!("ISEL \u{2192} G-code Converter v{}", APP_VERSION));
                // The transparent borderless window is created hidden; show it
                // explicitly. (Ground truth = WS_VISIBLE bit: measured on-par
                // with a control app; IsWindowVisible lies for every window in
                // a headless/disconnected session.)
                let _ = w.show();
            }
            #[cfg(target_os = "macos")]
            if let Some(w) = handle.get_webview_window("main") {
                let _ = window_vibrancy::apply_vibrancy(
                    &w,
                    window_vibrancy::NSVisualEffectMaterial::HudWindow,
                    Some(window_vibrancy::NSVisualEffectState::Active),
                    Some((18.0, 18.0)),
                );
            }
            // ISEL_PREVIEW=<out.ngc>: open the path preview at startup, so a
            // machine with no usable desktop can still prove the preview is
            // its own OS window (EnumWindows sees a second top-level window).
            if let Some(p) = std::env::var_os("ISEL_PREVIEW") {
                let path = std::path::PathBuf::from(p).to_string_lossy().into_owned();
                let opened = match commands::open_preview(handle, path.clone()) {
                    Ok(None) => true,
                    Ok(Some(msg)) => {
                        startup_log(&format!("ISEL_PREVIEW message: {msg}"));
                        false
                    }
                    Err(e) => {
                        startup_log(&format!("ISEL_PREVIEW error: {e}"));
                        false
                    }
                };
                startup_log(&format!("ISEL_PREVIEW={path:?} preview_open={opened}"));
            }
            startup_log("setup finished");
            Ok(())
        })
        .build(tauri::generate_context!())
    {
        Ok(app) => app,
        Err(e) => {
            // eframe's lesson kept: an ignored error means a silent exit with
            // no window. The usual cause is a missing WebView2 runtime.
            startup_log(&format!("build failed: {e}"));
            show_error(
                "Could not open a window.\nThis app needs the Microsoft Edge WebView2 runtime \
                 (built into Windows 11; delivered with Edge updates on Windows 10).\n\n\
                 Details: %TEMP%\\ISEL_to_Gcode_startup.log",
            );
            std::process::exit(1);
        }
    };
    app.run(|_, event| {
        if let tauri::RunEvent::Exit = event {
            startup_log("run exited");
        }
    });
    startup_log("run returned Ok (window closed)");
}
