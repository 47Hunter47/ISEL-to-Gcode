// Main window: port of python run_gui() (isel_to_gcode_gui.py lines 990-1197).
// Conversion runs on a worker thread and reports back over a channel, same as
// python's threading.Thread + root.after() marshalling.

use crate::convert::convert_file;
use crate::preview::Preview;
use crate::theme;
use egui::{Color32, RichText};
use std::path::PathBuf;
use std::sync::mpsc::{channel, Receiver};
use std::time::Duration;

pub const APP_VERSION: &str = env!("CARGO_PKG_VERSION");
// The palette lives in theme.rs (industrial SCADA look, sampled from the
// Siemens keyvisual). These aliases keep the UI code short.
const BG: Color32 = theme::PANEL;
const BTN: Color32 = theme::BTN;
const BLUE: Color32 = theme::ACCENT;
const GREEN_BTN: Color32 = theme::BTN_ACTIVE;
const LOG_GREEN: Color32 = theme::LOG;

enum Msg {
    Log(String),
    Progress(f64, String),
    Done(f64, String),
    Failed(String),
}

// The status line doubles as the state light: mint on success, alarm on failure.
fn status_color(s: &str) -> Color32 {
    if s == "Error occurred" {
        theme::ERR
    } else if s == "Conversion complete!" {
        theme::OK
    } else {
        theme::TEXT
    }
}

pub struct App {
    input: String,
    use_arc: bool,
    log: Vec<String>,
    progress: f32,
    status: String,
    converting: bool,
    last_output: Option<String>,
    rx: Option<Receiver<Msg>>,
    preview: Option<Preview>,
    dialog: Option<(String, String)>,
    // startup diagnostics: how many logic passes ran / did ui ever paint
    probe: u8,
    ui_probe: u8,
}

impl Default for App {
    fn default() -> Self {
        let mut log = vec![];
        for l in [
            "ISEL to G-code Converter Ready",
            "Drag and drop a file or click Browse to start",
            "--------------------------------------------------",
        ] {
            log.push(l.to_string());
        }
        App {
            input: String::new(),
            use_arc: false,
            log,
            progress: 0.0,
            status: "Ready".to_string(),
            converting: false,
            last_output: None,
            rx: None,
            preview: None,
            dialog: None,
            probe: 0,
            ui_probe: 0,
        }
    }
}

impl App {
    fn push_log(&mut self, m: impl Into<String>) {
        self.log.push(m.into());
    }

    fn start_convert(&mut self) {
        let in_path = self.input.clone();
        if in_path.is_empty() {
            self.dialog = Some(("Error".into(), "No input file selected".into()));
            return;
        }
        if !std::path::Path::new(&in_path).is_file() {
            self.dialog = Some(("Error".into(), format!("File not found: {}", in_path)));
            return;
        }
        let base = std::path::Path::new(&in_path)
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default();
        let Some(out) = rfd::FileDialog::new()
            .set_file_name(format!("{}.ngc", base))
            .add_filter("G-code Files", &["ngc"])
            .add_filter("All Files", &["*"])
            .save_file()
        else {
            return;
        };
        let out_path = out.to_string_lossy().into_owned();

        let arc = self.use_arc;
        self.converting = true;
        self.progress = 0.0;
        self.status = "Starting conversion...".to_string();
        self.log.clear();
        self.push_log("Starting conversion...");
        self.push_log(format!("Input:   {}", in_path));
        self.push_log(format!("Output:  {}", out_path));
        self.push_log(format!(
            "Mode:    {}",
            if arc {
                "G2/G3 arc commands + full-circle detection (R format)"
            } else {
                "G1 linearised"
            }
        ));
        self.push_log("--------------------------------------------------");

        let (tx, rx) = channel();
        self.rx = Some(rx);
        std::thread::spawn(move || {
            let tx_log = tx.clone();
            let mut log = move |m: &str| {
                let _ = tx_log.send(Msg::Log(m.to_string()));
            };
            let tx_prog = tx.clone();
            let mut progress: Option<Box<dyn FnMut(f64, &str)>> = Some(Box::new(
                move |p: f64, t: &str| {
                    let _ = tx_prog.send(Msg::Progress(p, t.to_string()));
                },
            ));
            match convert_file(&in_path, &out_path, arc, &mut log, &mut progress) {
                Ok(mins) => {
                    let _ = tx.send(Msg::Done(mins, out_path));
                }
                Err(e) => {
                    let _ = tx.send(Msg::Failed(e));
                }
            }
        });
    }

    fn drain(&mut self) {
        let mut msgs = vec![];
        if let Some(rx) = &self.rx {
            while let Ok(m) = rx.try_recv() {
                msgs.push(m);
            }
        }
        for m in msgs {
            match m {
                Msg::Log(s) => self.push_log(s),
                Msg::Progress(p, t) => {
                    self.progress = (p / 100.0) as f32;
                    self.status = t;
                }
                Msg::Done(mins, out) => {
                    self.converting = false;
                    self.rx = None;
                    self.last_output = Some(out);
                    self.progress = 1.0;
                    self.status = "Conversion complete!".to_string();
                    let m = mins as u64;
                    let s = ((mins - m as f64) * 60.0) as u64;
                    self.push_log("--------------------------------------------------");
                    self.push_log(format!(
                        "\u{2713} Conversion completed  ({})",
                        if self.use_arc { "G2/G3 arc commands" } else { "G1 linearised" }
                    ));
                    self.push_log(format!("\u{23f1} Estimated program time: {} min {} sec", m, s));
                    self.push_log(
                        "\u{26a0} Program time may change according to machine parameters",
                    );
                    self.dialog = Some((
                        "Completed".into(),
                        format!(
                            "Conversion finished!\nMode: {}\n\nEstimated time: {} min {} sec",
                            if self.use_arc { "G2/G3 arc commands" } else { "G1 linearised" },
                            m,
                            s
                        ),
                    ));
                }
                Msg::Failed(e) => {
                    self.converting = false;
                    self.rx = None;
                    self.status = "Error occurred".to_string();
                    self.push_log(format!("\u{2717} Error: {}", e));
                    self.dialog = Some(("Error".into(), e));
                }
            }
        }
    }

    fn open_preview(&mut self) {
        let Some(path) = self.last_output.clone() else { return };
        if !std::path::Path::new(&path).is_file() {
            return;
        }
        match crate::gcode::parse_gcode(&path) {
            Ok(data) if data.points.is_empty() => {
                self.dialog = Some(("Preview".into(), "No moves found in G-code file.".into()));
            }
            Ok(data) => {
                let name = std::path::Path::new(&path)
                    .file_name()
                    .map(|s| s.to_string_lossy().into_owned())
                    .unwrap_or_default();
                self.preview = Some(Preview::new(name, data));
            }
            Err(e) => {
                self.dialog = Some(("Preview Error".into(), format!("Could not read G-code:\n{}", e)));
            }
        }
    }
}

impl eframe::App for App {
    // eframe 0.37 split the old `update` into `logic` (no painting) + `ui`.
    fn logic(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.drain();

        // drag & drop
        let dropped: Vec<PathBuf> = ctx.input_mut(|i| {
            std::mem::take(&mut i.raw.dropped_files)
                .iter()
                .map(|f| f.path().to_path_buf())
                .collect()
        });
        for p in dropped {
            if p.is_file() {
                self.input = p.to_string_lossy().into_owned();
                self.push_log(format!("File imported: {}", self.input));
            } else {
                self.push_log(format!("Error: Not a valid file: {}", p.display()));
            }
        }

        // eframe keeps the window hidden until it has painted a frame, and on Windows an
        // invisible window gets no RedrawRequested, so nothing would ever paint and the
        // window would stay hidden forever. If egui reports it hidden, ask for a paint and
        // force it shown, and keep the loop ticking until it really is visible.
        let (visible, inner, monitor) = ctx.input(|i| {
            let vp = i.viewport();
            (
                vp.visible(),
                vp.inner_rect
                    .map(|r| (r.min.x as i32, r.min.y as i32, r.max.x as i32, r.max.y as i32)),
                vp.monitor_size.map(|s| (s.x as i32, s.y as i32)),
            )
        });
        if visible == Some(false) {
            ctx.send_viewport_cmd_to(
                egui::ViewportId::ROOT,
                egui::ViewportCommand::RequestPaintWhileHidden,
            );
            ctx.send_viewport_cmd_to(egui::ViewportId::ROOT, egui::ViewportCommand::Visible(true));
            ctx.request_repaint();
        }
        if self.probe < 5 {
            self.probe += 1;
            startup_log(&format!(
                "logic pass {} visible={:?} inner={:?} monitor={:?}",
                self.probe, visible, inner, monitor
            ));
        }
    }

    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        if self.ui_probe == 0 {
            self.ui_probe = 1;
            startup_log("ui pass ran (first paint)");
        }
        let ctx = ui.ctx().clone();
        egui::Frame::new()
            .fill(BG)
            .inner_margin(10.0)
            .show(ui, |ui| {
                ui.label(RichText::new("ISEL FILE").color(theme::TEXT));
                ui.add(
                    egui::TextEdit::singleline(&mut self.input)
                        .desired_width(f32::INFINITY)
                        .text_color(theme::TEXT),
                );

                ui.add_space(8.0);
                ui.horizontal(|ui| {
                    if ui
                        .add_enabled(
                            !self.converting,
                            egui::Button::new("BROWSE").fill(BTN).min_size(egui::vec2(110.0, 28.0)),
                        )
                        .clicked()
                    {
                        if let Some(p) = rfd::FileDialog::new().pick_file() {
                            self.input = p.to_string_lossy().into_owned();
                            self.push_log(format!("File selected: {}", self.input));
                        }
                    }
                    let convert_label = if self.converting { "CONVERTING..." } else { "CONVERT" };
                    if ui
                        .add_enabled(
                            !self.converting,
                            egui::Button::new(
                                RichText::new(convert_label).strong().color(
                                    // cyan fill wants dark glyphs, the dim mid-conversion fill wants light
                                    if self.converting { theme::TEXT } else { theme::ON_ACCENT },
                                ),
                            )
                                .fill(BLUE)
                                .min_size(egui::vec2(110.0, 28.0)),
                        )
                        .clicked()
                    {
                        self.start_convert();
                    }
                    let has_out = self.last_output.is_some();
                    if ui
                        .add_enabled(
                            has_out,
                            egui::Button::new(RichText::new("PREVIEW").strong())
                                .fill(GREEN_BTN)
                                .min_size(egui::vec2(110.0, 28.0)),
                        )
                        .clicked()
                    {
                        self.open_preview();
                    }
                });

                ui.checkbox(
                    &mut self.use_arc,
                    "Use G2/G3 arc commands  (smaller file, requires arc-capable controller)",
                );

                ui.add_space(6.0);
                ui.label(RichText::new(self.status.clone()).color(status_color(&self.status)));
                ui.add(
                    egui::ProgressBar::new(self.progress)
                        .desired_width(f32::INFINITY)
                        .fill(theme::ACCENT)
                        .desired_height(20.0),
                );

                ui.add_space(6.0);
                ui.label("CONVERSION LOG");
                egui::Frame::new()
                    .fill(theme::DEEP)
                    .inner_margin(4.0)
                    .show(ui, |ui| {
                        egui::ScrollArea::vertical()
                            .stick_to_bottom(true)
                            .auto_shrink([false, false])
                            .show(ui, |ui| {
                                for line in &self.log {
                                    ui.label(
                                        RichText::new(line).monospace().color(LOG_GREEN),
                                    );
                                }
                            });
                    });
            });

        if let Some(p) = &mut self.preview {
            p.show(&ctx);
            if !p.open {
                self.preview = None;
            }
        }

        if let Some((title, body)) = self.dialog.clone() {
            let mut close = false;
            egui::Window::new(title)
                .collapsible(false)
                .resizable(false)
                .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
                .show(&ctx, |ui| {
                    ui.label(body);
                    if ui.button("OK").clicked() {
                        close = true;
                    }
                });
            if close {
                self.dialog = None;
            }
        }

        if self.converting {
            ctx.request_repaint_after(Duration::from_millis(80));
        }
    }
}

pub fn run() {
    run_with(None);
}

/// `prefill` seeds the input path — set when the exe is launched with a file
/// argument (Windows "Open with" / drag onto the exe).
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

    let icon = image::load_from_memory(include_bytes!("../assets/icon.ico"))
        .ok()
        .map(|img| {
            let img = img.to_rgba8();
            let (w, h) = img.dimensions();
            egui::IconData {
                rgba: img.into_raw(),
                width: w,
                height: h,
            }
        });

    // eframe's default renderer is wgpu (DirectX); glow (OpenGL) is the fallback.
    // OpenGL comes first: it needs no async device creation, so a driver that
    // stalls inside wgpu can't block startup. ISEL_RENDERER=wgpu forces the other order.
    let mut order = [eframe::Renderer::Glow, eframe::Renderer::Wgpu];
    if std::env::var("ISEL_RENDERER").as_deref() == Ok("wgpu") {
        order = [eframe::Renderer::Wgpu, eframe::Renderer::Glow];
    }
    for renderer in order {
        startup_log(&format!("trying {renderer:?}"));
        let mut viewport = egui::ViewportBuilder::default()
            .with_inner_size([520.0, 620.0])
            .with_title(format!("ISEL \u{2192} G-code Converter v{}", APP_VERSION));
        if let Some(icon) = icon.clone() {
            viewport = viewport.with_icon(icon);
        }
        let options = eframe::NativeOptions {
            viewport,
            renderer,
            ..Default::default()
        };
        let prefill = prefill.clone();
        match eframe::run_native(
            "ISEL to G-code Converter",
            options,
            Box::new(move |cc| {
                // Fills, strokes, corner radii, selection and the text hierarchy
                // all come from theme::visuals().
                cc.egui_ctx.set_visuals(theme::visuals());
                let mut app = App::default();
                if let Some(p) = prefill {
                    app.input = p;
                    app.push_log(format!("File selected: {}", app.input));
                }
                startup_log("app creator finished");
                Ok(Box::new(app))
            }),
        ) {
            Ok(()) => {
                startup_log("run_native returned Ok (window closed)");
                return;
            }
            Err(e) => startup_log(&format!("renderer {renderer:?} failed: {e}")),
        }
    }
    show_error(
        "Could not open a window.\n\nDetails: %TEMP%\\ISEL_to_Gcode_startup.log",
    );
    std::process::exit(1);
}

// GUI state machine tests: no window needed, everything is channel + field state.
#[cfg(test)]
mod tests {
    use super::*;

    fn golden(name: &str) -> String {
        concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/").to_string() + name
    }

    #[test]
    fn preview_opens_from_completed_output() {
        let mut app = App::default();
        app.last_output = Some(golden("/sample_expected.arcs.ngc"));
        app.open_preview();
        let p = app.preview.expect("preview should open");
        assert!(p.open);
        assert_eq!(p.title, "sample_expected.arcs.ngc");
        assert!(!p.data.points.is_empty());
        assert!(app.dialog.is_none());
    }

    #[test]
    fn preview_reports_missing_file_without_opening_window() {
        let mut app = App::default();
        app.last_output = Some("C:/definitely/not/here.ngc".into());
        app.open_preview();
        assert!(app.preview.is_none());
        assert!(app.dialog.is_none(), "missing output is a silent no-op");
    }

    #[test]
    fn done_message_finalizes_state() {
        let mut app = App::default();
        let (tx, rx) = channel();
        app.rx = Some(rx);
        app.converting = true;
        app.push_log("Starting conversion...");
        tx.send(Msg::Progress(50.0, "Converted 1000 points".into())).unwrap();
        tx.send(Msg::Done(2.5, "C:/tmp/out.ngc".into())).unwrap();
        app.drain();
        assert!(!app.converting);
        assert!(app.rx.is_none());
        assert_eq!(app.last_output.as_deref(), Some("C:/tmp/out.ngc"));
        assert_eq!(app.status, "Conversion complete!");
        assert!((app.progress - 1.0).abs() < 1e-6);
        // progress updates the status label only (python parity), never the log
        let log = app.log.join("\n");
        assert!(!log.contains("Converted 1000 points"), "{log}");
        assert!(log.contains("2 min 30 sec"), "{log}");
        let (_, body) = app.dialog.expect("completion dialog");
        assert!(body.contains("2 min 30 sec"), "{body}");
    }

    #[test]
    fn failed_message_surfaces_error_dialog() {
        let mut app = App::default();
        let (tx, rx) = channel();
        app.rx = Some(rx);
        app.converting = true;
        tx.send(Msg::Failed("Input file not found".into())).unwrap();
        app.drain();
        assert!(!app.converting);
        assert_eq!(app.status, "Error occurred");
        let (title, body) = app.dialog.expect("error dialog");
        assert_eq!(title, "Error");
        assert!(body.contains("Input file not found"));
        assert!(app.log.join("\n").contains("Error: Input file not found"));
    }

    #[test]
    fn empty_input_starts_nothing_and_shows_error() {
        let mut app = App::default();
        app.start_convert();
        assert!(app.converting == false);
        let (_, body) = app.dialog.expect("error dialog");
        assert_eq!(body, "No input file selected");
    }

    #[test]
    fn missing_input_file_shows_error() {
        let mut app = App::default();
        app.input = "C:/definitely/not/here.isel".into();
        app.start_convert();
        assert!(!app.converting);
        let (_, body) = app.dialog.expect("error dialog");
        assert!(body.contains("File not found"), "{body}");
    }

    #[test]
    fn empty_gcode_yields_message_not_window() {
        let dir = std::env::temp_dir();
        let path = dir.join(format!("isel_empty_{}.ngc", std::process::id()));
        std::fs::write(&path, "M30\n").unwrap();
        let mut app = App::default();
        app.last_output = Some(path.to_string_lossy().into_owned());
        app.open_preview();
        assert!(app.preview.is_none());
        let (_, body) = app.dialog.expect("empty message");
        assert!(body.contains("No moves found"), "{body}");
        let _ = std::fs::remove_file(&path);
    }
}
