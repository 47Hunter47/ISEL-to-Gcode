// Toolpath preview window: port of python open_preview() (isel_to_gcode_gui.py
// lines 216-490) onto egui's painter. Top-down + isometric, pan, wheel zoom, fit.
// shortcut: rapid moves are solid amber, python drew them dashed; upgrade if the
// two move classes are ever confused at a glance.

use crate::gcode::{GcodeParsed, Mt};
use crate::theme;
use egui::{Color32, Pos2, Rect, RichText, Sense, Stroke};

// Same palette as the panels (theme.rs): cyan = cutting, mint = arc, amber =
// rapid. Everything else recedes so the toolpath is the only loud thing.
pub const COLOR_CUT: Color32 = theme::ACCENT;
pub const COLOR_RAPID: Color32 = theme::RAPID;
pub const COLOR_ARC: Color32 = theme::ARC;
pub const COLOR_BG: Color32 = theme::DEEP;
pub const COLOR_GRID: Color32 = theme::GRID;
pub const COLOR_AXIS: Color32 = theme::AXIS;
const PADDING: f32 = 48.0;

pub struct Preview {
    pub title: String,
    pub open: bool,
    pub data: GcodeParsed,
    iso: bool,
    ox: f32,
    oy: f32,
    scale: f32,
    last_rect: Rect,
}

fn thousands(n: usize) -> String {
    let s = n.to_string();
    let mut out = String::with_capacity(s.len() + s.len() / 3);
    for (i, c) in s.chars().enumerate() {
        if i > 0 && (s.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out
}

impl Preview {
    pub fn new(title: String, data: GcodeParsed) -> Self {
        Preview {
            title,
            open: true,
            data,
            iso: false,
            ox: 0.0,
            oy: 0.0,
            scale: 1.0,
            last_rect: Rect::NOTHING,
        }
    }

    /// world -> screen; ox/oy/scale applied here (python `project`).
    fn project(&self, w: (f64, f64, f64)) -> Pos2 {
        let (wx, wy, wz) = (w.0 as f32, w.1 as f32, w.2 as f32);
        if self.iso {
            let c = 30f32.to_radians().cos();
            let s = 30f32.to_radians().sin();
            Pos2::new(
                self.ox + (wx - wy) * c * self.scale,
                self.oy + ((wx + wy) * s - wz) * self.scale,
            )
        } else {
            Pos2::new(self.ox + wx * self.scale, self.oy - wy * self.scale)
        }
    }

    /// Bounds of the projected shape, origin-free (python `projected_bounds`).
    fn projected_bounds(&self) -> (f32, f32, f32, f32) {
        let (mut min_px, mut max_px) = (f32::INFINITY, f32::NEG_INFINITY);
        let (mut min_py, mut max_py) = (f32::INFINITY, f32::NEG_INFINITY);
        for p in &self.data.points {
            let (wx, wy, wz) = (p.x as f32, p.y as f32, p.z as f32);
            let (px, py) = if self.iso {
                let c = 30f32.to_radians().cos();
                let s = 30f32.to_radians().sin();
                ((wx - wy) * c, (wx + wy) * s - wz)
            } else {
                (wx, -wy)
            };
            min_px = min_px.min(px);
            max_px = max_px.max(px);
            min_py = min_py.min(py);
            max_py = max_py.max(py);
        }
        (min_px, max_px, min_py, max_py)
    }

    /// wheel zoom around a screen anchor: the world point under `c` stays put.
    fn zoom_at(&mut self, c: Pos2, factor: f32) {
        self.ox = c.x + (self.ox - c.x) * factor;
        self.oy = c.y + (self.oy - c.y) * factor;
        self.scale *= factor;
    }

    fn fit(&mut self, rect: Rect) {
        let (min_px, max_px, min_py, max_py) = self.projected_bounds();
        if !min_px.is_finite() {
            return;
        }
        let span_x = (max_px - min_px).max(1e-6);
        let span_y = (max_py - min_py).max(1e-6);
        let s = ((rect.width() - 2.0 * PADDING) / span_x)
            .min((rect.height() - 2.0 * PADDING) / span_y);
        self.scale = s;
        self.ox = rect.left() + (rect.width() - span_x * s) / 2.0 - min_px * s;
        self.oy = rect.top() + (rect.height() - span_y * s) / 2.0 - min_py * s;
        self.last_rect = rect;
    }

    /// Projected polylines in draw order (rapid -> cut -> arc). Split out of
    /// `show` so the geometry is testable without a desktop session.
    fn scene(&self) -> Vec<(Vec<Pos2>, Color32)> {
        let mut out = Vec::new();
        for (mt, color) in [
            (Mt::Rapid, COLOR_RAPID),
            (Mt::Cut, COLOR_CUT),
            (Mt::Arc, COLOR_ARC),
        ] {
            for (gmt, pts) in &self.data.groups {
                if *gmt != mt || pts.len() < 2 {
                    continue;
                }
                out.push((pts.iter().map(|p| self.project(*p)).collect(), color));
            }
        }
        out
    }

    /// A real OS window: an `egui::Window` is clipped to the main window and cannot
    /// be dragged outside it. Backends without multiple viewports fall back to an
    /// embedded window, which is the previous behaviour.
    pub fn show(&mut self, ctx: &egui::Context) {
        let mut open = self.open;
        let title = format!("Path Preview \u{2014} {}", self.title);
        ctx.show_viewport_immediate(
            egui::ViewportId::from_hash_of("path_preview"),
            egui::ViewportBuilder::default()
                .with_title(title)
                .with_inner_size([960.0, 720.0])
                .with_resizable(true),
            |ui, _class| {
                // the OS title-bar X asks us to close: stop showing the viewport next pass
                if ui.input(|i| i.viewport().close_requested()) {
                    open = false;
                }
                egui::Frame::new()
                    .fill(theme::PANEL)
                    .inner_margin(10.0)
                    .show(ui, |ui| {
                        ui.horizontal(|ui| {
                            ui.label(RichText::new("PATH PREVIEW").strong());
                            let mut legend = |color: Color32, name: &str| {
                                ui.label(RichText::new("\u{2501}\u{2501}").color(color).monospace());
                                ui.label(name);
                            };
                            legend(COLOR_CUT, "CUT");
                            legend(COLOR_ARC, "ARC");
                            legend(COLOR_RAPID, "RAPID");
                            let b = &self.data.bounds;
                            ui.label(
                                RichText::new(format!(
                                    "   {} cut  {} arc  {} rapid   W:{:.1}  H:{:.1} mm",
                                    thousands(self.data.counts.cut),
                                    thousands(self.data.counts.arc),
                                    thousands(self.data.counts.rapid),
                                    b.max_x - b.min_x,
                                    b.max_y - b.min_y
                                ))
                                .monospace()
                                .color(theme::TEXT_WEAK),
                            );
                        });
                        ui.horizontal(|ui| {
                            if ui.button("FIT").clicked() {
                                self.last_rect = Rect::NOTHING;
                            }
                            ui.checkbox(&mut self.iso, "Isometric");
                        });

                        let avail = ui.available_size();
                        let (response, painter) = ui.allocate_painter(avail, Sense::click_and_drag());
                        let rect = response.rect;
                        painter.rect_filled(rect, 0.0, COLOR_BG);

                        if self.last_rect != rect {
                            self.fit(rect);
                        }

                        if response.dragged() {
                            let d = response.drag_delta();
                            self.ox += d.x;
                            self.oy += d.y;
                        }
                        if response.hovered() {
                            // python: wheel up = zoom in. egui's scroll delta is positive downward.
                            let scroll = ui.input(|i| i.smooth_scroll_delta.y);
                            if scroll != 0.0 {
                                if let Some(c) = response.hover_pos() {
                                    let factor = if scroll > 0.0 { 1.0 / 1.15 } else { 1.15 };
                                    self.zoom_at(c, factor);
                                }
                            }
                        }

                        if !self.iso {
                            self.draw_grid(&painter, rect);
                        }

                        // draw order: rapid -> cut -> arc (arc on top so circles show)
                        for (line, color) in self.scene() {
                            painter.line(line, Stroke::new(1.0, color));
                        }
                    });
            },
        );
        self.open = open;
    }

    fn draw_grid(&self, painter: &egui::Painter, rect: Rect) {
        let b = &self.data.bounds;
        let s = self.scale.max(1e-9);
        let raw_step = (50.0 / s) as f64;
        let magnitude = 10f64.powf(raw_step.max(1e-9).log10().floor());
        let mut grid_step = magnitude;
        for mult in [1.0, 2.0, 5.0, 10.0] {
            if magnitude * mult >= raw_step {
                grid_step = magnitude * mult;
                break;
            }
        }
        let text_color = theme::TEXT_FAINT;
        let mut gx = (b.min_x / grid_step).floor() * grid_step;
        while gx <= b.max_x + grid_step {
            let p = self.project((gx, 0.0, 0.0));
            if p.x >= rect.left() && p.x <= rect.right() {
                painter.line_segment(
                    [Pos2::new(p.x, rect.top()), Pos2::new(p.x, rect.bottom())],
                    Stroke::new(1.0, COLOR_GRID),
                );
                painter.text(
                    Pos2::new(p.x + 2.0, rect.bottom() - 12.0),
                    egui::Align2::LEFT_BOTTOM,
                    format!("{:.0}", gx),
                    egui::FontId::monospace(7.0),
                    text_color,
                );
            }
            gx += grid_step;
        }
        let mut gy = (b.min_y / grid_step).floor() * grid_step;
        while gy <= b.max_y + grid_step {
            let p = self.project((0.0, gy, 0.0));
            if p.y >= rect.top() && p.y <= rect.bottom() {
                painter.line_segment(
                    [Pos2::new(rect.left(), p.y), Pos2::new(rect.right(), p.y)],
                    Stroke::new(1.0, COLOR_GRID),
                );
                painter.text(
                    Pos2::new(rect.left() + 4.0, p.y - 2.0),
                    egui::Align2::LEFT_BOTTOM,
                    format!("{:.0}", gy),
                    egui::FontId::monospace(7.0),
                    text_color,
                );
            }
            gy += grid_step;
        }
        let o = self.project((0.0, 0.0, 0.0));
        painter.line_segment(
            [Pos2::new(o.x, rect.top()), Pos2::new(o.x, rect.bottom())],
            Stroke::new(1.0, COLOR_AXIS),
        );
        painter.line_segment(
            [Pos2::new(rect.left(), o.y), Pos2::new(rect.right(), o.y)],
            Stroke::new(1.0, COLOR_AXIS),
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gcode::{Bounds, Counts, GcodeParsed, Mv, Mt};

    fn sample() -> GcodeParsed {
        let pts = vec![
            Mv { x: 0.0, y: 0.0, z: 0.0, mt: Mt::Rapid },
            Mv { x: 100.0, y: 0.0, z: 0.0, mt: Mt::Cut },
            Mv { x: 100.0, y: 50.0, z: 0.0, mt: Mt::Cut },
        ];
        GcodeParsed {
            points: pts.clone(),
            groups: vec![(Mt::Cut, vec![(0.0, 0.0, 0.0), (100.0, 0.0, 0.0), (100.0, 50.0, 0.0)])],
            counts: Counts { rapid: 1, cut: 2, arc: 0 },
            bounds: Bounds { min_x: 0.0, max_x: 100.0, min_y: 0.0, max_y: 50.0 },
        }
    }

    #[test]
    fn fit_centers_and_scales_to_rect() {
        let mut p = Preview::new("t".into(), sample());
        let rect = Rect::from_min_size(Pos2::new(0.0, 0.0), egui::vec2(500.0, 200.0));
        p.fit(rect);
        // height-limited: 200 - 2*48 padding = 104 px for 50 world units
        assert!((p.scale - 104.0 / 50.0).abs() < 1e-3, "scale {}", p.scale);
        let c0 = p.project((50.0, 25.0, 0.0)); // world center
        let rc = rect.center();
        assert!((c0.x - rc.x).abs() < 1.0 && (c0.y - rc.y).abs() < 1.0, "{:?}", c0);
    }

    #[test]
    fn zoom_keeps_anchor_fixed() {
        let mut p = Preview::new("t".into(), sample());
        let rect = Rect::from_min_size(Pos2::new(0.0, 0.0), egui::vec2(500.0, 200.0));
        p.fit(rect);
        let anchor_world = (20.0, 40.0, 0.0);
        let before = p.project(anchor_world);
        p.zoom_at(before, 1.15);
        let after = p.project(anchor_world);
        assert!((after.x - before.x).abs() < 0.01 && (after.y - before.y).abs() < 0.01);
        p.zoom_at(before, 1.0 / 1.15);
        let back = p.project(anchor_world);
        assert!((back.x - before.x).abs() < 0.01 && (back.y - before.y).abs() < 0.01);
    }

    #[test]
    fn iso_projection_matches_python() {
        // python: px=(wx-wy)*cos30, py=(wx+wy)*sin30 - wz (origin-free)
        let mut p = Preview::new("t".into(), sample());
        p.iso = true;
        p.ox = 0.0;
        p.oy = 0.0;
        p.scale = 1.0;
        let c = 30f32.to_radians().cos();
        let s = 30f32.to_radians().sin();
        let q = p.project((100.0, 50.0, 10.0));
        assert!((q.x - (100.0 - 50.0) * c).abs() < 1e-3, "{}", q.x);
        assert!((q.y - ((100.0 + 50.0) * s - 10.0)).abs() < 1e-3, "{}", q.y);
    }

    #[test]
    fn thousands_groups_digits() {
        assert_eq!(thousands(0), "0");
        assert_eq!(thousands(999), "999");
        assert_eq!(thousands(11060), "11,060");
    }
}

// geometry checks run against the real golden G-code, not hand-made points
#[cfg(test)]
mod scene_tests {
    use super::*;
    use egui::Pos2;

    fn load(golden: &str) -> Preview {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/").to_string() + golden;
        let data = crate::gcode::parse_gcode(&path).expect("parse golden");
        let mut p = Preview::new("golden".into(), data);
        p.fit(Rect::from_min_size(
            Pos2::new(0.0, 0.0),
            egui::vec2(800.0, 600.0),
        ));
        p
    }

    #[test]
    fn topdown_golden_fits_inside_canvas() {
        let p = load("/sample_expected.g1.ngc");
        let lines = p.scene();
        assert!(lines.len() >= 2, "lines {}", lines.len());
        let n_pts: usize = lines.iter().map(|l| l.0.len()).sum();
        assert!(n_pts > 1000, "points {}", n_pts);
        for (line, _) in &lines {
            for pt in line {
                assert!(pt.x >= -1.0 && pt.x <= 801.0, "x {}", pt.x);
                assert!(pt.y >= -1.0 && pt.y <= 601.0, "y {}", pt.y);
            }
        }
    }

    #[test]
    fn arcs_golden_yields_arc_and_cut_polylines() {
        // g1 golden: only rapid+cut; arcs golden adds arc groups
        let g1 = load("/sample_expected.g1.ngc").scene();
        let arcs = load("/sample_expected.arcs.ngc").scene();
        let colors = |v: &Vec<(Vec<Pos2>, Color32)>| -> Vec<usize> {
            [COLOR_RAPID, COLOR_CUT, COLOR_ARC]
                .iter()
                .map(|c| v.iter().filter(|l| l.1 == *c).count())
                .collect()
        };
        let c1 = colors(&g1);
        let c2 = colors(&arcs);
        assert!(c1[2] == 0, "g1 has no arcs: {:?}", c1);
        assert!(c2[2] > 0, "arcs file must produce arc polylines: {:?}", c2);
        assert!(arcs.iter().map(|l| l.0.len()).sum::<usize>() < g1.iter().map(|l| l.0.len()).sum(),
            "arc file should have far fewer points than the linearised one");
    }
}
