// Toolpath preview scene data: the projection math python `open_preview()`
// used (isel_to_gcode_gui.py lines 216-490) survives here as plain data, so it
// stays testable headless and the preview window only paints what it is given.
// shortcut: the whole scene is rebuilt per view event (pan/zoom/fit) instead of
// keeping a camera in the front end; move projection client-side if panning
// ever stutters over IPC.

use crate::gcode::{GcodeParsed, Mt};
use serde::{Deserialize, Serialize};

// Same palette as the python reference preview: cyan = cutting, mint = arc,
// amber = rapid. Everything else recedes so the toolpath is the only loud
// thing on screen. Hex strings so the canvas can paint them directly.
pub const COLOR_CUT: &str = "#00C0E0";
pub const COLOR_RAPID: &str = "#F0A020";
pub const COLOR_ARC: &str = "#3FD0B0";
pub const COLOR_BG: &str = "#000E1A";
pub const COLOR_GRID: &str = "#10374F";
pub const COLOR_AXIS: &str = "#1C5A78";
const PADDING: f32 = 48.0;

/// Screen transform (python `project`'s ox/oy/scale), egui-free.
#[derive(Clone, Copy, Serialize, Deserialize)]
pub struct Camera {
    pub ox: f32,
    pub oy: f32,
    pub scale: f32,
}

#[derive(Serialize)]
pub struct SceneLine {
    pub points: Vec<[f32; 2]>,
    pub color: &'static str,
}

#[derive(Serialize)]
pub struct GridLine {
    pub x1: f32,
    pub y1: f32,
    pub x2: f32,
    pub y2: f32,
    pub color: &'static str,
}

#[derive(Serialize)]
pub struct GridLabel {
    pub x: f32,
    pub y: f32,
    pub text: String,
}

#[derive(Serialize)]
pub struct PreviewScene {
    pub lines: Vec<SceneLine>,
    /// Grid and the X/Y zero axes in draw order; empty in isometric view
    /// (python drew neither there either).
    pub grid: Vec<GridLine>,
    pub labels: Vec<GridLabel>,
    pub camera: Camera,
    pub background: &'static str,
}

/// world -> screen (python `project`).
fn project(cam: &Camera, iso: bool, w: (f64, f64, f64)) -> [f32; 2] {
    let (wx, wy, wz) = (w.0 as f32, w.1 as f32, w.2 as f32);
    if iso {
        let c = 30f32.to_radians().cos();
        let s = 30f32.to_radians().sin();
        [cam.ox + (wx - wy) * c * cam.scale, cam.oy + ((wx + wy) * s - wz) * cam.scale]
    } else {
        [cam.ox + wx * cam.scale, cam.oy - wy * cam.scale]
    }
}

/// Bounds of the projected shape, origin-free (python `projected_bounds`).
fn projected_bounds(data: &GcodeParsed, iso: bool) -> (f32, f32, f32, f32) {
    let (mut min_px, mut max_px) = (f32::INFINITY, f32::NEG_INFINITY);
    let (mut min_py, mut max_py) = (f32::INFINITY, f32::NEG_INFINITY);
    for p in &data.points {
        let (wx, wy, wz) = (p.x as f32, p.y as f32, p.z as f32);
        let (px, py) = if iso {
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

/// Fit the shape into a (0,0,w,h) viewport, centered (python `fit`).
/// Degenerate data keeps the old default camera, like Preview::new left it.
pub fn fit(data: &GcodeParsed, iso: bool, width: f32, height: f32) -> Camera {
    let (min_px, max_px, min_py, max_py) = projected_bounds(data, iso);
    if !min_px.is_finite() {
        return Camera { ox: 0.0, oy: 0.0, scale: 1.0 };
    }
    let span_x = (max_px - min_px).max(1e-6);
    let span_y = (max_py - min_py).max(1e-6);
    let s = ((width - 2.0 * PADDING) / span_x).min((height - 2.0 * PADDING) / span_y);
    Camera {
        ox: (width - span_x * s) / 2.0 - min_px * s,
        oy: (height - span_y * s) / 2.0 - min_py * s,
        scale: s,
    }
}

/// Projected polylines in draw order (rapid -> cut -> arc) plus the grid,
/// axes and labels python painted under them.
pub fn scene(
    data: &GcodeParsed,
    width: f32,
    height: f32,
    iso: bool,
    camera: Option<&Camera>,
) -> PreviewScene {
    let cam = camera.copied().unwrap_or_else(|| fit(data, iso, width, height));

    let mut lines = Vec::new();
    for (mt, color) in [(Mt::Rapid, COLOR_RAPID), (Mt::Cut, COLOR_CUT), (Mt::Arc, COLOR_ARC)] {
        for (gmt, pts) in &data.groups {
            if *gmt != mt || pts.len() < 2 {
                continue;
            }
            lines.push(SceneLine {
                points: pts.iter().map(|p| project(&cam, iso, *p)).collect(),
                color,
            });
        }
    }

    let mut grid = Vec::new();
    let mut labels = Vec::new();
    if !iso {
        draw_grid(data, &cam, width, height, &mut grid, &mut labels);
    }

    PreviewScene {
        lines,
        grid,
        labels,
        camera: cam,
        background: COLOR_BG,
    }
}

/// Grid step, lines and labels from the old `Preview::draw_grid`, returned as
/// data instead of painted. The zero axes come back as grid lines too.
fn draw_grid(
    data: &GcodeParsed,
    cam: &Camera,
    width: f32,
    height: f32,
    grid: &mut Vec<GridLine>,
    labels: &mut Vec<GridLabel>,
) {
    let b = &data.bounds;
    let s = cam.scale.max(1e-9);
    let raw_step = (50.0 / s) as f64;
    let magnitude = 10f64.powf(raw_step.max(1e-9).log10().floor());
    let mut grid_step = magnitude;
    for mult in [1.0, 2.0, 5.0, 10.0] {
        if magnitude * mult >= raw_step {
            grid_step = magnitude * mult;
            break;
        }
    }
    let mut gx = (b.min_x / grid_step).floor() * grid_step;
    while gx <= b.max_x + grid_step {
        let p = project(cam, false, (gx, 0.0, 0.0));
        if p[0] >= 0.0 && p[0] <= width {
            grid.push(GridLine { x1: p[0], y1: 0.0, x2: p[0], y2: height, color: COLOR_GRID });
            labels.push(GridLabel { x: p[0] + 2.0, y: height - 12.0, text: format!("{:.0}", gx) });
        }
        gx += grid_step;
    }
    let mut gy = (b.min_y / grid_step).floor() * grid_step;
    while gy <= b.max_y + grid_step {
        let p = project(cam, false, (0.0, gy, 0.0));
        if p[1] >= 0.0 && p[1] <= height {
            grid.push(GridLine { x1: 0.0, y1: p[1], x2: width, y2: p[1], color: COLOR_GRID });
            labels.push(GridLabel { x: 4.0, y: p[1] - 2.0, text: format!("{:.0}", gy) });
        }
        gy += grid_step;
    }
    let o = project(cam, false, (0.0, 0.0, 0.0));
    grid.push(GridLine { x1: o[0], y1: 0.0, x2: o[0], y2: height, color: COLOR_AXIS });
    grid.push(GridLine { x1: 0.0, y1: o[1], x2: width, y2: o[1], color: COLOR_AXIS });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gcode::{Bounds, Counts, Mv, Mt};

    fn sample() -> GcodeParsed {
        let pts = vec![
            Mv { x: 0.0, y: 0.0, z: 0.0, mt: Mt::Rapid },
            Mv { x: 100.0, y: 0.0, z: 0.0, mt: Mt::Cut },
            Mv { x: 100.0, y: 50.0, z: 0.0, mt: Mt::Cut },
        ];
        GcodeParsed {
            points: pts,
            groups: vec![(Mt::Cut, vec![(0.0, 0.0, 0.0), (100.0, 0.0, 0.0), (100.0, 50.0, 0.0)])],
            counts: Counts { rapid: 1, cut: 2, arc: 0 },
            bounds: Bounds { min_x: 0.0, max_x: 100.0, min_y: 0.0, max_y: 50.0 },
        }
    }

    #[test]
    fn fit_centers_and_scales_to_rect() {
        // port of preview.rs's test of the same name (egui Rect -> plain w/h)
        let cam = fit(&sample(), false, 500.0, 200.0);
        // height-limited: 200 - 2*48 padding = 104 px for 50 world units
        assert!((cam.scale - 104.0 / 50.0).abs() < 1e-3, "scale {}", cam.scale);
        let c0 = project(&cam, false, (50.0, 25.0, 0.0)); // world center
        assert!((c0[0] - 250.0).abs() < 1.0 && (c0[1] - 100.0).abs() < 1.0, "{:?}", c0);
    }

    #[test]
    fn iso_projection_matches_python() {
        // python: px=(wx-wy)*cos30, py=(wx+wy)*sin30 - wz (origin-free)
        let cam = Camera { ox: 0.0, oy: 0.0, scale: 1.0 };
        let c = 30f32.to_radians().cos();
        let s = 30f32.to_radians().sin();
        let q = project(&cam, true, (100.0, 50.0, 10.0));
        assert!((q[0] - (100.0 - 50.0) * c).abs() < 1e-3, "{}", q[0]);
        assert!((q[1] - ((100.0 + 50.0) * s - 10.0)).abs() < 1e-3, "{}", q[1]);
    }

    #[test]
    fn zoom_anchor_math_mirror_keeps_world_point_under_cursor() {
        // The front end implements pan/zoom client-side with the same anchor
        // math the old zoom_at used (ox' = c.x + (ox-c.x)*f). This test pins
        // that formula against the projection it drives, so the JS mirror has
        // a Rust-verified reference.
        let cam = fit(&sample(), false, 500.0, 200.0);
        let anchor_world = (20.0, 40.0, 0.0);
        let before = project(&cam, false, anchor_world);
        let f = 1.15f32;
        let zoomed = Camera {
            ox: before[0] + (cam.ox - before[0]) * f,
            oy: before[1] + (cam.oy - before[1]) * f,
            scale: cam.scale * f,
        };
        let after = project(&zoomed, false, anchor_world);
        assert!((after[0] - before[0]).abs() < 0.01 && (after[1] - before[1]).abs() < 0.01);
    }
}

// geometry checks run against the real golden G-code, not hand-made points
#[cfg(test)]
mod scene_tests {
    use super::*;

    fn load(golden: &str) -> GcodeParsed {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/").to_string() + golden;
        crate::gcode::parse_gcode(&path).expect("parse golden")
    }

    #[test]
    fn topdown_golden_fits_inside_canvas() {
        let data = load("/sample_expected.g1.ngc");
        let sc = scene(&data, 800.0, 600.0, false, None);
        assert!(sc.lines.len() >= 2, "lines {}", sc.lines.len());
        let n_pts: usize = sc.lines.iter().map(|l| l.points.len()).sum();
        assert!(n_pts > 1000, "points {}", n_pts);
        for line in &sc.lines {
            for pt in &line.points {
                assert!(pt[0] >= -1.0 && pt[0] <= 801.0, "x {}", pt[0]);
                assert!(pt[1] >= -1.0 && pt[1] <= 601.0, "y {}", pt[1]);
            }
        }
    }

    #[test]
    fn arcs_golden_yields_arc_and_cut_polylines() {
        // g1 golden: only rapid+cut; arcs golden adds arc groups
        let g1 = scene(&load("/sample_expected.g1.ngc"), 800.0, 600.0, false, None);
        let arcs = scene(&load("/sample_expected.arcs.ngc"), 800.0, 600.0, false, None);
        let count = |v: &PreviewScene, color: &str| -> usize {
            v.lines.iter().filter(|l| l.color == color).count()
        };
        assert_eq!(count(&g1, COLOR_ARC), 0, "g1 has no arcs");
        assert!(count(&arcs, COLOR_ARC) > 0, "arcs file must produce arc polylines");
        let n = |v: &PreviewScene| v.lines.iter().map(|l| l.points.len()).sum::<usize>();
        assert!(n(&arcs) < n(&g1), "arc file should have far fewer points than the linearised one");
    }
}
