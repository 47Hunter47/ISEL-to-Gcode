// Toolpath preview scene data: the projection math python `open_preview()`
// used (isel_to_gcode_gui.py lines 216-490) survives here as plain data, so it
// stays testable headless and the preview window only paints what it is given.
// The scene is rebuilt when the file, view mode, or viewport changes (and on
// Fit); pan/zoom are a canvas transform in the front end over the projected
// scene, so a drag or wheel tick never crosses IPC.

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

/// View direction: yaw about Z and tilt out of the XY plane, in degrees.
/// One formula family covers both python presets exactly: tilt -90 is the top
/// view byte-for-byte, (45, atan(1/sqrt(2))) is python's isometric up to the
/// uniform 1/cos(tilt) = 1.224745 that `fit` absorbs into scale. A middle-drag
/// orbit interpolates freely between and beyond them (front end owns the drag;
/// each landed angle re-fits, like the old iso checkbox).
#[derive(Clone, Copy, Serialize, Deserialize)]
pub struct ViewDir {
    pub yaw_deg: f32,
    pub tilt_deg: f32,
}

pub const VIEW_TOP: ViewDir = ViewDir { yaw_deg: 0.0, tilt_deg: -90.0 };
pub const VIEW_ISO: ViewDir = ViewDir { yaw_deg: 45.0, tilt_deg: 35.264389682751755 };

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
    /// Grid and the X/Y zero axes in draw order; empty in any tilted view
    /// (python drew neither in iso, and a tilted grid plane would be a guess).
    pub grid: Vec<GridLine>,
    pub labels: Vec<GridLabel>,
    pub camera: Camera,
    pub background: &'static str,
}

/// Orthographic basis rows (sx_x, sx_y, sy_x, sy_y, sy_z): screen x/y per
/// world axis. At the pinned top preset use exact -1/0 — f32 cos(-90deg) is a
/// 4e-8 residue that would give wz a weight the python top view never had.
fn view_basis(v: &ViewDir) -> (f32, f32, f32, f32, f32) {
    let (ya, ti) = (v.yaw_deg.to_radians(), v.tilt_deg.to_radians());
    let (syaw, cyaw) = (ya.sin(), ya.cos());
    let (st, ct) = if v.tilt_deg == -90.0 { (-1.0f32, 0.0f32) } else { (ti.sin(), ti.cos()) };
    (cyaw, -syaw, syaw * st, cyaw * st, -ct)
}

/// world -> screen origin-free (the python formula's scale-free core).
fn project_raw(b: &(f32, f32, f32, f32, f32), w: (f64, f64, f64)) -> [f32; 2] {
    let (wx, wy, wz) = (w.0 as f32, w.1 as f32, w.2 as f32);
    [wx * b.0 + wy * b.1, wx * b.2 + wy * b.3 + wz * b.4]
}

/// world -> screen (python `project`).
fn project(cam: &Camera, v: &ViewDir, w: (f64, f64, f64)) -> [f32; 2] {
    let r = project_raw(&view_basis(v), w);
    [cam.ox + r[0] * cam.scale, cam.oy + r[1] * cam.scale]
}

/// Bounds of the projected shape, origin-free (python `projected_bounds`).
/// Every point lives in `data.groups` exactly once, so one walk covers all.
fn projected_bounds(data: &GcodeParsed, v: &ViewDir) -> (f32, f32, f32, f32) {
    let b = view_basis(v);
    let (mut min_px, mut max_px) = (f32::INFINITY, f32::NEG_INFINITY);
    let (mut min_py, mut max_py) = (f32::INFINITY, f32::NEG_INFINITY);
    for (_, pts) in &data.groups {
        for w in pts {
            let p = project_raw(&b, *w);
            min_px = min_px.min(p[0]);
            max_px = max_px.max(p[0]);
            min_py = min_py.min(p[1]);
            max_py = max_py.max(p[1]);
        }
    }
    (min_px, max_px, min_py, max_py)
}

/// Fit the shape into a (0,0,w,h) viewport, centered (python `fit`).
/// Degenerate data keeps the old default camera, like Preview::new left it.
pub fn fit(data: &GcodeParsed, v: &ViewDir, width: f32, height: f32) -> Camera {
    let (min_px, max_px, min_py, max_py) = projected_bounds(data, v);
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
    view: &ViewDir,
    camera: Option<&Camera>,
) -> PreviewScene {
    let cam = camera.copied().unwrap_or_else(|| fit(data, view, width, height));

    let mut lines = Vec::new();
    for (mt, color) in [(Mt::Rapid, COLOR_RAPID), (Mt::Cut, COLOR_CUT), (Mt::Arc, COLOR_ARC)] {
        for (gmt, pts) in &data.groups {
            if *gmt != mt || pts.len() < 2 {
                continue;
            }
            lines.push(SceneLine {
                points: pts.iter().map(|p| project(&cam, view, *p)).collect(),
                color,
            });
        }
    }

    let mut grid = Vec::new();
    let mut labels = Vec::new();
    if view.tilt_deg == -90.0 {
        draw_grid(data, &cam, view, width, height, &mut grid, &mut labels);
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
/// The step is grown a second time if needed so the total line count is
/// bounded: at extreme zoom the 50px target step iterates billions of times.
const GRID_LINE_CAP: f64 = 500.0; // per axis; the toolpath, not the grid, is the point

fn draw_grid(
    data: &GcodeParsed,
    cam: &Camera,
    view: &ViewDir,
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
    // cap: lines per axis is span/step, so a floor on the step bounds the loop
    // (NaN never arrives: bounds and step are finite once scale is finite)
    let span_x = (b.max_x - b.min_x).max(0.0);
    let span_y = (b.max_y - b.min_y).max(0.0);
    grid_step = grid_step.max(span_x / GRID_LINE_CAP).max(span_y / GRID_LINE_CAP);

    // Top-view basis (wz never enters; the grid plane lives at z=0) and a
    // world length whose projection always spans the viewport, so each grid
    // line is one full projected line — the old viewport-aligned stubs only
    // worked while yaw was 0 and are wrong for an orbited top view.
    let bs = view_basis(view);
    let pt = |wx: f64, wy: f64| {
        [cam.ox + (wx as f32 * bs.0 + wy as f32 * bs.1) * s,
         cam.oy + (wx as f32 * bs.2 + wy as f32 * bs.3) * s]
    };
    let reach = ((width + height) / s) as f64;
    // label anchor: where the line crosses the canvas bottom (gx) or left
    // edge (gy), interpolated - reproduces the python label placement exactly
    // at yaw 0 and stays correct as the view rotates.
    let bottom_label = |a: [f32; 2], bb: [f32; 2]| -> Option<[f32; 2]> {
        let (dx, dy) = (bb[0] - a[0], bb[1] - a[1]);
        let lo = if a[1] <= bb[1] { (a, bb) } else { (bb, a) };
        if lo.1[1] < height && (dy.abs() < 1e-9 || lo.1[1] == height) {
            return Some([lo.1[0], lo.1[1]]); // whole line above the bottom edge
        }
        if dy.abs() < 1e-9 || lo.0[1] > height {
            return None;
        }
        let t = (height - lo.0[1]) / dy.abs();
        Some([lo.0[0] + dx.signum() * dx.abs() * t.min(1.0), height])
    };
    let left_label = |a: [f32; 2], bb: [f32; 2]| -> Option<[f32; 2]> {
        let (dx, dy) = (bb[0] - a[0], bb[1] - a[1]);
        let lo = if a[0] <= bb[0] { (a, bb) } else { (bb, a) };
        if lo.1[0] > 0.0 && (dx.abs() < 1e-9 || lo.1[0] == 0.0) {
            return Some([lo.1[0], lo.1[1]]); // whole line right of the left edge
        }
        if dx.abs() < 1e-9 || lo.0[0] > 0.0 {
            return None;
        }
        let t = (0.0 - lo.0[0]) / dx.abs();
        Some([0.0, lo.0[1] + dy * t.min(1.0)])
    };

    let cy_mid = (b.min_y + b.max_y) / 2.0;
    let mut gx = (b.min_x / grid_step).floor() * grid_step;
    while gx <= b.max_x + grid_step {
        let (a, bb) = (pt(gx, cy_mid - reach), pt(gx, cy_mid + reach));
        grid.push(GridLine { x1: a[0], y1: a[1], x2: bb[0], y2: bb[1], color: COLOR_GRID });
        if let Some(p) = bottom_label(a, bb) {
            if p[0] >= 0.0 && p[0] <= width {
                labels.push(GridLabel { x: p[0] + 2.0, y: p[1] - 12.0, text: format!("{:.0}", gx) });
            }
        }
        gx += grid_step;
    }
    let cx_mid = (b.min_x + b.max_x) / 2.0;
    let mut gy = (b.min_y / grid_step).floor() * grid_step;
    while gy <= b.max_y + grid_step {
        let (a, bb) = (pt(cx_mid - reach, gy), pt(cx_mid + reach, gy));
        grid.push(GridLine { x1: a[0], y1: a[1], x2: bb[0], y2: bb[1], color: COLOR_GRID });
        if let Some(p) = left_label(a, bb) {
            if p[1] >= 0.0 && p[1] <= height {
                labels.push(GridLabel { x: p[0] + 4.0, y: p[1] - 2.0, text: format!("{:.0}", gy) });
            }
        }
        gy += grid_step;
    }
    let (ax0, ax1) = (pt(-reach, 0.0), pt(reach, 0.0));
    grid.push(GridLine { x1: ax0[0], y1: ax0[1], x2: ax1[0], y2: ax1[1], color: COLOR_AXIS });
    let (ay0, ay1) = (pt(0.0, -reach), pt(0.0, reach));
    grid.push(GridLine { x1: ay0[0], y1: ay0[1], x2: ay1[0], y2: ay1[1], color: COLOR_AXIS });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gcode::{Bounds, Counts, Mt};

    fn sample() -> GcodeParsed {
        GcodeParsed {
            groups: vec![
                (Mt::Rapid, vec![(0.0, 0.0, 0.0), (0.0, 0.0, 0.0)]),
                (Mt::Cut, vec![(0.0, 0.0, 0.0), (100.0, 0.0, 0.0), (100.0, 50.0, 0.0)]),
            ],
            counts: Counts { rapid: 1, cut: 2, arc: 0 },
            bounds: Bounds { min_x: 0.0, max_x: 100.0, min_y: 0.0, max_y: 50.0 },
        }
    }

    #[test]
    fn fit_centers_and_scales_to_rect() {
        // port of preview.rs's test of the same name (egui Rect -> plain w/h)
        let cam = fit(&sample(), &VIEW_TOP, 500.0, 200.0);
        // height-limited: 200 - 2*48 padding = 104 px for 50 world units
        assert!((cam.scale - 104.0 / 50.0).abs() < 1e-3, "scale {}", cam.scale);
        let c0 = project(&cam, &VIEW_TOP, (50.0, 25.0, 0.0)); // world center
        assert!((c0[0] - 250.0).abs() < 1.0 && (c0[1] - 100.0).abs() < 1.0, "{:?}", c0);
    }

    #[test]
    fn top_preset_is_the_python_formula_exactly() {
        // pinned preset: bit-for-bit python top view; wz must carry no weight
        // (the exact -1/0 basis stop, not f32 cos(-90deg)'s 4e-8 residue)
        let cam = Camera { ox: 7.0, oy: 9.0, scale: 2.0 };
        assert_eq!(project(&cam, &VIEW_TOP, (1.5, -2.0, 12345.0)), [7.0 + 1.5 * 2.0, 9.0 + 2.0 * 2.0]);
        assert_eq!(project(&cam, &VIEW_TOP, (1.5, -2.0, -999.0)), [10.0, 13.0]);
    }

    #[test]
    fn iso_basis_is_python_up_to_uniform_scale() {
        // python's iso (px=(wx-wy)cos30, py=(wx+wy)sin30-wz) is this exact
        // orthographic family at VIEW_ISO, uniformly scaled by
        // 1/cos(atan(1/sqrt(2))) = 1.2247448; fit() absorbs that into
        // cam.scale, so the rendered iso view is unchanged. The axis chirality
        // asserts guard against a mirror flip: python puts +Z up-screen and
        // world +(x+y) down-screen.
        let cam = Camera { ox: 0.0, oy: 0.0, scale: 1.0 };
        let k = 1.2247448f32;
        let q = project(&cam, &VIEW_ISO, (100.0, 50.0, 10.0));
        assert!((q[0] * k - (100.0 - 50.0) * 30f32.to_radians().cos()).abs() < 1e-2, "{}", q[0]);
        assert!((q[1] * k - ((100.0 + 50.0) * 30f32.to_radians().sin() - 10.0)).abs() < 1e-2, "{}", q[1]);
        let b = view_basis(&VIEW_ISO);
        let ax = project_raw(&b, (1.0, 0.0, 0.0));
        let ay = project_raw(&b, (0.0, 1.0, 0.0));
        let az = project_raw(&b, (0.0, 0.0, 1.0));
        let len = |p: [f32; 2]| p[0].hypot(p[1]);
        assert!((len(ax) - len(ay)).abs() < 1e-4 && (len(ay) - len(az)).abs() < 1e-4,
            "iso must foreshorten all three axes equally: {} {} {}", len(ax), len(ay), len(az));
        assert!(ax[0] > 0.0 && ax[1] > 0.0 && ay[0] < 0.0 && ay[1] > 0.0 && az[1] < 0.0,
            "{:?} {:?} {:?}", ax, ay, az);
    }

    #[test]
    fn zoom_anchor_math_mirror_keeps_world_point_under_cursor() {
        // The front end implements pan/zoom client-side with the same anchor
        // math the old zoom_at used (ox' = c.x + (ox-c.x)*f). This test pins
        // that formula against the projection it drives, so the JS mirror has
        // a Rust-verified reference.
        let cam = fit(&sample(), &VIEW_TOP, 500.0, 200.0);
        let anchor_world = (20.0, 40.0, 0.0);
        let before = project(&cam, &VIEW_TOP, anchor_world);
        let f = 1.15f32;
        let zoomed = Camera {
            ox: before[0] + (cam.ox - before[0]) * f,
            oy: before[1] + (cam.oy - before[1]) * f,
            scale: cam.scale * f,
        };
        let after = project(&zoomed, &VIEW_TOP, anchor_world);
        assert!((after[0] - before[0]).abs() < 0.01 && (after[1] - before[1]).abs() < 0.01);
    }

    #[test]
    fn extreme_zoom_and_degenerate_viewports_keep_grid_work_bounded() {
        // 1e9 px/mm over a 100mm shape: the 50px step target alone would ask
        // for ~10^10 vertical lines. The cap must grow the step instead.
        let cam = Camera { ox: 0.0, oy: 0.0, scale: 1e9 };
        let sc = scene(&sample(), 800.0, 600.0, &VIEW_TOP, Some(&cam));
        assert!(sc.grid.len() <= 2 * 500 + 2, "grid {}", sc.grid.len());
        assert!(sc.labels.len() <= 2 * 500 + 2, "labels {}", sc.labels.len());
        // tiny viewport on the same data: no panic, still bounded
        let sc1 = scene(&sample(), 1.0, 1.0, &VIEW_TOP, Some(&cam));
        assert!(sc1.grid.len() <= 2 * 500 + 2);
        // zoom far out: step explodes instead of the count; axes are always drawn
        let tiny = Camera { ox: 0.0, oy: 0.0, scale: 1e-9 };
        let sc2 = scene(&sample(), 800.0, 600.0, &VIEW_TOP, Some(&tiny));
        assert!(!sc2.grid.is_empty() && sc2.grid.len() <= 2 * 500 + 2, "grid {}", sc2.grid.len());
    }

    #[test]
    fn normal_zoom_grid_is_unchanged() {
        // the cap may only bite at extremes: at a fitted view the step stays
        // the python 50px target (demo2 fits ~18 lines/axis into an 800px view)
        let data = crate::gcode::parse_gcode(
            &(concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/").to_string() + "/demo2_expected.g1.ngc"))
            .expect("parse golden");
        let cam = fit(&data, &VIEW_TOP, 800.0, 600.0);
        let sc = scene(&data, 800.0, 600.0, &VIEW_TOP, Some(&cam));
        let lines: usize = sc.grid.iter().filter(|g| g.color == COLOR_GRID).count();
        assert!(lines > 0 && lines < 100, "grid lines {}", lines);
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
        let sc = scene(&data, 800.0, 600.0, &VIEW_TOP, None);
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
        let g1 = scene(&load("/sample_expected.g1.ngc"), 800.0, 600.0, &VIEW_TOP, None);
        let arcs = scene(&load("/sample_expected.arcs.ngc"), 800.0, 600.0, &VIEW_TOP, None);
        let count = |v: &PreviewScene, color: &str| -> usize {
            v.lines.iter().filter(|l| l.color == color).count()
        };
        assert_eq!(count(&g1, COLOR_ARC), 0, "g1 has no arcs");
        assert!(count(&arcs, COLOR_ARC) > 0, "arcs file must produce arc polylines");
        let n = |v: &PreviewScene| v.lines.iter().map(|l| l.points.len()).sum::<usize>();
        assert!(n(&arcs) < n(&g1), "arc file should have far fewer points than the linearised one");
    }
}
