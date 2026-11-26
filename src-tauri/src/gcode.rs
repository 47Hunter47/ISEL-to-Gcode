// Preview support: parse output G-code into drawable polylines.
// Port of python open_preview() parsing (isel_to_gcode_gui.py lines 224-302) +
// _arc_to_polyline (lines 139-213). Semantics pinned by tests/gcode_unit.rs.

use regex::Regex;
use std::sync::OnceLock;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Mt { Rapid, Cut, Arc }

#[derive(Clone, Copy)]
pub struct Mv { pub x: f64, pub y: f64, pub z: f64, pub mt: Mt }

#[derive(Default)]
pub struct Counts { pub rapid: usize, pub cut: usize, pub arc: usize }

#[derive(Default)]
pub struct Bounds { pub min_x: f64, pub max_x: f64, pub min_y: f64, pub max_y: f64 }

pub struct GcodeParsed {
    pub points: Vec<Mv>,
    pub groups: Vec<(Mt, Vec<(f64, f64, f64)>)>,
    pub counts: Counts,
    pub bounds: Bounds,
}

fn re(pat: &str) -> &'static Regex {
    static CACHE: OnceLock<std::collections::HashMap<&'static str, Regex>> = OnceLock::new();
    CACHE
        .get_or_init(|| {
            [
                r"^N\d+\s*",
                r"X(-?\d+(?:\.\d+)?)",
                r"Y(-?\d+(?:\.\d+)?)",
                r"Z(-?\d+(?:\.\d+)?)",
                r"R(-?\d+(?:\.\d+)?)",
            ]
            .iter()
            .map(|p| (*p, Regex::new(p).unwrap()))
            .collect()
        })
        .get(pat)
        .unwrap_or_else(|| panic!("regex not preloaded: {}", pat))
}

/// Port of python `_arc_to_polyline`; behavior pinned against reference output.
pub fn arc_to_polyline(x0: f64, y0: f64, x1: f64, y1: f64, r_signed: f64, is_g2: bool) -> Vec<(f64, f64)> {
    let steps: usize = 36;
    let r = r_signed.abs();
    if r < 1e-9 { return vec![(x0, y0), (x1, y1)]; }
    let dx = x1 - x0;
    let dy = y1 - y0;
    let d = dx.hypot(dy);
    if d < 1e-9 { return vec![(x0, y0), (x1, y1)]; } // full circle guard (python line 152)

    let mx = (x0 + x1) / 2.0;
    let my = (y0 + y1) / 2.0;
    let mut h2 = r * r - (d / 2.0) * (d / 2.0);
    if h2 < 0.0 { h2 = 0.0; }
    let h = h2.sqrt();
    let px = -dy / d;
    let py = dx / d;
    let cx1 = mx + h * px; let cy1 = my + h * py;
    let cx2 = mx - h * px; let cy2 = my - h * py;

    let c1_cross = dx * (cy1 - y0) - dy * (cx1 - x0);
    let (mut cx, mut cy) = if is_g2 {
        if c1_cross > 0.0 { (cx1, cy1) } else { (cx2, cy2) }
    } else {
        if c1_cross < 0.0 { (cx1, cy1) } else { (cx2, cy2) }
    };
    if r_signed < 0.0 {
        // major arc: take the OTHER center (python lines 196-198, float equality intended)
        if cx == cx1 { cx = cx2; cy = cy2; } else { cx = cx1; cy = cy1; }
    }

    let a0 = (y0 - cy).atan2(x0 - cx);
    let mut a1 = (y1 - cy).atan2(x1 - cx);
    if is_g2 { if a1 >= a0 { a1 -= 2.0 * std::f64::consts::PI; } }
    else { if a1 <= a0 { a1 += 2.0 * std::f64::consts::PI; } }

    let mut pts = Vec::with_capacity(steps + 1);
    for i in 0..=steps {
        let t = i as f64 / steps as f64;
        let a = a0 + (a1 - a0) * t;
        pts.push((cx + r * a.cos(), cy + r * a.sin()));
    }
    pts
}

pub fn parse_gcode(path: &str) -> Result<GcodeParsed, String> {
    let raw = std::fs::read_to_string(path)
        .map_err(|e| format!("Could not read G-code: {}", e))?;
    let mut points: Vec<Mv> = vec![];
    let (mut cx, mut cy, mut cz) = (0.0, 0.0, 0.0);
    for line in raw.lines() {
        let line = line.trim();
        if line.is_empty() { continue; }
        let ln = re(r"^N\d+\s*").replace(line, "").into_owned();
        if ln.is_empty() { continue; }
        let is_g0 = ln.starts_with("G0");
        let is_g1 = ln.starts_with("G1");
        let is_g2 = ln.starts_with("G2");
        let is_g3 = ln.starts_with("G3");
        if !(is_g0 || is_g1 || is_g2 || is_g3) { continue; }
        let num = |pat: &'static str| -> Option<f64> {
            re(pat).captures(&ln).and_then(|c| c[1].parse::<f64>().ok())
        };
        let nx = num(r"X(-?\d+(?:\.\d+)?)").unwrap_or(cx);
        let ny = num(r"Y(-?\d+(?:\.\d+)?)").unwrap_or(cy);
        let nz = num(r"Z(-?\d+(?:\.\d+)?)").unwrap_or(cz);
        let r_val = num(r"R(-?\d+(?:\.\d+)?)");
        if is_g0 {
            points.push(Mv { x: cx, y: cy, z: cz, mt: Mt::Rapid });
            points.push(Mv { x: nx, y: ny, z: nz, mt: Mt::Rapid });
        } else if is_g1 {
            points.push(Mv { x: cx, y: cy, z: cz, mt: Mt::Cut });
            points.push(Mv { x: nx, y: ny, z: nz, mt: Mt::Cut });
        } else if (is_g2 || is_g3) && r_val.is_some() {
            for (ax, ay) in arc_to_polyline(cx, cy, nx, ny, r_val.unwrap(), is_g2) {
                points.push(Mv { x: ax, y: ay, z: cz, mt: Mt::Arc });
            }
        } else {
            points.push(Mv { x: cx, y: cy, z: cz, mt: Mt::Cut });
            points.push(Mv { x: nx, y: ny, z: nz, mt: Mt::Cut });
        }
        cx = nx; cy = ny; cz = nz;
    }

    let mut counts = Counts::default();
    for p in &points {
        match p.mt { Mt::Rapid => counts.rapid += 1, Mt::Cut => counts.cut += 1, Mt::Arc => counts.arc += 1 }
    }
    let mut groups: Vec<(Mt, Vec<(f64, f64, f64)>)> = vec![];
    for p in &points {
        // python build_groups: consecutive runs of the same move type
        if groups.last().map(|g| g.0 != p.mt).unwrap_or(true) {
            groups.push((p.mt, vec![]));
        }
        groups.last_mut().unwrap().1.push((p.x, p.y, p.z));
    }
    let bounds = if points.is_empty() {
        Bounds::default()
    } else {
        let (mut min_x, mut max_x, mut min_y, mut max_y) = (f64::INFINITY, f64::NEG_INFINITY, f64::INFINITY, f64::NEG_INFINITY);
        for p in &points {
            if p.x < min_x { min_x = p.x; }
            if p.x > max_x { max_x = p.x; }
            if p.y < min_y { min_y = p.y; }
            if p.y > max_y { max_y = p.y; }
        }
        Bounds { min_x, max_x, min_y, max_y }
    };
    Ok(GcodeParsed { points, groups, counts, bounds })
}

