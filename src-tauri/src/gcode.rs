// Preview support: parse output G-code into drawable polylines.
// Port of python open_preview() parsing (isel_to_gcode_gui.py lines 224-302) +
// _arc_to_polyline (lines 139-213). Semantics pinned by tests/gcode_unit.rs.

use regex::Regex;
use std::sync::OnceLock;

#[derive(Clone, Copy, PartialEq, Eq, Debug, serde::Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Mt { Rapid, Cut, Arc }

#[derive(Default, Clone, Copy, Debug, serde::Serialize)]
pub struct Counts { pub rapid: usize, pub cut: usize, pub arc: usize }

#[derive(Default, Clone, Copy, Debug, serde::Serialize)]
pub struct Bounds { pub min_x: f64, pub max_x: f64, pub min_y: f64, pub max_y: f64 }

/// Every drawable point lives exactly once: inside `groups`, as consecutive
/// runs of the same move type (python build_groups). Counts and bounds are
/// accumulated while parsing, so nothing walks the geometry a second time.
#[derive(Clone, serde::Serialize)]
pub struct GcodeParsed {
    pub groups: Vec<(Mt, Vec<(f64, f64, f64)>)>,
    pub counts: Counts,
    pub bounds: Bounds,
}

impl GcodeParsed {
    pub fn has_moves(&self) -> bool { !self.groups.is_empty() }
}

fn re(pat: &str) -> &'static Regex {
    static CACHE: OnceLock<std::collections::HashMap<&'static str, Regex>> = OnceLock::new();
    CACHE
        .get_or_init(|| {
            [
                r"^N\d+\s*",
                r";.*",
                r"\([^)]*\)",
                r"^G(\d+)(.?)",
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
    let mut groups: Vec<(Mt, Vec<(f64, f64, f64)>)> = vec![];
    let mut counts = Counts::default();
    let (mut min_x, mut max_x, mut min_y, mut max_y) = (f64::INFINITY, f64::NEG_INFINITY, f64::INFINITY, f64::NEG_INFINITY);
    let mut any = false;
    // python build_groups, folded into the parse: consecutive runs of the same
    // move type; counts and bounds ride along so the geometry is never re-walked.
    let mut push = |mt: Mt, x: f64, y: f64, z: f64| {
        if groups.last().map(|g| g.0 != mt).unwrap_or(true) {
            groups.push((mt, vec![]));
        }
        groups.last_mut().unwrap().1.push((x, y, z));
        match mt { Mt::Rapid => counts.rapid += 1, Mt::Cut => counts.cut += 1, Mt::Arc => counts.arc += 1 }
        if x < min_x { min_x = x; }
        if x > max_x { max_x = x; }
        if y < min_y { min_y = y; }
        if y > max_y { max_y = y; }
        any = true;
    };
    let (mut cx, mut cy, mut cz) = (0.0, 0.0, 0.0);
    for line in raw.lines() {
        let line = line.trim();
        if line.is_empty() { continue; }
        // Strip comments before anything else: a `;` tail or `(...)` span can
        // carry X/Y/Z words and G-words that are prose, not program (no nesting
        // supported — none in converter output; upgrade if a real input needs it).
        let line = re(r";.*").replace_all(line, "").into_owned();
        let line = re(r"\([^)]*\)").replace_all(&line, "").into_owned();
        let ln = re(r"^N\d+\s*").replace(&line, "").into_owned();
        if ln.is_empty() { continue; }
        // Exact G-word, not a prefix: G17/G20/G21/G40/G90 are not motion, and
        // G01/G02/G03 are G1/G2/G3. \d+ is greedy, so what follows the digits
        // is the first non-digit char: a letter means the word ended (G1X10
        // still parses), a dot is a fractional non-motion code (G0.1, G91.1).
        let Some(g) = re(r"^G(\d+)(.?)").captures(&ln) else { continue };
        if &g[2] == "." { continue; }
        let Ok(code) = g[1].parse::<u32>() else { continue };
        let (is_g0, is_g1, is_g2, is_g3) = match code {
            0 => (true, false, false, false),
            1 => (false, true, false, false),
            2 => (false, false, true, false),
            3 => (false, false, false, true),
            _ => continue,
        };
        let num = |pat: &'static str| -> Option<f64> {
            re(pat).captures(&ln).and_then(|c| c[1].parse::<f64>().ok())
        };
        let nx = num(r"X(-?\d+(?:\.\d+)?)").unwrap_or(cx);
        let ny = num(r"Y(-?\d+(?:\.\d+)?)").unwrap_or(cy);
        let nz = num(r"Z(-?\d+(?:\.\d+)?)").unwrap_or(cz);
        let r_val = num(r"R(-?\d+(?:\.\d+)?)");
        if is_g0 {
            push(Mt::Rapid, cx, cy, cz);
            push(Mt::Rapid, nx, ny, nz);
        } else if is_g1 {
            push(Mt::Cut, cx, cy, cz);
            push(Mt::Cut, nx, ny, nz);
        } else if (is_g2 || is_g3) && r_val.is_some() {
            for (ax, ay) in arc_to_polyline(cx, cy, nx, ny, r_val.unwrap(), is_g2) {
                push(Mt::Arc, ax, ay, cz);
            }
        } else {
            push(Mt::Cut, cx, cy, cz);
            push(Mt::Cut, nx, ny, nz);
        }
        cx = nx; cy = ny; cz = nz;
    }

    let bounds = if any {
        Bounds { min_x, max_x, min_y, max_y }
    } else {
        Bounds::default()
    };
    Ok(GcodeParsed { groups, counts, bounds })
}

