// Literal port of isel_to_gcode_gui.py convert_file() (python lines 497-983).
// That reference file is no longer in the tree: git show 9b1fa1d^:isel_to_gcode_gui.py
// Arithmetic order and format strings are copied as-is: golden tests compare
// bytes against the python reference output.

use std::collections::HashMap;
use std::io::Write;
use std::sync::OnceLock;

pub const SCALE: f64 = 1000.0;
pub const VEL_RATIO: f64 = 16.6667;
pub const ARC_RESOLUTION: f64 = 0.05;
pub const DEFAULT_FEED: f64 = 1000.0;
pub const RAPID_FEED: f64 = 3000.0;
pub const ARC_FIT_MIN_DIAMETER: f64 = 5.0;
pub const ARC_FIT_RADIUS_TOL: f64 = 0.05;
pub const ARC_FIT_MIN_ARC_DEG: f64 = 355.0;
pub const ARC_FIT_MIN_POINTS: usize = 8;
pub const PROGRESS_STEP: u64 = 10000;

fn re(pat: &str) -> &'static regex::Regex {
    static CACHE: OnceLock<HashMap<&'static str, regex::Regex>> = OnceLock::new();
    CACHE
        .get_or_init(|| {
            [
                r"RPM(\d+)",
                r"VEL\s*(\d+(?:\.\d+)?)",
                r"X(-?\d+(?:\.\d+)?)",
                r"Y(-?\d+(?:\.\d+)?)",
                r"Z(-?\d+(?:\.\d+)?)",
                r"I(-?\d+(?:\.\d+)?)",
                r"J(-?\d+(?:\.\d+)?)",
            ]
            .iter()
            .map(|p| (*p, regex::Regex::new(p).unwrap()))
            .collect()
        })
        .get(pat)
        .unwrap_or_else(|| panic!("regex not preloaded: {}", pat))
}

/// python parse_coord(): fast int path, guarded float path, exact-regex fallback.
pub fn parse_coord(text: &str, axes: &[char]) -> HashMap<char, f64> {
    let mut coords: HashMap<char, f64> = HashMap::new();
    for token in text.split_whitespace().skip(1) {
        let axis = token.chars().next().unwrap();
        if !axes.contains(&axis) {
            continue;
        }
        let val = &token[axis.len_utf8()..];
        if val.is_empty() {
            continue;
        }
        // Fast path: plain integer. Python int()/SCALE == this f64 division.
        let neg = val.starts_with('-');
        let body = if neg { &val[1..] } else { val };
        // shortcut: ASCII digits only (python accepts unicode digits); ISEL files are ASCII.
        if !body.is_empty() && body.bytes().all(|b| b.is_ascii_digit()) {
            if let Ok(v) = body.parse::<i64>() {
                coords.entry(axis).or_insert((if neg { -v } else { v }) as f64 / SCALE);
                continue;
            }
        }
        // Guarded float path: must match the original regex -?\d+(?:\.\d+)?
        let bytes = val.as_bytes();
        let digit_start = bytes[0].is_ascii_digit()
            || (neg && bytes.len() > 1 && bytes[1].is_ascii_digit());
        if digit_start && !val.contains('e') && !val.contains('E') && val.matches('.').count() <= 1 {
            if let Ok(v) = val.parse::<f64>() {
                coords.entry(axis).or_insert(v / SCALE);
                continue;
            }
        }
        // Fallback: original per-axis regex on the whole line (first occurrence wins).
        coords.clear();
        for &a in axes {
            let pat = match a {
                'X' => r"X(-?\d+(?:\.\d+)?)",
                'Y' => r"Y(-?\d+(?:\.\d+)?)",
                'Z' => r"Z(-?\d+(?:\.\d+)?)",
                'I' => r"I(-?\d+(?:\.\d+)?)",
                _ => r"J(-?\d+(?:\.\d+)?)",
            };
            if let Some(cap) = re(pat).captures(text) {
                if let Ok(v) = cap[1].parse::<f64>() {
                    coords.insert(a, v / SCALE);
                }
            }
        }
        return coords;
    }
    coords
}

/// Least-squares circle through >=3 points (python fit_circle_to_points).
pub fn fit_circle(pts: &[(f64, f64)]) -> Option<(f64, f64, f64)> {
    let n = pts.len();
    if n < 3 {
        return None;
    }
    let (mut sum_x, mut sum_y, mut sum_x2, mut sum_y2, mut sum_xy) = (0f64, 0f64, 0f64, 0f64, 0f64);
    let (mut sum_x3, mut sum_y3, mut sum_xy2, mut sum_x2y) = (0f64, 0f64, 0f64, 0f64);
    for &(x, y) in pts {
        sum_x += x;
        sum_y += y;
        sum_x2 += x * x;
        sum_y2 += y * y;
        sum_xy += x * y;
        sum_x3 += x.powi(3);
        sum_y3 += y.powi(3);
        sum_xy2 += x * y.powi(2);
        sum_x2y += x.powi(2) * y;
    }
    let (a11, a12, a13) = (sum_x2, sum_xy, sum_x);
    let (a21, a22, a23) = (sum_xy, sum_y2, sum_y);
    let (a31, a32, a33) = (sum_x, sum_y, n as f64);
    let b1 = -(sum_x3 + sum_xy2);
    let b2 = -(sum_x2y + sum_y3);
    let b3 = -(sum_x2 + sum_y2);
    let det = a11 * (a22 * a33 - a23 * a32) - a12 * (a21 * a33 - a23 * a31)
        + a13 * (a21 * a32 - a22 * a31);
    if det.abs() < 1e-12 {
        return None;
    }
    let d = (b1 * (a22 * a33 - a23 * a32) - a12 * (b2 * a33 - a23 * b3)
        + a13 * (b2 * a32 - a22 * b3))
        / det;
    let e = (a11 * (b2 * a33 - a23 * b3) - b1 * (a21 * a33 - a23 * a31)
        + a13 * (a21 * b3 - b2 * a31))
        / det;
    let f = (a11 * (a22 * b3 - b2 * a32) - a12 * (a21 * b3 - b2 * a31)
        + b1 * (a21 * a32 - a22 * a31))
        / det;
    let cx = -d / 2.0;
    let cy = -e / 2.0;
    let r2 = cx * cx + cy * cy - f;
    if r2 <= 0.0 {
        return None;
    }
    Some((cx, cy, r2.sqrt()))
}

/// python try_arc_fit(): detects a full circle in a buffered polyline.
pub fn try_arc_fit(buffer: &[(f64, f64, f64)], start: (f64, f64, f64)) -> Option<(f64, f64, f64, bool)> {
    if buffer.len() < ARC_FIT_MIN_POINTS {
        return None;
    }
    let (mut zmin, mut zmax) = (f64::INFINITY, f64::NEG_INFINITY);
    for p in buffer {
        if p.2 < zmin {
            zmin = p.2;
        }
        if p.2 > zmax {
            zmax = p.2;
        }
    }
    if zmax - zmin > 1e-4 {
        return None;
    }
    let mut all_pts: Vec<(f64, f64)> = Vec::with_capacity(buffer.len() + 1);
    all_pts.push((start.0, start.1));
    for p in buffer {
        all_pts.push((p.0, p.1));
    }
    let (cx, cy, radius) = fit_circle(&all_pts)?;
    if radius * 2.0 < ARC_FIT_MIN_DIAMETER {
        return None;
    }
    for &(x, y) in &all_pts {
        if ((x - cx).hypot(y - cy) - radius).abs() > ARC_FIT_RADIUS_TOL {
            return None;
        }
    }
    let angles: Vec<f64> = all_pts.iter().map(|p| (p.1 - cy).atan2(p.0 - cx)).collect();
    let (mut cw_votes, mut ccw_votes) = (0usize, 0usize);
    for i in 1..all_pts.len() - 1 {
        let ax = all_pts[i].0 - all_pts[i - 1].0;
        let ay = all_pts[i].1 - all_pts[i - 1].1;
        let bx = all_pts[i + 1].0 - all_pts[i].0;
        let by = all_pts[i + 1].1 - all_pts[i].1;
        let cross = ax * by - ay * bx;
        if cross < 0.0 {
            cw_votes += 1;
        } else {
            ccw_votes += 1;
        }
    }
    let cw = cw_votes > ccw_votes;
    let mut total_angle = 0f64;
    for i in 1..angles.len() {
        let mut da = angles[i] - angles[i - 1];
        while da > std::f64::consts::PI {
            da -= 2.0 * std::f64::consts::PI;
        }
        while da < -std::f64::consts::PI {
            da += 2.0 * std::f64::consts::PI;
        }
        if cw {
            if da > 0.0 {
                da -= 2.0 * std::f64::consts::PI;
            }
        } else if da < 0.0 {
            da += 2.0 * std::f64::consts::PI;
        }
        total_angle += da.abs();
    }
    if total_angle.to_degrees() < ARC_FIT_MIN_ARC_DEG {
        return None;
    }
    Some((cx, cy, radius, cw))
}

#[derive(Clone, Copy, Default)]
pub struct Pos {
    pub x: f64,
    pub y: f64,
    pub z: f64,
}

pub struct Conv {
    use_arc: bool,
    line_no: u32,
    total_time_min: f64,
    current_feed: Option<f64>,
    pending_spindle: Option<String>,
    last_pos: Pos,
    buffer: Vec<Pos>,
    buffer_start: Option<Pos>,
    arc_fit_converted: u32,
    arc_fit_skipped: u32,
}

impl Conv {
    fn new(use_arc: bool) -> Self {
        Conv {
            use_arc,
            line_no: 1,
            total_time_min: 0.0,
            current_feed: None,
            pending_spindle: None,
            last_pos: Pos::default(),
            buffer: vec![],
            buffer_start: None,
            arc_fit_converted: 0,
            arc_fit_skipped: 0,
        }
    }

    fn nline(&mut self) -> String {
        let s = format!("N{:05} ", self.line_no);
        self.line_no += 1;
        s
    }

    fn feed(&self) -> f64 {
        self.current_feed.unwrap_or(0.0)
    }

    fn flush_spindle(&mut self) -> Vec<String> {
        match self.pending_spindle.take() {
            Some(sp) => vec![self.nline() + &sp],
            None => vec![],
        }
    }

    fn ensure_default_feed(&mut self, log: &mut dyn FnMut(&str)) -> Vec<String> {
        if self.current_feed.is_none() {
            self.current_feed = Some(DEFAULT_FEED);
            let line = self.nline() + &format!("F{:.0}", DEFAULT_FEED);
            log("Warning: No VEL found, using default F1000");
            return vec![line];
        }
        vec![]
    }

    /// python arc_to_g2g3(): single G2/G3 with R (negative R for spans > 180 deg).
    fn arc_to_g2g3(&mut self, start: Pos, end: Pos, center: Pos, cw: bool) -> Vec<String> {
        let (cx, cy) = (center.x, center.y);
        let r = (start.x - cx).hypot(start.y - cy);
        let a0 = (start.y - cy).atan2(start.x - cx);
        let mut a1 = (end.y - cy).atan2(end.x - cx);
        if cw {
            if a1 >= a0 {
                a1 -= 2.0 * std::f64::consts::PI;
            }
        } else if a1 <= a0 {
            a1 += 2.0 * std::f64::consts::PI;
        }
        let span = (a1 - a0).abs();
        let arc_len = span * r;
        let feed = self.feed();
        if feed != 0.0 {
            self.total_time_min += arc_len / feed;
        }
        let r_signed = if span <= std::f64::consts::PI { r } else { -r };
        let code = if cw { "G2" } else { "G3" };
        let f_str = if feed != 0.0 { format!(" F{:.0}", feed) } else { String::new() };
        vec![self.nline()
            + &format!(
                "{} X{:.3} Y{:.3} R{:.3}{}",
                code, end.x, end.y, r_signed, f_str
            )]
    }

    /// python fitted_arc_to_g2g3(): full circle as two semicircles.
    /// Reached from flush_moveabs_buffer when a buffered run closes a circle.
    fn fitted_arc_to_g2g3(
        &mut self,
        start: Pos,
        end: Pos,
        cx: f64,
        cy: f64,
        radius: f64,
        cw: bool,
    ) -> Vec<String> {
        let feed = self.feed();
        if feed != 0.0 {
            self.total_time_min += 2.0 * std::f64::consts::PI * radius / feed;
        }
        let a_start = (start.y - cy).atan2(start.x - cx);
        let a_mid = a_start + std::f64::consts::PI;
        let mx = cx + radius * a_mid.cos();
        let my = cy + radius * a_mid.sin();
        let chord = (mx - start.x).hypot(my - start.y);
        let r1 = chord / 2.0 + 0.001;
        let code = if cw { "G2" } else { "G3" };
        let f_str = if feed != 0.0 { format!(" F{:.0}", feed) } else { String::new() };
        let _ = end;
        vec![
            self.nline() + &format!("{} X{:.3} Y{:.3} R{:.3}{}", code, mx, my, r1, f_str),
            self.nline()
                + &format!(
                    "{} X{:.3} Y{:.3} R{:.3}{}",
                    code, start.x, start.y, r1, f_str
                ),
        ]
    }

    /// python linearize_arc_pure(): polyline at ARC_RESOLUTION mm steps.
    fn linearize_arc(&mut self, start: Pos, end: Pos, center: Pos, cw: bool) -> Vec<String> {
        let (cx, cy) = (center.x, center.y);
        let z = start.z;
        let r = (start.x - cx).hypot(start.y - cy);
        let a0 = (start.y - cy).atan2(start.x - cx);
        let mut a1 = (end.y - cy).atan2(end.x - cx);
        if cw {
            if a1 >= a0 {
                a1 -= 2.0 * std::f64::consts::PI;
            }
        } else if a1 <= a0 {
            a1 += 2.0 * std::f64::consts::PI;
        }
        let arc_len = (a1 - a0).abs() * r;
        let steps = ((arc_len / ARC_RESOLUTION) as i64).max(1);
        let da = (a1 - a0) / steps as f64;
        let feed = self.feed();
        let mut lines = Vec::with_capacity(steps as usize + 1);
        let (mut px, mut py) = (start.x, start.y);
        for i in 1..=steps {
            let a = a0 + da * i as f64;
            let tx = cx + r * a.cos();
            let ty = cy + r * a.sin();
            lines.push(self.nline() + &format!("G1 X{:.3} Y{:.3} Z{:.3}", tx, ty, z));
            if feed != 0.0 {
                self.total_time_min += (tx - px).hypot(ty - py) / feed;
            }
            px = tx;
            py = ty;
        }
        if (end.x - px).abs() > 0.001 || (end.y - py).abs() > 0.001 {
            lines.push(self.nline() + &format!("G1 X{:.3} Y{:.3} Z{:.3}", end.x, end.y, end.z));
            if feed != 0.0 {
                self.total_time_min += (end.x - px).hypot(end.y - py) / feed;
            }
        }
        lines
    }

    /// python emit_g1_buffer(): one G1 per point, only CHANGED axes printed.
    fn emit_g1_buffer(&mut self, buffer: &[Pos], start: Pos) -> Vec<String> {
        let mut lines = vec![];
        let mut prev = start;
        let feed = self.feed();
        for pt in buffer {
            let mut cmd = String::from("G1");
            if (pt.x - prev.x).abs() > 1e-6 {
                cmd.push_str(&format!(" X{:.3}", pt.x));
            }
            if (pt.y - prev.y).abs() > 1e-6 {
                cmd.push_str(&format!(" Y{:.3}", pt.y));
            }
            if (pt.z - prev.z).abs() > 1e-6 {
                cmd.push_str(&format!(" Z{:.3}", pt.z));
            }
            if cmd == "G1" {
                prev = *pt;
                continue;
            }
            lines.push(self.nline() + &cmd);
            let d = ((pt.x - prev.x).powi(2) + (pt.y - prev.y).powi(2) + (pt.z - prev.z).powi(2)).sqrt();
            if feed != 0.0 && d > 0.0 {
                self.total_time_min += d / feed;
            }
            prev = *pt;
        }
        lines
    }

    /// python flush_moveabs_buffer(). Full-circle fit is live here: the MOVEABS
    /// branch no longer flushes at its top, so a run can reach ARC_FIT_MIN_POINTS
    /// points and try_arc_fit turns closed circles into two semicircles in arcs mode.
    fn flush_moveabs_buffer(&mut self, log: &mut dyn FnMut(&str)) -> Vec<String> {
        if self.buffer.is_empty() {
            return vec![];
        }
        if self.use_arc && self.buffer.len() >= ARC_FIT_MIN_POINTS {
            let pts: Vec<(f64, f64, f64)> = self.buffer.iter().map(|p| (p.x, p.y, p.z)).collect();
            let start = self.buffer_start.unwrap_or_default();
            if let Some((cx, cy, radius, cw)) = try_arc_fit(&pts, (start.x, start.y, start.z)) {
                let lines = self.fitted_arc_to_g2g3(start, *self.buffer.last().unwrap(), cx, cy, radius, cw);
                self.last_pos = *self.buffer.last().unwrap();
                self.arc_fit_converted += 1;
                log(&format!(
                    "  Arc fit: full circle \u{f8}{:.2}mm \u{2192} {} R{:.3} (two semicircles)",
                    radius * 2.0,
                    if cw { "G2" } else { "G3" },
                    radius
                ));
                self.buffer.clear();
                return lines;
            }
        }
        let start = self.buffer_start.unwrap_or_default();
        let buffer = std::mem::take(&mut self.buffer);
        let lines = self.emit_g1_buffer(&buffer, start);
        if let Some(last) = buffer.last() {
            self.last_pos = *last;
        }
        if self.use_arc && buffer.len() >= ARC_FIT_MIN_POINTS {
            self.arc_fit_skipped += 1;
        }
        self.buffer.clear();
        lines
    }
}

pub fn convert_file(
    in_path: &str,
    out_path: &str,
    use_arc: bool,
    log: &mut dyn FnMut(&str),
    progress: &mut Option<Box<dyn FnMut(f64, &str)>>,
) -> Result<f64, String> {
    let bytes = match std::fs::read(in_path) {
        Ok(b) => b,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            return Err(format!("Input file not found: {}", in_path));
        }
        Err(e) => return Err(format!("Error reading input file: {}", e)),
    };
    // python reads with errors="replace"; lossy bytes cannot form an ASCII
    // command token, so parse behaviour is identical.
    let text = String::from_utf8_lossy(&bytes).into_owned();
    let total_lines = text
        .lines()
        .filter(|l| {
            let t = l.trim();
            !t.is_empty() && !t.starts_with(';')
        })
        .count() as u64;
    if total_lines == 0 {
        return Err("Input file contains no processable commands.".to_string());
    }
    if let Some(p) = progress {
        p(0.0, "Reading file\u{2026}");
    }

    let file = std::fs::File::create(out_path)
        .map_err(|e| format!("Cannot open output file for writing: {}", e))?;
    // python writes text mode on Windows => CRLF; goldens pin that.
    let mut out = std::io::BufWriter::with_capacity(256 * 1024, file);
    let mut emit = |lines: Vec<String>| -> std::io::Result<()> {
        for l in lines {
            out.write_all(l.as_bytes())?;
            out.write_all(b"\r\n")?;
        }
        Ok(())
    };

    let mut c = Conv::new(use_arc);
    let mut header = vec![];
    for kw in ["G21", "G17", "G40", "G49", "G90", "G94"] {
        header.push(c.nline() + kw);
    }
    if let Err(e) = emit(header) {
        return Err(format!("Cannot open output file for writing: {}", e));
    }

    let mut processed: u64 = 0;
    for raw in text.lines() {
        let line = raw.trim();
        if line.is_empty() || line.starts_with(';') {
            continue;
        }
        processed += 1;
        if let Some(p) = progress {
            if processed % PROGRESS_STEP == 0 || processed == total_lines {
                let pct = (processed as f64 / total_lines as f64) * 95.0;
                p(pct, &format!("Processing line {}/{}", processed, total_lines));
            }
        }

        let mut batch: Vec<String> = vec![];
        if line.starts_with("SPINDLE CW") {
            batch.extend(c.flush_moveabs_buffer(log));
            match re(r"RPM(\d+)").captures(line) {
                Some(cap) => {
                    let rpm = &cap[1];
                    if c.pending_spindle.is_none() {
                        c.pending_spindle = Some(format!("S{} M03", rpm));
                        log(&format!("Spindle: {} RPM (deferred until first Z retract)", rpm));
                    } else {
                        batch.push(c.nline() + &format!("S{} M03", rpm));
                        log(&format!("Spindle: {} RPM", rpm));
                    }
                }
                None => log(&format!("Warning: Could not parse RPM from: {}", line)),
            }
        } else if line.starts_with("FASTABS") {
            batch.extend(c.flush_moveabs_buffer(log));
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            let (mut lx, mut ly, mut lz) = (c.last_pos.x, c.last_pos.y, c.last_pos.z);
            if let Some(v) = cc.get(&'X') {
                lx = *v;
            }
            if let Some(v) = cc.get(&'Y') {
                ly = *v;
            }
            if let Some(v) = cc.get(&'Z') {
                lz = *v;
            }
            let mut cmd = String::from("G0");
            if cc.contains_key(&'X') {
                cmd.push_str(&format!(" X{:.3}", lx));
            }
            if cc.contains_key(&'Y') {
                cmd.push_str(&format!(" Y{:.3}", ly));
            }
            if cc.contains_key(&'Z') {
                cmd.push_str(&format!(" Z{:.3}", lz));
            }
            batch.push(c.nline() + &cmd);
            // First Z move is the safe retract: spindle starts only after it.
            if cc.contains_key(&'Z') {
                batch.extend(c.flush_spindle());
            }
            let d = ((lx - c.last_pos.x).powi(2) + (ly - c.last_pos.y).powi(2)
                + (lz - c.last_pos.z).powi(2))
            .sqrt();
            if d > 0.0 {
                c.total_time_min += d / RAPID_FEED;
            }
            c.last_pos = Pos { x: lx, y: ly, z: lz };
        } else if line.starts_with("MOVEABS") {
            // No buffer flush here: a run of MOVEABS must stay buffered so the
            // full-circle fit can see >=8 points (python line 876 equivalent).
            // The non-buffered path flushes in its else branch.
            batch.extend(c.flush_spindle());
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            batch.extend(c.ensure_default_feed(log));
            let mut target = c.last_pos;
            if let Some(v) = cc.get(&'X') {
                target.x = *v;
            }
            if let Some(v) = cc.get(&'Y') {
                target.y = *v;
            }
            if let Some(v) = cc.get(&'Z') {
                target.z = *v;
            }
            let has_xy = cc.contains_key(&'X') || cc.contains_key(&'Y');
            if c.use_arc && has_xy {
                if c.buffer.is_empty() {
                    c.buffer_start = Some(c.last_pos);
                }
                c.buffer.push(target);
                c.last_pos = target;
            } else {
                // python: a non-XY MOVEABS in arc mode ends the buffered run first
                if c.use_arc {
                    batch.extend(c.flush_moveabs_buffer(log));
                }
                // Direct G1: axes PRESENT on the line are printed even if unchanged.
                let mut cmd = String::from("G1");
                if cc.contains_key(&'X') {
                    cmd.push_str(&format!(" X{:.3}", target.x));
                }
                if cc.contains_key(&'Y') {
                    cmd.push_str(&format!(" Y{:.3}", target.y));
                }
                if cc.contains_key(&'Z') {
                    cmd.push_str(&format!(" Z{:.3}", target.z));
                }
                batch.push(c.nline() + &cmd);
                let feed = c.feed();
                if feed != 0.0 {
                    let d = ((target.x - c.last_pos.x).powi(2)
                        + (target.y - c.last_pos.y).powi(2)
                        + (target.z - c.last_pos.z).powi(2))
                    .sqrt();
                    c.total_time_min += d / feed;
                }
                c.last_pos = target;
            }
        } else if line.starts_with("VEL") {
            batch.extend(c.flush_moveabs_buffer(log));
            batch.extend(c.flush_spindle());
            match re(r"VEL\s*(\d+(?:\.\d+)?)").captures(line) {
                Some(cap) => {
                    let f = cap[1].parse::<f64>().unwrap_or(0.0) / VEL_RATIO;
                    c.current_feed = Some(f);
                    batch.push(c.nline() + &format!("F{:.0}", f));
                    log(&format!("Feed: F{:.0}", f));
                }
                None => log(&format!("Warning: Could not parse VEL from: {}", line)),
            }
        } else if line.starts_with("CWABS") || line.starts_with("CCWABS") {
            batch.extend(c.flush_moveabs_buffer(log));
            batch.extend(c.flush_spindle());
            let cw = line.starts_with("CWABS");
            batch.extend(c.ensure_default_feed(log));
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            let ij = parse_coord(line, &['I', 'J']);
            let mut end = c.last_pos;
            if let Some(v) = cc.get(&'X') {
                end.x = *v;
            }
            if let Some(v) = cc.get(&'Y') {
                end.y = *v;
            }
            if let Some(v) = cc.get(&'Z') {
                end.z = *v;
            }
            let center = Pos {
                x: ij.get(&'I').copied().unwrap_or(0.0),
                y: ij.get(&'J').copied().unwrap_or(0.0),
                z: 0.0,
            };
            let r_check = (c.last_pos.x - center.x).hypot(c.last_pos.y - center.y);
            if r_check < 1e-6 {
                log(&format!("Warning: zero-radius arc skipped on line [{}]", line));
            } else {
                if c.use_arc {
                    batch.extend(c.arc_to_g2g3(c.last_pos, end, center, cw));
                } else {
                    batch.extend(c.linearize_arc(c.last_pos, end, center, cw));
                }
                c.last_pos = end;
            }
        }

        if let Err(e) = emit(batch) {
            return Err(format!("Cannot open output file for writing: {}", e));
        }
    }

    let mut tail = c.flush_moveabs_buffer(log);
    tail.extend(c.flush_spindle());
    tail.push(c.nline() + "M05");
    tail.push(c.nline() + "M30");
    if let Err(e) = emit(tail) {
        return Err(format!("Cannot open output file for writing: {}", e));
    }
    if let Err(e) = out.flush() {
        return Err(format!("Cannot open output file for writing: {}", e));
    }
    if let Some(p) = progress {
        p(100.0, "Complete!");
    }
    if c.use_arc && (c.arc_fit_converted > 0 || c.arc_fit_skipped > 0) {
        log(&format!(
            "Arc fitting summary: {} circle(s) converted to G2/G3, {} segment group(s) left as G1",
            c.arc_fit_converted, c.arc_fit_skipped
        ));
    }
    Ok(c.total_time_min)
}
