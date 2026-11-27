// Fanuc-dialect G-code input (e.g. a SolidWorks CAM export): rewrite program lines
// into the equivalent ISEL NCP commands, then hand them to the untouched converter
// engine. This module never emits G-code and never changes geometry: it maps words
// only — modal motion, modal axis omission, inline modal F, R/I-J arcs.
//
// The dialect subset is what a real export contains (measured on a customer file:
// absolute mm (G90/G21), modal motion G00-G03 with address omission on ~99% of
// lines, inline modal feed, arcs as incremental I/J or R with negative R for
// major arcs, S+M03 spindle, T..M06 with a Z-bearing G43 line, G91 G28 homing
// blocks skipped, M02/M30 end the program). Everything that would silently change
// geometry or units is a loud error naming the line, never a guess:
// G20 inch, G18/G19 planes, G41/G42 compensation, G73/G81-G89 cycles, bare G91,
// G92, G04 dwell, M04, unknown G words.
//
// shortcut: one spindle, XY plane, mm; upgrade to per-operation state when a real
// multi-spindle or inch export appears.

enum Modal {
    None,
    Rapid,
    Cut,
    Cw,
    Ccw,
}

/// The engine's command words must never be reinterpreted: any ISEL signal in a
/// non-comment line makes the file ISEL, no matter what else it contains.
pub fn looks_like_fanuc(text: &str) -> bool {
    const ISEL_WORDS: [&str; 10] = [
        "MOVEABS", "FASTABS", "CWABS", "CCWABS", "VEL", "SPINDLE", "IMF_PBL", "PLANE",
        "GETTOOL", "WPZERO",
    ];
    let mut fanuc = false;
    for raw in text.lines() {
        let t = raw.trim();
        if t.is_empty() || t.starts_with(';') || t.starts_with('(') {
            continue;
        }
        for kw in ISEL_WORDS {
            if t.starts_with(kw) {
                return false;
            }
        }
        if (t.starts_with('O') && t[1..].starts_with(|c: char| c.is_ascii_digit())) || has_motion_g(t)
        {
            fanuc = true;
        }
    }
    fanuc
}

/// G0..G3 motion words, token-grepped: G17/G21/G90 are state words, not motion.
fn has_motion_g(t: &str) -> bool {
    let b = t.as_bytes();
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'G' && i + 1 < b.len() && b[i + 1].is_ascii_digit() {
            let mut j = i + 1;
            let mut v = 0u32;
            while j < b.len() && b[j].is_ascii_digit() {
                v = v * 10 + (b[j] - b'0') as u32;
                j += 1;
            }
            // G04 prints as "G4" and G0 as "G0"/"G00": accept bare digits 0..3 only
            // when the next char cannot continue the number.
            if v <= 3 && (j >= b.len() || !(b[j].is_ascii_digit() || b[j] == b'.')) {
                return true;
            }
            i = j;
        } else {
            i += 1;
        }
    }
    false
}

/// Split a block into (letter, value) words, tolerating the number shapes a post
/// emits: ".02" without a leading digit, "Z25." with a trailing dot, signed values.
/// A bare letter (no value) is skipped: it carries no target.
fn words(line: &str) -> Vec<(char, f64)> {
    let b = line.as_bytes();
    let mut out = vec![];
    let mut i = 0;
    while i < b.len() {
        if !b[i].is_ascii_alphabetic() {
            i += 1;
            continue;
        }
        let letter = (b[i] as char).to_ascii_uppercase();
        let mut j = i + 1;
        let num = j;
        if j < b.len() && (b[j] == b'-' || b[j] == b'+') {
            j += 1;
        }
        let digits0 = j;
        while j < b.len() && b[j].is_ascii_digit() {
            j += 1;
        }
        let mut dot = false;
        if j < b.len() && b[j] == b'.' {
            dot = true;
            j += 1;
            while j < b.len() && b[j].is_ascii_digit() {
                j += 1;
            }
        }
        if j > digits0 || (dot && j > digits0 + 1) {
            if let Ok(v) = line[num..j].parse::<f64>() {
                out.push((letter, v));
            }
            i = j;
        } else {
            i += 1;
        }
    }
    out
}

/// Drop the trailing `;` comment and every `(...)` span, like the preview parser
/// does; an unclosed `(` drops the rest of the line (posts never nest comments).
fn clean(raw: &str) -> String {
    let t = match raw.find(';') {
        Some(p) => &raw[..p],
        None => raw,
    };
    let mut out = String::with_capacity(t.len());
    let mut rest = t;
    while let Some(o) = rest.find('(') {
        out.push_str(&rest[..o]);
        rest = match rest[o + 1..].find(')') {
            Some(c) => &rest[o + c + 2..],
            None => "",
        };
    }
    out.push_str(rest);
    out.trim().to_string()
}

/// Pick the arc center for an R-format move by FANUC semantics: R > 0 is the
/// arc of <= 180 deg, R < 0 the major arc; the commanded direction fixes which
/// side the center lies on. (Note: gcode.rs's preview parser keeps the python
/// reference's side rule, which is the inverse for non-semicircles; input here
/// follows the controller, not the preview.)
fn arc_center(x0: f64, y0: f64, x1: f64, y1: f64, r_signed: f64, cw: bool) -> Option<(f64, f64)> {
    let r = r_signed.abs();
    let (dx, dy) = (x1 - x0, y1 - y0);
    let d = dx.hypot(dy);
    if d < 1e-9 {
        return None;
    }
    let h2 = r * r - (d / 2.0) * (d / 2.0);
    if h2 < -1e-9 {
        return None;
    }
    let h = h2.max(0.0).sqrt();
    let (mx, my) = ((x0 + x1) / 2.0, (y0 + y1) / 2.0);
    let (px, py) = (-dy / d, dx / d);
    let (cx1, cy1) = (mx + h * px, my + h * py);
    let (cx2, cy2) = (mx - h * px, my - h * py);
    // cross > 0 means c1 lies left of the travel direction. A CW arc keeps its
    // center right (minor) or left (major); a CCW arc mirrors that.
    let cross = dx * (cy1 - y0) - dy * (cx1 - x0);
    let want_left = cw == (r_signed < 0.0);
    if want_left {
        if cross > 0.0 {
            Some((cx1, cy1))
        } else {
            Some((cx2, cy2))
        }
    } else if cross > 0.0 {
        Some((cx2, cy2))
    } else {
        Some((cx1, cy1))
    }
}

fn axis_word(out: &mut String, letter: char, v_mm: f64) {
    // ISEL coordinates are thousandths of mm; CAM emits <= 3 decimals, so the
    // rounding below is exact for every real value.
    out.push_str(&format!(" {}{:.0}", letter, (v_mm * 1000.0).round()));
}

/// Rewrite a Fanuc-dialect program into ISEL command lines ("\n"-joined), the
/// exact lines the engine then converts. Errors name the 1-based source line.
pub fn rewrite_to_isel(text: &str) -> Result<String, String> {
    let mut out: Vec<String> = vec![];
    let (mut x, mut y, mut z) = (0f64, 0f64, 0f64);
    let mut feed: Option<f64> = None;
    let mut modal = Modal::None;
    let mut last_ij = (0f64, 0f64); // center vector of the last arc, modal repeats it
    let mut last_r: Option<f64> = None; // modal repeat after an R arc repeats the R word
    let mut last_style_ij = true;
    let mut abs_ij = false; // G90.1 makes I/J an absolute center instead
    let mut last_rpm: Option<f64> = None;

    for (idx, raw) in text.lines().enumerate() {
        let line_no = idx + 1;
        let line = clean(raw);
        if line.is_empty() || line.starts_with('%') || line.starts_with('$') {
            continue;
        }
        if line.starts_with('O') && line[1..].starts_with(|c: char| c.is_ascii_digit()) {
            continue; // program header, O0001
        }
        let w = words(&line);
        let err = |m: String| format!("Fanuc line {}: {} [{}]", line_no, m, line.trim());

        // Pre-scan: a homing return (G28/G30) drops the whole line — the G91 that
        // posts put on the same line must not reach its error arm first.
        if w.iter().any(|(l, v)| *l == 'G' && (*v == 28.0 || *v == 30.0)) {
            continue;
        }
        let (mut gx, mut gy, mut gz) = (None, None, None);
        let (mut gi, mut gj, mut gr, mut gf) = (None, None, None, None);
        // S is read before the word loop: "M03 S1200" must behave like "S1200 M03".
        let spindle_rpm = w.iter().find(|(l, _)| *l == 'S').map(|(_, v)| *v);
        let mut program_end = false;
        for (letter, v) in &w {
            match letter {
                'X' => gx = Some(*v),
                'Y' => gy = Some(*v),
                'Z' => gz = Some(*v),
                'I' => gi = Some(*v),
                'J' => gj = Some(*v),
                'R' => gr = Some(*v),
                'F' => gf = Some(*v),
                'S' => {}
                'M' => match *v as u32 {
                    3 => {
                        let rpm = spindle_rpm.or(last_rpm).ok_or_else(|| err("M03 without an S word".to_string()))?;
                        out.push(format!("SPINDLE CW RPM{:.0}", rpm));
                    }
                    4 => return Err(err("M04 spindle CCW: only CW spindle is supported".to_string())),
                    2 | 30 => program_end = true,
                    _ => {} // M00/M01/M05/M06/M08/M09 and unknown M words carry no geometry
                },
                'G' => {
                    let g = *v;
                    // Homing returns are machine moves, not part geometry: a line
                    // with G28/G30 is dropped whole (it may carry G91 + X0 Y0 Z0).
                    if g == 28.0 || g == 30.0 {
                        continue;
                    }
                    match g {
                        0.0 => modal = Modal::Rapid,
                        1.0 => modal = Modal::Cut,
                        2.0 => modal = Modal::Cw,
                        3.0 => modal = Modal::Ccw,
                        // state words the engine header already sets, or that mean
                        // "as produced": nothing to rewrite.
                        17.0 | 21.0 | 40.0 | 43.0 | 44.0 | 49.0 | 53.0 | 61.0 | 63.0
                        | 64.0 | 80.0 | 90.0 | 94.0 => {}
                        54.0..=59.0 => {} // work offsets: the part program is at zero
                        90.1 => abs_ij = true,
                        91.1 => abs_ij = false,
                        20.0 => return Err(err("G20 inch programming: export in mm".to_string())),
                        18.0 | 19.0 => return Err(err("G18/G19 interpolation plane: the converter is XY-only".to_string())),
                        41.0 | 42.0 => return Err(err("cutter compensation (G41/G42): bake the offset into the toolpath".to_string())),
                        73.0 | 81.0..=89.0 => return Err(err("canned drilling cycle: turn cycles off in the post".to_string())),
                        91.0 => return Err(err("incremental programming (G91): export absolute".to_string())),
                        92.0 => return Err(err("G92 work offset shift is not supported".to_string())),
                        4.0 => return Err(err("G04 dwell is not supported".to_string())),
                        10.0 => return Err(err("G10 is not supported".to_string())),
                        95.0 => return Err(err("G95 feed per revolution: export feed per minute".to_string())),
                        _ => return Err(err(format!("G{:?} is not supported", g))),
                    }
                }
                // Tool/housings/rotary words: T..M06 changes a tool the converter
                // does not model; H/D/P/Q/L are word arguments; A/B/C would mean
                // a rotary axis the geometry model cannot take.
                'T' | 'H' | 'D' | 'P' | 'Q' | 'L' | 'W' | 'K' | 'U' | 'V' | 'N' | 'E' => {}
                'A' | 'B' | 'C' => return Err(err(format!("{} rotary-axis word: 3-axis output only", letter))),
                _ => {}
            }
        }
        if spindle_rpm.is_some() {
            last_rpm = spindle_rpm;
        }

        // Feed first: a changed F becomes a VEL line before the move of this block
        // (and on its own for a feed-only block). VEL_RATIO round-trips the printed
        // F through the engine's division.
        if let Some(f) = gf {
            if Some(f) != feed {
                out.push(format!("VEL {:.4}", f * crate::convert::VEL_RATIO));
                feed = Some(f);
            }
        }

        let moved = gx.is_some() || gy.is_some() || gz.is_some();
        // I/J is a pair even for full circles with no X/Y endpoint. Otherwise an
        // I-only full-circle block reaches gj.unwrap() below and panics.
        if gi.is_some() != gj.is_some() {
            return Err(err("partial I/J arc pair".to_string()));
        }
        let arc_words = gi.is_some() || gr.is_some();
        match modal {
            Modal::Rapid if moved => {
                let mut l = String::from("FASTABS");
                if let Some(v) = gx {
                    axis_word(&mut l, 'X', v);
                    x = v;
                }
                if let Some(v) = gy {
                    axis_word(&mut l, 'Y', v);
                    y = v;
                }
                if let Some(v) = gz {
                    axis_word(&mut l, 'Z', v);
                    z = v;
                }
                out.push(l);
            }
            Modal::Cut if moved => {
                let mut l = String::from("MOVEABS");
                if let Some(v) = gx {
                    axis_word(&mut l, 'X', v);
                    x = v;
                }
                if let Some(v) = gy {
                    axis_word(&mut l, 'Y', v);
                    y = v;
                }
                if let Some(v) = gz {
                    axis_word(&mut l, 'Z', v);
                    z = v;
                }
                out.push(l);
            }
            Modal::Cw | Modal::Ccw if moved || arc_words => {
                let cw = matches!(modal, Modal::Cw);
                let (ex, ey) = (gx.unwrap_or(x), gy.unwrap_or(y));
                let center = if let Some(r) = gr {
                    last_r = Some(r);
                    last_style_ij = false;
                    arc_center(x, y, ex, ey, r, cw).ok_or_else(|| err("R arc: no center fits (zero chord or R below half the chord)".to_string()))?
                } else if gi.is_some() {
                    last_style_ij = true;
                    let (di, dj) = if abs_ij {
                        (gi.unwrap() - x, gj.unwrap() - y)
                    } else {
                        (gi.unwrap(), gj.unwrap())
                    };
                    last_ij = (di, dj);
                    (x + di, y + dj)
                } else if last_style_ij || last_r.is_none() {
                    // modal arc repeat: Fanuc reuses the previous center vector,
                    // an omitted endpoint means a full circle back to start
                    (x + last_ij.0, y + last_ij.1)
                } else {
                    // modal repeat after an R arc: Fanuc repeats the R word
                    arc_center(x, y, ex, ey, last_r.unwrap(), cw).ok_or_else(|| err("modal R arc: no center fits".to_string()))?
                };
                let mut l = String::from(if cw { "CWABS" } else { "CCWABS" });
                axis_word(&mut l, 'X', ex);
                axis_word(&mut l, 'Y', ey);
                if let Some(v) = gz {
                    axis_word(&mut l, 'Z', v);
                }
                axis_word(&mut l, 'I', center.0);
                axis_word(&mut l, 'J', center.1);
                out.push(l);
                x = ex;
                y = ey;
                if let Some(v) = gz {
                    z = v;
                }
            }
            Modal::None if moved => return Err(err("coordinates before any motion command".to_string())),
            _ => {}
        }

        if program_end {
            break;
        }
    }

    if out.is_empty() {
        return Err("Fanuc-format input contains no machining commands.".to_string());
    }
    Ok(out.join("\n"))
}
