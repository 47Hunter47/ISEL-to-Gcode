//! Optional adaptive corner feed reduction — G1 (MOVEABS) toolpath only.
//!
//! The analysis is read-only: it walks the same ISEL stream the converter
//! walks, groups continuous XY cutting moves, and returns a *relative* feed
//! factor per processed input line. No coordinate is moved, merged or dropped;
//! the converter still emits every original point of every original move — only
//! the F value on an existing line can differ.
//!
//! Direction estimation uses the chord between the vertex and the first point
//! at least `WINDOW_MM` of path length away, on both sides. A distance chord
//! instead of "previous/next segment" is what makes input rounding (±1 µm, the
//! output quantum) and dense-segment noise vanish: the chord spans at least a
//! millimetre, so a micrometre deviation can only shift the estimated direction
//! by a fraction of a degree.

use crate::convert::parse_coord;
use std::collections::HashMap;

/// Chord length for direction estimation. Chosen at the same order as the
/// output coordinate resolution (1 µm) times a thousand: noise of a few µm is
/// then worth a few hundredths of a degree, while a real turn inside 1 mm of
/// path is still measured in tens of degrees.
const WINDOW_MM: f64 = 1.0;
/// Same-point threshold. Output coordinates carry 3 decimals.
const TOL_MM: f64 = 1e-3;
/// Ceiling on the window walk so a pathological file of 1 µm steps cannot turn
/// the analysis into O(n · window/step).
const MAX_WALK: usize = 512;
/// The half-way feed step out of a corner only pays off on the dense geometry
/// around it. Segments are never split, so on a long straight it would slow down
/// metres of unrelated path; above this length the feed returns in one step.
const RAMP_MAX_SEG_MM: f64 = 4.0;

/// Direction change (degrees) -> feed factor. 0 = straight, 180 = reversal.
/// Values below 10° are treated as noise, hence the flat 100 % head.
const PROFILE: [(f64, f64); 9] = [
    (0.0, 1.0),
    (10.0, 1.0),
    (30.0, 0.9),
    (45.0, 0.75),
    (60.0, 0.6),
    (90.0, 0.4),
    (120.0, 0.3),
    (150.0, 0.25),
    (180.0, 0.25),
];

/// Feed factor for a direction change of `deg` degrees, linearly interpolated
/// between the profile stops (so 89.7° and 90.3° stay within ~0.3 % of 0.40).
pub fn factor_for(deg: f64) -> f64 {
    let d = deg.clamp(0.0, 180.0);
    for w in PROFILE.windows(2) {
        let ((d0, f0), (d1, f1)) = (w[0], w[1]);
        if d <= d1 {
            return f0 + (f1 - f0) * (d - d0) / (d1 - d0);
        }
    }
    0.25
}

/// Feed factors keyed by processed input line (the 1-based ordinal of a
/// non-blank, non-comment line, i.e. the same `processed` counter the converter
/// uses), plus how many corners were found.
pub struct CornerPlan {
    pub factors: HashMap<u64, f64>,
    pub corners: usize,
}

/// One continuous XY machining run. `pts[0]` is the position before the first
/// cutting move, so `lines[s]` is the input line of the move `pts[s] -> pts[s+1]`.
#[derive(Default)]
struct Group {
    pts: Vec<(f64, f64)>,
    lines: Vec<u64>,
}

impl Group {
    fn push(&mut self, from: (f64, f64), to: (f64, f64), line: u64) {
        if self.pts.is_empty() {
            self.pts.push(from);
        }
        self.pts.push(to);
        self.lines.push(line);
    }

    /// Analyse the finished run. Only whole runs reach here: a corner is never
    /// decided from a fragment, whatever the file size, because there is no
    /// chunking at all (the run itself is the look-ahead window).
    fn flush(&mut self, plan: &mut CornerPlan) {
        let pts = std::mem::take(&mut self.pts);
        let lines = std::mem::take(&mut self.lines);
        if pts.len() < 3 || lines.len() < 2 {
            return;
        }
        let mut factors = vec![1.0f64; lines.len()];
        for i in 1..pts.len() - 1 {
            // A duplicate vertex is the same corner again: the window walk already
            // looks through it, so it must not be counted (or applied) twice.
            if (pts[i].0 - pts[i - 1].0).abs() <= TOL_MM && (pts[i].1 - pts[i - 1].1).abs() <= TOL_MM
            {
                continue;
            }
            let f = factor_for(turn_deg(&pts, i));
            if f >= 1.0 - 1e-9 {
                continue;
            }
            plan.corners += 1;
            // The two segments meeting at the corner run at the corner feed; the
            // segments one further out go half way back, so the feed steps down
            // and up instead of jumping.
            let half = 1.0 - (1.0 - f) * 0.5;
            for (s, w, ramp) in [
                (i as isize - 2, half, true),
                (i as isize - 1, f, false),
                (i as isize, f, false),
                (i as isize + 1, half, true),
            ] {
                if s < 0 || s as usize >= factors.len() {
                    continue;
                }
                let s = s as usize;
                if ramp && segment_len(&pts, s) > RAMP_MAX_SEG_MM {
                    continue;
                }
                factors[s] = factors[s].min(w);
            }
        }
        for (s, f) in factors.iter().enumerate() {
            if *f < 1.0 - 1e-9 {
                plan.factors.insert(lines[s], *f);
            }
        }
    }
}

/// Plan the whole file before any G-code is written.
pub fn analyze(text: &str) -> CornerPlan {
    let mut plan = CornerPlan {
        factors: HashMap::new(),
        corners: 0,
    };
    let mut group = Group::default();
    // XY only: Z never participates in the geometry, it can only end a run.
    let mut last = (0.0f64, 0.0f64);
    let mut processed: u64 = 0;
    for raw in text.lines() {
        let line = raw.trim();
        if line.is_empty() || line.starts_with(';') {
            continue;
        }
        processed += 1;
        if line.starts_with("FASTABS") {
            group.flush(&mut plan);
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            if let Some(v) = cc.get(&'X') {
                last.0 = *v;
            }
            if let Some(v) = cc.get(&'Y') {
                last.1 = *v;
            }
        } else if line.starts_with("MOVEABS") {
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            let target = (
                cc.get(&'X').copied().unwrap_or(last.0),
                cc.get(&'Y').copied().unwrap_or(last.1),
            );
            if cc.contains_key(&'X') || cc.contains_key(&'Y') {
                group.push(last, target, processed);
            } else {
                // Pure Z (or empty) MOVEABS: not a cutting segment, ends the run.
                group.flush(&mut plan);
            }
            last = target;
        } else if line.starts_with("CWABS") || line.starts_with("CCWABS") {
            // Arcs are excluded from corner analysis; their end point becomes the
            // start of the next run.
            group.flush(&mut plan);
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            if let Some(v) = cc.get(&'X') {
                last.0 = *v;
            }
            if let Some(v) = cc.get(&'Y') {
                last.1 = *v;
            }
        } else if line.starts_with("SPINDLE CW") {
            group.flush(&mut plan);
        } else if line.starts_with("VEL") {
            // VEL deliberately does not split: the factors are relative to whatever
            // feed is in force when the line is emitted, so the geometry (and the
            // corners on it) continues across a feed change.
        } else {
            // Any other command (tool change, plane, end of program, an ISEL code
            // this build does not know) ends the run: two XY moves are never
            // bridged across a command whose meaning is not understood.
            group.flush(&mut plan);
        }
    }
    group.flush(&mut plan);
    plan
}

fn segment_len(pts: &[(f64, f64)], s: usize) -> f64 {
    (pts[s + 1].0 - pts[s].0).hypot(pts[s + 1].1 - pts[s].1)
}

fn turn_deg(pts: &[(f64, f64)], i: usize) -> f64 {
    let back = walk(pts, i, true);
    let ahead = walk(pts, i, false);
    let (ax, ay) = (pts[i].0 - back.0, pts[i].1 - back.1);
    let (bx, by) = (ahead.0 - pts[i].0, ahead.1 - pts[i].1);
    let (la, lb) = (ax.hypot(ay), bx.hypot(by));
    if la <= TOL_MM || lb <= TOL_MM {
        return 0.0; // one side has no direction: group end or a degenerate point
    }
    ((ax * bx + ay * by) / (la * lb))
        .clamp(-1.0, 1.0)
        .acos()
        .to_degrees()
}

/// First point at least `WINDOW_MM` of path length away from `pts[from]`, or the
/// run's end. Zero-length steps add nothing, so repeated coordinates fall
/// through to the next real point.
fn walk(pts: &[(f64, f64)], from: usize, back: bool) -> (f64, f64) {
    let mut acc = 0.0;
    let mut j = from;
    for _ in 0..MAX_WALK {
        let next = if back {
            j.checked_sub(1)
        } else {
            (j + 1 < pts.len()).then_some(j + 1)
        };
        let Some(n) = next else { break };
        acc += (pts[n].0 - pts[j].0).hypot(pts[n].1 - pts[j].1);
        j = n;
        if acc >= WINDOW_MM {
            break;
        }
    }
    pts[j]
}
