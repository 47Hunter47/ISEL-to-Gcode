//! Optional adaptive corner feed reduction — G1 (MOVEABS) toolpath only.
//!
//! The analysis is read-only: it walks the same ISEL stream the converter
//! walks, groups continuous XY cutting moves, and returns a *relative* feed
//! factor per processed input line. No coordinate is moved, merged or dropped;
//! the converter still emits every original point of every original move — only
//! the F value on an existing line can differ.
//!
//! Direction estimation asks for the chord `window_mm` (a `CornerParams` knob)
//! of *arc length* behind and ahead of a point, interpolated along the path.
//! Both the coordinate quantum and
//! the CAM's own rounding are far below that, so they cannot be read as geometry;
//! and because the chord follows the path rather than the CAM's vertices, a
//! 100 mm straight and a 0.3 mm polyline step are measured the same way.
//!
//! A sharp corner and a polyline arc both turn the direction; at a 1 mm chord a
//! small arc turns as much as a corner does. What separates them is where the
//! turn sits: a corner concentrates it at one point (the path a window away is
//! straight), an arc spreads it (the path a window away turns the same amount).
//! Corners get the angle profile; an arc gets ONE factor from its mean turn, so
//! it runs at a steady feed instead of a different F on every point.

use crate::convert::parse_coord;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Same-point threshold. Output coordinates carry 3 decimals.
const TOL_MM: f64 = 1e-3;
/// Reductions smaller than this are not worth an F command: they would litter
/// the file (and the machine) with 1 % feed steps.
const MIN_REDUCTION: f64 = 0.01;
/// Classification grid. Both the CAM's own coordinate rounding (0.001 mm) and the
/// cutter's practical resolution are far below this, so 0.01 mm is used as the
/// noise floor for classification — not as a claim about machine resolution.
/// Used for the analysis only — the emitted coordinates are untouched.
const SNAP_MM: f64 = 0.01;

/// The tunable policy of corner feed reduction: every number that used to be a
/// `const` in this file. They are policy, not physics — the defaults are what
/// measured well on real ISEL/CAM output. Serialized so the UI can edit each
/// one; whatever comes back goes through `cleaned()` before it is used.
#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(default = "CornerParams::default", rename_all = "camelCase")]
pub struct CornerParams {
    /// Chord length for direction estimation, mm of arc length. Chosen at the
    /// same order as the output coordinate resolution (1 µm) times a thousand:
    /// noise of a few µm is then worth a few hundredths of a degree, while a
    /// real turn inside this distance is still measured in tens of degrees.
    pub window_mm: f64,
    /// A turn smaller than this is noise; a curved section's mean is measured
    /// with the same floor, so a gentle arc stays at full feed.
    pub min_turn_deg: f64,
    /// A corner turns this much more than the path one window away. An arc
    /// turns a similar amount there, so it never passes.
    pub spike_ratio: f64,
    /// The half-way feed step out of a corner only pays off on the dense
    /// geometry around it. Segments are never split, so on a long straight it
    /// would slow down metres of unrelated path; above this length the feed
    /// returns in one step.
    pub ramp_max_seg_mm: f64,
    /// How far the segments flanking a corner climb back toward full feed:
    /// 0.5 is the half-way step, 0 steps straight back, 1 keeps them at the
    /// corner feed.
    pub ramp_mix: f64,
    /// A curved section at or above this radius keeps full feed; below it the
    /// feed follows `sqrt(r / arc_full_radius_mm)`. The square root is the
    /// shape of an acceleration limit (v = sqrt(a·r)) but the machine's
    /// acceleration is unknown, so this radius is a policy value, not a machine
    /// parameter. It makes a tight rounded U-turn slow (r=0.6 mm -> 27 %) and a
    /// broad one keep speed (r=50 mm -> 100 %).
    pub arc_full_radius_mm: f64,
    /// Angle profile: (turn degrees, feed factor) stops, linearly interpolated
    /// between them; the reversal floor is its sharpest factor.
    pub profile: Vec<(f64, f64)>,
}

impl Default for CornerParams {
    fn default() -> Self {
        CornerParams {
            window_mm: 1.0,
            min_turn_deg: 12.0,
            spike_ratio: 1.7,
            ramp_max_seg_mm: 4.0,
            ramp_mix: 0.5,
            arc_full_radius_mm: 8.0,
            profile: PROFILE.to_vec(),
        }
    }
}

fn num(v: f64, def: f64, lo: f64, hi: f64) -> f64 {
    if v.is_finite() {
        v.clamp(lo, hi)
    } else {
        def
    }
}

impl CornerParams {
    /// Clamp UI input to ranges the classifier survives: a non-finite or
    /// out-of-range number must not blind detection, and the profile always
    /// keeps a 0° head and a 180° tail so no lookup extrapolates above 1.
    pub fn cleaned(&self) -> CornerParams {
        let d = CornerParams::default();
        let mut profile: Vec<(f64, f64)> = self
            .profile
            .iter()
            .filter(|(a, f)| a.is_finite() && f.is_finite())
            .map(|(a, f)| (a.clamp(0.0, 180.0), f.clamp(0.01, 1.0)))
            .collect();
        profile.sort_by(|a, b| a.0.total_cmp(&b.0));
        profile.dedup_by(|a, b| a.0 == b.0);
        if profile.first().map_or(true, |r| r.0 > 0.0) {
            profile.insert(0, (0.0, 1.0));
        }
        if profile.last().map_or(true, |r| r.0 < 180.0) {
            let tail = profile.last().map_or(1.0, |r| r.1);
            profile.push((180.0, tail));
        }
        CornerParams {
            window_mm: num(self.window_mm, d.window_mm, 0.05, 20.0),
            min_turn_deg: num(self.min_turn_deg, d.min_turn_deg, 1.0, 90.0),
            spike_ratio: num(self.spike_ratio, d.spike_ratio, 1.0, 10.0),
            ramp_max_seg_mm: num(self.ramp_max_seg_mm, d.ramp_max_seg_mm, 0.0, 100.0),
            ramp_mix: num(self.ramp_mix, d.ramp_mix, 0.0, 1.0),
            arc_full_radius_mm: num(self.arc_full_radius_mm, d.arc_full_radius_mm, 0.05, 1000.0),
            profile,
        }
    }

    /// Sharpest factor the profile can hand out: the floor every rule shares.
    fn floor(&self) -> f64 {
        self.profile.iter().fold(1.0f64, |m, &(_, f)| m.min(f))
    }
}

/// Feed factor for a curved section of radius `r`: never above 1, never below
/// the sharpest stop of the angle profile.
fn radius_factor(r: f64, full_radius_mm: f64, floor: f64) -> f64 {
    if r >= full_radius_mm {
        1.0
    } else {
        (r / full_radius_mm).sqrt().clamp(floor, 1.0)
    }
}

/// Direction change (degrees) -> feed factor. 0 = straight, 180 = reversal.
/// Values below 10° are treated as noise, hence the flat 100 % head. The
/// frozen default profile the golden tests pin; conversions run on the
/// user-edited copy in `CornerParams`.
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

/// Feed factor for a direction change of `deg` degrees against the frozen
/// default profile, linearly interpolated between the stops (so 89.7° and
/// 90.3° stay within ~0.3 % of 0.40). Pinned by the tests; conversions use
/// `factor_with` with the user-edited profile.
pub fn factor_for(deg: f64) -> f64 {
    factor_with(deg, &PROFILE)
}

fn factor_with(deg: f64, profile: &[(f64, f64)]) -> f64 {
    let d = deg.clamp(0.0, 180.0);
    for w in profile.windows(2) {
        let ((d0, f0), (d1, f1)) = (w[0], w[1]);
        if d <= d1 {
            return f0 + (f1 - f0) * (d - d0) / (d1 - d0);
        }
    }
    profile.last().map(|w| w.1).unwrap_or(1.0)
}

/// Feed factors keyed by processed input line (the 1-based ordinal of a
/// non-blank, non-comment line, i.e. the same `processed` counter the converter
/// uses), plus how many corners and curved sections were found.
pub struct CornerPlan {
    pub factors: HashMap<u64, f64>,
    pub corners: usize,
    pub arcs: usize,
}

/// One continuous XY machining run. `pts[0]` is the position before the first
/// cutting move, so `lines[s]` is the input line of the move `pts[s] -> pts[s+1]`.
#[derive(Default)]
struct Group {
    pts: Vec<(f64, f64)>,
    lines: Vec<u64>,
}

/// The run as a path parameterised by arc length, so a direction can be asked
/// for at any distance along it — not only at the CAM's own vertices. That is
/// what makes a 100 mm straight and a 0.3 mm polyline step comparable: "1 mm
/// back" then always lands on the geometry itself.
struct Path {
    g: Vec<(f64, f64)>,
    /// cum[i] = arc length from the run's start to vertex i.
    cum: Vec<f64>,
    /// Chord half-length for turn/heading, from CornerParams.
    window: f64,
}

impl Path {
    fn new(g: Vec<(f64, f64)>, window: f64) -> Self {
        let mut cum = Vec::with_capacity(g.len());
        cum.push(0.0);
        for i in 1..g.len() {
            let d = (g[i].0 - g[i - 1].0).hypot(g[i].1 - g[i - 1].1);
            cum.push(cum[i - 1] + d);
        }
        Path { g, cum, window }
    }

    fn total(&self) -> f64 {
        *self.cum.last().unwrap_or(&0.0)
    }

    /// Point at arc length `s` (clamped to the run).
    fn at(&self, s: f64) -> (f64, f64) {
        let s = s.clamp(0.0, self.total());
        if s <= 0.0 {
            return self.g[0];
        }
        // last vertex whose arc length is <= s (clamped, so s == total is legal)
        let i = (self.cum.partition_point(|&c| c <= s) - 1).min(self.g.len() - 2);
        let (a, b) = (self.g[i], self.g[i + 1]);
        let seg = self.cum[i + 1] - self.cum[i];
        let f = if seg > 0.0 { (s - self.cum[i]) / seg } else { 0.0 };
        (a.0 + (b.0 - a.0) * f, a.1 + (b.1 - a.1) * f)
    }

    /// Signed direction change (degrees, + = left) at arc length `s`, from the
    /// chord `window` behind to the chord `window` ahead.
    fn turn(&self, s: f64) -> f64 {
        let b = self.at(s - self.window);
        let p = self.at(s);
        let a = self.at(s + self.window);
        let (ax, ay) = (p.0 - b.0, p.1 - b.1);
        let (bx, by) = (a.0 - p.0, a.1 - p.1);
        let (la, lb) = (ax.hypot(ay), bx.hypot(by));
        if la <= TOL_MM || lb <= TOL_MM {
            return 0.0; // no direction on one side: run end or a degenerate point
        }
        (ax * by - ay * bx).atan2(ax * bx + ay * by).to_degrees()
    }

    /// Heading (degrees) at arc length `s`, over a short chord. The difference
    /// between two headings is the turn of the path between them — counted once,
    /// unlike summing per-vertex window turns, whose windows overlap.
    fn heading(&self, s: f64) -> f64 {
        let b = self.at(s - self.window);
        let a = self.at(s + self.window);
        (a.1 - b.1).atan2(a.0 - b.0).to_degrees()
    }
}

/// Classification copy of the run: quantised to the machine's step, so CAM
/// rounding (0.001 mm) cannot show up as geometry. The output never sees this.
fn snap(g: &[(f64, f64)]) -> Vec<(f64, f64)> {
    g.iter()
        .map(|&(x, y)| {
            (
                (x / SNAP_MM).round() * SNAP_MM,
                (y / SNAP_MM).round() * SNAP_MM,
            )
        })
        .collect()
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
    ///
    /// A corner and a polyline arc both show a direction change, and at a 1 mm
    /// chord a small arc turns as sharply as a corner (this file's r=2.7 mm
    /// arcs read 22° per window). What separates them is *where* the turn sits:
    /// a corner is concentrated at one vertex, so the path a window away is
    /// straight; an arc turns a similar amount at the window point too.
    fn flush(&mut self, plan: &mut CornerPlan, cp: &CornerParams) {
        let pts = std::mem::take(&mut self.pts);
        let lines = std::mem::take(&mut self.lines);
        if pts.len() < 3 || lines.len() < 2 {
            return;
        }
        let p = Path::new(snap(&pts), cp.window_mm);
        let n = p.g.len();
        let t: Vec<f64> = (0..n).map(|i| p.turn(p.cum[i])).collect();
        let mut factors = vec![1.0f64; lines.len()];
        let mut curve = vec![false; n];
        for i in 1..n - 1 {
            // A repeated point is the same corner again: the window walk looks
            // through it, so it must not be counted (or applied) twice.
            if p.g[i] == p.g[i - 1] {
                continue;
            }
            let a = t[i].abs();
            if a <= cp.min_turn_deg {
                continue;
            }
            // How much the path turns one window away, in the same direction: a
            // curve continues, a corner does not.
            let s = p.cum[i];
            let cont = [s - cp.window_mm, s + cp.window_mm]
                .iter()
                .map(|&x| p.turn(x))
                .filter(|&b| b != 0.0 && (b > 0.0) == (t[i] > 0.0))
                .fold(0.0f64, |m, b| m.max(b.abs()));
            if a > cp.spike_ratio * cont {
                let f = factor_with(a, &cp.profile);
                if f >= 1.0 - MIN_REDUCTION {
                    continue;
                }
                plan.corners += 1;
                // The two segments meeting at the corner run at the corner feed;
                // the segments one further out climb back toward full feed by
                // ramp_mix, so the feed steps down and up instead of jumping.
                let half = 1.0 - (1.0 - f) * cp.ramp_mix;
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
                    if ramp && segment_len(&pts, s) > cp.ramp_max_seg_mm {
                        continue;
                    }
                    factors[s] = factors[s].min(w);
                }
            } else {
                curve[i] = true;
            }
        }
        // A curved section gets ONE factor from its mean turn, applied to the
        // segments the curve actually owns: a polyline arc then runs at a steady
        // feed instead of a different F on every 0.3 mm point, and the straight
        // that leads into it keeps full feed (its segment is not in the run).
        let mut i = 1;
        while i < n - 1 {
            if !curve[i] {
                i += 1;
                continue;
            }
            let start = i;
            while i + 1 < n - 1 && curve[i + 1] {
                i += 1;
            }
            let end = i;
            let mean = (start..=end).map(|k| t[k].abs()).sum::<f64>() / (end - start + 1) as f64;
            // Two views of the same section: how sharp each window step is (the
            // angle profile), and how tight the whole turn is. The section's total
            // turn is the difference of its end headings — summing the per-vertex
            // window turns would count the same degrees three times over, since
            // those windows overlap.
            let len = p.cum[end] - p.cum[start];
            // Unwrap successive local headings: a closed circle turns 360 deg,
            // which endpoint-only normalization aliases to zero.
            let headings: Vec<f64> = (start..=end)
                .map(|k| p.heading(p.cum[k]))
                .collect();
            let signed_total: f64 = headings
                .windows(2)
                .map(|w| ((w[1] - w[0] + 180.0).rem_euclid(360.0)) - 180.0)
                .sum();
            let total = signed_total.abs();
            let r = if total > 1e-6 { len / total.to_radians() } else { f64::INFINITY };
            // A curve never runs slower than a sharp corner that turns the same
            // total amount: rounding a corner must not make it slower than the
            // corner it replaced (a 90 deg turn spread over 0.1 mm steps is still
            // a 90 deg corner, and reads 0.40 here, not the radius floor).
            let f = factor_with(total, &cp.profile)
                .max(factor_with(mean, &cp.profile).min(radius_factor(r, cp.arc_full_radius_mm, cp.floor())));
            if f < 1.0 - MIN_REDUCTION {
                plan.arcs += 1;
                // The run's own segments, plus the one segment on each side that
                // carries the tangent point — but only when that segment is short:
                // a 12.6 mm straight leading into the arc keeps full feed.
                let lo = start.saturating_sub(1);
                for s in lo..=end.min(factors.len() - 1) {
                    let edge = s == lo || s == end;
                    if edge && segment_len(&pts, s) > cp.ramp_max_seg_mm {
                        continue;
                    }
                    factors[s] = factors[s].min(f);
                }
            }
            i += 1;
        }
        for (s, f) in factors.iter().enumerate() {
            if *f < 1.0 - 1e-9 {
                plan.factors.insert(lines[s], *f);
            }
        }
    }
}

/// Plan the whole file before any G-code is written, with the default policy.
pub fn analyze(text: &str) -> CornerPlan {
    analyze_with(text, &CornerParams::default())
}

/// Plan with `params` instead of the defaults; they are `cleaned()` here, so
/// callers may pass anything the UI produced.
pub fn analyze_with(text: &str, params: &CornerParams) -> CornerPlan {
    let p = params.cleaned();
    let floor = p.floor();
    let mut plan = CornerPlan {
        factors: HashMap::new(),
        corners: 0,
        arcs: 0,
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
            group.flush(&mut plan, &p);
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
            let xy_moved =
                (target.0 - last.0).abs() > TOL_MM || (target.1 - last.1).abs() > TOL_MM;
            if xy_moved {
                group.push(last, target, processed);
            } else if cc.contains_key(&'Z') {
                // X/Y repeated with only Z moving: that is a plunge, not a zero
                // length cut. It must never inherit a corner feed (the converter
                // times it in 3D), so the run ends here.
                group.flush(&mut plan, &p);
            }
            // A duplicate point with no Z change is neither: it is skipped and the
            // run continues, so it cannot break a corner in two.
            last = target;
        } else if line.starts_with("CWABS") || line.starts_with("CCWABS") {
            // An arc is one input line but many G1 moves in linearised output, so
            // it gets a single factor of its own, keyed by that line: a tight
            // rounded U-turn must slow down even in G1 mode.
            group.flush(&mut plan, &p);
            let cc = parse_coord(line, &['X', 'Y', 'Z']);
            let ij = parse_coord(line, &['I', 'J']);
            let end = (
                cc.get(&'X').copied().unwrap_or(last.0),
                cc.get(&'Y').copied().unwrap_or(last.1),
            );
            let center = (
                ij.get(&'I').copied().unwrap_or(0.0),
                ij.get(&'J').copied().unwrap_or(0.0),
            );
            let r = (last.0 - center.0).hypot(last.1 - center.1);
            if r > TOL_MM {
                let f = radius_factor(r, p.arc_full_radius_mm, floor);
                if f < 1.0 - MIN_REDUCTION {
                    plan.factors.insert(processed, f);
                    plan.arcs += 1;
                }
            }
            last = end;
        } else if line.starts_with("SPINDLE CW") {
            group.flush(&mut plan, &p);
        } else if line.starts_with("VEL") {
            // VEL deliberately does not split: the factors are relative to whatever
            // feed is in force when the line is emitted, so the geometry (and the
            // corners on it) continues across a feed change.
        } else {
            // Any other command (tool change, plane, end of program, an ISEL code
            // this build does not know) ends the run: two XY moves are never
            // bridged across a command whose meaning is not understood.
            group.flush(&mut plan, &p);
        }
    }
    group.flush(&mut plan, &p);
    plan
}

fn segment_len(pts: &[(f64, f64)], s: usize) -> f64 {
    (pts[s + 1].0 - pts[s].0).hypot(pts[s + 1].1 - pts[s].1)
}
