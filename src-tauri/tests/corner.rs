// Corner feed reduction: the toolpath is immutable, only F values may change.
use isel_to_gcode::convert::convert_file;
use isel_to_gcode::corner::{CornerPlan, analyze, factor_for};

/// Exact base feed of `VEL 16667`: the reference converter divides by 16.6667,
/// so the true value is 1000.018 — printed as F1000, and *time is the printed
/// value* in corner mode (the machine executes what the file says).
const FEED: f64 = 16667.0 / 16.6667;

/// ISEL program from an XY polyline (mm). Line 1 = FASTABS to the start point,
/// line 2 = VEL, line 3+ = one MOVEABS per point, so the move ending at
/// `pts[k+1]` (segment k) is processed line `k + 3`.
fn prog(pts: &[(f64, f64)]) -> String {
    let mut s = format!(
        "FASTABS X{:.0} Y{:.0}\nVEL 16667\n",
        pts[0].0 * 1000.0,
        pts[0].1 * 1000.0
    );
    for (x, y) in &pts[1..] {
        s.push_str(&format!("MOVEABS X{:.0} Y{:.0}\n", x * 1000.0, y * 1000.0));
    }
    s
}

/// Relative feed of segment `pts[s] -> pts[s+1]` (1.0 = untouched).
fn seg(plan: &CornerPlan, s: usize) -> f64 {
    plan.factors.get(&(s as u64 + 3)).copied().unwrap_or(1.0)
}

fn run(text: &str, arc: bool, corner: bool) -> (String, Vec<String>, f64) {
    static N: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    let id = N.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let dir = std::env::temp_dir();
    let inp = dir.join(format!("isel_c_{}_{}.isel", std::process::id(), id));
    let outp = dir.join(format!("isel_c_{}_{}.ngc", std::process::id(), id));
    std::fs::write(&inp, text).unwrap();
    let mut logs: Vec<String> = vec![];
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    let r = convert_file(
        inp.to_str().unwrap(),
        outp.to_str().unwrap(),
        arc,
        corner,
        &mut |m: &str| logs.push(m.to_string()),
        &mut none,
    );
    let ngc = std::fs::read_to_string(&outp).unwrap_or_default();
    std::fs::remove_file(&inp).ok();
    std::fs::remove_file(&outp).ok();
    (ngc, logs, r.unwrap_or(-1.0))
}

fn fixture(name: &str) -> String {
    std::fs::read_to_string(format!(
        "{}/../tests/fixtures/{}",
        env!("CARGO_MANIFEST_DIR"),
        name
    ))
    .unwrap()
}

/// G1 lines with any F suffix removed: every movement and coordinate.
fn g1_moves(ngc: &str) -> Vec<String> {
    ngc.lines()
        .filter(|l| l.contains(" G1"))
        .map(|l| match l.rsplit_once(" F") {
            Some((h, t)) if t.bytes().all(|b| b.is_ascii_digit()) => h.to_string(),
            _ => l.to_string(),
        })
        .collect()
}

/// Every line with its F words dropped: N number, command and coordinates must
/// match the original converter exactly; only F words may be added, changed or
/// dropped.
fn structure(ngc: &str) -> Vec<String> {
    ngc.lines()
        .map(|l| {
            l.split_whitespace()
                .filter(|t| !t.starts_with('F'))
                .collect::<Vec<_>>()
                .join(" ")
        })
        .collect()
}

/// The immutability contract: same lines, same numbering, same commands, same
/// coordinates — only F words differ.
fn assert_only_feeds_differ(off: &str, on: &str) {
    assert_eq!(off.lines().count(), on.lines().count(), "line count changed");
    assert_eq!(
        g1_moves(off),
        g1_moves(on),
        "a G1 movement or coordinate changed"
    );
    assert_eq!(structure(off), structure(on), "something other than F differs");
}

fn f_values(ngc: &str) -> Vec<f64> {
    ngc.lines()
        .filter_map(|l| l.rsplit_once(" F"))
        .filter_map(|(_, t)| t.parse::<f64>().ok())
        .collect()
}

// ── profile ────────────────────────────────────────────────────────────────
#[test]
fn profile_is_monotonic_and_smooth() {
    assert_eq!(factor_for(0.0), 1.0);
    assert_eq!(factor_for(10.0), 1.0);
    assert!((factor_for(90.0) - 0.4).abs() < 1e-12);
    assert!((factor_for(180.0) - 0.25).abs() < 1e-12);
    let mut prev = 1.0;
    for d in 0..=1800 {
        let f = factor_for(d as f64 / 10.0);
        assert!(f <= prev + 1e-12, "not monotonic at {} deg", d);
        assert!(prev - f < 0.02, "jump at {} deg", d);
        prev = f;
    }
    // 89.7 / 90.0 / 90.3 must be indistinguishable
    assert!((factor_for(89.7) - factor_for(90.3)).abs() < 0.005);
}

// ── geometric classification ───────────────────────────────────────────────
#[test]
fn straight_line_has_no_corners() {
    let p = analyze(&prog(&[(0.0, 0.0), (10.0, 0.0), (20.0, 0.0), (30.0, 0.0)]));
    assert_eq!(p.corners, 0);
    assert!(p.factors.is_empty());
}

#[test]
fn noisy_straight_line_has_no_corners() {
    // the spec's example: 10 mm segments with micrometre deviations
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.002),
        (20.0, -0.001),
        (30.0, 0.001),
        (40.0, 0.0),
    ]));
    assert_eq!(p.corners, 0);
}

#[test]
fn accumulated_rounding_error_has_no_corners() {
    // 400 dense steps with sub-micrometre wobble: no corner anywhere
    let pts: Vec<(f64, f64)> = (0..400)
        .map(|i| {
            let x = i as f64 * 0.25;
            (x, (x * 1.7).sin() * 0.0004)
        })
        .collect();
    assert_eq!(analyze(&prog(&pts)).corners, 0);
}

fn corner_factor(out_deg: f64) -> f64 {
    let a = out_deg.to_radians();
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.0),
        (10.0 + 10.0 * a.cos(), 10.0 * a.sin()),
    ]));
    assert_eq!(p.corners, 1, "{} deg", out_deg);
    seg(&p, 0)
}

#[test]
fn corner_severity_follows_the_profile() {
    for (deg, want) in [
        (45.0, 0.75),
        (60.0, 0.6),
        (90.0, 0.4),
        (120.0, 0.3),
        (180.0, 0.25),
    ] {
        let f = corner_factor(deg);
        assert!((f - want).abs() < 0.03, "{} deg -> {} want {}", deg, f, want);
    }
}

#[test]
fn imperfect_right_angles_agree() {
    let f = |dx: f64| {
        let p = analyze(&prog(&[(0.0, 0.0), (10.0, 0.0), (10.0 + dx, 10.0)]));
        seg(&p, 0)
    };
    let (a, b, c) = (f(-0.3), f(0.0), f(0.3));
    let spread = a.max(b).max(c) - a.min(b).min(c);
    assert!(spread < 0.03, "{} {} {} spread {}", a, b, c, spread);
}

#[test]
fn corner_slows_the_incoming_and_the_outgoing_segment() {
    // one corner at (100,0): the 100 mm segment that arrives at it and the 10 mm
    // segment that leaves it are both cut; the 100 mm segment after that is not
    let p = analyze(&prog(&[(0.0, 0.0), (100.0, 0.0), (100.0, 10.0), (100.0, 110.0)]));
    assert_eq!(p.corners, 1);
    assert!((seg(&p, 0) - 0.4).abs() < 0.02, "incoming: {}", seg(&p, 0));
    assert!((seg(&p, 1) - 0.4).abs() < 0.02, "outgoing: {}", seg(&p, 1));
    assert_eq!(seg(&p, 2), 1.0, "feed restored past the influence region");
}

#[test]
fn short_segments_ramp_the_feed_back_up() {
    // 1 mm segments leaving a 90 deg corner: full corner feed on the two segments
    // that meet there, half way on the next, then untouched. A segment is never
    // split, so a long one is never given a half-way value (see the test above).
    let mut pts = vec![(0.0, 0.0), (10.0, 0.0)];
    pts.extend((1..=6).map(|i| (10.0, i as f64)));
    let p = analyze(&prog(&pts));
    assert_eq!(p.corners, 1);
    assert!((seg(&p, 1) - 0.4).abs() < 0.02, "meeting the corner");
    assert!((seg(&p, 2) - 0.7).abs() < 0.02, "ramp back");
    assert_eq!(seg(&p, 4), 1.0, "and then full feed");
}

#[test]
fn consecutive_corners_take_the_strongest_reduction() {
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.0),
        (10.0, 10.0),
        (20.0, 10.0),
        (20.0, 20.0),
    ]));
    assert_eq!(p.corners, 3);
    // segment 2 sits between two 90 deg corners
    assert!(seg(&p, 2) < 0.45, "{}", seg(&p, 2));
}

#[test]
fn gentle_curve_stays_at_full_feed() {
    // r = 50 mm polyline, 0.25 mm steps: a curve, not a corner series
    let (r, n) = (50.0, 240);
    let pts: Vec<(f64, f64)> = (0..=n)
        .map(|i| {
            let a = i as f64 * 0.005;
            (r * a.sin(), r - r * a.cos())
        })
        .collect();
    assert_eq!(analyze(&prog(&pts)).corners, 0);
}

#[test]
fn small_feature_is_detected_not_erased() {
    // 1 mm square notch in a straight path: four right angles 1 mm apart. Pairs
    // of them turn the same way, so they may be reported as one curved section
    // instead of two corners — what matters is that every notch segment slows.
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.0),
        (10.0, 1.0),
        (11.0, 1.0),
        (11.0, 0.0),
        (20.0, 0.0),
    ]));
    assert!(p.corners + p.arcs >= 2, "corners {} arcs {}", p.corners, p.arcs);
    for s in 1..=3 {
        assert!(seg(&p, s) <= 0.45, "notch segment {} runs at {}", s, seg(&p, s));
    }
}

#[test]
fn very_short_segments_still_form_a_corner() {
    let mut pts: Vec<(f64, f64)> = (0..=100).map(|i| (i as f64 * 0.1, 0.0)).collect();
    pts.extend((1..=100).map(|i| (10.0, i as f64 * 0.1)));
    let p = analyze(&prog(&pts));
    let near = (96..=104).map(|s| seg(&p, s)).fold(1.0f64, f64::min);
    assert!((near - 0.4).abs() < 0.03, "corner feed was {}", near);
}

#[test]
fn zero_length_and_repeated_points_are_harmless() {
    // A repeated point is the same corner again, and it gets no factor of its own:
    // a line that does not move in XY must never be a carrier of a reduced feed.
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.0),
        (10.0, 0.0),
        (10.0, 10.0),
    ]));
    assert_eq!(p.corners, 1, "the duplicate is the same corner, counted once");
    assert_eq!(seg(&p, 0), 0.4, "segment into the corner");
    assert_eq!(seg(&p, 1), 1.0, "the duplicate line carries nothing");
    assert_eq!(seg(&p, 2), 0.4, "segment out of the corner");

    let rr = analyze(&prog(&[(5.0, 5.0), (5.0, 5.0), (5.0, 5.0)]));
    assert!(rr.factors.is_empty());
}

// ── global analysis: the same local geometry, wherever it sits ──────────────
/// A straight +X run of `n` 10 mm segments ending in a single 10 mm move up:
/// exactly one corner, on the last two segments (input lines n+2 and n+3).
fn corner_at(n: usize) -> CornerPlan {
    let mut pts: Vec<(f64, f64)> = (0..=n).map(|i| (i as f64 * 10.0, 0.0)).collect();
    pts.push((n as f64 * 10.0, 10.0));
    analyze(&prog(&pts))
}

#[test]
fn corner_decision_is_position_independent() {
    // the same corner after 8, after 4998 (lines 5000/5001) and after 12000
    // segments must give the same factors: there is no chunk boundary to split
    // a decision, the whole run is the analysis window
    for n in [8usize, 4998, 12000] {
        let p = corner_at(n);
        let want = [(n as u64 + 2, 0.4), (n as u64 + 3, 0.4)];
        assert_eq!(p.corners, 1, "n={}", n);
        assert_eq!(p.factors.len(), 2, "n={} {:?}", n, p.factors);
        for (line, f) in want {
            let got = p.factors.get(&line).copied().unwrap_or(-1.0);
            assert!((got - f).abs() < 1e-9, "n={} line {} -> {}", n, line, got);
        }
    }
}

#[test]
fn corner_straddling_line_5000_behaves_like_any_other() {
    let p = corner_at(4998);
    assert!(p.factors.contains_key(&5000) && p.factors.contains_key(&5001));
    let small = corner_at(8);
    // identical relative decision, not merely a similar one
    assert_eq!(p.factors.values().copied().fold(9.0f64, f64::min),
               small.factors.values().copied().fold(9.0f64, f64::min));
}

#[test]
fn corner_then_long_straight_returns_to_full_feed() {
    let p = analyze(&prog(&[
        (0.0, 0.0),
        (10.0, 0.0),
        (10.0, 10.0),
        (10.0, 40.0),
    ]));
    assert_eq!(seg(&p, 2), 1.0, "a 30 mm straight is not corner territory");
}

#[test]
fn rapid_breaks_the_run() {
    let p = analyze(
        "FASTABS X0 Y0\nVEL 16667\nMOVEABS X10000 Y0\nFASTABS X10000 Y10000\nMOVEABS X20000 Y10000\n",
    );
    assert_eq!(p.corners, 0, "a rapid reposition is not a machining corner");
    assert!(p.factors.is_empty());
}

#[test]
fn spindle_command_ends_the_run() {
    let p = analyze(
        "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nSPINDLE CW RPM1200\nMOVEABS X10000 Y10000\n",
    );
    assert_eq!(p.corners, 0);
}

#[test]
fn separate_runs_never_share_a_corner() {
    // run 1 has two right angles, run 2 is straight and inherits nothing
    let text = "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nMOVEABS X10000 Y10000\n\
MOVEABS X30000 Y10000\nFASTABS Z4000\nFASTABS X60000 Y0\nMOVEABS X70000 Y0\nMOVEABS X80000 Y0\n";
    let p = analyze(text);
    assert_eq!(p.corners, 2);
    assert!(p.factors.keys().all(|k| (2..=5).contains(k)), "{:?}", p.factors);
}

#[test]
fn unknown_command_ends_the_run() {
    let p = analyze(
        "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nGETTOOL 1\nMOVEABS X10000 Y10000\n",
    );
    assert_eq!(p.corners, 0);
}

// ── Z axis ────────────────────────────────────────────────────────────────
#[test]
fn z_only_moves_never_create_corners() {
    let p = analyze(
        "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X10000 Y0 Z-1000\nMOVEABS Z-2000\nMOVEABS X10000 Y10000 Z-2000\n",
    );
    assert_eq!(p.corners, 0);
}

#[test]
fn xyz_motion_is_judged_on_the_xy_projection() {
    let p = analyze(
        "VEL 16667\nMOVEABS X0 Y0 Z0\nMOVEABS X10000 Y0 Z-1000\nMOVEABS X10000 Y10000 Z-3000\n",
    );
    assert_eq!(p.corners, 1);
    assert!((seg(&p, 0) - 0.4).abs() < 0.02);
}

#[test]
fn repeated_xy_with_only_z_moving_is_never_a_corner_and_never_reduced() {
    // The regression: X/Y repeated on the line but only Z moves. It used to be a
    // zero-length XY segment that a nearby corner could hand a factor to, while
    // the converter times the move in 3D — so the plunge ran at corner feed.
    let text = "VEL 16667\nMOVEABS X0 Y0 Z0\nMOVEABS X100000 Y0 Z0\n\
MOVEABS X100000 Y0 Z-500\nMOVEABS X100000 Y10000 Z-500\n";
    let p = analyze(text);
    assert_eq!(p.corners, 0, "the run ends at the plunge: {:?}", p.factors);
    assert!(p.factors.is_empty());

    let (on, _, _) = run(text, false, true);
    let (off, _, _) = run(text, false, false);
    assert_only_feeds_differ(&off, &on);
    assert_eq!(on.matches(" F").count(), 1, "only the VEL F:\n{}", on);
}

#[test]
fn short_segment_straight_line_keeps_full_feed() {
    // 200 steps of 0.05 mm with a micrometre of wobble: dense, still straight
    let pts: Vec<(f64, f64)> = (0..200)
        .map(|i| (i as f64 * 0.05, if i % 3 == 0 { 0.001 } else { -0.001 }))
        .collect();
    let p = analyze(&prog(&pts));
    assert_eq!(p.corners, 0);
    assert!(p.arcs == 0, "no curved section either: {:?}", p.factors);
}

#[test]
fn cwabs_and_ccwabs_are_analysed() {
    // G1 mode linearises arcs; a tight rounded U-turn must slow down too.
    // r = 0.5 mm: sqrt(0.5/8) = 0.25 (the floor).
    let p = analyze("VEL 16667\nMOVEABS X0 Y0 Z-100\nCCWABS X1000 Y0 I500 J0\n");
    let f = p.factors.values().copied().next().expect("the arc is reduced");
    assert!((f - 0.25).abs() < 0.01, "tight arc factor {}", f);
    assert_eq!(p.factors.len(), 1, "one factor for the whole arc");

    // A broad arc keeps full feed: nothing to say.
    let wide = analyze("VEL 16667\nMOVEABS X0 Y0 Z-100\nCCWABS X100000 Y0 I50000 J0\n");
    assert!(wide.factors.is_empty(), "{:?}", wide.factors);
}

#[test]
fn tight_rounded_u_turn_slows_more_than_a_wide_one() {
    let u = |r: f64| {
        // half circle of radius r, cut as a polyline: 60 segments
        let n = 60;
        let pts: Vec<(f64, f64)> = (0..=n)
            .map(|i| {
                let a = std::f64::consts::PI * i as f64 / n as f64;
                (r * a.sin(), r * (1.0 - a.cos()))
            })
            .collect();
        let p = analyze(&prog(&pts));
        let f = p.factors.values().copied().fold(1.0f64, f64::min);
        (p, f)
    };
    let (tight, ft) = u(1.5);
    let (wide, fw) = u(40.0);
    assert_eq!(tight.arcs, 1, "one curved section");
    assert!(ft < 0.55, "tight r=1.5 U-turn at {}", ft);
    // Wide enough that no 1 mm window turns more than the noise floor: nothing
    // to reduce, which is the point (a broad U-turn keeps speed).
    assert!(wide.factors.is_empty(), "{:?}", wide.factors);
    assert_eq!(fw, 1.0);
}

#[test]
fn a_curve_is_never_slower_than_the_corner_it_rounds() {
    // Same 90 deg turn, three ways: sharp, rounded over 0.1 mm steps, rounded
    // over 10 mm steps. Rounding must not make the machine slower than the
    // sharp corner, so the sharp one is the floor.
    let sharp = analyze(&prog(&[(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)]));
    let fs = seg(&sharp, 0);
    assert!((fs - 0.4).abs() < 0.02, "sharp 90 deg {}", fs);

    let steps: Vec<(f64, f64)> = (0..=100).map(|i| (i as f64 * 0.1, 0.0)).collect();
    let mut rounded = steps.clone();
    rounded.extend((1..=100).map(|i| (10.0, i as f64 * 0.1)));
    let fr = (96..=104)
        .map(|s| seg(&analyze(&prog(&rounded)), s))
        .fold(1.0f64, f64::min);
    assert!(fr >= fs - 1e-9, "rounded {} slower than sharp {}", fr, fs);
    assert!(fr < 0.9, "still slowed: {}", fr);
}

#[test]
fn pure_z_line_never_carries_a_reduced_feed() {
    // right angle, then a plunge (Z only), then the next XY cut
    let text = "VEL 16667\nMOVEABS X0 Y0 Z0\nMOVEABS X10000 Y0 Z-100\n\
MOVEABS X10000 Y10000 Z-100\nMOVEABS Z-500\nMOVEABS X10000 Y20000 Z-500\n";
    let (on, _, _) = run(text, false, true);
    let z: Vec<&str> = on.lines().filter(|l| l.ends_with(" G1 Z-0.500 F1000")).collect();
    assert_eq!(z.len(), 1, "the plunge runs at the base feed:\n{}", on);
    // base feed already in force, so the next cut carries no F at all
    let next = on.lines().find(|l| l.contains("X10.000 Y20.000")).unwrap();
    assert!(!next.contains(" F"), "{}", next);
}

// ── output: geometry untouched, only F changes ─────────────────────────────
const ZIGZAG: &str = "GETTOOL 0\nSPINDLE CW RPM1600\nFASTABS Z4000\nFASTABS X0 Y0\nFASTABS Z-500\n\
VEL 16667\nMOVEABS X30000 Y0\nMOVEABS X30000 Y20000\nMOVEABS Z-1500\nMOVEABS X0 Y20000\n\
VEL 25000\nMOVEABS X0 Y40000\nFASTABS Z4000\n";

#[test]
fn corner_off_is_byte_identical_to_the_original_converter() {
    // no corner in the program: the flag on must not change one byte either
    let flat = "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nMOVEABS X20000 Y0\n";
    let (on, _, t_on) = run(flat, false, true);
    let (off, _, t_off) = run(flat, false, false);
    assert_eq!(on, off);
    assert_eq!(t_on, t_off);

    // and the flag off keeps the golden bytes
    let (off, _, _) = run(&fixture("sample.isel"), false, false);
    let want = std::fs::read(format!(
        "{}/../tests/fixtures/sample_expected.g1.ngc",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap();
    assert_eq!(off.into_bytes(), want);
}

#[test]
fn geometry_is_identical_with_and_without_corner_feed() {
    let (off, _, _) = run(ZIGZAG, false, false);
    let (on, logs, _) = run(ZIGZAG, false, true);
    assert_only_feeds_differ(&off, &on);
    assert!(logs.iter().any(|l| l.starts_with("Corner feed reduction: 2 corner")));
}

#[test]
fn real_programs_keep_their_lines_and_coordinates() {
    for name in ["sample", "demo1", "demo2"] {
        let text = fixture(&format!("{}.isel", name));
        let (off, _, _) = run(&text, false, false);
        let (on, _, _) = run(&text, false, true);
        assert_only_feeds_differ(&off, &on);
    }
}

#[test]
fn linearised_arcs_keep_their_geometry() {
    let text = fixture("sample.isel");
    let (off, _, _) = run(&text, false, false);
    let (on, _, _) = run(&text, false, true);
    assert_eq!(g1_moves(&off), g1_moves(&on));
    assert_only_feeds_differ(&off, &on);
}

#[test]
fn linearised_arc_after_a_corner_starts_at_base_feed() {
    // an arc polyline prints no F of its own: the restore rides on its first
    // line, so the whole arc cannot inherit the corner feed
    let text = "VEL 16667\nMOVEABS X0 Y0 Z-1000\nMOVEABS X10000 Y0 Z-1000\n\
MOVEABS X10000 Y10000 Z-1000\nCWABS X20000 Y10000 I15000 J10000\n";
    let (on, _, _) = run(text, false, true);
    let (off, _, _) = run(text, false, false);
    assert_only_feeds_differ(&off, &on);

    let lines: Vec<&str> = on.lines().collect();
    let corner = lines.iter().position(|l| l.contains("Y10.000")).unwrap();
    assert_eq!(lines[corner - 1].matches(" F400").count(), 1, "corner line");
    // The arc (r = 5 mm) now carries its own factor, sqrt(5/8) = 0.79, printed
    // once on its first G1 line; the corner's F400 must not leak into it.
    let first_arc = lines[corner + 1];
    assert!(first_arc.contains(" G1 X"), "{}", first_arc);
    assert!(
        first_arc.ends_with(" F791"),
        "arc runs at its own curvature feed: {}",
        first_arc
    );
    assert_eq!(first_arc.matches(" F").count(), 1, "one F for the whole arc");
    // every later arc line stays silent: the feed is already in force
    assert!(
        lines[corner + 2..].iter().take(20).all(|l| !l.contains(" F")),
        "no F per arc step"
    );
}

#[test]
fn feed_never_exceeds_the_vel_feed() {
    let (on, _, _) = run(ZIGZAG, false, true);
    let feeds = f_values(&on);
    assert!(!feeds.is_empty());
    assert!(feeds.iter().all(|f| *f <= 1500.0), "{:?}", feeds);
}

#[test]
fn rounded_slot_arcs_are_curves_not_corners() {
    // Mert's real slot program: two 12.6 mm straights joined by r≈2.7 mm end
    // arcs, cut twice. At a 1 mm chord a vertex on that arc reads ~22°, i.e.
    // more than the corner floor — the reported defect was exactly that: 128
    // "corners", an F on every line, feed wandering between F645 and F940.
    let p = analyze(&fixture("slot.isel"));
    assert_eq!(p.corners, 0, "a rounded slot has no sharp corner");
    assert!(p.arcs >= 2, "end arcs: {}", p.arcs);
    let mut vals: Vec<f64> = p.factors.values().copied().collect();
    vals.sort_by(|a, b| a.partial_cmp(b).unwrap());
    vals.dedup_by(|a, b| (*a - *b).abs() < 1e-9);
    // one factor per arc run (four here: two caps, cut twice), never one per vertex
    assert!(vals.len() <= p.arcs, "{} factors for {} arcs: {:?}", vals.len(), p.arcs, vals);
    assert!(
        vals.last().unwrap() - vals.first().unwrap() < 0.06,
        "the same geometry must give the same arc feed: {:?}",
        vals
    );
    // A 2.7 mm radius cap is a real U-turn: sqrt(2.7/8) is the curvature ceiling,
    // far below the ~6 % a per-window angle would give.
    assert!(
        vals.iter().all(|f| *f > 0.5 && *f < 0.72),
        "a 2.7 mm U-turn slows by radius, got {:?}",
        vals
    );
}

#[test]
fn rounded_slot_straights_keep_full_feed() {
    let text = fixture("slot.isel");
    let (on, logs, t_on) = run(&text, false, true);
    let (off_line, _, t_off) = run(&text, false, false);

    // 137 G1 moves; only the arcs' own segments and their restores carry an F
    let g1 = on.lines().filter(|l| l.contains(" G1")).count();
    let with_f = on.lines().filter(|l| l.contains(" G1") && l.contains(" F")).count();
    assert!(with_f > 0 && with_f < g1 / 5, "{} of {} G1 lines", with_f, g1);
    assert_eq!(g1, off_line.lines().filter(|l| l.contains(" G1")).count());

    // the straights are untouched; only the arcs' share of the path slows down
    assert!(t_on > t_off && t_on / t_off < 1.25, "{} vs {}", t_on, t_off);
    assert!(logs.iter().any(|l| l.contains("0 corner(s)")));
}

#[test]
fn corner_reduces_feed_and_restores_it() {
    // 100 mm approach, 10 mm exit, then a 40 mm straight
    let text = "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X100000 Y0\nMOVEABS X100000 Y10000\nMOVEABS X100000 Y50000\n";
    let (on, _, t_on) = run(text, false, true);
    let (_, _, t_off) = run(text, false, false);
    assert_eq!(on.matches(" F400").count(), 1, "one reduced F, no repeats");
    assert_eq!(on.matches(" F1000").count(), 2, "VEL + restore");
    // 110 mm at the printed F400, then the restore line keeps the exact VEL feed
    assert!((t_on - (110.0 / 400.0 + 40.0 / FEED)).abs() < 1e-9, "{}", t_on);
    assert!((t_off - 150.0 / FEED).abs() < 1e-9, "{}", t_off);
}

#[test]
fn cycle_time_uses_the_effective_feed() {
    let text = "VEL 16667\nMOVEABS X0 Y0\nMOVEABS X100000 Y0\nMOVEABS X100000 Y100000\n";
    let (_, _, on) = run(text, false, true);
    let (_, _, off) = run(text, false, false);
    assert!((off - 200.0 / FEED).abs() < 1e-9, "{}", off);
    assert!((on - (100.0 / 400.0 + 100.0 / 400.0)).abs() < 1e-9, "{}", on);
}

#[test]
fn vel_change_before_a_corner_sets_the_reference() {
    let text = "VEL 33333\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nMOVEABS X10000 Y10000\n";
    let (on, _, _) = run(text, false, true);
    assert!(on.contains(" F2000"), "{}", on); // base feed
    assert!(on.contains(" F800"), "{}", on); // 0.4 * 2000
}

#[test]
fn vel_change_after_a_corner_restores_the_new_feed() {
    // corner, then a 10 mm exit; VEL drops the feed before the next cut
    let text = "VEL 33333\nMOVEABS X0 Y0\nMOVEABS X10000 Y0\nMOVEABS X10000 Y5000\n\
VEL 16667\nMOVEABS X10000 Y15000\n";
    let (on, _, _) = run(text, false, true);
    assert!(on.contains(" F800"), "{}", on);
    assert!(on.contains("F1000"), "{}", on);
    let after: Vec<&str> = on.lines().filter(|l| l.contains("Y15.000")).collect();
    assert_eq!(after.len(), 1);
    assert!(
        !after[0].contains(" F"),
        "VEL already put the feed back: {}",
        after[0]
    );
}

#[test]
fn arc_mode_is_untouched_by_the_corner_flag() {
    let text = fixture("sample.isel");
    let (a, _, ta) = run(&text, true, false);
    let (b, logs, tb) = run(&text, true, true);
    assert_eq!(a, b, "G2/G3 output must not change");
    assert_eq!(ta, tb);
    assert!(logs.iter().any(|l| l.contains("not applied in G2/G3 arc mode")));
}

// ── CLI ───────────────────────────────────────────────────────────────────
#[test]
fn cli_keeps_its_flags() {
    let exe = env!("CARGO_BIN_EXE_ISEL_to_Gcode");
    let dir = std::env::temp_dir();
    let inp = format!("{}/../tests/fixtures/sample.isel", env!("CARGO_MANIFEST_DIR"));
    let golden = |tag: &str| {
        std::fs::read(format!(
            "{}/../tests/fixtures/sample_expected.{}.ngc",
            env!("CARGO_MANIFEST_DIR"),
            tag
        ))
        .unwrap()
    };
    let out = |tag: &str, flags: &[&str]| {
        let path = dir.join(format!("isel_cli_{}_{}.ngc", tag, std::process::id()));
        let st = std::process::Command::new(exe)
            .args(["convert", &inp, path.to_str().unwrap()])
            .args(flags)
            .output()
            .unwrap();
        assert!(st.status.success(), "{}: {:?}", tag, st);
        let bytes = std::fs::read(&path).unwrap();
        std::fs::remove_file(&path).ok();
        bytes
    };

    let plain = out("plain", &[]);
    assert_eq!(plain, golden("g1"), "default CLI output unchanged");
    assert_eq!(out("arcs", &["--arcs"]), golden("arcs"), "--arcs unchanged");
    // arc mode ignores corner reduction: identical bytes
    assert_eq!(out("both", &["--arcs", "--corners"]), golden("arcs"));

    let corners = String::from_utf8(out("corners", &["--corners"])).unwrap();
    let plain = String::from_utf8(plain).unwrap();
    assert_only_feeds_differ(&plain, &corners);
    let reduced = |s: &str| {
        s.lines()
            .filter(|l| l.contains(" G1") && l.contains(" F"))
            .count()
    };
    assert_eq!(reduced(&plain), 0, "the original never puts F on a G1 line");
    assert!(reduced(&corners) > 0, "the sample's corners are reduced");
}
