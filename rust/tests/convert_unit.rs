use isel_to_gcode::convert::{convert_file, fit_circle, parse_coord, try_arc_fit};

// ── parse_coord ─────────────────────────────────────────────────────────────
#[test] fn fast_int() { assert_eq!(parse_coord("MOVEABS X328495 Z-103460", &['X','Y','Z'])[&'X'], 328.495); }
#[test] fn bare_axis_ignored() { assert!(parse_coord("MOVEABS X", &['X','Y','Z']).is_empty()); }
#[test] fn first_occurrence_wins() { assert_eq!(parse_coord("MOVEABS X1000 X2000", &['X','Y','Z'])[&'X'], 1.0); }
#[test] fn guarded_float() { assert_eq!(parse_coord("MOVEABS X1.5", &['X','Y','Z'])[&'X'], 0.0015); }
#[test] fn fallback_dot_start() { let c = parse_coord("MOVEABS X.5 Y1000", &['X','Y','Z']);
    assert!(!c.contains_key(&'X')); assert_eq!(c[&'Y'], 1.0); }
#[test] fn fallback_no_space() { let c = parse_coord("FASTABS X123Y45", &['X','Y','Z']);
    assert_eq!(c[&'X'], 0.123); assert_eq!(c[&'Y'], 0.045); }
#[test] fn ij_axes() { let c = parse_coord("CWABS X1000 I500 J-250", &['I','J']);
    assert_eq!(c[&'I'], 0.5); assert_eq!(c[&'J'], -0.25); }

// ── circle fit ──────────────────────────────────────────────────────────────
#[test] fn fit_circle_lt3() { assert!(fit_circle(&[(0., 0.), (1., 0.)]).is_none()); }
#[test] fn fit_circle_collinear() { assert!(fit_circle(&[(0., 0.), (1., 1.), (2., 2.)]).is_none()); }
#[test] fn fit_circle_unit() {
    let (cx, cy, r) = fit_circle(&[(1., 0.), (0., 1.), (-1., 0.)]).unwrap();
    assert!((cx * cx + cy * cy + (r - 1.).powi(2)).abs() < 1e-3);
}
#[test] fn arc_fit_half_rejected() {
    // ~160° sweep: fewer than 355° => not a full circle
    let mut buf: Vec<(f64, f64, f64)> = vec![];
    for i in 1..=8 {
        let a = std::f64::consts::PI * i as f64 / 9.0;
        buf.push((a.cos() * 10., a.sin() * 10., 0.));
    }
    assert!(try_arc_fit(&buf, (10., 0., 0.)).is_none());
}

// ── flow tests through convert_file ────────────────────────────────────────
fn run(text: &str, arc: bool) -> (String, Vec<String>, f64) {
    static N: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    let id = N.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let dir = std::env::temp_dir();
    let inp = dir.join(format!("isel_t_{}_{}.isel", std::process::id(), id));
    let outp = dir.join(format!("isel_t_{}_{}.ngc", std::process::id(), id));
    std::fs::write(&inp, text).unwrap();
    let mut logs: Vec<String> = vec![];
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    let r = convert_file(inp.to_str().unwrap(), outp.to_str().unwrap(), arc,
                         &mut |m: &str| logs.push(m.to_string()), &mut none);
    // ngc comes back CRLF (python writes text-mode on Windows); tests compare logical lines
    let ngc = std::fs::read_to_string(&outp).unwrap_or_default();
    std::fs::remove_file(&inp).ok(); std::fs::remove_file(&outp).ok();
    (ngc, logs, r.unwrap_or(-1.0))
}

#[test] fn empty_input_error() {
    let dir = std::env::temp_dir();
    let inp = dir.join("isel_t_empty.isel"); let outp = dir.join("isel_t_eo.ngc");
    std::fs::write(&inp, "; only a comment\n\n").unwrap();
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    let r = convert_file(inp.to_str().unwrap(), outp.to_str().unwrap(), false,
                         &mut |_: &str| {}, &mut none);
    assert_eq!(r.unwrap_err(), "Input file contains no processable commands.");
}

#[test] fn moveabs_vel_flow() {
    let (ngc, logs, t) = run("MOVEABS X10000 Y10000 Z-1000\nVEL 16667\nMOVEABS X20000 Y10000 Z-1000\n", false);
    let l: Vec<&str> = ngc.lines().collect();
    assert_eq!(l[6], "N00007 F1000"); // default feed before any VEL
    assert!(logs.iter().any(|s| s.starts_with("Warning: No VEL found")));
    assert_eq!(l[7], "N00008 G1 X10.000 Y10.000 Z-1.000");
    assert_eq!(l[8], "N00009 F1000"); // 16667 / 16.6667 formats as 1000
    assert!(logs.iter().any(|s| s == "Feed: F1000"));
    assert_eq!(l[9], "N00010 G1 X20.000 Y10.000 Z-1.000");
    assert!(l[10].ends_with("M05")); assert!(l[11].ends_with("M30"));
    assert!(t > 0.02 && t < 0.03); // 14.14mm at 1000 + 10mm at 1000 mm/min
}

#[test] fn spindle_defer_then_immediate() {
    // python semantics: first SPINDLE defers; while pending exists the next
    // SPINDLE emits immediately; the (older) pending flushes after first Z.
    let (ngc, logs, _) = run("SPINDLE CW RPM1200\nSPINDLE CW RPM1800\nFASTABS X10000 Y10000 Z4000\n", false);
    let iz = ngc.lines().position(|s| s.contains("Z4.000")).unwrap();
    let i1800 = ngc.lines().position(|s| s.contains("S1800 M03")).unwrap();
    let i1200 = ngc.lines().position(|s| s.contains("S1200 M03")).unwrap();
    assert!(i1800 < iz, "second spindle must emit immediately");
    assert!(i1200 > iz, "deferred first spindle flushes after the Z line");
    assert!(logs.iter().any(|s| s.contains("deferred until first Z retract")));
    assert!(logs.iter().any(|s| s == "Spindle: 1800 RPM"));
}

#[test] fn bare_fastabs_emits_g0() {
    let (ngc, _, _) = run("FASTABS\n", false);
    assert!(ngc.lines().any(|s| s.ends_with(" G0")));
}

#[test] fn spindle_no_rpm_warning() {
    let (_, logs, _) = run("SPINDLE CW\n", false);
    assert!(logs.iter().any(|s| s.starts_with("Warning: Could not parse RPM from: SPINDLE CW")));
}

#[test] fn linearize_arc_396_steps() {
    // (20,20)->(30,30) center (26,24): r=sqrt(52)=7.2111, span=2.7468rad
    // arc_len 19.807mm / 0.05mm = 396 steps
    let (ngc, _, _) = run("FASTABS X20000 Y20000 Z-1000\nCWABS X30000 Y30000 I26000 J24000\n", false);
    let arc_lines: Vec<&str> = ngc.lines().filter(|s| s.contains(" G1 X")).collect();
    assert_eq!(arc_lines.len(), 396);
    let last = arc_lines.last().unwrap();
    assert!(last.contains("X30.000") && last.contains("Y30.000") && last.contains("Z-1.000"));
}

#[test] fn g2_line_format() {
    let (ngc, _, _) = run("FASTABS X20000 Y20000 Z-1000\nCWABS X30000 Y30000 I26000 J24000\n", true);
    assert!(ngc.contains("G2 X30.000 Y30.000 R7.211")); // span<pi => positive R
}

#[test] fn zero_radius_arc_skipped() {
    let (_, logs, _) = run("FASTABS X20000 Y20000 Z-1000\nCWABS X20000 Y20000 I20000 J20000\n", true);
    assert!(logs.iter().any(|s| s.contains("zero-radius arc skipped")));
}

#[test] fn duplicate_moveabs_mode_quirk() {
    // G1 mode: direct path re-emits axes that are merely PRESENT;
    // arcs mode: buffered path skips axes that are UNCHANGED.
    let dup = "G1 X20.000 Y10.000 Z-1.000";
    let input = "VEL 16667\nMOVEABS X20000 Y10000 Z-1000\nMOVEABS X20000 Y10000 Z-1000\n";
    let (ngc_g1, _, _) = run(input, false);
    let (ngc_arc, _, _) = run(input, true);
    assert_eq!(ngc_g1.matches(dup).count(), 2);
    assert_eq!(ngc_arc.matches(dup).count(), 1);
}

// Full-circle fit must run: a buffered run of MOVEABS forming a closed circle
// becomes two G3 semicircles (R = chord/2 + 0.001). Regression pin for the
// python bug where the MOVEABS branch flushed its own buffer (fit never ran).
#[test] fn full_circle_fit_two_semicircles() {
    let fixture = include_str!("../../tests/fixtures/sample.isel");
    let (ngc, logs, _) = run(fixture, true);
    assert!(logs.iter().any(|s| s.starts_with("  Arc fit: full circle")));
    assert!(logs.iter().any(|s| s.contains("1 circle(s) converted to G2/G3")));
    assert!(ngc.contains("G3 X15.000 Y25.000 R10.001 F500"), "first semicircle");
    assert!(ngc.contains("G3 X35.000 Y25.000 R10.001 F500"), "second semicircle");
    assert!(!ngc.contains("G1 X33.090"), "polyline points consumed by the fit");
    assert!(ngc.contains("G3 X20.000 Y20.000 R7.211 F500")); // CWABS/CCWABS arcs still work
}
