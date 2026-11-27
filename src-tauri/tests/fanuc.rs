use isel_to_gcode::convert::convert_file;
use isel_to_gcode::fanuc::{looks_like_fanuc, rewrite_to_isel};

const FIX: &str = include_str!("../../tests/fixtures/fanuc_demo.nc");

fn fx(name: &str) -> String {
    format!("{}/../tests/fixtures/{}", env!("CARGO_MANIFEST_DIR"), name)
}

// ── detection ───────────────────────────────────────────────────────────────
#[test] fn fanuc_fixture_detected() { assert!(looks_like_fanuc(FIX)); }

#[test] fn isel_fixtures_not_fanuc() {
    for f in ["sample.isel", "demo1.isel", "demo2.isel", "slot.isel"] {
        let text = std::fs::read_to_string(fx(f)).unwrap();
        assert!(!looks_like_fanuc(&text), "{} must stay ISEL", f);
    }
}

#[test] fn bare_g_state_words_are_not_fanuc() {
    // G17/G21/G90 alone carry no motion signal; an ISEL file quoting them in a
    // comment line is skipped by the detector.
    assert!(!looks_like_fanuc("IMF_PBL_V1.0\n; uses G17 and G21\nMOVEABS X1000\n"));
}

// ── rewrite: hand-computed expectation (each line maps one command) ─────────
#[test] fn rewrite_matches_hand_computed_isel() {
    let want = [
        "SPINDLE CW RPM1200",              // N0004 S+M03
        "FASTABS X25000 Y20000",           // N0006 G90 G00 X Y
        "FASTABS Z5000",                   // N0007 G43 Z.. H.. rides modal G0
        "VEL 4166.6750",                   // N0008 F250. -> 250 * VEL_RATIO
        "MOVEABS Z-1500",                  // N0008 G01 Z-1.5
        "MOVEABS X20000 Z-1750",           // N0009 modal G1, omitted Y stays
        "MOVEABS Y15000",                  // N0010 modal continuation
        "VEL 8333.3500",                   // N0011 feed-only block
        "MOVEABS X15000 Y10000",           // N0012 modal G1
        "CWABS X15000 Y20000 I15000 J15000", // N0013 G02 R5 chord=diameter: center = midpoint
        "MOVEABS X10000",                  // N0014 G01 restated
        "CCWABS X5000 Y15000 I10000 J15000", // N0015 G03, I/J incremental from (10,20)
        "FASTABS Z5000",                   // N0016 G00 Z5.
        "FASTABS X5000 Y5000",             // N0017 modal G0
    ].join("\n");
    // N0018/N0019 homing and everything after M30 are dropped by design.
    assert_eq!(rewrite_to_isel(FIX).unwrap(), want);
}

// ── parity: the Fanuc rewrite converts to the same program as the ISEL text
#[test] fn fanuc_input_converts_identical_to_isel_equivalent() {
    let isel = "IMF_PBL_V1.0 - TEST\nPLANE XY\nSPINDLE CW RPM1200\nFASTABS X25000 Y20000\n\
        FASTABS Z5000\nVEL 4166.6750\nMOVEABS Z-1500\nMOVEABS X20000 Z-1750\nMOVEABS Y15000\n\
        VEL 8333.3500\nMOVEABS X15000 Y10000\nCWABS X15000 Y20000 I15000 J15000\nMOVEABS X10000\n\
        CCWABS X5000 Y15000 I10000 J15000\nFASTABS Z5000\nFASTABS X5000 Y5000\n";
    let dir = std::env::temp_dir();
    let pid = std::process::id();
    let mut run = |name: &str, content: &str, arc: bool| -> (String, f64) {
        let inp = dir.join(format!("isel_fanuc_{}_{}.nc", name, pid));
        let out = dir.join(format!("isel_fanuc_{}_{}_o.ngc", name, pid));
        std::fs::write(&inp, content).unwrap();
        let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
        let t = convert_file(inp.to_str().unwrap(), out.to_str().unwrap(), arc, false,
                             &mut |_: &str| {}, &mut none).unwrap();
        let ngc = std::fs::read_to_string(&out).unwrap();
        std::fs::remove_file(&inp).ok(); std::fs::remove_file(&out).ok();
        (ngc, t)
    };
    for arc in [false, true] {
        let (a, ta) = run("f", FIX, arc);
        let (b, tb) = run("i", isel, arc);
        assert_eq!(a, b, "ngc differs (arc={})", arc);
        assert_eq!(ta, tb, "cycle time differs (arc={})", arc);
    }
}

// ── unsupported constructs fail loudly, naming the line ─────────────────────
fn must_err(src: &str, needle: &str) {
    let e = rewrite_to_isel(&format!("O1\n{}\nM30\n", src)).unwrap_err();
    assert!(e.contains(needle), "expected {:?} in {:?}", needle, e);
}

#[test] fn loud_errors() {
    must_err("G17 G20\nG00 X1.", "inch");
    must_err("G01 X1. F100.\nG81 X5. Y5. Z-3. R1.", "canned drilling cycle");
    must_err("G41 X5. Y5. D01", "cutter compensation");
    must_err("G91 G01 X5.", "incremental");
    must_err("G00 X1.\nS500 M04", "M04");
    must_err("G01 X1. F100.\nG04 P2.", "G04");
    must_err("G65 P1 L2", "not supported");
    must_err("X5. Y5.", "coordinates before any motion");
    must_err("G01 X1. F100.\nG02 X5. Y5. I2.", "partial I/J");
    // Full-circle blocks have no endpoint, but I/J must still be a complete pair.
    must_err("G01 X10. Y0. F100.\nG02 I-5.", "partial I/J");
    must_err("G01 X1. F100.\nG18 G02 X5. Y5. I2. J2.", "interpolation plane");
    must_err("G92 X10.", "G92");
    must_err("G95 G01 X1. F2.", "per revolution");
}

#[test] fn homing_line_with_g91_is_skipped_not_fatal() {
    let got = rewrite_to_isel("O1\nG00 X1. Y1.\nG91 G28 Z0.\nG28 X0. Y0.\nM30\n").unwrap();
    assert_eq!(got, "FASTABS X1000 Y1000");
}

// ── number shapes a post emits ──────────────────────────────────────────────
#[test] fn post_number_shapes() {
    // ".02" without leading digit, "25." with trailing dot, signed values
    let got = rewrite_to_isel("O1\nG00 X.02 Y25. Z-3.\nM30\n").unwrap();
    assert_eq!(got, "FASTABS X20 Y25000 Z-3000");
}

#[test] fn spindle_after_m_word_and_feed_roundtrip() {
    let got = rewrite_to_isel("O1\nM03 S900\nG00 X0.\nG01 X1. F333.33\nM30\n").unwrap();
    assert_eq!(got, "SPINDLE CW RPM900\nFASTABS X0\nVEL 5555.5111\nMOVEABS X1000");
    // the engine's VEL/16.6667 division must print the F the post wrote
    let dir = std::env::temp_dir();
    let inp = dir.join(format!("isel_fanuc_rt_{}.nc", std::process::id()));
    let out = dir.join(format!("isel_fanuc_rt_{}.ngc", std::process::id()));
    std::fs::write(&inp, "O1\nG00 X0.\nG01 X1. F333.33\nM30\n").unwrap();
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    convert_file(inp.to_str().unwrap(), out.to_str().unwrap(), false, false,
                 &mut |_: &str| {}, &mut none).unwrap();
    let ngc = std::fs::read_to_string(&out).unwrap();
    assert!(ngc.contains("F333"), "F must round-trip: {}", ngc);
    std::fs::remove_file(&inp).ok(); std::fs::remove_file(&out).ok();
}

// ── modal arcs: Fanuc semantics (R sign = minor/major; continuation repeats) ──
#[test] fn modal_arc_repeats_radius_and_vector() {
    let got = rewrite_to_isel(
        "O1\nG01 X10. Y0. F100.\nG02 X0. Y0. R5.\nX-10.\nG03 X10. Y0. I0. J5.\nM30\n").unwrap();
    let lines: Vec<&str> = got.lines().collect();
    assert_eq!(lines[2], "CWABS X0 Y0 I5000 J0"); // semicircle, center = midpoint (5,0)
    assert_eq!(lines[3], "CWABS X-10000 Y0 I-5000 J0"); // modal repeat of R5: center (-5,0)
    assert_eq!(lines[4], "CCWABS X10000 Y0 I-10000 J5000"); // incremental I/J from (-10,0): center (-10,5)
}

#[test] fn r_sign_selects_minor_or_major() {
    // G2 (0,0)->(0,6) with R+5: Fanuc minor arc, center to the right of travel
    let minor = rewrite_to_isel("O1\nG00 X0. Y0.\nG02 X0. Y6. R5.\nM30\n").unwrap();
    assert_eq!(minor.lines().last().unwrap(), "CWABS X0 Y6000 I4000 J3000");
    // R-5: Fanuc major arc, the other center
    let major = rewrite_to_isel("O1\nG00 X0. Y0.\nG02 X0. Y6. R-5.\nM30\n").unwrap();
    assert_eq!(major.lines().last().unwrap(), "CWABS X0 Y6000 I-4000 J3000");
    // mirrored for G3
    let ccw = rewrite_to_isel("O1\nG00 X0. Y0.\nG03 X0. Y6. R5.\nM30\n").unwrap();
    assert_eq!(ccw.lines().last().unwrap(), "CCWABS X0 Y6000 I-4000 J3000");
}

#[test] fn full_circle_bare_ij_line() {
    // Fanuc full circle: modal G2 with only I/J, no endpoint
    let got = rewrite_to_isel("O1\nG01 X10. Y0. F100.\nG02 I-5. J0.\nM30\n").unwrap();
    assert_eq!(got.lines().last().unwrap(), "CWABS X10000 Y0 I5000 J0");
}

#[test] fn g90_1_absolute_center() {
    let got = rewrite_to_isel("O1\nG90.1\nG01 X10. Y0. F100.\nG02 X0. Y0. I5. J0.\nM30\n").unwrap();
    assert_eq!(got.lines().last().unwrap(), "CWABS X0 Y0 I5000 J0"); // I/J is the center itself
}
