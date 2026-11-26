use isel_to_gcode::gcode::{arc_to_polyline, parse_gcode, GcodeParsed};

/// Write G-code text to a temp file and parse it. `tag` keeps parallel tests
/// from colliding on the same file name.
fn parse_str(content: &str, tag: &str) -> GcodeParsed {
    let path = std::env::temp_dir().join(format!(
        "isel_gcode_test_{}_{}.ngc", tag, std::process::id()));
    std::fs::write(&path, content).unwrap();
    let f = parse_gcode(&path.to_string_lossy()).unwrap();
    let _ = std::fs::remove_file(&path);
    f
}

fn near(a: (f64, f64), b: (f64, f64), tol: f64) -> bool {
    (a.0 - b.0).abs() < tol && (a.1 - b.1).abs() < tol
}

// Values below were captured from the python reference `_arc_to_polyline` on
// 2026-11-12 (see ledger); they pin the port, including the python center-side
// logic for positive-R G2 that yields a 270° CW sweep.
#[test] fn arc_positive_r_reference_points() {
    let pts = arc_to_polyline(5., 0., 0., 5., 5.0, true);
    assert_eq!(pts.len(), 37);
    assert!(near(pts[0], (5.0, 0.0), 1e-12));
    assert!(near(pts[18], (-3.535533905932737, -3.5355339059327373), 1e-9));
    assert!(near(pts[36], (0.0, 5.0), 1e-6));
}

#[test] fn arc_negative_r_flips_center() {
    let pts = arc_to_polyline(5., 0., 0., 5., -5.0, true);
    assert_eq!(pts.len(), 37);
    assert!(near(pts[18], (1.4644660940672627, 1.4644660940672622), 1e-9));
}

#[test] fn arc_degenerate_zero_distance() {
    let pts = arc_to_polyline(1., 1., 1., 1., 5.0, false);
    assert_eq!(pts.len(), 2);
    assert!(near(pts[0], (1., 1.), 1e-12) && near(pts[1], (1., 1.), 1e-12));
}

#[test] fn arc_degenerate_zero_radius() {
    let pts = arc_to_polyline(5., 0., 0., 5., 0.0, true);
    assert_eq!(pts.len(), 2);
    assert!(near(pts[0], (5., 0.), 1e-12) && near(pts[1], (0., 5.), 1e-12));
}

#[test] fn parse_g1_golden_counts_and_bounds() {
    let f = parse_gcode(concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/fixtures/sample_expected.g1.ngc")).unwrap();
    // python open_preview counts POINTS (move endpoints). 806 G1 moves => 1612
    // cut pts. The G17/G21 header lines are NOT motion: before exact G-word
    // parsing they matched startswith("G1")/("G2") and added 4 phantom cut pts
    // (asserted 1616). Those assertions encoded the parser bug; corrected
    // 2026-11-26 with the parser fix, golden output file unchanged.
    assert_eq!(f.counts.rapid, 6);
    assert_eq!(f.counts.cut, 1612);
    assert_eq!(f.counts.arc, 0);
    assert!((f.bounds.min_x - 0.0).abs() < 1e-9);
    assert!((f.bounds.max_x - 35.0).abs() < 1e-9);
    assert!((f.bounds.min_y - 0.0).abs() < 1e-9);
    assert!((f.bounds.max_y - 34.511).abs() < 1e-9);
    assert!(f.groups.len() >= 2);
}

// Stage-1 contract: motion is decided by an exact G-word, never by a prefix.

#[test] fn g17_g21_g20_setup_words_are_ignored() {
    let f = parse_str(
        "N1 G17\nN2 G18\nN3 G19\nN4 G20\nN5 G21\nN6 G40\nN7 G49\nN8 G54\n\
         N9 G90\nN10 G91\nN11 G94\nN12 G91.1\nN13 G4\nN14 G4.0\n", "setup_words");
    assert!(f.groups.is_empty(), "setup G-codes must not draw: {:?}", f.counts);
    assert_eq!((f.counts.rapid, f.counts.cut, f.counts.arc), (0, 0, 0));
    // a file of only setup commands is an empty preview (default bounds)
    assert_eq!(f.bounds.min_x, 0.0);
    assert_eq!(f.bounds.max_x, 0.0);
}

#[test] fn g0_and_g00_are_rapid() {
    let f = parse_str("N1 G0 X1 Y1\nN2 G00 X2 Y2 Z3\n", "g0_g00");
    assert_eq!((f.counts.rapid, f.counts.cut, f.counts.arc), (4, 0, 0));
}

#[test] fn g1_and_g01_are_cut_not_rapid() {
    // G01 startswith("G0") used to classify it as rapid.
    let f = parse_str("N1 G1 X1 Y1\nN2 G01 X2 Y2\nN3 G01X7 Y7\n", "g1_g01");
    assert_eq!((f.counts.rapid, f.counts.cut, f.counts.arc), (0, 6, 0));
    assert!((f.bounds.max_x - 7.0).abs() < 1e-9);
}

#[test] fn g2_g02_g3_g03_are_arcs() {
    let f = parse_str(
        "N1 G2 X5 Y0 R5\nN2 G02 X0 Y5 R5\nN3 G3 X5 Y5 R5\nN4 G03 X0 Y0 R5\n", "arcs");
    // each R-arc expands to 37 polyline points
    assert_eq!(f.counts.arc, 148);
    assert_eq!((f.counts.rapid, f.counts.cut), (0, 0));
}

#[test] fn n_numbered_motion_lines_parse() {
    let f = parse_str("N00001 G1 X10.000 Y20.000\nN00002 G0 X0.5 Y0.25\n", "nnum");
    assert_eq!((f.counts.rapid, f.counts.cut), (2, 2));
    assert!((f.bounds.max_x - 10.0).abs() < 1e-9);
    assert!((f.bounds.max_y - 20.0).abs() < 1e-9);
}

#[test] fn coordinates_in_comments_are_not_moves() {
    let f = parse_str(concat!(
        "N1 G1 X10 Y10 (travel to X99 Y99)\n",
        "N2 G1 (plunge at X77 Y77) X20 Y20 ; past X999 Y999\n",
        "; whole line G1 X88 Y88\n",
        "N3 (commented G1 X55 Y55)\n",
        "N4 G1 X30 (mid) Y30\n",
    ), "comments");
    // three moves, endpoints 10,20,30 — never 55/77/88/99/999
    assert_eq!((f.counts.rapid, f.counts.cut, f.counts.arc), (0, 6, 0));
    assert!((f.bounds.min_x - 0.0).abs() < 1e-9); // path starts at 0,0
    assert!((f.bounds.max_x - 30.0).abs() < 1e-9);
    assert!((f.bounds.max_y - 30.0).abs() < 1e-9);
}

#[test] fn non_motion_words_between_moves_leave_bounds_and_counts_clean() {
    let with_setup = parse_str(concat!(
        "N1 G1 X10 Y10\n",
        "N2 G90\nN3 G21\nN4 G20\nN5 G17\nN6 G40\nN7 G94\n",
        "N8 G1 X20 Y20\n",
    ), "setup_between");
    let plain = parse_str(
        "N1 G1 X10 Y10\nN2 G1 X20 Y20\n", "setup_between_plain");
    assert_eq!(with_setup.counts.cut, plain.counts.cut);
    assert_eq!(with_setup.counts.cut, 4);
    assert_eq!(with_setup.bounds.max_x, plain.bounds.max_x);
    assert_eq!(with_setup.bounds.min_x, 0.0);
    assert_eq!(with_setup.bounds.max_y, 20.0);
}

#[test] fn converter_header_block_adds_no_geometry() {
    // The header the converter itself emits, straight from the golden fixture.
    let f = parse_str(concat!(
        "N00001 G21\r\nN00002 G17\r\nN00003 G40\r\nN00004 G49\r\n",
        "N00005 G90\r\nN00006 G94\r\nN00007 G0 X0.000 Y0.000 Z4.000\r\n",
        "N00008 F1000\r\nN00009 G1 X10.000 Y10.000 Z-1.000\r\n",
    ), "header_block");
    assert_eq!((f.counts.rapid, f.counts.cut, f.counts.arc), (2, 2, 0));
}

