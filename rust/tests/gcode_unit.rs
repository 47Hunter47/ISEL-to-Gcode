use isel_to_gcode::gcode::{arc_to_polyline, parse_gcode};

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
    // python open_preview counts POINTS (move endpoints). 806 G1 moves + the
    // G17/G21 header lines that match startswith("G1")/("G2") => 808 x 2 cut pts.
    assert_eq!(f.counts.rapid, 6);
    assert_eq!(f.counts.cut, 1616);
    assert_eq!(f.counts.arc, 0);
    assert!((f.bounds.min_x - 0.0).abs() < 1e-9);
    assert!((f.bounds.max_x - 35.0).abs() < 1e-9);
    assert!((f.bounds.min_y - 0.0).abs() < 1e-9);
    assert!((f.bounds.max_y - 34.511).abs() < 1e-9);
    assert!(f.groups.len() >= 2);
}
