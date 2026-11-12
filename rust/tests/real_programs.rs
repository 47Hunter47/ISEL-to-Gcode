// Parity against Mert's two real CAM programs (demo1.ncp / demo2.ncp, sent
// 2026-11-12). The g1 goldens are the .ngc files he converted with the shipped
// python app, so this is byte-parity against published output, not a re-run.
// demo2 has 100 arc commands and linearises to 314k lines; demo1 has none, so
// both modes produce the same bytes.
use isel_to_gcode::convert::convert_file;

fn fx(name: &str) -> String {
    format!("{}/../tests/fixtures/{}", env!("CARGO_MANIFEST_DIR"), name)
}

fn run(name: &str, arc: bool, golden: &str, want_time: f64) {
    let out = std::env::temp_dir().join(format!(
        "isel_real_{}_{}_{}.ngc",
        name,
        if arc { "arcs" } else { "g1" },
        std::process::id()
    ));
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    let mut silent = |_: &str| {};
    let t = convert_file(&fx(&format!("{}.isel", name)), out.to_str().unwrap(), arc, &mut silent, &mut none)
        .unwrap_or_else(|e| panic!("{} convert failed: {}", name, e));

    let got = std::fs::read(&out).unwrap();
    let want = std::fs::read(fx(golden)).unwrap();
    assert_eq!(got.len(), want.len(), "{} ngc size differs", name);
    assert_eq!(
        String::from_utf8_lossy(&got),
        String::from_utf8_lossy(&want),
        "{} ngc bytes differ",
        name
    );
    // libm last-ULP drift; 1e-9 min = 0.06 s of program time
    assert!((t - want_time).abs() < 1e-9, "{} time: rust {} want {}", name, t, want_time);
    std::fs::remove_file(&out).ok();
}

#[test]
fn demo1_g1_matches_published() {
    run("demo1", false, "demo1_expected.g1.ngc", 13.363929830118765);
}

#[test]
fn demo1_arcs_matches_published() {
    // no CWABS/CCWABS in this program: arc mode must emit the same bytes
    run("demo1", true, "demo1_expected.g1.ngc", 13.363929830118765);
}

#[test]
fn demo2_g1_matches_published() {
    run("demo2", false, "demo2_expected.g1.ngc", 10.496234448475477);
}

#[test]
fn demo2_arcs_matches_python_reference() {
    run("demo2", true, "demo2_expected.arcs.ngc", 10.49623488494);
}
