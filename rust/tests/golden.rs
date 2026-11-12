use isel_to_gcode::convert::convert_file;

fn fx(name: &str) -> String {
    format!("{}/../tests/fixtures/{}", env!("CARGO_MANIFEST_DIR"), name)
}

fn run_golden(arc: bool, tag: &str) {
    let inp = fx("sample.isel");
    let out = std::env::temp_dir().join(format!("isel_golden_{}_{}.ngc", tag, std::process::id()));
    let mut logs: Vec<String> = vec![];
    let mut none: Option<Box<dyn FnMut(f64, &str)>> = None;
    let t = convert_file(&inp, out.to_str().unwrap(), arc, &mut |m: &str| logs.push(m.to_string()), &mut none)
        .unwrap_or_else(|e| panic!("convert failed: {}", e));

    let got_ngc = std::fs::read(&out).unwrap();
    let want_ngc = std::fs::read(fx(&format!("sample_expected.{}.ngc", tag))).unwrap();
    assert_eq!(
        String::from_utf8_lossy(&got_ngc),
        String::from_utf8_lossy(&want_ngc),
        "ngc bytes differ ({} mode)",
        tag
    );

    let got_log = logs.join("\n") + "\n";
    let want_log = std::fs::read_to_string(fx(&format!("sample_expected.{}.log", tag))).unwrap();
    assert_eq!(got_log, want_log, "log differs ({} mode)", tag);

    let want_t: f64 = std::fs::read_to_string(fx(&format!("sample_expected.{}.time", tag)))
        .unwrap()
        .trim()
        .parse()
        .unwrap();
    // libm last-ULP differences are expected across CRTs; 1e-9 min = 0.06 s of program time
    assert!(
        (t - want_t).abs() < 1e-9,
        "time differs ({} mode): rust {} vs py {}",
        tag,
        t,
        want_t
    );
    std::fs::remove_file(&out).ok();
}

#[test] fn golden_g1() { run_golden(false, "g1"); }
#[test] fn golden_arcs() { run_golden(true, "arcs"); }
