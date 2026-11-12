#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
use std::process::exit;

/// Headless CLI used by tests/CI:
///   ISEL_to_Gcode convert <in.isel> <out.ngc> [--arcs]
/// Also prints "TIME <minutes>" for byte/time comparisons against the python reference.
fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() >= 4 && args[1] == "convert" {
        let arc = args.get(4).map(|a| a == "--arcs").unwrap_or(false);
        let mut log = |m: &str| println!("{}", m);
        let mut progress: Option<Box<dyn FnMut(f64, &str)>> =
            Some(Box::new(|p: f64, t: &str| eprintln!("PROGRESS {:.1} {}", p, t)));
        match isel_to_gcode::convert::convert_file(&args[2], &args[3], arc, &mut log, &mut progress) {
            Ok(mins) => {
                println!("TIME {}", mins);
                exit(0);
            }
            Err(e) => {
                eprintln!("ERROR: {}", e);
                exit(1);
            }
        }
    }
    println!("GUI not built yet");
}
