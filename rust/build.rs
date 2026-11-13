// Explorer, the taskbar and shortcuts take an exe's icon from an embedded Win32
// resource; `with_icon` in app.rs only sets the icon of the live window. windres
// already ships with the GNU toolchain this project builds with (w64devkit
// locally, C:\mingw64 on CI), so this needs no extra crate.
use std::path::PathBuf;
use std::process::Command;

fn main() {
    println!("cargo:rerun-if-changed=assets/icon.ico");

    if std::env::var("CARGO_CFG_TARGET_ENV").as_deref() != Ok("gnu") {
        println!(
            "cargo:warning=exe icon skipped: build.rs only knows how to do this for windows-gnu"
        );
        return;
    }

    let out = PathBuf::from(std::env::var_os("OUT_DIR").expect("OUT_DIR"));
    let ico = PathBuf::from(std::env::var_os("CARGO_MANIFEST_DIR").expect("CARGO_MANIFEST_DIR"))
        .join("assets/icon.ico");
    // absolute, forward-slash path: a relative one would resolve against windres' cwd
    let rc = out.join("icon.rc");
    std::fs::write(
        &rc,
        format!(
            "1 ICON \"{}\"\n",
            ico.display().to_string().replace('\\', "/")
        ),
    )
    .expect("write icon.rc");

    let windres = ["windres", "x86_64-w64-mingw32-windres"]
        .into_iter()
        .find(|tool| {
            Command::new(tool)
                .arg("--version")
                .output()
                .is_ok_and(|o| o.status.success())
        })
        .unwrap_or_else(|| {
            panic!("windres not found on PATH: put the GNU binutils (w64devkit bin) on PATH so the exe gets its icon")
        });

    let obj = out.join("icon_res.o");
    let status = Command::new(windres)
        .args(["--output-format=coff", "--input"])
        .arg(&rc)
        .arg("--output")
        .arg(&obj)
        .status()
        .unwrap_or_else(|e| panic!("could not run {windres}: {e}"));
    assert!(status.success(), "{windres} failed on {}", rc.display());

    println!("cargo:rustc-link-arg-bins={}", obj.display());
}
