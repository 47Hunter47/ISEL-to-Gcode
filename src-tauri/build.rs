use std::path::Path;

fn main() {
    // The static WebView2 loader (vendored webview2-com-sys) links MSVC-built
    // objects; on windows-gnu they need the small CRT shim in shim/ because no
    // mingw-w64 library provides those symbols. The compiled objects are handed
    // to the linker for every binary and test this crate produces, so the exe
    // stays a single self-contained file.
    let target = std::env::var("TARGET").unwrap_or_default();
    if target == "x86_64-pc-windows-gnu" {
        let out = std::env::var("OUT_DIR").expect("OUT_DIR");
        let cc = std::env::var("CC").unwrap_or_else(|_| "gcc".to_string());
        for src in ["shim/msvc_crt_shim.c", "shim/msvc_crt_shim.s"] {
            let obj = Path::new(&out).join(format!(
                "{}.o",
                Path::new(src).file_name().unwrap().to_string_lossy()
            ));
            let status = std::process::Command::new(&cc)
                .args(["-c", "-O2", src, "-o"])
                .arg(&obj)
                .status()
                .expect("gcc not found: a mingw-w64 toolchain must be on PATH (see AGENTS.md)");
            assert!(status.success(), "failed to compile {src}");
            println!("cargo:rustc-link-arg={}", obj.display());
            println!("cargo:rerun-if-changed={src}");
        }

        // RT_MANIFEST (1/24): the v5 comctl32 in System32 has no
        // TaskDialogIndirect, so without the Common-Controls v6 dependency the
        // loader refuses to map the exe (STATUS_ENTRYPOINT_NOT_FOUND). GNU
        // links no manifest of its own and tauri-build does not emit one, so
        // this is the exe's only manifest resource.
        let mf = Path::new(&out).join("app.manifest");
        std::fs::write(&mf, include_str!("shim/app.manifest")).expect("write manifest");
        let rc = Path::new(&out).join("app.manifest.rc");
        std::fs::write(
            &rc,
            format!("1 24 \"{}\"\n", mf.display().to_string().replace('\\', "/")),
        )
        .expect("write manifest rc");
        let res = Path::new(&out).join("app.manifest.res");
        let status = std::process::Command::new("windres")
            .args(["--input"])
            .arg(&rc)
            .args(["--output"])
            .arg(&res)
            .args(["--output-format", "coff"])
            .status()
            .expect("windres not found: a mingw-w64 toolchain must be on PATH (see AGENTS.md)");
        assert!(status.success(), "failed to compile app.manifest.rc");
        println!("cargo:rustc-link-arg={}", res.display());
        println!("cargo:rerun-if-changed=shim/app.manifest");
    }

    tauri_build::build()
}
