# ISEL to G-code Converter

Windows GUI application that converts ISEL CAM files into standard G-code (NC) programs for CNC machines.

![version](https://img.shields.io/badge/version-2.5.1-blue)
![rust](https://img.shields.io/badge/rust-stable-orange)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)

## Features

- **Dark-mode GUI** with drag & drop or file browser
- **Coordinate conversion** — ISEL integer coordinates (×1000) to mm
- **Feed conversion** — ISEL `VEL` values to G-code `F` feed rates
- **Two output modes:**
  - `G1 linearised` — arcs expanded into short line segments
  - `G2/G3 arc commands` — true arc output with full-circle detection (R format, split into two semicircles for controller compatibility)
- **Spindle handling** — `S... M03` from `SPINDLE CW`, with `M05`/`M30` program end
- **Optional corner feed reduction** (off by default) — analyses the whole XY toolpath and lowers the feed around genuine sharp corners, then restores it. Curved sections are recognised as curves, not as a chain of corners, so each arc runs at one steady reduced feed and the straights keep full speed. Only the `F` word of existing movements changes: coordinates, movement count, line numbers and the machined path stay exactly the same
- **Toolpath preview** — top-down and isometric views with grid, pan, and zoom
- **Estimated cycle time** shown after conversion
- Single-file `.exe` — no installation required; renders with the Edge WebView2
  runtime that ships with Windows 11 (and arrives with Edge updates on Windows 10)

## Usage

1. Download the latest release (`ISEL_to_Gcode.exe`) from the [Releases](https://github.com/47Hunter47/ISEL-to-Gcode/releases) page.
2. Run the executable.
3. Drag & drop your ISEL file (`.isel`) or ISELCAM export (`.ncp`) onto the window (or click **Browse**, or drop the file onto the `ISEL_to_Gcode.exe` icon).
4. Optionally enable **Use G2/G3 arc commands** if your controller supports arcs.
5. Optionally enable **Corner Feed Reduction** to slow the machine down around sharp corners. Leave it off if your parts are mostly short zig-zag passes — every turn is a corner there, so the program takes noticeably longer.
6. Click **Convert** and choose a save location (`.ngc`).
7. Open **Preview** to inspect the toolpath before running the program.

> ⚠ The estimated program time is an approximation and may differ from actual machine time depending on machine parameters.

## Output format

Programs start with a standard header and end with spindle stop:

```
N00001 G21
N00002 G17
N00003 G40
N00004 G49
N00005 G90
N00006 G94
N00007 G0 Z...        ← first Z safe retract
N00008 S1200 M03      ← spindle starts after the tool is clear
...
M05
M30
```

The first spindle command is emitted **after** the first Z safe retract, so the spindle starts turning only once the tool is clear of the workpiece. Subsequent spindle speed changes are emitted immediately.

## Building from source

Requires Rust (the `src-tauri/rust-toolchain.toml` file pins the toolchain), on
Windows the GNU binutils that toolchain needs for linking (MSYS2/WinLibs or
`winget install BrechtSanders.WinLibs.POSIX.UCRT`), and Node.js 22+ for the UI.

```bat
npm install
npm run build
cd src-tauri
cargo build --release --features custom-protocol
```

`cargo build --release --features custom-protocol` embeds the built UI into the exe,
so the result stays a single file in `src-tauri/target/release/`. (Built without that
feature the exe looks for a dev server at `localhost:1420` instead and shows an empty
"page not found" window.) The same executable also runs
without the window for scripting: `ISEL_to_Gcode.exe convert in.isel out.ngc [--arcs] [--corners]`
(flags come after the two paths; `--corners` enables corner feed reduction).

`cargo test` compares the output byte-for-byte with reference conversions,
including two real customer programs in `tests/fixtures`, and checks that corner
feed reduction never changes a coordinate, a movement or a line number — only
the `F` values.

Pushing to `main` publishes a release automatically: the `Release` workflow reads the
version from `src-tauri/Cargo.toml`, runs the tests, builds and attaches the exe. If that
version is already released it does nothing, so a push that does not bump the version
publishes nothing.
