# AGENTS.md — ISEL-to-Gcode

Converts ISEL/CAM programs into the G-code that ISIN's heater-winding machines read.
Status, open work, bug history and test results are NOT kept here — they live in the
Obsidian note `Projeler/ISEL_to_Gcode.md`. This file is the rule book only.

The repository is public: documentation, comments, commit messages and release notes are
written in English.

## What lives where
- `rust/` is the whole program: library + one executable, `ISEL_to_Gcode`.
  `src/convert.rs` conversion, `src/corner.rs` optional corner feed reduction (read-only
  geometry analysis), `src/gcode.rs` G-code parser for the preview, `src/preview.rs` the
  drawing scene (testable without a window), `src/app.rs` the eframe UI and CLI entry,
  `src/main.rs` a thin entry point.
- `tests/fixtures/` holds the byte goldens (a synthetic fixture + two real customer
  programs). They are the contract; `rust/tests/` runs them.
- The Python program that produced the goldens is not in the tree. "python line N"
  references in comments belong to it; recover with
  `git show 9b1fa1d^:isel_to_gcode_gui.py`.

## Commands
- Local build/test: `export PATH="$HOME/tools/w64devkit/bin:$PATH"; cd rust && cargo test && cargo build --release`
- Windows GNU only. Keep `-C link-self-contained=yes` in `rust/.cargo/config.toml`.
- CLI: `ISEL_to_Gcode.exe convert <in.isel> <out.ngc> [--arcs] [--corners]` — flags come
  after the two paths.
- Release: `version` in `rust/Cargo.toml` is the single source of truth and is bumped by
  hand; commit, `git tag -a vX.Y.Z`, push the tag, `gh release create vX.Y.Z`. Only a
  *published* release triggers the CI build and attaches `ISEL_to_Gcode.exe`.

## Do
- Keep the output byte-for-byte identical to `tests/fixtures/*expected*`.
- Preserve the reference program's quirks (for example printing negative zero as
  `-0.000`); matching the reference beats being pretty.
- Run the conversion on a worker thread, never on the UI thread; the preview uses the
  worker's result.
- Check the return value of `eframe::run_native`; an ignored error means a silent exit
  with no window.
- Keep the startup log (`%TEMP%\ISEL_to_Gcode_startup.log`), the panic hook and the error
  box; on a machine where the window never appears, they are the only evidence.
- Render order is Glow (OpenGL) first, wgpu as fallback; `ISEL_RENDERER=glow|wgpu`
  switches it temporarily.
- Define colours in one place: `src/theme.rs`. Do not paint widgets by hand;
  `theme::apply()` writes both the Dark and the Light bucket and pins the theme
  preference — under Windows light mode egui's white defaults (progress track, checkbox)
  otherwise leak back in.
- UI evidence on a headless machine: with `ISEL_SHOT=path.bmp` the app writes one frame
  and exits. `app::tests::shot_bmp_is_bottom_up_bgr` protects the writer's channel/row
  order.
- The exe file icon comes from the Win32 resource `rust/build.rs` embeds with `windres`
  (`assets/icon.ico`); `with_icon` only paints the live window. `windres` must be on PATH —
  without it build.rs deliberately fails, so an icon-less exe is never published silently.
- The preview window is a separate OS window (`show_viewport_immediate`); moving it back
  to `egui::Window` squeezes it inside the main window. Headless evidence:
  `ISEL_PREVIEW=<out.ngc>` opens the preview at startup and a second top-level window
  shows up via `EnumWindows`.
- Never call a blocking shell dialog (`rfd::FileDialog`) from the UI pass; it runs
  the modal IFileDialog on the calling thread and freezes the window for as long as
  the shell needs (seconds on a network share). Spawn it, drain the answer on the UI
  thread, and show that something is in flight.
- Keep the `.ngc`/`.isel`/`.ncp` CRLF pins in `.gitattributes`.
- After a code change refresh both graphs: codegraph `sync` and `graphify update .`.

## Don't
- In corner feed reduction do not split segments, add or remove points, round coordinates
  or smooth the path; only the F value of existing G1 lines is scaled, and with the option
  off the output stays byte-identical.
- Do not bring back a Python/.NET/Electron layer; delivery is a single exe with no
  installation.
- Never regenerate a golden from current Rust output just to make a test pass.
- Do not switch the local or CI build to MSVC `link.exe`; it crashes on this machine.
- Do not delete the synthetic fixture goldens (`sample_expected.*`); they cover paths the
  real programs never take (default feed when no VEL is present, R arcs, zero-length
  protection).
- `graphify-out/`, `graph_update.log`, `.codegraph/` and `rust/target/` are not committed.
- Do not run `cargo fmt`: the code is hand-formatted, rustfmt rewrites the whole tree and
  buries the real change.
- Do not use subagents in this project.

## Code graph (graphify) — automatic use

- **Before opening project files** take a quick look at the architecture:
  ```bash
  read_file graphify-out/GRAPH_REPORT.md   # God Nodes + Communities (first 30 lines are enough)
  ```
- **Before a large refactor or multi-file change** query the graph:
  ```bash
  python -m graphify affected "module_or_function" --depth 2 --graph graphify-out/graph.json
  python -m graphify query "question" --graph graphify-out/graph.json
  ```
- Not needed for small single-file changes — `search_files` is enough.
