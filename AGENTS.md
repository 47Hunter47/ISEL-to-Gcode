# AGENTS.md — ISEL-to-Gcode

Converts ISEL/CAM programs into the G-code that cnc milling machines read.
Status, open work, bug history and test results are NOT kept here — they live in the
This file is the rule book only.

The repository is public: documentation, comments, commit messages and release notes are
written in English.

## What lives where
- `src-tauri/` is the whole Rust program: library + one executable, `ISEL_to_Gcode`.
  `src/convert.rs` conversion, `src/corner.rs` optional corner feed reduction (read-only
  geometry analysis), `src/gcode.rs` G-code parser for the preview, `src/preview.rs` the
  preview camera math (egui-free, testable without a window), `src/commands.rs` the Tauri
  commands the UI calls, `src/app.rs` tauri window creation + startup diagnostics,
  `src/main.rs` a thin entry point (CLI first, GUI after).
- The web UI lives in `src/` (Vite + React + TS + Tailwind v4 + shadcn/ui);
  `npm run build` produces `dist/`, which the tauri build script embeds into the exe.
  Delivery stays a single file: never bundle an installer (`tauri build` bundlers off).
- `tests/fixtures/` holds the byte goldens (a synthetic fixture + two real customer
  programs). They are the contract; `src-tauri/tests/` runs them.
- The Python program that produced the goldens is not in the tree. "python line N"
  references in comments belong to it; recover with
  `git show 9b1fa1d^:isel_to_gcode_gui.py`.

## Commands
- Local build/test: `export PATH="$HOME/tools/w64devkit/bin:$PATH`;
  `npm install && npm run build` (dist/ must exist before the cargo steps), then
  `cd src-tauri && cargo test && cargo build --release`.
- Windows GNU only. Keep `-C link-self-contained=yes` in `src-tauri/.cargo/config.toml`.
- UI dev loop: `npm run tauri dev` (vite on port 1420 + hot reload).
- CLI: `ISEL_to_Gcode.exe convert <in.isel> <out.ngc> [--arcs] [--corners]` — flags come
  after the two paths.
- Finishing a change is `./tools/ship.sh "<conventional commit message>"`: it commits the
  work, bumps the version implied by the commits since the last bump (`feat` = minor,
  `fix` = patch, otherwise none), updates `Cargo.toml`/`Cargo.lock`/`tauri.conf.json`/the
  README badge in the same commit and pushes. A plain `git push` does not bump, and a
  pre-push hook cannot do it: git resolves the refs before the hook runs, so a commit
  made inside the hook is never pushed.
- Release: pushing to `main` publishes it. The `Release` workflow derives the version from
  `version` in `src-tauri/Cargo.toml` (never typed, so it cannot drift), builds the web
  front end, builds and attaches the exe at the pushed commit. If that version is already
  released the job stops before building, so a push without a bump publishes nothing.
  Publishing a release by hand still works: it only attaches the exe to it.

## Do
- Keep the output byte-for-byte identical to `tests/fixtures/*expected*`.
- Preserve the reference program's quirks (for example printing negative zero as
  `-0.000`); matching the reference beats being pretty.
- Run conversion and shell dialogs inside async Tauri commands (they execute on tauri's
  thread pool), never on the window thread; `rfd`'s modal IFileDialog runs on the
  calling thread, so calling it from a sync command freezes the window for as long as
  the shell needs (seconds on a network share).
- Check the result of the tauri builder in `app::run_with`; an ignored error means a
  silent exit with no window. The usual cause is a missing Edge WebView2 runtime — the
  error box names it.
- Keep the startup log (`%TEMP%\ISEL_to_Gcode_startup.log`), the panic hook and the error
  box; on a machine where the window never appears, they are the only evidence.
- Define colours in two places: the toolpath palette is the `COLOR_*` constants in
  `preview.rs` (hex strings painted onto the canvas; the canvas adds nothing of its own
  except the grid label color, which mirrors TEXT_FAINT); the UI chrome palette is CSS
  custom properties in `src/styles.css` (`--isel-*`). The preview window paints exactly
  what Rust projects.
- Keep pan/zoom client-side in the preview page (pan offset + the 3-line anchor zoom
  only; the Rust tests pin that formula) so dragging never waits on an IPC round trip;
  all projected geometry comes from `preview.rs`.
- The preview is its own OS window (webview label `preview`, `preview.html?p=<path>`).
  Headless evidence: `ISEL_PREVIEW=<out.ngc>` opens the preview at startup and a second
  top-level window shows up via `EnumWindows`.
- The exe file icon and Win32 version resource come from the resource `tauri-build`
  embeds from `bundle.icon` in `tauri.conf.json` (assets/icon.ico); pure Rust, no windres.
- Conversion log/status strings shown to the user are the python port's wording, spacing
  included (see the byte-parity comments in `src/App.tsx`).
- The UI needs the WebView2 runtime (built into Windows 11, delivered with Edge updates
  on Windows 10). Keep the "could not open a window" error actionable.
- Keep the `.ngc`/`.isel`/`.ncp` CRLF pins in `.gitattributes`.
- After a code change refresh both graphs: codegraph `sync` and `graphify update .`.

## Don't
- In corner feed reduction do not split segments, add or remove points, round coordinates
  or smooth the path; only the F value of existing G1 lines is scaled, and with the
  option off the output stays byte-identical.
- Do not bring back a Python/.NET/Electron layer; delivery is a single exe with no
  installation. Tauri is the one allowed web stack because it embeds everything in the
  Rust exe.
- Never regenerate a golden from current Rust output just to make a test pass.
- Do not switch the local or CI build to MSVC `link.exe`; it crashes on this machine.
- Do not delete the synthetic fixture goldens (`sample_expected.*`); they cover paths the
  real programs never take (default feed when no VEL is present, R arcs, zero-length
  protection).
- Do not bundle an installer or ship `WebView2Loader.dll` next to the exe; CI's artifact
  stays the bare exe.
- `graphify-out/`, `graph_update.log`, `.codegraph/`, `src-tauri/target/`, `dist/` and
  `node_modules/` are not committed.
- Do not run `cargo fmt`: the code is hand-formatted, rustfmt rewrites the whole tree and
  buries the real change.
- Do not use subagents in this project.

## Codegraph + graphify — automatic use

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
