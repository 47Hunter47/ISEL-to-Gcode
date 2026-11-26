# Tauri v2 UI migration — design

## Intent
Replace the egui/eframe front end with a Raycast/Linear-style dark UI built on
Tauri v2 + React (TypeScript) + Tailwind CSS + shadcn/ui + lucide-react, with an
OS-native blurred window (acrylic/mica) and a custom titlebar. The conversion
program itself must keep working exactly as it does today.

## Goals
- All Rust business logic (`convert.rs`, `corner.rs`, `gcode.rs`, preview math)
  is moved, not rewritten; output stays byte-identical to `tests/fixtures/`.
- The CLI (`ISEL_to_Gcode convert <in> <out> [--arcs] [--corners]`), the golden
  tests, the startup log, the panic hook + error box and the windres exe icon
  all keep working.
- Delivery stays a single, no-install exe: the built web frontend is embedded
  into the exe by Tauri's build script; CI keeps its current shape
  (test job + `cargo build --release` producing `ISEL_to_Gcode.exe`).

## Non-goals
- No installer/bundler output (no NSIS/MSI). No new runtime daemons.
- No changes to conversion semantics, feed handling or G-code formatting.

## Anti-goals
- Anything that makes the exe depend on files shipped next to it.
- Per-frame IPC round-trips for preview pan/zoom (old workshop machines must
  stay smooth).

## Constraints
- Windows GNU toolchain only (w64devkit locally, `C:\mingw64` on CI);
  `-C link-self-contained=yes` stays; MSVC `link.exe` is banned on this host.
- `tools/ship.sh` derives the version bump from `Cargo.toml`; it must keep
  finding the version line. Two workflows build/attach the exe automatically.
- Tauri needs the WebView2 runtime: built into Win11, delivered with Edge on
  Win10. If it is missing the app must show a clear error box, not a blank
  window.

## Approaches considered
1. **Tauri v2, assets embedded, CI shape unchanged (chosen).** Single exe,
   GNU-compatible spike first; CI just gains `npm run build` before
   `cargo build --release`.
2. Restyle egui only — rejected: the user explicitly asked to remove the egui
   layer and its rendering stack.
3. Tauri with bundled installer — rejected: violates the single-exe rule.

## Decided design
- `rust/` is renamed to `src-tauri/` (git history preserved); the Vite +
  React + TS frontend lives at the repo root (`src/`, `dist/`).
- Commands (async, so the window thread never blocks): `pick_file`,
  `save_file(default_name)`, `convert(in, out, arcs, corners)` streaming
  `log` / `progress` / `done` / `failed` events, and `parse_gcode(path)`
  returning the parsed toolpath as JSON. `convert_file` keeps its signature.
- `main.rs` parses the `convert` CLI before Tauri starts; a dropped-file
  prefill arg still seeds the input field; `run` result is checked, on failure
  the existing startup log + error box explain why.
- Window: `transparent: true`, `decorations: false`; `setup` applies
  `apply_mica`/`apply_acrylic` (Windows) or `apply_vibrancy` (macOS) guarded
  by a runtime check; custom titlebar with `data-tauri-drag-region` and
  minimize/maximize/close buttons (capabilities: `core:default` plus the
  window permissions used).
- Preview stays a separate OS window (a second WebviewWindow). `preview.rs`
  keeps fit/zoom/iso/grid math in Rust, freed of egui types, with its tests
  ported 1:1; the window receives the parsed scene once as JSON and does
  pan/zoom/fit/iso and drawing on a JS canvas (formula port of
  `project`/`fit`/`zoom_at`, so dragging never round-trips).
- Removed with egui: `theme.rs` and its unit test, `ISEL_SHOT` BMP evidence,
  `ISEL_RENDERER`. The SCADA palette and the new translucent-dark UI palette
  live as CSS variables; monospace log, status light and the two option
  checkboxes keep their semantics.
- `AGENTS.md` rules that only existed for eframe are rewritten for Tauri in
  the same change; README usage/requirements updated; Obsidian project note
  carries the status at ship time.

## Left to the builder
File names inside the frontend, the exact shadcn component set (expected:
button, checkbox, progress, scroll-area, dialog), exact TS payload types
mirroring the Rust serde shapes, and the vite/tsconfig scaffolding details.
