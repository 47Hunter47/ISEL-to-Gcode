# Tauri v2 UI migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (AGENTS.md bans subagents here). Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the egui front end with Tauri v2 + React/TS + Tailwind + shadcn/ui without touching conversion logic; delivery stays a single exe CI builds automatically.

**Architecture:** `rust/` → `src-tauri/`; commands + events bridge the React windows to the unchanged Rust core; the built `dist/` is embedded into the exe by the tauri build script so `cargo build --release` (now CI-prepended by `npm run build`) still yields one file.

**Tech Stack:** tauri 2, tauri-build 2, window-vibrancy, rfd (unchanged), serde; Vite + React + TypeScript, Tailwind CSS v4, shadcn/ui, lucide-react, @tauri-apps/api.

**Spec:** `docs/superpowers/specs/2026-11-26-tauri-ui-design.md`

## Global Constraints
- Windows GNU only; keep `-C link-self-contained=yes`; MSVC link.exe banned; PATH needs w64devkit bin (locally) / C:\mingw64 (CI).
- Conversion output stays byte-identical to `tests/fixtures/*expected*`; `convert_file`'s signature is unchanged.
- CLI stays `ISEL_to_Gcode convert <in> <out> [--arcs] [--corners]` with `TIME`/`PROGRESS`/`ERROR` console lines.
- Single exe: no bundler output; CI artifacts stay `ISEL_to_Gcode.exe` only.
- Public repo: English code/docs/commits; no personal machine paths in docs.
- No `cargo fmt`. Do not regenerate goldens.

## Review Focus
- Missing WebView2 runtime: user must get the rfd error box + startup log line, not a silent blank window.
- A slow shell dialog (network share) must not freeze the window — every `rfd` call lives in an async command's worker thread.
- CLI regression: `convert` must be handled before `tauri::Builder` so CI/tests keep the console contract.
- Preview pan/zoom over IPC: scene transfer happens once (static geometry); per-view events only re-run the camera; upgrade trigger is documented in code.
- `transparent` + acrylic/mica must degrade to the opaque CSS background (not black) when vibrancy APIs are unavailable (RDP, older Windows).

---

### Task 0: GNU spike gate (no repo files)

**Files:** scratch `tauri-spike/src-tauri` (already scaffolded)

- [x] `cargo check` finishes with 0 errors on `stable-x86_64-pc-windows-gnu`.
- [x] `cargo build --release` produces `spike.exe`; `ls target/release/` shows no extra DLLs (WebView2Loader statically linked).
- [ ] If either fails: stop, report evidence to the user before porting anything.

### Task 1: Repo restructure (behavior unchanged)

**Files:**
- Move: `rust/` → `src-tauri/` (git mv; `target/` is untracked, just moved/removed).
- Modify: `tools/ship.sh` (all `rust/Cargo.toml|Cargo.lock` paths → `src-tauri/…`), `.gitignore` (`rust/target/` → `src-tauri/target/`, add `node_modules/`), `.github/workflows/build.yaml` + `main.yml` (`working-directory: rust` → `src-tauri`; artifact `src-tauri/target/release/ISEL_to_Gcode.exe`), `README.md` build section.

- [ ] `git mv rust src-tauri`
- [ ] Update every `rust/` path reference: `grep -rn 'rust/' tools .github README.md .gitignore` returns nothing build-relevant.
- [ ] `cd src-tauri && cargo test` passes (w64devkit on PATH).
- [ ] Commit `refactor: move Rust program to src-tauri for Tauri layout`

### Task 2: Rust core freed of egui + commands layer

**Files:**
- Modify: `src-tauri/Cargo.toml` (drop eframe/egui/image; add `tauri = "2"`, `window-vibrancy`, serde derives; keep rfd/regex; add tauri-build), `src-tauri/build.rs` (tauri_build + existing windres logic), `src-tauri/src/lib.rs`, `src-tauri/src/gcode.rs` (add `Serialize` derives only), `src-tauri/src/preview.rs`, `src-tauri/src/main.rs`.
- Create: `src-tauri/src/commands.rs`.
- Delete: `src-tauri/src/theme.rs`, `src-tauri/tests/theme_unit.rs`, egui UI half of `app.rs` (replaced by `commands.rs` + `tauri.rs`).

**Interfaces:**
- Produces: `commands::pick_file() -> Option<String>`, `commands::save_file(default_name: String) -> Option<String>`, `commands::convert(app: tauri::AppHandle, input: String, output: String, arcs: bool, corners: bool) -> Result<(), String>` (emits `convert-log(String)`, `convert-progress{percent,text}`, `convert-done{minutes,output}`, `convert-failed{error}`), `commands::parse_gcode(path: String) -> Result<GcodeJson, String>`, `commands::preview_scene(view: PreviewView) -> Result<PreviewScene, String>`, `commands::open_preview_window(app, path) -> Result<(),String>` (second WebviewWindow `preview.html?f=<path>`), `commands::get_prefill() -> Option<String>`.
- `PreviewView { width: f32, height: f32, iso: bool, ox: f32, oy: f32, scale: f32, path: String }` → `PreviewScene { lines: Vec<{ points: Vec<[f32;2]>, color: String }>, grid: Vec<…>, axis: …, fit: {ox,oy,scale} }` (camera math ported from `preview.rs` 1:1, tests ported 1:1).

- [ ] Preview camera tests first (ported names): `fit_centers_and_scales_to_rect`, `zoom_keeps_anchor_fixed`, `iso_projection_matches_python`, `topdown_golden_fits_inside_canvas`, `arcs_golden_yields_arc_and_cut_polylines` — run, fail.
- [ ] Implement camera refactor (egui types → plain f32/`[f32;2]`), ported tests pass.
- [ ] `commands.rs` per Interfaces; conversion worker + emitter mirrors old `Msg` semantics (done-time formatting `"{} min {} sec"` identical).
- [ ] `tauri.rs`: setup = panic hook + startup log (same file/name), vibrancy mica→acrylic fallback, title/version, prefill managed state; `run()` checks builder result → `show_error` on Err (WebView2-missing message names `%TEMP%\ISEL_to_Gcode_startup.log`).
- [ ] `main.rs`: CLI branch unchanged first; then `isel_to_gcode::tauri::run_with(prefill)`.
- [ ] `cargo test` + `cargo check` clean. CLI golden smoke: `ISEL_to_Gcode convert` run against `tests/fixtures/sample.isel` matches golden bytes.
- [ ] Commit `refactor: expose Rust core as Tauri commands (egui dependency removed)`

### Task 3: Frontend app shell (build gate)

**Files:** Create root `package.json`, `vite.config.ts`, `tsconfig.json`, `index.html`, `preview.html`, `src/main.tsx`, `src/App.tsx`, `src/PreviewPage.tsx`, `src/components/ui/*` (shadcn), `src/lib/{types.ts,api.ts}`, `src/styles.css`; `src-tauri/tauri.conf.json`, `src-tauri/capabilities/default.json`.

- [ ] Scaffold Vite React-TS + Tailwind v4 + shadcn init (dark) + lucide-react + @tauri-apps/api; `npx shadcn add button checkbox progress dialog scroll-area separator tooltip`.
- [ ] `types.ts` mirrors Task 2 serde shapes exactly (field names identical); `api.ts` typed invoke/event wrappers.
- [ ] UI per spec: titlebar (drag region + min/max/close), bg `#121214`/80 + backdrop blur, cards `border-white/[0.08] bg-white/[0.03] rounded-xl`, indigo-600 primary, hover `white/[0.07]`, transitions 150ms, monospace log pane (phosphor-mint), footer status bar with shortcuts; drop-zone wires Tauri drag-drop events; dialog texts byte-match old completion/error strings.
- [ ] Preview page: canvas renderer over `preview_scene` result; FIT button + Isometric toggle; once-per-open geometry fetch (comment the shortcut).
- [ ] `npm run build` (tsc + vite) zero errors; `cargo check` in src-tauri still clean.
- [ ] Commit `feat: React + shadcn UI shell over Tauri commands`

### Task 4: CI + docs

**Files:** Modify both workflows (setup-node@v4 + `npm ci` + `npm run build` before cargo steps; npm cache), `AGENTS.md` (eframe-specific rules → Tauri rules), `README.md` (WebView2 requirement line, build instructions).

- [ ] Workflows: node 22 LTS + build frontend in repo root, cargo steps from `src-tauri/`.
- [ ] AGENTS.md rewritten rules keep startup-log/dialog/golden/ship invariants; drop ISEL_SHOT/ISEL_RENDERER/theme rules.
- [ ] Commit `chore: CI builds the embedded web front end; docs updated`

### Task 5: Verify + ship

- [ ] Fresh `cargo test` + `cargo build --release` + `npm run build` all green; `ls src-tauri/target/release/ISEL_to_Gcode.exe` and confirm nothing else needed next to it (copy to empty dir, run).
- [ ] Launch exe: startup log grows with window-shown line; window appears (computer_use screenshot); kill cleanly. CLI: `convert` sample fixture matches golden bytes.
- [ ] `./tools/ship.sh "feat: ultra-modern Tauri v2 UI (React + shadcn/ui)"` → version 2.4.0, push.
- [ ] `gh` run watch: CI + Release both green, release v2.4.0 attaches exe. Update Obsidian `Projeler/ISEL_to_Gcode.md` status.
