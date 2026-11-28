// Mirrors the serde shapes in src-tauri/src (field names identical on purpose:
// Rust is the single source, these types only describe the wire).

export type Mt = "rapid" | "cut" | "arc";

export interface Camera {
  ox: number;
  oy: number;
  scale: number;
}

export interface SceneLine {
  points: [number, number][];
  color: string;
}

export interface GridLine {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
}

export interface GridLabel {
  x: number;
  y: number;
  text: string;
}

export interface PreviewScene {
  lines: SceneLine[];
  grid: GridLine[];
  labels: GridLabel[];
  camera: Camera;
  background: string;
}

// commands::GcodeSummary — the stats bar's data only; the geometry stays on
// the Rust side and arrives projected, inside PreviewScene.
export interface GcodeSummary {
  counts: { rapid: number; cut: number; arc: number };
  bounds: { min_x: number; max_x: number; min_y: number; max_y: number };
}

export interface PreviewReq {
  path: string;
  width: number;
  height: number;
  iso: boolean;
  camera: Camera | null;
}

// convert worker events (src-tauri/src/commands.rs)
export interface LogMsg {
  line: string;
}
export interface ProgressMsg {
  percent: number;
  text: string;
}
export interface DoneMsg {
  minutes: number;
  output: string;
}
export interface FailedMsg {
  error: string;
}

// Corner feed reduction policy — mirrors CornerParams in src-tauri/src/corner.rs
// (camelCase via serde rename). get_corner_defaults supplies the numbers; this
// file must not copy them, or the two copies would drift.
export interface CornerParams {
  windowMm: number;
  minTurnDeg: number;
  spikeRatio: number;
  rampMaxSegMm: number;
  rampMix: number;
  arcFullRadiusMm: number;
  profile: [number, number][];
}
