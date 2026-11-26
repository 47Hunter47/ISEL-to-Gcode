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

// gcode::GcodeParsed (only the header fields the preview window paints).
export interface GcodeParsed {
  groups: [Mt, [number, number, number][]][];
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
