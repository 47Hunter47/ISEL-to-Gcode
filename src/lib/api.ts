// Typed wrappers over every invoke of the Rust command layer. Argument object
// keys are camelCase: tauri maps them to the snake_case command parameters.
import { invoke } from "@tauri-apps/api/core";
import type { DoneMsg, FailedMsg, GcodeParsed, LogMsg, PreviewReq, PreviewScene, ProgressMsg } from "./types";

export const api = {
  checkInput: (path: string) => invoke<void>("check_input", { path }),
  pickFile: () => invoke<string | null>("pick_file"),
  saveFile: (defaultName: string) => invoke<string | null>("save_file", { defaultName }),
  convert: (input: string, output: string, arcs: boolean, corners: boolean) =>
    invoke<void>("convert", { input, output, arcs, corners }),
  getPrefill: () => invoke<string | null>("get_prefill"),
  getVersion: () => invoke<string>("get_version"),
  openPreview: (path: string) => invoke<string | null>("open_preview_window", { path }),
  parseGcode: (path: string) => invoke<GcodeParsed>("parse_gcode", { path }),
  previewScene: (req: PreviewReq) => invoke<PreviewScene>("preview_view", { req }),
};

export type { DoneMsg, FailedMsg, LogMsg, ProgressMsg };
