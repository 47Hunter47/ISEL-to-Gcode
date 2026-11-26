import { Flame, Minus, Square, X } from "lucide-react";
import { getCurrentWindow } from "@tauri-apps/api/window";

const btn =
  "flex h-9 w-11 items-center justify-center text-zinc-400 transition-colors duration-150 hover:bg-white/[0.07] hover:text-zinc-100";

// decorations:false window: this 36px strip is the titlebar. The
// data-tauri-drag-region attribute is what makes the empty area draggable
// (tauri listens for it). Close hover is the Win11 red, not the Win10 one.
export default function Titlebar({ version }: { version: string }) {
  // outside tauri (bare-browser dev peek) there is no window to control
  const w = "__TAURI_INTERNALS__" in window ? getCurrentWindow() : null;
  return (
    <div data-tauri-drag-region className="flex h-9 shrink-0 items-center justify-between pl-3 select-none">
      <div data-tauri-drag-region className="pointer-events-none flex items-center gap-2">
        <Flame size={14} strokeWidth={1.5} className="text-fluent-accent" />
        <span className="text-[12px] tracking-wide text-zinc-300/80">
          ISEL &rarr; G-code Converter <span className="font-mono text-zinc-600">v{version}</span>
        </span>
      </div>
      <div className="flex">
        <button aria-label="Minimize" className={btn} onClick={() => w?.minimize()}>
          <Minus size={15} strokeWidth={1.5} />
        </button>
        <button aria-label="Maximize" className={btn} onClick={() => w?.toggleMaximize()}>
          <Square size={11} strokeWidth={1.5} />
        </button>
        <button
          aria-label="Close"
          className={`${btn} hover:bg-fluent-close hover:text-white hover:[&>svg]:text-white`}
          onClick={() => w?.close()}
        >
          <X size={15} strokeWidth={1.5} />
        </button>
      </div>
    </div>
  );
}
