import { Maximize2, Minus, X } from "lucide-react";
import { getCurrentWindow } from "@tauri-apps/api/window";

const btn =
  "flex h-8 w-11 items-center justify-center text-zinc-400 transition-colors duration-150 hover:bg-white/[0.07] hover:text-zinc-100";

// decorations:false window: this strip is the titlebar. The data-tauri-drag-region
// attribute is what makes the empty area draggable (tauri listens for it).
export default function Titlebar({ version }: { version: string }) {
  const w = getCurrentWindow();
  return (
    <div data-tauri-drag-region className="flex h-9 shrink-0 items-center justify-between pl-3 select-none">
      <div data-tauri-drag-region className="pointer-events-none text-[12px] tracking-wide text-zinc-400">
        ISEL &rarr; G-code Converter <span className="font-mono text-zinc-600">v{version}</span>
      </div>
      <div className="flex">
        <button aria-label="Minimize" className={btn} onClick={() => w.minimize()}>
          <Minus size={15} strokeWidth={1.5} />
        </button>
        <button aria-label="Maximize" className={btn} onClick={() => w.toggleMaximize()}>
          <Maximize2 size={12} strokeWidth={1.5} />
        </button>
        <button
          aria-label="Close"
          className={`${btn} hover:bg-[#E81123] hover:text-white hover:[&>svg]:text-white`}
          onClick={() => w.close()}
        >
          <X size={15} strokeWidth={1.5} />
        </button>
      </div>
    </div>
  );
}
