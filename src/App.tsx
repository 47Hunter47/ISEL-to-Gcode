import { useEffect, useRef, useState } from "react";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Progress } from "@/components/ui/progress";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Separator } from "@/components/ui/separator";
import { Loader2 } from "lucide-react";
import Titlebar from "./Titlebar";
import { api } from "./lib/api";
import type { DoneMsg, FailedMsg, LogMsg, ProgressMsg } from "./lib/types";

// Log/status strings below are kept byte-identical to the egui app's python
// port (spacing included); the reference program's wording is the contract.
const SEP = "--------------------------------------------------";
const initialLog = [
  "ISEL to G-code Converter Ready",
  "Drag and drop a file or click Browse to start",
  SEP,
];

interface DialogMsg {
  title: string;
  body: string;
}

// python's file_stem for the default save name
const stem = (path: string) => {
  const name = path.replace(/\\/g, "/").split("/").pop() ?? "";
  return name.replace(/\.[^.]*$/, "");
};

const modeLabel = (arcs: boolean) => (arcs ? "G2/G3 arc commands" : "G1 linearised");

export default function App() {
  const [input, setInput] = useState("");
  const [useArc, setUseArc] = useState(false);
  const [corner, setCorner] = useState(false);
  const [log, setLog] = useState<string[]>(initialLog);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState("Ready");
  const [converting, setConverting] = useState(false);
  const [waitingDialog, setWaitingDialog] = useState(false);
  const [lastOutput, setLastOutput] = useState<string | null>(null);
  const [msg, setMsg] = useState<DialogMsg | null>(null);
  const [dropping, setDropping] = useState(false);
  const [version, setVersion] = useState("");

  const busy = converting || waitingDialog;
  // mirrors for the async callbacks (state inside closures goes stale)
  const vals = useRef({ input, useArc, corner });
  vals.current = { input, useArc, corner };
  const unlistens = useRef<UnlistenFn[]>([]);
  const logBox = useRef<HTMLDivElement>(null);

  const pushLog = (line: string) => setLog((l) => [...l, line]);

  const browse = async () => {
    setWaitingDialog(true);
    setStatus("Waiting for the file dialog...");
    const p = await api.pickFile();
    setWaitingDialog(false);
    setStatus("Ready");
    if (p) {
      setInput(p);
      pushLog(`File selected: ${p}`);
    }
  };

  const cleanup = () => {
    unlistens.current.forEach((u) => u());
    unlistens.current = [];
  };

  const convert = async () => {
    const v = vals.current;
    try {
      await api.checkInput(v.input);
    } catch (e) {
      setMsg({ title: "Error", body: String(e) });
      return;
    }
    setWaitingDialog(true);
    setStatus("Waiting for the file dialog...");
    const out = await api.saveFile(`${stem(v.input)}.ngc`);
    setWaitingDialog(false);
    if (!out) {
      setStatus("Ready"); // cancelled: nothing was started, nothing to undo
      return;
    }
    // begin_convert's header lines, same strings as the egui port
    setLog([
      "Starting conversion...",
      `Input:   ${v.input}`,
      `Output:  ${out}`,
      `Mode:    ${v.useArc ? "G2/G3 arc commands + full-circle detection (R format)" : "G1 linearised"}`,
      `Corners:  ${v.corner ? "adaptive feed reduction (G1 only, geometry unchanged)" : "off"}`,
      SEP,
    ]);
    setStatus("Starting conversion...");
    setProgress(0);
    setConverting(true);

    unlistens.current = await Promise.all([
      listen<LogMsg>("convert-log", (e) => pushLog(e.payload.line)),
      listen<ProgressMsg>("convert-progress", (e) => {
        setProgress(e.payload.percent / 100); // progress updates the status label only (python parity)
        setStatus(e.payload.text);
      }),
      listen<DoneMsg>("convert-done", (e) => {
        const m = Math.floor(e.payload.minutes);
        const s = Math.floor((e.payload.minutes - m) * 60);
        setConverting(false);
        setLastOutput(e.payload.output);
        setProgress(1);
        setStatus("Conversion complete!");
        setLog((l) => [
          ...l,
          SEP,
          `\u2713 Conversion completed  (${modeLabel(vals.current.useArc)})`,
          `\u23F1 Estimated program time: ${m} min ${s} sec`,
          "\u26A0 Program time may change according to machine parameters",
        ]);
        setMsg({
          title: "Completed",
          body: `Conversion finished!\nMode: ${modeLabel(vals.current.useArc)}\n\nEstimated time: ${m} min ${s} sec`,
        });
        cleanup();
      }),
      listen<FailedMsg>("convert-failed", (e) => {
        setConverting(false);
        setStatus("Error occurred");
        pushLog(`\u2717 Error: ${e.payload.error}`);
        setMsg({ title: "Error", body: e.payload.error });
        cleanup();
      }),
    ]);

    try {
      await api.convert(v.input, out, v.useArc, v.corner);
    } catch (e) {
      setConverting(false);
      setStatus("Error occurred");
      pushLog(`\u2717 Error: ${e}`);
      setMsg({ title: "Error", body: String(e) });
      cleanup();
    }
  };

  const preview = async () => {
    if (!lastOutput) return;
    try {
      // a vanished output file is a silent no-op (decided Rust side)
      const m = await api.openPreview(lastOutput);
      if (m) setMsg({ title: "Preview", body: m });
    } catch (e) {
      setMsg({ title: "Preview Error", body: String(e) });
    }
  };

  useEffect(() => {
    api.getVersion().then(setVersion).catch(() => setVersion("?"));
    api.getPrefill().then((p) => {
      if (p) {
        setInput(p);
        pushLog(`File selected: ${p}`);
      }
    });
    // drag & drop replaces egui's dropped_files; checkInput verifies it is a
    // real file (the web layer has no stat of its own)
    let un: UnlistenFn | undefined;
    getCurrentWebview()
      .onDragDropEvent((e) => {
        const p = e.payload;
        if (p.type === "enter" || p.type === "over") setDropping(true);
        else if (p.type === "leave") setDropping(false);
        else if (p.type === "drop") {
          setDropping(false);
          for (const path of p.paths) {
            api.checkInput(path).then(
              () => {
                setInput(path);
                pushLog(`File imported: ${path}`);
              },
              () => pushLog(`Error: Not a valid file: ${path}`),
            );
          }
        }
      })
      .then((u: UnlistenFn) => (un = u));
    return () => {
      un?.();
      unlistens.current.forEach((u) => u());
    };
  }, []);

  useEffect(() => {
    const box = logBox.current;
    if (box) box.scrollTop = box.scrollHeight; // stick to the bottom like the old ScrollArea
  }, [log]);

  // no deps on purpose: the handler rebinds each render so it always sees
  // current busy/dialog state
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Enter" && !busy && !msg) {
        e.preventDefault();
        void convert();
      } else if ((e.ctrlKey || e.metaKey) && (e.key === "b" || e.key === "o") && !busy) {
        e.preventDefault();
        void browse();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  const statusColor =
    status === "Error occurred"
      ? "text-isel-err"
      : status === "Conversion complete!"
        ? "text-isel-ok"
        : "text-zinc-300";

  return (
    <div className="flex h-screen flex-col overflow-hidden rounded-lg border border-white/[0.08] bg-isel-panel/80 text-[13px] text-zinc-200 backdrop-blur-xl">
      <Titlebar version={version} />

      <div className="flex-1 space-y-4 overflow-hidden p-4">
        <div>
          <Label className="text-[11px] tracking-widest text-zinc-500">ISEL FILE</Label>
          <Input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            spellCheck={false}
            className="mt-1 border-white/[0.08] bg-white/[0.03] font-mono text-[12px] text-zinc-100 placeholder:text-zinc-600"
            placeholder="Drop an ISEL file here or click Browse"
          />
        </div>

        <div className="flex gap-2">
          <Button
            variant="secondary"
            disabled={busy}
            onClick={browse}
            className="w-28 border-white/[0.08] bg-white/[0.05] text-zinc-100 transition-all duration-150 hover:bg-white/[0.09]"
          >
            Browse
          </Button>
          <Button
            disabled={busy}
            onClick={convert}
            className="w-28 bg-indigo-600 font-semibold text-white transition-all duration-150 hover:bg-indigo-500 disabled:opacity-50"
          >
            {converting ? "Converting..." : "Convert"}
          </Button>
          <Button
            disabled={!lastOutput}
            onClick={preview}
            className="w-28 border-emerald-500/20 bg-emerald-500/10 font-semibold text-emerald-300 transition-all duration-150 hover:bg-emerald-500/20"
          >
            Preview
          </Button>
        </div>

        <div className="space-y-2 rounded-xl border border-white/[0.08] bg-white/[0.03] p-4">
          <label className="flex items-center gap-2">
            <Checkbox id="arcs" checked={useArc} onCheckedChange={(c) => setUseArc(c === true)} />
            <Label htmlFor="arcs" className="cursor-pointer font-normal">
              Use G2/G3 arc commands&nbsp;&nbsp;(smaller file, requires arc-capable controller)
            </Label>
          </label>
          <label className="flex items-center gap-2">
            <Checkbox id="corners" checked={corner} onCheckedChange={(c) => setCorner(c === true)} />
            <Label htmlFor="corners" className="cursor-pointer font-normal">
              Corner Feed Reduction&nbsp;&nbsp;(slower feed around sharp G1 corners)
            </Label>
          </label>
        </div>

        <div>
          <div className="flex items-center gap-2">
            {busy && <Loader2 size={14} className="animate-spin text-indigo-400" />}
            <span className={statusColor}>{status}</span>
          </div>
          <Progress
            value={Math.round(progress * 100)}
            className="mt-2 [&_[data-slot=progress-track]]:bg-white/[0.06]"
          />
        </div>

        <div className="flex min-h-0 flex-1 flex-col">
          <Label className="mb-1 text-[11px] tracking-widest text-zinc-500">CONVERSION LOG</Label>
          <div
            ref={logBox}
            className={`min-h-0 flex-1 overflow-y-auto rounded-xl border p-3 font-mono text-[11.5px] leading-5 ${
              dropping ? "border-indigo-400/60 bg-indigo-500/10" : "border-white/[0.08] bg-isel-deep/80"
            }`}
          >
            {log.map((line, i) => (
              <div key={i} className="whitespace-pre-wrap text-isel-log">
                {line}
              </div>
            ))}
          </div>
        </div>
      </div>

      <Separator className="bg-white/[0.08]" />
      <footer className="flex h-7 shrink-0 items-center justify-between px-4 text-[11px] text-zinc-500 select-none">
        <span>
          Enter <span className="text-zinc-600">convert</span> · Ctrl+B <span className="text-zinc-600">browse</span> · drop
          a file anywhere
        </span>
        <span className="font-mono">v{version}</span>
      </footer>

      <Dialog open={msg !== null} onOpenChange={(o) => !o && setMsg(null)}>
        <DialogContent className="border-white/[0.08] bg-isel-dialog text-zinc-100 backdrop-blur-xl">
          <DialogHeader>
            <DialogTitle>{msg?.title}</DialogTitle>
          </DialogHeader>
          <div className="whitespace-pre-wrap text-[13px] text-zinc-300">{msg?.body}</div>
          <DialogFooter>
            <Button onClick={() => setMsg(null)} className="bg-indigo-600 text-white hover:bg-indigo-500">
              OK
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
