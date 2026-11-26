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
import { FileCog, Info, Loader2 } from "lucide-react";
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

// sidebar nav: the converter is the app; "about" gives the fluent nav pattern
// real content instead of a fake second page
const NAV = [
  { id: "converter", label: "Converter", icon: FileCog },
  { id: "about", label: "About", icon: Info },
] as const;

type Page = (typeof NAV)[number]["id"];

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
  const [page, setPage] = useState<Page>("converter");

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

  // bare-browser peek (no tauri host): skip the window IPC, keep the UI alive
  const isTauri = "__TAURI_INTERNALS__" in window;

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
    if (!isTauri) return; // bare-browser peek: no IPC host to talk to
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
    <div className="flex h-screen flex-col overflow-hidden rounded-lg border border-white/[0.08] text-[13px] text-zinc-200">
      <Titlebar version={version} />

      <div className="flex min-h-0 flex-1">
        {/* sidebar sits straight on the mica; lighter than the content surface */}
        <nav className="flex w-[220px] shrink-0 flex-col gap-1 bg-white/[0.02] p-3 select-none">
          {NAV.map((item) => {
            const active = page === item.id;
            const Icon = item.icon;
            return (
              <button
                key={item.id}
                onClick={() => setPage(item.id)}
                className={`relative flex items-center gap-2 rounded-md px-3 py-2 text-left transition-colors duration-150 ${
                  active ? "bg-white/[0.08] text-white" : "text-zinc-400 hover:bg-white/[0.04] hover:text-zinc-200"
                }`}
              >
                {active && (
                  <span className="absolute left-0 top-1/2 h-4 w-[3px] -translate-y-1/2 rounded-full bg-fluent-accent" />
                )}
                <Icon size={15} strokeWidth={1.5} />
                {item.label}
              </button>
            );
          })}
        </nav>

        <main className="flex min-w-0 flex-1 flex-col rounded-tl-lg bg-isel-panel/85 px-5 py-4">
          <div className="flex items-center justify-between gap-3">
            <h1 className="text-[20px] font-semibold text-white">{NAV.find((n) => n.id === page)?.label}</h1>
            {page === "converter" && (
              <div className="flex gap-2">
                <Button variant="secondary" disabled={busy} onClick={browse} className="w-28">
                  Browse
                </Button>
                <Button disabled={busy} onClick={convert} className="w-28 font-semibold">
                  {converting ? "Converting..." : "Convert"}
                </Button>
                <Button
                  variant="secondary"
                  disabled={!lastOutput}
                  onClick={preview}
                  className="w-28 font-semibold text-[#6ccb5f] hover:bg-[rgba(108,203,95,0.16)]"
                >
                  Preview
                </Button>
              </div>
            )}
          </div>

          {page === "about" ? (
            <section className="fluent-card select-text mt-4">
              <div>ISEL &rarr; G-code Converter v{version}</div>
              <div className="mt-2 text-zinc-400">Converts ISEL/CAM heater-winding programs to machine G-code.</div>
              <div className="mt-2 font-mono text-[12px] text-zinc-300">
                CLI: ISEL_to_Gcode.exe convert in.isel out.ngc [--arcs] [--corners]
              </div>
            </section>
          ) : (
            <div className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto pt-4">
              <section className="fluent-card shrink-0">
                <Label className="text-[11px] tracking-widest text-zinc-500">ISEL FILE</Label>
                <Input
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  spellCheck={false}
                  className="mt-1 bg-[#1f1f1f] font-mono text-[12px] text-zinc-100 placeholder:text-zinc-600"
                  placeholder="Drop an ISEL file here or click Browse"
                />
              </section>

              <div className="flex shrink-0 gap-3">
                <section className="fluent-card flex-1 space-y-2">
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
                </section>

                <section className="fluent-card flex-1">
                  <div className="flex items-center gap-2">
                    {busy && <Loader2 size={14} className="animate-spin text-fluent-accent" />}
                    <span className={statusColor}>{status}</span>
                  </div>
                  <Progress
                    value={Math.round(progress * 100)}
                    className="mt-2 [&_[data-slot=progress-track]]:bg-white/[0.06]"
                  />
                </section>
              </div>

              <section className="fluent-card flex min-h-0 flex-1 flex-col">
                <Label className="mb-1 text-[11px] tracking-widest text-zinc-500">CONVERSION LOG</Label>
                <div
                  ref={logBox}
                  className={`select-text min-h-0 flex-1 overflow-y-auto rounded-lg border p-2 font-mono text-[11.5px] leading-5 ${
                    dropping ? "border-fluent-accent/60 bg-fluent-accent/10" : "border-white/[0.08] bg-isel-deep/80"
                  }`}
                >
                  {log.map((line, i) => (
                    <div key={i} className="whitespace-pre-wrap text-isel-log">
                      {line}
                    </div>
                  ))}
                </div>
              </section>
            </div>
          )}
        </main>
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
        <DialogContent className="border-white/[0.14] bg-isel-dialog text-zinc-100 backdrop-blur-xl">
          <DialogHeader>
            <DialogTitle>{msg?.title}</DialogTitle>
          </DialogHeader>
          <div className="select-text whitespace-pre-wrap text-[13px] text-zinc-300">{msg?.body}</div>
          <DialogFooter>
            <Button onClick={() => setMsg(null)}>OK</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
