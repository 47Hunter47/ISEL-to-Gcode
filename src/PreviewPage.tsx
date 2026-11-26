import { useCallback, useEffect, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { api } from "./lib/api";
import type { Camera, GcodeParsed, PreviewScene } from "./lib/types";

// Canvas paints exactly what preview.rs projects; the only client-side math is
// the pan offset and the 3-line anchor zoom (mirror of preview.rs's
// zoom_at / preview.rs's test pins that formula). Keeping pan client-side
// means a drag never waits on an IPC round trip.
// The label color mirrors preview.rs's python-side TEXT_FAINT (#2F5F73).
const LABEL_COLOR = "#2F5F73";
const nf = new Intl.NumberFormat("en-US");

const pathFromUrl = () => new URLSearchParams(window.location.search).get("p") ?? "";

export default function PreviewPage() {
  const [path, setPath] = useState(pathFromUrl);
  const [iso, setIso] = useState(false);
  const [scene, setScene] = useState<PreviewScene | null>(null);
  const [stats, setStats] = useState<GcodeParsed | null>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const camRef = useRef<Camera | null>(null); // null asks the server to fit
  const inFlight = useRef(false);
  const dirty = useRef(false);
  const drag = useRef<{ x: number; y: number } | null>(null);

  const request = useCallback(async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    if (inFlight.current) {
      dirty.current = true; // coalesce rapid pan/zoom events into one view per round trip
      return;
    }
    inFlight.current = true;
    try {
      do {
        dirty.current = false;
        const { width, height } = canvas.getBoundingClientRect();
        const s = await api.previewScene({
          path,
          width: Math.max(1, Math.round(width)),
          height: Math.max(1, Math.round(height)),
          iso,
          camera: camRef.current,
        });
        camRef.current = s.camera;
        setScene(s);
      } while (dirty.current);
    } catch (e) {
      console.error("preview_view failed:", e);
    } finally {
      inFlight.current = false;
    }
  }, [path, iso]);

  useEffect(() => {
    camRef.current = null; // new file: refit
    api.parseGcode(path).then(setStats).catch(() => setStats(null));
    request();
  }, [path, request]);

  // refit when the window resizes, like the old rect-change refit
  useEffect(() => {
    const ro = new ResizeObserver(() => {
      camRef.current = null;
      request();
    });
    if (wrapRef.current) ro.observe(wrapRef.current);
    return () => ro.disconnect();
  }, [request]);

  // a reused preview window gets its new file through this event
  useEffect(() => {
    const un = listen<{ path: string }>("load-path", (e) => setPath(e.payload.path));
    return () => void un.then((u) => u());
  }, []);

  const fit = () => {
    camRef.current = null;
    request();
  };

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const onWheel = (e: WheelEvent) => {
      e.preventDefault();
      const cam = camRef.current;
      if (!cam) return;
      const rect = canvas.getBoundingClientRect();
      const cx = e.clientX - rect.left;
      const cy = e.clientY - rect.top;
      const f = e.deltaY > 0 ? 1 / 1.15 : 1.15; // wheel up = zoom in (python parity)
      camRef.current = { ox: cx + (cam.ox - cx) * f, oy: cy + (cam.oy - cy) * f, scale: cam.scale * f };
      request();
    };
    canvas.addEventListener("wheel", onWheel, { passive: false });
    return () => canvas.removeEventListener("wheel", onWheel);
  }, [request]);

  const onMouseDown = (e: React.MouseEvent) => {
    drag.current = { x: e.clientX, y: e.clientY };
    const onMove = (ev: MouseEvent) => {
      const cam = camRef.current;
      const d = drag.current;
      if (!cam || !d) return;
      camRef.current = { ...cam, ox: cam.ox + (ev.clientX - d.x), oy: cam.oy + (ev.clientY - d.y) };
      drag.current = { x: ev.clientX, y: ev.clientY };
      request();
    };
    const onUp = () => {
      drag.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  // paint everything the Rust scene carries; nothing is recomputed here
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !scene) return;
    const dpr = window.devicePixelRatio || 1;
    const { width, height } = canvas.getBoundingClientRect();
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(height * dpr);
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = scene.background;
    ctx.fillRect(0, 0, width, height);

    ctx.lineWidth = 1;
    for (const g of scene.grid) {
      ctx.strokeStyle = g.color;
      ctx.beginPath();
      ctx.moveTo(g.x1, g.y1);
      ctx.lineTo(g.x2, g.y2);
      ctx.stroke();
    }
    ctx.fillStyle = LABEL_COLOR;
    ctx.font = "7px monospace";
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    for (const l of scene.labels) ctx.fillText(l.text, l.x, l.y);

    for (const line of scene.lines) {
      ctx.strokeStyle = line.color;
      ctx.beginPath();
      line.points.forEach(([x, y], i) => (i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y)));
      ctx.stroke();
    }
  }, [scene]);

  const b = stats?.bounds;
  const legend = (color: string, name: string) => (
    <span className="flex items-center gap-1.5">
      <span className="inline-block h-[2px] w-5" style={{ background: color }} />
      {name}
    </span>
  );

  return (
    <div className="flex h-full flex-col bg-isel-panel text-[13px] text-zinc-200 select-none">
      <div className="flex items-center justify-between gap-4 px-4 pt-3">
        <div className="flex items-center gap-4">
          <span className="text-[12px] font-semibold tracking-widest text-zinc-300">PATH PREVIEW</span>
          <span className="flex items-center gap-3 text-[11px] text-zinc-400">
            {legend("#00C0E0", "CUT")}
            {legend("#3FD0B0", "ARC")}
            {legend("#F0A020", "RAPID")}
          </span>
        </div>
        <span className="font-mono text-[11px] text-zinc-500">
          {stats &&
            `${nf.format(stats.counts.cut)} cut  ${nf.format(stats.counts.arc)} arc  ${nf.format(stats.counts.rapid)} rapid   W:${(b!.max_x - b!.min_x).toFixed(1)}  H:${(b!.max_y - b!.min_y).toFixed(1)} mm`}
        </span>
      </div>
      <div className="flex items-center gap-3 px-4 py-2">
        <Button size="sm" variant="secondary" onClick={fit} className="border-white/[0.08] bg-white/[0.05] text-zinc-100 transition-all duration-150 hover:bg-white/[0.09]">
          Fit
        </Button>
        <div className="flex items-center gap-2">
          <Checkbox
            id="iso"
            checked={iso}
            onCheckedChange={(c) => {
              setIso(c === true); // the camera survives the toggle, like the egui version
            }}
          />
          <Label htmlFor="iso" className="font-normal">
            Isometric
          </Label>
        </div>
      </div>
      <div ref={wrapRef} className="min-h-0 flex-1 px-4 pb-4">
        <canvas
          ref={canvasRef}
          onMouseDown={onMouseDown}
          className="h-full w-full cursor-crosshair rounded-xl border border-white/[0.08]"
        />
      </div>
    </div>
  );
}
