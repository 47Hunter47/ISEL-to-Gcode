import { useCallback, useEffect, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { api } from "./lib/api";
import type { Camera, GcodeSummary, PreviewScene } from "./lib/types";

// Canvas paints exactly what preview.rs projects, redrawn under a single
// client-side translate+scale: screen' = cam.ox + (cam.scale/base.scale) *
// (screen - base.ox), which is project()'s own world->screen semantics seen
// from the base camera the server fitted. So the 3-line anchor zoom (mirror
// of preview.rs's zoom_at, pinned by preview.rs's test) and panning stay
// client-side and never re-project points or cross IPC; the server scene is
// only re-requested for a new file, the ISO toggle, a resize refit, or Fit.
// The label color mirrors preview.rs's python-side TEXT_FAINT (#2F5F73).
const LABEL_COLOR = "#2F5F73";
const nf = new Intl.NumberFormat("en-US");

const pathFromUrl = () => new URLSearchParams(window.location.search).get("p") ?? "";

export default function PreviewPage() {
  const [path, setPath] = useState(pathFromUrl);
  const [iso, setIso] = useState(false);
  const [scene, setScene] = useState<PreviewScene | null>(null);
  const [stats, setStats] = useState<GcodeSummary | null>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const camRef = useRef<Camera | null>(null); // null asks the server to fit
  // Latest-request-wins queue: every view change bumps want.rev with the state
  // to request. One IPC flight at a time; after each response the flight
  // re-reads want, so a newer path/iso/camera always runs — a request is never
  // dropped and never replayed from a stale closure. A response may adopt its
  // server camera only when its request asked for a fit (cam null); otherwise
  // the camera on screen is the user's and an obsolete response must not roll
  // back a pan/zoom that happened mid-flight (preview.rs echoes sent cameras).
  const want = useRef({ path: "", iso: false, cam: null as Camera | null, rev: 0 });
  const applied = useRef(0);
  const running = useRef(false);
  const drag = useRef<{ x: number; y: number } | null>(null);

  const request = useCallback((patch?: { path?: string; iso?: boolean }) => {
    if (patch) Object.assign(want.current, patch);
    want.current.cam = camRef.current;
    want.current.rev += 1;
    if (running.current) return; // the live flight will re-read want when it lands
    running.current = true;
    void (async () => {
      try {
        while (applied.current !== want.current.rev) {
          const canvas = canvasRef.current;
          if (!canvas) break;
          const w = { ...want.current }; // snapshot: this flight is about w
          const { width, height } = canvas.getBoundingClientRect();
          const s = await api.previewScene({
            path: w.path,
            width: Math.max(1, Math.round(width)),
            height: Math.max(1, Math.round(height)),
            iso: w.iso,
            camera: w.cam,
          });
          applied.current = w.rev;
          if (w.cam === null) camRef.current = s.camera;
          setScene(s);
        }
      } catch (e) {
        console.error("preview_view failed:", e);
      } finally {
        running.current = false;
      }
    })();
  }, []);

  // paint everything the Rust scene carries; only the camera transform is
  // applied here (see the file header — same semantics as project())
  const rafRef = useRef(0);
  const paint = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas || !scene) return;
    const dpr = window.devicePixelRatio || 1;
    const { width, height } = canvas.getBoundingClientRect();
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(height * dpr);
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    const base = scene.camera;
    const cam = camRef.current ?? base;
    const a = cam.scale / base.scale;
    const e = cam.ox - a * base.ox;
    const f = cam.oy - a * base.oy;

    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = scene.background;
    ctx.fillRect(0, 0, width, height);

    ctx.setTransform(dpr * a, 0, 0, dpr * a, dpr * e, dpr * f);
    ctx.lineWidth = 1 / a; // strokes keep their 1px visual weight under scale
    for (const g of scene.grid) {
      ctx.strokeStyle = g.color;
      ctx.beginPath();
      ctx.moveTo(g.x1, g.y1);
      ctx.lineTo(g.x2, g.y2);
      ctx.stroke();
    }

    // labels follow their grid line but stay 7px: the text is a world
    // coordinate, so only its position belongs to the transform
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = LABEL_COLOR;
    ctx.font = "7px monospace";
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    for (const l of scene.labels) ctx.fillText(l.text, a * l.x + e, a * l.y + f);

    ctx.setTransform(dpr * a, 0, 0, dpr * a, dpr * e, dpr * f);
    for (const line of scene.lines) {
      ctx.strokeStyle = line.color;
      ctx.beginPath();
      line.points.forEach(([x, y], i) => (i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y)));
      ctx.stroke();
    }
  }, [scene]);

  const schedulePaint = useCallback(() => {
    // one repaint per frame, however many wheel/move events arrive
    if (rafRef.current) return;
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0;
      paint();
    });
  }, [paint]);

  useEffect(() => {
    paint();
  }, [paint]);

  useEffect(() => () => cancelAnimationFrame(rafRef.current), []);

  const statsReq = useRef(0);
  useEffect(() => {
    const my = ++statsReq.current; // ignore stats from an obsolete file
    api.previewStats(path)
      .then((s) => {
        if (statsReq.current === my) setStats(s);
      })
      .catch(() => {
        if (statsReq.current === my) setStats(null);
      });
  }, [path]);

  useEffect(() => {
    camRef.current = null; // new file or view change: refit
    request({ path, iso });
  }, [path, iso, request]);

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
      schedulePaint(); // client-side only: the canvas transform does the zoom
    };
    canvas.addEventListener("wheel", onWheel, { passive: false });
    return () => canvas.removeEventListener("wheel", onWheel);
  }, [schedulePaint]);

  const onMouseDown = (e: React.MouseEvent) => {
    drag.current = { x: e.clientX, y: e.clientY };
    const onMove = (ev: MouseEvent) => {
      const cam = camRef.current;
      const d = drag.current;
      if (!cam || !d) return;
      camRef.current = { ...cam, ox: cam.ox + (ev.clientX - d.x), oy: cam.oy + (ev.clientY - d.y) };
      drag.current = { x: ev.clientX, y: ev.clientY };
      schedulePaint(); // client-side only: dragging is a canvas translation
    };
    const onUp = () => {
      drag.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

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
