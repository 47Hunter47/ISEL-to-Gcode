try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.96"

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import os
import math
import threading

try:
    import numpy as np
    _NUMPY = True
except ImportError:
    _NUMPY = False

# ── machine constants ─────────────────────────────────────────────────────────
SCALE          = 1000.0
VEL_RATIO      = 16.6667
SAFE_Z         = 4.0
ARC_RESOLUTION = 0.05
DEFAULT_FEED   = 1000.0
RAPID_FEED     = 3000.0

# ── arc fitting constants (G1 → G2/G3 detection) ─────────────────────────────
# Minimum diameter for a full-circle arc to be converted to G2/G3
ARC_FIT_MIN_DIAMETER  = 6.0     # mm — circles smaller than this stay as G1
# How much each point is allowed to deviate from the fitted circle center
ARC_FIT_RADIUS_TOL    = 0.05    # mm — tighter = stricter circle detection
# A full circle must span at least this many degrees (360 = only full circles)
ARC_FIT_MIN_ARC_DEG   = 355.0   # degrees — allows tiny floating-point gap at start/end
# Minimum number of G1 segments needed to consider arc fitting
ARC_FIT_MIN_POINTS    = 8       # fewer segments = likely not a circle

# ── UI colors ─────────────────────────────────────────────────────────────────
BG  = "#1e1e1e"
FG  = "#ffffff"
BTN = "#2d2d2d"

# ── preview colors ────────────────────────────────────────────────────────────
COLOR_CUT   = "#00c8ff"
COLOR_RAPID = "#ff4444"
COLOR_BG    = "#121212"
COLOR_GRID  = "#2a2a2a"
COLOR_AXIS  = "#3a3a3a"


# ─────────────────────────────────────────────────────────────────────────────
#  ARC FITTING — detect full circles in a sequence of G1 points
# ─────────────────────────────────────────────────────────────────────────────

def fit_circle_to_points(pts):
    """
    Fit a circle to a list of (x, y) points using the algebraic method.
    Returns (cx, cy, radius) or None if fitting fails.

    This is a numerically stable least-squares approach — does not need numpy.
    Works by solving the linear system derived from x²+y²+Dx+Ey+F=0.
    """
    n = len(pts)
    if n < 3:
        return None

    # Build the matrix and vector for the least-squares system
    sum_x  = sum(p[0] for p in pts)
    sum_y  = sum(p[1] for p in pts)
    sum_x2 = sum(p[0]**2 for p in pts)
    sum_y2 = sum(p[1]**2 for p in pts)
    sum_xy = sum(p[0]*p[1] for p in pts)
    sum_x3 = sum(p[0]**3 for p in pts)
    sum_y3 = sum(p[1]**3 for p in pts)
    sum_xy2 = sum(p[0]*p[1]**2 for p in pts)
    sum_x2y = sum(p[0]**2*p[1] for p in pts)

    # Solve 3x3 linear system using Cramer's rule
    # [ sum_x2  sum_xy  sum_x ] [D]   [ -(sum_x3 + sum_xy2) ]
    # [ sum_xy  sum_y2  sum_y ] [E] = [ -(sum_x2y + sum_y3) ]
    # [ sum_x   sum_y   n     ] [F]   [ -(sum_x2 + sum_y2)  ]

    a11, a12, a13 = sum_x2, sum_xy, sum_x
    a21, a22, a23 = sum_xy, sum_y2, sum_y
    a31, a32, a33 = sum_x,  sum_y,  float(n)

    b1 = -(sum_x3 + sum_xy2)
    b2 = -(sum_x2y + sum_y3)
    b3 = -(sum_x2 + sum_y2)

    det = (a11*(a22*a33 - a23*a32)
           - a12*(a21*a33 - a23*a31)
           + a13*(a21*a32 - a22*a31))

    if abs(det) < 1e-12:
        return None  # degenerate / collinear points

    D = ((b1*(a22*a33 - a23*a32)
          - a12*(b2*a33 - a23*b3)
          + a13*(b2*a32 - a22*b3)) / det)

    E = ((a11*(b2*a33 - a23*b3)
          - b1*(a21*a33 - a23*a31)
          + a13*(a21*b3 - b2*a31)) / det)

    F = ((a11*(a22*b3 - b2*a32)
          - a12*(a21*b3 - b2*a31)
          + b1*(a21*a32 - a22*a31)) / det)

    cx = -D / 2.0
    cy = -E / 2.0
    r2 = cx**2 + cy**2 - F
    if r2 <= 0:
        return None

    return cx, cy, math.sqrt(r2)


def try_arc_fit(buffer_pts, start_pos):
    """
    Try to detect a full circle in buffer_pts (list of (x, y) positions).
    start_pos is the position before the first buffer point.

    Returns (cx, cy, radius, cw) if a valid full circle is found,
    or None if the points do not form a full circle.

    Rules enforced here:
    - Minimum ARC_FIT_MIN_POINTS segments
    - All points within ARC_FIT_RADIUS_TOL of the fitted center
    - Arc span must be >= ARC_FIT_MIN_ARC_DEG (full circle check)
    - Diameter must be >= ARC_FIT_MIN_DIAMETER
    - Z must not change (no helical moves)
    """
    if len(buffer_pts) < ARC_FIT_MIN_POINTS:
        return None

    # Check Z is constant — no helical arcs
    z_vals = [p[2] for p in buffer_pts]
    if max(z_vals) - min(z_vals) > 1e-4:
        return None

    xy_pts = [(p[0], p[1]) for p in buffer_pts]

    # Include the start position in the fit for better accuracy
    all_pts = [(start_pos["X"], start_pos["Y"])] + xy_pts

    result = fit_circle_to_points(all_pts)
    if result is None:
        return None

    cx, cy, radius = result

    # Check minimum diameter
    if radius * 2.0 < ARC_FIT_MIN_DIAMETER:
        return None

    # Check that every point lies on the circle within tolerance
    for x, y in all_pts:
        dist = math.hypot(x - cx, y - cy)
        if abs(dist - radius) > ARC_FIT_RADIUS_TOL:
            return None

    # Compute total arc span by summing angular steps between consecutive points
    all_xy = [(start_pos["X"], start_pos["Y"])] + xy_pts
    angles = [math.atan2(p[1] - cy, p[0] - cx) for p in all_xy]

    # Determine dominant direction (CW or CCW) from cross products
    cw_votes = 0
    ccw_votes = 0
    for i in range(1, len(all_xy) - 1):
        ax = all_xy[i][0] - all_xy[i-1][0]
        ay = all_xy[i][1] - all_xy[i-1][1]
        bx = all_xy[i+1][0] - all_xy[i][0]
        by = all_xy[i+1][1] - all_xy[i][1]
        cross = ax * by - ay * bx
        if cross < 0:
            cw_votes += 1
        else:
            ccw_votes += 1

    cw = cw_votes > ccw_votes

    # Sum angular steps in the dominant direction
    total_angle = 0.0
    for i in range(1, len(angles)):
        da = angles[i] - angles[i-1]
        # Normalize to [-pi, pi]
        while da > math.pi:
            da -= 2 * math.pi
        while da < -math.pi:
            da += 2 * math.pi
        # CW = negative da; CCW = positive da
        if cw:
            if da > 0:
                da -= 2 * math.pi
        else:
            if da < 0:
                da += 2 * math.pi
        total_angle += abs(da)

    total_deg = math.degrees(total_angle)

    if total_deg < ARC_FIT_MIN_ARC_DEG:
        return None  # not a full circle — could be half-circle or corner radius

    return cx, cy, radius, cw


# ─────────────────────────────────────────────────────────────────────────────
#  PREVIEW WINDOW
# ─────────────────────────────────────────────────────────────────────────────

def open_preview(parent, gcode_path):
    """
    Parse output G-code and draw toolpath on a canvas.
    - Segments grouped into polylines per move type for speed
    - Toggle between 2D top-down and isometric projection
    - Pan: left-click drag  |  Zoom: mouse wheel  |  Fit: button
    """

    # ── parse G-code ──────────────────────────────────────────────────────────
    points = []   # (x, y, z, is_rapid)
    cx, cy, cz = 0.0, 0.0, 0.0
    is_rapid   = False

    try:
        with open(gcode_path, "r", encoding="utf-8", errors="replace") as f:
            for raw in f:
                ln = re.sub(r"^N\d+\s*", "", raw.strip())
                if not ln:
                    continue
                if ln.startswith("G0"):
                    is_rapid = True
                elif re.match(r"G[123]\b", ln):
                    is_rapid = False
                else:
                    continue
                xm = re.search(r"X(-?\d+(?:\.\d+)?)", ln)
                ym = re.search(r"Y(-?\d+(?:\.\d+)?)", ln)
                zm = re.search(r"Z(-?\d+(?:\.\d+)?)", ln)
                nx = float(xm.group(1)) if xm else cx
                ny = float(ym.group(1)) if ym else cy
                nz = float(zm.group(1)) if zm else cz
                points.append((cx, cy, cz, is_rapid))
                cx, cy, cz = nx, ny, nz
        if points:
            points.append((cx, cy, cz, is_rapid))
    except Exception as e:
        messagebox.showerror("Preview Error", f"Could not read G-code:\n{e}",
                             parent=parent)
        return

    if not points:
        messagebox.showinfo("Preview", "No moves found in G-code file.",
                            parent=parent)
        return

    # ── group consecutive same-type segments into polylines ───────────────────
    def build_groups(pts):
        groups = []
        cur    = None
        for x, y, z, rapid in pts:
            if cur is None or cur["rapid"] != rapid:
                cur = {"rapid": rapid, "pts": []}
                groups.append(cur)
            cur["pts"].append((x, y, z))
        return groups

    raw_groups = build_groups(points)

    all_x = [p[0] for p in points]
    all_y = [p[1] for p in points]
    min_x, max_x = min(all_x), max(all_x)
    min_y, max_y = min(all_y), max(all_y)
    span_x = max_x - min_x or 1.0
    span_y = max_y - min_y or 1.0

    cut_count   = sum(1 for p in points if not p[3])
    rapid_count = sum(1 for p in points if p[3])

    # ── window ────────────────────────────────────────────────────────────────
    win = tk.Toplevel(parent)
    win.title(f"Path Preview — {os.path.basename(gcode_path)}")
    win.geometry("960x720")
    win.configure(bg=BG)
    win.resizable(True, True)

    # ── toolbar ───────────────────────────────────────────────────────────────
    toolbar = tk.Frame(win, bg=BTN, pady=5)
    toolbar.pack(fill=tk.X)

    tk.Label(toolbar, text="Path Preview", bg=BTN, fg=FG,
             font=("Consolas", 10, "bold")).pack(side=tk.LEFT, padx=10)

    tk.Label(toolbar, text="━━", bg=BTN, fg=COLOR_CUT,
             font=("Consolas", 11)).pack(side=tk.LEFT, padx=(16, 2))
    tk.Label(toolbar, text="Cut", bg=BTN, fg=FG,
             font=("Consolas", 9)).pack(side=tk.LEFT)

    tk.Label(toolbar, text="━━", bg=BTN, fg=COLOR_RAPID,
             font=("Consolas", 11)).pack(side=tk.LEFT, padx=(12, 2))
    tk.Label(toolbar, text="Rapid", bg=BTN, fg=FG,
             font=("Consolas", 9)).pack(side=tk.LEFT)

    stats = (f"   {cut_count:,} cut  {rapid_count:,} rapid"
             f"   W:{span_x:.1f}  H:{span_y:.1f} mm")
    tk.Label(toolbar, text=stats, bg=BTN, fg="#777777",
             font=("Consolas", 9)).pack(side=tk.LEFT, padx=6)

    iso_var = tk.BooleanVar(value=False)

    def toggle_iso():
        fit_view()

    iso_chk = tk.Checkbutton(
        toolbar, text="Isometric", variable=iso_var, command=toggle_iso,
        bg=BTN, fg=FG, selectcolor="#3a7afe",
        activebackground=BTN, activeforeground=FG, font=("Consolas", 9)
    )
    iso_chk.pack(side=tk.RIGHT, padx=6)

    tk.Button(toolbar, text="Fit", bg="#3a7afe", fg="white",
              font=("Consolas", 9), relief="flat", padx=8,
              command=lambda: fit_view()).pack(side=tk.RIGHT, padx=6)

    # ── canvas ────────────────────────────────────────────────────────────────
    canvas = tk.Canvas(win, bg=COLOR_BG, highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)

    # ── view state ────────────────────────────────────────────────────────────
    view = {"ox": 0.0, "oy": 0.0, "scale": 1.0}
    drag = {"x": 0, "y": 0, "active": False}
    PADDING   = 48
    ISO_ANGLE = math.radians(30)
    ISO_CX    = math.cos(ISO_ANGLE)
    ISO_CY    = math.sin(ISO_ANGLE)

    def project(wx, wy, wz):
        s = view["scale"]
        if iso_var.get():
            px = (wx - wy) * ISO_CX
            py = (wx + wy) * ISO_CY - wz
            return view["ox"] + px * s, view["oy"] + py * s
        else:
            return view["ox"] + wx * s, view["oy"] - wy * s

    def projected_bounds():
        if iso_var.get():
            pxs = [(p[0] - p[1]) * ISO_CX for p in points]
            pys = [(p[0] + p[1]) * ISO_CY - p[2] for p in points]
        else:
            pxs = [p[0] for p in points]
            pys = [-p[1] for p in points]
        return min(pxs), max(pxs), min(pys), max(pys)

    def draw_all():
        canvas.delete("all")
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 2 or ch < 2:
            return

        # grid (2D only)
        if not iso_var.get():
            s = view["scale"]
            raw_step  = 50.0 / s
            magnitude = 10 ** math.floor(math.log10(max(raw_step, 1e-9)))
            grid_step = magnitude
            for mult in (1, 2, 5, 10):
                if magnitude * mult >= raw_step:
                    grid_step = magnitude * mult
                    break
            gx = math.floor(min_x / grid_step) * grid_step
            while gx <= max_x + grid_step:
                px, _ = project(gx, 0, 0)
                if 0 <= px <= cw:
                    canvas.create_line(px, 0, px, ch, fill=COLOR_GRID, width=1)
                    canvas.create_text(px + 2, ch - 12, text=f"{gx:.0f}",
                                       fill="#4a4a4a", font=("Consolas", 7),
                                       anchor="w")
                gx += grid_step
            gy = math.floor(min_y / grid_step) * grid_step
            while gy <= max_y + grid_step:
                _, py = project(0, gy, 0)
                if 0 <= py <= ch:
                    canvas.create_line(0, py, cw, py, fill=COLOR_GRID, width=1)
                    canvas.create_text(4, py - 2, text=f"{gy:.0f}",
                                       fill="#4a4a4a", font=("Consolas", 7),
                                       anchor="sw")
                gy += grid_step
            ox_px, oy_px = project(0, 0, 0)
            canvas.create_line(ox_px, 0, ox_px, ch, fill=COLOR_AXIS, width=1)
            canvas.create_line(0, oy_px, cw, oy_px, fill=COLOR_AXIS, width=1)

        # draw: rapid behind, cut on top
        for pass_rapid in (True, False):
            color = COLOR_RAPID if pass_rapid else COLOR_CUT
            dash  = (4, 4) if pass_rapid else None
            for grp in raw_groups:
                if grp["rapid"] != pass_rapid:
                    continue
                pts = grp["pts"]
                if len(pts) < 2:
                    continue
                flat = []
                for wx, wy, wz in pts:
                    px, py = project(wx, wy, wz)
                    flat.extend((px, py))
                xs = flat[0::2]
                ys = flat[1::2]
                if (min(xs) > cw or max(xs) < 0 or
                        min(ys) > ch or max(ys) < 0):
                    continue
                if dash:
                    canvas.create_line(*flat, fill=color, width=1, dash=dash)
                else:
                    canvas.create_line(*flat, fill=color, width=1)

    def fit_view(event=None):
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 2 or ch < 2:
            win.after(60, fit_view)
            return
        px_min, px_max, py_min, py_max = projected_bounds()
        span_px = px_max - px_min or 1.0
        span_py = py_max - py_min or 1.0
        sx = (cw - 2 * PADDING) / span_px
        sy = (ch - 2 * PADDING) / span_py
        s  = min(sx, sy)
        view["scale"] = s
        view["ox"]    = (cw - span_px * s) / 2 - px_min * s
        view["oy"]    = (ch - span_py * s) / 2 - py_min * s
        draw_all()

    def on_press(e):
        drag["x"] = e.x; drag["y"] = e.y; drag["active"] = True
        canvas.config(cursor="fleur")

    def on_release(e):
        drag["active"] = False
        canvas.config(cursor="")

    def on_drag(e):
        if not drag["active"]: return
        view["ox"] += e.x - drag["x"]
        view["oy"] += e.y - drag["y"]
        drag["x"] = e.x; drag["y"] = e.y
        draw_all()

    def on_zoom(e):
        factor = 1.15 if e.delta > 0 else (1 / 1.15)
        view["ox"] = e.x + (view["ox"] - e.x) * factor
        view["oy"] = e.y + (view["oy"] - e.y) * factor
        view["scale"] *= factor
        draw_all()

    canvas.bind("<ButtonPress-1>",   on_press)
    canvas.bind("<ButtonRelease-1>", on_release)
    canvas.bind("<B1-Motion>",       on_drag)
    canvas.bind("<MouseWheel>",      on_zoom)
    canvas.bind("<Button-4>",
                lambda e: on_zoom(type("E",(),{"delta":1,"x":e.x,"y":e.y})()))
    canvas.bind("<Button-5>",
                lambda e: on_zoom(type("E",(),{"delta":-1,"x":e.x,"y":e.y})()))
    canvas.bind("<Configure>",       lambda e: fit_view())

    win.after(100, fit_view)


# ─────────────────────────────────────────────────────────────────────────────
#  CORE CONVERSION
# ─────────────────────────────────────────────────────────────────────────────

def convert_file(input_path, output_path, log_fn,
                 progress_callback=None, use_arc_commands=False):

    total_time_min  = 0.0
    current_feed    = None
    last_pos        = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    line_no         = 1

    # Counters for arc fitting results (logged at end)
    arc_fit_converted = 0
    arc_fit_skipped   = 0

    def nline():
        nonlocal line_no
        s = f"N{line_no:05d} "
        line_no += 1
        return s

    def parse_coord(text, axes=("X", "Y", "Z")):
        coords = {}
        for axis in axes:
            m = re.search(rf"{axis}(-?\d+(?:\.\d+)?)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    def arc_to_g2g3(start, end, center, cw):
        """Convert a known arc (from CWABS/CCWABS) to a single G2/G3 R-format line."""
        nonlocal total_time_min, current_feed
        cx, cy = center["X"], center["Y"]
        r  = math.hypot(start["X"] - cx, start["Y"] - cy)
        a0 = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1 = math.atan2(end["Y"]   - cy, end["X"]   - cx)
        if cw:
            if a1 >= a0: a1 -= 2 * math.pi
        else:
            if a1 <= a0: a1 += 2 * math.pi
        span    = abs(a1 - a0)
        arc_len = span * r
        if current_feed:
            total_time_min += arc_len / current_feed
        r_signed = r if span <= math.pi else -r
        code  = "G2" if cw else "G3"
        gline = nline() + f"{code} X{end['X']:.4f} Y{end['Y']:.4f} R{r_signed:.4f}"
        return end.copy(), [gline]

    def fitted_arc_to_g2g3(start_pos, end_pos, cx, cy, radius, cw):
        """
        Emit a single G2/G3 line for a full-circle detected by arc fitting.
        For a full circle, end point = start point (G2/G3 full circle syntax).
        R is negative because span > 180° (full circle = 360°).
        """
        nonlocal total_time_min, current_feed
        arc_len = 2.0 * math.pi * radius
        if current_feed:
            total_time_min += arc_len / current_feed
        # Full circle: R must be negative (span > 180°)
        r_signed = -radius
        code  = "G2" if cw else "G3"
        # End = start for full circle
        gline = nline() + (f"{code} X{start_pos['X']:.4f} Y{start_pos['Y']:.4f}"
                           f" R{r_signed:.4f}")
        return end_pos.copy(), [gline]

    def emit_g1_buffer(buffer_pts, buf_start_pos):
        """
        Flush buffered MOVEABS points as regular G1 lines.
        Called when arc fitting fails or is not attempted.
        """
        nonlocal total_time_min, current_feed
        lines = []
        prev  = buf_start_pos
        for pt in buffer_pts:
            cmd = "G1"
            # Only emit axes that changed to keep output clean
            if abs(pt["X"] - prev.get("X", 0)) > 1e-6:
                cmd += f" X{pt['X']:.3f}"
            if abs(pt["Y"] - prev.get("Y", 0)) > 1e-6:
                cmd += f" Y{pt['Y']:.3f}"
            if abs(pt["Z"] - prev.get("Z", 0)) > 1e-6:
                cmd += f" Z{pt['Z']:.3f}"
            if cmd == "G1":
                # No axes changed — skip duplicate point
                prev = pt
                continue
            lines.append(nline() + cmd)
            d = math.sqrt(sum((pt[k] - prev[k]) ** 2 for k in "XYZ"))
            if current_feed and d > 0:
                total_time_min += d / current_feed
            prev = pt
        return lines

    def linearize_arc_numpy(start, end, center, cw):
        nonlocal total_time_min, current_feed, line_no
        cx, cy = center["X"], center["Y"]
        z      = start["Z"]
        r      = math.hypot(start["X"] - cx, start["Y"] - cy)
        a0     = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1     = math.atan2(end["Y"]   - cy, end["X"]   - cx)
        if cw:
            if a1 >= a0: a1 -= 2 * math.pi
        else:
            if a1 <= a0: a1 += 2 * math.pi
        arc_len = abs(a1 - a0) * r
        steps   = max(1, int(arc_len / ARC_RESOLUTION))
        angles  = np.linspace(a0, a1, steps + 1)[1:]
        xs      = cx + r * np.cos(angles)
        ys      = cy + r * np.sin(angles)
        dx = np.diff(np.concatenate([[start["X"]], xs]))
        dy = np.diff(np.concatenate([[start["Y"]], ys]))
        if current_feed:
            total_time_min += float(np.sum(np.hypot(dx, dy))) / current_feed
        lines = [
            f"N{line_no + i:05d} G1 X{x:.3f} Y{y:.3f} Z{z:.3f}"
            for i, (x, y) in enumerate(zip(xs, ys))
        ]
        line_no += steps
        pos = {"X": float(xs[-1]), "Y": float(ys[-1]), "Z": z}
        fp  = end.copy()
        if abs(fp["X"] - pos["X"]) > 0.001 or abs(fp["Y"] - pos["Y"]) > 0.001:
            dist = math.hypot(fp["X"] - pos["X"], fp["Y"] - pos["Y"])
            lines.append(nline() + f"G1 X{fp['X']:.3f} Y{fp['Y']:.3f} Z{fp['Z']:.3f}")
            if current_feed:
                total_time_min += dist / current_feed
            pos = fp
        return pos, lines

    def linearize_arc_pure(start, end, center, cw):
        nonlocal total_time_min, current_feed, line_no
        cx, cy = center["X"], center["Y"]
        z      = start["Z"]
        r      = math.hypot(start["X"] - cx, start["Y"] - cy)
        a0     = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1     = math.atan2(end["Y"]   - cy, end["X"]   - cx)
        if cw:
            if a1 >= a0: a1 -= 2 * math.pi
        else:
            if a1 <= a0: a1 += 2 * math.pi
        arc_len = abs(a1 - a0) * r
        steps   = max(1, int(arc_len / ARC_RESOLUTION))
        da      = (a1 - a0) / steps
        lines   = []
        px, py  = start["X"], start["Y"]
        for i in range(1, steps + 1):
            a  = a0 + da * i
            tx = cx + r * math.cos(a)
            ty = cy + r * math.sin(a)
            lines.append(f"N{line_no:05d} G1 X{tx:.3f} Y{ty:.3f} Z{z:.3f}")
            line_no += 1
            if current_feed:
                total_time_min += math.hypot(tx - px, ty - py) / current_feed
            px, py = tx, ty
        pos = {"X": px, "Y": py, "Z": z}
        fp  = end.copy()
        if abs(fp["X"] - pos["X"]) > 0.001 or abs(fp["Y"] - pos["Y"]) > 0.001:
            lines.append(nline() + f"G1 X{fp['X']:.3f} Y{fp['Y']:.3f} Z{fp['Z']:.3f}")
            if current_feed:
                total_time_min += math.hypot(fp["X"] - pos["X"],
                                             fp["Y"] - pos["Y"]) / current_feed
            pos = fp
        return pos, lines

    # Select arc handler for CWABS/CCWABS lines
    if use_arc_commands:
        handle_arc = arc_to_g2g3
    elif _NUMPY:
        handle_arc = linearize_arc_numpy
    else:
        handle_arc = linearize_arc_pure

    # ── count lines ───────────────────────────────────────────────────────────
    try:
        with open(input_path, "r", encoding="utf-8", errors="replace") as f:
            total_lines = sum(
                1 for ln in f
                if ln.strip() and not ln.strip().startswith(";")
            )
    except FileNotFoundError:
        raise FileNotFoundError(f"Input file not found: {input_path}")
    except Exception as e:
        raise Exception(f"Error reading input file: {e}")

    if total_lines == 0:
        raise Exception("Input file contains no processable commands.")

    if progress_callback:
        progress_callback(0, "Reading file…")

    try:
        out_f = open(output_path, "w", encoding="utf-8", buffering=256 * 1024)
    except Exception as e:
        raise Exception(f"Cannot open output file for writing: {e}")

    def emit(line):
        out_f.write(line + "\n")

    # ── MOVEABS arc-fitting buffer ─────────────────────────────────────────────
    # When use_arc_commands=True, consecutive MOVEABS lines are buffered here
    # instead of being written immediately. When a non-MOVEABS command arrives
    # (or at end of file), the buffer is flushed: either as a single G2/G3 if
    # a full circle is detected, or as normal G1 lines otherwise.
    moveabs_buffer   = []      # list of {"X", "Y", "Z"} dicts
    buffer_start_pos = None    # last_pos snapshot before buffer started

    def flush_moveabs_buffer():
        """
        Try arc fitting on the accumulated buffer.
        If a full circle is found → emit G2/G3.
        Otherwise → emit G1 lines.
        Returns (new_last_pos, lines_written_count)
        """
        nonlocal arc_fit_converted, arc_fit_skipped, last_pos

        if not moveabs_buffer:
            return

        if use_arc_commands and len(moveabs_buffer) >= ARC_FIT_MIN_POINTS:
            # Attempt arc fitting
            xy_pts = [(p["X"], p["Y"], p["Z"]) for p in moveabs_buffer]
            fit    = try_arc_fit(xy_pts, buffer_start_pos)
            if fit is not None:
                cx, cy, radius, cw = fit
                end_pos = moveabs_buffer[-1]
                new_pos, arc_lines = fitted_arc_to_g2g3(
                    buffer_start_pos, end_pos, cx, cy, radius, cw
                )
                for al in arc_lines:
                    emit(al)
                last_pos = new_pos
                arc_fit_converted += 1
                log_fn(f"  Arc fit: full circle ⌀{radius*2:.2f}mm → "
                       f"{'G2' if cw else 'G3'} R{-radius:.4f}")
                moveabs_buffer.clear()
                return

        # Arc fitting not attempted or failed → flush as G1
        g1_lines = emit_g1_buffer(moveabs_buffer, buffer_start_pos)
        for gl in g1_lines:
            emit(gl)
        if moveabs_buffer:
            last_pos = moveabs_buffer[-1]
        if use_arc_commands and len(moveabs_buffer) >= ARC_FIT_MIN_POINTS:
            arc_fit_skipped += 1
        moveabs_buffer.clear()

    try:
        emit(nline() + "G21")
        emit(nline() + "G17")
        emit(nline() + "G40")
        emit(nline() + "G49")
        emit(nline() + "G90")
        emit(nline() + "G94")

        with open(input_path, "r", encoding="utf-8", errors="replace") as in_f:
            lines_processed = 0
            for raw in in_f:
                line = raw.strip()
                if not line or line.startswith(";"):
                    continue
                lines_processed += 1
                if progress_callback and (
                    lines_processed % 100 == 0 or lines_processed == total_lines
                ):
                    pct = (lines_processed / total_lines) * 95
                    progress_callback(pct,
                        f"Processing line {lines_processed}/{total_lines}")

                try:
                    if line.startswith("SPINDLE CW"):
                        # Non-MOVEABS command → flush buffer first
                        flush_moveabs_buffer()
                        m = re.search(r"RPM(\d+)", line)
                        if m:
                            emit(nline() + f"S{int(m.group(1))} M03")
                            log_fn(f"Spindle: {m.group(1)} RPM")
                        else:
                            log_fn(f"Warning: Could not parse RPM from: {line}")

                    elif line.startswith("FASTABS"):
                        # Non-MOVEABS command → flush buffer first
                        flush_moveabs_buffer()
                        c      = parse_coord(line)
                        target = last_pos.copy()
                        target.update(c)
                        cmd = "G0"
                        for k in ("X", "Y", "Z"):
                            if k in c:
                                cmd += f" {k}{target[k]:.3f}"
                        emit(nline() + cmd)
                        d = math.sqrt(sum((target[k] - last_pos[k]) ** 2
                                         for k in "XYZ"))
                        if d > 0:
                            total_time_min += d / RAPID_FEED
                        last_pos = target

                    elif line.startswith("MOVEABS"):
                        c      = parse_coord(line)
                        target = last_pos.copy()
                        target.update(c)
                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            emit(nline() + f"F{current_feed:.0f}")
                            log_fn(f"Warning: No VEL found, using default F{current_feed:.0f}")

                        # Pure-Z moves (no XY) are never part of a circle —
                        # flush buffer first, then emit G1 immediately
                        has_xy = "X" in c or "Y" in c

                        if use_arc_commands and has_xy:
                            # XY move in arc mode — buffer for circle detection
                            if not moveabs_buffer:
                                # Snapshot position before buffer starts
                                buffer_start_pos = last_pos.copy()
                            moveabs_buffer.append(target.copy())
                            # Keep last_pos current so next line has correct reference
                            last_pos = target
                        else:
                            # Pure-Z move (or arc mode off):
                            # flush any pending XY buffer, then write G1 immediately
                            if use_arc_commands:
                                flush_moveabs_buffer()
                            cmd = "G1"
                            for k in ("X", "Y", "Z"):
                                if k in c:
                                    cmd += f" {k}{target[k]:.3f}"
                            emit(nline() + cmd)
                            if current_feed:
                                d = math.sqrt(sum((target[k] - last_pos[k]) ** 2
                                                 for k in "XYZ"))
                                total_time_min += d / current_feed
                            last_pos = target

                    elif line.startswith("VEL"):
                        # Non-MOVEABS command → flush buffer first
                        flush_moveabs_buffer()
                        m = re.search(r"VEL\s*(\d+(?:\.\d+)?)", line)
                        if m:
                            current_feed = float(m.group(1)) / VEL_RATIO
                            emit(nline() + f"F{current_feed:.0f}")
                            log_fn(f"Feed: F{current_feed:.0f}")
                        else:
                            log_fn(f"Warning: Could not parse VEL from: {line}")

                    elif line.startswith("CWABS") or line.startswith("CCWABS"):
                        # Non-MOVEABS command → flush buffer first
                        flush_moveabs_buffer()
                        cw = line.startswith("CWABS")
                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            emit(nline() + f"F{current_feed:.0f}")
                            log_fn(f"Warning: No VEL found, using default F{current_feed:.0f}")
                        c  = parse_coord(line, axes=("X", "Y", "Z"))
                        ij = parse_coord(line, axes=("I", "J"))
                        end = last_pos.copy()
                        end.update(c)
                        # I/J are absolute coordinates in ISEL (not offsets)
                        center = {
                            "X": ij.get("I", 0.0),
                            "Y": ij.get("J", 0.0),
                        }
                        r_check = math.hypot(last_pos["X"] - center["X"],
                                             last_pos["Y"] - center["Y"])
                        if r_check < 1e-6:
                            log_fn(f"Warning: zero-radius arc skipped on line [{line}]")
                            continue
                        new_pos, arc_lines = handle_arc(last_pos, end, center, cw)
                        for al in arc_lines:
                            emit(al)
                        last_pos = new_pos

                except Exception as e:
                    log_fn(f"Warning: Error on line [{line}]: {e}")
                    continue

            # End of file — flush any remaining MOVEABS buffer
            flush_moveabs_buffer()

        emit(nline() + "M05")
        emit(nline() + "M30")

    finally:
        out_f.close()

    if progress_callback:
        progress_callback(100, "Complete!")

    # Log arc fitting summary if arc mode was active
    if use_arc_commands and (arc_fit_converted > 0 or arc_fit_skipped > 0):
        log_fn(f"Arc fitting summary: {arc_fit_converted} circle(s) converted to G2/G3, "
               f"{arc_fit_skipped} segment group(s) left as G1")

    return total_time_min


# ─────────────────────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────────────────────

def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("520x520")
    root.configure(bg=BG)

    input_var    = tk.StringVar()
    use_arc_var  = tk.BooleanVar(value=False)
    converting   = False
    last_output  = {"path": None}   # path of last successfully converted file

    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def update_progress(value, status_text=""):
        progress_bar["value"] = value
        if status_text:
            progress_label.config(text=status_text)
        root.update_idletasks()

    def drop(event):
        if converting: return
        path = event.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            log(f"File imported: {path}")
        else:
            log(f"Error: Not a valid file: {path}")

    def browse_input():
        if converting: return
        path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if path:
            input_var.set(path)
            log(f"File selected: {path}")

    def show_preview():
        path = last_output["path"]
        if path and os.path.isfile(path):
            open_preview(root, path)

    def _do_convert(in_path, out_path, arc_mode):
        nonlocal converting

        def safe_progress(val, txt=""):
            root.after(0, update_progress, val, txt)

        def safe_log(msg):
            root.after(0, log, msg)

        try:
            total_time = convert_file(
                in_path, out_path, safe_log, safe_progress,
                use_arc_commands=arc_mode
            )
            m        = int(total_time)
            s        = int((total_time - m) * 60)
            mode_str = "G2/G3 arc commands" if arc_mode else (
                       "G1 + numpy" if _NUMPY else "G1 linearised")

            def on_done():
                last_output["path"] = out_path
                preview_btn.config(state="normal")   # enable preview button

                log("-" * 50)
                log(f"✓ Conversion completed  ({mode_str})")
                log(f"⏱ Estimated program time: {m} min {s} sec")
                log("⚠ Program time may change according to machine parameters")
                progress_label.config(text="Conversion complete!")
                messagebox.showinfo(
                    "Completed",
                    f"Conversion finished!\nMode: {mode_str}\n\n"
                    f"Estimated time: {m} min {s} sec"
                )

            root.after(0, on_done)

        except Exception as e:
            root.after(0, lambda: log(f"✗ Error: {e}"))
            root.after(0, lambda: progress_label.config(text="Error occurred"))
            root.after(0, lambda: messagebox.showerror("Error", str(e)))

        finally:
            def re_enable():
                nonlocal converting
                converting = False
                convert_btn.config(state="normal", text="Convert")
                browse_btn.config(state="normal")
            root.after(0, re_enable)

    def convert():
        nonlocal converting
        if converting: return
        if not input_var.get():
            messagebox.showerror("Error", "No input file selected")
            return
        in_path = input_var.get()
        if not os.path.isfile(in_path):
            messagebox.showerror("Error", f"File not found: {in_path}")
            return
        base     = os.path.splitext(os.path.basename(in_path))[0]
        out_path = filedialog.asksaveasfilename(
            defaultextension=".ngc",
            initialfile=base + ".ngc",
            filetypes=[("G-code Files", "*.ngc"), ("All Files", "*.*")]
        )
        if not out_path: return

        arc_mode   = use_arc_var.get()
        converting = True
        preview_btn.config(state="disabled")         # reset preview until done
        convert_btn.config(state="disabled", text="Converting...")
        browse_btn.config(state="disabled")
        progress_bar["value"] = 0
        progress_label.config(text="Starting conversion...")
        logbox.delete(1.0, tk.END)
        log("Starting conversion...")
        log(f"Input:   {in_path}")
        log(f"Output:  {out_path}")
        if arc_mode:
            log("Mode:    G2/G3 arc commands + full-circle detection (R format)")
        elif _NUMPY:
            log("Mode:    G1 linearised  (numpy accelerated)")
        else:
            log("Mode:    G1 linearised  (pure Python)")
        log("-" * 50)

        threading.Thread(
            target=_do_convert,
            args=(in_path, out_path, arc_mode),
            daemon=True
        ).start()

    # ── layout ────────────────────────────────────────────────────────────────
    tk.Label(root, text="ISEL File", bg=BG, fg=FG,
             font=("Arial", 10)).pack(pady=5)

    entry = tk.Entry(root, textvariable=input_var, width=62, bg=BTN, fg=FG)
    entry.pack(padx=10)
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind("<<Drop>>", drop)

    btn_frame = tk.Frame(root, bg=BG)
    btn_frame.pack(pady=10)

    browse_btn = tk.Button(btn_frame, text="Browse", command=browse_input,
                           bg=BTN, fg=FG, width=12)
    browse_btn.pack(side=tk.LEFT, padx=5)

    convert_btn = tk.Button(btn_frame, text="Convert", command=convert,
                            bg="#3a7afe", fg="white", width=12,
                            font=("Arial", 10, "bold"))
    convert_btn.pack(side=tk.LEFT, padx=5)

    # Preview button — disabled until a successful conversion exists
    preview_btn = tk.Button(btn_frame, text="Preview", command=show_preview,
                            bg="#2a6a2a", fg="white", width=12,
                            font=("Arial", 10, "bold"), state="disabled")
    preview_btn.pack(side=tk.LEFT, padx=5)

    numpy_text  = ("⚡ numpy detected – arc linearisation accelerated"
                   if _NUMPY else "⚠ numpy not found – pip install numpy")
    numpy_color = "#00ff88" if _NUMPY else "#ffaa00"
    tk.Label(root, text=numpy_text, bg=BG, fg=numpy_color,
             font=("Arial", 8)).pack(pady=(0, 2))

    arc_frame = tk.Frame(root, bg=BG)
    arc_frame.pack(pady=(2, 8))
    tk.Checkbutton(
        arc_frame,
        text="Use G2/G3 arc commands  (smaller file, requires arc-capable controller)",
        variable=use_arc_var, bg=BG, fg=FG, selectcolor=BTN,
        activebackground=BG, activeforeground=FG, font=("Arial", 9),
    ).pack()

    progress_frame = tk.Frame(root, bg=BG)
    progress_frame.pack(pady=5, padx=10, fill=tk.X)

    progress_label = tk.Label(progress_frame, text="Ready", bg=BG, fg=FG,
                               font=("Arial", 9))
    progress_label.pack()

    style = ttk.Style()
    style.theme_use("default")
    style.configure("custom.Horizontal.TProgressbar",
                    troughcolor=BTN, background="#3a7afe", thickness=20)
    progress_bar = ttk.Progressbar(
        progress_frame, style="custom.Horizontal.TProgressbar",
        orient="horizontal", mode="determinate", length=490
    )
    progress_bar.pack(pady=5)

    tk.Label(root, text="Conversion Log", bg=BG, fg=FG,
             font=("Arial", 9)).pack(pady=(5, 2))

    logbox = tk.Text(root, height=10, bg="#121212", fg="#00ff88",
                     font=("Consolas", 9))
    logbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    log("ISEL to G-code Converter Ready")
    log("Drag and drop a file or click Browse to start")
    log("-" * 50)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
