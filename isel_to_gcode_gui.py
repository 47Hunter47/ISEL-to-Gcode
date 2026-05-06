try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.93"

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

# ── UI colors ─────────────────────────────────────────────────────────────────
BG  = "#1e1e1e"
FG  = "#ffffff"
BTN = "#2d2d2d"

# ── preview colors ────────────────────────────────────────────────────────────
COLOR_CUT   = "#00c8ff"   # G1 cutting moves  — cyan
COLOR_RAPID = "#ff4444"   # G0 rapid moves    — red
COLOR_BG    = "#121212"   # canvas background
COLOR_GRID  = "#2a2a2a"   # grid lines
COLOR_AXIS  = "#3a3a3a"   # axis lines


# ─────────────────────────────────────────────────────────────────────────────
#  PREVIEW WINDOW
# ─────────────────────────────────────────────────────────────────────────────

def open_preview(parent, gcode_path):
    """
    Parse the output G-code file and draw all moves on a canvas.
    G0  → rapid (red, thin dashed)
    G1/G2/G3 → cutting (cyan, solid)
    Pan with left-click drag, zoom with mouse wheel.
    """

    # ── parse G-code into move segments ──────────────────────────────────────
    segments = []   # list of (x0,y0, x1,y1, is_rapid)

    cx_pos, cy_pos = 0.0, 0.0
    is_rapid = False

    try:
        with open(gcode_path, "r", encoding="utf-8", errors="replace") as f:
            for raw in f:
                line = raw.strip()
                # strip line number
                line = re.sub(r"^N\d+\s*", "", line)

                if not line:
                    continue

                # detect move type
                if line.startswith("G0"):
                    is_rapid = True
                elif re.match(r"G[123]\b", line):
                    is_rapid = False
                else:
                    # non-move command — skip but don't reset position
                    continue

                # extract X / Y (ignore Z for 2D preview)
                xm = re.search(r"X(-?\d+(?:\.\d+)?)", line)
                ym = re.search(r"Y(-?\d+(?:\.\d+)?)", line)

                nx = float(xm.group(1)) if xm else cx_pos
                ny = float(ym.group(1)) if ym else cy_pos

                segments.append((cx_pos, cy_pos, nx, ny, is_rapid))
                cx_pos, cy_pos = nx, ny

    except Exception as e:
        messagebox.showerror("Preview Error", f"Could not read G-code:\n{e}",
                             parent=parent)
        return

    if not segments:
        messagebox.showinfo("Preview", "No moves found in G-code file.",
                            parent=parent)
        return

    # ── compute bounding box ──────────────────────────────────────────────────
    all_x = [s[0] for s in segments] + [s[2] for s in segments]
    all_y = [s[1] for s in segments] + [s[3] for s in segments]
    min_x, max_x = min(all_x), max(all_x)
    min_y, max_y = min(all_y), max(all_y)
    span_x = max_x - min_x or 1.0
    span_y = max_y - min_y or 1.0

    # ── window setup ──────────────────────────────────────────────────────────
    win = tk.Toplevel(parent)
    win.title(f"Path Preview — {os.path.basename(gcode_path)}")
    win.geometry("900x700")
    win.configure(bg=BG)
    win.resizable(True, True)

    # ── toolbar ───────────────────────────────────────────────────────────────
    toolbar = tk.Frame(win, bg=BTN, pady=4)
    toolbar.pack(fill=tk.X)

    tk.Label(toolbar, text="Path Preview", bg=BTN, fg=FG,
             font=("Consolas", 10, "bold")).pack(side=tk.LEFT, padx=10)

    # legend
    tk.Label(toolbar, text="━━", bg=BTN, fg=COLOR_CUT,
             font=("Consolas", 11)).pack(side=tk.LEFT, padx=(20, 2))
    tk.Label(toolbar, text="Cutting (G1)", bg=BTN, fg=FG,
             font=("Consolas", 9)).pack(side=tk.LEFT)

    tk.Label(toolbar, text="━━", bg=BTN, fg=COLOR_RAPID,
             font=("Consolas", 11)).pack(side=tk.LEFT, padx=(16, 2))
    tk.Label(toolbar, text="Rapid (G0)", bg=BTN, fg=FG,
             font=("Consolas", 9)).pack(side=tk.LEFT)

    # stats
    cut_count   = sum(1 for s in segments if not s[4])
    rapid_count = sum(1 for s in segments if s[4])
    stats_text  = (f"  |  Moves: {cut_count:,} cut   {rapid_count:,} rapid"
                   f"  |  W:{span_x:.1f}mm  H:{span_y:.1f}mm")
    tk.Label(toolbar, text=stats_text, bg=BTN, fg="#888888",
             font=("Consolas", 9)).pack(side=tk.LEFT, padx=10)

    tk.Button(toolbar, text="Fit", bg="#3a7afe", fg="white",
              font=("Consolas", 9), relief="flat", padx=8,
              command=lambda: fit_view()).pack(side=tk.RIGHT, padx=8)

    # ── canvas ────────────────────────────────────────────────────────────────
    canvas = tk.Canvas(win, bg=COLOR_BG, highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)

    # ── view state ────────────────────────────────────────────────────────────
    view = {"ox": 0.0, "oy": 0.0, "scale": 1.0}
    drag = {"x": 0, "y": 0, "active": False}

    PADDING = 40   # px padding around content when fitting

    def world_to_canvas(wx, wy):
        """Convert mm world coords to canvas pixel coords."""
        s  = view["scale"]
        px = view["ox"] + wx * s
        # Y axis: G-code Y+ is "up", canvas Y+ is down — flip
        py = view["oy"] - wy * s
        return px, py

    def draw_all():
        canvas.delete("all")
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 2 or ch < 2:
            return

        # grid (light)
        s = view["scale"]
        # choose grid spacing so cells are ~50-100 px
        raw_step = 50.0 / s
        # round to nice number
        magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 1
        for mult in (1, 2, 5, 10):
            if magnitude * mult >= raw_step:
                grid_step = magnitude * mult
                break
        else:
            grid_step = magnitude * 10

        # vertical grid lines
        x_start = math.floor(min_x / grid_step) * grid_step - grid_step
        x_end   = max_x + grid_step
        gx = x_start
        while gx <= x_end:
            px, _ = world_to_canvas(gx, 0)
            if 0 <= px <= cw:
                canvas.create_line(px, 0, px, ch, fill=COLOR_GRID, width=1)
                canvas.create_text(px + 2, ch - 12, text=f"{gx:.0f}",
                                   fill="#555555", font=("Consolas", 7),
                                   anchor="w")
            gx += grid_step

        # horizontal grid lines
        y_start = math.floor(min_y / grid_step) * grid_step - grid_step
        y_end   = max_y + grid_step
        gy = y_start
        while gy <= y_end:
            _, py = world_to_canvas(0, gy)
            if 0 <= py <= ch:
                canvas.create_line(0, py, cw, py, fill=COLOR_GRID, width=1)
                canvas.create_text(4, py - 2, text=f"{gy:.0f}",
                                   fill="#555555", font=("Consolas", 7),
                                   anchor="sw")
            gy += grid_step

        # origin cross
        ox_px, oy_px = world_to_canvas(0, 0)
        canvas.create_line(ox_px, 0, ox_px, ch, fill=COLOR_AXIS, width=1)
        canvas.create_line(0, oy_px, cw, oy_px, fill=COLOR_AXIS, width=1)

        # draw moves — rapid first (underneath), then cuts on top
        for pass_rapid in (True, False):
            color = COLOR_RAPID if pass_rapid else COLOR_CUT
            width = 1 if pass_rapid else 1
            dash  = (4, 4) if pass_rapid else None

            for x0, y0, x1, y1, seg_rapid in segments:
                if seg_rapid != pass_rapid:
                    continue
                px0, py0 = world_to_canvas(x0, y0)
                px1, py1 = world_to_canvas(x1, y1)
                # skip if fully out of view
                if (max(px0, px1) < 0 or min(px0, px1) > cw or
                        max(py0, py1) < 0 or min(py0, py1) > ch):
                    continue
                if dash:
                    canvas.create_line(px0, py0, px1, py1,
                                       fill=color, width=width, dash=dash)
                else:
                    canvas.create_line(px0, py0, px1, py1,
                                       fill=color, width=width)

    def fit_view(event=None):
        """Fit the entire toolpath into the canvas with padding."""
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 2 or ch < 2:
            win.after(50, fit_view)
            return

        scale_x = (cw - 2 * PADDING) / span_x
        scale_y = (ch - 2 * PADDING) / span_y
        s = min(scale_x, scale_y)

        # center
        view["scale"] = s
        view["ox"] = (cw - span_x * s) / 2 - min_x * s
        view["oy"] = (ch - span_y * s) / 2 + max_y * s
        draw_all()

    # ── interaction ───────────────────────────────────────────────────────────

    def on_mouse_press(event):
        drag["x"] = event.x
        drag["y"] = event.y
        drag["active"] = True
        canvas.config(cursor="fleur")

    def on_mouse_release(event):
        drag["active"] = False
        canvas.config(cursor="")

    def on_mouse_drag(event):
        if not drag["active"]:
            return
        dx = event.x - drag["x"]
        dy = event.y - drag["y"]
        view["ox"] += dx
        view["oy"] += dy
        drag["x"] = event.x
        drag["y"] = event.y
        draw_all()

    def on_zoom(event):
        # mouse wheel: zoom toward cursor position
        factor = 1.15 if event.delta > 0 else (1 / 1.15)
        mx, my = event.x, event.y
        view["ox"] = mx + (view["ox"] - mx) * factor
        view["oy"] = my + (view["oy"] - my) * factor
        view["scale"] *= factor
        draw_all()

    canvas.bind("<ButtonPress-1>",   on_mouse_press)
    canvas.bind("<ButtonRelease-1>", on_mouse_release)
    canvas.bind("<B1-Motion>",       on_mouse_drag)
    canvas.bind("<MouseWheel>",      on_zoom)          # Windows / macOS
    canvas.bind("<Button-4>",        lambda e: on_zoom(
        type("E", (), {"delta": 1, "x": e.x, "y": e.y})()))   # Linux scroll up
    canvas.bind("<Button-5>",        lambda e: on_zoom(
        type("E", (), {"delta": -1, "x": e.x, "y": e.y})()))  # Linux scroll down

    canvas.bind("<Configure>", lambda e: fit_view())

    # initial draw after window is laid out
    win.after(100, fit_view)


# ─────────────────────────────────────────────────────────────────────────────
#  CORE CONVERSION
# ─────────────────────────────────────────────────────────────────────────────

def convert_file(input_path, output_path, log_fn,
                 progress_callback=None, use_arc_commands=False):

    total_time_min = 0.0
    current_feed   = None
    last_pos       = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    line_no        = 1

    def nline():
        nonlocal line_no
        s = f"N{line_no:05d} "
        line_no += 1
        return s

    def parse_coord(text, axes=("X", "Y", "Z")):
        """
        Parse axis values from an ISEL command string.
        All values are divided by SCALE (1000) to convert from microns to mm.
        I/J in ISEL are ABSOLUTE coordinates (not offsets from current pos).
        """
        coords = {}
        for axis in axes:
            m = re.search(rf"{axis}(-?\d+(?:\.\d+)?)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    # ── G2/G3 with R format ───────────────────────────────────────────────────
    def arc_to_g2g3(start, end, center, cw):
        nonlocal total_time_min, current_feed

        cx, cy = center["X"], center["Y"]
        r  = math.hypot(start["X"] - cx, start["Y"] - cy)

        a0 = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1 = math.atan2(end["Y"]   - cy, end["X"]   - cx)

        if cw:
            if a1 >= a0: a1 -= 2 * math.pi
        else:
            if a1 <= a0: a1 += 2 * math.pi

        span = abs(a1 - a0)
        arc_len = span * r
        if current_feed:
            total_time_min += arc_len / current_feed

        r_signed = r if span <= math.pi else -r
        code  = "G2" if cw else "G3"
        gline = (nline() +
                 f"{code} X{end['X']:.4f} Y{end['Y']:.4f} R{r_signed:.4f}")

        return end.copy(), [gline]

    # ── arc linearisation – numpy ─────────────────────────────────────────────
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

        angles = np.linspace(a0, a1, steps + 1)[1:]
        xs     = cx + r * np.cos(angles)
        ys     = cy + r * np.sin(angles)

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
            lines.append(nline() +
                         f"G1 X{fp['X']:.3f} Y{fp['Y']:.3f} Z{fp['Z']:.3f}")
            if current_feed:
                total_time_min += dist / current_feed
            pos = fp

        return pos, lines

    # ── arc linearisation – pure Python ──────────────────────────────────────
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

        lines = []
        px, py = start["X"], start["Y"]

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
            lines.append(nline() +
                         f"G1 X{fp['X']:.3f} Y{fp['Y']:.3f} Z{fp['Z']:.3f}")
            if current_feed:
                total_time_min += math.hypot(fp["X"] - pos["X"],
                                             fp["Y"] - pos["Y"]) / current_feed
            pos = fp

        return pos, lines

    # ── choose handler ────────────────────────────────────────────────────────
    if use_arc_commands:
        handle_arc = arc_to_g2g3
    elif _NUMPY:
        handle_arc = linearize_arc_numpy
    else:
        handle_arc = linearize_arc_pure

    # ── first pass: count lines ───────────────────────────────────────────────
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

    # ── open output ───────────────────────────────────────────────────────────
    try:
        out_f = open(output_path, "w", encoding="utf-8", buffering=256 * 1024)
    except Exception as e:
        raise Exception(f"Cannot open output file for writing: {e}")

    def emit(line):
        out_f.write(line + "\n")

    # ── main loop ─────────────────────────────────────────────────────────────
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
                    progress_callback(
                        pct, f"Processing line {lines_processed}/{total_lines}"
                    )

                try:
                    if line.startswith("SPINDLE CW"):
                        m = re.search(r"RPM(\d+)", line)
                        if m:
                            emit(nline() + f"S{int(m.group(1))} M03")
                            log_fn(f"Spindle: {m.group(1)} RPM")
                        else:
                            log_fn(f"Warning: Could not parse RPM from: {line}")

                    elif line.startswith("FASTABS"):
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
                        m = re.search(r"VEL\s*(\d+(?:\.\d+)?)", line)
                        if m:
                            current_feed = float(m.group(1)) / VEL_RATIO
                            emit(nline() + f"F{current_feed:.0f}")
                            log_fn(f"Feed: F{current_feed:.0f}")
                        else:
                            log_fn(f"Warning: Could not parse VEL from: {line}")

                    elif line.startswith("CWABS") or line.startswith("CCWABS"):
                        cw = line.startswith("CWABS")
                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            emit(nline() + f"F{current_feed:.0f}")
                            log_fn(f"Warning: No VEL found, using default F{current_feed:.0f}")

                        c  = parse_coord(line, axes=("X", "Y", "Z"))
                        ij = parse_coord(line, axes=("I", "J"))

                        end = last_pos.copy()
                        end.update(c)

                        # FIX v1.92: ISEL I/J are ABSOLUTE coordinates,
                        # not offsets from current position.
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

        emit(nline() + "M05")
        emit(nline() + "M30")

    finally:
        out_f.close()

    if progress_callback:
        progress_callback(100, "Complete!")

    return total_time_min


# ─────────────────────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────────────────────

def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("520x500")
    root.configure(bg=BG)

    input_var   = tk.StringVar()
    use_arc_var = tk.BooleanVar(value=False)
    converting  = False

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
        if converting:
            return
        path = event.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            log(f"File imported: {path}")
        else:
            log(f"Error: Not a valid file: {path}")

    def browse_input():
        if converting:
            return
        path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if path:
            input_var.set(path)
            log(f"File selected: {path}")

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
                # ── auto-open preview ─────────────────────────────────────
                open_preview(root, out_path)

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
        if converting:
            return
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
        if not out_path:
            return

        arc_mode   = use_arc_var.get()
        converting = True
        convert_btn.config(state="disabled", text="Converting...")
        browse_btn.config(state="disabled")
        progress_bar["value"] = 0
        progress_label.config(text="Starting conversion...")

        logbox.delete(1.0, tk.END)
        log("Starting conversion...")
        log(f"Input:   {in_path}")
        log(f"Output:  {out_path}")
        if arc_mode:
            log("Mode:    G2/G3 arc commands (R format)")
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

    browse_btn = tk.Button(
        btn_frame, text="Browse", command=browse_input,
        bg=BTN, fg=FG, width=12
    )
    browse_btn.pack(side=tk.LEFT, padx=5)

    convert_btn = tk.Button(
        btn_frame, text="Convert", command=convert,
        bg="#3a7afe", fg="white", width=12, font=("Arial", 10, "bold")
    )
    convert_btn.pack(side=tk.LEFT, padx=5)

    numpy_text  = "⚡ numpy detected – arc linearisation accelerated" if _NUMPY \
                  else "⚠ numpy not found – pip install numpy"
    numpy_color = "#00ff88" if _NUMPY else "#ffaa00"
    tk.Label(root, text=numpy_text, bg=BG, fg=numpy_color,
             font=("Arial", 8)).pack(pady=(0, 2))

    arc_frame = tk.Frame(root, bg=BG)
    arc_frame.pack(pady=(2, 8))

    tk.Checkbutton(
        arc_frame,
        text="Use G2/G3 arc commands  (smaller file, requires arc-capable controller)",
        variable=use_arc_var,
        bg=BG, fg=FG, selectcolor=BTN,
        activebackground=BG, activeforeground=FG,
        font=("Arial", 9),
    ).pack()

    progress_frame = tk.Frame(root, bg=BG)
    progress_frame.pack(pady=5, padx=10, fill=tk.X)

    progress_label = tk.Label(
        progress_frame, text="Ready", bg=BG, fg=FG, font=("Arial", 9)
    )
    progress_label.pack()

    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "custom.Horizontal.TProgressbar",
        troughcolor=BTN, background="#3a7afe", thickness=20
    )
    progress_bar = ttk.Progressbar(
        progress_frame, style="custom.Horizontal.TProgressbar",
        orient="horizontal", mode="determinate", length=490
    )
    progress_bar.pack(pady=5)

    tk.Label(root, text="Conversion Log", bg=BG, fg=FG,
             font=("Arial", 9)).pack(pady=(5, 2))

    logbox = tk.Text(
        root, height=10, bg="#121212", fg="#00ff88", font=("Consolas", 9)
    )
    logbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    log("ISEL to G-code Converter Ready")
    log("Drag and drop a file or click Browse to start")
    log("-" * 50)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
