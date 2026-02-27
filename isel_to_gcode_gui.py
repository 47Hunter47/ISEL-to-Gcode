try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.7"

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

BG  = "#1e1e1e"
FG  = "#ffffff"
BTN = "#2d2d2d"


# ─────────────────────────────────────────────────────────────────────────────
#  CORE CONVERSION
# ─────────────────────────────────────────────────────────────────────────────

def convert_file(input_path, output_path, log_fn,
                 progress_callback=None, use_arc_commands=False):

    total_time_min = 0.0
    current_feed   = None
    last_pos       = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    line_no        = 1
    start_done     = False

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

    # ── G2/G3 with R format ───────────────────────────────────────────────────
    #
    # WHY R and not I/J:
    # ──────────────────
    # ISEL stores arc center as offset from START (I/J) and end point as
    # absolute coords (X/Y) – both in the same integer unit space.
    # After /1000 conversion, floating-point rounding makes:
    #   r_from_start  ≠  r_from_end   (can differ by >1 mm for large arcs)
    #
    # With I/J format the controller checks both radii and alarms.
    # With R format the controller uses the START point + R to draw the arc
    # and drives to the programmed END point – no radius consistency check.
    # Logosol CNC explicitly supports: G2 X.. Y.. R.. F..
    #
    # R sign convention (standard):
    #   R > 0  →  arc ≤ 180°  (minor arc)
    #   R < 0  →  arc > 180°  (major arc)

    def arc_to_g2g3(start, end, center, cw):
        nonlocal total_time_min, current_feed

        cx, cy = center["X"], center["Y"]
        r  = math.hypot(start["X"] - cx, start["Y"] - cy)

        # arc angle span for time estimate + R sign
        a0 = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1 = math.atan2(end["Y"]   - cy, end["X"]   - cx)

        if cw:
            if a1 >= a0: a1 -= 2 * math.pi
        else:
            if a1 <= a0: a1 += 2 * math.pi

        span = abs(a1 - a0)           # always positive, 0 < span ≤ 2π

        arc_len = span * r
        if current_feed:
            total_time_min += arc_len / current_feed

        # R sign: negative means major arc (> 180°)
        r_signed = r if span <= math.pi else -r

        code  = "G2" if cw else "G3"
        # Use ISEL's original end point – no reprojection needed with R format
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
        for hdr in ("G21", "G17", "G90"):
            emit(hdr)

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
                        c = parse_coord(line)
                        if not start_done:
                            if last_pos["Z"] < SAFE_Z:
                                emit(nline() + f"G0 Z{SAFE_Z:.3f}")
                                last_pos["Z"] = SAFE_Z
                            start_done = True
                        target = last_pos.copy()
                        target.update(c)
                        cmd = "G0"
                        for k in ("X", "Y", "Z"):
                            if k in c:
                                cmd += f" {k}{target[k]:.3f}"
                        emit(nline() + cmd)
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
                        end    = last_pos.copy()
                        end.update(c)
                        center = {
                            "X": last_pos["X"] + ij.get("I", 0.0),
                            "Y": last_pos["Y"] + ij.get("J", 0.0),
                        }

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
