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

# ── machine constants ─────────────────────────────────────────────────────────
SCALE          = 1000.0   # ISEL units → mm  (1 unit = 0.001 mm)
VEL_RATIO      = 16.6667  # ISEL VEL → G-code F (mm/min)
SAFE_Z         = 4.0      # mm – retract height before first rapid
ARC_RESOLUTION = 0.05     # mm – chord length when linearising arcs (G1 mode)
DEFAULT_FEED   = 1000.0    # mm/min – fallback if no VEL found before first move

BG  = "#1e1e1e"
FG  = "#ffffff"
BTN = "#2d2d2d"


# ─────────────────────────────────────────────────────────────────────────────
#  CORE CONVERSION
# ─────────────────────────────────────────────────────────────────────────────

def convert_file(input_path, output_path, log, progress_callback=None,
                 use_arc_commands=False):
    """
    Convert an ISEL NC file to G-code.

    use_arc_commands=False  → arcs linearised to G1 segments (safe / universal)
    use_arc_commands=True   → arcs emitted as G2/G3 (smaller file,
                               requires arc-capable controller)
    """

    total_time_min = 0.0
    current_feed   = None
    last_pos       = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    line_no        = 1
    start_done     = False

    # ── helpers ──────────────────────────────────────────────────────────────

    def nline():
        nonlocal line_no
        s = f"N{line_no:05d} "
        line_no += 1
        return s

    def move_distance(p1, p2):
        return math.sqrt(
            (p2["X"] - p1["X"]) ** 2 +
            (p2["Y"] - p1["Y"]) ** 2 +
            (p2["Z"] - p1["Z"]) ** 2
        )

    def parse_coord(text, axes=("X", "Y", "Z")):
        """
        Extract axis values from an ISEL command string.
        Handles integers AND decimals, e.g. X-3250, X12.5, X-3.25
        Divides by SCALE to convert ISEL units → mm.
        """
        coords = {}
        for axis in axes:
            m = re.search(rf"{axis}(-?\d+(?:\.\d+)?)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    # ── arc as single G2/G3 line ──────────────────────────────────────────────

    def arc_to_g2g3(start, end, center, cw):
        """
        Emit one G2/G3 line using I/J centre offsets.

        FIX: "radius to end of arc differs from radius to start"
        ─────────────────────────────────────────────────────────
        The error occurs because floating-point division (ISEL integer / 1000.0)
        leaves a tiny difference between:
            r_start = hypot(start - center)
            r_end   = hypot(end   - center)

        Most controllers check this and alarm if the difference exceeds their
        internal tolerance (often 0.001–0.005 mm).

        Solution: recompute the end point by projecting it back onto the circle
        defined by the start radius.  This guarantees r_start == r_end to full
        floating-point precision, so the controller's check always passes.

        Z is intentionally omitted from the G2/G3 line (ISEL arcs are always
        XY-planar; mixing Z in the same block confuses many controllers).
        """
        nonlocal total_time_min, current_feed

        cx, cy = center["X"], center["Y"]

        # radius from start point (authoritative)
        r = math.hypot(start["X"] - cx, start["Y"] - cy)

        # --- arc angle span ---
        a0 = math.atan2(start["Y"] - cy, start["X"] - cx)
        a1 = math.atan2(end["Y"]   - cy, end["X"]   - cx)

        if cw:
            if a1 >= a0:
                a1 -= 2 * math.pi
        else:
            if a1 <= a0:
                a1 += 2 * math.pi

        arc_len = abs(a1 - a0) * r
        if current_feed:
            total_time_min += arc_len / current_feed

        # --- corrected end point (lies exactly on the circle) ---
        end_x = cx + r * math.cos(a1)
        end_y = cy + r * math.sin(a1)

        # I/J are offsets from START to centre
        i_off = cx - start["X"]
        j_off = cy - start["Y"]

        code  = "G2" if cw else "G3"
        gline = (
            nline() +
            f"{code} X{end_x:.4f} Y{end_y:.4f}"
            f" I{i_off:.4f} J{j_off:.4f}"
        )
        # update end with corrected coordinates
        corrected_end = end.copy()
        corrected_end["X"] = end_x
        corrected_end["Y"] = end_y

        return corrected_end, [gline]

    # ── arc as linearised G1 segments ─────────────────────────────────────────

    def linearize_arc(start, end, center, cw):
        nonlocal total_time_min, current_feed

        sx, sy = start["X"],  start["Y"]
        ex, ey = end["X"],    end["Y"]
        cx, cy = center["X"], center["Y"]

        r  = math.hypot(sx - cx, sy - cy)
        a0 = math.atan2(sy - cy, sx - cx)
        a1 = math.atan2(ey - cy, ex - cx)

        if cw:
            if a1 >= a0:
                a1 -= 2 * math.pi
        else:
            if a1 <= a0:
                a1 += 2 * math.pi

        arc_len = abs(a1 - a0) * r
        steps   = max(1, int(arc_len / ARC_RESOLUTION))
        da      = (a1 - a0) / steps

        lines = []
        pos   = start.copy()

        for i in range(1, steps + 1):
            a      = a0 + da * i
            target = {
                "X": cx + r * math.cos(a),
                "Y": cy + r * math.sin(a),
                "Z": start["Z"],
            }
            lines.append(
                nline() +
                f"G1 X{target['X']:.3f} Y{target['Y']:.3f} Z{target['Z']:.3f}"
            )
            if current_feed:
                total_time_min += move_distance(pos, target) / current_feed
            pos = target

        # exact end-point correction
        final_pos = end.copy()
        if abs(final_pos["X"] - pos["X"]) > 0.001 or \
           abs(final_pos["Y"] - pos["Y"]) > 0.001:
            lines.append(
                nline() +
                f"G1 X{final_pos['X']:.3f} Y{final_pos['Y']:.3f} Z{final_pos['Z']:.3f}"
            )
            if current_feed:
                total_time_min += move_distance(pos, final_pos) / current_feed
            pos = final_pos

        return pos, lines

    # ── select arc handler once ───────────────────────────────────────────────

    handle_arc = arc_to_g2g3 if use_arc_commands else linearize_arc

    # ── first pass: count non-comment lines for progress bar ─────────────────

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
        progress_callback(0, "Reading file...")

    # ── open output for streaming write ──────────────────────────────────────

    try:
        out_f = open(output_path, "w", encoding="utf-8")
    except Exception as e:
        raise Exception(f"Cannot open output file for writing: {e}")

    def emit(line):
        out_f.write(line + "\n")

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
                    lines_processed % 50 == 0 or lines_processed == total_lines
                ):
                    pct = (lines_processed / total_lines) * 95
                    progress_callback(
                        pct, f"Processing line {lines_processed}/{total_lines}"
                    )

                try:
                    # ── SPINDLE ───────────────────────────────────────────
                    if line.startswith("SPINDLE CW"):
                        m = re.search(r"RPM(\d+)", line)
                        if m:
                            rpm = int(m.group(1))
                            emit(nline() + f"S{rpm} M03")
                            log(f"Spindle: {rpm} RPM")
                        else:
                            log(f"Warning: Could not parse RPM from: {line}")

                    # ── FASTABS (rapid move) ──────────────────────────────
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

                    # ── MOVEABS (feed move) ───────────────────────────────
                    elif line.startswith("MOVEABS"):
                        c      = parse_coord(line)
                        target = last_pos.copy()
                        target.update(c)

                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            emit(nline() + f"F{current_feed:.0f}")
                            log(f"Warning: No VEL command found, using default F{current_feed:.0f}")

                        cmd = "G1"
                        for k in ("X", "Y", "Z"):
                            if k in c:
                                cmd += f" {k}{target[k]:.3f}"
                        emit(nline() + cmd)

                        if current_feed:
                            total_time_min += move_distance(last_pos, target) / current_feed
                        last_pos = target

                    # ── VEL (feed rate) ───────────────────────────────────
                    elif line.startswith("VEL"):
                        m = re.search(r"VEL\s*(\d+(?:\.\d+)?)", line)
                        if m:
                            vel          = float(m.group(1))
                            current_feed = vel / VEL_RATIO
                            emit(nline() + f"F{current_feed:.0f}")
                            log(f"Feed: F{current_feed:.0f}")
                        else:
                            log(f"Warning: Could not parse VEL from: {line}")

                    # ── CWABS / CCWABS (arc move) ─────────────────────────
                    elif line.startswith("CWABS") or line.startswith("CCWABS"):
                        cw = line.startswith("CWABS")

                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            emit(nline() + f"F{current_feed:.0f}")
                            log(f"Warning: No VEL command found, using default F{current_feed:.0f}")

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
                    log(f"Warning: Error processing line: {line}")
                    log(f"  Error: {e}")
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
    root.geometry("520x490")
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
            mode_str = "G2/G3 arc commands" if arc_mode else "G1 linearised"

            def on_done():
                log("-" * 50)
                log(f"✓ Conversion completed  ({mode_str})")
                log(f"⏱ Estimated program time: {m} min {s} sec")
                log("⚠ Program time may change according to machine parameters")
                progress_label.config(text="Conversion complete!")
                messagebox.showinfo(
                    "Completed",
                    f"Conversion finished successfully!\n"
                    f"Mode: {mode_str}\n\n"
                    f"Estimated time:\n{m} min {s} sec"
                )

            root.after(0, on_done)

        except FileNotFoundError as e:
            root.after(0, lambda: log(f"✗ Error: {e}"))
            root.after(0, lambda: progress_label.config(text="Error occurred"))
            root.after(0, lambda: messagebox.showerror("File Error", str(e)))

        except UnicodeDecodeError:
            root.after(0, lambda: log("✗ Error: File encoding problem"))
            root.after(0, lambda: progress_label.config(text="Error occurred"))
            root.after(0, lambda: messagebox.showerror(
                "Encoding Error",
                "Could not read file. Please check file encoding."
            ))

        except Exception as e:
            root.after(0, lambda: log(f"✗ Error: {e}"))
            root.after(0, lambda: progress_label.config(text="Error occurred"))
            root.after(0, lambda: messagebox.showerror(
                "Error", f"Conversion failed:\n{e}"
            ))

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
        log(f"Input:  {in_path}")
        log(f"Output: {out_path}")
        log(f"Mode:   {'G2/G3 arc commands' if arc_mode else 'G1 linearised'}")
        log("-" * 50)

        threading.Thread(
            target=_do_convert,
            args=(in_path, out_path, arc_mode),
            daemon=True
        ).start()

    # ── layout ────────────────────────────────────────────────────────────────

    tk.Label(root, text="ISEL File", bg=BG, fg=FG,
             font=("Arial", 10)).pack(pady=5)

    entry = tk.Entry(root, textvariable=input_var, width=60, bg=BTN, fg=FG)
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

    # G2/G3 toggle
    arc_frame = tk.Frame(root, bg=BG)
    arc_frame.pack(pady=(0, 8))

    tk.Checkbutton(
        arc_frame,
        text="Use G2/G3 arc commands  (smaller file, requires arc-capable controller)",
        variable=use_arc_var,
        bg=BG, fg=FG,
        selectcolor=BTN,
        activebackground=BG,
        activeforeground=FG,
        font=("Arial", 9),
    ).pack()

    # Progress
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
        orient="horizontal", mode="determinate", length=480
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


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_gui()
