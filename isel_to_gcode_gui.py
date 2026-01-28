try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.2"

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import os
import math

SCALE = 1000.0
VEL_RATIO = 16.6667
SAFE_Z = 4.0
ARC_RESOLUTION = 0.01  # mm
DEFAULT_FEED = 100.0  # Default feed rate if no VEL command

BG = "#1e1e1e"
FG = "#ffffff"
BTN = "#2d2d2d"


def convert_file(input_path, output_path, log):
    total_time_min = 0.0
    current_feed = None
    last_pos = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    line_no = 1
    start_done = False

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
        coords = {}
        for axis in axes:
            m = re.search(rf"{axis}(-?\d+)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    def linearize_arc(start, end, center, cw):
        nonlocal total_time_min, current_feed

        sx, sy = start["X"], start["Y"]
        ex, ey = end["X"], end["Y"]
        cx, cy = center["X"], center["Y"]

        r = math.hypot(sx - cx, sy - cy)
        a0 = math.atan2(sy - cy, sx - cx)
        a1 = math.atan2(ey - cy, ex - cx)

        if cw:
            if a1 >= a0:
                a1 -= 2 * math.pi
        else:
            if a1 <= a0:
                a1 += 2 * math.pi

        arc_len = abs(a1 - a0) * r
        steps = max(1, int(arc_len / ARC_RESOLUTION))
        da = (a1 - a0) / steps

        pos = start.copy()

        for i in range(1, steps + 1):
            a = a0 + da * i
            target = {
                "X": cx + r * math.cos(a),
                "Y": cy + r * math.sin(a),
                "Z": start["Z"]
            }

            gcode.append(
                nline() +
                f"G1 X{target['X']:.3f} Y{target['Y']:.3f} Z{target['Z']:.3f}"
            )

            if current_feed:
                total_time_min += move_distance(pos, target) / current_feed

            pos = target

        # Ensure final position matches end point exactly
        final_pos = end.copy()
        if abs(final_pos["X"] - pos["X"]) > 0.001 or abs(final_pos["Y"] - pos["Y"]) > 0.001:
            gcode.append(
                nline() +
                f"G1 X{final_pos['X']:.3f} Y{final_pos['Y']:.3f} Z{final_pos['Z']:.3f}"
            )
            if current_feed:
                total_time_min += move_distance(pos, final_pos) / current_feed
            pos = final_pos

        return pos

    gcode = ["G21", "G17", "G90"]

    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith(";"):
                    continue

                try:
                    if line.startswith("SPINDLE CW"):
                        match = re.search(r"RPM(\d+)", line)
                        if match:
                            rpm = int(match.group(1))
                            gcode.append(nline() + f"S{rpm} M03")
                            log(f"Spindle: {rpm} RPM")
                        else:
                            log(f"Warning: Could not parse RPM from: {line}")

                    elif line.startswith("FASTABS"):
                        c = parse_coord(line)

                        if not start_done:
                            # Only move Z to safe height if current Z is lower
                            if last_pos["Z"] < SAFE_Z:
                                gcode.append(nline() + f"G0 Z{SAFE_Z:.3f}")
                                last_pos["Z"] = SAFE_Z
                            start_done = True

                        target = last_pos.copy()
                        target.update(c)

                        cmd = "G0"
                        for k in ["X", "Y", "Z"]:
                            cmd += f" {k}{target[k]:.3f}"

                        gcode.append(nline() + cmd)
                        last_pos = target

                    elif line.startswith("MOVEABS"):
                        c = parse_coord(line)
                        target = last_pos.copy()
                        target.update(c)

                        # Ensure feed rate is set
                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            gcode.append(nline() + f"F{current_feed:.0f}")
                            log(f"Warning: No VEL command found, using default F{current_feed:.0f}")

                        cmd = "G1"
                        for k in ["X", "Y", "Z"]:
                            cmd += f" {k}{target[k]:.3f}"

                        gcode.append(nline() + cmd)

                        if current_feed:
                            total_time_min += move_distance(last_pos, target) / current_feed

                        last_pos = target

                    elif line.startswith("VEL"):
                        match = re.search(r"VEL\s*(\d+)", line)
                        if match:
                            vel = int(match.group(1))
                            current_feed = vel / VEL_RATIO
                            gcode.append(nline() + f"F{current_feed:.0f}")
                            log(f"Feed: F{current_feed:.0f}")
                        else:
                            log(f"Warning: Could not parse VEL from: {line}")

                    elif line.startswith("CWABS") or line.startswith("CCWABS"):
                        cw = line.startswith("CWABS")

                        # Ensure feed rate is set for arc moves
                        if current_feed is None:
                            current_feed = DEFAULT_FEED
                            gcode.append(nline() + f"F{current_feed:.0f}")
                            log(f"Warning: No VEL command found, using default F{current_feed:.0f}")

                        c = parse_coord(line, axes=("X", "Y", "Z"))
                        ij = parse_coord(line, axes=("I", "J"))

                        end = last_pos.copy()
                        end.update(c)

                        center = {
                            "X": last_pos["X"] + ij.get("I", 0.0),
                            "Y": last_pos["Y"] + ij.get("J", 0.0),
                        }

                        last_pos = linearize_arc(last_pos, end, center, cw)

                except Exception as e:
                    log(f"Warning: Error processing line: {line}")
                    log(f"  Error: {str(e)}")
                    continue

    except FileNotFoundError:
        raise FileNotFoundError(f"Input file not found: {input_path}")
    except UnicodeDecodeError:
        raise UnicodeDecodeError(
            'utf-8', b'', 0, 1, 
            f"Could not read file {input_path}. File encoding may not be UTF-8."
        )
    except Exception as e:
        raise Exception(f"Error reading input file: {str(e)}")

    gcode += [nline() + "M05", nline() + "M30"]

    try:
        with open(output_path, "w", encoding='utf-8') as f:
            f.write("\n".join(gcode))
    except Exception as e:
        raise Exception(f"Error writing output file: {str(e)}")

    return total_time_min


def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("520x400")
    root.configure(bg=BG)

    input_var = tk.StringVar()

    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def drop(event):
        path = event.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            log(f"File imported: {path}")
        else:
            log(f"Error: Not a valid file: {path}")

    def browse_input():
        path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if path:
            input_var.set(path)
            log(f"File selected: {path}")

    def convert():
        if not input_var.get():
            messagebox.showerror("Error", "No input file selected")
            return

        in_path = input_var.get()
        
        if not os.path.isfile(in_path):
            messagebox.showerror("Error", f"File not found: {in_path}")
            return

        base = os.path.splitext(os.path.basename(in_path))[0]

        out_path = filedialog.asksaveasfilename(
            defaultextension=".ngc",
            initialfile=base + ".ngc",
            filetypes=[("G-code Files", "*.ngc"), ("All Files", "*.*")]
        )

        if not out_path:
            return

        try:
            logbox.delete(1.0, tk.END)
            log("Starting conversion...")
            log(f"Input: {in_path}")
            log(f"Output: {out_path}")
            log("-" * 50)
            
            total_time = convert_file(in_path, out_path, log)
            
            m = int(total_time)
            s = int((total_time - m) * 60)

            log("-" * 50)
            log(f"✓ Conversion completed successfully")
            log(f"⏱ Estimated program time: {m} min {s} sec")
            log("⚠ Program time may change according to machine parameters")

            messagebox.showinfo(
                "Completed",
                f"Conversion finished successfully!\n\nEstimated time:\n{m} min {s} sec"
            )
        
        except FileNotFoundError as e:
            log(f"✗ Error: {str(e)}")
            messagebox.showerror("File Error", str(e))
        except UnicodeDecodeError as e:
            log(f"✗ Error: File encoding problem")
            messagebox.showerror("Encoding Error", "Could not read file. Please check file encoding.")
        except Exception as e:
            log(f"✗ Error: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed:\n{str(e)}")

    # GUI Layout
    tk.Label(root, text="ISEL File", bg=BG, fg=FG, font=("Arial", 10)).pack(pady=5)

    entry = tk.Entry(root, textvariable=input_var, width=60, bg=BTN, fg=FG)
    entry.pack(padx=10)
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind("<<Drop>>", drop)

    btn_frame = tk.Frame(root, bg=BG)
    btn_frame.pack(pady=10)

    tk.Button(
        btn_frame, 
        text="Browse", 
        command=browse_input,
        bg=BTN,
        fg=FG,
        width=12
    ).pack(side=tk.LEFT, padx=5)

    tk.Button(
        btn_frame, 
        text="Convert", 
        bg="#3a7afe", 
        fg="white", 
        command=convert,
        width=12,
        font=("Arial", 10, "bold")
    ).pack(side=tk.LEFT, padx=5)

    tk.Label(root, text="Conversion Log", bg=BG, fg=FG, font=("Arial", 9)).pack(pady=(5, 2))

    logbox = tk.Text(root, height=12, bg="#121212", fg="#00ff88", font=("Consolas", 9))
    logbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # Welcome message
    log("ISEL to G-code Converter Ready")
    log("Drag and drop a file or click Browse to start")
    log("-" * 50)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
