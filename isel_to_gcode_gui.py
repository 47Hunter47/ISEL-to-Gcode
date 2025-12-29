try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.0"

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import os
import math

# ---------------- CONFIG ----------------
SCALE = 1000.0
VEL_RATIO = 16.6667

BG = "#1e1e1e"
FG = "#ffffff"
BTN = "#2d2d2d"

COLOR_G0_XY = "#bbbbbb"
COLOR_G1_XY = "#00ff88"
COLOR_G0_Z  = "#4aa3ff"
COLOR_G1_Z  = "#ff5555"

ISO_SCALE = 40
Z_SCALE = 25
CANVAS_SIZE = 420
# ----------------------------------------

def parse_isel(input_path):
    moves = []
    last = {"X": 0.0, "Y": 0.0, "Z": 0.0}

    def get(axis, text):
        m = re.search(rf"{axis}(-?\d+)", text)
        return float(m.group(1)) / SCALE if m else last[axis]

    with open(input_path) as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            if line.startswith(("FASTABS", "MOVEABS")):
                mode = "G0" if line.startswith("FASTABS") else "G1"

                new = {
                    "X": get("X", line),
                    "Y": get("Y", line),
                    "Z": get("Z", line),
                }

                moves.append((mode, last.copy(), new.copy()))
                last = new

    return moves

def iso(x, y, z):
    ix = (x - y) * ISO_SCALE
    iy = (x + y) * ISO_SCALE * 0.5 - z * Z_SCALE
    return ix, iy


def draw_simulation(canvas, moves):
    canvas.delete("all")

    cx = CANVAS_SIZE // 2
    cy = CANVAS_SIZE // 2 + 60

    for mode, p1, p2 in moves:
        x1, y1 = iso(p1["X"], p1["Y"], p1["Z"])
        x2, y2 = iso(p2["X"], p2["Y"], p2["Z"])

        x1 += cx
        y1 += cy
        x2 += cx
        y2 += cy

        # Z movement?
        if p1["X"] == p2["X"] and p1["Y"] == p2["Y"]:
            color = COLOR_G0_Z if mode == "G0" else COLOR_G1_Z
        else:
            color = COLOR_G0_XY if mode == "G0" else COLOR_G1_XY

        canvas.create_line(
            x1, y1, x2, y2,
            fill=color,
            width=2
        )


def convert_file(input_path, output_path, log):
    total_time_min = 0.0
    current_feed = None
    last_pos = {"X": None, "Y": None, "Z": None}

    def move_distance(p1, p2):
        d = 0.0
        for a in ["X", "Y", "Z"]:
            if p1[a] is not None and p2[a] is not None:
                d += (p2[a] - p1[a]) ** 2
        return math.sqrt(d)

    def parse_coord(text):
        coords = {}
        for axis in ['X', 'Y', 'Z']:
            m = re.search(rf"{axis}(-?\d+)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    gcode = ["G21", "G17", "G90"]

    with open(input_path) as f:
        for line in f:
            line = line.strip()

            if line.startswith("VEL"):
                vel = int(re.search(r"VEL\s*(\d+)", line).group(1))
                current_feed = vel / VEL_RATIO
                gcode.append(f"F{current_feed:.0f}")
                log(f"Feed: F{current_feed:.0f}")

            elif line.startswith("SPINDLE CW"):
                rpm = re.search(r"RPM(\d+)", line).group(1)
                gcode.append(f"S{rpm} M03")
                log(f"Spindle: {rpm} RPM")

            elif line.startswith("FASTABS"):
                c = parse_coord(line)
                cmd = "G0"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(cmd)
                last_pos.update(c)

            elif line.startswith("MOVEABS"):
                c = parse_coord(line)
                cmd = "G1"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(cmd)

                if current_feed and all(last_pos[a] is not None for a in c):
                    total_time_min += move_distance(last_pos, c) / current_feed

                last_pos.update(c)

    gcode += ["M05", "M30"]

    with open(output_path, "w") as f:
        f.write("\n".join(gcode))

    m = int(total_time_min)
    s = int((total_time_min - m) * 60)
    log(f"⏱ Estimated time: {m} min {s} sec")
    log("Note: program time may change according to machine parameters")


def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("520x620")
    root.configure(bg=BG)

    input_var = tk.StringVar()

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def load_and_draw(path):
        moves = parse_isel(path)
        draw_simulation(canvas, moves)

    def drop(event):
        path = event.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            log(f"File loaded: {path}")
            load_and_draw(path)

    def browse():
        p = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if p:
            input_var.set(p)
            load_and_draw(p)

    def convert():
        if not input_var.get():
            return
        base = os.path.splitext(os.path.basename(input_var.get()))[0]
        out = filedialog.asksaveasfilename(
            defaultextension=".ngc",
            initialfile=base + ".ngc"
        )
        if out:
            convert_file(input_var.get(), out, log)
            messagebox.showinfo("OK", "Convert completed")

    tk.Label(root, text="ISEL File", bg=BG, fg=FG).pack()
    entry = tk.Entry(root, textvariable=input_var, bg=BTN, fg=FG, width=55)
    entry.pack()
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind("<<Drop>>", drop)

    tk.Button(root, text="Browse", command=browse).pack(pady=5)

    canvas = tk.Canvas(
        root,
        width=CANVAS_SIZE,
        height=CANVAS_SIZE,
        bg="#111111",
        highlightthickness=0
    )
    canvas.pack(pady=10)

    tk.Button(root, text="Convert", bg="#3a7afe", fg="white", command=convert).pack()

    logbox = tk.Text(root, height=6, bg="#121212", fg="#00ff88")
    logbox.pack(fill="both", expand=True, pady=5)

    root.mainloop()


if __name__ == "__main__":
    run_gui()


