try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.0"

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os, re, math

# ---------------- CONFIG ----------------
SCALE = 1000.0
VEL_RATIO = 16.6667

BG = "#1e1e1e"
FG = "#ffffff"
BTN = "#2d2d2d"

# ----------------------------------------


# ========== ISEL PARSER (CORE FIX) ==========
def parse_isel_moves(path, log):
    moves = []

    pos = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    buf = pos.copy()
    last_axis = None
    current_feed = None
    total_time = 0.0

    def flush(mode):
        nonlocal pos, total_time
        if buf != pos:
            moves.append((mode, pos.copy(), buf.copy()))
            if mode == "G1" and current_feed:
                d = math.sqrt(
                    (buf["X"] - pos["X"]) ** 2 +
                    (buf["Y"] - pos["Y"]) ** 2 +
                    (buf["Z"] - pos["Z"]) ** 2
                )
                total_time += d / current_feed
            pos = buf.copy()

    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith(";"):
                continue

            if line.startswith("SPINDLE CW"):
                rpm = re.search(r"RPM(\d+)", line)
                if rpm:
                    log(f"Spindle: {rpm.group(1)} RPM")

            elif line.startswith("VEL"):
                vel = int(re.search(r"VEL\s*(\d+)", line).group(1))
                current_feed = vel / VEL_RATIO
                log(f"Feed: F{current_feed:.0f}")

            elif line.startswith("FASTABS"):
                flush("G1")
                for a in "XYZ":
                    m = re.search(rf"{a}(-?\d+)", line)
                    if m:
                        buf[a] = float(m.group(1)) / SCALE
                flush("G0")
                last_axis = None

            elif line.startswith("MOVEABS"):
                for a in "XYZ":
                    m = re.search(rf"{a}(-?\d+)", line)
                    if m:
                        if last_axis == a:
                            flush("G1")
                        buf[a] = float(m.group(1)) / SCALE
                        last_axis = a

    flush("G1")
    return moves, total_time


# ========== GCODE WRITER ==========
def write_gcode(moves, output_path):
    n = 1
    out = ["G21", "G17", "G90"]

    def nl(cmd):
        nonlocal n
        s = f"N{n:05d} {cmd}"
        n += 1
        return s

    for mode, p1, p2 in moves:
        cmd = mode
        for a in "XYZ":
            if p1[a] != p2[a]:
                cmd += f" {a}{p2[a]:.3f}"
        out.append(nl(cmd))

    out.append(nl("M05"))
    out.append(nl("M30"))

    with open(output_path, "w") as f:
        f.write("\n".join(out))


# ========== FAKE ISOMETRIC PREVIEW ==========
def draw_preview(canvas, moves):
    canvas.delete("all")

    def iso(p):
        return (
            p["X"] - p["Y"],
            (p["X"] + p["Y"]) * 0.5 - p["Z"]
        )

    for mode, p1, p2 in moves:
        x1, y1 = iso(p1)
        x2, y2 = iso(p2)

        z_only = (p1["X"] == p2["X"] and p1["Y"] == p2["Y"])

        if mode == "G0":
            color = "#5555ff" if not z_only else "#00ffff"
        else:
            color = "#00ff00" if not z_only else "#ffcc00"

        canvas.create_line(
            x1 + 250, y1 + 200,
            x2 + 250, y2 + 200,
            fill=color,
            width=2
        )


# ========== GUI ==========
def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("720x520")
    root.configure(bg=BG)

    input_var = tk.StringVar()

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def drop(e):
        path = e.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            preview()

    def browse():
        input_var.set(
            filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        )
        preview()

    def preview():
        if not input_var.get():
            return
        moves, est = parse_isel_moves(input_var.get(), log)
        draw_preview(canvas, moves)

        m = int(est)
        s = int((est - m) * 60)
        log(f"⏱ Estimated time: {m} min {s} sec")
        log("Note: program time may change according to machine parameters")

    def convert():
        if not input_var.get():
            messagebox.showerror("Error", "No input file")
            return

        base = os.path.splitext(os.path.basename(input_var.get()))[0]
        out = filedialog.asksaveasfilename(
            defaultextension=".ngc",
            initialfile=base + ".ngc",
            filetypes=[("G-code", "*.ngc"), ("All", "*.*")]
        )
        if not out:
            return

        moves, _ = parse_isel_moves(input_var.get(), log)
        write_gcode(moves, out)
        messagebox.showinfo("OK", "Conversion completed")

    tk.Label(root, text="ISEL File", bg=BG, fg=FG).pack(pady=4)

    entry = tk.Entry(root, textvariable=input_var, width=70,
                     bg=BTN, fg=FG, insertbackground=FG)
    entry.pack()
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind("<<Drop>>", drop)

    tk.Button(root, text="Browse", command=browse).pack(pady=4)
    tk.Button(root, text="Convert", command=convert,
              bg="#3a7afe", fg="white").pack(pady=6)

    canvas = tk.Canvas(root, width=500, height=260,
                       bg="#111111", highlightthickness=0)
    canvas.pack(pady=8)

    logbox = tk.Text(root, height=6, bg="#121212",
                     fg="#00ff88", insertbackground="white")
    logbox.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
