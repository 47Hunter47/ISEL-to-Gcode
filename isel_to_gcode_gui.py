try:
    from version import APP_VERSION
except ImportError:
    APP_VERSION = "1.0"

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import os

SCALE = 1000.0
VEL_RATIO = 16.6667

BG = "#1e1e1e"
FG = "#ffffff"
BTN = "#2d2d2d"


def convert_file(input_path, output_path, log):
    total_time_min = 0.0
    current_feed = None
    last_pos = {"X": None, "Y": None, "Z": None}
    line_no = 1

    def nline():
        nonlocal line_no
        s = f"N{line_no:05d} "
        line_no += 1
        return s

    def move_distance(p1, p2):
        dist = 0.0
        for a in ["X", "Y", "Z"]:
            if p1[a] is not None and p2[a] is not None:
                dist += (p2[a] - p1[a]) ** 2
        return dist ** 0.5

    def parse_coord(text):
        coords = {}
        for axis in ["X", "Y", "Z"]:
            m = re.search(rf"{axis}(-?\d+)", text)
            if m:
                coords[axis] = float(m.group(1)) / SCALE
        return coords

    gcode = ["G21", "G17", "G90"]

    with open(input_path) as f:
        for line in f:
            line = line.strip()

            if not line or line.startswith(";"):
                continue

            if line.startswith("SPINDLE CW"):
                rpm = re.search(r"RPM(\d+)", line).group(1)
                gcode.append(nline() + f"S{rpm} M03")
                log(f"Spindle: {rpm} RPM")

            elif line.startswith("FASTABS"):
                c = parse_coord(line)
                cmd = "G0"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(nline() + cmd)
                last_pos.update(c)

            elif line.startswith("MOVEABS"):
                c = parse_coord(line)
                cmd = "G1"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(nline() + cmd)

                if current_feed and all(last_pos[a] is not None for a in c):
                    dist = move_distance(last_pos, c)
                    total_time_min += dist / current_feed

                last_pos.update(c)

            elif line.startswith("VEL"):
                vel = int(re.search(r"VEL\s*(\d+)", line).group(1))
                current_feed = vel / VEL_RATIO
                gcode.append(nline() + f"F{current_feed:.0f}")
                log(f"Feed: F{current_feed:.0f}")

    gcode += [nline() + "M05", nline() + "M30"]

    with open(output_path, "w") as f:
        f.write("\n".join(gcode))

    return total_time_min


def run_gui():
    root = TkinterDnD.Tk()
    root.title(f"ISEL → G-code Converter v{APP_VERSION}")
    root.geometry("520x380")
    root.configure(bg=BG)

    input_var = tk.StringVar()

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def drop(event):
        path = event.data.strip("{}")
        if os.path.isfile(path):
            input_var.set(path)
            log(f"File imported: {path}")

    def browse_input():
        input_var.set(
            filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        )

    def convert():
        if not input_var.get():
            messagebox.showerror("Error", "No input file selected")
            return

        in_path = input_var.get()
        base = os.path.splitext(os.path.basename(in_path))[0]

        out_path = filedialog.asksaveasfilename(
            defaultextension=".ngc",
            initialfile=base + ".ngc",
            filetypes=[("G-code Files", "*.ngc"), ("All Files", "*.*")]
        )

        if not out_path:
            return

        try:
            total_time = convert_file(in_path, out_path, log)
            m = int(total_time)
            s = int((total_time - m) * 60)

            log(f"⏱ Estimated program time: {m} min {s} sec")
            log("⚠ Program time may change according to machine parameters")

            messagebox.showinfo(
                "Completed",
                f"Conversion finished.\n\nEstimated time:\n{m} min {s} sec"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Label(root, text="ISEL File", bg=BG, fg=FG).pack(pady=5)

    entry = tk.Entry(
        root,
        textvariable=input_var,
        width=60,
        bg=BTN,
        fg=FG,
        insertbackground=FG
    )
    entry.pack()
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind("<<Drop>>", drop)

    tk.Button(root, text="Browse", command=browse_input).pack(pady=5)

    tk.Button(
        root,
        text="Convert",
        bg="#3a7afe",
        fg="white",
        command=convert
    ).pack(pady=10)

    logbox = tk.Text(
        root,
        height=9,
        bg="#121212",
        fg="#00ff88",
        insertbackground="white"
    )
    logbox.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
