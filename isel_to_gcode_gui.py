import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os

SCALE = 1000.0
VEL_RATIO = 16.6667


def convert_file(input_path, output_path, log):
    line_no = 1

    def nline():
        nonlocal line_no
        s = f"N{line_no:05d} "
        line_no += 1
        return s

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

            #if line.startswith("GETTOOL"):
            #    tool = re.search(r"GETTOOL (\d+)", line).group(1)
            #    gcode.append(nline() + f"T{tool} M06")
            #    log("Takım değiştirildi")

            elif line.startswith("SPINDLE CW"):
                rpm = re.search(r"RPM(\d+)", line).group(1)
                gcode.append(nline() + f"S{rpm} M03")
                log(f"Spindle: {rpm} RPM")

            elif line.startswith("FASTABS"):
                c = parse_coord(line)
                cmd = "G0"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(nline() + cmd)

            elif line.startswith("MOVEABS"):
                c = parse_coord(line)
                cmd = "G1"
                for k, v in c.items():
                    cmd += f" {k}{v:.3f}"
                gcode.append(nline() + cmd)

            elif line.startswith("VEL"):
                vel = int(line.split()[1])
                feed = vel / VEL_RATIO
                gcode.append(nline() + f"F{feed:.0f}")
                log(f"Feed: F{feed:.0f}")

    gcode += [nline() + "M05", nline() + "M30"]

    with open(output_path, "w") as f:
        f.write("\n".join(gcode))


def run_gui():
    root = tk.Tk()
    root.title("ISEL → G-code Dönüştürücü")
    root.geometry("480x320")

    input_var = tk.StringVar()

    def browse_input():
        input_var.set(
            filedialog.askopenfilename(
                filetypes=[
                    ("ISEL NC Files", "*.nc"),
                    ("All Files", "*.*")
                ]
            )
        )

    def log(msg):
        logbox.insert(tk.END, msg + "\n")
        logbox.see(tk.END)

    def convert():
        if not input_var.get():
            messagebox.showerror("Hata", "Giriş dosyası seçilmedi")
            return

        in_path = input_var.get()
        base, _ = os.path.splitext(in_path)
        out_path = base + ".ngc"

        try:
            convert_file(in_path, out_path, log)
            messagebox.showinfo("Tamam", f"Dönüştürme tamamlandı:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Hata", str(e))

    tk.Label(root, text="ISEL (.nc) Dosyası").pack()
    tk.Entry(root, textvariable=input_var, width=50).pack()
    tk.Button(root, text="Gözat", command=browse_input).pack(pady=5)

    tk.Button(
        root,
        text="DÖNÜŞTÜR",
        bg="green",
        fg="white",
        command=convert
    ).pack(pady=10)

    logbox = tk.Text(root, height=8)
    logbox.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
