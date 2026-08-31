# ISEL to G-code Converter

Windows GUI application that converts ISEL CAM files into standard G-code (NC) programs for CNC machines.

![version](https://img.shields.io/badge/version-2.0-blue)
![python](https://img.shields.io/badge/python-3.10+-green)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)

## Features

- **Dark-mode GUI** with drag & drop or file browser
- **Coordinate conversion** — ISEL integer coordinates (×1000) to mm
- **Feed conversion** — ISEL `VEL` values to G-code `F` feed rates
- **Two output modes:**
  - `G1 linearised` — arcs expanded into short line segments (numpy-accelerated when available)
  - `G2/G3 arc commands` — true arc output with full-circle detection (R format, split into two semicircles for controller compatibility)
- **Spindle handling** — `S... M03` from `SPINDLE CW`, with `M05`/`M30` program end
- **Toolpath preview** — top-down and isometric views with grid, pan, and zoom
- **Estimated cycle time** shown after conversion
- Single-file `.exe` (PyInstaller) — no Python installation required

## Usage

1. Download the latest release (`ISEL_to_Gcode.exe`) from the [Releases](https://github.com/47Hunter47/ISEL-to-Gcode/releases) page.
2. Run the executable.
3. Drag & drop your ISEL file onto the window (or click **Browse**).
4. Optionally enable **Use G2/G3 arc commands** if your controller supports arcs.
5. Click **Convert** and choose a save location (`.ngc`).
6. Open **Preview** to inspect the toolpath before running the program.

> ⚠ The estimated program time is an approximation and may differ from actual machine time depending on machine parameters.

## Output format

Programs start with a standard header and end with spindle stop:

```
N00001 G21
N00002 G17
N00003 G40
N00004 G49
N00005 G90
N00006 G94
N00007 G0 Z...        ← first Z safe retract
N00008 S1200 M03      ← spindle starts after the tool is clear
...
M05
M30
```

The first spindle command is emitted **after** the first Z safe retract, so the spindle starts turning only once the tool is clear of the workpiece. Subsequent spindle speed changes are emitted immediately.

## Building from source

Requires Python 3.10+ on Windows.

```bat
pip install pyinstaller tkinterdnd2 numpy
pyinstaller --onefile --windowed --icon=icon.ico --collect-all tkinterdnd2 --hidden-import=numpy --name ISEL_to_Gcode isel_to_gcode_gui.py
```

The executable is created in `dist\`.

A release build is also triggered automatically by GitHub Actions when a new release is published.
