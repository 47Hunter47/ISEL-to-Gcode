"""Sync APP_VERSION from the GitHub release tag into version.py and version_info.txt.

Called by CI on release events. On manual runs (no tag) it leaves versions untouched.
"""
import os
import re
import sys

tag = os.environ.get("GITHUB_REF_NAME", "")
m = re.match(r"^v?(\d+(?:\.\d+)*)$", tag)
if not m:
    print(f"No release tag in GITHUB_REF_NAME ({tag!r}); skipping version sync.")
    sys.exit(0)

ver = m.group(1)
parts = (ver.split(".") + ["0", "0", "0", "0"])[:4]

# version.py - imported by the GUI at runtime
with open("version.py", "w") as f:
    f.write(f'APP_VERSION = "{parts[0]}.{parts[1]}"\n')

# version_info.txt - PyInstaller EXE metadata (kept consistent)
try:
    t = open("version_info.txt").read()
except FileNotFoundError:
    t = None
if t is not None:
    t = re.sub(r"filevers=\([\d, ]+\)", f"filevers=({', '.join(parts)})", t)
    t = re.sub(r"prodvers=\([\d, ]+\)", f"prodvers=({', '.join(parts)})", t)
    full = ".".join(parts[:3])
    t = re.sub(r"StringStruct\('FileVersion', '[^']*'\)", f"StringStruct('FileVersion', '{full}')", t)
    t = re.sub(r"StringStruct\('ProductVersion', '[^']*'\)", f"StringStruct('ProductVersion', '{full}')", t)
    open("version_info.txt", "w").write(t)

print(f"Version synced to {parts[0]}.{parts[1]}")
