"""Sync APP_VERSION from the GitHub release tag into version.py and version_info.txt.

Runs in CI on release events. Reads GITHUB_REF_NAME (the release tag), rewrites
version.py and version_info.txt to match, and prints the new version.
On manual runs (no release tag) it leaves everything untouched.
"""
import os
import re
import sys


def main():
    tag = os.environ.get("GITHUB_REF_NAME", "")
    m = re.match(r"^v?(\d+(?:\.\d+)*)$", tag)
    if not m:
        print(f"No release tag in GITHUB_REF_NAME ({tag!r}); skipping version sync.")
        return

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


if __name__ == "__main__":
    main()
