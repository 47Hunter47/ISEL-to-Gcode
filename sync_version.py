"""Release versiyon senkronizasyonu.

Her release'de:
- Tag bir versiyonsa (v2.1, 2.1) o kullanilir.
- Tag versiyon degilse, main'deki guncel versiyon otomatik artirilir (2.0 -> 2.1).
- workflow_dispatch (manuel calistirma) icin versiyon artirilmaz.
- version.py + README badge main'e GitHub API ile commit edilir.
- .version dosyasina hedef versiyon yazilir (artifact adi icin).
"""
import base64
import json
import os
import re
import sys
import urllib.request

API = f"https://api.github.com/repos/{os.environ['GITHUB_REPOSITORY']}"
TOKEN = os.environ["GITHUB_TOKEN"]


def api(path, method="GET", data=None):
    """GitHub Contents API cagrisi."""
    req = urllib.request.Request(
        API + path,
        data=json.dumps(data).encode() if data is not None else None,
        headers={
            "Authorization": f"token {TOKEN}",
            "Accept": "application/vnd.github+json",
        },
        method=method,
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def read_remote(path):
    """Dosya icerigi + sha dondur (main)."""
    d = api(f"/contents/{path}")
    return base64.b64decode(d["content"]).decode("utf-8"), d["sha"]


def write_remote(path, content, sha, message):
    """Dosyayi main'e commit et."""
    api(f"/contents/{path}", method="PUT", data={
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode(),
        "sha": sha,
        "branch": "main",
    })


def parse_version(s):
    """Metinden ilk sayisal versiyonu cikar (2.0, v2.1 -> '2.1')."""
    m = re.search(r"(\d+(?:\.\d+)*)", s or "")
    return m.group(1) if m else None


def next_version(cur):
    """Son bileseni artir: 2.0 -> 2.1, 2.1.3 -> 2.1.4."""
    parts = cur.split(".")
    parts[-1] = str(int(parts[-1]) + 1)
    return ".".join(parts)


def main():
    event = os.environ.get("GITHUB_EVENT_NAME", "")
    tag = os.environ.get("GITHUB_REF_NAME", "")

    vpy, vpy_sha = read_remote("version.py")
    cur = parse_version(vpy)
    if not cur:
        sys.exit(f"ERROR: main'deki version.py icinde versiyon bulunamadi: {vpy!r}")

    # Hedef versiyonu belirle
    if event == "workflow_dispatch":
        target = cur  # manuel calistirmada artis yok
        print(f"workflow_dispatch -> versiyon degistirilmedi: {target}")
    else:
        tag_ver = parse_version(tag)
        if tag_ver:
            target = tag_ver
            print(f"Tag versiyonu kullaniliyor: {tag} -> {target}")
        else:
            target = next_version(cur)
            print(f"Tag versiyon degil ({tag!r}) -> otomatik artis: {cur} -> {target}")

    # Build icin yerel dosyalar
    new_vpy = f'APP_VERSION = "{target}"\n'
    with open("version.py", "w") as f:
        f.write(new_vpy)

    readme, readme_sha = read_remote("README.md")
    new_readme = re.sub(
        r"version-\d+(?:\.\d+)*-(\w+)", f"version-{target}-\\1", readme
    )
    with open("README.md", "w") as f:
        f.write(new_readme)

    # main'e commit et (degisiklik varsa)
    if new_vpy != vpy:
        write_remote("version.py", new_vpy, vpy_sha,
                     f"chore: sync version to {target}")
        print(f"version.py main'e commit edildi: {target}")
    else:
        print("version.py zaten guncel")

    if new_readme != readme:
        write_remote("README.md", new_readme, readme_sha,
                     f"chore: sync version badge to {target}")
        print(f"README.md main'e commit edildi: {target}")
    else:
        print("README.md zaten guncel")

    # Artifact adi icin versiyonu yaz
    with open(".version", "w") as f:
        f.write(target)

    print(f"Versiyon senkronize edildi: {target}")


if __name__ == "__main__":
    main()
