#!/usr/bin/env python3
"""Package src/ into dist/table-of-authorities.oxt.

An .oxt is a zip. There is no build step beyond that; the component is Python
and LibreOffice runs it from the package.
"""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
OUT = ROOT / "dist" / "table-of-authorities.oxt"


def build() -> Path:
    OUT.parent.mkdir(exist_ok=True)
    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as package:
        for path in sorted(SRC.rglob("*")):
            if path.is_file() and "__pycache__" not in path.parts:
                package.write(path, path.relative_to(SRC).as_posix())
    return OUT


if __name__ == "__main__":
    out = build()
    with zipfile.ZipFile(out) as package:
        names = package.namelist()
    print(f"built {out.relative_to(ROOT)} ({out.stat().st_size} bytes)")
    for name in sorted(names):
        print(f"  {name}")
    sys.exit(0)
