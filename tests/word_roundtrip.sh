#!/usr/bin/env bash
# Round trip a marked document through real Microsoft Word. WSL + Word only.
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
LO="${LO_PROGRAM:-$HOME/.local/opt/lo/root/opt/libreoffice26.8/program}"
export LD_LIBRARY_PATH="$HOME/.local/opt/lo/root/usr/lib/x86_64-linux-gnu:$LO"
export PATH="$LO:$PATH"
command -v powershell.exe >/dev/null || { echo "not WSL; skipping"; exit 0; }
python3 "$ROOT/build.py" >/dev/null
"$LO/unopkg" add -f "$ROOT/dist/table-of-authorities.oxt" 2>&1 | grep -v platform || true
exec "$LO/python" "$ROOT/tests/word_roundtrip.py"
