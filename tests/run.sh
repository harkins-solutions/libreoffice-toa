#!/usr/bin/env bash
# Build, install the extension, and run the tests against it.
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
LO="${LO_PROGRAM:-$HOME/.local/opt/lo/root/opt/libreoffice26.8/program}"
export LD_LIBRARY_PATH="$HOME/.local/opt/lo/root/usr/lib/x86_64-linux-gnu:$LO"
export PATH="$LO:$PATH"

python3 "$ROOT/build.py"
"$LO/unopkg" add -f "$ROOT/dist/table-of-authorities.oxt" 2>&1 | grep -v "platform" || true
exec "$LO/python" "$ROOT/tests/run_tests.py"
