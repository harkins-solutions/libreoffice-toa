#!/usr/bin/env bash
# Build the extension, install it, and run the tests against it.
#
# Finds LibreOffice in this order:
#   $LO_PROGRAM, a system install, then a local extracted copy.
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"

find_program_dir() {
  [ -n "$LO_PROGRAM" ] && { echo "$LO_PROGRAM"; return; }
  for candidate in /usr/lib/libreoffice/program \
                   /usr/lib64/libreoffice/program \
                   "$HOME"/.local/opt/lo/root/opt/libreoffice*/program \
                   /opt/libreoffice*/program; do
    [ -x "$candidate/soffice" ] && { echo "$candidate"; return; }
  done
  echo ""
}

LO="$(find_program_dir)"
[ -z "$LO" ] && { echo "no LibreOffice found; set LO_PROGRAM"; exit 1; }
export PATH="$LO:$PATH"
# A locally extracted copy needs its bundled libraries on the path; a system
# install already has them.
[ -d "$HOME/.local/opt/lo/root/usr/lib/x86_64-linux-gnu" ] && \
  export LD_LIBRARY_PATH="$HOME/.local/opt/lo/root/usr/lib/x86_64-linux-gnu:$LO"

# Debian and Ubuntu wire uno into the system python; the upstream builds ship
# their own. Either is fine, but they are not the same interpreter.
if [ -x "$LO/python" ]; then PY="$LO/python"; else PY="python3"; fi
UNOPKG="$LO/unopkg"; [ -x "$UNOPKG" ] || UNOPKG="unopkg"

echo "LibreOffice: $LO"
echo "python:      $PY"

python3 "$ROOT/build.py"
"$UNOPKG" add -f "$ROOT/dist/table-of-authorities.oxt" 2>&1 | grep -v -i platform || true
exec "$PY" "$ROOT/tests/run_tests.py"
