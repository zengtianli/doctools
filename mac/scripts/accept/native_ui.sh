#!/bin/bash
set -euo pipefail
DIR="$(cd "$(dirname "$0")/../.." && pwd -P)"
source "$HOME/Dev/tools/dev/lib/tools/macapp/xcode_env.sh"
xcode_env_use macosx
cd "$DIR"
TASK_BUILD="$(mktemp -d "${TMPDIR:-/tmp}/dockit-native-accept.XXXXXX")"
trap 'rm -rf "$TASK_BUILD"' EXIT
if ! xcodebuild -project DocTools.xcodeproj -scheme DocTools -configuration Release \
    -derivedDataPath "$TASK_BUILD" CODE_SIGNING_ALLOWED=NO build > "$TASK_BUILD/build.log" 2>&1; then
    tail -80 "$TASK_BUILD/build.log" >&2
    exit 1
fi
export SOP_OUT_DIR="${SOP_OUT_DIR:-$TASK_BUILD/results}"
mkdir -p "$SOP_OUT_DIR"
"$HOME/Dev/.venv/bin/python" - "$TASK_BUILD/Build/Products/Release/DocTools.app/Contents/MacOS/DocTools" "$TASK_BUILD" <<'PY'
import os
import subprocess
import sys
import traceback

scratch_paths = sorted({sys.argv[2], os.path.normpath(sys.argv[2]), os.path.realpath(sys.argv[2])}, key=len, reverse=True)

def emit(value, stream):
    if not value:
        return
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="replace")
    for path in scratch_paths:
        value = value.replace(path, "<scratch>")
    stream.write(value.replace(os.path.expanduser("~"), "~"))
    stream.flush()

try:
    result = subprocess.run([sys.argv[1], "--ui-self-test"], timeout=60, capture_output=True, text=True)
except subprocess.TimeoutExpired as error:
    emit(error.stdout, sys.stdout)
    emit(error.stderr, sys.stderr)
    emit(traceback.format_exc(), sys.stderr)
    raise SystemExit(1)
except OSError:
    emit(traceback.format_exc(), sys.stderr)
    raise SystemExit(1)
emit(result.stdout, sys.stdout)
emit(result.stderr, sys.stderr)
raise SystemExit(result.returncode)
PY
