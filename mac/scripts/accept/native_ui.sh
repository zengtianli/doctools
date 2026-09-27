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
"$HOME/Dev/.venv/bin/python" - "$TASK_BUILD/Build/Products/Release/DocTools.app/Contents/MacOS/DocTools" <<'PY'
import subprocess
import sys
result = subprocess.run([sys.argv[1], "--ui-self-test"], timeout=60)
raise SystemExit(result.returncode)
PY
