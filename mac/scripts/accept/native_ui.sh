#!/bin/bash
set -euo pipefail
DIR="$(cd "$(dirname "$0")/../.." && pwd -P)"
source "$HOME/Dev/tools/dev/lib/tools/macapp/xcode_env.sh"
xcode_env_use macosx
cd "$DIR"
TASK_BUILD="$(mktemp -d "${TMPDIR:-/tmp}/dockit-native-accept.XXXXXX")"
trap 'rm -rf "$TASK_BUILD"' EXIT
# 共享构建通道负责仓外副本、排队与签名；75 原样交回 Chapter，不能记成通过或构建失败。
# 验收输出只给运行中的自检。构建仅继承工具链/系统定位和既有 lane 互斥参数，
# 不继承 shell 的凭证、Homebrew 编译器覆盖或工作树内的 SOP_OUT_DIR。
if "$HOME/Dev/.venv/bin/python" - "$DIR" "$TASK_BUILD" > "$TASK_BUILD/receipt.json" <<'PY'
import os
from pathlib import Path
import subprocess
import sys

keys = {"HOME", "PATH", "TMPDIR", "DEVELOPER_DIR", "LANG", "LC_ALL", "USER", "LOGNAME", "SHELL",
        "SIM_LANE_LOCK", "SIM_LANE_EXTRA_LOCK"}
environment = {key: value for key, value in os.environ.items() if key in keys}
lane = Path.home() / "Dev/tools/dev/lib/tools/macapp/ios/sim_lane.py"
result = subprocess.run([sys.executable, str(lane), "build", "--project-dir", sys.argv[1],
                         "--scheme", "DocTools", "--platform", "mac", "--configuration", "Release",
                         "--work-dir", sys.argv[2], "--json"], env=environment)
raise SystemExit(result.returncode)
PY
then
    :
else
    STATUS=$?
    cat "$TASK_BUILD/receipt.json" >&2
    exit "$STATUS"
fi
export SOP_OUT_DIR="${SOP_OUT_DIR:-$TASK_BUILD/results}"
mkdir -p "$SOP_OUT_DIR"
"$HOME/Dev/.venv/bin/python" - "$TASK_BUILD/receipt.json" "$TASK_BUILD" <<'PY'
import json
import os
import plistlib
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
    receipt = json.loads(open(sys.argv[1]).read())
    if receipt.get("ok") is not True:
        raise ValueError("Shared build did not return a successful receipt")
    app = receipt["app_path"]
    with open(os.path.join(app, "Contents/Info.plist"), "rb") as file:
        executable = plistlib.load(file)["CFBundleExecutable"]
    result = subprocess.run([os.path.join(app, "Contents/MacOS", executable), "--ui-self-test"],
                            timeout=60, capture_output=True, text=True)
except subprocess.TimeoutExpired as error:
    emit(error.stdout, sys.stdout)
    emit(error.stderr, sys.stderr)
    emit(traceback.format_exc(), sys.stderr)
    raise SystemExit(1)
except (OSError, ValueError, KeyError):
    emit(traceback.format_exc(), sys.stderr)
    raise SystemExit(1)
emit(result.stdout, sys.stdout)
emit(result.stderr, sys.stderr)
raise SystemExit(result.returncode)
PY
