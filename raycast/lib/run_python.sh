#!/bin/bash
# ============================================================
# Raycast 统一运行器
# 位置：raycast/lib/run_python.sh
# ============================================================
source "$(dirname "${BASH_SOURCE[0]}")/../../lib/common.sh"

run_python() {
    local script_name="$1"; shift
    # 注入 ~/Dev/tools/dev/lib 到 PYTHONPATH (doctools 多文件依赖 file_ops/display 等 canonical 模块)
    export PYTHONPATH="$HOME/Dev/tools/dev/lib${PYTHONPATH:+:$PYTHONPATH}"
    exec uv run --project "$PROJECT_ROOT" python3 "$SCRIPTS_DIR/$script_name" "$@"
}

run_streamlit() {
    local project_dir="$1"; shift
    local app_file="${1:-app.py}"; shift
    cd "$PROJECT_ROOT/projects/$project_dir" || exit 1
    exec streamlit run "$app_file" "$@"
}

run_shell() {
    local script_name="$1"; shift
    exec bash "$SCRIPTS_DIR/$script_name" "$@"
}
