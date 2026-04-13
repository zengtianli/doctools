#!/bin/bash
# ============================================================
# Raycast 统一运行器
# 位置：raycast/lib/run_python.sh
# ============================================================
source "$(dirname "${BASH_SOURCE[0]}")/../../lib/common.sh"

run_python() {
    local script_name="$1"; shift
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
