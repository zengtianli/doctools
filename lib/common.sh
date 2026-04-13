#!/bin/bash
# ============================================================
# 通用函数库 - Raycast 脚本公共库
# 位置：lib/common.sh
# ============================================================
# 使用说明: source "$(dirname "$0")/../lib/common.sh"
# ============================================================

# 防止重复 source
[[ -n "$_COMMON_SH_LOADED" ]] && return 0
_COMMON_SH_LOADED=1

# ===== 动态路径检测 =====
# LIB_DIR: 本文件所在目录
readonly _COMMON_SH_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# PROJECT_ROOT: 项目根目录 (lib 的上级)
readonly PROJECT_ROOT="$(cd "$_COMMON_SH_DIR/.." && pwd)"
# SCRIPTS_DIR: 脚本目录
readonly SCRIPTS_DIR="$PROJECT_ROOT/scripts"

# ===== 环境变量设置 =====
# 确保 uv 和 homebrew 的 bin 在 PATH 中（Raycast 非交互式 shell 不继承 profile）
for _bin_dir in "$HOME/.local/bin" "/opt/homebrew/bin"; do
    [ -d "$_bin_dir" ] && export PATH="$_bin_dir:$PATH"
done
# 加载用户环境变量（API keys 等）
_ENV_ZSH="$HOME/Documents/sync/zsh/config/env.zsh"
if [ -f "$_ENV_ZSH" ]; then
    source "$_ENV_ZSH"
fi
# 检测 Python 路径
if command -v python3 &>/dev/null; then
    readonly PYTHON_PATH="$(command -v python3)"
else
    readonly PYTHON_PATH="/usr/bin/env python3"
fi

# ===== 使用统计 =====
readonly USAGE_LOG="$HOME/.useful_scripts_usage.log"

log_script_usage() {
    local script_name="$1"
    local category="${2:-unknown}"
    local timestamp=$(date "+%Y-%m-%d %H:%M:%S")
    
    if [ ! -f "$USAGE_LOG" ]; then
        echo "timestamp,script_name,category" > "$USAGE_LOG"
    fi
    
    echo "$timestamp,$script_name,$category" >> "$USAGE_LOG" 2>/dev/null || true
}

# ===== Finder 操作函数 =====

# 获取 Finder 选中的多个文件（每行一个路径）
# 注意：返回换行分隔，配合 while IFS= read -r 使用
get_finder_selection_multiple() {
    osascript <<'EOF'
tell application "Finder"
    set selectedItems to selection as list
    if (count of selectedItems) = 0 then
        return ""
    end if
    set posixPaths to {}
    repeat with i from 1 to count of selectedItems
        set end of posixPaths to POSIX path of (item i of selectedItems as alias)
    end repeat
    set AppleScript's text item delimiters to "
"
    set pathsText to posixPaths as text
    set AppleScript's text item delimiters to ""
    return pathsText
end tell
EOF
}

# 获取 Finder 选中的单个文件
get_finder_selection_single() {
    osascript <<'EOF'
tell application "Finder"
    if (count of (selection as list)) > 0 then
        POSIX path of (item 1 of (selection as list) as alias)
    else
        ""
    end if
end tell
EOF
}

# 获取 Finder 当前目录
get_finder_current_dir() {
    osascript <<'EOF'
tell application "Finder"
    if (count of (selection as list)) > 0 then
        set firstItem to item 1 of (selection as list)
        if class of firstItem is folder then
            POSIX path of (firstItem as alias)
        else
            POSIX path of (container of firstItem as alias)
        end if
    else
        POSIX path of (insertion location as alias)
    end if
end tell
EOF
}

# ===== 文件检查函数 =====

# 检查文件扩展名（兼容 Bash 3.2）
check_file_extension() {
    local file="$1"
    local expected_ext="$2"
    local actual_ext="${file##*.}"
    
    [[ "$(echo "$actual_ext" | tr '[:upper:]' '[:lower:]')" == "$(echo "$expected_ext" | tr '[:upper:]' '[:lower:]')" ]]
}

# 验证文件路径安全性
validate_file_path() {
    local path="$1"
    if [[ "$path" =~ \.\./|\||\; ]]; then
        show_error "不安全的文件路径: $path"
        return 1
    fi
    return 0
}

# 检查文件是否存在
check_file_exists() {
    local file="$1"
    if [ ! -f "$file" ]; then
        show_error "文件不存在: $file"
        return 1
    fi
    return 0
}

# ===== 消息显示函数 =====

show_success() {
    echo "✅ $1"
}

show_error() {
    echo "❌ $1"
}

show_warning() {
    echo "⚠️ $1"
}

show_info() {
    echo "ℹ️ $1"
}

show_processing() {
    echo "🔄 $1"
}

# ===== 工具函数 =====

# 安全切换目录
safe_cd() {
    local target_dir="$1"
    if cd "$target_dir" 2>/dev/null; then
        return 0
    else
        show_error "无法进入目录: $target_dir"
        return 1
    fi
}

# 检查命令是否存在
check_command_exists() {
    local cmd="$1"
    if ! command -v "$cmd" &> /dev/null; then
        show_error "$cmd 未安装"
        return 1
    fi
    return 0
}

# ===== 获取脚本所在目录 =====
# 在 Raycast 脚本中使用：
# SCRIPT_DIR="$(get_script_dir)"
# 然后用 "$SCRIPT_DIR/xxx.py" 调用同目录的 Python 脚本

get_script_dir() {
    cd "$(dirname "${BASH_SOURCE[1]}")" && pwd
}

# ===== 在 Ghostty 中执行命令 =====

run_in_ghostty() {
    local command="$1"
    local command_escaped=$(printf "%s" "$command" | sed 's/"/\\"/g')
    
    osascript <<EOF
tell application "Ghostty"
    activate
    tell application "System Events"
        keystroke "n" using command down
    end tell
end tell
EOF
    
    sleep 1
    
    osascript <<EOF
tell application "Ghostty"
    activate
    delay 0.2
    set the clipboard to "$command_escaped"
    tell application "System Events"
        keystroke "v" using command down
        delay 0.1
        key code 36
    end tell
end tell
EOF
}
