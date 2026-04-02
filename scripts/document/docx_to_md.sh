#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-to-md
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName DOCX
# @raycast.description DOCX转Markdown（使用markitdown）- 支持多选

source "$(dirname "$0")/../../lib/common.sh"
log_script_usage "docx/to_md" "docx"

# ===== 检查依赖 =====
HAS_MARKITDOWN=false
HAS_DOXX=false

if command -v markitdown &> /dev/null; then
    HAS_MARKITDOWN=true
fi

if command -v doxx &> /dev/null; then
    HAS_DOXX=true
fi

if [ "$HAS_MARKITDOWN" = false ] && [ "$HAS_DOXX" = false ]; then
    show_error "必须安装 markitdown (pip install markitdown) 或 doxx"
    exit 1
fi

# ===== 获取文件列表（优先命令行参数，否则 Finder 选择）=====
if [ $# -gt 0 ]; then
    # 命令行参数：支持相对路径和绝对路径
    SELECTED_FILES=""
    for arg in "$@"; do
        if [[ "$arg" = /* ]]; then
            file_path="$arg"
        else
            file_path="$(pwd)/$arg"
        fi
        if [ -n "$SELECTED_FILES" ]; then
            SELECTED_FILES="$SELECTED_FILES"$'\n'"$file_path"
        else
            SELECTED_FILES="$file_path"
        fi
    done
else
    SELECTED_FILES=$(get_finder_selection_multiple)
fi

if [ -z "$SELECTED_FILES" ]; then
    show_error "没有在 Finder 中选择任何文件，也没有传入命令行参数"
    exit 1
fi

# ===== 转换单个文件 =====
convert_docx_to_md() {
    local file="$1"
    local output_file="${file%.docx}.md"
    
    # 优先使用 markitdown
    if [ "$HAS_MARKITDOWN" = true ]; then
        if markitdown "$file" > "$output_file" 2>/dev/null; then
            return 0
        fi
    fi
    
    # 尝试 doxx
    if [ "$HAS_DOXX" = true ]; then
        if doxx --export markdown "$file" > "$output_file" 2>/dev/null; then
            return 0
        fi
    fi
    
    [ -f "$output_file" ] && rm -f "$output_file"
    return 1
}

# ===== 统计变量 =====
success_count=0
fail_count=0
skip_count=0
total_count=0

# ===== 处理每个文件 =====
while IFS= read -r SELECTED_FILE; do
    [ -z "$SELECTED_FILE" ] && continue
    
    ((total_count++))
    
    if ! check_file_extension "$SELECTED_FILE" "docx"; then
        show_warning "跳过非 DOCX 文件: $(basename "$SELECTED_FILE")"
        ((skip_count++))
        continue
    fi
    
    if [ ! -f "$SELECTED_FILE" ]; then
        show_error "文件不存在: $(basename "$SELECTED_FILE")"
        ((fail_count++))
        continue
    fi
    
    md_file="${SELECTED_FILE%.docx}.md"
    if [ -f "$md_file" ]; then
        show_warning "输出文件已存在，跳过: $(basename "$md_file")"
        ((skip_count++))
        continue
    fi
    
    show_processing "转换: $(basename "$SELECTED_FILE")"
    
    if convert_docx_to_md "$SELECTED_FILE"; then
        show_success "已转换: $(basename "$SELECTED_FILE") -> $(basename "$md_file")"
        ((success_count++))
    else
        show_error "转换失败: $(basename "$SELECTED_FILE")"
        ((fail_count++))
    fi
    
    echo ""
done <<< "$SELECTED_FILES"

# ===== 显示统计信息 =====
echo "========================================"
echo "📊 处理完成统计"
echo "========================================"
echo "✅ 成功: $success_count 个文件"
[ $fail_count -gt 0 ] && echo "❌ 失败: $fail_count 个文件"
[ $skip_count -gt 0 ] && echo "⏭️  跳过: $skip_count 个文件"
echo "📁 总计: $total_count 个文件"

[ $fail_count -gt 0 ] && exit 1
exit 0
