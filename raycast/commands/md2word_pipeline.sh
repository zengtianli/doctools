#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title md2word-pipeline
# @raycast.description MD/DOCX 一键转 Word（套模板 → 文本修复 → 图名居中）
# @raycast.mode fullOutput
# @raycast.icon 📝
# @raycast.packageName Document Processing
source ~/Dev/devtools/lib/log_usage.sh

source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"

PYTHON="uv run --project $PROJECT_ROOT python3"
SCRIPT_DIR="$SCRIPTS_DIR/document"

# ── 处理单个文件 ─────────────────────────────────────────────
process_file() {
    local input_file="$1"
    local ext="${input_file##*.}"
    ext="$(echo "$ext" | tr '[:upper:]' '[:lower:]')"
    local dir="$(dirname "$input_file")"
    local stem="$(basename "$input_file" ".$ext")"

    if [[ "$ext" != "md" && "$ext" != "docx" ]]; then
        show_warning "跳过非 md/docx 文件: $(basename "$input_file")"
        return 1
    fi

    echo "── $(basename "$input_file") ──"

    # 最终输出文件名
    local final_file
    if [[ "$ext" == "md" ]]; then
        final_file="$dir/$stem.docx"
    else
        final_file="$dir/${stem}_styled.docx"
    fi

    # Step 1: 转换 / 套模板
    local step1_out
    if [[ "$ext" == "md" ]]; then
        echo "  ▶ 1/3 MD → DOCX"
        $PYTHON "$SCRIPT_DIR/md_docx_template.py" "$input_file"
        step1_out="$dir/$stem.docx"
    else
        echo "  ▶ 1/3 套模板样式"
        $PYTHON "$SCRIPT_DIR/docx_apply_template.py" "$input_file"
        step1_out="$dir/${stem}_styled.docx"
    fi

    if [[ ! -f "$step1_out" ]]; then
        show_error "Step 1 失败: $(basename "$input_file")"
        return 1
    fi

    # Step 2: 文本格式修复
    echo "  ▶ 2/3 文本修复"
    $PYTHON "$SCRIPT_DIR/docx_text_formatter.py" "$step1_out"
    local step2_out="$dir/$(basename "$step1_out" .docx)_fixed.docx"

    if [[ ! -f "$step2_out" ]]; then
        show_error "Step 2 失败: $(basename "$input_file")"
        return 1
    fi

    # Step 3: 图名居中
    echo "  ▶ 3/3 图名样式"
    $PYTHON "$SCRIPT_DIR/docx_apply_image_caption.py" "$step2_out"

    # 清理中间文件
    mv "$step2_out" "$final_file"
    if [[ "$step1_out" != "$final_file" && -f "$step1_out" ]]; then
        rm "$step1_out"
    fi
    [[ -f "${step2_out}.backup" ]] && rm "${step2_out}.backup"

    show_success "完成 → $(basename "$final_file")"
    echo ""
}

# ── 主流程 ───────────────────────────────────────────────────
files="$(get_finder_selection_multiple)"

if [[ -z "$files" ]]; then
    show_error "请在 Finder 中选中 .md 或 .docx 文件"
    exit 1
fi

# 统计文件数
file_count=0
while IFS= read -r f; do
    [[ -n "$f" ]] && ((file_count++))
done <<< "$files"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📝 MD2Word Pipeline — $file_count 个文件"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

success=0
fail=0
while IFS= read -r f; do
    [[ -z "$f" ]] && continue
    f="${f%/}"
    if [[ ! -f "$f" ]]; then
        show_warning "文件不存在: $f"
        ((fail++))
        continue
    fi
    if process_file "$f"; then
        ((success++))
    else
        ((fail++))
    fi
done <<< "$files"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
show_success "完成 $success 个${fail:+，失败 $fail 个}"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
