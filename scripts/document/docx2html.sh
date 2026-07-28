#!/usr/bin/env bash
# /docx view · 任意 docx → 本地富 HTML 预览（pandoc + 中文优化 CSS）
# usage: docx2html.sh <docx 路径> [输出目录]
#        DOCX_VIEW_NO_OPEN=1 抑制自动 open（CI / agent 场景必设）
set -euo pipefail

SRC="${1:?usage: docx2html.sh <docx> [out_dir]}"
[[ ! -f "$SRC" ]] && { echo "源文件不存在: $SRC" >&2; exit 1; }
[[ "${SRC##*.}" != "docx" ]] && { echo "需要 .docx 后缀，收到: $SRC" >&2; exit 1; }

# 解析绝对路径
SRC_ABS="$(cd "$(dirname "$SRC")" && pwd)/$(basename "$SRC")"
SRC_DIR="$(dirname "$SRC_ABS")"
BASENAME="$(basename "$SRC" .docx)"
OUT_DIR="${2:-$SRC_DIR/preview}"

mkdir -p "$OUT_DIR"
# 绝对化：下面要 cd 进去（pandoc --extract-media 落 cwd），相对 OUT_DIR 会二次拼接
OUT_DIR="$(cd "$OUT_DIR" && pwd)"
cd "$OUT_DIR"

# CSS 与本脚本同级 assets/（2026-07-27 从退役 skill docx-view 并入 doctools）
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CSS="$SCRIPT_DIR/assets/zh-cn.css"
[[ ! -f "$CSS" ]] && { echo "CSS 缺失: $CSS" >&2; exit 1; }

# 把 CSS 包成 <style> 注入 head（让 HTML 自包含，移动也能看）
HEAD_TMP="$(mktemp -t docxview-head.XXXXXX)"
trap 'rm -f "$HEAD_TMP"' EXIT
{
  printf '<style>\n'
  cat "$CSS"
  printf '\n</style>\n'
} > "$HEAD_TMP"

OUT_HTML="$OUT_DIR/$BASENAME.html"

pandoc "$SRC_ABS" \
  -f docx -t html5 \
  --standalone \
  --toc --toc-depth=3 \
  --extract-media=media \
  --metadata title="$BASENAME" \
  --metadata lang=zh-CN \
  -H "$HEAD_TMP" \
  -o "$OUT_HTML"

# 报告
# 注意：不要写 `ls media | wc -l` —— 无图 docx 时 ls 失败，set -euo pipefail 会让脚本
# 在这里静默 exit 1，报告不打印、也不 open（2026-07-27 实测到的老 bug）。
SIZE=$(ls -lh "$OUT_HTML" | awk '{print $5}')
if [[ -d media ]]; then
  MEDIA_COUNT=$(find media -type f | wc -l | tr -d ' ')
else
  MEDIA_COUNT=0
fi

echo "✅ 已生成: $OUT_HTML"
echo "   大小:   $SIZE"
echo "   media:  ${MEDIA_COUNT} 张图"
echo "   URL:    file://$OUT_HTML"

# 自动打开（除非 DOCX_VIEW_NO_OPEN=1）
if [[ "${DOCX_VIEW_NO_OPEN:-0}" != "1" ]] && command -v open >/dev/null; then
  open "$OUT_HTML"
fi
