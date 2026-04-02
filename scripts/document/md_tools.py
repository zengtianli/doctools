#!/usr/bin/env python3
"""
md_tools.py - Markdown 工具集

将 6 个独立的 Markdown 处理脚本合并为一个统一入口。

子命令:
    format      文本格式自动修复（引号/标点/单位）
    merge       合并多个 Markdown 文件为一个
    split       按一级标题拆分 Markdown 文件
    strip       删除 Markdown 文件中所有 blockquote
    to-docx     Markdown 转 Docx（Pandoc 版）
    to-html     Markdown 渲染为 HTML 并在浏览器中打开
    frontmatter 批量生成 YAML frontmatter（LLM）

用法:
    python3 md_tools.py <subcommand> [args...]
    python3 md_tools.py format file.md
    python3 md_tools.py merge a.md b.md
    python3 md_tools.py split input.md
    python3 md_tools.py strip md_final/ --fix
    python3 md_tools.py to-docx input.md -o output.docx
    python3 md_tools.py to-html file.md
    python3 md_tools.py frontmatter docs/ --dry-run
"""

import argparse
import os
import re
import subprocess
import sys
import tempfile
import time
import webbrowser
from pathlib import Path

# ── lib path setup ──────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))
from display import show_error, show_info, show_processing, show_success, show_warning
from file_ops import (
    check_file_extension,
    fatal_error,
    validate_input_file,
)
from finder import get_input_files
from progress import ProgressTracker
from dockit.text import fix_punctuation, fix_quotes, fix_units

# ── version info ────────────────────────────────────────────────────
SCRIPT_VERSION = "3.0.0"
SCRIPT_AUTHOR = "tianli"
SCRIPT_UPDATED = "2026-03-25"


def format_process_file(input_file):
    """处理单个文件的格式修复"""
    input_path = Path(input_file)

    if not input_path.exists():
        show_error(f"文件不存在 - {input_file}")
        return False

    output_path = input_path.parent / f"{input_path.stem}_fixed{input_path.suffix}"

    try:
        show_processing(f"正在读取文件: {input_path.name}")
        with open(input_path, encoding="utf-8") as f:
            content = f.read()

        show_processing("正在处理引号...")
        fixed_content, quote_count, _ = fix_quotes(content)

        show_processing("正在处理标点符号...")
        fixed_content, punct_count = fix_punctuation(fixed_content)

        show_processing("正在转换单位...")
        fixed_content, unit_count = fix_units(fixed_content)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(fixed_content)

        show_success("处理完成!")
        print(f"   - 共替换了 {quote_count} 个引号")
        print(f"   - 共替换了 {punct_count} 个标点符号")
        print(f"   - 共转换了 {unit_count} 个单位")
        print(f"   - 输出文件: {output_path.name}")

        return True

    except Exception as e:
        show_error(f"处理失败: {e}")
        return False


def cmd_format(args):
    """format 子命令入口"""
    files = args.files
    if not files:
        files = get_input_files([], expected_ext="md")

    if not files:
        fatal_error(
            "缺少文件名参数\n\n"
            "使用方法:\n"
            "    python3 md_tools.py format 文件名.md\n"
            "    python3 md_tools.py format file1.md file2.md\n"
            "    或在 Finder 中选择 .md 文件后运行"
        )

    print("=" * 50)
    print("文本格式自动修复工具")
    print("=" * 50)

    tracker = ProgressTracker()

    for file_path in files:
        print(f"\n处理文件: {Path(file_path).name}")
        success = format_process_file(str(file_path))
        if success:
            tracker.add_success()
        else:
            tracker.add_error()

    print("\n" + "=" * 50)
    tracker.show_summary("文件处理")


# ════════════════════════════════════════════════════════════════════
#  merge - 合并多个 Markdown 文件
# ════════════════════════════════════════════════════════════════════

def merge_md_files(md_files: list, output_file: Path):
    """合并多个 Markdown 文件"""
    tracker = ProgressTracker()

    md_files = sorted(md_files)

    show_info(f"准备合并 {len(md_files)} 个 Markdown 文件")

    try:
        with open(output_file, "w", encoding="utf-8") as f_out:
            for i, md_file in enumerate(md_files, 1):
                md_path = Path(md_file)

                if not validate_input_file(md_path):
                    tracker.add_skip()
                    continue

                if not check_file_extension(md_path, "md"):
                    show_warning(f"跳过非 Markdown 文件: {md_path.name}")
                    tracker.add_skip()
                    continue

                show_info(f"处理 ({i}/{len(md_files)}): {md_path.name}")

                with open(md_path, encoding="utf-8") as f_in:
                    content = f_in.read()
                    f_out.write(content)
                    if i < len(md_files):
                        f_out.write("\n\n")

                tracker.add_success()

        show_success(f"合并完成，已保存为: {output_file.name}")
        tracker.show_summary("文件合并")

    except Exception as e:
        fatal_error(f"合并失败: {e}")


def cmd_merge(args):
    """merge 子命令入口"""
    files = args.files
    if not files:
        files = get_input_files([], expected_ext="md")

    if not files:
        fatal_error("请提供至少一个 Markdown 文件，或在 Finder 中选择文件后运行")

    # 判断最后一个参数是否为输出文件
    if args.output:
        output_file = Path(args.output)
        md_files = files
    elif files[-1].endswith(".md") and len(files) > 1 and not Path(files[-1]).exists():
        output_file = Path(files[-1])
        md_files = files[:-1]
    else:
        first_file = Path(files[0]).resolve()
        output_file = first_file.parent / "merged.md"
        md_files = files

    merge_md_files(md_files, output_file)


# ════════════════════════════════════════════════════════════════════
#  split - 按一级标题拆分 Markdown 文件
# ════════════════════════════════════════════════════════════════════

def split_slugify(title: str) -> str:
    """将标题转换为文件名友好格式"""
    title = re.sub(r"^#*\s*\d*\.?\s*", "", title)
    slug = title.lower().strip()
    slug = re.sub(r"[^\w\s-]", "", slug)
    slug = re.sub(r"[\s-]+", "_", slug)
    return slug[:50]


def split_markdown(input_path: Path) -> list[tuple[str, str]]:
    """
    拆分 Markdown 文件
    返回: [(filename, content), ...]
    """
    content = input_path.read_text(encoding="utf-8")

    pattern = r"^(# .+)$"
    parts = re.split(pattern, content, flags=re.MULTILINE)

    results = []
    idx = 0

    # 第一部分: # 之前的内容（标题、摘要等）
    if parts[0].strip():
        results.append((f"{idx:02d}_title.md", parts[0].strip()))
        idx += 1

    # 后续部分: 成对出现（标题, 内容）
    for i in range(1, len(parts), 2):
        if i + 1 < len(parts):
            heading = parts[i]
            body = parts[i + 1] if i + 1 < len(parts) else ""

            slug = split_slugify(heading)
            filename = f"{idx:02d}_{slug}.md"

            full_content = f"{heading}\n{body}".strip()
            results.append((filename, full_content))
            idx += 1

    return results


def cmd_split(args):
    """split 子命令入口"""
    input_file = args.input
    if not input_file:
        files = get_input_files([], expected_ext="md")
        if files:
            input_file = files[0]

    if not input_file:
        fatal_error("请提供一个 Markdown 文件，或在 Finder 中选择文件后运行")

    input_path = Path(input_file).resolve()

    if not validate_input_file(input_path):
        sys.exit(1)
    if not check_file_extension(input_path, "md"):
        fatal_error(f"不是 Markdown 文件: {input_path.name}")

    # 输出目录
    output_dir = input_path.parent / f"{input_path.stem}_split"
    output_dir.mkdir(exist_ok=True)

    parts = split_markdown(input_path)

    show_info(f"输出目录: {output_dir}")
    show_info(f"拆分为 {len(parts)} 个文件:")

    for filename, content in parts:
        output_path = output_dir / filename
        output_path.write_text(content, encoding="utf-8")
        lines = content.count("\n") + 1
        show_success(f"{filename} ({lines} 行)")

    show_success("拆分完成!")


# ════════════════════════════════════════════════════════════════════
#  strip - 删除 Markdown 文件中所有 blockquote
# ════════════════════════════════════════════════════════════════════

def strip_blockquotes(text: str) -> str:
    """删除所有 blockquote 行，并清理残留的连续空行"""
    lines = text.split("\n")
    result = []
    in_code = False

    for line in lines:
        stripped = line.strip()

        # 跟踪代码块
        if stripped.startswith("```"):
            in_code = not in_code
            result.append(line)
            continue

        # 代码块内不处理
        if in_code:
            result.append(line)
            continue

        # 跳过 blockquote 行
        if stripped.startswith(">"):
            continue

        result.append(line)

    # 清理连续空行（最多保留 1 个）
    cleaned = []
    prev_empty = False
    for line in result:
        is_empty = line.strip() == ""
        if is_empty and prev_empty:
            continue
        cleaned.append(line)
        prev_empty = is_empty

    return "\n".join(cleaned)


def strip_count_blockquotes(text: str) -> int:
    """统计 blockquote 行数（排除代码块内）"""
    lines = text.split("\n")
    count = 0
    in_code = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if stripped.startswith(">"):
            count += 1
    return count


def strip_process_file(filepath: Path, do_fix: bool, output_dir: Path | None) -> dict:
    """处理单个文件"""
    text = filepath.read_text(encoding="utf-8")
    bq_count = strip_count_blockquotes(text)

    if bq_count > 0:
        print(f"  {filepath.name}: {bq_count} 行 blockquote")
    else:
        print(f"  {filepath.name}: 无 blockquote")

    if do_fix and bq_count > 0:
        fixed_text = strip_blockquotes(text)

        if output_dir:
            output_dir.mkdir(parents=True, exist_ok=True)
            out_path = output_dir / filepath.name
        else:
            out_path = filepath

        out_path.write_text(fixed_text, encoding="utf-8")
        show_success(f"已删除 {bq_count} 行 -> {out_path}")

    return {"file": filepath.name, "blockquotes": bq_count}


def cmd_strip(args):
    """strip 子命令入口"""
    input_path = Path(args.input)
    if not input_path.exists():
        show_error(f"路径不存在: {input_path}")
        sys.exit(1)

    if input_path.is_dir():
        md_files = sorted(input_path.glob("*.md"))
        if not md_files:
            show_error(f"目录中没有 .md 文件: {input_path}")
            sys.exit(1)
        show_info(f"发现 {len(md_files)} 个 MD 文件")
    else:
        md_files = [input_path]

    output_dir = Path(args.output_dir) if args.output_dir else None

    all_stats = []
    for f in md_files:
        stats = strip_process_file(f, args.fix, output_dir)
        all_stats.append(stats)

    total_bq = sum(s["blockquotes"] for s in all_stats)
    files_with_bq = sum(1 for s in all_stats if s["blockquotes"] > 0)

    if len(all_stats) > 1:
        print(f"\n合计: {total_bq} 行 blockquote（{files_with_bq}/{len(all_stats)} 个文件）")

    if total_bq > 0 and not args.fix:
        print("\n使用 --fix 执行删除")
        sys.exit(1)


# ════════════════════════════════════════════════════════════════════
#  to-docx - Markdown 转 Docx（Pandoc 版）
# ════════════════════════════════════════════════════════════════════

# 默认模板路径
DOCX_DEFAULT_TEMPLATE = str(Path.home() / "Downloads/归档/其他文档/template.docx")


def docx_get_finder_selection():
    """获取 Finder 选中的文件"""
    script = """
    tell application "Finder"
        set sel to selection
        if (count of sel) > 0 then
            return POSIX path of (item 1 of sel as alias)
        end if
    end tell
    """
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    return result.stdout.strip()


def docx_convert(md_path, output_path=None, template_path=None):
    """用 Pandoc 将 Markdown 转换为 Docx"""
    md_path = Path(md_path).resolve()
    if not md_path.exists():
        show_error(f"文件不存在: {md_path}")
        sys.exit(1)

    if output_path is None:
        output_path = md_path.with_suffix(".docx")
    else:
        output_path = Path(output_path).resolve()

    # 构建 Pandoc 命令
    cmd = [
        "pandoc",
        str(md_path),
        "-o",
        str(output_path),
        "--from",
        "markdown",
        "--to",
        "docx",
    ]

    # 使用模板（reference-doc）
    tpl = template_path or DOCX_DEFAULT_TEMPLATE
    if tpl and Path(tpl).exists():
        cmd += ["--reference-doc", str(tpl)]
        show_info(f"模板: {Path(tpl).name}")

    show_processing(f"输入: {md_path.name}")

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        show_error(f"Pandoc 转换失败:\n{result.stderr}")
        sys.exit(1)

    show_success(f"输出: {output_path}")
    return str(output_path)


def cmd_to_docx(args):
    """to-docx 子命令入口"""
    md_path = args.input
    if not md_path:
        finder_file = docx_get_finder_selection()
        if finder_file and finder_file.endswith(".md"):
            md_path = finder_file
            show_info(f"从 Finder 获取: {os.path.basename(finder_file)}")
        else:
            show_error("请在 Finder 中选择一个 .md 文件")
            sys.exit(1)

    docx_convert(md_path, args.output, args.template)


# ════════════════════════════════════════════════════════════════════
#  to-html - Markdown 渲染为 HTML 并在浏览器中打开
# ════════════════════════════════════════════════════════════════════

# 内联 CSS，不依赖外部文件
HTML_CSS = """
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
    line-height: 1.6; color: #1f2937; background: #f9fafb;
    max-width: 900px; margin: 0 auto; padding: 2rem;
  }
  h1 { font-size: 1.8rem; border-bottom: 2px solid #e5e7eb; padding-bottom: 0.5rem; margin: 1.5rem 0 1rem; color: #111827; }
  h2 { font-size: 1.4rem; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3rem; margin: 1.3rem 0 0.8rem; color: #1f2937; }
  h3 { font-size: 1.15rem; margin: 1rem 0 0.5rem; color: #374151; }
  p { margin: 0.5rem 0; }
  a { color: #2563eb; text-decoration: none; }
  a:hover { text-decoration: underline; }
  code { background: #f3f4f6; padding: 0.15rem 0.4rem; border-radius: 4px; font-size: 0.9em; }
  pre { background: #1f2937; color: #f9fafb; padding: 1rem; border-radius: 8px; overflow-x: auto; margin: 0.8rem 0; }
  pre code { background: none; padding: 0; color: inherit; }
  blockquote { border-left: 4px solid #3b82f6; padding: 0.5rem 1rem; margin: 0.8rem 0; background: #eff6ff; color: #1e40af; }
  table { border-collapse: collapse; width: 100%; margin: 0.8rem 0; }
  th, td { border: 1px solid #d1d5db; padding: 0.5rem 0.75rem; text-align: left; }
  th { background: #f3f4f6; font-weight: 600; }
  tr:nth-child(even) { background: #f9fafb; }
  ul, ol { padding-left: 1.5rem; margin: 0.5rem 0; }
  li { margin: 0.2rem 0; }
  hr { border: none; border-top: 1px solid #e5e7eb; margin: 1.5rem 0; }
  .nav { background: #1f2937; color: white; padding: 0.8rem 1.5rem; margin: -2rem -2rem 2rem; }
  .nav a { color: #93c5fd; margin-right: 1rem; }
  .file-list { list-style: none; padding: 0; }
  .file-list li { padding: 0.5rem 0; border-bottom: 1px solid #e5e7eb; }
  .file-list .date { color: #6b7280; font-size: 0.85em; margin-left: 0.5rem; }
  .emoji { font-size: 1.1em; }
  input[type="checkbox"] { margin-right: 0.3rem; }
</style>
"""


def html_md_to_html_simple(md_text):
    """简易 MD -> HTML 转换，不依赖第三方库"""
    html = md_text

    # 先处理代码块
    code_blocks = []

    def save_code(m):
        code_blocks.append(m.group(1))
        return f"__CODE_BLOCK_{len(code_blocks) - 1}__"

    html = re.sub(r"```[\w]*\n(.*?)```", save_code, html, flags=re.DOTALL)

    inline_codes = []

    def save_inline(m):
        inline_codes.append(m.group(1))
        return f"__INLINE_CODE_{len(inline_codes) - 1}__"

    html = re.sub(r"`([^`]+)`", save_inline, html)

    # 表格
    def convert_table(m):
        lines = m.group(0).strip().split("\n")
        rows = []
        for i, line in enumerate(lines):
            line = line.strip()
            if not line.startswith("|"):
                continue
            cells = [c.strip() for c in line.split("|")[1:-1]]
            if i == 1 and all(re.match(r"^[-:]+$", c) for c in cells):
                continue  # separator row
            tag = "th" if i == 0 else "td"
            row = "".join(f"<{tag}>{c}</{tag}>" for c in cells)
            rows.append(f"<tr>{row}</tr>")
        return f"<table>{''.join(rows)}</table>"

    html = re.sub(r"(\|.+\|[\n\r]+)+", convert_table, html)

    # Headers
    html = re.sub(r"^### (.+)$", r"<h3>\1</h3>", html, flags=re.MULTILINE)
    html = re.sub(r"^## (.+)$", r"<h2>\1</h2>", html, flags=re.MULTILINE)
    html = re.sub(r"^# (.+)$", r"<h1>\1</h1>", html, flags=re.MULTILINE)

    # Bold, italic
    html = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html)
    html = re.sub(r"\*(.+?)\*", r"<em>\1</em>", html)

    # Checkbox
    html = re.sub(r"- \[x\]", r'<li><input type="checkbox" checked disabled>', html)
    html = re.sub(r"- \[ \]", r'<li><input type="checkbox" disabled>', html)

    # Lists
    html = re.sub(r"^- (.+)$", r"<li>\1</li>", html, flags=re.MULTILINE)

    # Blockquote
    html = re.sub(r"^> (.+)$", r"<blockquote>\1</blockquote>", html, flags=re.MULTILINE)

    # HR
    html = re.sub(r"^---+$", r"<hr>", html, flags=re.MULTILINE)

    # Links
    html = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', html)

    # Paragraphs (lines that aren't already HTML)
    lines = html.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if (
            stripped
            and not stripped.startswith("<")
            and not stripped.startswith("__CODE")
            and not stripped.startswith("__INLINE")
        ):
            result.append(f"<p>{line}</p>")
        else:
            result.append(line)
    html = "\n".join(result)

    # Restore code blocks
    for i, block in enumerate(code_blocks):
        html = html.replace(f"__CODE_BLOCK_{i}__", f"<pre><code>{block}</code></pre>")
    for i, code in enumerate(inline_codes):
        html = html.replace(f"__INLINE_CODE_{i}__", f"<code>{code}</code>")

    return html


def html_render_file(md_path, output_dir=None):
    """渲染单个 MD 文件为 HTML"""
    md_path = Path(md_path)
    md_text = md_path.read_text(encoding="utf-8")

    title = md_path.stem
    title_match = re.search(r"^# (.+)$", md_text, re.MULTILINE)
    if title_match:
        title = title_match.group(1)

    body = html_md_to_html_simple(md_text)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    {HTML_CSS}
</head>
<body>
    <div class="nav">
        <strong>{md_path.name}</strong>
        <span style="float:right;color:#9ca3af;font-size:0.85em">{md_path.parent}</span>
    </div>
    {body}
</body>
</html>"""

    if output_dir:
        out = Path(output_dir) / f"{md_path.stem}.html"
    else:
        out = Path(tempfile.mkdtemp()) / f"{md_path.stem}.html"

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    return out


def html_render_directory(dir_path, output_dir=None):
    """渲染整个目录为带索引的 HTML"""
    dir_path = Path(dir_path)
    md_files = sorted(dir_path.glob("**/*.md"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not md_files:
        show_error(f"目录 {dir_path} 中没有找到 .md 文件")
        sys.exit(1)

    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp())
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    file_links = []
    for md in md_files:
        out = html_render_file(md, output_dir)
        rel = md.relative_to(dir_path)
        mtime = os.path.getmtime(md)
        from datetime import datetime

        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        file_links.append((rel, out.name, date_str))

    items = "\n".join(
        f'<li><a href="{name}">{rel}</a><span class="date">{date}</span></li>'
        for rel, name, date in file_links
    )

    index_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{dir_path.name} — 文件浏览</title>
    {HTML_CSS}
</head>
<body>
    <div class="nav"><strong>{dir_path}</strong> <span style="color:#9ca3af">({len(md_files)} 个文件)</span></div>
    <h1>{dir_path.name}</h1>
    <ul class="file-list">{items}</ul>
</body>
</html>"""

    index_path = output_dir / "index.html"
    index_path.write_text(index_html, encoding="utf-8")
    return index_path


def html_render_directory_from_files(file_list, output_dir):
    """从文件列表生成带索引的 HTML 目录"""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    file_links = []
    for f in file_list:
        md_path = Path(f)
        out = html_render_file(md_path, output_dir)
        mtime = os.path.getmtime(md_path)
        from datetime import datetime

        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        file_links.append((md_path.name, out.name, date_str))

    items = "\n".join(
        f'<li><a href="{name}">{rel}</a><span class="date">{date}</span></li>'
        for rel, name, date in file_links
    )

    index_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>文件浏览</title>
    {HTML_CSS}
</head>
<body>
    <div class="nav"><strong>多文件预览</strong> <span style="color:#9ca3af">({len(file_list)} 个文件)</span></div>
    <h1>文件列表</h1>
    <ul class="file-list">{items}</ul>
</body>
</html>"""

    index_path = output_dir / "index.html"
    index_path.write_text(index_html, encoding="utf-8")
    return index_path


def cmd_to_html(args):
    """to-html 子命令入口"""
    targets = args.targets

    if not targets:
        show_error("用法: md_tools.py to-html <file.md|directory/> [file2.md ...]")
        sys.exit(1)

    target = targets[0]

    if os.path.isdir(target):
        out = html_render_directory(target, args.output_dir)
        show_success(f"索引页: {out}")
        if not args.no_open:
            webbrowser.open(f"file://{out}")
    elif len(targets) > 1:
        tmpdir = Path(tempfile.mkdtemp())
        for f in targets:
            html_render_file(f, tmpdir)
        out = html_render_directory_from_files(targets, tmpdir)
        if not args.no_open:
            webbrowser.open(f"file://{out}")
    else:
        out = html_render_file(target, args.output_dir)
        show_success(f"{out}")
        if not args.no_open:
            webbrowser.open(f"file://{out}")


# ════════════════════════════════════════════════════════════════════
#  frontmatter - 批量生成 MD frontmatter（LLM）
# ════════════════════════════════════════════════════════════════════

FRONTMATTER_IGNORE_DIRS = {".git", "_site", "__pycache__", "node_modules", ".DS_Store"}

FRONTMATTER_SYSTEM_PROMPT = (
    "你是文档分析助手。根据文档内容生成 YAML frontmatter。"
    "只返回 frontmatter 块（含 --- 分隔符），不要其他文字。"
    "description 用中文，≤50字。tags 3-5 个，用中英文均可。"
    '格式：\n---\ndescription: "描述"\ntags: [tag1, tag2, tag3]\n---'
)


def frontmatter_has_frontmatter(text: str) -> bool:
    """检查文本是否已有 YAML frontmatter（以 --- 开头且有第二个 ---）"""
    if not text.startswith("---\n") and not text.startswith("---\r\n"):
        return False
    second = text.find("---", 4)
    return second != -1


def frontmatter_call_llm(content: str) -> str:
    """调用 LLM 生成 frontmatter。"""
    from llm_client import chat

    return chat(
        system=FRONTMATTER_SYSTEM_PROMPT,
        message=f"为以下文档生成 frontmatter：\n\n{content[:2000]}",
    )


def frontmatter_parse_response(text: str) -> str | None:
    """从 API 响应中提取 --- 之间的 frontmatter 块。

    容错处理模型返回多余文字的情况。
    返回完整的 frontmatter 块（含 --- 分隔符和尾部换行），失败返回 None。
    """
    match = re.search(r"---\s*\n(.*?)\n---", text, re.DOTALL)
    if not match:
        return None
    inner = match.group(1).strip()
    if not inner:
        return None
    return f"---\n{inner}\n---\n"


def frontmatter_scan_files(src_dir: str) -> list[tuple[str, str]]:
    """递归扫描目录下的 .md 文件，跳过 FRONTMATTER_IGNORE_DIRS。

    返回 [(relative_path, absolute_path), ...] 按相对路径排序。
    """
    src = Path(src_dir).resolve()
    results = []
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in FRONTMATTER_IGNORE_DIRS]
        for f in files:
            if f.endswith(".md"):
                abs_path = os.path.join(root, f)
                rel_path = os.path.relpath(abs_path, src)
                results.append((rel_path, abs_path))
    results.sort(key=lambda x: x[0])
    return results


def frontmatter_process_file(filepath: str, content: str) -> str | None:
    """对单个文件调用 API 生成 frontmatter。

    返回 frontmatter 字符串，失败返回 None。
    """
    try:
        raw = frontmatter_call_llm(content)
    except Exception as e:
        print(f"  Error: {e}")
        return None

    fm = frontmatter_parse_response(raw)
    if fm is None:
        print(f"  Failed to parse response: {raw[:100]}")
        return None
    return fm


def cmd_frontmatter(args):
    """frontmatter 子命令入口"""
    src_dir = args.src_dir
    if not os.path.isdir(src_dir):
        show_error(f"'{src_dir}' is not a directory")
        sys.exit(1)

    # 收集文件
    if args.file:
        abs_path = os.path.join(os.path.abspath(src_dir), args.file)
        if not os.path.isfile(abs_path):
            show_error(f"file not found: {abs_path}")
            sys.exit(1)
        files = [(args.file, abs_path)]
    else:
        files = frontmatter_scan_files(src_dir)

    if not files:
        print("No .md files found.")
        return

    # 过滤已有 frontmatter 的文件
    to_process = []
    skipped = 0
    for rel, abs_path in files:
        try:
            content = Path(abs_path).read_text(encoding="utf-8")
        except Exception as e:
            print(f"Warning: cannot read {rel}: {e}")
            skipped += 1
            continue
        if frontmatter_has_frontmatter(content):
            skipped += 1
            continue
        to_process.append((rel, abs_path, content))

    total = len(to_process)

    # dry-run 模式
    if args.dry_run:
        print(f"Files to process ({total}):")
        for rel, _, _ in to_process:
            print(f"  {rel}")
        print(f"\nTotal: {total} files to process, {skipped} skipped (already have frontmatter)")
        return

    if total == 0:
        print(f"All files already have frontmatter. Skipped {skipped} files.")
        return

    print(f"Processing {total} files (skipping {skipped} with existing frontmatter)...\n")

    processed = 0
    failed = 0

    for i, (rel, abs_path, content) in enumerate(to_process, 1):
        print(f"[{i}/{total}] Processing: {rel}")

        fm = frontmatter_process_file(abs_path, content)
        if fm is None:
            failed += 1
            print("  FAILED")
            continue

        # 写入文件头部
        try:
            new_content = fm + "\n" + content
            Path(abs_path).write_text(new_content, encoding="utf-8")
            processed += 1
            desc_match = re.search(r'description:\s*"(.+?)"', fm)
            desc = desc_match.group(1) if desc_match else "(parsed)"
            print(f"  OK: {desc}")
        except Exception as e:
            failed += 1
            print(f"  Write error: {e}")

        # Rate limiting: 0.5s 间隔
        if i < total:
            time.sleep(0.5)

    print(f"\nDone! Processed {processed} files, skipped {skipped}, failed {failed}")


# ════════════════════════════════════════════════════════════════════
#  CLI - argparse 主入口
# ════════════════════════════════════════════════════════════════════

def build_parser():
    """构建 argparse 解析器"""
    parser = argparse.ArgumentParser(
        prog="md_tools.py",
        description="Markdown 工具集 - 格式修复/合并/拆分/去引用/转Docx/转HTML/Frontmatter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""子命令示例:
  %(prog)s format file.md                        格式修复（引号/标点/单位）
  %(prog)s merge a.md b.md                       合并多个文件
  %(prog)s merge a.md b.md -o combined.md        合并并指定输出
  %(prog)s split input.md                        按一级标题拆分
  %(prog)s strip md_final/                       检查 blockquote（只统计）
  %(prog)s strip md_final/ --fix                 删除 blockquote
  %(prog)s to-docx input.md                      转换为 docx
  %(prog)s to-docx input.md -t template.docx     用模板转换
  %(prog)s to-html file.md                       渲染为 HTML
  %(prog)s to-html dir/                          目录批量渲染
  %(prog)s frontmatter docs/                     批量生成 frontmatter
  %(prog)s frontmatter docs/ --dry-run           只列出待处理文件
  %(prog)s frontmatter docs/ --file README.md    只处理单个文件
""",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {SCRIPT_VERSION} (by {SCRIPT_AUTHOR}, {SCRIPT_UPDATED})",
    )

    subparsers = parser.add_subparsers(dest="command", help="可用子命令")

    # ── format ──
    p_format = subparsers.add_parser(
        "format",
        help="文本格式自动修复（引号/标点/单位）",
        description="修复 Markdown 文件中的引号、英文标点和中文单位格式",
    )
    p_format.add_argument("files", nargs="*", help="要处理的 Markdown 文件（支持多个）")

    # ── merge ──
    p_merge = subparsers.add_parser(
        "merge",
        help="合并多个 Markdown 文件为一个",
        description="将多个 Markdown 文件按文件名排序后合并为一个文件",
    )
    p_merge.add_argument("files", nargs="*", help="要合并的 Markdown 文件")
    p_merge.add_argument("-o", "--output", help="输出文件名（默认 merged.md）")

    # ── split ──
    p_split = subparsers.add_parser(
        "split",
        help="按一级标题拆分 Markdown 文件",
        description="将一个 Markdown 文件按 # 一级标题拆分为多个文件",
    )
    p_split.add_argument("input", nargs="?", help="要拆分的 Markdown 文件")

    # ── strip ──
    p_strip = subparsers.add_parser(
        "strip",
        help="删除 Markdown 文件中所有 blockquote",
        description="删除 Markdown 文件中所有 blockquote（> 开头的行）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  md_tools.py strip md_final/                            检查模式（只统计）
  md_tools.py strip md_final/ --fix                      删除（覆盖原文件）
  md_tools.py strip md_final/ --fix --output-dir md_out/ 删除（输出到新目录）
""",
    )
    p_strip.add_argument("input", help="MD 文件或目录路径")
    p_strip.add_argument("--fix", action="store_true", help="执行删除")
    p_strip.add_argument("--output-dir", help="输出目录（默认覆盖原文件）")

    # ── to-docx ──
    p_docx = subparsers.add_parser(
        "to-docx",
        help="Markdown 转 Docx（Pandoc 版）",
        description="使用 Pandoc 将 Markdown 文件转换为 Docx 格式",
    )
    p_docx.add_argument("input", nargs="?", help="Markdown 文件路径")
    p_docx.add_argument("-o", "--output", help="输出 docx 文件路径")
    p_docx.add_argument("-t", "--template", help="参考模板 docx（Pandoc --reference-doc）")

    # ── to-html ──
    p_html = subparsers.add_parser(
        "to-html",
        help="Markdown 渲染为 HTML 并在浏览器中打开",
        description="将 Markdown 文件或整个目录渲染为 HTML 页面",
    )
    p_html.add_argument("targets", nargs="*", help="MD 文件或目录（支持多个文件）")
    p_html.add_argument("-o", "--output-dir", help="输出目录（默认用临时目录）")
    p_html.add_argument("--no-open", action="store_true", help="不自动打开浏览器")

    # ── frontmatter ──
    p_fm = subparsers.add_parser(
        "frontmatter",
        help="批量生成 YAML frontmatter（LLM）",
        description="扫描目录下的 Markdown 文件，调用 LLM 生成 description + tags 的 YAML frontmatter",
    )
    p_fm.add_argument("src_dir", help="要扫描的目录路径")
    p_fm.add_argument("--dry-run", action="store_true", help="只列出会处理的文件，不调用 API")
    p_fm.add_argument("--file", dest="file", help="只处理单个文件（相对于 src_dir 的路径）")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # 路由到对应子命令
    commands = {
        "format": cmd_format,
        "merge": cmd_merge,
        "split": cmd_split,
        "strip": cmd_strip,
        "to-docx": cmd_to_docx,
        "to-html": cmd_to_html,
        "frontmatter": cmd_frontmatter,
    }

    handler = commands.get(args.command)
    if handler:
        handler(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
