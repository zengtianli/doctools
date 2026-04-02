#!/usr/bin/env python3
"""
Bullet Point 转段落/表格工具

将 Markdown 文件中的 bullet point 列表转换为符合公文规范的段落或表格形式。
调用智谱 API 进行智能转换，排除表格内和代码块内的 `-` 符号。

用法：
    python3 bullet_to_paragraph.py <输入路径> [--dry-run] [--output-dir <目录>]

参数：
    输入路径        .md 文件或包含 .md 文件的目录
    --dry-run       仅检测，不修改文件
    --output-dir    输出到指定目录（默认直接覆盖原文件）

示例：
    # 单文件检测
    python3 bullet_to_paragraph.py file.md --dry-run

    # 单文件转换
    python3 bullet_to_paragraph.py file.md

    # 目录批量转换
    python3 bullet_to_paragraph.py ./md/

    # 输出到新目录
    python3 bullet_to_paragraph.py ./md/ --output-dir ./md_clean
"""

import argparse
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))
from display import show_error, show_info, show_success, show_warning
from file_ops import show_help_footer, show_help_header, show_version_info

SCRIPT_VERSION = "2.0.0"
SCRIPT_AUTHOR = "tianli"
SCRIPT_UPDATED = "2026-03-14"


def show_version():
    """显示版本信息"""
    show_version_info(SCRIPT_VERSION, SCRIPT_AUTHOR, SCRIPT_UPDATED)


def show_help():
    """显示帮助信息"""
    show_help_header(sys.argv[0], "将 Markdown 中的 bullet point 转换为公文规范格式")
    print("    <输入路径>             .md 文件或包含 .md 文件的目录")
    print("    --dry-run              仅检测，不修改文件")
    print("    --output-dir <目录>    输出到指定目录（默认覆盖原文件）")
    show_help_footer()


def get_client():
    """保留接口兼容，返回 None（已改用 claude CLI）"""
    return None


NUMBERED_LIST_RE = re.compile(r"^\d+\.\s+")


def _is_bullet_line(line: str) -> bool:
    """判断是否为 bullet point 行（`- ` 开头）"""
    return line.startswith("- ")


def _is_numbered_line(line: str) -> bool:
    """判断是否为有序列表行（`数字. ` 开头）"""
    return bool(NUMBERED_LIST_RE.match(line))


def _is_list_line(line: str) -> bool:
    """判断是否为列表行（bullet point 或有序列表）"""
    return _is_bullet_line(line) or _is_numbered_line(line)


def extract_bullet_blocks(content: str) -> list:
    """提取文档中的列表块（bullet point 和有序列表），排除表格内和代码块内的"""
    lines = content.split("\n")
    blocks = []
    current_block = []
    block_start_line = -1
    in_block = False
    block_type = None  # 'bullet' 或 'numbered'
    in_table = False
    in_code = False

    for i, line in enumerate(lines):
        # 代码块边界
        if line.strip().startswith("```"):
            in_code = not in_code

        # 表格边界
        if not in_code:
            if line.strip().startswith("|") and "|" in line[1:]:
                in_table = True
            elif in_table and not line.strip().startswith("|"):
                in_table = False

        # 只处理非表格、非代码块内的列表行
        is_bullet = _is_bullet_line(line) and not in_table and not in_code
        is_numbered = _is_numbered_line(line) and not in_table and not in_code

        if is_bullet or is_numbered:
            current_type = "bullet" if is_bullet else "numbered"
            if not in_block:
                in_block = True
                block_start_line = i
                block_type = current_type
                current_block.append(line)
            elif current_type == block_type:
                # 同类型列表，继续当前块
                current_block.append(line)
            else:
                # 类型切换，先保存当前块，再开始新块
                blocks.append(
                    {
                        "start_line": block_start_line,
                        "end_line": i - 1,
                        "lines": list(current_block),
                        "block_type": block_type,
                        "context_before": "\n".join(lines[max(0, block_start_line - 3) : block_start_line]),
                        "context_after": "\n".join(lines[i : min(len(lines), i + 3)]),
                    }
                )
                current_block = [line]
                block_start_line = i
                block_type = current_type
        else:
            if in_block and current_block:
                blocks.append(
                    {
                        "start_line": block_start_line,
                        "end_line": i - 1,
                        "lines": list(current_block),
                        "block_type": block_type,
                        "context_before": "\n".join(lines[max(0, block_start_line - 3) : block_start_line]),
                        "context_after": "\n".join(lines[i : min(len(lines), i + 3)]),
                    }
                )
                current_block = []
                in_block = False
                block_type = None

    if in_block and current_block:
        blocks.append(
            {
                "start_line": block_start_line,
                "end_line": len(lines) - 1,
                "lines": list(current_block),
                "block_type": block_type,
                "context_before": "\n".join(lines[max(0, block_start_line - 3) : block_start_line]),
                "context_after": "",
            }
        )

    return blocks


def clean_api_output(text: str) -> str:
    """清理 API 输出，去除 <thinking> 标签和多余解释"""
    # 去除 <thinking>...</thinking>
    text = re.sub(r"<thinking>.*?</thinking>", "", text, flags=re.DOTALL)
    # 去除可能的开头解释
    text = text.strip()
    # 如果输出以 "以下是" 或 "转换后" 等解释开头，跳到实际内容
    for prefix in ["以下是", "转换后的内容", "转换结果"]:
        if text.startswith(prefix):
            lines = text.split("\n", 1)
            if len(lines) > 1:
                text = lines[1].strip()
    return text


def convert_bullet_block(client, block: dict) -> str:
    """调用 Claude 转换单个 bullet block"""
    from llm_client import chat as llm_chat

    bullet_content = "\n".join(block["lines"])
    block_type = block.get("block_type", "bullet")
    type_desc = "有序列表（数字编号开头）" if block_type == "numbered" else "bullet point（`- ` 开头的列表）"
    prompt = f"""将以下{type_desc}转换为符合公文规范的格式。

规则：
1. 不允许使用 bullet point（`- ` 开头的列表）或有序列表（`1. ` `2. ` 等数字编号开头的列表）
2. 并列关系的要点（3项以上）→ 表格
3. 流程步骤 → 段落，用"第一""第二"连接
4. 2-3 项说明 → 段落，用"一是...二是...三是..."
5. 保持原有信息完整
6. 禁用词：确保、我们、我司

上下文：
{block["context_before"]}

待转换：
{bullet_content}

下文：
{block["context_after"]}

仅输出转换后的内容，不要解释、不要思考过程。"""

    try:
        raw = llm_chat(
            system="你是公文格式转换器。只输出转换后的内容，不要解释。",
            message=prompt,
        )
        return clean_api_output(raw)
    except Exception as e:
        show_error(f"API 调用失败: {e}")
        return bullet_content


def process_file(client, file_path: Path, output_path: Path, dry_run: bool) -> dict:
    """处理单个文件，返回统计信息"""
    with open(file_path, encoding="utf-8") as f:
        content = f.read()

    blocks = extract_bullet_blocks(content)
    bullet_count = sum(len(b["lines"]) for b in blocks)

    stats = {"file": file_path.name, "blocks": len(blocks), "bullets": bullet_count, "converted": 0, "remaining": 0}

    if not blocks:
        show_info(f"{file_path.name}: 无 bullet point")
        return stats

    if dry_run:
        show_info(f"{file_path.name}: {len(blocks)} 块, {bullet_count} 个 bullet point")
        for b in blocks:
            print(f"    L{b['start_line'] + 1}-{b['end_line'] + 1}: {b['lines'][0][:60]}...")
        return stats

    show_info(f"{file_path.name}: 转换 {len(blocks)} 块...")
    lines = content.split("\n")

    for block in reversed(blocks):
        converted = convert_bullet_block(client, block)
        # 验证：转换结果不应包含 bullet point 或有序列表
        new_bullets = [ln for ln in converted.split("\n") if _is_list_line(ln)]
        if new_bullets:
            show_warning(f"L{block['start_line'] + 1}: 转换后仍有 {len(new_bullets)} 个列表项，重试...")
            converted = convert_bullet_block(client, block)
            new_bullets = [ln for ln in converted.split("\n") if _is_list_line(ln)]
            if new_bullets:
                show_error(f"L{block['start_line'] + 1}: 重试仍有列表项，保留原文")
                stats["remaining"] += len(block["lines"])
                continue

        lines[block["start_line"] : block["end_line"] + 1] = converted.split("\n")
        stats["converted"] += len(block["lines"])

    result = "\n".join(lines)

    # 最终验证
    final_bullets = len([ln for ln in result.split("\n") if _is_list_line(ln) and not _in_table_or_code(result, ln)])
    stats["remaining"] = final_bullets

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result)

    if final_bullets == 0:
        show_success(f"完成 → {output_path.name}")
    else:
        show_warning(f"剩余 {final_bullets} 个 bullet point → {output_path.name}")
    return stats


def _in_table_or_code(content: str, target_line: str) -> bool:
    """简单判断某行是否在表格或代码块内"""
    in_code = False
    in_table = False
    for line in content.split("\n"):
        if line.strip().startswith("```"):
            in_code = not in_code
        if not in_code and line.strip().startswith("|") and "|" in line[1:]:
            in_table = True
        elif in_table and not line.strip().startswith("|"):
            in_table = False
        if line == target_line and (in_code or in_table):
            return True
    return False


def main():
    parser = argparse.ArgumentParser(description="将 Markdown 中的 bullet point 转换为公文规范格式（表格或段落）")
    parser.add_argument("input_path", type=Path, help=".md 文件或包含 .md 文件的目录")
    parser.add_argument("--dry-run", action="store_true", help="仅检测，不修改文件")
    parser.add_argument("--output-dir", type=Path, default=None, help="输出目录（默认覆盖原文件）")
    parser.add_argument("--version", action="store_true", help="显示版本信息")
    args = parser.parse_args()

    if args.version:
        show_version()
        sys.exit(0)

    input_path = args.input_path

    # 收集待处理文件
    if input_path.is_file():
        if input_path.suffix != ".md":
            show_error(f"不是 .md 文件: {input_path}")
            sys.exit(1)
        md_files = [input_path]
    elif input_path.is_dir():
        md_files = sorted(input_path.glob("*.md"))
        if not md_files:
            show_error(f"目录中无 .md 文件: {input_path}")
            sys.exit(1)
    else:
        show_error(f"路径不存在: {input_path}")
        sys.exit(1)

    # 初始化 API 客户端
    client = get_client()

    mode = "检测模式" if args.dry_run else "转换模式"
    show_info(f"{mode} | 共 {len(md_files)} 个文件")
    print()

    all_stats = []
    for md_file in md_files:
        output = (args.output_dir / md_file.name) if args.output_dir else md_file
        stats = process_file(client, md_file, output, args.dry_run)
        all_stats.append(stats)

    # 汇总
    total_blocks = sum(s["blocks"] for s in all_stats)
    total_bullets = sum(s["bullets"] for s in all_stats)
    total_converted = sum(s["converted"] for s in all_stats)
    total_remaining = sum(s["remaining"] for s in all_stats)

    print(f"\n{'=' * 40}")
    show_info(f"总计: {total_bullets} 个 bullet point, {total_blocks} 个块")
    if not args.dry_run:
        show_info(f"已转换: {total_converted}, 剩余: {total_remaining}")
        if total_remaining > 0:
            show_warning("仍有未转换的 bullet point，建议再跑一次")
    print(f"{'=' * 40}")


if __name__ == "__main__":
    main()
