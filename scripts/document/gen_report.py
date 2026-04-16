#!/usr/bin/env python3
"""基于参考报告逐章生成新县市报告。

用法:
    python3 gen_report.py \\
        --ref-dir ref_chapters/ \\
        --data-dir taizhou-天台/data/ \\
        --config eco-flow-report.yaml \\
        --output generated/ \\
        --county 天台县 --city 台州市

    python3 gen_report.py ... --chapters 01,02    # 只生成指定章节
    python3 gen_report.py ... --dry-run            # 只打印 prompt
    python3 gen_report.py ... --force              # 覆盖已有章节
"""

import argparse
import glob
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd
import yaml

# ── 路径设置 ──
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR.parent.parent / "lib"))
from llm_client import chat

# 规则目录
RULES_DIR = Path.home() / "Dev" / "cc-harness" / "rules" / "gen-report"


# ═══════════════════════════════════════════════════════════════
#  配置加载
# ═══════════════════════════════════════════════════════════════

def load_config(config_name: str, overrides: dict) -> dict:
    """加载章节策略 YAML，合并命令行覆盖变量。"""
    path = RULES_DIR / config_name
    if not path.exists():
        path = RULES_DIR / f"{config_name}.yaml"
    if not path.exists():
        print(f"配置不存在: {path}", file=sys.stderr)
        sys.exit(1)

    with open(path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    # 合并变量
    config["variables"] = {
        "county": overrides.get("county", ""),
        "county_short": overrides.get("county", "").replace("县", "").replace("区", "").replace("市", ""),
        "city": overrides.get("city", ""),
        "city_short": overrides.get("city", "").replace("市", ""),
        **overrides,
    }
    return config


def list_configs() -> list[str]:
    """列出可用配置。"""
    if not RULES_DIR.exists():
        return []
    return [f.stem for f in RULES_DIR.glob("*.yaml")]


# ═══════════════════════════════════════════════════════════════
#  数据读取
# ═══════════════════════════════════════════════════════════════

def resolve_data_files(patterns: list[str], data_dir: str) -> list[str]:
    """展开 data_files 中的 glob 模式，返回实际文件路径。"""
    files = []
    for pattern in patterns:
        expanded = pattern.replace("{data_dir}", data_dir)
        matches = sorted(glob.glob(expanded))
        files.extend(matches)
    return files


def read_data_context(file_paths: list[str], max_xlsx_rows: int = 50) -> str:
    """读取数据文件，拼接为 LLM 可用的上下文。"""
    parts = []
    for fp in file_paths:
        p = Path(fp)
        if not p.exists():
            continue

        if p.suffix in (".md", ".txt"):
            content = p.read_text(encoding="utf-8")
            parts.append(f"### 数据文件: {p.name}\n\n{content}")

        elif p.suffix in (".xlsx", ".xls"):
            try:
                df = pd.read_excel(fp, engine="openpyxl" if p.suffix == ".xlsx" else None)
                if len(df) > max_xlsx_rows:
                    df = df.head(max_xlsx_rows)
                    truncated = f"（仅显示前 {max_xlsx_rows} 行）"
                else:
                    truncated = ""
                md_table = df.to_markdown(index=False)
                parts.append(f"### 数据文件: {p.name} {truncated}\n\n{md_table}")
            except Exception as e:
                parts.append(f"### 数据文件: {p.name}\n\n[读取失败: {e}]")

    return "\n\n".join(parts)


# ═══════════════════════════════════════════════════════════════
#  章节拆分
# ═══════════════════════════════════════════════════════════════

def split_by_heading(content: str, heading_prefix: str = "## ") -> list[dict]:
    """按指定标题级别拆分内容。"""
    sections = []
    current_title = ""
    current_lines = []

    for line in content.split("\n"):
        if line.startswith(heading_prefix):
            if current_lines:
                sections.append({
                    "title": current_title,
                    "text": "\n".join(current_lines),
                })
            current_title = line.strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_lines:
        sections.append({
            "title": current_title,
            "text": "\n".join(current_lines),
        })

    return sections


# ═══════════════════════════════════════════════════════════════
#  Prompt 构建
# ═══════════════════════════════════════════════════════════════

SYSTEM_REPLACE = """你是水利报告编辑。将参考报告中的地名替换为目标县市，保持原文结构和措辞不变。

替换规则：
{replacements}

要求：
1. 仅替换地名，不改变内容结构和专业表述
2. 替换后检查上下文通顺，如有因地名变化导致的不通顺，做最小修改
3. 数字和数据如与地名强相关（如"全县XX座水库"），用 [待填] 标记
4. 输出纯 Markdown，不加任何说明"""

SYSTEM_REWRITE = """你是资深水利工程师，正在编写《{county}小型水库工程生态流量核定与保障实施方案》。

参考以下已完成报告的对应章节（来自其他县市），学习其结构、论证逻辑和行文风格，结合提供的{county}数据，编写{county}的对应章节。

写作要求：
1. 严格保持参考报告的章节结构（标题层级、段落组织）
2. 所有数据必须来自提供的数据文件，不要编造
3. 如数据不足，用 [待填：XXX] 标记需要补充的内容
4. 保持乙方立场，用"建议""按要求确定"等措辞
5. 表格数据直接使用提供的数据，保持 Markdown 表格格式
6. 输出纯 Markdown，不加任何说明"""

SYSTEM_REGENERATE = """你是资深水利工程师，正在编写《{county}小型水库工程生态流量核定与保障实施方案》的结论章。

基于以下已完成章节的内容，撰写结论。参考报告的结论章节作为格式参考。

要求：
1. 汇总核定成果（水库数量、生态流量目标值）
2. 概括保障措施（泄放设施、监测设施、调度方案）
3. 提出后续工作建议
4. 保持乙方立场，措辞严谨
5. 输出纯 Markdown"""


def build_replacements_text(config: dict) -> str:
    """构建替换规则文本。"""
    variables = config["variables"]
    replacements = config.get("replacements", {})
    lines = []
    for old, new_template in replacements.items():
        new_val = new_template
        for k, v in variables.items():
            new_val = new_val.replace("{" + k + "}", v)
        if old != new_val:
            lines.append(f"- 「{old}」→「{new_val}」")
    return "\n".join(lines) if lines else "（无需替换）"


def build_prompt(chapter_config: dict, ref_content: str,
                 data_context: str, config: dict,
                 generated_chapters: dict) -> tuple[str, str]:
    """根据策略构建 system + user prompt。"""
    strategy = chapter_config["strategy"]
    variables = config["variables"]
    county = variables["county"]

    if strategy == "replace":
        system = SYSTEM_REPLACE.format(
            replacements=build_replacements_text(config)
        )
        user = f"请替换以下章节中的地名：\n\n{ref_content}"

    elif strategy == "rewrite":
        system = SYSTEM_REWRITE.format(county=county)
        user_parts = [f"## 参考报告章节（其他县市）\n\n{ref_content}"]
        if data_context:
            user_parts.append(f"## {county}数据\n\n{data_context}")
        user = "\n\n".join(user_parts)

    elif strategy == "regenerate":
        system = SYSTEM_REGENERATE.format(county=county)
        user_parts = [f"## 参考报告结论（格式参考）\n\n{ref_content}"]
        # 收集已生成章节
        depends = chapter_config.get("depends_on", [])
        for dep_id in depends:
            if dep_id in generated_chapters:
                user_parts.append(
                    f"## 已完成章节\n\n{generated_chapters[dep_id][:3000]}"
                )
        user = "\n\n".join(user_parts)

    else:
        return "", ""

    return system, user


# ═══════════════════════════════════════════════════════════════
#  章节生成
# ═══════════════════════════════════════════════════════════════

def find_ref_file(ref_dir: str, chapter_id: str) -> str | None:
    """根据 chapter_id 前缀匹配参考章节文件。"""
    pattern = os.path.join(ref_dir, f"{chapter_id}_*.md")
    matches = sorted(glob.glob(pattern))
    return matches[0] if matches else None


def generate_chapter(chapter_config: dict, ref_dir: str, data_dir: str,
                     output_dir: str, config: dict,
                     generated_chapters: dict,
                     dry_run: bool = False) -> str | None:
    """生成单个章节。"""
    ch_id = chapter_config["id"]
    strategy = chapter_config["strategy"]
    title = chapter_config["title"]
    model = chapter_config.get("model", "haiku")

    if strategy == "skip":
        print(f"  [跳过] {title}")
        return None

    # 读取参考章节
    ref_file = find_ref_file(ref_dir, ch_id)
    if not ref_file:
        print(f"  [警告] 未找到参考章节文件: {ch_id}_*.md")
        return None
    ref_content = Path(ref_file).read_text(encoding="utf-8")

    # 读取数据文件
    data_context = ""
    data_files_patterns = chapter_config.get("data_files", [])
    if data_files_patterns:
        files = resolve_data_files(data_files_patterns, data_dir)
        if files:
            data_context = read_data_context(files)
        else:
            print(f"  [警告] 未找到数据文件: {data_files_patterns}")

    # 判断是否需要分段生成
    split_by = chapter_config.get("split_by")
    if split_by and strategy == "rewrite" and len(ref_content) > 2000:
        return generate_chapter_segmented(
            chapter_config, ref_content, data_context,
            config, generated_chapters, model, dry_run
        )

    # 构建 prompt
    system, user = build_prompt(
        chapter_config, ref_content, data_context,
        config, generated_chapters
    )

    if dry_run:
        print(f"\n{'='*60}")
        print(f"[DRY RUN] 章节: {title} | 策略: {strategy} | 模型: {model}")
        print(f"System prompt ({len(system)} 字):")
        print(system[:500] + "..." if len(system) > 500 else system)
        print(f"\nUser message ({len(user)} 字):")
        print(user[:500] + "..." if len(user) > 500 else user)
        return None

    # 调用 LLM
    print(f"  调用 LLM ({model})...", end=" ", flush=True)
    try:
        result = chat(system=system, message=user, model=model)
        print(f"完成 ({len(result)} 字)")
        return result
    except Exception as e:
        print(f"失败: {e}")
        return None


def generate_chapter_segmented(chapter_config: dict, ref_content: str,
                                data_context: str, config: dict,
                                generated_chapters: dict,
                                model: str, dry_run: bool) -> str | None:
    """分段生成大章节（按二级标题拆分）。"""
    split_by = chapter_config.get("split_by", "## ")
    sections = split_by_heading(ref_content, split_by)
    title = chapter_config["title"]

    if not sections:
        sections = [{"title": title, "text": ref_content}]

    print(f"  分 {len(sections)} 段生成")
    results = []

    for i, section in enumerate(sections):
        sec_title = section["title"] or f"段落{i}"
        print(f"    [{i+1}/{len(sections)}] {sec_title[:40]}", end=" ", flush=True)

        # 为每段构建 prompt（复用 rewrite 策略）
        sec_config = {**chapter_config, "strategy": "rewrite"}
        system, user = build_prompt(
            sec_config, section["text"], data_context,
            config, generated_chapters
        )

        if dry_run:
            print(f"[DRY RUN] ({len(section['text'])} 字)")
            continue

        try:
            result = chat(system=system, message=user, model=model)
            print(f"完成 ({len(result)} 字)")
            results.append(result)
        except Exception as e:
            print(f"失败: {e}")
            results.append(f"[生成失败: {sec_title}]\n\n")

    return "\n\n".join(results) if results else None


# ═══════════════════════════════════════════════════════════════
#  主流程
# ═══════════════════════════════════════════════════════════════

def generate_all(config: dict, ref_dir: str, data_dir: str,
                 output_dir: str, chapter_filter: list | None = None,
                 dry_run: bool = False, force: bool = False):
    """按顺序逐章生成。"""
    os.makedirs(output_dir, exist_ok=True)
    chapters = config.get("chapters", [])
    generated_chapters = {}  # id -> content，供 regenerate 策略使用

    for ch in chapters:
        ch_id = ch["id"]
        title = ch["title"]

        # 过滤
        if chapter_filter and ch_id not in chapter_filter:
            continue

        # 检查依赖
        depends = ch.get("depends_on", [])
        for dep in depends:
            if dep not in generated_chapters:
                dep_file = os.path.join(output_dir, f"{dep}_*.md")
                dep_matches = sorted(glob.glob(dep_file))
                if dep_matches:
                    generated_chapters[dep] = Path(dep_matches[0]).read_text(encoding="utf-8")

        # 检查已有文件
        out_pattern = os.path.join(output_dir, f"{ch_id}_*.md")
        existing = sorted(glob.glob(out_pattern))
        if existing and not force and not dry_run:
            print(f"[{ch_id}] {title} — 已存在，跳过（用 --force 覆盖）")
            generated_chapters[ch_id] = Path(existing[0]).read_text(encoding="utf-8")
            continue

        print(f"[{ch_id}] {title} ({ch.get('strategy', '?')})")
        result = generate_chapter(
            ch, ref_dir, data_dir, output_dir, config,
            generated_chapters, dry_run
        )

        if result and not dry_run:
            # 保存到文件
            safe_title = re.sub(r"[^\w\u4e00-\u9fff]+", "_", title)[:30]
            out_file = os.path.join(output_dir, f"{ch_id}_{safe_title}.md")
            Path(out_file).write_text(result, encoding="utf-8")
            print(f"  → {out_file}")
            generated_chapters[ch_id] = result

    if not dry_run:
        print(f"\n{'='*50}")
        print(f"生成完成，输出目录: {output_dir}")
        out_files = sorted(Path(output_dir).glob("*.md"))
        total_chars = sum(f.stat().st_size for f in out_files)
        print(f"共 {len(out_files)} 个文件，{total_chars} 字符")


def main():
    parser = argparse.ArgumentParser(
        description="基于参考报告逐章生成新县市报告",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--ref-dir", help="参考报告的拆分 MD 目录")
    parser.add_argument("--data-dir", help="目标县市的 data 目录")
    parser.add_argument("--config", default="eco-flow-report",
                        help="策略配置名称（默认: eco-flow-report）")
    parser.add_argument("--output", help="输出目录")
    parser.add_argument("--county", help="目标县市名（如 天台县）")
    parser.add_argument("--city", help="目标城市名（如 台州市）")
    parser.add_argument("--chapters", help="只生成指定章节（逗号分隔，如 01,02）")
    parser.add_argument("--model", help="覆盖所有章节的模型")
    parser.add_argument("--dry-run", action="store_true", help="只打印 prompt，不调用 LLM")
    parser.add_argument("--force", action="store_true", help="覆盖已有章节文件")
    parser.add_argument("--list-configs", action="store_true", help="列出可用配置")
    args = parser.parse_args()

    if args.list_configs:
        for c in list_configs():
            print(f"  {c}")
        return

    # 检查必填参数
    missing = [n for n in ("ref_dir", "data_dir", "output", "county", "city")
               if not getattr(args, n)]
    if missing:
        parser.error(f"缺少必填参数: {', '.join('--' + m.replace('_', '-') for m in missing)}")

    # 加载配置
    config = load_config(args.config, {
        "county": args.county,
        "city": args.city,
    })

    # 覆盖模型
    if args.model:
        for ch in config.get("chapters", []):
            ch["model"] = args.model

    # 解析章节过滤
    chapter_filter = None
    if args.chapters:
        chapter_filter = [c.strip() for c in args.chapters.split(",")]

    print(f"配置: {args.config}")
    print(f"目标: {args.county}（{args.city}）")
    print(f"参考: {args.ref_dir}")
    print(f"数据: {args.data_dir}")
    print(f"输出: {args.output}")
    print()

    generate_all(
        config, args.ref_dir, args.data_dir, args.output,
        chapter_filter=chapter_filter,
        dry_run=args.dry_run,
        force=args.force,
    )


if __name__ == "__main__":
    main()
