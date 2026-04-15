#!/usr/bin/env python3
"""4维度 LLM 深度审阅工具 — 逐章检查，输出 Word 批注。

用法:
    python3 review_deep.py input.docx
    python3 review_deep.py input.docx --rules eco-flow-report
    python3 review_deep.py input.docx --rules eco-flow-report --dim completeness
    python3 review_deep.py input.docx -o reviewed.docx
"""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml

# ── 路径设置 ──
SCRIPT_DIR = Path(__file__).resolve().parent
DOCX_TOOLS = SCRIPT_DIR / "docx_tools.py"
PYTHON = "/opt/homebrew/bin/python3"

# llm_client
sys.path.insert(0, str(SCRIPT_DIR.parent.parent / "lib"))
from llm_client import chat

# 规则目录：cc-harness/rules/review-deep/
RULES_DIR = Path.home() / "Dev" / "cc-harness" / "rules" / "review-deep"


def load_rules(rules_name: str) -> dict:
    """加载规则文件。"""
    path = RULES_DIR / f"{rules_name}.yaml"
    if not path.exists():
        # 尝试不带后缀
        path = RULES_DIR / rules_name
        if not path.exists():
            print(f"规则文件不存在: {path}", file=sys.stderr)
            print(f"可用规则: {list_rules()}", file=sys.stderr)
            sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def list_rules() -> list[str]:
    """列出所有可用规则。"""
    if not RULES_DIR.exists():
        return []
    return [f.stem for f in RULES_DIR.glob("*.yaml")]


def extract_chapters(docx_path: str) -> list[dict]:
    """调用 docx_tools.py extract --split-chapters --json 提取章节。"""
    # 先用 split-chapters 模式提取到临时目录
    with tempfile.TemporaryDirectory() as tmpdir:
        cmd = [PYTHON, str(DOCX_TOOLS), "extract", docx_path,
               "--split-chapters", "-o", tmpdir]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"提取失败: {result.stderr}", file=sys.stderr)
            sys.exit(1)

        chapters = []
        for md_file in sorted(Path(tmpdir).glob("*.md")):
            text = md_file.read_text(encoding="utf-8")
            # 从文件名提取章节标题
            name = md_file.stem
            # 格式: 00_前言, 01_1_编制目的与依据
            title = re.sub(r"^\d+_", "", name)
            chapters.append({"title": title, "text": text, "file": name})

        return chapters


def build_prompt(chapter_text: str, dimension: dict, doc_name: str) -> tuple[str, str]:
    """构建单维度检查的 system + user prompt。"""
    rules_text = "\n".join(f"- {r}" for r in dimension["rules"])

    system = f"""你是一位资深水利工程报告审阅专家。
你正在审阅《{doc_name}》，当前检查维度是【{dimension['name']}】：{dimension['description']}。

检查规则：
{rules_text}

输出要求：
1. 逐条检查上述规则，发现问题就报告
2. 每个问题用 JSON 对象表示，格式：
   {{"find": "原文中的问题片段（10-30字，能在原文中精确定位）", "replace": "同find值", "comment": "【{dimension['name']}】问题描述和修改建议"}}
3. 如果建议替换文字，replace 填替换后的内容
4. 输出一个 JSON 数组，没有问题则输出 []
5. 只输出 JSON，不要其他文字"""

    user = f"请审阅以下章节内容：\n\n{chapter_text[:8000]}"

    return system, user


def parse_llm_response(response: str) -> list[dict]:
    """从 LLM 响应中提取 JSON 数组。"""
    # 尝试直接解析
    response = response.strip()

    # 去掉 markdown code fence
    if response.startswith("```"):
        lines = response.split("\n")
        lines = [l for l in lines if not l.startswith("```")]
        response = "\n".join(lines).strip()

    # 找到 JSON 数组
    match = re.search(r"\[[\s\S]*\]", response)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    return []


def review_chapter(chapter: dict, dimensions: dict, doc_name: str,
                   target_dims: list[str] | None = None,
                   model: str = "haiku") -> list[dict]:
    """对单个章节执行多维度检查。"""
    all_issues = []

    for dim_key, dim_config in dimensions.items():
        if target_dims and dim_key not in target_dims:
            continue

        print(f"  检查 [{dim_config['name']}] ...", end=" ", flush=True)
        system, user = build_prompt(chapter["text"], dim_config, doc_name)

        try:
            response = chat(system=system, message=user, model=model)
            issues = parse_llm_response(response)
            print(f"发现 {len(issues)} 个问题")
            all_issues.extend(issues)
        except Exception as e:
            print(f"失败: {e}")

    return all_issues


def apply_reviews(docx_path: str, output_path: str, rules: list[dict]) -> int:
    """调用 docx_tools.py track-changes review 写入批注。"""
    if not rules:
        return 0

    # 写入临时 rules 文件
    rules_file = Path(docx_path).parent / "_review_deep_rules.json"
    with open(rules_file, "w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)

    cmd = [PYTHON, str(DOCX_TOOLS), "track-changes", "review",
           docx_path, "-o", output_path, "-r", str(rules_file),
           "--author", "LLM深度审阅"]
    result = subprocess.run(cmd, capture_output=True, text=True)

    # 清理临时文件
    rules_file.unlink(missing_ok=True)

    if result.returncode != 0:
        print(f"写入批注失败: {result.stderr}", file=sys.stderr)
        return 0

    # 从输出中提取数量
    match = re.search(r"(\d+)\s*处", result.stdout)
    return int(match.group(1)) if match else len(rules)


def main():
    parser = argparse.ArgumentParser(
        description="4维度 LLM 深度审阅工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", nargs="?", help="输入 .docx 文件")
    parser.add_argument("-o", "--output", help="输出文件路径（默认: <原名>_reviewed.docx）")
    parser.add_argument("--rules", default="eco-flow-report",
                        help="规则名称（默认: eco-flow-report）")
    parser.add_argument("--dim", nargs="*",
                        help="指定检查维度（completeness/structure/tone/consistency），默认全部")
    parser.add_argument("--model", default="haiku",
                        help="LLM 模型（haiku/sonnet/opus，默认: haiku）")
    parser.add_argument("--list-rules", action="store_true",
                        help="列出可用规则")
    parser.add_argument("--dry-run", action="store_true",
                        help="只输出问题清单，不写入批注")
    args = parser.parse_args()

    if args.list_rules:
        for r in list_rules():
            rules_data = load_rules(r)
            print(f"  {r}: {rules_data.get('description', '')}")
        return

    if not args.input:
        parser.error("请提供输入文件路径")

    if not os.path.exists(args.input):
        print(f"文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 加载规则
    rules_config = load_rules(args.rules)
    dimensions = rules_config.get("dimensions", {})
    doc_name = rules_config.get("name", Path(args.input).stem)

    print(f"文档: {args.input}")
    print(f"规则: {args.rules} ({rules_config.get('description', '')})")
    print(f"维度: {', '.join(args.dim) if args.dim else '全部'}")
    print(f"模型: {args.model}")
    print()

    # 提取章节
    print("提取章节...")
    chapters = extract_chapters(args.input)
    print(f"共 {len(chapters)} 个章节\n")

    # 逐章检查
    all_rules = []
    for i, chapter in enumerate(chapters):
        print(f"[{i + 1}/{len(chapters)}] {chapter['title']}")
        if len(chapter["text"].strip()) < 50:
            print("  (内容过短，跳过)")
            continue
        issues = review_chapter(chapter, dimensions, doc_name,
                                target_dims=args.dim, model=args.model)
        all_rules.extend(issues)
        print()

    # 汇总
    print(f"{'=' * 50}")
    print(f"共发现 {len(all_rules)} 个问题")

    if not all_rules:
        print("无问题，审阅完成。")
        return

    # 按维度统计
    dim_counts = {}
    for r in all_rules:
        comment = r.get("comment", "")
        for dim_config in dimensions.values():
            if dim_config["name"] in comment:
                dim_counts[dim_config["name"]] = dim_counts.get(dim_config["name"], 0) + 1
                break

    for dim_name, count in dim_counts.items():
        print(f"  {dim_name}: {count} 个")

    if args.dry_run:
        print("\n问题清单:")
        for r in all_rules:
            print(f"  - [{r.get('comment', '')}]")
            if r.get("find") != r.get("replace"):
                print(f"    {r['find']} → {r['replace']}")
        return

    # 写入批注
    output = args.output or str(
        Path(args.input).parent / f"{Path(args.input).stem}_reviewed.docx"
    )
    print(f"\n写入批注到: {output}")
    count = apply_reviews(args.input, output, all_rules)
    print(f"完成: {count} 处批注已写入")


if __name__ == "__main__":
    main()
