#!/usr/bin/env python3
"""
报告/标书质量检查 + 自动修复脚本

检查项（本地，不调 API）：
  1. 禁用词检查（确保、我们、我司）
  2. Bullet point 检查（^- 开头，排除表格和代码块）
  3. 数据来源标注检查（含数字的句子是否标注来源）
  4. 标书评分对齐检查（--bid 模式）
  5. 重复行检测（相邻或间隔1行的完全相同非空行）
  6. 有序列表检测（数字. 开头的行）

自动修复（--fix 模式）：
  1. 禁用词替换
  2. Bullet point → 段落（调用 bullet_to_paragraph 模块）
  3. 重复行删除（保留第一次出现）
  4. 有序列表 → 段落（调用 bullet_to_paragraph 模块）
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))
from display import show_error, show_info, show_success, show_warning
from file_ops import show_version_info

# 尝试导入 bullet_to_paragraph 模块
try:
    from bullet_to_paragraph import convert_bullet_block, extract_bullet_blocks

    BULLET_MODULE_AVAILABLE = True
except ImportError:
    BULLET_MODULE_AVAILABLE = False

SCRIPT_VERSION = "1.1.0"
SCRIPT_AUTHOR = "tianli"
SCRIPT_UPDATED = "2026-03-14"

# ── 禁用词配置 ──────────────────────────────────────────────

FORBIDDEN_WORDS = {
    "确保": "保障",
    "我们": "本项目团队",
    "我司": "本单位",
}

# ── 数据模式：匹配含数字+单位的句子片段 ────────────────────

DATA_PATTERN = re.compile(
    r"\d+[\d.,]*\s*(?:万|亿|千|百)?"
    r"(?:m³|m²|km²|km|hm²|万m³|亿m³|万元|亿元|元|%|‰|℃|mm|cm|MW|kW|GW|万kW"
    r"|吨|万吨|亿吨|公顷|亩|万亩|人|万人|户|万户|座|个|处|条|段|台|套|站"
    r"|立方米每秒|m³/s|L/s)"
)

SOURCE_PATTERN = re.compile(r"\[(?:数据)?来源[：:]")

# 更宽泛的来源识别：根据/依据/按照/《》引用 等
SOURCE_INDICATORS = re.compile(
    r"根据|依据|按照|参照|来源于|数据来源|来自|出自|引自|摘自|"
    r"《[^》]+》"
)

# 已有 SOURCE 注释
EXISTING_SOURCE_COMMENT = re.compile(r"<!--\s*SOURCE:")

# 招标文件相关章节关键词
BID_SECTION_KEYWORDS = [
    "验收",
    "工作要求",
    "工作内容",
    "付款",
    "提交标准",
    "预期成果",
    "成果提交",
    "成果验收",
    "质量标准",
]

# 要求/标准句式（表示数据来自招标要求）
REQUIREMENT_PHRASES = [
    "不少于",
    "达到",
    "控制在",
    "不低于",
    "不超过",
    "不小于",
    "≥",
    "≤",
    "覆盖率",
    "通过率",
    "合格率",
]


# ── 工具函数 ────────────────────────────────────────────────


def is_in_code_block(lines: list[str], line_idx: int) -> bool:
    """判断某行是否在代码块（```）内"""
    in_code = False
    for i in range(line_idx):
        if lines[i].strip().startswith("```"):
            in_code = not in_code
    return in_code


def is_in_table(line: str) -> bool:
    """判断某行是否是表格行"""
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|")


# ── 检查函数 ────────────────────────────────────────────────


def check_forbidden_words(lines: list[str]) -> list[dict]:
    """检查禁用词，返回问题列表"""
    issues = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        for word, replacement in FORBIDDEN_WORDS.items():
            col = line.find(word)
            while col != -1:
                # 提取上下文（前后各 15 字符）
                start = max(0, col - 15)
                end = min(len(line), col + len(word) + 15)
                context = line[start:end].strip()
                issues.append(
                    {
                        "type": "禁用词",
                        "line": i + 1,
                        "word": word,
                        "replacement": replacement,
                        "context": context,
                        "fixable": True,
                    }
                )
                col = line.find(word, col + len(word))
    return issues


def check_bullet_points(lines: list[str]) -> list[dict]:
    """检查 bullet point，返回问题列表"""
    issues = []
    in_block = False
    block_start = -1
    block_count = 0

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i) or is_in_table(line):
            if in_block:
                issues.append(
                    {
                        "type": "bullet",
                        "line_start": block_start + 1,
                        "line_end": block_start + block_count,
                        "count": block_count,
                        "fixable": BULLET_MODULE_AVAILABLE,
                    }
                )
                in_block = False
                block_count = 0
            continue

        if re.match(r"^- ", line):
            if not in_block:
                in_block = True
                block_start = i
                block_count = 0
            block_count += 1
        else:
            if in_block:
                issues.append(
                    {
                        "type": "bullet",
                        "line_start": block_start + 1,
                        "line_end": block_start + block_count,
                        "count": block_count,
                        "fixable": BULLET_MODULE_AVAILABLE,
                    }
                )
                in_block = False
                block_count = 0

    # 文件末尾仍在 block 中
    if in_block:
        issues.append(
            {
                "type": "bullet",
                "line_start": block_start + 1,
                "line_end": block_start + block_count,
                "count": block_count,
                "fixable": BULLET_MODULE_AVAILABLE,
            }
        )

    return issues


def _find_nearest_heading(lines: list[str], line_idx: int) -> str:
    """向上找最近的标题"""
    for i in range(line_idx, -1, -1):
        stripped = lines[i].strip()
        if stripped.startswith("#"):
            return stripped
    return ""


def _has_nearby_table(lines: list[str], line_idx: int, radius: int = 5) -> bool:
    """检查附近是否有表格"""
    return any(is_in_table(lines[i]) for i in range(max(0, line_idx - radius), min(len(lines), line_idx + radius)))


def _classify_data_source(lines: list[str], line_idx: int) -> str:
    """判断数据的可能来源类别"""
    line = lines[line_idx]

    # 1. 要求/标准句式 → 招标文件
    if any(p in line for p in REQUIREMENT_PHRASES):
        return "招标文件"

    # 2. 所在章节为招标要求相关 → 招标文件
    heading = _find_nearest_heading(lines, line_idx)
    if any(kw in heading for kw in BID_SECTION_KEYWORDS):
        return "招标文件"

    # 3. 附近有表格 → 表格数据（表格本身就是来源证据）
    if _has_nearby_table(lines, line_idx):
        return "表格数据"

    return "PLACEHOLDER"


def check_data_sources(lines: list[str]) -> list[dict]:
    """检查含数字的句子是否有数据来源标注，并分类来源"""
    issues = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i) or is_in_table(line):
            continue
        # 跳过标题行
        if line.strip().startswith("#"):
            continue

        matches = DATA_PATTERN.findall(line)
        if not matches:
            continue

        # 已有来源标注（多种形式）
        if SOURCE_PATTERN.search(line) or SOURCE_INDICATORS.search(line):
            continue
        if EXISTING_SOURCE_COMMENT.search(line):
            continue
        # 检查下一行
        if i + 1 < len(lines):
            next_line = lines[i + 1]
            if SOURCE_PATTERN.search(next_line) or SOURCE_INDICATORS.search(next_line):
                continue

        # 分类来源
        source_type = _classify_data_source(lines, i)

        # 表格附近的数据：表格本身就是来源证据，不报警
        if source_type == "表格数据":
            continue

        issues.append(
            {
                "type": "数据来源",
                "line": i + 1,
                "data": matches[0],
                "source_type": source_type,
                "fixable": True,
            }
        )

    return issues


# ASCII 图形字符集（用于排除 ASCII 图形行）
ASCII_ART_CHARS = set("│┌└┃─┐┘▼├┤┬┴╪╭╮╯╰═╬╠╣╦╩║")

# 有序列表正则
NUMBERED_LIST_RE = re.compile(r"^\d+\.\s+")


def _is_ascii_art_line(line: str) -> bool:
    """判断是否为 ASCII 图形行"""
    return any(ch in line for ch in ASCII_ART_CHARS)


def _is_table_separator(line: str) -> bool:
    """判断是否为表格分隔行（如 |------|------|）"""
    stripped = line.strip()
    if not stripped.startswith("|"):
        return False
    # 去掉 | 后只剩 -、:、空格
    inner = stripped.replace("|", "").replace("-", "").replace(":", "").strip()
    return inner == ""


def _is_table_header(line: str) -> bool:
    """判断是否为表格表头行（| xxx | xxx |）"""
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|")


def check_duplicate_lines(lines: list[str]) -> list[dict]:
    """检测连续出现的重复行（相邻或间隔1行的完全相同非空行）

    排除：代码块内的行、表格分隔行、表格表头行、ASCII 图形行、空行
    """
    issues = []
    in_code = False

    for i, line in enumerate(lines):
        stripped = line.strip()

        # 跟踪代码块状态
        if stripped.startswith("```"):
            in_code = not in_code
            continue

        # 跳过代码块内
        if in_code:
            continue

        # 跳过空行
        if not stripped:
            continue

        # 跳过表格分隔行、表格表头行、ASCII 图形行
        if _is_table_separator(line) or _is_ascii_art_line(line):
            continue
        if _is_table_header(line):
            continue

        # 检查与前一行（相邻）是否重复
        if i >= 1:
            prev = lines[i - 1].strip()
            if prev == stripped and prev:
                issues.append(
                    {
                        "type": "重复行",
                        "line": i + 1,
                        "content": stripped[:60],
                        "fixable": True,
                    }
                )
                continue

        # 检查与间隔1行的前一行是否重复（中间是空行）
        if i >= 2:
            between = lines[i - 1].strip()
            prev2 = lines[i - 2].strip()
            if between == "" and prev2 == stripped and prev2:
                issues.append(
                    {
                        "type": "重复行",
                        "line": i + 1,
                        "content": stripped[:60],
                        "fixable": True,
                    }
                )

    return issues


def check_numbered_lists(lines: list[str]) -> list[dict]:
    """检测有序列表（数字. 开头的行），排除代码块和表格内"""
    issues = []
    in_block = False
    block_start = -1
    block_count = 0

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i) or is_in_table(line):
            if in_block:
                issues.append(
                    {
                        "type": "numbered_list",
                        "line_start": block_start + 1,
                        "line_end": block_start + block_count,
                        "count": block_count,
                        "fixable": BULLET_MODULE_AVAILABLE,
                    }
                )
                in_block = False
                block_count = 0
            continue

        if NUMBERED_LIST_RE.match(line):
            if not in_block:
                in_block = True
                block_start = i
                block_count = 0
            block_count += 1
        else:
            if in_block:
                issues.append(
                    {
                        "type": "numbered_list",
                        "line_start": block_start + 1,
                        "line_end": block_start + block_count,
                        "count": block_count,
                        "fixable": BULLET_MODULE_AVAILABLE,
                    }
                )
                in_block = False
                block_count = 0

    # 文件末尾仍在 block 中
    if in_block:
        issues.append(
            {
                "type": "numbered_list",
                "line_start": block_start + 1,
                "line_end": block_start + block_count,
                "count": block_count,
                "fixable": BULLET_MODULE_AVAILABLE,
            }
        )

    return issues


def check_table_intro(lines: list[str], min_chars: int = 80) -> list[dict]:
    """检查表格前的引导段落是否充分（≥min_chars字）

    公文规范要求表格前必须有充分的引导段落，展开政策依据、适用性说明、
    本项目情况等，禁止"具体如下表所示"式的一句话引导。
    """
    issues = []
    in_code = False

    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue

        # 检测表格起始行（| xxx | xxx | 且下一行是分隔行 |---|---|）
        if not (stripped.startswith("|") and stripped.endswith("|")):
            continue
        if i + 1 >= len(lines):
            continue
        _ = lines[i + 1].strip()
        if not _is_table_separator(lines[i + 1]):
            continue

        # 找到一个表格，向上搜索引导段落
        intro_text = ""
        for j in range(i - 1, max(i - 5, -1), -1):
            prev = lines[j].strip()
            if not prev:
                continue
            if prev.startswith("#"):
                # 遇到标题，说明表格紧跟标题后面，引导语缺失
                break
            if prev.startswith("|"):
                # 上面还有表格，跳过
                break
            intro_text = prev
            break

        # 计算引导语长度（去除 markdown 标记）
        clean_intro = re.sub(r"[*#>\[\]`]", "", intro_text)
        if len(clean_intro) < min_chars:
            issues.append(
                {
                    "type": "表格引导语",
                    "line": i + 1,
                    "intro_len": len(clean_intro),
                    "intro_text": intro_text[:60],
                    "min_chars": min_chars,
                    "fixable": False,  # 需要 API 或人工扩写
                }
            )

    return issues


def check_scoring_alignment(lines: list[str], scoring_file: Path) -> list[dict]:
    """检查标书内容是否响应了评分标准中的每一项"""
    issues = []

    if not scoring_file.exists():
        show_error(f"评分标准文件不存在: {scoring_file}")
        return issues

    scoring_text = scoring_file.read_text(encoding="utf-8")
    content_text = "\n".join(lines)

    # 解析评分项：匹配 "第X项" 或 序号+内容+分值 的模式
    # 支持格式：
    #   1. xxx（5分）
    #   | 1 | xxx | 5 |
    #   第1项（5分）：xxx
    scoring_items = []

    # 模式1：序号. 内容（X分）
    for m in re.finditer(r"(\d+)[.、]\s*(.+?)[（(](\d+)分[）)]", scoring_text):
        scoring_items.append(
            {
                "index": int(m.group(1)),
                "desc": m.group(2).strip(),
                "score": int(m.group(3)),
            }
        )

    # 模式2：表格行 | 序号 | 内容 | 分值 |
    if not scoring_items:
        for m in re.finditer(r"\|\s*(\d+)\s*\|\s*(.+?)\s*\|\s*(\d+)\s*\|", scoring_text):
            scoring_items.append(
                {
                    "index": int(m.group(1)),
                    "desc": m.group(2).strip(),
                    "score": int(m.group(3)),
                }
            )

    for item in scoring_items:
        # 提取关键词（取描述中长度 >= 2 的中文词组）
        keywords = re.findall(r"[\u4e00-\u9fff]{2,}", item["desc"])
        # 至少有一半关键词出现在正文中才算响应
        if not keywords:
            continue
        matched = sum(1 for kw in keywords if kw in content_text)
        ratio = matched / len(keywords)
        item["responded"] = ratio >= 0.5
        issues.append(
            {
                "type": "评分对齐",
                "index": item["index"],
                "score": item["score"],
                "desc": item["desc"][:30],
                "responded": item["responded"],
                "fixable": False,
            }
        )

    return issues


# ── 修复函数 ────────────────────────────────────────────────


def fix_forbidden_words(text: str) -> str:
    """替换禁用词"""
    # 先处理 "我们" 的上下文敏感替换
    # "我们团队/公司/单位" → "本项目团队/本单位/本单位"
    text = re.sub(r"我们(团队)", r"本项目\1", text)
    text = re.sub(r"我们(公司|单位)", r"本\1", text)
    # 剩余的独立 "我们" → "本项目团队"
    text = text.replace("我们", "本项目团队")

    # 其他禁用词直接替换
    text = text.replace("确保", "保障")
    text = text.replace("我司", "本单位")
    return text


def _get_llm_client():
    """保留接口兼容，返回占位对象（已改用 claude CLI）"""
    return True  # non-None so callers proceed


def fix_bullet_points(text: str) -> str:
    """调用 bullet_to_paragraph 模块修复 bullet point"""
    if not BULLET_MODULE_AVAILABLE:
        show_warning("bullet_to_paragraph 模块不可用，跳过 bullet point 修复")
        return text

    client = _get_llm_client()
    if client is None:
        show_warning("API 不可用，跳过 bullet point 修复")
        return text

    blocks = extract_bullet_blocks(text)
    if not blocks:
        return text

    # 从后往前替换，避免偏移
    lines = text.split("\n")
    for block in reversed(blocks):
        converted = convert_bullet_block(client, block)
        start = block["start_line"]
        end = block["end_line"] + 1
        lines[start:end] = converted.split("\n")

    return "\n".join(lines)


def fix_data_sources(text: str, issues: list[dict]) -> str:
    """在数据所在行末尾插入来源标注 <!-- SOURCE: XXX -->"""
    lines = text.split("\n")

    # 按行号从大到小处理，避免偏移
    sorted_issues = sorted(issues, key=lambda x: x["line"], reverse=True)

    for issue in sorted_issues:
        line_idx = issue["line"] - 1
        source = issue.get("source_type", "PLACEHOLDER")
        line = lines[line_idx].rstrip()
        lines[line_idx] = f"{line} <!-- SOURCE: {source} -->"

    return "\n".join(lines)


def fix_duplicate_lines(text: str, issues: list[dict]) -> str:
    """删除重复行（保留第一次出现的）"""
    if not issues:
        return text
    lines = text.split("\n")
    # 收集需要删除的行号（0-based）
    lines_to_remove = set()
    for issue in issues:
        line_idx = issue["line"] - 1
        lines_to_remove.add(line_idx)
        # 如果重复行与前一行之间有空行（间隔1行的情况），也删除中间的空行
        if line_idx >= 2 and lines[line_idx - 1].strip() == "":
            prev2 = lines[line_idx - 2].strip()
            if prev2 == lines[line_idx].strip():
                lines_to_remove.add(line_idx - 1)

    result = [line for i, line in enumerate(lines) if i not in lines_to_remove]
    return "\n".join(result)


def fix_numbered_lists(text: str) -> str:
    """调用 bullet_to_paragraph 模块修复有序列表（与 bullet point 相同处理方式）"""
    if not BULLET_MODULE_AVAILABLE:
        show_warning("bullet_to_paragraph 模块不可用，跳过有序列表修复")
        return text

    client = _get_llm_client()
    if client is None:
        show_warning("API 不可用，跳过有序列表修复")
        return text

    blocks = extract_bullet_blocks(text)
    # 只处理 numbered 类型的块
    numbered_blocks = [b for b in blocks if b.get("block_type") == "numbered"]
    if not numbered_blocks:
        return text

    lines = text.split("\n")
    for block in reversed(numbered_blocks):
        converted = convert_bullet_block(client, block)
        start = block["start_line"]
        end = block["end_line"] + 1
        lines[start:end] = converted.split("\n")

    return "\n".join(lines)


# ── 报告输出 ────────────────────────────────────────────────


def format_report(
    filepath: str,
    forbidden: list,
    bullets: list,
    data_sources: list,
    scoring: list,
    duplicates: list = None,
    numbered: list = None,
    table_intros: list = None,
) -> str:
    """格式化输出检查报告"""
    if duplicates is None:
        duplicates = []
    if numbered is None:
        numbered = []
    if table_intros is None:
        table_intros = []

    parts = []
    parts.append(f"文件: {filepath}")
    parts.append("")

    total = 0
    fixable = 0

    # 禁用词
    if forbidden:
        parts.append(f"[禁用词] 发现 {len(forbidden)} 处")
        for item in forbidden:
            parts.append(f'  L{item["line"]}: "...{item["context"]}..." → 建议改为 "{item["replacement"]}"')
        total += len(forbidden)
        fixable += sum(1 for i in forbidden if i["fixable"])
    else:
        parts.append("[禁用词] 无问题")
    parts.append("")

    # Bullet point
    if bullets:
        bullet_count = sum(b["count"] for b in bullets)
        parts.append(f"[Bullet Point] 发现 {bullet_count} 处")
        for item in bullets:
            parts.append(f"  L{item['line_start']}-{item['line_end']}: {item['count']} 个 bullet point")
        total += bullet_count
        fixable += sum(b["count"] for b in bullets if b["fixable"])
    else:
        parts.append("[Bullet Point] 无问题")
    parts.append("")

    # 重复行
    if duplicates:
        parts.append(f"[重复行] 发现 {len(duplicates)} 处")
        for item in duplicates:
            parts.append(f'  L{item["line"]}: "{item["content"]}"')
        total += len(duplicates)
        fixable += sum(1 for i in duplicates if i["fixable"])
    else:
        parts.append("[重复行] 无问题")
    parts.append("")

    # 有序列表
    if numbered:
        numbered_count = sum(n["count"] for n in numbered)
        parts.append(f"[有序列表] 发现 {numbered_count} 处")
        for item in numbered:
            parts.append(f"  L{item['line_start']}-{item['line_end']}: {item['count']} 个有序列表项")
        total += numbered_count
        fixable += sum(n["count"] for n in numbered if n["fixable"])
    else:
        parts.append("[有序列表] 无问题")
    parts.append("")

    # 数据来源
    if data_sources:
        # 按来源类型分组统计
        by_type = {}
        for item in data_sources:
            st = item.get("source_type", "PLACEHOLDER")
            by_type.setdefault(st, []).append(item)
        parts.append(f"[数据来源] {len(data_sources)} 处需要标注")
        for st, items in by_type.items():
            parts.append(f"  → {st}: {len(items)} 处")
        for item in data_sources:
            st = item.get("source_type", "PLACEHOLDER")
            parts.append(f'  L{item["line"]}: "{item["data"]}" → {st}')
        total += len(data_sources)
        fixable += len(data_sources)
    else:
        parts.append("[数据来源] 无问题")
    parts.append("")

    # 表格引导语
    if table_intros:
        parts.append(f"[表格引导语] {len(table_intros)} 处引导语过短")
        for item in table_intros:
            parts.append(
                f"  L{item['line']}: 引导语仅{item['intro_len']}字"
                f"（要求≥{item['min_chars']}字）"
                f' — "{item["intro_text"]}..."'
            )
        total += len(table_intros)
    else:
        parts.append("[表格引导语] 无问题")
    parts.append("")

    # 评分对齐
    if scoring:
        parts.append("[评分对齐]")
        for item in scoring:
            mark = "+" if item["responded"] else "x"
            parts.append(
                f"  {mark} 第{item['index']}项（{item['score']}分）: "
                f"{'已响应' if item['responded'] else '未找到对应内容'} "
                f"— {item['desc']}"
            )
        not_responded = [i for i in scoring if not i["responded"]]
        total += len(not_responded)
        parts.append("")

    # 汇总
    manual = total - fixable
    parts.append("=" * 40)
    parts.append(f"总计: {total} 个问题")
    parts.append(f"  可自动修复: {fixable} 个")
    parts.append(f"  需人工处理: {manual} 个")
    parts.append("=" * 40)

    return "\n".join(parts)


# ── 主流程 ──────────────────────────────────────────────────


def process_file(filepath: Path, args) -> dict:
    """处理单个文件，返回统计信息"""
    text = filepath.read_text(encoding="utf-8")
    lines = text.split("\n")

    # 执行检查
    forbidden = check_forbidden_words(lines)
    bullets = check_bullet_points(lines)
    duplicates = check_duplicate_lines(lines)
    numbered = check_numbered_lists(lines)
    data_sources = check_data_sources(lines)
    table_intros = check_table_intro(lines)

    scoring = []
    if args.bid and args.scoring:
        scoring = check_scoring_alignment(lines, Path(args.scoring))

    # 输出报告
    report = format_report(
        str(filepath),
        forbidden,
        bullets,
        data_sources,
        scoring,
        duplicates=duplicates,
        numbered=numbered,
        table_intros=table_intros,
    )
    print(f"\n{'=' * 3} 质量检查报告 {'=' * 3}\n")
    print(report)

    # 统计
    total_issues = (
        len(forbidden)
        + sum(b["count"] for b in bullets)
        + len(duplicates)
        + sum(n["count"] for n in numbered)
        + len(data_sources)
        + len(table_intros)
        + len([i for i in scoring if not i["responded"]])
    )
    total_fixable = (
        sum(1 for i in forbidden if i["fixable"])
        + sum(b["count"] for b in bullets if b["fixable"])
        + sum(1 for i in duplicates if i["fixable"])
        + sum(n["count"] for n in numbered if n["fixable"])
        + sum(1 for i in data_sources if i.get("fixable"))
    )

    # 自动修复
    if args.fix and total_fixable > 0:
        fixed_text = fix_forbidden_words(text)
        # 重复行修复（在其他修复之前，避免行号偏移）
        if duplicates:
            fixed_text = fix_duplicate_lines(fixed_text, duplicates)
        if BULLET_MODULE_AVAILABLE:
            fixed_text = fix_bullet_points(fixed_text)
            fixed_text = fix_numbered_lists(fixed_text)
        # 数据来源标注（需在重复行删除后重新计算行号，此处简化处理）
        if data_sources and not duplicates:
            # 仅在无重复行删除时才按原始行号标注，否则行号已偏移
            fixed_text = fix_data_sources(fixed_text, data_sources)
        elif data_sources and duplicates:
            # 重复行删除后行号已变，重新检测数据来源
            new_lines = fixed_text.split("\n")
            new_data_sources = check_data_sources(new_lines)
            if new_data_sources:
                fixed_text = fix_data_sources(fixed_text, new_data_sources)

        # 确定输出路径
        if args.output_dir:
            out_dir = Path(args.output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / filepath.name
        else:
            out_path = filepath

        out_path.write_text(fixed_text, encoding="utf-8")
        show_success(f"已修复并保存: {out_path}")

    return {"total": total_issues, "fixable": total_fixable, "file": str(filepath)}


def main():
    parser = argparse.ArgumentParser(
        description="报告/标书质量检查 + 自动修复",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  %(prog)s 01.md                              报告模式检查
  %(prog)s 01.md --bid --scoring 评分标准.md   标书模式检查
  %(prog)s 01.md --fix                        检查 + 自动修复
  %(prog)s ./md/                              检查整个目录
  %(prog)s ./md/ --bid --scoring s.md --fix --output-dir ./md_fixed/
""",
    )
    parser.add_argument("input", nargs="?", help="MD 文件或目录路径")
    parser.add_argument("--bid", action="store_true", help="标书模式（额外检查评分对齐）")
    parser.add_argument("--scoring", help="评分标准文件路径（配合 --bid 使用）")
    parser.add_argument("--fix", action="store_true", help="自动修复可修复的问题")
    parser.add_argument("--output-dir", help="修复后文件的输出目录（默认覆盖原文件）")
    parser.add_argument("--version", action="store_true", help="显示版本信息")

    args = parser.parse_args()

    if args.version:
        show_version_info(SCRIPT_VERSION, SCRIPT_AUTHOR, SCRIPT_UPDATED)
        return

    if not args.input:
        parser.print_help()
        sys.exit(1)

    if args.bid and not args.scoring:
        show_warning("标书模式建议指定 --scoring 评分标准文件，否则跳过评分对齐检查")

    input_path = Path(args.input)

    if not input_path.exists():
        show_error(f"路径不存在: {input_path}")
        sys.exit(1)

    # 收集文件
    if input_path.is_dir():
        md_files = sorted(input_path.glob("*.md"))
        if not md_files:
            show_error(f"目录中没有 .md 文件: {input_path}")
            sys.exit(1)
        show_info(f"发现 {len(md_files)} 个 MD 文件")
    else:
        md_files = [input_path]

    # 逐文件处理
    all_stats = []
    for f in md_files:
        stats = process_file(f, args)
        all_stats.append(stats)

    # 多文件汇总
    if len(all_stats) > 1:
        grand_total = sum(s["total"] for s in all_stats)
        grand_fixable = sum(s["fixable"] for s in all_stats)
        print(f"\n{'=' * 3} 全部文件汇总 {'=' * 3}")
        print(f"文件数: {len(all_stats)}")
        print(f"总问题: {grand_total}")
        print(f"  可自动修复: {grand_fixable}")
        print(f"  需人工处理: {grand_total - grand_fixable}")
        print("=" * 40)

    # 退出码：有问题返回 1
    total = sum(s["total"] for s in all_stats)
    if total > 0 and not args.fix:
        sys.exit(1)


if __name__ == "__main__":
    main()
