#!/usr/bin/env python3
"""
AI 敏感词扫描器 (scan_sensitive_words.py)

调用智谱 API 扫描标书 .md 文件，智能识别敏感词并维护敏感词列表。

识别类型：
  - 组织名称/公司名称（竞争对手或错误引用）
  - 不符合乙方立场的措辞（过于强硬/绝对的用词）
  - 项目名称交叉污染（A项目文档中出现B项目专有名词）

用法：
  python3 scan_sensitive_words.py <目录路径>
  python3 scan_sensitive_words.py <目录路径> --config path/to/sensitive_words.json
  python3 scan_sensitive_words.py <目录路径> --json
  python3 scan_sensitive_words.py <目录路径> --update
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

# === 路径设置 ===
SCRIPTS_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SCRIPTS_ROOT / "lib"))

from display import show_error, show_info, show_processing, show_success, show_warning
from file_ops import find_files_by_extension

# === 常量 ===
CHUNK_SIZE = 8000  # 按字符数分块


# === API 调用 ===


def create_client():
    """保留接口兼容，返回 None（已改用 claude CLI）"""
    return None


def build_prompt(content: str, existing_words: list[str], filename: str) -> str:
    """构建发送给 AI 的分析 prompt"""
    existing_list = "、".join(existing_words[:50]) if existing_words else "（暂无）"

    return f"""你是一个标书文档审阅专家。请分析以下标书文档内容，识别其中的敏感词。

## 需要识别的敏感词类型

1. **组织名称/公司名称**：可能是竞争对手名称、错误引用的单位名称、不应出现在本标书中的第三方名称。注意区分：招标方/甲方名称是正常的，但其他公司名称可能是错误引用。

2. **不符合乙方立场的措辞**：标书是乙方（投标方）撰写的，以下措辞不恰当：
   - 对甲方提要求或命令的词（如"必须"、"应当"、"要求甲方"）
   - 过于绝对的承诺（如"确保"、"保证"、"杜绝"、"绝对"）
   - 语气过于强硬或紧迫的词（如"亟需"、"迫在眉睫"）
   - 非正式用语（如"我们"、"我司"、"我方"）
   - 甲方口吻的词（如"验收合格后"用在不恰当的语境中）

3. **项目名称交叉污染**：文档中出现的可能属于其他项目的专有名词（地名、项目名、河流名、水库名等），这些词可能是从其他标书复制时遗漏的。

## 重要：以下不算敏感词，请勿报告

- **标书规范用语**：如"采购人""投标人""本项目""本项目团队""本单位""投标方""招标方""中标人""评标委员会"等，这些是标书行业标准术语，不是敏感词。
- **正式行文措辞**：如"根据""按照""依据""鉴于""为此"等连接词，属于公文规范用语。
- **本项目自身名称**：本文件所属项目的名称、地名、河流名等，属于本项目内容，不是交叉污染。

## 已有敏感词列表（无需重复报告）
{existing_list}

## 当前文件名
{filename}

## 文档内容
{content}

## 输出要求

请直接输出纯 JSON 数组（不要用 markdown 代码块包裹，不要加任何解释文字）。
每个元素包含：
- word: 敏感词原文
- category: 类型（"organization" | "wording" | "cross_project"）
- reason: 为什么认为这是敏感词
- suggest: 建议替换为什么（如果适用，否则为空字符串）
- severity: 严重程度（"high" | "medium" | "low"）
- context: 该词出现的上下文片段（前后各约 10 个字）

如果没有发现任何敏感词，直接返回 []

注意：JSON 中的字符串值如果包含引号，请用反斜杠转义。直接输出 JSON，不要包裹在 ```json ``` 中。"""


def _parse_json_response(text: str) -> list[dict] | None:
    """尝试从 AI 返回的文本中解析 JSON 数组，带多层容错"""
    # 第 1 步：strip markdown 代码块
    cleaned = text.strip()
    if cleaned.startswith("```"):
        # 去掉 ```json 或 ``` 开头行和结尾 ```
        cleaned = re.sub(r"^```(?:json)?\s*\n?", "", cleaned)
        cleaned = re.sub(r"\n?```\s*$", "", cleaned)
        cleaned = cleaned.strip()

    # 第 2 步：直接尝试 json.loads
    try:
        result = json.loads(cleaned)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # 第 3 步：用 regex 提取最外层 JSON 数组
    match = re.search(r"\[[\s\S]*\]", cleaned)
    if match:
        try:
            result = json.loads(match.group())
            if isinstance(result, list):
                return result
        except json.JSONDecodeError:
            pass

    return None


def call_api(client, prompt: str, max_retries: int = 2) -> list[dict] | None:
    """调用 Claude 分析内容，带 retry 和容错解析"""
    from llm_client import chat as llm_chat

    for attempt in range(1, max_retries + 1):
        try:
            text = llm_chat(
                system="你是标书文档审阅专家。只输出纯 JSON 数组，不要 markdown 代码块。",
                message=prompt,
            )

            result = _parse_json_response(text)
            if result is not None:
                return result

            if attempt < max_retries:
                show_warning(f"  JSON 解析失败（第 {attempt} 次），重试中...")
                continue
            else:
                show_warning(f"  JSON 解析失败（已重试 {max_retries} 次），跳过此块")
                show_warning(f"  原始返回: {text[:200]}")
                return None

        except Exception as e:
            show_error(f"API 调用失败: {e}")
            if attempt < max_retries:
                show_warning(f"  第 {attempt} 次失败，重试中...")
                continue
            return None


# === 配置文件操作 ===


def find_config(scan_dir: str) -> str | None:
    """从扫描目录向上查找 sensitive_words.json"""
    current = Path(scan_dir).resolve()
    while current != current.parent:
        config = current / "sensitive_words.json"
        if config.is_file():
            return str(config)
        current = current.parent
    return None


def load_config(config_path: str) -> dict:
    """加载 sensitive_words.json"""
    try:
        with open(config_path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        show_error(f"无法读取配置文件 {config_path}: {e}")
        sys.exit(1)


def get_existing_words(config: dict) -> list[str]:
    """从配置中提取所有已有敏感词"""
    words = []
    for item in config.get("forbidden_words", []):
        words.append(item["word"])
    for item in config.get("project_isolation", []):
        words.append(item["word"])
    return words


def get_whitelist(config: dict) -> set[str]:
    """从配置中提取白名单词汇"""
    return set(config.get("whitelist", []))


def save_config(config: dict, config_path: str):
    """保存 sensitive_words.json"""
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
        f.write("\n")


# === 文件处理 ===


def chunk_text(text: str, chunk_size: int = CHUNK_SIZE) -> list[str]:
    """按字符数分块，尽量在段落边界切割"""
    if len(text) <= chunk_size:
        return [text]

    chunks = []
    while text:
        if len(text) <= chunk_size:
            chunks.append(text)
            break

        # 在 chunk_size 附近找段落边界
        cut = text.rfind("\n\n", 0, chunk_size)
        if cut == -1 or cut < chunk_size // 2:
            cut = text.rfind("\n", 0, chunk_size)
        if cut == -1 or cut < chunk_size // 2:
            cut = chunk_size

        chunks.append(text[:cut])
        text = text[cut:].lstrip("\n")

    return chunks


# 标书正式术语白名单：这些词不应被报告为敏感词
FORMAL_TERMS_WHITELIST = {
    "采购人",
    "投标人",
    "招标人",
    "中标人",
    "投标方",
    "招标方",
    "本项目",
    "本项目团队",
    "本单位",
    "本公司",
    "评标委员会",
    "采购代理机构",
    "监理单位",
    "建设单位",
    "项目负责人",
    "项目经理",
    "技术负责人",
    "根据",
    "按照",
    "依据",
    "鉴于",
    "为此",
}


def verify_findings(findings: list[dict], content: str, whitelist: set[str] | None = None) -> list[dict]:
    """验证 AI 返回的词是否真的出现在文件内容中，过滤幻觉和白名单/正式术语"""
    # 合并内置正式术语白名单和配置白名单
    combined_whitelist = FORMAL_TERMS_WHITELIST.copy()
    if whitelist:
        combined_whitelist.update(whitelist)

    verified = []
    for f in findings:
        word = f.get("word", "")
        if not word:
            continue
        # 过滤白名单词汇
        if word in combined_whitelist:
            continue
        # 验证词是否真的在文件内容中
        if word in content:
            verified.append(f)
    return verified


def scan_file(client, filepath: Path, existing_words: list[str], whitelist: set[str] | None = None) -> list[dict]:
    """扫描单个文件"""
    try:
        content = filepath.read_text(encoding="utf-8")
    except Exception as e:
        show_warning(f"无法读取 {filepath}: {e}")
        return []

    if not content.strip():
        return []

    chunks = chunk_text(content)
    all_findings = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            show_processing(f"  分块 {i + 1}/{len(chunks)}")

        prompt = build_prompt(chunk, existing_words, filepath.name)
        findings = call_api(client, prompt)

        if findings:
            for f in findings:
                f["file"] = str(filepath)
            all_findings.extend(findings)

    # 过滤 AI 幻觉：验证词是否真的出现在文件中，同时过滤白名单
    verified = verify_findings(all_findings, content, whitelist)
    if len(verified) < len(all_findings):
        diff = len(all_findings) - len(verified)
        show_warning(f"  过滤 {diff} 个幻觉词条（文件中不存在）")

    return verified


# === 去重与过滤 ===


def deduplicate_findings(
    findings: list[dict], existing_words: list[str], whitelist: set[str] | None = None
) -> list[dict]:
    """去重并过滤已有的敏感词和白名单词汇"""
    seen = set()
    existing_set = set(existing_words)
    # 合并内置正式术语白名单和配置白名单
    combined_whitelist = FORMAL_TERMS_WHITELIST.copy()
    if whitelist:
        combined_whitelist.update(whitelist)

    result = []

    for f in findings:
        word = f.get("word", "")
        if not word or word in existing_set or word in combined_whitelist:
            continue
        key = (word, f.get("category", ""))
        if key in seen:
            continue
        seen.add(key)
        result.append(f)

    return result


# === 输出格式 ===

CATEGORY_LABELS = {
    "organization": "组织名称",
    "wording": "措辞问题",
    "cross_project": "项目串名",
}

SEVERITY_LABELS = {
    "high": "高",
    "medium": "中",
    "low": "低",
}


def print_table(findings: list[dict]):
    """人类可读的表格输出"""
    if not findings:
        show_success("未发现新的可疑敏感词。")
        return

    print(f"\n{'=' * 80}")
    print(f"  发现 {len(findings)} 个可疑敏感词")
    print(f"{'=' * 80}\n")

    for i, f in enumerate(findings, 1):
        cat = CATEGORY_LABELS.get(f.get("category", ""), f.get("category", "未知"))
        sev = SEVERITY_LABELS.get(f.get("severity", ""), f.get("severity", ""))
        word = f.get("word", "")
        reason = f.get("reason", "")
        suggest = f.get("suggest", "")
        context = f.get("context", "")
        filepath = f.get("file", "")

        print(f"  [{i:2d}] {word}")
        print(f"       类型: {cat}  |  严重程度: {sev}")
        print(f"       原因: {reason}")
        if suggest:
            print(f"       建议: -> {suggest}")
        if context:
            print(f"       上下文: ...{context}...")
        if filepath:
            print(f"       文件: {Path(filepath).name}")
        print()

    print(f"{'=' * 80}")


def print_json(findings: list[dict]):
    """结构化 JSON 输出"""
    print(json.dumps(findings, ensure_ascii=False, indent=2))


# === 更新模式 ===


def interactive_update(findings: list[dict], config: dict, config_path: str):
    """交互式确认并更新 sensitive_words.json"""
    if not findings:
        show_info("没有新词需要添加。")
        return

    added_count = 0
    print("\n逐条确认是否加入敏感词表（y=加入 / n=跳过 / q=退出）：\n")

    for f in findings:
        word = f.get("word", "")
        cat = CATEGORY_LABELS.get(f.get("category", ""), f.get("category", ""))
        reason = f.get("reason", "")
        suggest = f.get("suggest", "")

        print(f"  [{cat}] {word}")
        print(f"    原因: {reason}")
        if suggest:
            print(f"    建议替换: {suggest}")

        try:
            choice = input("    加入词表? (y/n/q): ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if choice == "q":
            break
        elif choice == "y":
            category = f.get("category", "")
            if category == "cross_project":
                config.setdefault("project_isolation", []).append(
                    {
                        "word": word,
                        "reason": reason,
                    }
                )
            else:
                entry = {"word": word, "reason": reason}
                if suggest:
                    entry["suggest"] = suggest
                config.setdefault("forbidden_words", []).append(entry)
            added_count += 1
            show_success(f"已添加: {word}")
        print()

    if added_count > 0:
        save_config(config, config_path)
        show_success(f"已更新 {config_path}，新增 {added_count} 个敏感词。")
    else:
        show_info("未添加任何新词。")


# === 主流程 ===


def main():
    parser = argparse.ArgumentParser(
        description="AI 敏感词扫描器 - 调用智谱 API 扫描标书文档识别敏感词",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  %(prog)s ~/Dev/Work/zdwp/docs/                    # 扫描目录，输出表格
  %(prog)s ~/Dev/Work/zdwp/docs/ --json              # 输出 JSON
  %(prog)s ~/Dev/Work/zdwp/docs/ --update            # 扫描后交互式更新词表
  %(prog)s ~/Dev/Work/zdwp/docs/ --config words.json # 指定配置文件
        """,
    )
    parser.add_argument("directory", help="要扫描的目录路径")
    parser.add_argument("--config", help="sensitive_words.json 路径（默认自动向上查找）")
    parser.add_argument("--update", action="store_true", help="交互式确认后更新敏感词表")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出（面向 CC）")

    args = parser.parse_args()

    # 验证目录
    scan_dir = Path(args.directory).resolve()
    if not scan_dir.is_dir():
        show_error(f"目录不存在: {scan_dir}")
        sys.exit(1)

    # 查找配置文件
    config_path = args.config
    if not config_path:
        config_path = find_config(str(scan_dir))

    config = {}
    existing_words = []
    whitelist = set()
    if config_path:
        config = load_config(config_path)
        existing_words = get_existing_words(config)
        whitelist = get_whitelist(config)
        if not args.json:
            show_info(f"配置文件: {config_path}")
            show_info(f"已有敏感词: {len(existing_words)} 个")
            if whitelist:
                show_info(f"白名单词汇: {len(whitelist)} 个")
    else:
        if not args.json:
            show_warning("未找到 sensitive_words.json，将从零开始扫描。")
        if args.update:
            show_error("--update 模式需要配置文件，请用 --config 指定。")
            sys.exit(1)

    # 收集 .md 文件
    md_files = find_files_by_extension(scan_dir, "md", recursive=True)
    if not md_files:
        show_warning(f"目录中没有 .md 文件: {scan_dir}")
        sys.exit(0)

    if not args.json:
        show_info(f"找到 {len(md_files)} 个 .md 文件")

    # 创建 API 客户端
    client = create_client()

    # 逐文件扫描
    all_findings = []
    for i, filepath in enumerate(md_files, 1):
        if not args.json:
            show_processing(f"扫描 ({i}/{len(md_files)}): {filepath.name}")
        findings = scan_file(client, filepath, existing_words, whitelist)
        all_findings.extend(findings)

    # 去重
    unique_findings = deduplicate_findings(all_findings, existing_words, whitelist)

    if not args.json:
        show_info(f"AI 共返回 {len(all_findings)} 个词条，去重后 {len(unique_findings)} 个新词")

    # 输出
    if args.json:
        print_json(unique_findings)
    else:
        print_table(unique_findings)

    # 更新模式
    if args.update:
        interactive_update(unique_findings, config, config_path)


if __name__ == "__main__":
    main()
