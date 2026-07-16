#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_identity_gate.py — 标书终稿身份泄漏门（只读校验，不改 docx）。

场景：浙江水利院投标工作流，"陪标"（pei 模式）= 通用技术稿严禁出现单位身份；
主标（main 模式）允许实名与自有业绩，但仍禁工具痕迹与协作署名。

扫描范围 = docx 全部 .xml/.rels part（正文/页眉页脚/docProps 元数据）+ 文件名本身。

用法:
    python3 bid_identity_gate.py <docx路径> [--mode main|pei] [--rules <yaml路径>]

exit 0 = PASS；exit 2 = 有 findings；exit 1 = 用法/IO 错误。
stdout 最后一行必是 "PASS" 或 "FAIL <n> findings"。

规则 YAML（--rules，可缺省）只取 identity_banned: [公司全名/人名...] 做项目级增补。
参考: shaoxing-eco-flow-2026/scripts/gen_bid_docx.py::assert_no_identity（此处独立成完整 CLI）。
"""
import argparse
import re
import sys
import zipfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    etree = None

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag):
    return "{%s}%s" % (W_NS, tag)


# ── 内置禁词 ────────────────────────────────────────────────────────────────
# pei 模式全启用；main 模式跳过实名类（公司名/设计院/院自指/业绩归属——主标允许实名和自有业绩）
BANNED_PEI = [
    "我院", "本院", "我们", "我司",
    "勘测设计院", "设计院", "浙水院", "有限责任公司",
    "python-docx",
    # 业绩归属表述 = 陪标身份泄漏（换单位即假业绩）
    "移植投标人", "借鉴投标人", "投标人已在", "投标人承担浙江省内", "投标人全省",
    # 协作署名
    "数字化团队",
]
BANNED_MAIN = ["python-docx", "数字化团队"]

# 裸「院」自指检测的保护词（先剥再查残留「院」，仅 word/document.xml 段文本层）
YUAN_PROTECT = ["院士", "医院", "法院", "科学院", "研究院", "学院", "剧院", "国务院"]

# 院内自指变体（内置于裸院检测覆盖面说明；院内/院级/院总工 含「院」故裸院检测天然覆盖）

# docProps/core.xml 需为空的字段（pei fatal / main warning）
CORE_FIELDS = ["creator", "lastModifiedBy", "description"]

# WordprocessingML 文本 part（跨 run 拼接后扫，防禁词被 run 切断漏检）
WORDML_TEXT_RE = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")

# 文件名拼音人名缩写前缀（如 ztl714.docx / zhl-v2.docx）→ 两模式 warning
PINYIN_PREFIX_RE = re.compile(r"^([a-zA-Z]{2,3})(?=[\d_\-.]|$)")


# ── 规则 YAML（stdlib 极简解析，够用即可；有 PyYAML 则优先） ─────────────────
def _mini_yaml(text):
    """只支持本 gate 需要的子集: 顶级 `key:` + 缩进 `- 字符串` 列表（可带引号）。"""
    data, key = {}, None
    for raw in text.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        m = re.match(r"^([A-Za-z_][\w]*)\s*:\s*(.*)$", line)
        if m and not line.startswith((" ", "\t", "-")):
            key = m.group(1)
            rest = m.group(2).strip()
            if rest.startswith("[") and rest.endswith("]"):
                items = [x.strip().strip("'\"") for x in rest[1:-1].split(",") if x.strip()]
                data[key] = items
                key = None
            else:
                data[key] = []
            continue
        if key is not None and stripped.startswith("- "):
            item = stripped[2:].strip()
            if item.startswith("[") and item.endswith("]"):
                data[key].append([x.strip().strip("'\"") for x in item[1:-1].split(",")])
            else:
                data[key].append(item.strip().strip("'\""))
    return data


def load_rules(path):
    if path is None:
        return {}
    p = Path(path)
    if not p.is_file():
        print("规则文件不存在: %s" % p, file=sys.stderr)
        sys.exit(1)
    text = p.read_text(encoding="utf-8")
    try:
        import yaml  # type: ignore
        return yaml.safe_load(text) or {}
    except ImportError:
        return _mini_yaml(text)


# ── 扫描 ────────────────────────────────────────────────────────────────────
def extract_wordml_text(xml_str):
    return "".join(WORDML_TEXT_RE.findall(xml_str))


def scan_banned(parts, banned, findings):
    """全 part 扫禁词。WordprocessingML part 用跨 run 拼接文本，其余用原始 XML。"""
    for name, xml in sorted(parts.items()):
        hay = xml
        if "<w:t" in xml:
            hay = extract_wordml_text(xml)
        for word in banned:
            c = hay.count(word)
            if c:
                findings.append("[身份] %s %s x%d" % (name, word, c))


def scan_bare_yuan(doc_xml_bytes, banned_active, findings):
    """裸「院」自指（仅 word/document.xml 段文本层）。先剥保护词与已单报的含院禁词。"""
    if etree is None:
        # 无 lxml 兜底: 整文层面 regex（无段号）
        text = extract_wordml_text(doc_xml_bytes.decode("utf-8", "ignore"))
        for pr in _yuan_strip_list(banned_active):
            text = text.replace(pr, "")
        if "院" in text:
            findings.append("[身份] word/document.xml 裸院 x%d（无 lxml，段号不可用）" % text.count("院"))
        return
    root = etree.fromstring(doc_xml_bytes)
    strip_list = _yuan_strip_list(banned_active)
    for i, p in enumerate(root.iter(w("p")), 1):
        txt = "".join(t.text or "" for t in p.iter(w("t")))
        if "院" not in txt:
            continue
        cleaned = txt
        for pr in strip_list:
            cleaned = cleaned.replace(pr, "")
        for m in re.finditer("院", cleaned):
            s, e = max(0, m.start() - 15), min(len(cleaned), m.end() + 15)
            findings.append("[身份] P%d 裸院: …%s…" % (i, cleaned[s:e]))


def _yuan_strip_list(banned_active):
    """裸院检测前要剥掉的词: 保护词 + 已由禁词扫描单独上报的含「院」禁词（防重复计）。长词优先。"""
    lst = list(YUAN_PROTECT) + [b for b in banned_active if "院" in b]
    return sorted(set(lst), key=len, reverse=True)


def scan_core_props(parts, mode, findings, warnings):
    core = parts.get("docProps/core.xml")
    if core is None:
        return
    for field in CORE_FIELDS:
        m = re.search(r"<(?:\w+:)?%s[^>]*>([^<]+)</(?:\w+:)?%s>" % (field, field), core)
        if m and m.group(1).strip():
            line = "[元数据] docProps/core.xml %s=%s（应为空）" % (field, m.group(1).strip())
            if mode == "pei":
                findings.append(line)
            else:
                warnings.append(line)


def scan_filename(docx, banned, findings, warnings):
    name = docx.name
    for word in banned:
        if word in name:
            findings.append("[身份] <文件名> %s x1" % word)
    m = PINYIN_PREFIX_RE.match(docx.stem)
    if m:
        warnings.append("[文件名] %s 前缀 '%s' 疑似拼音人名缩写 → 建议改中性名（warning，不计红）" % (name, m.group(1)))


# ── 主流程 ──────────────────────────────────────────────────────────────────
def run(docx, mode, rules_path):
    if not docx.is_file():
        print("文件不存在: %s" % docx, file=sys.stderr)
        sys.exit(1)
    rules = load_rules(rules_path)
    extra = [x for x in (rules.get("identity_banned") or []) if x]

    if mode == "pei":
        banned = list(BANNED_PEI) + [x for x in extra if x not in BANNED_PEI]
    else:
        # main: 实名类（公司名/设计院/院自指/业绩归属/项目级实名增补）全跳过
        banned = list(BANNED_MAIN)

    try:
        with zipfile.ZipFile(str(docx)) as z:
            parts_bytes = {n: z.read(n) for n in z.namelist()
                           if n.endswith(".xml") or n.endswith(".rels")}
    except zipfile.BadZipFile:
        print("非法 docx（不是 zip）: %s" % docx, file=sys.stderr)
        sys.exit(1)
    parts = {n: b.decode("utf-8", "ignore") for n, b in parts_bytes.items()}

    findings, warnings = [], []
    scan_banned(parts, banned, findings)
    if mode == "pei":
        doc_bytes = parts_bytes.get("word/document.xml")
        if doc_bytes is not None:
            scan_bare_yuan(doc_bytes, banned, findings)
    scan_core_props(parts, mode, findings, warnings)
    scan_filename(docx, banned, findings, warnings)

    # ── 报告 ──
    print("bid_identity_gate · %s · mode=%s · rules=%s · 内置禁词 %d + 增补 %d"
          % (docx.name, mode, rules_path or "(无)",
             len(BANNED_PEI) if mode == "pei" else len(BANNED_MAIN),
             len(extra) if mode == "pei" else 0))
    print("扫描 part 数: %d（.xml/.rels）+ 文件名" % len(parts))
    for f in findings:
        print(f)
    for wline in warnings:
        print("[warn] " + wline)
    if findings:
        print("FAIL %d findings" % len(findings))
        sys.exit(2)
    print("PASS")
    sys.exit(0)


def main():
    ap = argparse.ArgumentParser(description="标书终稿身份泄漏门（只读）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--mode", choices=["main", "pei"], default="pei",
                    help="pei=陪标(默认,全禁词+裸院+元数据fatal) / main=主标(仅工具痕迹+协作署名)")
    ap.add_argument("--rules", default=None, help="项目级规则 YAML（identity_banned 增补）")
    args = ap.parse_args()
    run(args.docx, args.mode, args.rules)


if __name__ == "__main__":
    main()
