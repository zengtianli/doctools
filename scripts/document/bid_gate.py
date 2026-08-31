#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_gate.py — 标书终稿门检族（原 bid_* 6 个入口脚本 2026-07-31 合并为子命令）。

子命令（argparse 表面与原脚本零改动，检测逻辑 SSOT 仍在同目录 bid_residue_lib.py）:
  run      <docx> [--mode main|pei] [--rules Y] [--apply]   原 bid_final.py（driver：扫描→清理→身份门→付印门）
  scan     <docx> [--mode] [--rules]                        原 bid_residue_scan.py（8 类残留只读扫描）
  sweep    <docx> [--mode] [--rules] [--check|--apply]      原 bid_finalize_sweep.py（类 1-7 确定性清理，三护栏）
  identity <docx> [--mode] [--rules]                        原 bid_identity_gate.py（第 8 类身份泄漏只读门）
  print    <docx> [--mode] [--rules]                        原 bid_print_ready.py（付印只读门）
  toc      <项目目录|docx> [--chapters Y]                    bid_toc_gate.py（目录 ≡ 招标表，2026-08-30 立）
  deref    <target> [--check] [--md] [--manual-pairs Y]     原 bid_deref.py（交叉引用去耦合，表面与众不同：无 --mode）

exit 契约不变: 0 = PASS · 2 = 有门红/待人判 · 1 = 用法/IO 错误。
run 的 stdout 是被机器消费的契约（work_ops.cmd_bidgate 正则解析「门控汇总:」与
「sha256 <12hex>」两行）——banner 保留旧脚本名字面量、汇总行逐字不动。
driver 由 subprocess 改为直接函数调用，用 redirect_stdout/redirect_stderr + rstrip
复刻旧 capture_output 语义，逐字节等价（真件两模式 diff 验证）。

surgical 写盘（sweep --apply / deref apply）后挂 docx_parts.assert_parts_intact
部件完整性断言（2026-07-31 立的第二判据）：丢部件/多改部件即抛，不静默。
"""
import argparse
import contextlib
import glob
import hashlib
import io
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bid_residue_lib as lib  # noqa: E402
import bid_toc_gate  # noqa: E402  目录一致性门（2026-08-30 /govern）

sys.path.append(str(Path(__file__).resolve().parents[2] / "lib"))
import caption_re  # noqa: E402  题注判据 SSOT
from docx_parts import assert_parts_intact  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_NS = W
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def w(t):
    return "{%s}%s" % (W, t)


def ln(e):
    return etree.QName(e).localname


# ════════════════════════════════════════════════════════════════════════════
# scan — 原 bid_residue_scan.py：标书终稿 8 类残留【只读】扫描器
#
# 逐段（lxml 解 word/document.xml，段=w:p，段文本=拼接 w:t）匹配内置 taxonomy
# + --rules YAML 项目增补，输出分类报告：每条 [类别N] P<段号4位> 摘录(前80字) 处置建议。
# 另扫 docProps/core.xml 与全部 XML part（第8类身份泄漏）。
# --mode pei（默认）全量 8 类；main 跳过第8类实名类（公司名/院自指/业绩归属/
#        裸院/元数据署名），仍报工具痕迹（python-docx）。
# ════════════════════════════════════════════════════════════════════════════

def scan_apply_path(docx_path, args=None) -> dict:
    """只读扫 8 类残留，返回 findings 报告；不开 Document、不写盘、不 sys.exit。

    为什么是 apply_path 而不是 apply(doc, args)：第 8 类身份泄漏要扫 docProps/core.xml、
    页眉页脚、styles.xml 这些 document.xml 以外的部件，python-docx 的 Document 对象根本
    看不到它们 —— 拿 doc 当入口只能扫到正文，覆盖面会静默缩水一半，正是最该抓的那一半。

    args 取 mode / rules 两个属性（Namespace 或 None 都行）。
    """
    mode = getattr(args, "mode", "pei")
    rules = lib.load_rules(getattr(args, "rules", None))
    _, parts = lib.load_parts(docx_path)
    if "word/document.xml" not in parts:
        raise ValueError(f"不是有效 docx（缺 word/document.xml）: {docx_path}")
    findings = lib.scan_parts(parts, mode=mode, rules=rules)
    by_cat = {}
    for f in findings:
        by_cat[f["cat"]] = by_cat.get(f["cat"], 0) + 1
    return {"changed": 0, "findings": findings, "分类计数": by_cat,
            "残留条数": len(findings), "mode": mode}


def cmd_scan(args) -> int:
    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    try:
        report = scan_apply_path(args.docx, args)
    except (OSError, ValueError, RuntimeError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1
    findings = report["findings"]

    print(f"bid_residue_scan · {args.docx.name} · mode={args.mode}"
          f" · rules={args.rules.name if args.rules else '(内置通用)'}")
    if findings:
        by_cat = report["分类计数"]
        print("分类计数:", " ".join(f"类别{c}({lib.CAT_NAMES[c]})={n}" for c, n in sorted(by_cat.items())))
        for line in lib.format_findings(findings):
            print(line)
        print(f"FAIL {len(findings)} findings")
        return 2
    print("8 类残留全零（协作标记/拟hedge/评分脚手架/内部编号/二次残渣/断裂引用/口径meta/身份泄漏）")
    print("PASS")
    return 0


# ════════════════════════════════════════════════════════════════════════════
# sweep — 原 bid_finalize_sweep.py：标书终稿确定性清理引擎（残留类 1-7）
#
# 第8类身份泄漏不在此改（交 identity 子命令）。surgical：zipfile 逐 entry 复制，
# 只重写 word/document.xml；改前 .bak-YYYYmmdd-HHMMSS 备份。
# 内置行为：
#   ① 剥〔E-xx...〕依据码（含括号内 worklib#/招标段号，整个〔..〕剥除；公文文号〔2025〕不动）
#   ② rules delete_startswith / delete_exact 整段删（评分脚手架块/照抄评标办法裸句）
#   ③ rules exact 跨 run 保格式精确替换
#   ④ rules caption_renumber 题注按正文出现序幂等重编号（占位法防连环；已按序则跳过）
#   ⑤ 通用二次残渣修（，此处）→）· ，本处）→）· ，）→）· 、）→）· （）→删）
# 默认 --check 干跑打印计数不落盘；--apply 才写（备份 + 三护栏红则不写 exit 2）。
# 落盘后自动复扫（bid_residue_lib.scan_parts 类别 1-7），复扫归零才 PASS。
# ════════════════════════════════════════════════════════════════════════════

DEBRIS_PAIRS = [("，此处）", "）"), ("，本处）", "）"), ("，）", "）"), ("、）", "）"), ("（）", "")]


def renumber_captions(root, full_before, prefix):
    """题注前缀按正文首次出现序重编号。幂等：首现序已 = 1..k 则跳过。占位法防连环替换。"""
    pat = re.compile(re.escape(prefix) + r"(\d+)")
    first = {}
    for m in pat.finditer(full_before):
        first.setdefault(int(m.group(1)), m.start())
    if not first:
        return 0
    nums = sorted(first)
    seq = sorted(first, key=lambda n: first[n])
    if seq == nums == list(range(1, len(nums) + 1)):
        return 0  # 已按出现序连续编号，跳过（重复跑不再转乱）
    mapping = {old: rank for rank, old in enumerate(seq, 1)}
    n = 0
    for p in lib.paragraphs(root):
        for old in mapping:
            n += lib.replace_all(p, f"{prefix}{old}", f"{prefix}@{old}@")
        for old, new in mapping.items():
            lib.replace_all(p, f"{prefix}@{old}@", f"{prefix}{new}")
    return n


def sweep(root, rules):
    """in-memory 执行清理，返回计数 dict。"""
    counts = {"删段": 0, "剥E码": 0, "EXACT": 0, "重编号": 0, "残渣修": 0}
    full_before = "".join(lib.ptext(p) for p in lib.paragraphs(root))
    # ② 整段删
    for p in lib.paragraphs(root):
        s = lib.ptext(p).strip()
        if any(s.startswith(x) for x in rules["delete_startswith"]) or s in rules["delete_exact"]:
            p.getparent().remove(p)
            counts["删段"] += 1
    # ①③④⑤ 段内替换
    for p in lib.paragraphs(root):
        counts["剥E码"] += lib.regex_strip_para(p, lib.E_RE, "")
        for old, new in rules["exact"]:
            counts["EXACT"] += lib.replace_all(p, old, new)
    for prefix in rules["caption_renumber"]:
        counts["重编号"] += renumber_captions(root, full_before, prefix)
    for p in lib.paragraphs(root):
        for old, new in DEBRIS_PAIRS:
            counts["残渣修"] += lib.replace_all(p, old, new)
    return counts


def sweep_apply_path(docx_path, args=None) -> dict:
    """读 docx → 清理类 1-7 → 三护栏 →（护栏绿且 args.apply 时）备份+落盘+复扫，返回报告。

    **为什么是 apply_path 而不是 apply(doc, args)** —— 不是因为要改 document.xml 以外的部件，
    而是因为这道清理的价值全在「护栏红就不许写盘」这条否决权上。apply(doc, args) 的契约
    把存盘交给调用方，护栏只能以返回值的形式提建议：调用方照样可以拿着一份已经被改坏的
    内存树 save 下去，而且 pipeline_lib.run_pipeline 就是这么干的（它把 step 异常吞成
    error 后继续存盘）。把否决权交出去 = 把 fail-closed 降级成 fail-open，那还不如不要护栏。
    所以写盘这一步必须和护栏待在同一个函数里，接口只能是 apply_path。

    args 取 mode / rules / apply 三个属性；apply 缺省为 False = 干跑不落盘（破坏性动作
    默认关，符合「默认非交互直接干」但「默认不写」的取舍：写坏一份终稿的代价远高于多跑一次）。
    """
    docx_path = Path(docx_path)
    mode = getattr(args, "mode", "pei")
    do_apply = bool(getattr(args, "apply", False))
    rules = lib.load_rules(getattr(args, "rules", None))
    names, parts = lib.load_parts(docx_path)
    if "word/document.xml" not in parts:
        raise ValueError(f"不是有效 docx（缺 word/document.xml）: {docx_path}")

    terms = lib.protect_terms(rules)
    root = lib.parse_document(parts)
    before = "".join(lib.ptext(p) for p in lib.paragraphs(root))
    g0 = lib.guards(before, terms)

    counts = sweep(root, rules)
    total = sum(counts.values())

    after = "".join(lib.ptext(p) for p in lib.paragraphs(root))
    g1 = lib.guards(after, terms)
    bad = lib.guard_diff(g0, g1)

    report = {"changed": total, "counts": counts, "护栏": bad, "applied": do_apply,
              "wrote": False, "backup": None, "residual": None, "mode": mode}
    if bad or not do_apply:
        return report          # 护栏红 / 干跑：一个字节都不写

    # --apply 落盘：备份 + 只重写 document.xml（surgical，其余 entry 逐个原样复制）
    if total:
        bak = docx_path.with_suffix(docx_path.suffix + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
        shutil.copy2(str(docx_path), str(bak))
        parts["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(str(docx_path), "w", zipfile.ZIP_DEFLATED) as z:
            for n in names:
                z.writestr(n, parts[n])
        # 部件完整性断言（fail-closed）：基线 = 改前备份；除 document.xml 外任何部件
        # 丢失/字节变化即抛 PartIntegrityError，不静默。verbose=False 保持 --apply
        # stdout 与合并前逐字一致。
        assert_parts_intact(bak, docx_path, verbose=False)
        report["wrote"] = True
        report["backup"] = bak.name

    # 落盘后自动复扫（类别 1-7；第8类归 identity_gate）
    _, parts2 = lib.load_parts(docx_path)
    report["residual"] = lib.scan_parts(parts2, mode=mode, rules=rules, cats=range(1, 8))
    return report


def cmd_sweep(args) -> int:
    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    try:
        report = sweep_apply_path(args.docx, args)
    except (OSError, ValueError, RuntimeError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    counts, bad, total = report["counts"], report["护栏"], report["changed"]
    tag = "[APPLY] " if args.apply else "[CHECK] "
    print(tag + " · ".join(f"{k} {v}" for k, v in counts.items()))
    print("三护栏:", "✅ 全绿（数字集合不变/术语计数不变/保护从句不减）" if not bad
          else "❌ " + "；".join(bad))

    if bad:
        print("护栏红，不落盘")
        print(f"FAIL {len(bad)} findings")
        return 2

    if not args.apply:
        # 干跑：计数 0 = 幂等无事可做 PASS；有计数 = 待清理项
        if total:
            print(f"FAIL {total} findings")
            return 2
        print("PASS")
        return 0

    if report["wrote"]:
        print(f"已写 {args.docx.name} · 备份 {report['backup']}")
    else:
        print("无待清理项，未改动文件（幂等）")

    residual = report["residual"]
    if residual:
        print(f"复扫（类1-7）未归零，剩 {len(residual)} 条（需人工改写或补 rules）:")
        for line in lib.format_findings(residual):
            print("  " + line)
        print(f"FAIL {len(residual)} findings")
        return 2
    print("复扫（类1-7）归零")
    print("PASS")
    return 0


# ════════════════════════════════════════════════════════════════════════════
# identity — 原 bid_identity_gate.py：标书终稿身份泄漏门（只读校验，不改 docx）
#
# 场景：浙江水利院投标工作流，"陪标"（pei 模式）= 通用技术稿严禁出现单位身份；
# 主标（main 模式）允许实名与自有业绩，但仍禁工具痕迹与协作署名。
# 扫描范围 = docx 全部 .xml/.rels part（正文/页眉页脚/docProps 元数据）+ 文件名本身。
# stdout 最后一行必是 "PASS" 或 "FAIL <n> findings"。
# 规则 YAML（--rules，可缺省）只取 identity_banned: [公司全名/人名...] 做项目级增补。
# 参考: shaoxing-eco-flow-2026/scripts/gen_bid_docx.py::assert_no_identity。
# ════════════════════════════════════════════════════════════════════════════

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
    """载入项目级规则 YAML。找不到文件抛 ValueError（不 sys.exit）——

    这个函数被 apply_path 间接调用，而 apply_path 可能跑在别人的 pipeline 里；
    在库函数里 sys.exit 会把整条 pipeline 连坐掐掉，调用方连「哪一步、为什么」都拿不到。
    退出码归 main。
    """
    if path is None:
        return {}
    p = Path(path)
    if not p.is_file():
        raise ValueError("规则文件不存在: %s" % p)
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


def identity_apply_path(docx_path, args=None) -> dict:
    """只读扫身份泄漏，返回 findings/warnings 报告；不改 docx、不写盘、不 sys.exit。

    为什么是 apply_path 而不是 apply(doc, args)：本门的扫描面是「docx 里全部 .xml/.rels
    部件 + 文件名本身」。陪标最常漏的两处恰恰在正文之外 —— docProps/core.xml 的
    creator/lastModifiedBy，和页眉页脚里的院名。python-docx 的 Document 对象够不着它们，
    改用 doc 当入口等于把这道门的覆盖面砍掉一半，而且砍掉的是最容易忘的一半。

    args 取 mode / rules 两个属性（Namespace 或 None 都行）。
    """
    docx = Path(docx_path)
    if not docx.is_file():
        raise ValueError("文件不存在: %s" % docx)
    mode = getattr(args, "mode", "pei")
    rules_path = getattr(args, "rules", None)
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
        raise ValueError("非法 docx（不是 zip）: %s" % docx)
    parts = {n: b.decode("utf-8", "ignore") for n, b in parts_bytes.items()}

    findings, warnings = [], []
    scan_banned(parts, banned, findings)
    if mode == "pei":
        doc_bytes = parts_bytes.get("word/document.xml")
        if doc_bytes is not None:
            scan_bare_yuan(doc_bytes, banned, findings)
    scan_core_props(parts, mode, findings, warnings)
    scan_filename(docx, banned, findings, warnings)

    return {"changed": 0, "findings": findings, "warnings": warnings,
            "findings数": len(findings), "warnings数": len(warnings),
            "mode": mode, "禁词数": len(banned), "增补数": len(extra) if mode == "pei" else 0,
            "part数": len(parts)}


def identity_run(docx, mode, rules_path) -> int:
    """CLI 外壳：调 identity_apply_path 拿报告 → 打印 → 返回退出码（原 sys.exit 改 return）。"""
    args = argparse.Namespace(mode=mode, rules=rules_path)
    try:
        report = identity_apply_path(docx, args)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 1

    print("bid_identity_gate · %s · mode=%s · rules=%s · 内置禁词 %d + 增补 %d"
          % (docx.name, mode, rules_path or "(无)",
             len(BANNED_PEI) if mode == "pei" else len(BANNED_MAIN),
             report["增补数"]))
    print("扫描 part 数: %d（.xml/.rels）+ 文件名" % report["part数"])
    for f in report["findings"]:
        print(f)
    for wline in report["warnings"]:
        print("[warn] " + wline)
    if report["findings"]:
        print("FAIL %d findings" % report["findings数"])
        return 2
    print("PASS")
    return 0


def cmd_identity(args) -> int:
    return identity_run(args.docx, args.mode, args.rules)


# ════════════════════════════════════════════════════════════════════════════
# print — 原 bid_print_ready.py：标书终稿「打印级」只读验收门
#
# 定位：终稿 = 直接可打印。本引擎**只读不写**，扫「打印出来会露馅」的硬伤：
#   fatal ① ASCII 字符画字符（┌┐└┘…═║）出现在正文
#         ② 图/表占位标记残留（〔图位 / 〔界面图位 / 段首【图 / 段首 fenced ``` hint）
#         ③ 题注编号断裂：每类前缀（图 X- / 表 X-，X=章号）节内编号须连续 1..k 且与出现序一致
#            （只认「紧邻图/表」的真题注段，天然排除图目录/表目录/正文引用行）
#         ④ zip 损坏 / word/document.xml 不可解析
#         ⑤ 空段落孤立标点（段文本 strip 后仅剩 ，。、）； —— regex 清理二次残渣兜底）
#         ⑥ OOXML 语义（Word「无法读取的内容」修复弹窗触发点）
#   warn  ① docProps creator/lastModifiedBy 非空 ② media 数为 0 ③ 无 TOC ④ 页面非 A4
# stdout 最后一行必为 "PASS" 或 "FAIL <n> findings"（n = fatal 数；仅 warn → PASS）
# --rules 项目级 YAML 增补：identity_banned 命中在本门降级为 warn 提示（fatal 归身份门）
#         caption_restart_heading: 正则，命中的标题开启新作用域，题注编号在作用域内重启
# 题注校验参考 shaoxing normalize_captions.py（相邻图/表判定）与 finalize2.py（重编号段）。
# ════════════════════════════════════════════════════════════════════════════

DC = "http://purl.org/dc/elements/1.1/"
CP = ("http://schemas.openxmlformats.org/package/2006/metadata/core-properties")

# ── fatal ① ASCII 字符画字符集 ──────────────────────────────────
ASCII_ART_CHARS = set("┌┐└┘├┤┬┴┼╪╞╡╫╬═║╔╗╚╝╠╣╦╩")

# ── fatal ② 占位标记 ────────────────────────────────────────────
PLACEHOLDER_CONTAINS = ["〔图位", "〔界面图位"]          # 段内任意位置命中即 fatal
PLACEHOLDER_PREFIX_RE = [
    (re.compile(r"^【图"), "段首【图…占位题注"),
    (re.compile(r"^```"), "fenced hint 首行 ``` 残留"),
]

# ── fatal ③ 题注（真题注 = 紧邻图/表）────────────────────────────
# 2026-08-01 判据下沉 lib/caption_re.BID_STRICT_CAPTION。**右界断言保留** —— 全仓唯一
# 一份，防把正文内联「见图3-12的说明」误吃成题注，是合并时最容易丢的能力。
# 行为变化：短横补齐五种（旧的是全仓唯一含 U+2011 却**不含全角－**的一份，与 renum
# 恰好互补 → renum 写出的 `图1－2` 它判不出，报假 fatal「题注编号断裂」）。
CAP_SPEC = caption_re.BID_STRICT_CAPTION
CAP_RE = caption_re.pattern(CAP_SPEC)

# ── fatal ⑤ 孤立标点残渣 ───────────────────────────────────────
ISOLATED_PUNCT = set("，。、）；")

# ── warn ④ A4（twips，容差 ±60）────────────────────────────────
A4_PORTRAIT = (11906, 16838)
A4_TOL = 60


def ptext(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def parse_simple_yaml(path: Path):
    """极简 YAML 解析（本管线 schema 专用：key: + '- 字符串' / '- [a, b]' / 行内 [a, b]）。"""

    def unquote(s):
        s = s.strip()
        if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
            s = s[1:-1]
        return s

    def flow_list(s):
        inner = s.strip()[1:-1]
        return [unquote(x) for x in inner.split(",") if x.strip()]

    data, key = {}, None
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.split("#", 1)[0].rstrip() if not raw.lstrip().startswith("#") else ""
        if not line.strip():
            continue
        if not raw[:1].isspace() and ":" in line:
            key, _, rest = line.partition(":")
            key = key.strip()
            rest = rest.strip()
            data[key] = flow_list(rest) if rest.startswith("[") else ([] if not rest else [unquote(rest)])
            if rest and not rest.startswith("["):
                key = None  # 标量已收，防后续误挂
            continue
        s = line.strip()
        if key is not None and s.startswith("- "):
            item = s[2:].strip()
            data[key].append(flow_list(item) if item.startswith("[") else unquote(item))
    return data


def check(docx: Path, mode: str, rules: dict):
    fatal, warn = [], []

    # ── fatal ④ zip / xml 可解析性（挂了直接短路返回）──────────
    try:
        zf = zipfile.ZipFile(str(docx))
    except (zipfile.BadZipFile, OSError) as e:
        fatal.append(f"zip 损坏或不可读: {e}")
        return fatal, warn
    with zf:
        names = zf.namelist()
        if zf.testzip() is not None:
            fatal.append("zip CRC 校验失败（文件损坏）")
            return fatal, warn
        if "word/document.xml" not in names:
            fatal.append("缺 word/document.xml")
            return fatal, warn
        doc_bytes = zf.read("word/document.xml")
        core_bytes = zf.read("docProps/core.xml") if "docProps/core.xml" in names else None
        media_n = sum(1 for n in names if n.startswith("word/media/") and not n.endswith("/"))
    try:
        root = etree.fromstring(doc_bytes)
    except etree.XMLSyntaxError as e:
        fatal.append(f"word/document.xml 不可解析: {e}")
        return fatal, warn

    # ── fatal ⑥ OOXML 语义（Word「无法读取的内容」修复弹窗触发点）──
    with zipfile.ZipFile(str(docx)) as zf2:
        for n in names:
            if n == "word/document.xml" or not n.endswith((".xml", ".rels")):
                continue
            try:
                etree.fromstring(zf2.read(n))
            except etree.XMLSyntaxError as e:
                fatal.append(f"part 不可解析: {n}: {e}")
        rels_bytes = zf2.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in names else b""
    for tc in root.iter(w("tc")):
        if not [c for c in tc if ln(c) in ("p", "tbl", "sdt", "altChunk")]:
            row = ["".join(t.text or "" for t in c.iter(w("t")))[:20]
                   for c in tc.getparent().findall(w("tc"))] if tc.getparent() is not None else []
            fatal.append(f"空表格单元格（CT_Tc 必须含块级元素,Word 必弹修复）· 行内容 {row[:3]}")
    for tr in root.iter(w("tr")):
        if not tr.findall(w("tc")) and tr.find(w("sdt")) is None:
            fatal.append("空表行（tr 无 tc）")
    for tb in root.iter(w("tbl")):
        if not [c for c in tb if ln(c) == "tr"]:
            fatal.append("空表格（tbl 无 tr）")
    _bs = [b.get(w("id")) for b in root.iter(w("bookmarkStart"))]
    _be = [b.get(w("id")) for b in root.iter(w("bookmarkEnd"))]
    _um = set(_bs) ^ set(_be)
    if _um:
        warn.append(f"书签起止不配对 id={sorted(_um)}")
    _dp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
    _ids = [d.get("id") for d in root.iter(_dp)]
    _dup = [i for i in set(_ids) if _ids.count(i) > 1]
    if _dup:
        fatal.append(f"docPr id 重复: {_dup}")
    _used = set(re.findall(rb'r:(?:id|embed)="(rId\d+)"', doc_bytes))
    _defined = set(re.findall(rb'Id="(rId\d+)"', rels_bytes))
    _undef = _used - _defined
    if _undef:
        fatal.append(f"引用了未定义的 rId: {sorted(x.decode() for x in _undef)}")

    body = root.find(w("body"))
    kids = list(body) if body is not None else []
    all_paras = list(root.iter(w("p")))
    pno = {id(p): i for i, p in enumerate(all_paras, 1)}  # 全文段号（P 序，1-based）

    # ── fatal ① ASCII 字符画 ───────────────────────────────────
    for p in all_paras:
        txt = ptext(p)
        hits = sorted({c for c in txt if c in ASCII_ART_CHARS})
        if hits:
            fatal.append(f"ASCII 字符画字符 {''.join(hits)} · P{pno[id(p)]} · 「{txt.strip()[:40]}」")

    # ── fatal ② 占位标记 ───────────────────────────────────────
    for p in all_paras:
        txt = ptext(p)
        st = txt.strip()
        for pat in PLACEHOLDER_CONTAINS:
            if pat in txt:
                fatal.append(f"占位标记「{pat}」残留 · P{pno[id(p)]} · 「{st[:40]}」")
        for rx, desc in PLACEHOLDER_PREFIX_RE:
            if rx.match(st):
                fatal.append(f"{desc} · P{pno[id(p)]} · 「{st[:40]}」")

    # ── fatal ③ 题注编号断裂（只认紧邻图/表的真题注）────────────
    def has_graphic(e):
        return ln(e) == "p" and (
            e.find(f".//{w('drawing')}") is not None or e.find(f".//{w('pict')}") is not None)

    # rules: caption_restart_heading = 正则；命中的段落开启新作用域，题注编号在此重启
    # （投标文件分节交付时常见：每节自成编号体系，跨节允许同号）
    restart_rx = None
    _rr = rules.get("caption_restart_heading") or []
    if _rr:
        try:
            restart_rx = re.compile(_rr[0])
        except re.error:
            restart_rx = None

    groups = {}  # (kind, chapter, scope) -> [(num, P段号, 题注文本)]
    fig_caps = 0
    scope = 0
    for i, e in enumerate(kids):
        if ln(e) != "p":
            continue
        st = ptext(e).strip()
        if restart_rx is not None and st and restart_rx.match(st):
            scope += 1
        m = CAP_RE.match(st)
        if not m:
            continue
        kind, chap, num = m.group("kind"), m.group("sec"), m.group("seq")
        num = int(num)
        prev = kids[i - 1] if i > 0 else None
        nxt = kids[i + 1] if i + 1 < len(kids) else None
        if kind == "图":
            adjacent = (prev is not None and has_graphic(prev)) or (nxt is not None and has_graphic(nxt))
        else:
            adjacent = (prev is not None and ln(prev) == "tbl") or (nxt is not None and ln(nxt) == "tbl")
        if not adjacent:
            continue  # 图目录/表目录/正文引用行，不计
        if kind == "图":
            fig_caps += 1
        groups.setdefault((kind, chap, scope), []).append((num, pno[id(e)], st[:30]))
    for (kind, chap, sc), items in sorted(groups.items()):
        seq = [n for n, _, _ in items]
        want = list(range(1, len(seq) + 1))
        if seq != want:
            where = " ".join(f"{kind}{chap}-{n}@P{pn}" for n, pn, _ in items)
            tag = f"（作用域 {sc}）" if restart_rx is not None else ""
            fatal.append(f"题注编号断裂: 「{kind} {chap}-」{tag}出现序 {seq} ≠ 期望 {want} · {where}")

    # ── fatal ⑤ 孤立标点残渣 ───────────────────────────────────
    for p in all_paras:
        st = ptext(p).strip(" \t　\xa0")
        if st and all(c in ISOLATED_PUNCT for c in st):
            fatal.append(f"空段落孤立标点「{st}」· P{pno[id(p)]}")

    # ── warn ① docProps 身份提示 ──────────────────────────────
    if core_bytes is not None:
        try:
            core = etree.fromstring(core_bytes)
            creator = core.findtext("{%s}creator" % DC) or ""
            lastmod = core.findtext("{%s}lastModifiedBy" % CP) or ""
            if creator.strip():
                warn.append(f"docProps creator 非空: 「{creator.strip()}」（身份门 fatal，此处提示）")
            if lastmod.strip():
                warn.append(f"docProps lastModifiedBy 非空: 「{lastmod.strip()}」（身份门 fatal，此处提示）")
        except etree.XMLSyntaxError:
            warn.append("docProps/core.xml 不可解析（无法核 creator/lastModifiedBy）")

    # ── warn(规则增补) identity_banned 命中提示 ─────────────────
    full = "".join(ptext(p) for p in all_paras)
    for word in rules.get("identity_banned", []):
        c = full.count(word)
        if c:
            warn.append(f"身份禁词（--rules）「{word}」正文命中 x{c}（fatal 归身份门，此处提示）")

    # ── warn ② media 数 ───────────────────────────────────────
    if media_n == 0:
        warn.append("media 数为 0（纯文字标书可能合法；若应含插图请核）")

    # ── warn ③ TOC ─────────────────────────────────────────────
    doc_str = doc_bytes.decode("utf-8", "ignore")
    has_toc = ("TOC" in doc_str and re.search(r"<w:instrText[^>]*>[^<]*TOC", doc_str)) \
        or 'w:val="Table of Contents"' in doc_str \
        or any(ptext(p).strip() in ("目录", "目　录", "目  录") for p in all_paras)
    if not has_toc:
        warn.append("未检出 TOC（无目录域/目录段）")

    # ── warn ④ A4 ─────────────────────────────────────────────
    bad_pages = []
    for i, sect in enumerate(root.iter(w("sectPr")), 1):
        pg = sect.find(w("pgSz"))
        if pg is None:
            bad_pages.append(f"节{i}: 无 pgSz")
            continue
        try:
            pw, ph = int(pg.get(w("w"))), int(pg.get(w("h")))
        except (TypeError, ValueError):
            bad_pages.append(f"节{i}: pgSz 属性缺失")
            continue
        dims = tuple(sorted((pw, ph)))
        want = tuple(sorted(A4_PORTRAIT))
        if not all(abs(a - b) <= A4_TOL for a, b in zip(dims, want)):
            bad_pages.append(f"节{i}: {pw}x{ph} twips 非 A4")
    if bad_pages:
        warn.append("页面非 A4: " + "; ".join(bad_pages))

    # 报告头附加统计（非 finding）
    warn_stats = f"[stat] 段落 {len(all_paras)} · media {media_n} · 真图题注 {fig_caps} · 题注组 {len(groups)}"
    return fatal, warn, warn_stats


def print_apply_path(docx_path, args=None) -> dict:
    """只读跑打印级验收，返回 fatal/warn 报告；不改 docx、不写盘、不 sys.exit。

    为什么是 apply_path 而不是 apply(doc, args)：本门有一半检查项在 document.xml 之外 ——
    zip CRC、各 part 的 XML 可解析性、word/_rels 里的 rId 定义、docProps 元数据、media 计数。
    这些正是「Word 打开弹修复框」的高发处，拿 python-docx 的 Document 当入口一条都看不到。

    args 取 mode / rules 两个属性；rules 是 YAML 路径，None = 不加载项目级增补。
    """
    docx_path = Path(docx_path)
    mode = getattr(args, "mode", "pei")
    rules_path = getattr(args, "rules", None)
    rules = parse_simple_yaml(Path(rules_path)) if rules_path else {}
    result = check(docx_path, mode, rules)
    fatal, warn = result[0], result[1]
    stat = result[2] if len(result) == 3 else None   # zip/xml 挂时短路返回二元组，没有统计行
    return {"changed": 0, "fatal": fatal, "warn": warn, "stat": stat,
            "fatal数": len(fatal), "warn数": len(warn), "mode": mode}


def cmd_print(a) -> int:
    if not a.docx.exists():
        print(f"找不到文件: {a.docx}", file=sys.stderr)
        return 1
    rules = {}
    if a.rules is not None:
        if not a.rules.exists():
            print(f"找不到规则文件: {a.rules}", file=sys.stderr)
            return 1
        try:
            rules = parse_simple_yaml(a.rules)
        except Exception as e:
            print(f"规则 YAML 解析失败: {e}", file=sys.stderr)
            return 1

    print(f"== bid_print_ready · {a.docx.name} · mode={a.mode}"
          f"{' · rules=' + a.rules.name if a.rules else ''} ==")

    result = check(a.docx, a.mode, rules)
    if len(result) == 2:            # 短路（zip/xml 挂）
        fatal, warn = result
        stat = None
    else:
        fatal, warn, stat = result

    for f in fatal:
        print(f"[fatal] {f}")
    for v in warn:
        print(f"[warn]  {v}")
    if stat:
        print(stat)

    if fatal:
        print(f"小结: fatal {len(fatal)} · warn {len(warn)}")
        print(f"FAIL {len(fatal)} findings")
        return 2
    print(f"小结: fatal 0 · warn {len(warn)}")
    print("PASS")
    return 0


# ════════════════════════════════════════════════════════════════════════════
# deref — 原 bid_deref.py：标书正文交叉引用去耦合（陪标：合稿人会删/调章节，
# 写死编号=断链耦合）
#
# 处理形态（2026-07-16 用户钦定：B/D 整句删除、E/F 一并去耦合）：
#   A 括号引用     （7.2）/（详见 7.6.6.3）/（表 7.2-2）      → 整个括号删
#   B 动词+章节号  方案细节详见 7.2 / 承接 3.4               → 删从句；全句皆引用→删整句
#   C 裸章节号     详见剩余的 7.1.3.2 平原河网双指标法        → 删号留名词
#   D 第N章        详见第8章 / 第5章以…统领（叙事主语）       → 删从句；主语位→标 MANUAL
#   E/F 表图引用   见表 7.1-1 / 见图 7.4-1                   → 题注在邻近(±5段)→"见下表/下图"；跨位→题注名
# 删除类操作后自动顺稿（双标点/悬空逗号/空句）。主语位叙事句不自动删，输出 MANUAL 清单人判。
# 护栏：数字集合差异必须全部属于「已知章节号/表图号 token」；SL/T·GB/T 标准号、
# 数量词（亿/万/m³/km²/%）、公文文号〔YYYY〕、日期不受影响（机器核验非承诺）。
# exit 0=完成(或 check 通过)；2=有 MANUAL 待人判或护栏红；1=用法/IO 错。
# ════════════════════════════════════════════════════════════════════════════

NUM = r"\d+(?:\.\d+){1,3}"
NUMRANGE = rf"{NUM}(?:\s*[—\-－~～]\s*{NUM})?"
VERBS = r"(?:详见|参见|承接|对应|支撑|衔接|落地到|喂给|指向|见)"
UNIT_AFTER = r"[亿万倍年月日%‰米mkK]|m³|km²|万元|亿元"

def ptext_items(p):
    return [[t, t.text or ""] for r in p.iter(w("r")) for t in r.findall(w("t"))]

def deref_ptext(p): return "".join(it[1] for it in ptext_items(p))

def replace_span(p, old, new):
    """跨 run 精确替换一次，保格式。返回是否成功。"""
    items = ptext_items(p)
    full = "".join(it[1] for it in items)
    idx = full.find(old)
    if idx < 0: return False
    a, b = idx, idx + len(old); pos = 0; done = False
    for t, txt in items:
        L = len(txt); s, e = pos, pos + L; pos = e
        if e <= a or s >= b: continue
        left = txt[:a - s] if s < a else ""; right = txt[b - s:] if e > b else ""
        if not done:
            t.text = left + new + right; t.set(XML_SPACE, "preserve"); done = True
        else:
            t.text = left + right
    return True


def build_maps(texts, skip=None):
    """章节号→标题名；(图|表, 完整号)→题注名+段号。skip=表格段号集合（单元格会污染映射）。"""
    sec = {}
    caps = {}
    for i, t in enumerate(texts):
        if skip and i in skip:
            continue
        s = t.strip()
        m = re.match(rf"^({NUM})[ 　](.{{2,40}})$", s)
        if m and len(s) < 50 and not re.search(NUM, m.group(2)):
            sec[m.group(1)] = m.group(2).strip()
        m = re.match(rf"^(图|表)\s*(\d+(?:[\.\-－]\d+)*(?:[\-－]\d+)?)[ 　](.+)$", s)
        if m and len(s) < 60:
            caps[(m.group(1), m.group(2))] = (m.group(3).strip(), i)
    return sec, caps


SEC_SHAPE = re.compile(r"^(?:[3-9]|10)(?:\.\d+){1,3}$")


def smooth(t):
    """删除类操作后的残渣顺稿。"""
    prev = None
    while prev != t:
        prev = t
        t = re.sub(r"（\s*）|\(\s*\)", "", t)
        t = re.sub(r"，\s*，", "，", t)
        t = re.sub(r"；\s*；", "；", t)
        t = re.sub(r"，\s*。", "。", t)
        t = re.sub(r"；\s*。", "。", t)
        t = re.sub(r"：\s*。", "。", t)
        t = re.sub(r"^\s*[，；]", "", t)
        t = re.sub(r"([。；：])\s*，", r"\1", t)
    return t


SENT_SPLIT = re.compile(r"(?<=[。！？])")
CLAUSE_SPLIT = re.compile(r"(?<=[，；：])")

def cell_deref(text, sec, manual, loc):
    """表格单元格：章节形编号→标题名（映射列整列是号，删空会成残渣）。数量词不动。"""
    def tok_repl(m):
        tok = m.group(0)
        if not SEC_SHAPE.match(tok):
            return tok  # 0.01 / 99.9 / 2.5 等数量词
        if tok in sec:
            return sec[tok]
        manual.append((loc, f"表格单元格未知章节号 {tok}: {text[:60]}"))
        return tok
    t = re.sub(NUM, tok_repl, text)
    t = re.sub(r"([一-鿿]）?)\s*[／/]\s*([一-鿿])", r"\1／\2", t)
    return smooth(t)


def deref_text(text, sec, caps, para_idx, manual, loc):
    """对一段纯文本做去耦合，返回新文本（不含顺稿前后对比逻辑）。"""
    t = text

    # --- E/F 表图引用 ---
    def fig_repl(m):
        kind, num = m.group(2), m.group(3)
        cap = caps.get((kind, num)) or caps.get((kind, num.replace(".", "-")))
        lead = m.group(1) or ""
        if cap:
            title, ci = cap
            if para_idx is not None and abs(ci - para_idx) <= 5:
                pos = "下" if ci >= (para_idx or 0) else "上"
                return (lead or "见") + pos + ("表" if kind == "表" else "图")
            # 跨位置：题注名若已紧跟在引用后（“图 X-Y“名”…”）→ 只删号；否则替换成题注名
            after = t[m.end():m.end() + len(title) + 4]
            if title[:6] in after:
                return ""
            return (lead or "") + ("《" + title + "》" if False else title)
        manual.append((loc, f"{kind} {num} 题注未定位: …{t[max(0,m.start()-15):m.end()+15]}…"))
        return m.group(0)
    t = re.sub(rf"(见|如|详见)?({'表'}|{'图'})\s*(\d+(?:[\.\-－]\d+)+)", fig_repl, t)

    # --- A 括号引用（章节号；含编号列表与带前缀描述词的引用括号） ---
    NUMLIST = rf"{NUMRANGE}(?:\s*[、，/／]\s*{NUMRANGE})*"
    t = re.sub(rf"（\s*(?:{VERBS})?\s*{NUMLIST}\s*(?:节|章)?\s*）", "", t)
    t = re.sub(rf"（[^（）]{{0,20}}?(?:详见|参见|见|衔接|承接)\s*{NUMLIST}[^（）]{{0,12}}）", "", t)
    t = re.sub(rf"（\s*(?:{VERBS})?\s*第\s*\d+\s*章\s*）", "", t)

    # --- B 动词+号 从句/整句删除；“第N章”一律 MANUAL（叙事句机器改必坏语感） ---
    ref_clause = re.compile(rf"{VERBS}\s*{NUMRANGE}|{VERBS}\s*第\s*\d+\s*章")
    sents = SENT_SPLIT.split(t)
    out_sents = []
    for sent in sents:
        if not sent.strip():
            continue
        has_verb_ref = bool(ref_clause.search(sent))
        has_chapter = bool(re.search(r"第\s*\d+\s*章", sent))
        if has_chapter and not has_verb_ref:
            manual.append((loc, f"第N章叙事句需人判: {sent.strip()[:90]}"))
            out_sents.append(sent); continue
        if not has_verb_ref:
            out_sents.append(sent); continue
        clauses = CLAUSE_SPLIT.split(sent)
        kept = [c for c in clauses if not ref_clause.search(c)]
        if not any(re.search(r"[一-鿿]{4,}", c) for c in kept):
            continue  # 全句皆引用 → 整句删
        new_sent = "".join(kept)
        if re.search(r"第\s*\d+\s*章", new_sent):
            manual.append((loc, f"删从句后仍含第N章需人判: {new_sent.strip()[:90]}"))
        out_sents.append(new_sent)
    t = "".join(out_sents)

    # --- C 裸章节号：已知章节号 token（含区间），后不接数量单位 ---
    def bare_repl(m):
        tok = m.group(1)
        head = tok.split("—")[0].split("-")[0].strip()
        if head in sec or re.match(r"^[3-9]\.\d", head) or re.match(r"^10\.", head):
            return ""
        return m.group(0)
    t = re.sub(rf"(?<![\d.．表图SLGB/T\-])({NUMRANGE})(?:\s*(?:节|各节|小节))?\s*(?!{UNIT_AFTER})(?=[一-鿿（(，。；、]|$)", bare_repl, t)

    return smooth(t)


def frag_scan(text, loc):
    """删句/删号后断句残渣检测：括号失衡、（，、空括号、已在…详述类断句。"""
    probs = []
    d = 0; broken = False
    for ch in text:
        if ch == "（": d += 1
        elif ch == "）":
            d -= 1
            if d < 0: broken = True; break
    if broken or d != 0:
        probs.append((loc, "括号失衡: " + text[:60]))
    for pat, msg in (("（，", "（，残渣"), ("，）", "，）残渣"), ("（）", "空括号")):
        if pat in text:
            probs.append((loc, msg + ": " + text[:60]))
    if re.search(r"已在[ 　]*[、，]?[ 　]*详述", text):
        probs.append((loc, "断句(已在…详述): " + text[:60]))
    return probs


def frag_scan_pair(old, new, loc):
    """成对断句检测：旧文以句号收、新文以逗号类悬尾 = 删句吃掉句号（自愈后不应再现）。"""
    probs = frag_scan(new, loc)
    ns, os_ = new.rstrip(), old.rstrip()
    if os_ and ns and os_[-1] in "。！？" and ns[-1] in "，、；：":
        probs.append((loc, f"删句吃句号(段尾悬「{ns[-1]}」): …{ns[-30:]}"))
    return probs


def ooxml_scan(root):
    """OOXML 语义扫：空 tc/tr/tbl（Word「无法读取的内容」修复弹窗触发点）。"""
    probs = []
    for tc in root.iter(w("tc")):
        if not [c for c in tc if c.tag in (w("p"), w("tbl"))]:
            probs.append(("tc", "空单元格（Word 修复弹窗触发,CT_Tc 必须含块级元素）"))
    for tr in root.iter(w("tr")):
        if not tr.findall(w("tc")):
            probs.append(("tr", "空表行"))
    for tbl in root.iter(w("tbl")):
        if not [c for c in tbl if c.tag == w("tr")]:
            probs.append(("tbl", "空表格"))
    return probs


def run_docx(docx: Path, check: bool, manual_pairs_path=None):
    with zipfile.ZipFile(str(docx)) as z:
        names = z.namelist(); parts = {n: z.read(n) for n in names}
    root = etree.fromstring(parts["word/document.xml"])
    paras = list(root.iter(w("p")))
    texts = [deref_ptext(p) for p in paras]
    in_tbl = set()
    for i, p in enumerate(paras):
        if any(True for _ in p.iterancestors(w("tbl"))):
            in_tbl.add(i)
    sec, caps = build_maps(texts, skip=in_tbl)
    print(f"章节映射 {len(sec)} 条 · 题注映射 {len(caps)} 条")

    manual_pairs = []
    if manual_pairs_path:
        import yaml
        manual_pairs = [(o, n) for o, n in yaml.safe_load(open(manual_pairs_path, encoding="utf-8")) or []]

    manual, changes = [], []
    protected = re.compile(r"SL/T\s*\d|GB/T\s*\d|〔\d{4}〕")
    for i, p in enumerate(paras):
        old = texts[i]
        if not old.strip(): continue
        cur = old
        for o, n in manual_pairs:
            if o in cur:
                cur = cur.replace(o, n)
        s = cur.strip()
        if re.match(rf"^(?:\d+(?:\.\d+){{0,3}}|(?:图|表)\s*[\d\.\-－]+)[ 　]", s) and len(s) < 60 and i not in in_tbl:
            if cur != old:
                changes.append((i, old, cur))
            continue  # 标题/题注本体不动
        if i in in_tbl and re.search(NUM, s) and len(s) < 80 and not re.search(r"[。！？]", s):
            new = cell_deref(cur, sec, manual, f"P{i:04d}")
        else:
            new = deref_text(cur, sec, caps, i, manual, f"P{i:04d}")
        if new != old:
            # 尾标点自愈：删句吃掉了句号（旧文以句号收、新文以逗号/顿号/分号悬尾）→ 补回句号
            ns, os_ = new.rstrip(), old.rstrip()
            if os_ and ns and os_[-1] in "。！？" and ns[-1] in "，、；：":
                new = ns[:-1] + "。"
            changes.append((i, old, new))

    # 护栏：数字差异必须全部是章节/表图号 token（或其连字符拆分组件）
    def digit_multiset(t):
        t2 = protected.sub("", t)
        return re.findall(r"\d+(?:\.\d+)*", t2)
    allowed = set(sec)
    for (_, num) in caps:
        allowed.add(num)
        allowed.update(re.split(r"[\-－.．]", num))
    removed_ok, removed_bad = 0, []
    for i, old, new in changes:
        from collections import Counter
        diff = Counter(digit_multiset(old)) - Counter(digit_multiset(new))
        for tok, cnt in diff.items():
            base = tok.split(".")[0]
            if tok in allowed or (base.isdigit() and 3 <= int(base) <= 10):
                removed_ok += cnt
            else:
                removed_bad.append((f"P{i:04d}", tok))
        gained = Counter(digit_multiset(new)) - Counter(digit_multiset(old))
        for tok, cnt in gained.items():
            if tok in allowed or "．" in tok:
                continue  # 单元格号→标题名带出的下级号（标题文本自身含号已被 build_maps 排除，此处放行映射产物）
            removed_bad.append((f"P{i:04d}", f"新增数字 {tok}x{cnt}"))

    # 断句残渣预扫：每条改动的新文本过 frag_scan_pair（check 模式即可见,apply 模式硬拦）
    frags = []
    for i, old, new in changes:
        frags.extend(frag_scan_pair(old, new, f"P{i:04d}"))

    print(f"改动段 {len(changes)} · 删除编号 token {removed_ok} · MANUAL {len(manual)}")
    for i, old, new in changes:
        print(f"--- P{i:04d}")
        print(f"  旧: {old[:150]}")
        print(f"  新: {new[:150] if new.strip() else '(整段删空)'}")
    if manual:
        print("== MANUAL 需人判 ==")
        for loc, msg in manual: print(f"  {loc}: {msg}")
    if removed_bad:
        print("== 护栏红（非章节号数字变动，拒绝落盘）==")
        for loc, tok in removed_bad: print(f"  {loc}: {tok}")
        sys.exit(2)
    if frags:
        print("== 断句残渣红（删句留尾巴，拒绝落盘；改 manual-pairs 后重跑）==")
        for loc, msg in frags: print(f"  {loc}: {msg}")
        if not check: sys.exit(2)
    if check:
        print("[CHECK] 未落盘"); sys.exit(2 if (manual or frags) else 0)

    for i, old, new in changes:
        if new.strip():
            assert replace_span(paras[i], old, new), f"P{i} 替换失败"
        else:
            par = paras[i].getparent()
            siblings = [c for c in par if c.tag in (w("p"), w("tbl"))]
            if par.tag == w("tc") and len(siblings) <= 1:
                # tc 内唯一块级元素禁删光（空 tc = Word 修复弹窗）,清 run 留空段
                for r_ in list(paras[i].findall(w("r"))):
                    paras[i].remove(r_)
            else:
                par.remove(paras[i])

    # apply 后 OOXML 语义复扫,红则不写盘
    oox = ooxml_scan(root)
    if oox:
        print("== OOXML 语义红（拒绝落盘）==")
        for loc, msg in oox: print(f"  {loc}: {msg}")
        sys.exit(2)
    bak = str(docx) + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copy2(str(docx), bak)
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(str(docx), "w", zipfile.ZIP_DEFLATED) as z:
        for n in names: z.writestr(n, parts[n])
    # 部件完整性断言（fail-closed）：基线 = 改前备份；除 document.xml 外任何部件
    # 丢失/字节变化即抛 PartIntegrityError，不静默。verbose=False 保持 apply
    # stdout 与合并前逐字一致。
    assert_parts_intact(Path(bak), docx, verbose=False)
    print(f"已写 {docx.name} 备份 {os.path.basename(bak)}")
    sys.exit(2 if manual else 0)


def run_md(dirpath: str, check: bool):
    files = sorted(glob.glob(os.path.join(dirpath, "*.md")))
    all_lines = []
    for f in files:
        for ln_ in open(f, encoding="utf-8"):
            all_lines.append(ln_.rstrip("\n"))
    sec, caps = build_maps(all_lines)
    # md 标题形如 "## 7.2 xxx" / "#### 7.1.3.2 xxx"，补进映射
    for f in files:
        for ln_ in open(f, encoding="utf-8"):
            m = re.match(rf"^#+\s+({NUM})[ 　](.+)$", ln_.strip())
            if m: sec[m.group(1)] = m.group(2).strip()
    print(f"[md] 章节映射 {len(sec)} 条 · 题注映射 {len(caps)} 条")
    manual, total = [], 0
    for f in files:
        lines = open(f, encoding="utf-8").read().split("\n")
        out = []
        changed = 0
        for j, ln_ in enumerate(lines):
            if re.match(rf"^#+\s", ln_) or re.match(rf"^(图|表)\s*\d", ln_.strip()):
                out.append(ln_); continue
            new = deref_text(ln_, sec, caps, None, manual, f"{os.path.basename(f)}:{j+1}")
            if new != ln_:
                changed += 1
                print(f"--- {os.path.basename(f)}:{j+1}")
                print(f"  旧: {ln_[:140]}")
                print(f"  新: {new[:140] if new.strip() else '(删空)'}")
            out.append(new)
        if changed and not check:
            open(f, "w", encoding="utf-8").write("\n".join(out))
        total += changed
    print(f"[md] 改动行 {total} · MANUAL {len(manual)}")
    if manual:
        print("== MANUAL 需人判 ==")
        for loc, msg in manual: print(f"  {loc}: {msg}")
    sys.exit(2 if manual else 0)


def cmd_deref(args) -> int:
    if args.md:
        run_md(args.target, args.check)
    else:
        run_docx(Path(args.target), args.check, args.manual_pairs)
    return 0  # run_md/run_docx 自带 sys.exit 收口，正常不会走到这


# ════════════════════════════════════════════════════════════════════════════
# run — 原 bid_final.py：标书终稿机械管线 driver（残留扫描 → 清理 → 身份门 → 付印门）
#
# 流程:
#   ① residue scan（8 类残留只读扫描）
#   ② --apply 时: finalize sweep --apply（确定性清理类1-7，备份+三护栏）→ residue 复扫须归零
#   ③ identity gate（第8类身份泄漏门）
#   ④ print ready（付印门）
#   全绿 → 「终稿 PASS」+ sha256 前 12 位指纹；任一红 → exit 2 并汇总哪道门红。
# exit 0 = PASS；exit 2 = 有门红；exit 1 = 用法/IO 错。
#
# 2026-07-31 合并后由 subprocess 改为直接函数调用：banner 保留旧脚本名字面量
# （逐字节 stdout 契约，work_ops.cmd_bidgate 正则消费），redirect_stdout/redirect_stderr
# + rstrip 复刻旧 capture_output 语义。
# ════════════════════════════════════════════════════════════════════════════

def run_engine(script_name, fn, ns, mode, rules, extra=()):
    """直调一道门，banner 与输出裁剪逐字节复刻旧 subprocess driver。"""
    tail = ["--mode", mode]
    if rules:
        tail += ["--rules", str(rules)]
    tail += list(extra)
    print(f"── {script_name} {' '.join(tail)} " + "─" * max(4, 60 - len(script_name)))
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = fn(ns)
    if out.getvalue():
        print(out.getvalue().rstrip())
    if err.getvalue():
        print(err.getvalue().rstrip(), file=sys.stderr)
    return rc


def cmd_run(args) -> int:
    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    if args.rules and not args.rules.is_file():
        print(f"错误: 规则文件不存在 {args.rules}", file=sys.stderr)
        return 1

    gates = {}  # 门名 → exit code

    def _scan_ns():
        return argparse.Namespace(docx=args.docx, mode=args.mode, rules=args.rules)

    # ① 残留扫描
    rc = run_engine("bid_residue_scan.py", cmd_scan, _scan_ns(), args.mode, args.rules)
    if rc == 1:
        print("residue_scan 用法/IO 错误，中止", file=sys.stderr)
        return 1

    # ② --apply: 确定性清理 + 复扫归零
    if args.apply:
        sweep_ns = argparse.Namespace(docx=args.docx, mode=args.mode, rules=args.rules,
                                      check=False, apply=True)
        rc_sweep = run_engine("bid_finalize_sweep.py", cmd_sweep, sweep_ns,
                              args.mode, args.rules, extra=["--apply"])
        if rc_sweep == 1:
            print("finalize_sweep 用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates["finalize_sweep"] = rc_sweep
        rc = run_engine("bid_residue_scan.py", cmd_scan, _scan_ns(), args.mode, args.rules)  # 复扫须归零
        if rc == 1:
            print("residue_scan 复扫用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates["residue_rescan"] = rc
    else:
        gates["residue_scan"] = rc

    # ③④ 身份门 + 付印门
    identity_ns = argparse.Namespace(docx=args.docx, mode=args.mode,
                                     rules=(str(args.rules) if args.rules else None))
    print_ns = argparse.Namespace(docx=args.docx, mode=args.mode, rules=args.rules)
    for script_name, name, fn, gns in (
        ("bid_identity_gate.py", "identity_gate", cmd_identity, identity_ns),
        ("bid_print_ready.py", "print_ready", cmd_print, print_ns),
    ):
        rc = run_engine(script_name, fn, gns, args.mode, args.rules)
        if rc == 1:
            print(f"{name} 用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates[name] = rc

    # ⑤ 目录门：正文章号 ≡ chapters.yaml ≡ 招标那张规定顺序的表（2026-08-30 /govern）
    #    本 driver 是标书专用编排 —— 定位不到 _project.yaml 就是红，不静默跳过。
    print("── bid_toc_gate.py " + "─" * 44)
    gates["toc_gate"] = bid_toc_gate.main([str(args.docx)])

    # ── 汇总 ──
    red = [g for g, rc in gates.items() if rc != 0]
    print("═" * 68)
    print("门控汇总:", " · ".join(f"{g}={'✅' if rc == 0 else '❌'}" for g, rc in gates.items()))
    if red:
        print(f"红门: {', '.join(red)}")
        print(f"FAIL {len(red)} findings")
        return 2
    sha = hashlib.sha256(args.docx.read_bytes()).hexdigest()[:12]
    print(f"终稿 PASS · {args.docx.name} · sha256 {sha}")
    print("PASS")
    return 0


# ════════════════════════════════════════════════════════════════════════════
# argparse 表面（每个子命令与原脚本零改动）
# ════════════════════════════════════════════════════════════════════════════

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="bid_gate.py",
        description="标书终稿门检族（run=五门 driver · scan/sweep/identity/print/toc=单门 · deref=交叉引用去耦合）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("run", help="终稿机械管线 driver（原 bid_final.py）",
                       description="标书终稿机械管线 driver")
    p.add_argument("docx", type=Path)
    p.add_argument("--mode", choices=["main", "pei"], default="pei")
    p.add_argument("--rules", type=Path, default=None)
    p.add_argument("--apply", action="store_true", help="执行确定性清理（否则只跑各道门）")
    p.set_defaults(fn=cmd_run)

    p = sub.add_parser("scan", help="8 类残留只读扫描（原 bid_residue_scan.py）",
                       description="标书终稿 8 类残留只读扫描器")
    p.add_argument("docx", type=Path)
    p.add_argument("--mode", choices=["main", "pei"], default="pei")
    p.add_argument("--rules", type=Path, default=None)
    p.set_defaults(fn=cmd_scan)

    p = sub.add_parser("sweep", help="类 1-7 确定性清理（原 bid_finalize_sweep.py）",
                       description="标书终稿确定性清理引擎（残留类 1-7）")
    p.add_argument("docx", type=Path)
    p.add_argument("--mode", choices=["main", "pei"], default="pei")
    p.add_argument("--rules", type=Path, default=None)
    g = p.add_mutually_exclusive_group()
    g.add_argument("--check", action="store_true", help="干跑打印计数（默认）")
    g.add_argument("--apply", action="store_true", help="落盘（备份+三护栏机检）")
    p.set_defaults(fn=cmd_sweep)

    p = sub.add_parser("toc", help="目录一致性门：章号 ≡ chapters.yaml ≡ 招标表",
                       description="正文章号/章名/分值 ≡ chapters.yaml ≡ 招标文件那张规定顺序的表")
    p.add_argument("target", type=Path, help="项目目录 或 技术标 docx")
    p.add_argument("--chapters", type=Path, default=None)
    p.set_defaults(fn=lambda a: bid_toc_gate.main(
        [str(a.target)] + (["--chapters", str(a.chapters)] if a.chapters else [])))

    p = sub.add_parser("identity", help="第 8 类身份泄漏只读门（原 bid_identity_gate.py）",
                       description="标书终稿身份泄漏门（只读）")
    p.add_argument("docx", type=Path)
    p.add_argument("--mode", choices=["main", "pei"], default="pei",
                   help="pei=陪标(默认,全禁词+裸院+元数据fatal) / main=主标(仅工具痕迹+协作署名)")
    p.add_argument("--rules", default=None, help="项目级规则 YAML（identity_banned 增补）")
    p.set_defaults(fn=cmd_identity)

    p = sub.add_parser("print", help="付印只读门（原 bid_print_ready.py）",
                       description="标书终稿打印级只读验收门")
    p.add_argument("docx", type=Path)
    p.add_argument("--mode", choices=["main", "pei"], default="pei")
    p.add_argument("--rules", type=Path, default=None)
    p.set_defaults(fn=cmd_print)

    p = sub.add_parser("deref", help="交叉引用去耦合（原 bid_deref.py；docx surgical 或 --md 源 md）",
                       description="标书正文交叉引用去耦合")
    p.add_argument("target", type=str)
    p.add_argument("--check", action="store_true")
    p.add_argument("--md", action="store_true")
    p.add_argument("--manual-pairs", type=str, default=None)
    p.set_defaults(fn=cmd_deref)

    args = ap.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
