#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_residue_lib.py — 标书终稿残留检测共享库（bid_residue_scan / bid_finalize_sweep 共用）。

8 类残留 taxonomy（内置通用检测，项目差异走 --rules YAML 增补）：
  1 协作标记   【待..】/占位版/（7.6.6，数字化团队）/祝稿
  2 拟hedge    （拟）/，拟）/标"拟"/拟按 —— 护术语 比拟/模拟/拟合/拟定/拟派
  3 评分脚手架 得满分要求/进行综合评定裸句/评分表·评分子项·评分点/评委/技术分
  4 内部编号   〔E-xx〕/worklib#/招标段号 段4xx —— 公文文号〔2025〕保留
  5 二次残渣   ，此处）/，本处）/空括号/悬空逗号
  6 断裂引用   题注编号与正文出现序不符 / 编号不连续
  7 口径meta   原件待核/二手转述/待核定占位
  8 身份泄漏   公司名/院自指/业绩归属/裸「院」/docProps creator·lastModifiedBy 非空/python-docx 痕迹
               （--mode main 跳过实名类：公司名/院自指/业绩归属/裸院/元数据署名；仍报工具痕迹）

本库只读检测 + 通用 XML 工具（跨 run 保格式替换供 sweep 用），不含任何落盘逻辑。
依赖：stdlib + lxml；规则 YAML 解析用 PyYAML（缺失时抛清晰错误）。
风格对齐 doctools 家族（docx_qa.py）与 shaoxing finalize2.py / gen_bid_docx.py 交付门。
"""
import re
import zipfile
from bisect import bisect_right
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(t): return f"{{{W}}}{t}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# ── 内置护栏词表（SSOT：与 shaoxing finalize2 / gen_bid_docx 交付门对齐）──────
PROTECT_TERMS = ["比拟", "模拟", "拟合", "拟定", "拟派"]
PROTECT_CLAUSES = ["以采购人核定", "以合同约定为准", "按现行有效版本核定"]
YUAN_PROTECT = ["院士", "医院", "法院", "科学院", "研究院", "学院", "剧院", "国务院"]
SELF_REF = ["我院", "本院", "院内", "院级", "院总工"]
PERF_ATTRIB = ["移植投标人", "借鉴投标人", "投标人已在", "投标人承担浙江省内", "投标人全省"]
TOOL_TRACES = ["python-docx"]

# ── 内置检测 pattern ─────────────────────────────────────────────
COLLAB_RE = re.compile(r"【待[^】]*】")
COLLAB_TOKENS = ["占位版", "数字化团队", "祝稿", "待祝稿合并"]
HEDGE_PATS = [re.compile(p) for p in
              (r"（拟[）：，。；]", r"（拟$", r"，拟[）；，]", r"标[“\"]拟", r"拟按", r"（拟）")]
SCORE_STARTSWITH = ["得满分要求"]
SCORE_TOKENS = ["进行综合评定"]
SCORE_RE = re.compile(r"评分表|评分子项|评分点|评委|技术分(?![析工层级包布配])")
E_RE = re.compile(r"〔E-[^〕]*〕")
INTERNAL_TOKENS = ["worklib#"]
SEG_RE = re.compile(r"(?<![0-9])(?:招标)?段4\d\d(?![0-9])")  # 招标 python-docx 段号=内部码；公文文号〔2025〕不在此列
DEBRIS_TOKENS = ["，此处）", "，本处）", "（）", "，）", "、）", "（，", "（、", "不复写"]


def paren_unbalanced(text):
    """段内全角括号失衡（删句留尾巴的典型形态,如「本章不复写）：」）。"""
    d = 0
    for ch in text:
        if ch == "（":
            d += 1
        elif ch == "）":
            d -= 1
            if d < 0:
                return True
    return d != 0
META_TOKENS = ["原件待核", "二手转述", "待核定占位", "可假定占位", "待台账", "待核实", "TODO", "TBD"]
CAPTION_RE = re.compile(r"([表图])\s?(\d+(?:\.\d+)?)-(\d+)")
NUM_RE = re.compile(r"[0-9０-９]+(?:[.．][0-9０-９]+)?%?")

CAT_NAMES = {
    1: "协作标记", 2: "拟hedge", 3: "评分脚手架", 4: "内部编号",
    5: "二次残渣", 6: "断裂引用", 7: "口径meta", 8: "身份泄漏",
    9: "交叉引用耦合",
}
CAT_ADVICE = {
    1: "整段删标记+去协作署名，正文能自立才算删完",
    2: "谨慎改写去 hedge（护术语/保护从句），非硬删",
    3: "删脚手架块/照抄裸句；评分元语言改「招标要求/两方面」；去直呼评委",
    4: "剥内部码（〔E-xx〕/worklib#/招标段号）；公文文号〔2025〕保留",
    5: "修截断残渣（上一轮 regex 替换自产物，每轮替换后必复扫）",
    6: "题注按正文出现序重编号/修指向（占位法防连环替换）",
    7: "改成正常引用口径或删；带数据的改写守数据红线",
    8: "交 bid_identity_gate 处理：去公司名/院自指/业绩归属，清 docProps 元数据",
    9: "交 bid_deref 去耦合（陪标合稿人会删/调章节，正文写死编号=断链）：括号引用删/动词句删/号→标题名；题注本体编号保留",
}

# 类9 交叉引用耦合（pei 专属；2026-07-16 用户钦定「标书里不要交叉引用，合稿人会调整」）
XREF_PATS = [
    re.compile(r"（\s*(?:详见|参见|见)?\s*\d+\.\d[\d.]*(?:\s*[、，/／]\s*\d+\.\d[\d.]*)*\s*(?:节|章)?\s*）"),
    re.compile(r"(?:详见|参见|承接|支撑|衔接|对应|落地到)\s*\d+\.\d"),
    re.compile(r"第\s*[3-9]\s*章"),
    re.compile(r"(?:见|如|详见)\s*(?:表|图)\s*\d+[\.\-－]\d"),
]
XREF_SKIP = re.compile(r"^(?:\d+(?:\.\d+){0,3}|(?:图|表)\s*[\d\.\-－]+)[ 　]")


# ── docx 读取 / 段落工具 ─────────────────────────────────────────
def load_parts(docx_path):
    """读 docx 全部 entry → (names, {name: bytes})。IOError/BadZipFile 由调用方兜。"""
    with zipfile.ZipFile(str(docx_path)) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}
    return names, parts


def parse_document(parts):
    return etree.fromstring(parts["word/document.xml"])


def ptext(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def paragraphs(root):
    return list(root.iter(w("p")))


# ── 跨 run 保格式替换（沿用 finalize2/docx_qa 同款）────────────────
def replace_once(p, old, new):
    items = [[t, t.text or ""] for r in p.iter(w("r")) for t in r.findall(w("t"))]
    full = "".join(it[1] for it in items)
    idx = full.find(old)
    if idx < 0:
        return False
    a, b = idx, idx + len(old)
    pos = 0
    done = False
    for t, txt in items:
        L = len(txt)
        s, e = pos, pos + L
        pos = e
        if e <= a or s >= b:
            continue
        left = txt[:a - s] if s < a else ""
        right = txt[b - s:] if e > b else ""
        if not done:
            t.text = left + new + right
            t.set(XML_SPACE, "preserve")
            done = True
        else:
            t.text = left + right
    return True


def replace_all(p, old, new):
    n = 0
    while replace_once(p, old, new):
        n += 1
        if n > 300:
            break
    return n


def regex_strip_para(p, pattern, repl=""):
    """段内正则命中逐个用 replace_once 精确删（保格式），返回命中数。"""
    full = ptext(p)
    n = 0
    for h in pattern.findall(full) if pattern.groups == 0 else [m.group(0) for m in pattern.finditer(full)]:
        if replace_once(p, h, repl):
            n += 1
    return n


# ── 规则 YAML ────────────────────────────────────────────────────
RULE_KEYS = {"delete_startswith": list, "delete_exact": list, "exact": list,
             "caption_renumber": list, "protect_terms": list, "identity_banned": list}


def load_rules(path=None):
    """载入项目级规则 YAML；path=None 时返回空规则（只用内置通用规则）。"""
    rules = {k: [] for k in RULE_KEYS}
    if not path:
        return rules
    try:
        import yaml
    except ImportError:
        raise RuntimeError("解析 --rules 需要 PyYAML（pip install pyyaml），当前环境缺失")
    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    for k in RULE_KEYS:
        v = data.get(k) or []
        if not isinstance(v, list):
            raise ValueError(f"规则 {k} 必须是列表，得到 {type(v).__name__}")
        rules[k] = v
    for pair in rules["exact"]:
        if not (isinstance(pair, (list, tuple)) and len(pair) == 2):
            raise ValueError(f"exact 每项必须 [old, new] 二元组，得到 {pair!r}")
    return rules


def protect_terms(rules):
    return list(dict.fromkeys(PROTECT_TERMS + list(rules.get("protect_terms") or [])))


# ── 三护栏 ───────────────────────────────────────────────────────
def numset(txt):
    """数字集合：先剥〔...〕与内部段号码，再取数字去重集合（与 finalize2 对齐）。"""
    txt = re.sub(r"〔[^〕]*〕", "", txt)
    txt = re.sub(r"招标段4\d\d|段4\d\d", "", txt)
    return sorted(set(NUM_RE.findall(txt)))


def guards(full, terms):
    return {
        "numset": numset(full),
        "terms": {k: full.count(k) for k in terms},
        "clauses": {c: full.count(c) for c in PROTECT_CLAUSES},
    }


def guard_diff(g0, g1):
    """返回违规描述列表；空列表 = 三护栏全绿。"""
    bad = []
    if g0["numset"] != g1["numset"]:
        bad.append(f"数字集合变化: 差={set(g0['numset']) ^ set(g1['numset'])}")
    for k, v in g0["terms"].items():
        if g1["terms"].get(k) != v:
            bad.append(f"受保护术语「{k}」计数 {v}→{g1['terms'].get(k)}")
    for c, v in g0["clauses"].items():
        if g1["clauses"].get(c, 0) < v:
            bad.append(f"保护从句「{c}」计数减少 {v}→{g1['clauses'].get(c, 0)}")
    return bad


# ── 检测主入口 ───────────────────────────────────────────────────
def _strip_terms(s, terms):
    for t in terms:
        s = s.replace(t, "")
    return s


def _mk(cat, para, text, marks):
    excerpt = re.sub(r"\s+", " ", text.strip())[:80]
    return {"cat": cat, "para": para, "excerpt": excerpt, "marks": marks,
            "advice": CAT_ADVICE[cat]}


def scan_parts(parts, mode="pei", rules=None, cats=None):
    """扫描 docx parts → findings 列表（dict: cat/para/excerpt/marks/advice）。

    mode: pei（默认，全量）| main（跳过第8类实名类：公司名/院自指/业绩归属/裸院/元数据署名，
          仍报工具痕迹）。cats: 只扫指定类别集合（None=全部 1-8）。
    """
    rules = rules or load_rules(None)
    cats = set(cats) if cats else set(range(1, 10))
    terms = protect_terms(rules)
    root = parse_document(parts)
    paras = paragraphs(root)
    findings = []

    ptexts = [ptext(p) for p in paras]
    for i, s in enumerate(ptexts, 1):
        st = s.strip()
        if not st:
            continue
        # 1 协作标记
        if 1 in cats:
            marks = COLLAB_RE.findall(st) + [t for t in COLLAB_TOKENS if t in st]
            for t in rules["delete_startswith"]:
                if st.startswith(t) and t not in SCORE_STARTSWITH:
                    marks.append(f"delete_startswith:{t}")
            if marks:
                findings.append(_mk(1, i, st, marks))
        # 2 拟hedge（先剥受保护术语再匹配）
        if 2 in cats and "拟" in st:
            bare = _strip_terms(st, terms)
            marks = [m for pat in HEDGE_PATS for m in pat.findall(bare)]
            if marks:
                findings.append(_mk(2, i, st, marks))
        # 3 评分脚手架
        if 3 in cats:
            marks = [f"段首:{t}" for t in SCORE_STARTSWITH if st.startswith(t)]
            marks += [t for t in SCORE_TOKENS if t in st]
            marks += SCORE_RE.findall(st)
            if st in rules["delete_exact"]:
                marks.append("delete_exact 裸句")
            if marks:
                findings.append(_mk(3, i, st, marks))
        # 9 交叉引用耦合（pei 专属；标题/题注本体豁免）
        if 9 in cats and mode == "pei" and not (XREF_SKIP.match(st) and len(st) < 60):
            marks = [m for pat in XREF_PATS for m in pat.findall(st)]
            if marks:
                findings.append(_mk(9, i, st, marks))
        # 4 内部编号
        if 4 in cats:
            marks = E_RE.findall(st) + [t for t in INTERNAL_TOKENS if t in st] + SEG_RE.findall(st)
            if marks:
                findings.append(_mk(4, i, st, marks))
        # 5 二次残渣
        if 5 in cats:
            marks = [t for t in DEBRIS_TOKENS if t in st]
            if paren_unbalanced(st):
                marks.append("段内（）失衡")
            if re.search(r"已在[ 　]*[、，]?[ 　]*详述", st):
                marks.append("断句(已在…详述)")
            if marks:
                findings.append(_mk(5, i, st, marks))
        # 7 口径meta
        if 7 in cats:
            marks = [t for t in META_TOKENS if t in st]
            if marks:
                findings.append(_mk(7, i, st, marks))
        # 8 身份泄漏（正文实名类，main 跳过）
        if 8 in cats and mode != "main":
            marks = [t for t in (rules["identity_banned"] + SELF_REF + PERF_ATTRIB) if t in st]
            bare = _strip_terms(st, YUAN_PROTECT + rules["identity_banned"] + SELF_REF)
            if "院" in bare:
                marks.append(f"裸「院」x{bare.count('院')}")
            if marks:
                findings.append(_mk(8, i, st, marks))

    # 6 断裂引用（文档级：题注编号 vs 正文出现序）
    if 6 in cats:
        findings += _scan_captions(ptexts, rules)

    # 8 非正文 part：工具痕迹（两模式都报）+ 实名类（main 跳过）+ docProps 元数据
    if 8 in cats:
        findings += _scan_parts_identity(parts, mode, rules)

    findings.sort(key=lambda f: (f["cat"], f["para"]))
    return findings


def _scan_captions(ptexts, rules):
    """题注组检查：每组（表/图 × 章前缀）首次出现序须单调 + 编号 1..k 连续。"""
    offs, full, pos = [], [], 0
    for s in ptexts:
        offs.append(pos)
        full.append(s)
        pos += len(s)
    text = "".join(full)

    def para_of(charpos):
        return bisect_right(offs, charpos)

    groups = {}
    for m in CAPTION_RE.finditer(text):
        key = (m.group(1), m.group(2))
        groups.setdefault(key, []).append((int(m.group(3)), m.start()))
    # rules caption_renumber 前缀也拆成组核（通用正则已覆盖标准「表 X-n」形态，此处兜非标前缀）
    for prefix in rules["caption_renumber"]:
        pat = re.compile(re.escape(prefix) + r"(\d+)")
        hits = [(int(m.group(1)), m.start()) for m in pat.finditer(text)]
        if hits:
            groups.setdefault(("R", prefix), []).extend(hits)

    out, seen = [], set()
    for (kind, pref), hits in sorted(groups.items(), key=lambda kv: str(kv[0])):
        first = {}
        for n, p in hits:
            first.setdefault(n, p)
        nums = sorted(first)
        seq = sorted(first, key=lambda n: first[n])
        label = f"{kind} {pref}-" if kind in ("表", "图") else pref
        if label in seen:  # rules caption_renumber 前缀与通用组重叠时去重
            continue
        seen.add(label)
        if seq != nums:
            out.append(_mk(6, para_of(first[seq[0]]), f"题注组「{label}」出现序 {seq} ≠ 编号序 {nums}",
                           [f"{label}{n}" for n in seq]))
        elif nums != list(range(1, len(nums) + 1)):
            out.append(_mk(6, para_of(first[nums[0]]), f"题注组「{label}」编号不连续: {nums}",
                           [f"{label}{n}" for n in nums]))
    return out


def _scan_parts_identity(parts, mode, rules):
    out = []
    xml_parts = {n: b.decode("utf-8", "ignore") for n, b in parts.items()
                 if n.endswith(".xml") or n.endswith(".rels")}
    # 工具痕迹：两模式都报
    for name, xml in xml_parts.items():
        marks = [f"{t} x{xml.count(t)}" for t in TOOL_TRACES if t in xml]
        if marks:
            out.append(_mk(8, 0, f"{name}: 工具痕迹", marks))
    if mode == "main":
        return out
    # 实名类：正文以外的 part（页眉/页脚/styles/docProps 等；document.xml 已逐段报过）
    for name, xml in xml_parts.items():
        if name == "word/document.xml":
            continue
        marks = [f"{t} x{xml.count(t)}" for t in (rules["identity_banned"] + SELF_REF + PERF_ATTRIB)
                 if t in xml]
        if marks:
            out.append(_mk(8, 0, f"{name}: 身份词", marks))
    # docProps 元数据署名
    core = xml_parts.get("docProps/core.xml", "")
    if core:
        croot = etree.fromstring(parts["docProps/core.xml"])
        for tag, label in (("{http://purl.org/dc/elements/1.1/}creator", "creator"),
                           ("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy",
                            "lastModifiedBy")):
            el = croot.find(tag)
            if el is not None and (el.text or "").strip():
                out.append(_mk(8, 0, f"docProps/core.xml: {label}=「{el.text.strip()}」非空",
                               [f"{label} 非空"]))
    return out


def format_findings(findings):
    """findings → 人类可读报告行列表（不含 PASS/FAIL 尾行）。"""
    lines = []
    for f in findings:
        pno = f"P{f['para']:04d}" if f["para"] else "PMETA"
        marks = "、".join(str(m) for m in dict.fromkeys(f["marks"]))[:60]
        lines.append(f"[类别{f['cat']}·{CAT_NAMES[f['cat']]}] {pno} {f['excerpt']}"
                     f"  ⟨命中: {marks}⟩ → {f['advice']}")
    return lines
