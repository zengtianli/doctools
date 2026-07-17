#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_deref — 标书正文交叉引用去耦合（陪标：合稿人会删/调章节，写死编号=断链耦合）。

处理形态（2026-07-16 用户钦定：B/D 整句删除、E/F 一并去耦合）：
  A 括号引用     （7.2）/（详见 7.6.6.3）/（表 7.2-2）      → 整个括号删
  B 动词+章节号  方案细节详见 7.2 / 承接 3.4               → 删从句；全句皆引用→删整句
  C 裸章节号     详见剩余的 7.1.3.2 平原河网双指标法        → 删号留名词
  D 第N章        详见第8章 / 第5章以…统领（叙事主语）       → 删从句；主语位→标 MANUAL
  E/F 表图引用   见表 7.1-1 / 见图 7.4-1                   → 题注在邻近(±5段)→"见下表/下图"；跨位→题注名
删除类操作后自动顺稿（双标点/悬空逗号/空句）。主语位叙事句不自动删，输出 MANUAL 清单人判。

护栏：数字集合差异必须全部属于「已知章节号/表图号 token」；SL/T·GB/T 标准号、
数量词（亿/万/m³/km²/%）、公文文号〔YYYY〕、日期不受影响（机器核验非承诺）。

用法:
  python3 bid_deref.py <docx> [--check] [--manual-pairs <yaml>]   # docx surgical
  python3 bid_deref.py --md <chapters目录> [--check]              # 源 md 同逻辑
exit 0=完成(或 check 通过)；2=有 MANUAL 待人判或护栏红；1=用法/IO 错。
"""
import argparse, glob, os, re, shutil, sys, zipfile
from datetime import datetime
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(t): return f"{{{W}}}{t}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

NUM = r"\d+(?:\.\d+){1,3}"
NUMRANGE = rf"{NUM}(?:\s*[—\-－~～]\s*{NUM})?"
VERBS = r"(?:详见|参见|承接|对应|支撑|衔接|落地到|喂给|指向|见)"
UNIT_AFTER = r"[亿万倍年月日%‰米mkK]|m³|km²|万元|亿元"

def ptext_items(p):
    return [[t, t.text or ""] for r in p.iter(w("r")) for t in r.findall(w("t"))]

def ptext(p): return "".join(it[1] for it in ptext_items(p))

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
    texts = [ptext(p) for p in paras]
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
    print(f"已写 {docx.name} 备份 {os.path.basename(bak)}")
    sys.exit(2 if manual else 0)


def run_md(dirpath: str, check: bool):
    files = sorted(glob.glob(os.path.join(dirpath, "*.md")))
    all_lines = []
    for f in files:
        for ln in open(f, encoding="utf-8"):
            all_lines.append(ln.rstrip("\n"))
    sec, caps = build_maps(all_lines)
    # md 标题形如 "## 7.2 xxx" / "#### 7.1.3.2 xxx"，补进映射
    for f in files:
        for ln in open(f, encoding="utf-8"):
            m = re.match(rf"^#+\s+({NUM})[ 　](.+)$", ln.strip())
            if m: sec[m.group(1)] = m.group(2).strip()
    print(f"[md] 章节映射 {len(sec)} 条 · 题注映射 {len(caps)} 条")
    manual, total = [], 0
    for f in files:
        lines = open(f, encoding="utf-8").read().split("\n")
        out = []
        changed = 0
        for j, ln in enumerate(lines):
            if re.match(rf"^#+\s", ln) or re.match(rf"^(图|表)\s*\d", ln.strip()):
                out.append(ln); continue
            new = deref_text(ln, sec, caps, None, manual, f"{os.path.basename(f)}:{j+1}")
            if new != ln:
                changed += 1
                print(f"--- {os.path.basename(f)}:{j+1}")
                print(f"  旧: {ln[:140]}")
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


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("target", type=str)
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--md", action="store_true")
    ap.add_argument("--manual-pairs", type=str, default=None)
    a = ap.parse_args()
    if a.md:
        run_md(a.target, a.check)
    else:
        run_docx(Path(a.target), a.check, a.manual_pairs)
