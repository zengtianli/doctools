#!/usr/bin/env python3
"""标书正文体例门与定点修复（图文顺序 / 来源展开 / 过渡套话 / 四级编号）。

Why（2026-09-17 海宁标立）：一次性脚本按「段落列表」重排 docx，把表格与 sectPr 甩到
文首、整份结构损坏；又在图后插入方向说反的套话，把页码级来源压成裸号〔S-05〕。
本工具把这四件事收成确定性入口：只改 word/document.xml，其余部件逐字节保留；
每个写入动作先过结构不变量，不过不落盘。

    bid_body.py check  <docx>                                  # 只读门，违规 exit 2
    bid_body.py filler <docx> [--apply]                        # 删题注后的机器过渡句
    bid_body.py h4num  <docx> [--apply]                        # （十九）→ 19）
    bid_body.py figs   <docx> [--apply]                        # 图+题注 挪到正文之后
    bid_body.py cite   <docx> --trace <溯源稿.md> --rules <yaml> [--apply]
    bid_body.py note   <docx> --after <完整标题> --text <一段话> [--apply]   # 标题下插一段正文（幂等）

不带 --apply = 干跑只报告。写回 = 备份 .bak-时间戳 + 原地 + 并发写回门。
体例口径：标题 → 正文 → 图 → 题注；来源写到文件名和页码，内部来源号不进交付正文。
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.append(str(Path(__file__).resolve().parents[2] / "lib"))
from docx_parts import assert_parts_intact  # noqa: E402
from bid_residue_lib import W, ptext, replace_once, w  # noqa: E402
from docx_write_gate import WriteGate  # noqa: E402

NS = {"w": W}
CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
FILLER = re.compile(r"^(下图|下表|如上图所示|如下图所示|如下图表所示)")
MARK = re.compile(r"〔[^〔〕]*S-0\d[^〔〕]*〕")
CAP_PAREN = re.compile(r"（[^（）]*(?:S-0\d|拟议)[^（）]*）。?$")
TRACE_MARK = re.compile(r"\[([^\[\]]*S-0\d[^\[\]]*)\]")


# ── 读写 ─────────────────────────────────────────────────────────
class Doc:
    def __init__(self, path):
        self.path = Path(path)
        self.gate = WriteGate(self.path)
        with zipfile.ZipFile(self.path) as z:
            self.infos = z.infolist()
            self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        self.root = etree.fromstring(self.parts["word/document.xml"])
        self.body = self.root.find(w("body"))
        styles = etree.fromstring(self.parts["word/styles.xml"])
        self.smap = {}
        for s in styles.findall(w("style")):
            n = s.find(w("name"))
            self.smap[s.get(w("styleId"))] = n.get(w("val")) if n is not None else ""

    def style(self, p):
        ps = p.find("w:pPr/w:pStyle", NS)
        return self.smap.get(ps.get(w("val")), "") if ps is not None else "Normal"

    def is_heading(self, el):
        return el.tag == w("p") and self.style(el).lower().startswith("heading")

    def level(self, el):
        m = re.search(r"(\d)", self.style(el))
        return int(m.group(1)) if m else 9

    def is_caption(self, el, kind="图表"):
        if el.tag != w("p"):
            return False
        st, t = self.style(el), ptext(el).strip()
        if "图" in kind and ("图名" in st or re.match(r"图\s*\d+[-.－]\d+", t)):
            return True
        return "表" in kind and ("表名" in st or bool(re.match(r"表\s*\d+[-.－]\d+", t)))

    @staticmethod
    def has_img(el):
        return el.tag == w("p") and el.find(".//w:drawing", NS) is not None

    def is_text(self, el):
        return (el.tag == w("p") and not self.is_heading(el) and not self.is_caption(el)
                and not self.has_img(el) and bool(ptext(el).strip()))

    def save(self):
        self.gate.assert_unchanged()
        bak = self.path.with_name(self.path.name + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
        n = 1
        while bak.exists():                            # 同秒连跑不得盖掉上一份备份
            n += 1
            bak = bak.with_name(bak.name.split("~")[0] + f"~{n}")
        shutil.copy2(self.path, bak)
        self.parts["word/document.xml"] = etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = self.path.with_name(self.path.name + ".tmp-bid-body")
        with zipfile.ZipFile(tmp, "w") as z:
            for i in self.infos:
                z.writestr(i, self.parts[i.filename], compress_type=i.compress_type)
        assert_parts_intact(bak, tmp, allow_changed={"word/document.xml"}, verbose=False)
        tmp.replace(self.path)
        print(f"✓ 备份 {bak.name}\n✓ 写回 {self.path.name}（仅 word/document.xml）")


def _h(el):
    return hashlib.md5(etree.tostring(el)).hexdigest()


def invariants(doc, content=True):
    """结构不变量快照。content=True（只挪/删不改字）时连元素内容一并锁定。"""
    kids = list(doc.body)
    snap = {
        "count": len(kids),
        "sect_last": bool(kids) and kids[-1].tag == w("sectPr"),
        "tags": sorted(k.tag for k in kids),
        "imgs": sum(1 for k in kids if doc.has_img(k)),
        "tbl_has_cap": [doc.is_caption(k.getprevious(), "表") for k in kids
                        if k.tag == w("tbl") and k.getprevious() is not None],
    }
    if content:
        snap["hashes"] = sorted(_h(k) for k in kids)
        snap["tbl_caps"] = [(_h(k), _h(k.getprevious())) for k in kids
                            if k.tag == w("tbl") and k.getprevious() is not None
                            and doc.is_caption(k.getprevious(), "表")]
    return snap


def assert_invariants(doc, before, removed=()):
    after = invariants(doc, content="hashes" in before)
    if removed:
        gone = sorted(removed)
        for key in ("hashes",):
            rest = list(before[key])
            for g in gone:
                rest.remove(g)
            before = dict(before, **{key: rest})
        before["count"] -= len(gone)
        tags = list(before["tags"])
        for _ in gone:
            tags.remove(w("p"))
        before["tags"] = tags
    bad = [k for k in before if before[k] != after[k]]
    if bad or not after["sect_last"]:
        sys.exit(f"⛔ 结构不变量被破坏 {bad or ['sect_last']}，未写回")


# ── 扫描 ─────────────────────────────────────────────────────────
def find_fillers(doc):
    out = []
    for el in doc.body:
        prev = el.getprevious()
        if el.tag != w("p") or prev is None or not doc.is_caption(prev):
            continue
        t = ptext(el).strip()
        core = re.sub(r"^[图表]\s*\d+[-.－]\d+\s*", "", ptext(prev).strip())
        core = re.split(r"[（(。]", core)[0].strip()
        if FILLER.match(t) and len(t) < 90 and (core and core in t or "相关内容如下" in t):
            out.append(el)
    return out


def cn2int(s):
    if "十" in s:
        a, _, b = s.partition("十")
        return CN.get(a, 1) * 10 + CN.get(b, 0)
    return CN.get(s, 0)


def find_h4(doc):
    out = []
    for el in doc.body:
        if doc.is_heading(el) and doc.level(el) == 4:
            m = re.match(r"（([一二三四五六七八九十]+)）", ptext(el).strip())
            if m:
                out.append((el, m.group(0), f"{cn2int(m.group(1))}）"))
    return out


def h4_sequence_issues(doc):
    issues, seq, parent = [], 0, ""
    for el in doc.body:
        if not doc.is_heading(el):
            continue
        if doc.level(el) < 4:
            seq, parent = 0, ptext(el).strip()[:24]
            continue
        m = re.match(r"(\d+)）", ptext(el).strip())
        if not m:
            continue
        n = int(m.group(1))
        if n != seq + 1:
            issues.append(f"{parent} 下编号 {seq}→{n}：{ptext(el).strip()[:30]}")
        seq = n
    return issues


def plan_figs(doc):
    """图紧跟标题 → 计划挪到其后第一段连续正文的末尾。返回 [(块, 停靠元素, 说明)]。"""
    kids, plans, stuck = list(doc.body), [], []
    for i, el in enumerate(kids):
        if not doc.has_img(el) or i == 0 or not doc.is_heading(kids[i - 1]):
            continue
        block = [el]
        j = i + 1
        if j < len(kids) and doc.is_caption(kids[j], "图"):
            block.append(kids[j])
            j += 1
        cap = ptext(block[-1]).strip()[:28] if len(block) > 1 else "（无题注）"
        while j < len(kids) and doc.is_heading(kids[j]):
            j += 1
        k = j
        while k < len(kids) and doc.is_text(kids[k]):
            k += 1
        nxt = kids[k] if k < len(kids) else None       # 末段若是引出表格的短行，图停在它之前，不拆散表
        if (k > j and nxt is not None and (nxt.tag == w("tbl") or doc.is_caption(nxt, "表"))
                and len(ptext(kids[k - 1]).strip()) < 30
                and not re.search(r"[。；！？]$", ptext(kids[k - 1]).strip())):
            k -= 1
        if k == j:
            stuck.append(f"{cap}：其后没有可承接的正文段")
            continue
        plans.append((block, kids[k], f"{cap} → 下移 {k - j} 段正文之后，止于「{ptext(kids[k - 1]).strip()[-18:]}」"))
    return plans, stuck


def scan(doc):
    kids = list(doc.body)
    f = {}
    f["结构"] = ([] if kids and kids[-1].tag == w("sectPr") else ["sectPr 不在 body 末位"]) + \
        [f"表格位于首个标题之前（第{i}位）" for i, k in enumerate(kids) if k.tag == w("tbl")
         and not any(doc.is_heading(x) for x in kids[:i])]
    plans, stuck = plan_figs(doc)
    f["图紧跟标题"] = [p[2].split(" → ")[0] for p in plans] + stuck
    f["过渡套话"] = [ptext(e).strip()[:40] for e in find_fillers(doc)]
    h4 = find_h4(doc)
    f["四级编号"] = [f"{old}→{new} {ptext(e).strip()[:24]}" for e, old, new in h4] + h4_sequence_issues(doc)
    allp = list(doc.root.iter(w("p")))
    f["内部来源号"] = [ptext(p).strip()[:36] for p in allp
                  if re.search(r"S-0\d", ptext(p))]
    f["过程稿式来源写法"] = [ptext(p).strip()[:36] for p in allp if re.search(r"(?<!不)同规划(PDF|第|，|印刷)|(PDF|印刷)第\d", ptext(p))]
    return f


# ── 子命令 ───────────────────────────────────────────────────────
def cmd_check(a):
    doc = Doc(a.docx)
    f = scan(doc)
    print(f"[门检] {doc.path.name}")
    bad = 0
    for k, v in f.items():
        print(f"  {'✗' if v else '✓'} {k}: {len(v)}")
        for x in v[: a.show]:
            print(f"      · {x}")
        bad += len(v)
    sys.exit(2 if bad else 0)


def cmd_filler(a):
    doc = Doc(a.docx)
    before = invariants(doc)
    hits = find_fillers(doc)
    for e in hits:
        print("  删:", ptext(e).strip())
    print(f"共 {len(hits)} 句")
    if a.apply and hits:
        removed = [_h(e) for e in hits]
        for e in hits:
            doc.body.remove(e)
        assert_invariants(doc, before, removed)
        doc.save()


def cmd_h4num(a):
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    hits = find_h4(doc)
    for e, old, new in hits:
        print(f"  {old} → {new}  {ptext(e).strip()[len(old):][:30]}")
    if a.apply and hits:
        for e, old, new in hits:
            replace_once(e, old, new)
    issues = h4_sequence_issues(doc)
    print(f"共 {len(hits)} 处；改后序号断点 {len(issues)}")
    for x in issues:
        print("  ⚠", x)
    if a.apply and hits:
        if issues and not a.force:
            sys.exit("⛔ 改后序号不连续，未写回（确认无误可加 --force）")
        assert_invariants(doc, before)
        doc.save()


def cmd_figs(a):
    doc = Doc(a.docx)
    before = invariants(doc)
    plans, stuck = plan_figs(doc)
    for _, _, note in plans:
        print("  ", note)
    for s in stuck:
        print("  ⚠ 未处理:", s)
    print(f"共 {len(plans)} 张待挪，{len(stuck)} 张需人工")
    if a.apply and plans:
        for block, stop, _ in plans:
            for el in block:
                doc.body.remove(el)
            for el in block:
                stop.addprevious(el)
        assert_invariants(doc, before)
        if plan_figs(doc)[0]:
            sys.exit("⛔ 挪后仍有图紧跟标题，未写回")
        doc.save()


# ── 来源展开 ─────────────────────────────────────────────────────
def load_rules(path):
    import yaml
    r = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    r.setdefault("aliases", {})
    r.setdefault("overrides", {})
    r.setdefault("pdf_offset", {})
    r.setdefault("drop_notes", [])
    return r


def _shift(expr, off):
    return re.sub(r"\d+", lambda m: str(int(m.group(0)) - off), expr)


def render(detail, rules, errs):
    """溯源稿方括号内的明细 → 读者可读的出处；无法确定的进 errs，不猜。"""
    if detail in rules["overrides"]:
        v = rules["overrides"][detail]
        return v if not v or v.startswith("（") else "（" + v + "）"
    out, sid, last = [], None, None
    for seg in [s.strip() for s in re.split(r"[；;]", detail) if s.strip()]:
        if seg in rules["overrides"]:
            if rules["overrides"][seg]:
                out.append(rules["overrides"][seg])
            continue
        m = re.match(r"(S-0\d(?:、S-0\d)*)[，,]?\s*(.*)$", seg)
        rest = seg
        if m:
            sid, rest = m.group(1), m.group(2)
        if not rest:
            errs.append(f"裸号无明细: {detail}")
            continue
        if rest in rules["overrides"]:
            if rules["overrides"][rest]:
                out.append(rules["overrides"][rest])
            continue
        if rest in rules["drop_notes"]:
            continue
        if "PDF第" not in rest and "印刷第" not in rest:
            if sid == "S-01" and m:
                rest = rest if re.match(r"招标文件|合同", rest) else "招标文件" + rest
            elif m and sid != "S-01":
                errs.append(f"无页码且无 override: {seg}")
                continue
            out.append(rest)
            continue
        if sid == "S-01":
            rest = rest if re.match(r"招标文件|合同", rest) else "招标文件" + re.sub(r"^招标需求", "“招标需求”", rest)
            out.append(rest.replace("印刷", ""))
            continue
        name = None
        for k in sorted(rules["aliases"], key=len, reverse=True):
            if k in rest:
                name, rest = rules["aliases"][k], rest.replace(k, "", 1)
                break
        if name is None and re.match(r"(同规划)?，?(PDF|印刷)第", rest) and last:
            name, rest = last, rest.replace("同规划", "", 1)
        if name is None:
            errs.append(f"文件名无法识别: {seg}")
            continue
        last = name
        rest = rest.lstrip("，, 》")
        if "印刷" in rest:
            pages = rest.split("印刷", 1)[1]
        else:
            off = rules["pdf_offset"].get(name)
            pm = re.match(r"PDF(第[\d、—\-–第页]+页)(.*)$", rest)
            if off is None or not pm:
                errs.append(f"仅有PDF页码且无已核偏移: {seg}")
                continue
            pages = _shift(pm.group(1), off) + pm.group(2)
        if re.search(r"PDF|印刷", pages):
            errs.append(f"页码式过于复杂，需 override: {seg}")
            continue
        if out and out[-1].startswith(name + "第") and pages.startswith("第"):
            out[-1] += "、" + pages                     # 同一文件的相邻页次并成一条
        else:
            out.append(name + pages)
    return "（" + "；".join(out) + "）" if out else ""


def norm(t):
    t = MARK.sub("", TRACE_MARK.sub("", t))
    return re.sub(r"[\s*#>|`]+", "", t)


def cmd_cite(a):
    doc = Doc(a.docx)
    rules = load_rules(a.rules)
    before = invariants(doc, content=False)
    trace, last_alias = {}, None
    names = sorted(rules["aliases"], key=len, reverse=True)
    for line in Path(a.trace).read_text(encoding="utf-8").splitlines():
        ds = []
        for d in TRACE_MARK.findall(line):
            segs = []
            for seg in re.split(r"[；;]", d):          # 「同规划」跨段指代 → 按溯源稿顺序还原成上一份具名文件
                if "同规划" in seg and last_alias:
                    seg = seg.replace("同规划", last_alias, 1)
                hit = next((k for k in names if k in seg), None)
                last_alias = hit or last_alias
                segs.append(seg)
            ds.append("；".join(segs))
        if ds:
            trace.setdefault(norm(line), ds)
    keys = list(trace)
    errs, todo = [], []
    reps = [x for key, xs in (rules.get("replace") or {}).items() if key in doc.path.name for x in xs]
    for old, new in reps:                              # 整句改写（表注等），按文件名分组，须恰好命中一段
        hit = [p for p in doc.root.iter(w("p")) if old in ptext(p)]
        if not hit and sum(new in ptext(p) for p in doc.root.iter(w("p"))) == 1:
            continue                                   # 已改过，幂等跳过
        if len(hit) != 1:
            errs.append(f"replace 命中 {len(hit)} 段（应为 1）: {old}")
            continue
        replace_once(hit[0], old, new)
        print(f"  整句 {old} → {new}")
    for p in doc.root.iter(w("p")):
        full = ptext(p)
        marks = MARK.findall(full)
        if not marks or doc.is_caption(p):
            continue
        k = norm(full)
        if k not in trace:
            near = difflib.get_close_matches(k, keys, n=1, cutoff=0.8)
            k = near[0] if near else None
        if k is None:
            errs.append(f"溯源稿找不到对应段: {full[:40]}")
            continue
        ds = trace[k]
        if len(ds) != len(marks):
            errs.append(f"来源号数量不一致 docx{len(marks)}≠md{len(ds)}: {full[:40]}")
            continue
        todo.append((p, [(mk, render(d, rules, errs), d) for mk, d in zip(marks, ds)]))
    caps = [p for p in doc.root.iter(w("p")) if doc.is_caption(p) and CAP_PAREN.search(ptext(p).strip())]
    cap_plan = []
    for p in caps:
        old = CAP_PAREN.search(ptext(p).strip()).group(0)
        inner = old.strip("。")[1:-1]
        new = rules["overrides"].get(inner)
        if new is None and "S-05" in inner:
            errs.append(f"题注来源需 override: {inner}")
            continue
        cap_plan.append((p, old, new or ""))
    cell_plan = []
    for tbl in doc.root.iter(w("tbl")):
        last = [None]
        for p in tbl.iter(w("p")):
            t = ptext(p).strip()
            if not re.search(r"(PDF|印刷)第\d", t) or MARK.search(t) or len(t) > 80:
                continue
            for k in sorted(rules["aliases"], key=len, reverse=True):
                if k in t:
                    last[0] = k
                    break
            src = t.replace("同规划", last[0] or "同规划", 1) if t.startswith("同规划") else t
            new = rules["overrides"].get(t) or render("S-05，" + src, rules, errs)[1:-1]
            if new:
                cell_plan.append((p, t, new))
    n = sum(len(x[1]) for x in todo)
    for p, items in todo:
        for mk, new, d in items:
            print(f"  [{d}]\n      → {new or '（删除标记）'}")
    for p, old, new in cap_plan:
        print(f"  题注 {ptext(p).strip()[:14]}… {old} → {new or '（去括注）'}")
    for p, old, new in cell_plan:
        print(f"  表格来源格 {old} → {new}")
    print(f"正文来源 {n} 处，题注 {len(cap_plan)} 处，表格来源格 {len(cell_plan)} 处，问题 {len(errs)}")
    for e in sorted(set(errs)):
        print("  ⚠", e)
    if not a.apply:
        return
    if errs:
        sys.exit("⛔ 有未决来源，未写回（补 rules 的 overrides/aliases 后重跑）")
    for p, items in todo:
        for mk, new, _ in items:
            full = ptext(p)
            i = full.find(mk)
            lead = full[i - 1] if i > 0 and full[i - 1] in "。；！？" else ""
            if not replace_once(p, lead + mk, new + lead):
                sys.exit(f"⛔ 替换失败: {full[:40]}")
    for p, old, new in cap_plan + cell_plan:
        if not replace_once(p, old, new):
            sys.exit(f"⛔ 替换失败: {old}")
    left = [ptext(p)[:30] for p in doc.root.iter(w("p")) if re.search(r"S-0\d|(PDF|印刷)第\d", ptext(p))]
    if left:
        sys.exit(f"⛔ 仍有内部来源号残留 {len(left)} 处，未写回: {left[:3]}")
    assert_invariants(doc, before)
    doc.save()


def cmd_note(a):
    """在指定标题后插入一段正文：克隆其后最近一段正文的段落/字符格式，已存在同文则跳过。"""
    import copy
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    if any(ptext(p).strip() == a.text.strip() for p in doc.body.iter(w("p"))):
        print("已存在同文段落，跳过")
        return
    heads = [k for k in doc.body if doc.is_heading(k) and ptext(k).strip() == a.after.strip()]
    if len(heads) != 1:
        sys.exit(f"⛔ 标题「{a.after}」命中 {len(heads)} 处（应为 1）")
    model = next((k for k in heads[0].itersiblings() if doc.is_text(k) and len(ptext(k)) > 40), None)
    if model is None:
        sys.exit("⛔ 找不到可克隆格式的正文段")
    new = copy.deepcopy(model)
    runs = new.findall(w("r"))
    for r in runs[1:]:
        new.remove(r)
    for extra in new.xpath("./*[not(self::w:pPr) and not(self::w:r)]", namespaces=NS):
        new.remove(extra)
    for t in runs[0].findall(w("t"))[1:]:
        runs[0].remove(t)
    runs[0].find(w("t")).text = a.text.strip()
    for attr in list(new.attrib):                       # paraId 等须唯一，克隆件不带
        del new.attrib[attr]
    print(f"  「{a.after}」之后插入：{a.text.strip()}")
    if a.apply:
        heads[0].addnext(new)
        before["count"] += 1
        before["tags"] = sorted(before["tags"] + [w("p")])
        assert_invariants(doc, before)
        doc.save()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in [("check", cmd_check), ("filler", cmd_filler), ("h4num", cmd_h4num),
                     ("figs", cmd_figs), ("cite", cmd_cite), ("note", cmd_note)]:
        s = sub.add_parser(name)
        s.add_argument("docx")
        s.set_defaults(fn=fn)
        if name == "check":
            s.add_argument("--show", type=int, default=4)
        else:
            s.add_argument("--apply", action="store_true")
        if name == "h4num":
            s.add_argument("--force", action="store_true")
        if name == "note":
            s.add_argument("--after", required=True)
            s.add_argument("--text", required=True)
        if name == "cite":
            s.add_argument("--trace", required=True)
            s.add_argument("--rules", required=True)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
