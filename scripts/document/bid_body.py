#!/usr/bin/env python3
"""标书正文体例门与定点修复（图文顺序 / 来源展开 / 过渡套话 / 条目编号）。

Why（2026-09-17 海宁标立）：一次性脚本按「段落列表」重排 docx，把表格与 sectPr 甩到
文首、整份结构损坏；又在图后插入方向说反的套话，把页码级来源压成裸号〔S-05〕。
本工具把这四件事收成确定性入口：只改 word/document.xml，其余部件逐字节保留；
每个写入动作先过结构不变量，不过不落盘。

    bid_body.py check  <docx>                                  # 只读门，违规 exit 2
    bid_body.py filler <docx> [--apply]                        # 删题注后的机器过渡句
    bid_body.py h4num  <docx>                                  # 已退役：标题不再用 N）编号，见 cmd_h4num
    bid_body.py figs   <docx> [--apply]                        # 图+题注 挪到正文之后
    bid_body.py cite   <docx> --trace <溯源稿.md> --rules <yaml> [--apply]
    bid_body.py h3group <docx> --plan <yaml> [--apply]          # 二级标题下平铺的 N）条目分组：插三级标题+引导段，组内重排 1）
    bid_body.py uncite <docx> [--apply]                        # 清掉正文里的页码级出处括注（先用出处对照留底）
    bid_body.py resub  <docx> --rules <yaml> [--apply]          # 按项目 rules 的 resub 列表做保格式正则替换（去“拟”等过程稿口气）
    bid_body.py media  <docx> --old <旧图目录> --new <新图目录>… [--apply]  # 内嵌图按旧图哈希匹配、同名同像素尺寸替换
    bid_body.py swapfig <docx> --plan <plan.json> [--apply]   # 重画图按图号换入：可改比例（按版宽重设尺寸）、可改图题名，旧图部件无引用即移除
    bid_body.py note   <docx> --after <完整标题> --text <一段话> [--apply]   # 标题下插一段正文（幂等）
    bid_body.py insert <docx> --patch <补丁.md>… [--apply]     # 定稿扩充：按锚点段原文插入多段，原段一字不动（幂等）

不带 --apply = 干跑只报告。写回 = 备份 .bak-时间戳 + 原地 + 并发写回门。
体例口径：标题逐级下挂（1 / 1.1 / 1.1.1 / 1.1.1.1），1）2）是正文段、不用标题样式，编号在每个上级标题下
从 1）起；标题 → 正文 → 图 → 题注；交付正文靠句内点名交代出处（“《××规划》记载…”），
页码级来源留在稿外的出处对照表，不以括注形式进正文；内部来源号、过程稿口气（“拟”）不进交付正文。
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


ITEM = re.compile(r"(\d+)）")
HEAD_ITEM = re.compile(r"\d+）|（[一二三四五六七八九十]+）")


def heading_items(doc):
    """标题样式里的 N）/（十九）：条目编号属于正文段，标题只用 1.1.1.1 式数字编号。"""
    return [el for el in doc.body if doc.is_heading(el) and HEAD_ITEM.match(ptext(el).strip())]


def item_sequence_issues(doc, headings_as_items=False):
    """正文条目 N）在每个上级标题下从 1）连续编号；遇到标题清零。同一标题下可以有几组并列
    清单（如按类别列依据），各组从 1）重起不算断点；跳号、重号、不从 1）起才报。
    headings_as_items=True 时把「N）开头的标题」也当条目（h3group 处理旧稿用）。"""
    issues, seq, parent = [], 0, ""
    for el in doc.body:
        if el.tag != w("p"):
            continue
        t = ptext(el).strip()
        m = ITEM.match(t)
        if doc.is_heading(el) and not (headings_as_items and m):
            seq, parent = 0, t[:24]
            continue
        if not m or not (doc.is_text(el) or doc.is_heading(el)):
            continue
        n = int(m.group(1))
        if n != seq + 1 and n != 1:
            issues.append(f"{parent} 下编号 {seq}→{n}：{t[:30]}")
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
    f["标题样式N）"] = [ptext(e).strip()[:30] for e in heading_items(doc)]
    f["条目编号"] = item_sequence_issues(doc)
    skip, last = [], None                              # 标题层级跳级：1 / 1.1 / 1.1.1 / 1.1.1.1，1）为正文段
    for k in kids:
        if doc.is_heading(k):
            if last is not None and doc.level(k) > doc.level(last) + 1:
                skip.append(f"「{ptext(last).strip()[:22]}」下直接挂 {doc.level(k)} 级「{ptext(k).strip()[:16]}」")
            last = k
    f["标题跳级"] = skip
    allp = list(doc.root.iter(w("p")))
    f["内部来源号"] = [ptext(p).strip()[:36] for p in allp
                  if re.search(r"S-0\d", ptext(p))]
    f["正文出处括注"] = [x[:40] for p in allp for x in find_src_parens(ptext(p))]
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
    """已退役（2026-09-18）。原做法把四级标题的（十九）改成 19），仍是标题样式里的 N）；
    现行体例四级标题为 1.1.1.1，1）2）是正文段。标题样式 N）由 check 报红，改稿按体例人工或另行降为正文段。"""
    sys.exit("⛔ h4num 已退役：标题不再用 N）编号（四级标题用 1.1.1.1，1）2）为正文段）。"
             "先跑 check 查「标题样式N）」。")


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


SRC_PAREN = re.compile(r"（((?:[^（）]|（[^（）]*）)*)）")
SRC_HEAD = re.compile(r"^(?:来源：)?(?:招标文件|合同第|《|海宁市20\d\d|海宁市非居民|夹浦)")
STUB = re.compile(r"^[^。；，]{0,12}[：据见]?[。；]?$")


def find_src_parens(text):
    return [m.group(0) for m in SRC_PAREN.finditer(text) if SRC_HEAD.match(m.group(1))]


def cmd_uncite(a):
    """交付正文不挂页码级出处括注：删括注、顺手收拾删后留下的残句；整段只剩引导残句的删段。"""
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    n, drop, odd = 0, [], []
    for p in list(doc.root.iter(w("p"))):
        full = ptext(p)
        hits = find_src_parens(full)
        if not hits:
            continue
        rest = full
        for h in hits:
            rest = rest.replace(h, "", 1)
        rest = rest.strip()
        if STUB.match(rest) and p.getparent() is doc.body and not doc.is_caption(p):
            drop.append(p)
            print(f"  删段: {full[:60]}")
            n += len(hits)
            continue
        for h in hits:
            cur = ptext(p) if a.apply else full
            i = cur.find(h)
            start = max(cur.rfind("。", 0, i), cur.rfind("；", 0, i)) + 1
            m = re.search(r"[。；]", cur[i + len(h):])
            end = i + len(h) + (m.end() if m else 0)
            sent = cur[start:end]
            target = h
            if len(sent.replace(h, "").strip("。；")) < 14:        # 括注前只剩“表中…分别见”这类引导残句 → 整句删
                target = sent
                print(f"  删句: {sent[:70]}")
            elif cur[i - 1:i] in "见：，" or (cur[i - 1:i] == "据" and cur[i - 2:i] not in ("依据", "证据", "数据", "凭据")):
                odd.append(cur[max(0, i - 20):i + 30])
            else:
                print(f"  删括注: …{cur[max(0, i - 16):i]}|{h[:70]}")
            if a.apply and not replace_once(p, target, ""):
                sys.exit(f"⛔ 替换失败: {target[:40]}")
            n += 1
    print(f"括注 {n} 处，其中整段删除 {len(drop)} 段；需人工看的残句 {len(odd)}")
    for x in odd:
        print("  ⚠", x)
    if a.apply:
        if odd:
            sys.exit("⛔ 有残句，未写回（先处理或写进 replace 规则）")
        for p in drop:
            doc.body.remove(p)
        before["count"] -= len(drop)
        tags = list(before["tags"])
        for _ in drop:
            tags.remove(w("p"))
        before["tags"] = tags
        assert_invariants(doc, before)
        doc.save()


def replace_span(p, a, b, new):
    """按段内字符区间 [a,b) 保格式替换（跨 run）；replace_once 的按位置版。"""
    pos, done = 0, False
    for t in [t for r in p.iter(w("r")) for t in r.findall(w("t"))]:
        txt = t.text or ""
        s0, e0 = pos, pos + len(txt)
        pos = e0
        if e0 <= a or s0 >= b:
            continue
        left = txt[:a - s0] if s0 < a else ""
        right = txt[b - s0:] if e0 > b else ""
        t.text = left + ("" if done else new) + right
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        done = True
    return done


def cmd_resub(a):
    """rules.resub = [[正则, 替换], …] 依次作用于每个段落；护栏：去空白后的数字序列不变。"""
    rules = load_rules(a.rules)
    subs = [(re.compile(pat), rep) for pat, rep in rules.get("resub") or []]
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    total, shown = 0, 0
    for p in doc.root.iter(w("p")):
        orig = ptext(p)
        for rx, rep in subs:
            for m in reversed(list(rx.finditer(ptext(p)))):
                cur = ptext(p)
                if shown < a.show:
                    print(f"  {cur[max(0, m.start() - 10):m.end() + 8]} → {cur[max(0, m.start() - 10):m.start()]}{m.expand(rep)}{cur[m.end():m.end() + 8]}")
                    shown += 1
                replace_span(p, m.start(), m.end(), m.expand(rep))
                total += 1
        if re.findall(r"\d+(?:\.\d+)?", orig) != re.findall(r"\d+(?:\.\d+)?", ptext(p)):
            sys.exit(f"⛔ 数字序列被改动，未写回: {orig[:50]}")
    print(f"替换 {total} 处")
    if a.apply and total:
        assert_invariants(doc, before)
        doc.save()


def cmd_media(a):
    """图源改过后把新 PNG 换进 docx：内嵌图字节哈希 == 旧图目录里某文件 → 用新图目录里的同名文件替换。
    只换 word/media 下命中的部件，像素尺寸必须一致（版面不变），document.xml 不动。"""
    import io
    from PIL import Image
    path = Path(a.docx)
    gate = WriteGate(path)
    olds = {hashlib.md5(f.read_bytes()).hexdigest(): f.name for f in Path(a.old).glob("*.png")}
    news = {f.name: f for d in a.new for f in Path(d).glob("*.png")}
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    changed = set()
    for name, data in parts.items():
        if not name.startswith("word/media/"):
            continue
        src = olds.get(hashlib.md5(data).hexdigest())
        if not src or src not in news:
            continue
        nb = news[src].read_bytes()
        if Image.open(io.BytesIO(nb)).size != Image.open(io.BytesIO(data)).size:
            sys.exit(f"⛔ {src} 新旧像素尺寸不一致，未写回")
        if nb != data:
            parts[name] = nb
            changed.add(name)
            print(f"  {name[11:]} ← {src}")
    print(f"替换 {len(changed)} 张")
    if a.apply and changed:
        gate.assert_unchanged()
        bak = path.with_name(path.name + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
        n = 1
        while bak.exists():
            n += 1
            bak = bak.with_name(bak.name.split("~")[0] + f"~{n}")
        shutil.copy2(path, bak)
        tmp = path.with_name(path.name + ".tmp-bid-body")
        with zipfile.ZipFile(tmp, "w") as z:
            for i in infos:
                z.writestr(i, parts[i.filename], compress_type=i.compress_type)
        assert_parts_intact(bak, tmp, allow_changed=changed, verbose=False)
        tmp.replace(path)
        print(f"✓ 备份 {bak.name}\n✓ 写回 {path.name}（仅 {len(changed)} 个 media 部件）")


def cmd_swapfig(a):
    """重画后的图按图号换进 docx：允许新图比例与旧图不同（按版宽重设显示尺寸），可同时改图题名称（图号不动）。
    plan.json = [{"no": "7-2", "png": "新图.png", "caption": "新图名（可省）"}, …]。
    新图作为新图片部件写入并改指向；旧图片部件无人引用即移除，不留孤儿 media（health gate 会判红）。"""
    import io
    import json
    from PIL import Image
    emu_cm = 360000
    max_w, max_h = int(a.width * emu_cm), int(a.max_height * emu_cm)
    plan = json.loads(Path(a.plan).read_text(encoding="utf-8"))
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    RELS = "word/_rels/document.xml.rels"
    PR = "http://schemas.openxmlformats.org/package/2006/relationships"
    IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    rels = etree.fromstring(doc.parts[RELS])
    used_ids = {r.get("Id") for r in rels}
    changed, added = {"word/document.xml", RELS}, set()
    old_rids = []
    for it in plan:
        no = it["no"]
        caps = [k for k in doc.body if doc.is_caption(k, "图") and re.match(rf"图\s*{re.escape(no)}(\s|$)", ptext(k).strip())]
        if len(caps) != 1:
            sys.exit(f"⛔ 图{no} 题注命中 {len(caps)} 处（应为 1）")
        cap = caps[0]
        pic = next((k for k in (cap.getprevious(), cap.getnext()) if k is not None and doc.has_img(k)), None)
        if pic is None:
            sys.exit(f"⛔ 图{no} 题注前后都没有图片段")
        blips = pic.findall(f".//{{{A}}}blip")
        if len(blips) != 1:
            sys.exit(f"⛔ 图{no} 所在段有 {len(blips)} 个图片引用（应为 1）")
        png = Path(it["png"]).expanduser()
        data = png.read_bytes()
        pw, ph = Image.open(io.BytesIO(data)).size
        cx = max_w
        cy = int(cx * ph / pw)
        if cy > max_h:
            cy, cx = max_h, int(max_h * pw / ph)
        n = 1
        while f"rId{n}" in used_ids:
            n += 1
        rid = f"rId{n}"
        used_ids.add(rid)
        k = 1
        while f"word/media/swapfig{k}.png" in doc.parts:
            k += 1
        part = f"word/media/swapfig{k}.png"
        doc.parts[part] = data
        added.add(part)
        etree.SubElement(rels, f"{{{PR}}}Relationship", Id=rid, Type=IMG, Target=f"media/swapfig{k}.png")
        old_rids.append(blips[0].get(f"{{{R}}}embed"))
        blips[0].set(f"{{{R}}}embed", rid)
        for tag in (f"{{{WP}}}extent", f"{{{A}}}ext"):
            for e in pic.iter(tag):
                e.set("cx", str(cx)); e.set("cy", str(cy))
        msg = f"  图{no}  {png.name} {pw}×{ph} → {cx / emu_cm:.1f}×{cy / emu_cm:.1f}cm"
        if it.get("caption"):
            m = re.match(rf"(图\s*{re.escape(no)}\s*)(.*)$", ptext(cap).strip())
            if m.group(2) != it["caption"] and not replace_once(cap, m.group(2), it["caption"]):
                sys.exit(f"⛔ 图{no} 题注改名失败")
            msg += f"  题注：{ptext(cap).strip()}"
        print(msg)
    # 旧图片关系：正文不再引用即删；其指向的 media 部件别处也不引用即删
    xml = etree.tostring(doc.root).decode()
    removed = set()
    for rid in set(old_rids):
        if f'"{rid}"' in xml:
            continue
        rel = next(r for r in rels if r.get("Id") == rid)
        target = "word/" + rel.get("Target").lstrip("/").removeprefix("word/")
        rels.remove(rel)
        still = any(r.get("Target").endswith(target[5:]) for r in rels) or any(
            n.endswith(".rels") and n != RELS and target[5:].encode() in doc.parts[n] for n in doc.parts)
        if not still and target in doc.parts:
            del doc.parts[target]
            removed.add(target)
    ct = doc.parts["[Content_Types].xml"]
    if b'Extension="png"' not in ct:
        doc.parts["[Content_Types].xml"] = ct.replace(
            b"</Types>", b'<Default Extension="png" ContentType="image/png"/></Types>')
        changed.add("[Content_Types].xml")
    print(f"换图 {len(plan)} 张；移除旧图片部件 {len(removed)} 个")
    if not a.apply:
        return
    assert_invariants(doc, before)
    doc.gate.assert_unchanged()
    path = doc.path
    bak = path.with_name(path.name + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    n = 1
    while bak.exists():
        n += 1
        bak = bak.with_name(bak.name.split("~")[0] + f"~{n}")
    shutil.copy2(path, bak)
    doc.parts["word/document.xml"] = etree.tostring(doc.root, xml_declaration=True, encoding="UTF-8", standalone=True)
    doc.parts[RELS] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_name(path.name + ".tmp-bid-body")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for i in doc.infos:
            if i.filename in doc.parts:
                z.writestr(i, doc.parts[i.filename], compress_type=i.compress_type)
        for name in sorted(added):
            z.writestr(name, doc.parts[name])
    from docx_parts import diff_parts, media_census
    d = diff_parts(bak, tmp, changed)
    bad = [f"丢失 {x}" for x in d.lost if x not in removed] + [f"改写 {x}" for x in d.changed] + \
          [f"新增 {x}" for x in d.added if x not in added]
    c0, c1 = media_census(bak), media_census(tmp)
    if c1["dangling"] or c1["refs"] != c0["refs"]:
        bad.append(f"图引用 {c0['refs']}→{c1['refs']}，悬空 {c1['dangling']}")
    if bad:
        tmp.unlink()
        sys.exit("⛔ 部件校验未通过，未写回：" + "；".join(bad))
    tmp.replace(path)
    print(f"✓ 备份 {bak.name}\n✓ 写回 {path.name}（document.xml、关系表；新增 {len(added)} 张、移除 {len(removed)} 张图片部件）")


def _clone(model, text):
    """克隆段落外壳（pPr + 首个 run 的 rPr），换成一段纯文本；不带 paraId/书签/图。"""
    import copy
    new = copy.deepcopy(model)
    runs = new.findall(w("r"))
    for extra in new.xpath("./*[not(self::w:pPr)]", namespaces=NS):
        if extra is not runs[0]:
            new.remove(extra)
    for child in list(runs[0]):
        if child.tag not in (w("rPr"),):
            runs[0].remove(child)
    t = etree.SubElement(runs[0], w("t"))
    t.text = text
    for attr in list(new.attrib):
        del new.attrib[attr]
    return new


def cmd_h3group(a):
    """把某二级标题下平铺的 N）条目按计划分组：每组前插三级标题与引导段，组内条目编号从 1）重排，
    可在组末追加新条目（正文段「m）…」及其后续段）。条目以 N）开头的正文段为准，旧稿里 N）开头的
    四级标题也计入并照样重排（它们仍会被 check 报「标题样式N）」）。只插入、不移动既有内容。"""
    import yaml
    plan = yaml.safe_load(Path(a.plan).read_text(encoding="utf-8"))
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    kids = list(doc.body)
    h3_model = next((k for k in kids if doc.is_heading(k) and doc.level(k) == 3), None)
    tx_model = next((k for k in kids if doc.is_text(k) and len(ptext(k)) > 60
                     and k.find(".//w:drawing", NS) is None), None)
    if None in (h3_model, tx_model):
        sys.exit("⛔ 文档内找不到可克隆的三级标题/正文段")
    titles = {ptext(k).strip() for k in kids if doc.is_heading(k)}
    added = 0
    for sec in plan["sections"]:
        heads = [k for k in kids if doc.is_heading(k) and doc.level(k) == 2 and ptext(k).strip() == sec["h2"]]
        if len(heads) != 1:
            sys.exit(f"⛔ 二级标题「{sec['h2']}」命中 {len(heads)} 处（应为 1）")
        items, end = [], None
        for k in heads[0].itersiblings():
            if k.tag == w("sectPr") or (doc.is_heading(k) and doc.level(k) <= 2):
                end = k
                break
            if doc.is_heading(k) and doc.level(k) == 3:
                sys.exit(f"⛔ 「{sec['h2']}」下已有三级标题「{ptext(k).strip()}」，不重复分组")
            if (doc.is_text(k) or doc.is_heading(k)) and ITEM.match(ptext(k).strip()):
                items.append(k)
        if sum(g["n"] for g in sec["groups"]) != len(items):
            sys.exit(f"⛔ 「{sec['h2']}」下 N）条目 {len(items)} 个 ≠ 计划 {sum(g['n'] for g in sec['groups'])} 个")
        print(f"{sec['h2']}（{len(items)} 条 → {len(sec['groups'])} 组）")
        i = 0
        for g in sec["groups"]:
            if g["title"] in titles:
                sys.exit(f"⛔ 标题「{g['title']}」已存在")
            members = items[i:i + g["n"]]
            i += g["n"]
            anchor_end = items[i] if i < len(items) else end
            members[0].addprevious(_clone(h3_model, g["title"]))
            added += 1
            for para in ([g["lead"]] if isinstance(g.get("lead"), str) else g.get("lead") or []):
                members[0].addprevious(_clone(tx_model, para.strip()))
                added += 1
            print(f"  {g['title']}")
            for n, el in enumerate(members, 1):
                old = ITEM.match(ptext(el).strip())
                if old.group(0) != f"{n}）":
                    replace_once(el, old.group(0), f"{n}）")
                print(f"      {ptext(el).strip()[:40]}")
            for m, blk in enumerate(g.get("append") or [], len(members) + 1):
                head = blk.get("item") or blk["h4"]          # h4 为旧键名，仍按正文条目生成
                anchor_end.addprevious(_clone(tx_model, f"{m}）{head}"))
                added += 1
                for para in blk.get("paras") or []:
                    anchor_end.addprevious(_clone(tx_model, para.strip()))
                    added += 1
                print(f"      {m}）{head}  [新增 {len(blk.get('paras') or [])} 段]")
    issues = item_sequence_issues(doc, headings_as_items=True)
    print(f"新增段落 {added}；条目序号断点 {len(issues)}")
    for x in issues:
        print("  ⚠", x)
    if a.apply:
        if issues:
            sys.exit("⛔ 序号不连续，未写回")
        before["count"] += added
        before["tags"] = sorted(before["tags"] + [w("p")] * added)
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


def _parse_patch(paths):
    """补丁格式：`@after <锚点段原文或其唯一开头>` 起一块，其后每个非空行是一段新正文。"""
    blocks = []
    for path in paths:
        cur = None
        for line in Path(path).read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if not s:
                continue
            if s.startswith("@after "):
                cur = (s[7:].strip(), [])
                blocks.append(cur)
            elif cur is None:
                sys.exit(f"⛔ {path}: 首个 @after 之前有正文：{s[:30]}")
            else:
                cur[1].append(s)
    return blocks


def _clone_para(model, text):
    import copy
    new = copy.deepcopy(model)
    for k in list(new):
        if k.tag != w("pPr"):
            new.remove(k)
    for attr in list(new.attrib):                       # paraId 等须唯一，克隆件不带
        del new.attrib[attr]
    r = etree.SubElement(new, w("r"))
    rpr = model.find("w:r/w:rPr", NS)
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, w("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return new


def cmd_insert(a):
    """用户定稿后的扩充：在锚点段之后插入多段新正文，原有段落一字不动、顺序不变。

    Why（2026-09-18 海宁 6.2.3 扩充）：用户在 Word 定稿后要求“再扩充”，整章重生成会冲掉手改；
    按段号定位又会在用户再改一次后漂移。这里按锚点段原文定位（须唯一命中），新段克隆锚点
    的段落/首 run 格式——锚点是标题时改克隆其后第一段正文；首段新文已存在则整块跳过。
    """
    doc = Doc(a.docx)
    before = invariants(doc, content=False)
    old_texts = [ptext(p) for p in doc.body.iter(w("p"))]
    existing = {t.strip() for t in old_texts}
    body_ps = [k for k in doc.body if k.tag == w("p")]
    plan = []
    for anchor_text, lines in _parse_patch(a.patch):
        hits = [k for k in body_ps if ptext(k).strip().startswith(anchor_text)]
        if len(hits) != 1:
            sys.exit(f"⛔ 锚点「{anchor_text[:30]}」命中 {len(hits)} 处（应为 1）")
        anchor = hits[0]
        if lines and lines[0] in existing:
            print(f"  已存在「{lines[0][:20]}…」，跳过此块")
            continue
        model = anchor
        if doc.is_heading(anchor):
            model = next((k for k in anchor.itersiblings() if doc.is_text(k)), None)
            if model is None:
                sys.exit(f"⛔ 标题「{anchor_text[:30]}」之后找不到可克隆格式的正文段")
        plan.append((anchor, model, lines))
        print(f"  「{anchor_text[:24]}」之后插入 {len(lines)} 段")
    added = sum(len(x[2]) for x in plan)
    print(f"共插入 {added} 段")
    if not a.apply or not added:
        return
    last = {}
    for anchor, model, lines in plan:
        tail = last.get(id(anchor), anchor)             # 同锚点多块按补丁顺序接续
        for s in lines:
            el = _clone_para(model, s)
            tail.addnext(el)
            tail = el
        last[id(anchor)] = tail
    before["count"] += added
    before["tags"] = sorted(before["tags"] + [w("p")] * added)
    assert_invariants(doc, before)
    new_texts = iter(ptext(p) for p in doc.body.iter(w("p")))
    if not all(any(t == n for n in new_texts) for t in old_texts):   # 原段须按原顺序全部保留
        sys.exit("⛔ 原有段落缺失或顺序改变，未写回")
    doc.save()

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in [("check", cmd_check), ("filler", cmd_filler), ("h4num", cmd_h4num),
                     ("figs", cmd_figs), ("cite", cmd_cite), ("note", cmd_note),
                     ("h3group", cmd_h3group), ("uncite", cmd_uncite),
                     ("resub", cmd_resub), ("media", cmd_media), ("swapfig", cmd_swapfig), ("insert", cmd_insert)]:
        s = sub.add_parser(name)
        s.add_argument("docx")
        s.set_defaults(fn=fn)
        if name == "check":
            s.add_argument("--show", type=int, default=4)
        else:                                          # h4num 已退役，仍收 --apply 以便给出退役说明
            s.add_argument("--apply", action="store_true")
        if name == "media":
            s.add_argument("--old", required=True)
            s.add_argument("--new", required=True, action="append")
        if name == "swapfig":
            s.add_argument("--plan", required=True)
            s.add_argument("--width", type=float, default=15.5, help="显示宽度 cm（版心宽）")
            s.add_argument("--max-height", type=float, default=21.0, help="显示高度上限 cm")
        if name == "resub":
            s.add_argument("--rules", required=True)
            s.add_argument("--show", type=int, default=25)
        if name == "h3group":
            s.add_argument("--plan", required=True)
        if name == "insert":
            s.add_argument("--patch", required=True, action="append")
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
