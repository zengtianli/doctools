#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_qa.py — DOCX 交付件「校核终稿化」+「去AI味」QA 引擎(/docx finalize · /docx de-ai 后端)。

为什么单独成引擎：把「讨论稿→终稿」「去AI味」这两步从手搓 scratchpad 固化下来。
检测 100% 确定性(词表+正则+段号),改写人判(apply 执行人批准后的 ops.json)。

三动作：
  scan   <docx> --mode residue|aiflavor|both [--profile bid|report] [--json OUT] [--suggest]
         只读。扫草稿残留 / AI味,按类别打印命中(带段号)。--suggest 对"可机械删"的括注
         产出建议 ops.json(纯删除项),供 apply 前人工审。
  apply  <docx> <ops.json> [--no-backup]
         执行人批准的 ops：run-safe 段内替换(保格式) + 删表列。**每个文本 op 必须全文
         唯一命中,否则整体不保存(防 anchor 不符误伤)**。默认先备份 .bak-时间戳。
  verify <docx> [--profile bid] [--mode residue|aiflavor|both]
         重扫。residue 模式全类归零则 exit 0(终稿干净);aiflavor 只报数(留人判)。

run-safe 替换原理(沿用 docx_renumber_figures 同款,跨 run 安全)：段内 concat 全文 →
字符偏移定位 → 只改命中区间所在 w:t,区间外格式/run 一字不动。绝不用 python-docx(剥 OLE/媒体)。

ops.json 格式:
  {"text":   [{"old": "唯一原文", "new": "改后(删除则空串)"}, ...],
   "delcol": [{"table_header": "对应缺口"}, ...]}   # 删表头含该词的整列
"""
import argparse, json, os, re, shutil, sys, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
def w(t): return "{%s}%s" % (W, t)

# ── 词表 profile ──────────────────────────────────────────────
# 草稿残留(bid profile)：讨论稿才有、终稿必须删/改的"内部脚手架"
RESIDUE_BID = [
    ("内部缺口码",   r"缺口\s*G\s*[1-7]"),
    ("裸缺口码",     r"(?<![0-9A-Za-z])G[1-7](?![0-9A-Za-z])"),
    ("专题内部码",   r"专题0[1-9]"),
    ("待核实标签",   r"待区级核实|待核实|待台账"),
    ("待补标签",     r"待补[（(]"),
    ("评分痕迹",     r"响应子项\s*\d+|评分要点|子项\s*\d+"),
    ("口径分级meta", r"口径分级标注|行末标"),
    ("草稿注记",     r"对齐数据红线|文字底稿"),
    ("定稿语气",     r"定稿前|定稿时|标书定稿|定稿以|定稿引用"),
]
# 报告类用更克制的子集(没有评分/专题/缺口码这些标书专有物)
RESIDUE_REPORT = [
    ("待核实标签",   r"待核实|待台账|待补充|待复核"),
    ("草稿注记",     r"〔待|【待|XX(?![0-9])|××|待定|TODO|TBD"),
    ("定稿语气",     r"定稿前|定稿时|初稿|草稿阶段"),
]
# 可机械删的纯括注(--suggest 只对这些产删除建议;带"："的标签前缀需人改,不建议)
SUGGEST_DELETE = [
    r"（缺口G[1-7][^）]*）", r"（响应子项\d+）", r"（对齐数据红线）",
    r"（行末标[^）]*）", r"（缺口G[1-7]，定稿时核实）",
]

# AI味(通用 profile)：套话/口号/排比/句式模板
AI_BUZZWORDS = ["闭环","赋能","抓手","锚定","全链条","保驾护航","夯实","筑牢","谱写",
                "新篇章","新画卷","顶层设计","靶向","全周期","一盘棋","有力支撑","保障有力"]
AI_SELFPRAISE = ["内容齐全","结构完整","表达准确","条理清晰","逻辑严谨","层次分明",
                 "全面符合","成效显著","亮点纷呈","科学合理","切实可行"]

def load(docx):
    zin = zipfile.ZipFile(docx)
    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    root = etree.fromstring(data["word/document.xml"])
    return data, root

def paras_text(root):
    out = []
    for p in root.iter(w("p")):
        out.append("".join(t.text or "" for t in p.iter(w("t"))))
    return out

def full_text(root):
    return "".join(t.text or "" for t in root.iter(w("t")))

# ── run-safe 段内替换(跨 run,保格式) ──────────────────────────
def replace_in_para(p, old, new):
    items = []
    for r in p.iter(w("r")):
        for t in r.findall(w("t")):
            items.append([t, t.text or ""])
    full = "".join(it[1] for it in items)
    idx = full.find(old)
    if idx < 0:
        return 0
    a, b = idx, idx + len(old); pos = 0; done = False
    for t, txt in items:
        L = len(txt); s, e = pos, pos + L; pos = e
        if e <= a or s >= b:
            continue
        left = txt[:a-s] if s < a else ""
        right = txt[b-s:] if e > b else ""
        if not done:
            t.text = left + new + right; t.set(XML_SPACE, "preserve"); done = True
        else:
            t.text = left + right
    return 1

def tctext(tc): return "".join(t.text or "" for t in tc.iter(w("t")))

def delete_col_by_header(root, header_kw):
    body = root.find(w("body"))
    for tbl in body.iter(w("tbl")):
        rows = tbl.findall(w("tr"))
        if not rows:
            continue
        hdr = [tctext(tc).strip() for tc in rows[0].findall(w("tc"))]
        col = next((i for i, c in enumerate(hdr) if header_kw in c), None)
        if col is None:
            continue
        ncell = len(hdr)
        for tr in rows:                       # 列删除要求各行 tc 数一致(无横向合并错位)
            if len(tr.findall(w("tc"))) != ncell:
                return ("IRREGULAR", header_kw, ncell)
        grid = tbl.find(w("tblGrid"))
        if grid is not None:
            gcols = grid.findall(w("gridCol"))
            if col < len(gcols):
                grid.remove(gcols[col])
        for tr in rows:
            tr.remove(tr.findall(w("tc"))[col])
        return ("OK", header_kw, col)
    return ("NOTFOUND", header_kw, None)

# ── scan ─────────────────────────────────────────────────────
def scan(docx, mode, profile, jsonout, suggest):
    _, root = load(docx)
    paras = paras_text(root)
    full = "".join(paras)
    print("文件 %s" % os.path.basename(docx))
    print("段落 %d  字数 %d\n" % (len(paras), len(full)))
    total = 0

    if mode in ("residue", "both"):
        table = RESIDUE_BID if profile == "bid" else RESIDUE_REPORT
        print("=== 草稿残留 (profile=%s) ===" % profile)
        for cat, pat in table:
            rx = re.compile(pat)
            hits = []
            for i, p in enumerate(paras):
                for m in rx.finditer(p):
                    hits.append((i, m.group(0)))
            total += len(hits)
            mark = "" if not hits else "  ← 终稿应删/改"
            print("  %-12s %d%s" % (cat, len(hits), mark))
            for i, s in hits[:6]:
                ctx = paras[i].strip()
                print("      #%-4d %s" % (i, (ctx[:78] + "…") if len(ctx) > 78 else ctx))
            if len(hits) > 6:
                print("      … 另 %d 处" % (len(hits) - 6))
        print()

    if mode in ("aiflavor", "both"):
        print("=== AI味 ===")
        # 1 套话词频
        print("  [套话词频 ≥3]")
        for kw in AI_BUZZWORDS:
            c = full.count(kw)
            if kw == "全周期":   # 表格单元格里的"全周期"(短cell)不算口头禅
                c = sum(1 for p in paras if "全周期" in p and len(p) > 15)
            if c >= 3:
                print("      %-10s %d" % (kw, c))
        # 2 引号顺口溜(X得Y、X得Z 三连或带引号的口语四字)
        slogan = []
        for i, p in enumerate(paras):
            if len(re.findall(r"[“\"][^”\"，。；]{2,8}得[^”\"，。；]{1,8}[”\"]", p)) >= 2:
                slogan.append(i)
        if slogan:
            print("  [引号顺口溜段] %s" % slogan)
            for i in slogan[:4]:
                print("      #%-4d %s" % (i, paras[i].strip()[:78]))
        # 3 自评四连
        praise = []
        for i, p in enumerate(paras):
            if sum(1 for kw in AI_SELFPRAISE if kw in p) >= 3:
                praise.append(i)
        if praise:
            print("  [自评四连段] %s" % praise)
            for i in praise[:4]:
                print("      #%-4d %s" % (i, paras[i].strip()[:78]))
        # 4 口号排比(结构性重复才算;排除纯领域名词枚举如"防洪、治涝、供水")
        # 判据 = 单段 ≥3 个 "X到位/X有力/.." 评价四字 或 "X有Y/X得Y" 短句并列
        SLOGAN_TAILS = ("到位|有力|高效|畅通|内控|兜底|落地|清晰|完整|齐全|规范|可控|"
                        "有序|扎实|严谨|分明|纷呈|显著|可行|务实|留痕|销项")
        pat_tail = re.compile(r"[^，。、；：“”\"]{1,4}(?:" + SLOGAN_TAILS + r")[、，]")
        pat_youde = re.compile(r"[^，。、；：“”\"]{1,3}[有得][^，。、；：“”\"]{1,3}[、，]")
        pailie = []
        for i, p in enumerate(paras):
            n = len(pat_tail.findall(p)) + len(pat_youde.findall(p))
            if n >= 3:
                pailie.append(i)
        if pailie:
            print("  [口号排比段(结构性四字并列)] %s" % pailie)
            for i in pailie[:6]:
                print("      #%-4d %s" % (i, paras[i].strip()[:78]))
        # 5 句式模板
        for label, pat in [("通过…实现/确保", r"通过[^，。；]{2,30}[，、][^，。；]{0,30}(实现|确保|形成)"),
                           ("不仅…而且/还", r"不仅[^，。；]{2,40}(而且|还|更)")]:
            n = len(re.findall(pat, full))
            if n:
                print("  [%s] %d 处" % (label, n))
        print("\n  注：套话/排比/数字口号部分是标书体裁惯例,评审常受用,删需人判(非全删)。")
        print()

    if suggest and mode in ("residue", "both"):
        ops = {"text": [], "delcol": []}
        seen = set()
        for pat in SUGGEST_DELETE:
            for m in re.finditer(pat, full):
                s = m.group(0)
                if s not in seen:
                    seen.add(s); ops["text"].append({"old": s, "new": ""})
        outp = jsonout or (os.path.splitext(docx)[0] + ".qa_suggest.json")
        with open(outp, "w") as f:
            json.dump(ops, f, ensure_ascii=False, indent=2)
        print("=== --suggest: 可机械删的纯括注 %d 项 → %s ===" % (len(ops["text"]), outp))
        print("  (仅纯删除项;带标签前缀/需改写的须人工补 ops,审后 apply)\n")

    if mode in ("residue", "both"):
        print("草稿残留合计: %d  %s" % (total, "✅ 干净(疑似终稿)" if total == 0 else "← 见上"))
    return total

# ── apply ────────────────────────────────────────────────────
def apply(docx, opsfile, no_backup):
    with open(opsfile) as f:
        ops = json.load(f)
    data, root = load(docx)
    text_ops = ops.get("text", [])
    delcols = ops.get("delcol", [])

    print("=== 文本 op 命中校验(必须各=1) ===")
    hits = []
    for op in text_ops:
        old, new = op["old"], op.get("new", "")
        h = sum(replace_in_para(p, old, new) for p in root.iter(w("p")))
        hits.append(h)
        flag = "OK" if h == 1 else "XXXX 命中%d ≠1" % h
        print("  %d  %s  | %s" % (h, old[:30], flag))
    if any(h != 1 for h in hits):
        print("\n⚠️ 有 op 未唯一命中 —— 整体不保存,请核对 anchor。")
        return 2

    print("\n=== 删表列 ===") if delcols else None
    for dc in delcols:
        res = delete_col_by_header(root, dc["table_header"])
        print("  %s -> %s" % (dc["table_header"], res))
        if res[0] in ("IRREGULAR", "NOTFOUND"):
            print("\n⚠️ 删列异常(%s) —— 不保存。" % res[0])
            return 3

    if not no_backup:
        bak = docx + ".bak-" + os.popen("date +%Y%m%d-%H%M%S").read().strip()
        shutil.copy(docx, bak)
        print("\n备份 %s" % bak)
    data["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = docx + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for nm, d in data.items():
        zout.writestr(nm, d)
    zout.close(); os.replace(tmp, docx)
    print("✅ 已保存(%d 文本 op + %d 删列)" % (len(text_ops), len(delcols)))
    return 0

# ── verify ───────────────────────────────────────────────────
def verify(docx, profile, mode):
    total = scan(docx, mode, profile, None, False)
    if mode == "aiflavor":
        print("\n(aiflavor 只报数,不判 PASS/FAIL — 改与不改是体裁判断)")
        return 0
    ok = (total == 0)
    print("\n总判: %s" % ("PASS ✅ 终稿干净" if ok else "FAIL ❌ 仍有草稿残留"))
    return 0 if ok else 1

def main():
    ap = argparse.ArgumentParser(description="DOCX 终稿化/去AI味 QA 引擎")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sp = sub.add_parser("scan"); sp.add_argument("docx")
    sp.add_argument("--mode", choices=["residue","aiflavor","both"], default="both")
    sp.add_argument("--profile", choices=["bid","report"], default="bid")
    sp.add_argument("--json"); sp.add_argument("--suggest", action="store_true")
    ap2 = sub.add_parser("apply"); ap2.add_argument("docx"); ap2.add_argument("ops")
    ap2.add_argument("--no-backup", action="store_true")
    vp = sub.add_parser("verify"); vp.add_argument("docx")
    vp.add_argument("--profile", choices=["bid","report"], default="bid")
    vp.add_argument("--mode", choices=["residue","aiflavor","both"], default="residue")
    a = ap.parse_args()
    if a.cmd == "scan":
        scan(a.docx, a.mode, a.profile, a.json, a.suggest); sys.exit(0)
    if a.cmd == "apply":
        sys.exit(apply(a.docx, a.ops, a.no_backup))
    if a.cmd == "verify":
        sys.exit(verify(a.docx, a.profile, a.mode))

if __name__ == "__main__":
    main()
