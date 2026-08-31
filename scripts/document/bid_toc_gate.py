#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_toc_gate.py — 标书目录一致性门：正文章号 ≡ chapters.yaml ≡ 招标文件那张规定顺序的表。

Why（2026-08-30 用户钦定 /govern，桐乡实证）：
  标书目录不是作者自由发挥的地方。评委拿招标文件那张表逐项翻分，翻不到就扣分。
  桐乡原稿自造了 13 章（「项目理解」「总体思路」各占一章），于是「安全保证 = 评分第 6 项」
  写在第 8 章，其后全部错位 2。用户原话：「就一模一样，按照招标要求的序号、分数条目来，
  一点都不能改」。这类错所有既有门（残留/身份/付印）全绿也抓不到——它们不看目录。

判据（顺序固定，别自己发挥）：
  ① 招标文件有「投标文件组成 / 装订顺序」表 → 整册装订顺序跟它；
  ② 招标文件有「评标办法」评分表        → 正文章号跟它（章名、分值照抄原文）；
  ③ 两张都有 → 两张都要对上（嵌套：装订表管排列，评分表管正文章号）；
  ④ 一张都没有 → 才允许自拟，且必须先问用户。
  招标表里没有的内容不占章号（放不编号卷首/附件）。

契约：
  exit 0 = 全绿 · 1 = 用法/IO 错误 · 2 = 有不一致 / 无 chapters.yaml / 枚举为空。
  「枚举为空一律非零」——拒绝在空集上报绿。

用法：
  bid_toc_gate.py <项目目录|技术标docx>  [--chapters <chapters.yaml>]

chapters.yaml 两种 schema 都认：
  A) chapters: [{item_no, title, score, file}]           ← 章号 = 招标表项号（桐乡式）
  B) number_base: N + sequence: [{slug,title,subs?}]     ← 章号 = number_base + 序位（温州式）
豁免：项目 _project.yaml 写 `toc_gate: legacy   # <理由>` → 打印豁免行并 exit 0，不静默跳过。
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

import yaml
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CN = "零一二三四五六七八九十"
CN_NUM = {c: i for i, c in enumerate(CN)}


def cn2int(s):
    """一 / 十 / 十一 / 二十三 → int"""
    if not s:
        return None
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + CN_NUM.get(s[1:], 0)
    if "十" in s:
        a, b = s.split("十")
        return CN_NUM.get(a, 0) * 10 + (CN_NUM.get(b, 0) if b else 0)
    return CN_NUM.get(s)


def int2cn(n):
    if n <= 10:
        return CN[n]
    if n < 20:
        return "十" + CN[n - 10]
    return CN[n // 10] + "十" + (CN[n % 10] if n % 10 else "")


# ── 输入定位 ────────────────────────────────────────────────────────────────
def find_project(p: Path):
    for d in [p if p.is_dir() else p.parent, *(p if p.is_dir() else p.parent).parents]:
        if (d / "_project.yaml").is_file():
            return d
    return None


def find_chapters_yaml(root: Path):
    for rel in ("chapters.yaml", "技术标/chapters.yaml", "scripts/chapters.yaml",
                "成果/chapters.yaml"):
        f = root / rel
        if f.is_file():
            return f
    return None


# ── chapters.yaml → 期望章表 ────────────────────────────────────────────────
def expected_from_yaml(y: dict):
    """→ ([(章键str, 章名, 分值or None)], schema名)。
    章键 = 显示章号的字符串形式："6" / "12.3"，两种 schema 与两种标题写法共用同一比较键。"""
    if isinstance(y.get("chapters"), list) and y["chapters"]:
        out = []
        for c in y["chapters"]:
            no = c.get("item_no", c.get("no"))
            out.append((str(int(no)), str(c["title"]), c.get("score")))
        return out, "chapters"
    if isinstance(y.get("sequence"), list) and y["sequence"]:
        base = int(y.get("number_base", 1))
        out = []
        for i, s in enumerate(y["sequence"]):
            n = base + i
            subs = s.get("subs") or []
            if subs:
                for k, sb in enumerate(subs, 1):
                    out.append((f"{n}.{k}", str(sb["title"]), None))
            else:
                out.append((str(n), str(s["title"]), None))
        return out, "sequence"
    return [], "?"


# ── 正文实际章标题 ──────────────────────────────────────────────────────────
# 两种合法写法：「第六章 安全保证（5分）」与「12.3 流域生态流量调度方案」——
# 章号写法是项目风格，不是判据；判据是章号序与章名。
HEAD_CN = re.compile(r"^第([一二三四五六七八九十]+)章\s*(.+?)(?:（(\d+)分）)?\s*$")
HEAD_AR = re.compile(r"^(\d+(?:\.\d+)*)[\s、.]\s*(.+?)(?:（(\d+)分）)?\s*$")


def parse_head(txt):
    """→ (章键str, 章名, 分值or None) 或 None"""
    m = HEAD_CN.match(txt)
    if m:
        return str(cn2int(m.group(1))), m.group(2).strip(), int(m.group(3)) if m.group(3) else None
    m = HEAD_AR.match(txt)
    if m:
        return m.group(1), m.group(2).strip(), int(m.group(3)) if m.group(3) else None
    return None


def heads_from_md(paths):
    got = []
    for p in paths:
        for line in p.read_text(encoding="utf-8").split("\n"):
            if not line.startswith("# "):
                continue
            h = parse_head(line[2:].strip())
            if h:
                got.append((*h, p.name))
    return got


def heads_from_docx(docx: Path):
    x = etree.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    got = []
    for para in x.iter("{%s}p" % W):
        st = para.find(".//{%s}pStyle" % W)
        sid = st.get("{%s}val" % W) if st is not None else ""
        if sid not in ("Heading1", "1", "a3", "标题1"):
            continue
        txt = "".join(tt.text or "" for tt in para.iter("{%s}t" % W)).strip()
        h = parse_head(txt)
        if h:
            got.append((*h, docx.name))
    return got


# ── 招标一手源 ──────────────────────────────────────────────────────────────
def tender_scoring_items(root: Path):
    """招标原件 md 里「评标办法」评分表 → [(no, name, score)]，解析不到返回 []。"""
    mds = sorted((root / "招标文件").glob("*.md")) if (root / "招标文件").is_dir() else []
    for md in mds:
        s = md.read_text(encoding="utf-8", errors="ignore")
        raw = re.findall(r"\|\s*([1-9]\d?)、\s*([^|\n]{2,40}?)\s*\|", s)
        seen, items = set(), []
        for no, name in raw:
            no = int(no)
            if no in seen:
                continue
            seen.add(no)
            name = re.sub(r"\[|\]\{\.s\d+\}", "", name).strip()
            m = re.match(r"^(.*?)（(\d+)分）$", name)
            items.append((no, (m.group(1) if m else name).strip(),
                          int(m.group(2)) if m else None))
        items = [i for i in items if i[1]]
        if len(items) >= 5 and [i[0] for i in items] == list(range(1, len(items) + 1)):
            return items, md.name
    return [], None


# ── 招标「形式硬要求」机检 ─────────────────────────────────────────────────
# Why（2026-08-30 用户钦定「招标要求的重要性是最高的」）：招标「投标文件的编制」一节里
# 页码/封面/目录这类要求白纸黑字写着「导致被误读、漏读或者查找不到相关内容的，是投标人的
# 责任」，却因为不在评分表里而整轮没人列过。内容对齐由上面的章号比对兜，形式要求由这里兜。
def check_form_requirements(docx: Path, reqs: list):
    """reqs = chapters.yaml 的 form_requirements。→ (findings, checked_n)"""
    z = zipfile.ZipFile(docx)
    names = z.namelist()
    doc = z.read("word/document.xml").decode("utf-8", "ignore")
    ftr = "".join(z.read(n).decode("utf-8", "ignore") for n in names if "footer" in n)
    hdr = "".join(z.read(n).decode("utf-8", "ignore") for n in names if "header" in n)
    body_txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", doc))
    front = body_txt[:1200]          # 封面只可能在最前面

    def has_page_number():
        blob = ftr + hdr + doc
        return bool(re.search(r"\bPAGE\b", blob))

    def has_toc():
        return ("TOC \\o" in doc) or ("TOC \\h" in doc) or ("目　　录" in front) or ("目录" in front)

    AUTO = {
        "页码": (has_page_number, "页眉页脚里没有 PAGE 域 —— 全文无页码"),
        "目录": (has_toc, "没有目录域，卷首也没有「目录」段"),
    }
    findings, n = [], 0
    for r in reqs:
        name = str(r.get("要求") or r.get("item") or "").strip()
        src = str(r.get("条款") or r.get("clause") or "").strip()
        n += 1
        by = str(r.get("by") or "").strip()
        if by and by not in ("投标人", "本稿"):
            continue
        if name in AUTO:
            fn, msg = AUTO[name]
            if not fn():
                findings.append(f"[形式要求] {name}（招标 {src}）：{msg}")
            continue
        kws = r.get("必含关键词") or r.get("keywords") or []
        if kws:
            miss = [k for k in kws if k not in front]
            if miss:
                findings.append(f"[形式要求] {name}（招标 {src}）：卷首缺 {'、'.join(miss)}")
            continue
        by = str(r.get("by") or "").strip()
        if by and by not in ("投标人", "本稿"):
            continue  # 责任方不是本稿（如统稿单位统一编页），已在 chapters.yaml 显式声明
        if not str(r.get("location") or "").strip():
            findings.append(f"[形式要求] {name}（招标 {src}）：既无机检也没写 location —— 无法证明已落位")
    return findings, n


def main(argv=None):
    ap = argparse.ArgumentParser(prog="bid_toc_gate.py", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target", type=Path, help="项目目录 或 技术标 docx")
    ap.add_argument("--chapters", type=Path, default=None)
    a = ap.parse_args(argv)

    a.target = a.target.resolve()
    if not a.target.exists():
        print(f"错误: 不存在 {a.target}", file=sys.stderr)
        return 1
    root = find_project(a.target)
    if root is None:
        print(f"错误: 从 {a.target} 向上找不到 _project.yaml，无法定位标书项目根", file=sys.stderr)
        return 1

    print(f"== bid_toc_gate · {root.name} ==")

    # 豁免（必须显式写，且带理由）
    proj = yaml.safe_load((root / "_project.yaml").read_text(encoding="utf-8")) or {}
    if str(proj.get("toc_gate", "")).strip().startswith("legacy"):
        print(f"[豁免] _project.yaml toc_gate: {proj['toc_gate']}")
        print("PASS（豁免，非核验通过；再版前必须建 chapters.yaml）")
        return 0

    cy = a.chapters or find_chapters_yaml(root)
    if cy is None:
        print("[FAIL] 找不到 chapters.yaml —— 章号没有 SSOT，等于没有依据")
        print("       建它：从招标文件那张规定顺序的表派生（装订顺序表 → 评分表 → 两张都有就都对上）")
        print("       确属规则之前的历史稿 → _project.yaml 写 `toc_gate: legacy  # 理由`")
        return 2
    y = yaml.safe_load(cy.read_text(encoding="utf-8")) or {}
    exp, schema = expected_from_yaml(y)
    if not exp:
        print(f"[FAIL] {cy} 里 chapters/sequence 枚举为空 —— 拒绝在空集上报绿")
        return 2
    print(f"chapters.yaml: {cy.relative_to(root)} · schema={schema} · {len(exp)} 章")

    err = []

    # ① chapters.yaml ≡ 招标评分表（一手源）
    items, src = tender_scoring_items(root)
    if items:
        print(f"招标评分表一手源: 招标文件/{src} · {len(items)} 项")
        if len(items) == len(exp):
            for (no, name, score), (eno, etitle, escore) in zip(items, exp):
                if str(no) != eno:
                    err.append(f"章号 {eno} ≠ 评分项 {no}")
                elif name[:4] != etitle[:4]:
                    err.append(f"第 {eno} 章 章名「{etitle}」≠ 评分表「{name}」")
                elif score is not None and escore is not None and score != escore:
                    err.append(f"第 {eno} 章 分值 {escore} ≠ 评分表 {score}")
        else:
            print(f"[note] 章数 {len(exp)} ≠ 评分项数 {len(items)}"
                  f" —— 本项目章号可能跟的是「投标文件组成」装订表，跳过评分表逐项比对")
    else:
        print("[note] 未从招标原件 md 解析出评分表（可能未转 md），跳过一手源比对")

    # ② chapters.yaml ≡ scoring.json
    sj = root / "招标文件/scoring.json"
    if sj.is_file():
        tb = json.loads(sj.read_text(encoding="utf-8")).get("tech_business", {})
        si = tb.get("items", [])
        if si and len(si) == len(exp):
            for (eno, etitle, escore), j in zip(exp, si):
                if eno != str(j["no"]) or etitle != j["name"]:
                    err.append(f"chapters.yaml 第 {eno} 章「{etitle}」≠ scoring.json 项{j['no']}「{j['name']}」")
                elif escore is not None and escore != j["max"]:
                    err.append(f"第 {eno} 章 分值 {escore} ≠ scoring.json {j['max']}")
        if si and tb.get("score") and sum(i["max"] for i in si) != tb["score"]:
            err.append(f"scoring.json 分值合计 {sum(i['max'] for i in si)} ≠ {tb['score']}")

    # ③ chapters.yaml ≡ 正文实际标题
    got, where = [], None
    if a.target.is_file() and a.target.suffix.lower() == ".docx":
        got, where = heads_from_docx(a.target), a.target.name
    else:
        files = [root / c["file"] for c in y.get("chapters", []) if c.get("file")]
        files = [f for f in files if f.is_file()]
        if not files:
            for d in ("成果/md", "技术标/chapters", "成果/chapters", "成果"):
                cand = sorted((root / d).glob("ch*.md")) if (root / d).is_dir() else []
                if cand:
                    files, where = cand, d
                    break
        else:
            where = "chapters.yaml 声明的章文件"
        got = heads_from_md(files)
    if not got:
        print(f"[FAIL] 正文里一条「第X章 …」都没扫到（{where or a.target}）—— 拒绝在空集上报绿")
        return 2
    print(f"正文章标题来源: {where} · 扫到 {len(got)} 章")

    # delivery_scope：交付件只含部分章（陪标分册/技术线通用稿）时，必须**显式声明**范围，
    # 声明后只校验范围内的章。没声明就按全量比 —— 不给"少几章"留默认放过的口子。
    scope = ((y.get("delivery_scope") or {}).get("chapters")) or []
    if scope:
        exp = [e for e in exp if e[0] in {str(s) for s in scope}]
        print(f"交付范围: 第 {'、'.join(str(s) for s in scope)} 章"
              f"（{(y.get('delivery_scope') or {}).get('reason', '')}）")
    if len(got) != len(exp):
        err.append(f"正文 {len(got)} 章 ≠ 应交付 {len(exp)} 章："
                   f"正文={[g[0] for g in got]} 期望={[e[0] for e in exp]}")
    else:
        for (gno, gtitle, gscore, gsrc), (eno, etitle, escore) in zip(got, exp):
            if gno != eno:
                err.append(f"{gsrc}: 章号 {gno} 应为 {eno}")
            elif gtitle != etitle:
                err.append(f"{gsrc}: 第 {eno} 章 正文「{gtitle}」≠ chapters.yaml「{etitle}」")
            elif escore is not None and gscore is not None and gscore != escore:
                err.append(f"{gsrc}: 第 {eno} 章 分值（{gscore}分）≠ 应为（{escore}分）")

    # ④ 装订顺序表（招标「投标文件组成」）逐项落位
    bo = y.get("binding_order") or []
    if bo:
        miss = [b for b in bo if not str(b.get("location", "")).strip()]
        if miss:
            err.append("binding_order 有 %d 项没写 location（招标装订表要求的件没落位）: %s"
                       % (len(miss), "、".join(str(b.get("no", "?")) for b in miss)))
        print(f"装订顺序表: {len(bo)} 项 · 已落位 {len(bo) - len(miss)}")

    # ⑤ 招标形式硬要求（页码/封面/目录/正副本字样…），仅在 chapters.yaml 声明时启用
    reqs = y.get("form_requirements") or []
    if reqs:
        if a.target.is_file() and a.target.suffix.lower() == ".docx":
            fe, fn_ = check_form_requirements(a.target, reqs)
            err.extend(fe)
            print(f"招标形式硬要求: {fn_} 条 · 不符 {len(fe)}")
        else:
            print(f"招标形式硬要求: {len(reqs)} 条（只对 docx 生效，本次目标是目录，跳过）")

    if err:
        print("[FAIL] 目录一致性门")
        for e in err:
            print("  ✗", e)
        print(f"FAIL {len(err)} findings")
        return 2
    print(f"[PASS] 招标表 ≡ chapters.yaml ≡ 正文 · {len(exp)} 章"
          + (f" / {sum(e[2] for e in exp if e[2])} 分" if all(e[2] for e in exp) else ""))
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
