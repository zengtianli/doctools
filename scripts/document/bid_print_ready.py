#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_print_ready.py — 标书终稿「打印级」只读验收门（doctools 家族）。

定位：终稿 = 直接可打印。本引擎**只读不写**，扫「打印出来会露馅」的硬伤：
  fatal ① ASCII 字符画字符（┌┐└┘…═║）出现在正文
        ② 图/表占位标记残留（〔图位 / 〔界面图位 / 段首【图 / 段首 fenced ``` hint）
        ③ 题注编号断裂：每类前缀（图 X- / 表 X-，X=章号）节内编号须连续 1..k 且与出现序一致
           （只认「紧邻图/表」的真题注段，天然排除图目录/表目录/正文引用行）
        ④ zip 损坏 / word/document.xml 不可解析
        ⑤ 空段落孤立标点（段文本 strip 后仅剩 ，。、）； —— regex 清理二次残渣兜底）
  warn  ① docProps creator/lastModifiedBy 非空（身份门管 fatal，这里仅提示）
        ② media 数为 0（纯文字标书可能合法）
        ③ 无 TOC（无目录域/目录段）
        ④ 页面非 A4（横向 A4 视为合法，宽表横向属正常装帧）

CLI 契约（终稿管线统一）：
  python3 bid_print_ready.py <docx> [--mode main|pei] [--rules <yaml>]
  exit 0 = PASS · exit 2 = 有 fatal findings · exit 1 = 用法/IO 错误
  stdout 最后一行必为 "PASS" 或 "FAIL <n> findings"（n = fatal 数；仅 warn → PASS）
  --mode 默认 pei（本门为只读校验，两模式检查项相同，仅记录在报告头）
  --rules 项目级 YAML 增补：identity_banned 命中在本门降级为 warn 提示（fatal 归身份门）

题注校验参考 shaoxing normalize_captions.py（相邻图/表判定）与 finalize2.py（重编号段）。
只用 stdlib + lxml；零写操作，无备份需求。
"""
import argparse
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DC = "http://purl.org/dc/elements/1.1/"
CP = ("http://schemas.openxmlformats.org/package/2006/metadata/core-properties")


def w(t):
    return "{%s}%s" % (W, t)


def ln(e):
    return etree.QName(e).localname


# ── fatal ① ASCII 字符画字符集 ──────────────────────────────────
ASCII_ART_CHARS = set("┌┐└┘├┤┬┴┼╪╞╡╫╬═║╔╗╚╝╠╣╦╩")

# ── fatal ② 占位标记 ────────────────────────────────────────────
PLACEHOLDER_CONTAINS = ["〔图位", "〔界面图位"]          # 段内任意位置命中即 fatal
PLACEHOLDER_PREFIX_RE = [
    (re.compile(r"^【图"), "段首【图…占位题注"),
    (re.compile(r"^```"), "fenced hint 首行 ``` 残留"),
]

# ── fatal ③ 题注（真题注 = 紧邻图/表）────────────────────────────
CAP_RE = re.compile(r"^(图|表)\s*([0-9]+(?:[.．][0-9]+)*)\s*[-‑–—]\s*([0-9]+)(?=[　\s）】]|$)")

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

    groups = {}  # (kind, chapter) -> [(num, P段号, 题注文本)]
    fig_caps = 0
    for i, e in enumerate(kids):
        if ln(e) != "p":
            continue
        st = ptext(e).strip()
        m = CAP_RE.match(st)
        if not m:
            continue
        kind, chap, num = m.group(1), m.group(2), int(m.group(3))
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
        groups.setdefault((kind, chap), []).append((num, pno[id(e)], st[:30]))
    for (kind, chap), items in sorted(groups.items()):
        seq = [n for n, _, _ in items]
        want = list(range(1, len(seq) + 1))
        if seq != want:
            where = " ".join(f"{kind}{chap}-{n}@P{pn}" for n, pn, _ in items)
            fatal.append(f"题注编号断裂: 「{kind} {chap}-」出现序 {seq} ≠ 期望 {want} · {where}")

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


def main():
    ap = argparse.ArgumentParser(description="标书终稿打印级只读验收门")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--mode", choices=["main", "pei"], default="pei")
    ap.add_argument("--rules", type=Path, default=None)
    a = ap.parse_args()

    if not a.docx.exists():
        print(f"找不到文件: {a.docx}", file=sys.stderr)
        sys.exit(1)
    rules = {}
    if a.rules is not None:
        if not a.rules.exists():
            print(f"找不到规则文件: {a.rules}", file=sys.stderr)
            sys.exit(1)
        try:
            rules = parse_simple_yaml(a.rules)
        except Exception as e:
            print(f"规则 YAML 解析失败: {e}", file=sys.stderr)
            sys.exit(1)

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
        sys.exit(2)
    print(f"小结: fatal 0 · warn {len(warn)}")
    print("PASS")
    sys.exit(0)


if __name__ == "__main__":
    main()
