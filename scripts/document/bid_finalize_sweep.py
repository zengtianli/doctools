#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_finalize_sweep.py — 标书终稿确定性清理引擎（残留类 1-7）。

第8类身份泄漏不在此改（交 bid_identity_gate.py）。surgical：zipfile 逐 entry 复制，
只重写 word/document.xml；改前 .bak-YYYYmmdd-HHMMSS 备份。

内置行为：
  ① 剥〔E-xx...〕依据码（含括号内 worklib#/招标段号，整个〔..〕剥除；公文文号〔2025〕不动）
  ② rules delete_startswith / delete_exact 整段删（评分脚手架块/照抄评标办法裸句）
  ③ rules exact 跨 run 保格式精确替换
  ④ rules caption_renumber 题注按正文出现序幂等重编号（占位法防连环；已按序则跳过）
  ⑤ 通用二次残渣修（，此处）→）· ，本处）→）· ，）→）· 、）→）· （）→删）

用法: python3 bid_finalize_sweep.py <docx> [--mode main|pei] [--rules <yaml>] [--check|--apply]
  默认 --check 干跑打印计数不落盘；--apply 才写（备份 + 三护栏红则不写 exit 2）。
  落盘后自动复扫（bid_residue_lib.scan_parts 类别 1-7），复扫归零才 PASS。
exit 0 = PASS（无事可做/清理后归零）；exit 2 = 有待清理项/护栏红/复扫未归零；exit 1 = 用法/IO 错误。
"""
import argparse
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bid_residue_lib as lib

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


def main():
    ap = argparse.ArgumentParser(description="标书终稿确定性清理引擎（残留类 1-7）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--mode", choices=["main", "pei"], default="pei")
    ap.add_argument("--rules", type=Path, default=None)
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--check", action="store_true", help="干跑打印计数（默认）")
    g.add_argument("--apply", action="store_true", help="落盘（备份+三护栏机检）")
    args = ap.parse_args()

    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    try:
        rules = lib.load_rules(args.rules)
        names, parts = lib.load_parts(args.docx)
        if "word/document.xml" not in parts:
            print(f"错误: 不是有效 docx（缺 word/document.xml）: {args.docx}", file=sys.stderr)
            return 1
    except (OSError, ValueError, RuntimeError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    terms = lib.protect_terms(rules)
    root = lib.parse_document(parts)
    before = "".join(lib.ptext(p) for p in lib.paragraphs(root))
    g0 = lib.guards(before, terms)

    counts = sweep(root, rules)
    total = sum(counts.values())

    after = "".join(lib.ptext(p) for p in lib.paragraphs(root))
    g1 = lib.guards(after, terms)
    bad = lib.guard_diff(g0, g1)

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

    # --apply 落盘：备份 + 只重写 document.xml
    if total:
        bak = args.docx.with_suffix(args.docx.suffix + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
        shutil.copy2(str(args.docx), str(bak))
        parts["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(str(args.docx), "w", zipfile.ZIP_DEFLATED) as z:
            for n in names:
                z.writestr(n, parts[n])
        print(f"已写 {args.docx.name} · 备份 {bak.name}")
    else:
        print("无待清理项，未改动文件（幂等）")

    # 落盘后自动复扫（类别 1-7；第8类归 identity_gate）
    _, parts2 = lib.load_parts(args.docx)
    residual = lib.scan_parts(parts2, mode=args.mode, rules=rules, cats=range(1, 8))
    if residual:
        print(f"复扫（类1-7）未归零，剩 {len(residual)} 条（需人工改写或补 rules）:")
        for line in lib.format_findings(residual):
            print("  " + line)
        print(f"FAIL {len(residual)} findings")
        return 2
    print("复扫（类1-7）归零")
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
