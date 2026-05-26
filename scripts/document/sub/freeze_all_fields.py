#!/usr/bin/env python3
# distilled from qual-supply/scripts/freeze_all_fields.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
freeze_all_fields.py
=====================

单功能描述
----------
把 docx 内所有 Word 字段域 (`<w:fldChar>` / `<w:instrText>` / `<w:fldSimple>`)
**冻结为静态文本** (plain text): 保留字段渲染结果, 删除字段壳 (begin/separate/end
fldChar 所在 run + instrText 所在 run + fldSimple 包裹元素), 合稿时 Word 不会再
重算 TOC / PAGEREF / SEQ / 公式等字段。

触发场景
--------
- 多份 docx 合稿前, TOC + PAGEREF 字段会被 Word 重算导致目录页码全乱
- 主报告"准成品"阶段, 罗马数字页码 / SEQ 编号需冻结为字面文本
- 验收前预防性 freeze, 保证客户打开 Word 不触发 F9 update

CLI
---
    python3 scripts/freeze_all_fields.py <docx_path> \\
        [--types TOC,PAGEREF,SEQ,=,REF,STYLEREF,DATE,...] \\
        [--dry-run] [--no-backup] [--report <json>]

默认行为
--------
- `--types all` (冻结所有字段类型)
- 自动 `--backup` 写 `.bak-N-YYYY-MM-DD.docx` (除非 `--no-backup`)
- 写前 `lsof` 自检 Word/WPS 占用

算法
----
fldSimple (简单字段):
  <w:fldSimple w:instr="..."><w:r>...</w:r></w:fldSimple>
  → 把 fldSimple 内子元素提升到 parent 同位置, 删 fldSimple 包裹

fldChar + instrText (复杂字段, 可嵌套):
  begin run → instr runs → separate run → result runs → end run
  用栈匹配 begin/end (支持嵌套, TOC > PAGEREF > = 三层都见过)
  - 删 begin run (含 fldChar begin)
  - 删 instrText runs (begin 到 separate 之间)
  - 删 separate run (含 fldChar separate)
  - 删 end run (含 fldChar end)
  - **保留** separate 到 end 之间的 result runs (这是渲染文本)

类型过滤:
  --types TOC,PAGEREF 只 freeze 这两类字段
  通过 instr 首词 (TOC / PAGEREF / SEQ / = / REF / STYLEREF / DATE / TIME ...) 判断
  保留 (skip) 的字段维持原样

不许做
------
- 改非字段段的 plain text 内容
- 改 header/footer (本脚本只动 word/document.xml)
- 改 styles.xml / numbering.xml
- 用 sed/awk 改 XML 字符串
- commit / push
"""
import argparse
import datetime
import json
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def lsof_check(p: Path) -> None:
    try:
        r = subprocess.run(["lsof", str(p)], capture_output=True, text=True, timeout=5)
        if r.stdout.strip():
            print(f"[ERR] 文件被占用 (Word/WPS 没关?):\n{r.stdout}", file=sys.stderr)
            sys.exit(2)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass


def backup(p: Path) -> Path:
    today = datetime.date.today().isoformat()
    n = 1
    while True:
        b = p.with_name(f"{p.stem}.bak-{n}-{today}{p.suffix}")
        if not b.exists():
            shutil.copy2(p, b)
            return b
        n += 1


def find_enclosing_run(elem):
    """Walk up to the nearest <w:r> ancestor (or None)."""
    cur = elem
    while cur is not None and etree.QName(cur).localname != "r":
        cur = cur.getparent()
    return cur


def freeze_simple_fields(root, type_filter):
    """fldSimple → expand inner runs in place. type_filter is None or set of allowed types."""
    removed = []
    kept = []
    for fs in list(root.iter(f"{W}fldSimple")):
        instr = (fs.get(f"{W}instr") or "").strip()
        first = instr.split()[0] if instr else "EMPTY"
        if type_filter is not None and first not in type_filter:
            kept.append(first)
            continue
        parent = fs.getparent()
        if parent is None:
            continue
        idx = parent.index(fs)
        children = list(fs)
        for child in children:
            fs.remove(child)
            parent.insert(idx, child)
            idx += 1
        parent.remove(fs)
        removed.append(first)
    return removed, kept


def freeze_complex_fields(root, type_filter):
    """fldChar begin/separate/end matched via stack. instrText runs deleted.
    Result runs (separate..end) preserved.

    Returns (frozen_types: list[str], skipped_types: list[str],
             instrtext_runs_removed: int, fldchar_runs_removed: int)
    """
    # Build document order list of all fldChars and instrTexts (interleaved by doc order)
    # We need order info to know which instrTexts belong to which field.

    # Collect every fldChar with its enclosing run.
    # Iterate fldChars in document order.
    body = root.find(f".//{W}body")
    if body is None:
        body = root

    fldchars = []  # list of (elem, run, fldCharType)
    for fc in body.iter(f"{W}fldChar"):
        run = find_enclosing_run(fc)
        if run is None:
            continue
        ft = fc.get(f"{W}fldCharType")
        fldchars.append((fc, run, ft))

    # Build a quick index: for each instrText, find its enclosing run.
    instr_runs = []  # (run, text)
    for it in body.iter(f"{W}instrText"):
        run = find_enclosing_run(it)
        if run is None:
            continue
        instr_runs.append((run, it.text or ""))

    # Now walk body in document order, building field stack.
    # We need a flat traversal yielding either fldChar runs or instrText runs in doc order.
    # Easier: do a single iter and pick up both kinds.
    events = []  # (run, kind, payload)  kind: 'begin'|'separate'|'end'|'instr'
    seen_runs_for_event = set()
    for elem in body.iter():
        tag = etree.QName(elem).localname
        if tag == "fldChar":
            run = find_enclosing_run(elem)
            if run is None:
                continue
            ft = elem.get(f"{W}fldCharType")
            events.append((run, ft, None))
        elif tag == "instrText":
            run = find_enclosing_run(elem)
            if run is None:
                continue
            # If multiple instrText in same run (rare), we still want to record once per run
            events.append((run, "instr", elem.text or ""))

    # Stack-based pairing.
    # Each stack frame: {begin_run, separate_run, end_run, instr_text (concat), instr_runs (list of runs)}
    stack = []
    fields = []  # closed fields, in close order
    for run, kind, payload in events:
        if kind == "begin":
            stack.append({
                "begin_run": run,
                "separate_run": None,
                "end_run": None,
                "instr_text": "",
                "instr_runs": [],
                "post_separate": False,
            })
        elif kind == "instr":
            if stack and not stack[-1]["post_separate"]:
                stack[-1]["instr_text"] += payload
                if run not in stack[-1]["instr_runs"]:
                    stack[-1]["instr_runs"].append(run)
        elif kind == "separate":
            if stack:
                stack[-1]["separate_run"] = run
                stack[-1]["post_separate"] = True
        elif kind == "end":
            if stack:
                ctx = stack.pop()
                ctx["end_run"] = run
                fields.append(ctx)

    # Now decide which fields to freeze.
    runs_to_remove = []  # list of runs (preserve order, dedup later)
    frozen_types = []
    skipped_types = []
    for f in fields:
        instr = f["instr_text"].strip()
        first = instr.split()[0] if instr else "EMPTY"
        if type_filter is not None and first not in type_filter:
            skipped_types.append(first)
            continue
        frozen_types.append(first)
        # Mark runs for removal: begin, all instr runs, separate, end.
        # Result runs (between separate and end) are NOT in this list → preserved.
        if f["begin_run"] is not None:
            runs_to_remove.append(f["begin_run"])
        for r in f["instr_runs"]:
            runs_to_remove.append(r)
        if f["separate_run"] is not None:
            runs_to_remove.append(f["separate_run"])
        if f["end_run"] is not None:
            runs_to_remove.append(f["end_run"])

    # Dedup by id, preserve order
    seen = set()
    fldchar_run_count = 0
    instr_run_count = 0
    final_remove = []
    for r in runs_to_remove:
        rid = id(r)
        if rid in seen:
            continue
        seen.add(rid)
        # Classify for stats: if run contains fldChar → fldchar run; if contains instrText → instr run
        has_fc = r.find(f"{W}fldChar") is not None
        has_it = r.find(f"{W}instrText") is not None
        if has_fc:
            fldchar_run_count += 1
        elif has_it:
            instr_run_count += 1
        final_remove.append(r)

    # Execute removal
    for r in final_remove:
        parent = r.getparent()
        if parent is not None:
            parent.remove(r)

    return frozen_types, skipped_types, instr_run_count, fldchar_run_count


def parse_types(s: str):
    s = s.strip()
    if not s or s.lower() == "all":
        return None
    return set(t.strip() for t in s.split(",") if t.strip())


def main():
    ap = argparse.ArgumentParser(description="Freeze all Word fields to static text.")
    ap.add_argument("docx", help="Path to .docx")
    ap.add_argument("--types", default="all",
                    help="Comma-separated field types to freeze (e.g. TOC,PAGEREF,SEQ,=). "
                         "'all' = freeze every field. Default: all")
    ap.add_argument("--dry-run", action="store_true", help="Report only, do not modify file")
    ap.add_argument("--no-backup", action="store_true", help="Skip writing .bak-N-<date>.docx")
    ap.add_argument("--report", help="Write JSON report to this path")
    args = ap.parse_args()

    p = Path(args.docx).resolve()
    if not p.exists():
        print(f"[ERR] 文件不存在: {p}", file=sys.stderr)
        sys.exit(1)

    type_filter = parse_types(args.types)

    if not args.dry_run:
        lsof_check(p)

    # Load
    with zipfile.ZipFile(p, "r") as z:
        names = z.namelist()
        doc_xml = z.read("word/document.xml")

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_xml, parser)

    # Pre-stats
    before_fldchar = len(root.findall(f".//{W}fldChar"))
    before_instr = len(root.findall(f".//{W}instrText"))
    before_fldsimple = len(root.findall(f".//{W}fldSimple"))

    # Type distribution (before)
    type_dist = Counter()
    for it in root.findall(f".//{W}instrText"):
        txt = (it.text or "").strip()
        first = txt.split()[0] if txt else "EMPTY"
        type_dist[first] += 1
    for fs in root.findall(f".//{W}fldSimple"):
        instr = (fs.get(f"{W}instr") or "").strip()
        first = instr.split()[0] if instr else "EMPTY"
        type_dist[first] += 1

    # 1. fldSimple
    simple_removed, simple_kept = freeze_simple_fields(root, type_filter)
    # 2. complex
    complex_frozen, complex_skipped, instr_runs_removed, fldchar_runs_removed = \
        freeze_complex_fields(root, type_filter)

    # Post-stats
    after_fldchar = len(root.findall(f".//{W}fldChar"))
    after_instr = len(root.findall(f".//{W}instrText"))
    after_fldsimple = len(root.findall(f".//{W}fldSimple"))

    report = {
        "docx": str(p),
        "dry_run": args.dry_run,
        "type_filter": sorted(type_filter) if type_filter else "all",
        "before": {
            "fldChar": before_fldchar,
            "instrText": before_instr,
            "fldSimple": before_fldsimple,
        },
        "types_distribution": dict(type_dist),
        "fldSimple_frozen": len(simple_removed),
        "fldSimple_frozen_types": dict(Counter(simple_removed)),
        "fldSimple_skipped_types": dict(Counter(simple_kept)),
        "fldChar_complex_frozen": len(complex_frozen),
        "fldChar_complex_frozen_types": dict(Counter(complex_frozen)),
        "fldChar_complex_skipped_types": dict(Counter(complex_skipped)),
        "instrtext_runs_removed": instr_runs_removed,
        "fldchar_runs_removed": fldchar_runs_removed,
        "after": {
            "fldChar": after_fldchar,
            "instrText": after_instr,
            "fldSimple": after_fldsimple,
        },
    }

    print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.report:
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                     encoding="utf-8")

    if args.dry_run:
        print("[dry-run] no file written")
        return

    if not args.no_backup:
        bp = backup(p)
        print(f"[backup] {bp.name}")

    # Re-serialize and write back
    new_doc_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Rewrite zip
    tmp = p.with_suffix(p.suffix + ".tmp")
    with zipfile.ZipFile(p, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc_xml
            zout.writestr(item, data)
    tmp.replace(p)
    print(f"[done] wrote {p}")


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline: 走全 freeze (all field types); dry_run from args"""
    types = getattr(args, "freeze_types", "all") if args else "all"
    dry = bool(getattr(args, "dry_run", False)) if args else False
    p = Path(docx_path).resolve()
    type_filter = parse_types(types)
    with zipfile.ZipFile(p, "r") as z:
        doc_xml = z.read("word/document.xml")
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_xml, parser)
    before = {
        "fldChar": len(root.findall(f".//{W}fldChar")),
        "instrText": len(root.findall(f".//{W}instrText")),
        "fldSimple": len(root.findall(f".//{W}fldSimple")),
    }
    simple_removed, _ = freeze_simple_fields(root, type_filter)
    complex_frozen, _, instr_runs_removed, fldchar_runs_removed = \
        freeze_complex_fields(root, type_filter)
    after = {
        "fldChar": len(root.findall(f".//{W}fldChar")),
        "instrText": len(root.findall(f".//{W}instrText")),
        "fldSimple": len(root.findall(f".//{W}fldSimple")),
    }
    n_changed = len(simple_removed) + len(complex_frozen)
    if not dry and n_changed > 0:
        new_doc_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = p.with_suffix(p.suffix + ".tmp")
        with zipfile.ZipFile(p, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = new_doc_xml
                zout.writestr(item, data)
        tmp.replace(p)
    return {
        "changed": n_changed,
        "before": before,
        "after": after,
        "fldSimple_frozen": len(simple_removed),
        "fldChar_complex_frozen": len(complex_frozen),
    }


if __name__ == "__main__":
    main()
