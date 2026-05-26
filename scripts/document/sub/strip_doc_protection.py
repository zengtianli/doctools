#!/usr/bin/env python3
# distilled from qual-supply/scripts/strip_doc_protection.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
strip_doc_protection.py
========================

单功能描述
----------
从 docx 的 `word/settings.xml` 移除文档保护 / 编辑限制 / viewProtection 等
约束类元素, 确保多份子 docx 合稿时不会被子文档的保护设置卡 (合稿后 Word/
WPS 弹"文档已被保护" / "受编辑限制"等阻断, 或后续脚本写回 XML 失败)。

触发场景
--------
- 合稿前预防性清: 子 docx 来自 WPS / Word 不同模板, 可能携带遗留的
  documentProtection / trackChanges / formsDesign 等设置, 整合后 Word 弹保护
- 验收前预防性清: 客户打开主报告不希望见到"此文档受保护"提示
- pipeline 末段统一清: 与 freeze_all_fields / strip_revisions 串联

被清元素 (默认全清)
-------------------
- <w:documentProtection ...>  文档保护 (含 w:edit 属性 readOnly/comments/trackedChanges/forms)
- <w:writeProtection ...>     写保护
- <w:readModeInkLockDown>     阅读模式墨水锁
- <w:formsDesign>             表单设计模式
- <w:enforcement>             修订强制 (若以独立元素存在)
- <w:trackChanges>            修订追踪开关 (与 strip_revisions.py 重叠, 本脚本兜底)

保留 (不动)
-----------
- <w:rsids>                   修订追踪 ID (不影响保护, 不动)
- sectPr 内 protection 标记   页面属性, 不动

CLI
---
    python3 scripts/strip_doc_protection.py <docx_path> \\
        [--dry-run] [--no-backup] [--report <json>]

默认行为
--------
- 全清 6 类保护元素 (一次性, 无 type filter)
- 自动 `--backup` 写 `.bak-N-YYYY-MM-DD.docx` (除非 `--no-backup`)
- 写前 `lsof` 自检 Word/WPS 占用

不许做
------
- 删 sectPr 内 protection 标记 (那是页面属性)
- 删整个 settings.xml (只删指定子元素)
- 改 word/document.xml / styles.xml / numbering.xml
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

# 默认清除的保护类元素 local-name
PROTECTION_ELEMENTS = [
    "documentProtection",
    "writeProtection",
    "readModeInkLockDown",
    "formsDesign",
    "enforcement",
    "trackChanges",  # 兜底, 与 strip_revisions.py 重叠
]

SETTINGS_PATH = "word/settings.xml"


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


def _strip(root, dry_run: bool):
    """Remove protection elements from settings root. Return Counter of removed by tag."""
    removed = Counter()
    for tag in PROTECTION_ELEMENTS:
        # findall on settings root: protection elements are direct children of <w:settings>
        # but tolerate descendant for safety
        for el in list(root.iter(f"{W}{tag}")):
            parent = el.getparent()
            if parent is None:
                continue
            removed[tag] += 1
            if not dry_run:
                parent.remove(el)
    return removed


def _audit(root) -> Counter:
    """Pre/post audit: count protection elements in settings root."""
    c = Counter()
    for tag in PROTECTION_ELEMENTS:
        c[tag] = len(list(root.iter(f"{W}{tag}")))
    return c


def main():
    ap = argparse.ArgumentParser(description="Strip document protection / edit limits / "
                                              "viewProtection from docx settings.xml.")
    ap.add_argument("docx", help="Path to .docx")
    ap.add_argument("--dry-run", action="store_true", help="Report only, do not modify file")
    ap.add_argument("--no-backup", action="store_true", help="Skip writing .bak-N-<date>.docx")
    ap.add_argument("--report", help="Write JSON report to this path")
    args = ap.parse_args()

    p = Path(args.docx).resolve()
    if not p.exists():
        print(f"[ERR] 文件不存在: {p}", file=sys.stderr)
        sys.exit(1)

    if not args.dry_run:
        lsof_check(p)

    # Load settings.xml
    with zipfile.ZipFile(p, "r") as z:
        names = z.namelist()
        if SETTINGS_PATH not in names:
            report = {
                "docx": str(p),
                "dry_run": args.dry_run,
                "note": f"no {SETTINGS_PATH}",
                "removed": {},
                "total": 0,
            }
            print(json.dumps(report, ensure_ascii=False, indent=2))
            if args.report:
                Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                             encoding="utf-8")
            return
        settings_bytes = z.read(SETTINGS_PATH)

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(settings_bytes, parser)

    before = _audit(root)
    removed = _strip(root, dry_run=args.dry_run)
    after = _audit(root) if not args.dry_run else {k: before[k] for k in before}

    report = {
        "docx": str(p),
        "dry_run": args.dry_run,
        "before": {k: v for k, v in before.items() if v > 0},
        "removed": {k: v for k, v in removed.items() if v > 0},
        "total": sum(removed.values()),
        "after": {k: v for k, v in after.items() if v > 0},
    }

    print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.report:
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                     encoding="utf-8")

    if args.dry_run:
        print("[dry-run] no file written")
        return

    if sum(removed.values()) == 0:
        print("[noop] no protection elements found, file unchanged")
        return

    if not args.no_backup:
        bp = backup(p)
        print(f"[backup] {bp.name}")

    new_settings_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                        standalone=True)

    # Rewrite zip
    tmp = p.with_suffix(p.suffix + ".tmp")
    with zipfile.ZipFile(p, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == SETTINGS_PATH:
                data = new_settings_bytes
            zout.writestr(item, data)
    tmp.replace(p)
    print(f"[done] wrote {p}")


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline 入口: 移除全部保护类元素, dry_run from args."""
    dry = bool(getattr(args, "dry_run", False)) if args else False
    p = Path(docx_path).resolve()

    with zipfile.ZipFile(p, "r") as z:
        names = z.namelist()
        if SETTINGS_PATH not in names:
            return {"changed": 0, "removed": {}, "total": 0, "note": f"no {SETTINGS_PATH}"}
        settings_bytes = z.read(SETTINGS_PATH)

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(settings_bytes, parser)

    before = _audit(root)
    removed = _strip(root, dry_run=dry)
    n_changed = sum(removed.values())

    if not dry and n_changed > 0:
        new_settings_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                            standalone=True)
        tmp = p.with_suffix(p.suffix + ".tmp")
        with zipfile.ZipFile(p, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == SETTINGS_PATH:
                    data = new_settings_bytes
                zout.writestr(item, data)
        tmp.replace(p)

    after = _audit(root) if not dry else {k: before[k] for k in before}

    return {
        "changed": n_changed,
        "before": {k: v for k, v in before.items() if v > 0},
        "removed": {k: v for k, v in removed.items() if v > 0},
        "after": {k: v for k, v in after.items() if v > 0},
        "total": n_changed,
    }


if __name__ == "__main__":
    main()
