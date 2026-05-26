#!/usr/bin/env python3
# distilled from qual-supply/scripts/strip_style_outlinelvl.py (2026-05-25 W1)
"""strip_style_outlinelvl.py — 从 word/styles.xml 移除 caption 族样式自身的 <w:outlineLvl>

## 单功能
docx 的 `word/styles.xml` 里,样式定义自带 <w:pPr><w:outlineLvl/> 时,
所有引用该样式的段落都会被 Word 视为进入大纲层级(显示在导航窗格),
即使段级 <w:outlineLvl> 已被清理也无效——段级会继承样式级。
本脚本仅删除指定 caption 族样式(默认 ZDWP 表名 / ZDWP 图名 / Caption / 0 表名 /
0 图名称 / 0图)上的 <w:outlineLvl> 元素,使图/表 caption 段彻底退出 Word 导航大纲。

## 触发场景
- 跑 strip_outlinelvl_from_captions.py(段级清理)后 Word 导航仍显示 caption
- audit_caption_outline.py 报 caption 段被视为大纲层级
- styles.xml 自查发现 caption 族样式带 outlineLvl

## 严禁动什么
- Heading 1-9 / Title / TOC Heading / 任何 "标题N" / "1.1.1.x N级标题" 等真章节样式
- docx 段级 outlineLvl(那是 strip_outlinelvl_from_captions.py 的职责)
- 样式的 type / rPr / 其他 pPr 子元素
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"

DEFAULT_CAPTION_STYLES = {
    "ZDWP 表名",
    "ZDWP 图名",
    "Caption",
    "0 表名",
    "0 图名称",
    "0图",
}

# 永远不许动的"真标题"样式(白名单兜底,即使用户误把它们传进 --styles 也拒绝)
PROTECTED_STYLE_PREFIXES = (
    "heading ",   # heading 1..9
    "Heading ",
)
PROTECTED_STYLE_NAMES = {
    "Title",
    "TOC Heading",
    "标题1", "标题2", "标题3", "标题4", "标题5", "标题6",
    "1.1.1.1四级标题",
    "--1.1.1三级标题",
}


def is_protected(name: str) -> bool:
    if name in PROTECTED_STYLE_NAMES:
        return True
    for p in PROTECTED_STYLE_PREFIXES:
        if name.startswith(p):
            return True
    # 形如 "标题N" / "X级标题" 兜底
    if "标题" in name and any(c.isdigit() for c in name):
        return True
    return False


def find_outlinelvl_styles(root) -> list[tuple[str, str, int]]:
    """返回所有带 outlineLvl 的样式列表 [(name, styleId, lvl), ...]"""
    out = []
    for s in root.findall(f".//{W}style"):
        n = s.find(f"{W}name")
        nm = n.get(W + "val") if n is not None else "?"
        sid = s.get(W + "styleId")
        ol = s.find(f"{W}pPr/{W}outlineLvl")
        if ol is not None:
            out.append((nm, sid, int(ol.get(W + "val"))))
    return out


def patch_styles_xml(docx_path: Path, target_names: set[str], dry_run: bool):
    with zipfile.ZipFile(docx_path, "r") as z:
        styles_bytes = z.read("word/styles.xml")

    root = etree.fromstring(styles_bytes)
    removed = []
    skipped_protected = []

    for s in root.findall(f".//{W}style"):
        n = s.find(f"{W}name")
        nm = n.get(W + "val") if n is not None else None
        if nm is None or nm not in target_names:
            continue
        if is_protected(nm):
            skipped_protected.append(nm)
            continue
        pPr = s.find(f"{W}pPr")
        if pPr is None:
            continue
        ol = pPr.find(f"{W}outlineLvl")
        if ol is None:
            continue
        lvl = ol.get(W + "val")
        sid = s.get(W + "styleId")
        if not dry_run:
            pPr.remove(ol)
        removed.append({"name": nm, "styleId": sid, "outlineLvl": lvl})

    if dry_run:
        return removed, skipped_protected, None

    new_bytes = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with zipfile.ZipFile(docx_path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            if it.filename == "word/styles.xml":
                zout.writestr(it, new_bytes)
            else:
                zout.writestr(it, zin.read(it.filename))
    shutil.move(str(tmp), str(docx_path))
    return removed, skipped_protected, new_bytes


def next_backup_path(docx_path: Path) -> Path:
    today = date.today().isoformat()
    n = 1
    while True:
        cand = docx_path.with_name(
            f"{docx_path.stem}.bak-{n}-{today}.docx"
        )
        if not cand.exists():
            return cand
        n += 1


def check_open(docx_path: Path):
    """lsof 自检:打开中拒绝写"""
    import subprocess
    r = subprocess.run(
        ["lsof", str(docx_path)], capture_output=True, text=True
    )
    if r.stdout.strip():
        print(f"[abort] {docx_path} 被打开(Word/WPS),请关闭后重试:\n{r.stdout}",
              file=sys.stderr)
        sys.exit(2)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("docx", help="目标 docx")
    ap.add_argument(
        "--styles",
        help=f"逗号分隔的样式名集合,默认: {','.join(sorted(DEFAULT_CAPTION_STYLES))}",
    )
    ap.add_argument("--dry-run", action="store_true", help="只统计不写")
    ap.add_argument("--no-backup", action="store_true", help="跳过 .bak 备份")
    ap.add_argument("--report", help="写 JSON 报告路径")
    args = ap.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"[abort] 不存在: {docx_path}", file=sys.stderr)
        sys.exit(2)

    if not args.dry_run:
        check_open(docx_path)

    if args.styles:
        target = {s.strip() for s in args.styles.split(",") if s.strip()}
    else:
        target = set(DEFAULT_CAPTION_STYLES)

    # 诊断:全列 outlineLvl 样式
    with zipfile.ZipFile(docx_path, "r") as z:
        sb = z.read("word/styles.xml")
    root = etree.fromstring(sb)
    pre_all = find_outlinelvl_styles(root)
    print(f"[diag] styles.xml 共 {len(pre_all)} 个样式带 outlineLvl(before)")
    pre_in_target = [t for t in pre_all if t[0] in target and not is_protected(t[0])]
    print(f"[diag] 命中 target 且非保护样式: {len(pre_in_target)}")
    for nm, sid, lvl in pre_in_target:
        print(f"  - {nm} (id={sid}, lvl={lvl})")

    # 备份
    backup_path = None
    if not args.dry_run and not args.no_backup:
        backup_path = next_backup_path(docx_path)
        shutil.copy2(docx_path, backup_path)
        print(f"[backup] {backup_path.name}")

    removed, skipped_protected, _ = patch_styles_xml(
        docx_path, target, dry_run=args.dry_run
    )

    print(f"[result] removed={len(removed)}  dry_run={args.dry_run}")
    for r in removed:
        print(f"  - {r['name']} (id={r['styleId']}, was lvl={r['outlineLvl']})")
    if skipped_protected:
        print(f"[guard] 拒绝处理保护样式: {skipped_protected}")

    if args.report:
        rep = {
            "docx": str(docx_path),
            "dry_run": args.dry_run,
            "backup": str(backup_path) if backup_path else None,
            "target_styles": sorted(target),
            "pre_outlinelvl_styles": [
                {"name": n, "styleId": s, "outlineLvl": l} for n, s, l in pre_all
            ],
            "removed": removed,
            "skipped_protected": skipped_protected,
        }
        Path(args.report).write_text(
            json.dumps(rep, ensure_ascii=False, indent=2)
        )
        print(f"[report] {args.report}")


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    styles = getattr(args, "strip_styles", None) if args else None
    if styles:
        target = {s.strip() for s in styles.split(",") if s.strip()}
    else:
        target = set(DEFAULT_CAPTION_STYLES)
    dry = bool(getattr(args, "dry_run", False)) if args else False
    removed, skipped_protected, _ = patch_styles_xml(
        Path(docx_path), target, dry_run=dry
    )
    return {
        "changed": len(removed),
        "removed": removed,
        "skipped_protected": skipped_protected,
    }


if __name__ == "__main__":
    main()
