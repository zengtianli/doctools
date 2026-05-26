#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
strip_orphan_media.py
=====================

单功能描述
----------
从 docx (zip) 移除 ``word/media/*`` 中**未被任何 rId 引用**的孤儿媒体文件
(常见为模板 placeholder 残留的 ``.wmf`` / ``.emf`` / ``.png`` 等)。

算法
----
1. ``zipfile`` 打开 docx (只读)
2. 收集所有关系文件 (``word/_rels/document.xml.rels`` + 所有
   ``word/_rels/header*.xml.rels`` / ``footer*.xml.rels`` /
   ``footnotes.xml.rels`` / ``endnotes.xml.rels`` / ``comments.xml.rels``)
3. 解析每个 rels, 收集 ``Target`` 指向 ``media/*`` 的项 (相对路径转 ``word/media/<name>``)
4. 列 ``word/media/`` 下所有 entry, 差集 = orphan
5. 重打包: 写新 zip 跳过 orphan 文件

触发场景
--------
- 模板继承时图替换不彻底, 留下旧 ``.wmf`` placeholder 占体积
- ``docx_cli.py audit images`` 报 orphan media >0 时清场
- 减小 docx 体积 (有时 50%+)

CLI
---
    python3 sub/strip_orphan_media.py <docx_path> \\
        [-o OUT | --inplace] [--dry-run] [--no-backup] [--report <json>]

默认行为
--------
- 默认 ``--inplace`` (留 ``.bak-N-YYYY-MM-DD.docx``)
- ``-o OUT`` 写到新路径, **不**改原文件 + **不**留 bak
- ``--dry-run`` 列将删的文件名 + 释放字节数, 不写
- 写前 ``lsof`` 自检 Word/WPS 占用

不做
----
- 不动 ``word/media/`` 之外的资源 (embeddings/charts 等)
- 不试图"修复"挂的 rId (那是 image relink 的活)
- 仅删未被引用的 media; 被引用的一律保留
"""
from __future__ import annotations

import argparse
import datetime
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = f"{{{NS_REL}}}"


# ---------------- 共用工具 ----------------

def find_next_backup(docx_path: Path) -> Path:
    today = datetime.date.today().isoformat()
    n = 1
    while True:
        cand = docx_path.with_name(
            f"{docx_path.stem}.bak-{n}-{today}{docx_path.suffix}"
        )
        if not cand.exists():
            return cand
        n += 1


def lsof_check(p: Path) -> str | None:
    try:
        r = subprocess.run(
            ["lsof", str(p)], capture_output=True, text=True, timeout=5
        )
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


# ---------------- 核心扫描 ----------------

# rels 文件可能在 word/_rels/ 下任意 *.xml.rels
_RELS_RE = re.compile(r"^word/_rels/.+\.xml\.rels$")
_MEDIA_RE = re.compile(r"^word/media/.+$")


def _collect_referenced_media(z: zipfile.ZipFile) -> set[str]:
    """收集所有 rels 文件里 Target 指向 media/* 的项,返回 zip 内规范化路径集合 (word/media/<name>)."""
    referenced: set[str] = set()
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    for name in z.namelist():
        if not _RELS_RE.match(name):
            continue
        try:
            data = z.read(name)
            root = etree.fromstring(data, parser=parser)
        except (etree.XMLSyntaxError, KeyError):
            continue
        if root is None:
            continue
        for rel in root.findall(f"{REL}Relationship"):
            target = rel.get("Target") or ""
            # Target 一般是相对 word/ 的: "media/image1.wmf" 或 "../media/x" 罕见
            # 也可能是绝对 "/word/media/x" (少见)
            t = target.replace("\\", "/").lstrip("/")
            if t.startswith("media/"):
                referenced.add("word/" + t)
            elif "/media/" in t:
                # 罕见情况, 提取 media/ 后缀
                idx = t.find("media/")
                referenced.add("word/" + t[idx:])
    return referenced


def scan_orphans(docx_path: Path) -> dict:
    """只读扫: 返回 {referenced, media_in_zip, orphans, orphan_bytes}."""
    with zipfile.ZipFile(str(docx_path), "r") as z:
        referenced = _collect_referenced_media(z)
        media_in_zip: dict[str, int] = {}  # name -> compressed size (近似释放空间)
        for info in z.infolist():
            if _MEDIA_RE.match(info.filename):
                media_in_zip[info.filename] = info.compress_size
        orphans = sorted(set(media_in_zip) - referenced)
        orphan_bytes = sum(media_in_zip[n] for n in orphans)
    return {
        "referenced_count": len(referenced),
        "media_in_zip_count": len(media_in_zip),
        "orphans": orphans,
        "orphan_count": len(orphans),
        "orphan_compressed_bytes": orphan_bytes,
    }


def rewrite_skip(docx_path: Path, out_path: Path, orphans: set[str]) -> int:
    """重打包: 跳过 orphans 集合中的文件. 返回实际跳过数."""
    skipped = 0
    tmp = out_path.with_suffix(out_path.suffix + ".tmp")
    with zipfile.ZipFile(str(docx_path), "r") as zin:
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                if it.filename in orphans:
                    skipped += 1
                    continue
                zout.writestr(it, zin.read(it.filename))
    shutil.move(str(tmp), str(out_path))
    return skipped


# ---------------- CLI ----------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("docx", type=Path)
    mx = parser.add_mutually_exclusive_group()
    mx.add_argument("-o", "--output", type=Path, default=None,
                    help="写到新路径(不动原文件,不留 bak)")
    mx.add_argument("--inplace", action="store_true", default=True,
                    help="原地改写(默认),自动留 .bak-N-YYYY-MM-DD")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-backup", action="store_true",
                        help="inplace 模式下不留 .bak")
    parser.add_argument("--report", type=Path, default=None)
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 2

    inplace = args.output is None
    if not args.dry_run and inplace:
        occ = lsof_check(args.docx)
        if occ:
            print(f"[ERR] 文件被占用 (Word/WPS 在开?), 立即停止:\n{occ}", file=sys.stderr)
            return 3

    print(f"[INFO] 扫描 {args.docx.name}")
    scan = scan_orphans(args.docx)
    print(f"  [scan] referenced={scan['referenced_count']} "
          f"media_in_zip={scan['media_in_zip_count']} "
          f"orphans={scan['orphan_count']} "
          f"orphan_bytes(compressed)={scan['orphan_compressed_bytes']}")

    report = {
        "docx": str(args.docx.resolve()),
        "dry_run": args.dry_run,
        "inplace": inplace,
        "output": str(args.output.resolve()) if args.output else None,
        "backup": None,
        "wrote": False,
        "scan": scan,
        "skipped": 0,
        "size_before": args.docx.stat().st_size,
        "size_after": None,
    }

    if scan["orphan_count"] == 0:
        print("[INFO] 无 orphan media, 不写")
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(
                json.dumps(report, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"[INFO] report -> {args.report}")
        return 0

    # dry-run: 列名
    if args.dry_run:
        print("[DRY-RUN] 将删除以下 orphan media:")
        for n in scan["orphans"][:50]:
            print(f"  - {n}")
        if scan["orphan_count"] > 50:
            print(f"  ... 共 {scan['orphan_count']} 个 (仅显示前 50)")
        print(f"[DRY-RUN] 预计释放 (compressed) {scan['orphan_compressed_bytes']} bytes")
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(
                json.dumps(report, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"[INFO] report -> {args.report}")
        return 0

    # 真写
    out_path = args.output if args.output else args.docx
    if inplace and not args.no_backup:
        bak = find_next_backup(args.docx)
        shutil.copy2(args.docx, bak)
        report["backup"] = str(bak)
        print(f"[INFO] 备份 -> {bak.name}")

    skipped = rewrite_skip(args.docx, out_path, set(scan["orphans"]))
    report["skipped"] = skipped
    report["size_after"] = out_path.stat().st_size
    report["wrote"] = True

    delta = report["size_before"] - report["size_after"]
    print(f"[OK] 删除 {skipped} 个 orphan media, "
          f"体积 {report['size_before']} -> {report['size_after']} "
          f"(减 {delta} bytes / {delta/1024:.1f} KB)")

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"[INFO] report -> {args.report}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
