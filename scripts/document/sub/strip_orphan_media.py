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
import re
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree

# sub/ 自身进 sys.path —— docx_cli 的 _dispatch 用 spec_from_file_location 加载,
# 不带脚本目录, 裸 import _cli_common 会 ImportError (append 不是 insert(0))
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parent))
import _cli_common as _cc  # noqa: E402  家族 main() 样板 SSOT

# 仓根 lib 进 sys.path —— 部件完整性断言（B 类：删除是有意的,用 diff_parts 对账）
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
from docx_parts import PartIntegrityError, diff_parts  # noqa: E402

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = f"{{{NS_REL}}}"


# ---------------- 核心扫描 ----------------

# rels 文件可能在 word/_rels/ 下任意 *.xml.rels
_RELS_RE = re.compile(r"^word/_rels/.+\.xml\.rels$")
_MEDIA_RE = re.compile(r"^word/media/.+$")

# Office namespaces used to detect rId usage in body / header / footer xml
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_R_ATTRS = (f"{{{NS_R}}}embed", f"{{{NS_R}}}link", f"{{{NS_R}}}id")


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


def _scan_used_rids_in_xml(z: zipfile.ZipFile, xml_name: str) -> set[str]:
    """Scan an OOXML part for actually-used rIds (r:embed / r:link / r:id attrs)."""
    used: set[str] = set()
    try:
        data = z.read(xml_name)
    except KeyError:
        return used
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    try:
        root = etree.fromstring(data, parser=parser)
    except etree.XMLSyntaxError:
        return used
    if root is None:
        return used
    for elem in root.iter():
        for attr in NS_R_ATTRS:
            rid = elem.get(attr)
            if rid:
                used.add(rid)
    return used


def _collect_used_media_via_body(z: zipfile.ZipFile) -> set[str]:
    """Deep scan: only count media actually referenced by an rId used in
    document.xml / header*.xml / footer*.xml / footnotes.xml / endnotes.xml / comments.xml.

    Returns set of zip-paths like "word/media/imageN.png" that are TRULY used.
    """
    # 1. Map each rels file to its used-rIds (from the corresponding XML part)
    body_xmls = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ]
    for n in z.namelist():
        if re.match(r"^word/(header|footer)\d*\.xml$", n):
            body_xmls.append(n)

    used_per_part: dict[str, set[str]] = {}
    for xn in body_xmls:
        used_per_part[xn] = _scan_used_rids_in_xml(z, xn)

    # 2. For each rels file, intersect its Relationship rIds with the used-rIds
    #    of the corresponding XML part, then collect Target → zip-path.
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    truly_referenced: set[str] = set()

    def _rels_for(part_xml: str) -> str:
        # word/document.xml -> word/_rels/document.xml.rels
        # word/header1.xml -> word/_rels/header1.xml.rels
        d, fn = part_xml.rsplit("/", 1)
        return f"{d}/_rels/{fn}.rels"

    for part_xml, used_rids in used_per_part.items():
        if not used_rids:
            continue
        rels_name = _rels_for(part_xml)
        try:
            data = z.read(rels_name)
            root = etree.fromstring(data, parser=parser)
        except (KeyError, etree.XMLSyntaxError):
            continue
        if root is None:
            continue
        for rel in root.findall(f"{REL}Relationship"):
            rid = rel.get("Id")
            if rid not in used_rids:
                continue
            target = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
            if target.startswith("media/"):
                truly_referenced.add("word/" + target)
            elif "/media/" in target:
                idx = target.find("media/")
                truly_referenced.add("word/" + target[idx:])
    return truly_referenced


def scan_orphans(docx_path: Path, deep: bool = False) -> dict:
    """只读扫: 返回 {referenced, media_in_zip, orphans, orphan_bytes}.

    deep=False (默认): 只看 rels 是否仍 list 此 media (相容旧行为).
    deep=True: 还看 document.xml/header*/footer*/footnotes/endnotes/comments 里
               有没有真的用到此 rId. 用于 split / table-extract 这类
               "body 已裁但 rels 未裁" 的场景, 此时多数 media 在 rels 里
               还在但已无 body 引用 -> 标 orphan.
    """
    with zipfile.ZipFile(str(docx_path), "r") as z:
        if deep:
            referenced = _collect_used_media_via_body(z)
        else:
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
        "deep": deep,
    }


def _rewrite_rels_drop_orphans(rels_xml: bytes, orphan_targets: set[str]) -> bytes:
    """Drop <Relationship> entries whose Target points to a removed media file.

    orphan_targets: zip-paths like "word/media/imageN.png" that have been deleted.
    Returns rewritten rels XML bytes; if nothing changed, returns input unchanged.
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    try:
        root = etree.fromstring(rels_xml, parser=parser)
    except etree.XMLSyntaxError:
        return rels_xml
    if root is None:
        return rels_xml
    changed = False
    for rel in list(root.findall(f"{REL}Relationship")):
        target = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
        if target.startswith("media/"):
            zp = "word/" + target
        elif "/media/" in target:
            idx = target.find("media/")
            zp = "word/" + target[idx:]
        else:
            continue
        if zp in orphan_targets:
            root.remove(rel)
            changed = True
    if not changed:
        return rels_xml
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def rewrite_skip(docx_path: Path, out_path: Path, orphans: set[str]) -> int:
    """重打包: 跳过 orphans 集合中的文件 + 同步清理 rels 里指向它们的 Relationship.
    返回实际跳过数 (媒体文件)."""
    skipped = 0
    tmp = out_path.with_suffix(out_path.suffix + ".tmp")
    with zipfile.ZipFile(str(docx_path), "r") as zin:
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                if it.filename in orphans:
                    skipped += 1
                    continue
                data = zin.read(it.filename)
                # If it's a rels file, drop any Relationship pointing to a deleted media
                if _RELS_RE.match(it.filename):
                    data = _rewrite_rels_drop_orphans(data, orphans)
                zout.writestr(it, data)
    # B 类完整性对账（move 之前，docx_path 未动 = 天然基线，断言炸则源件/产物均无损）：
    # ① 丢的 == 判定的 orphan 集（一个不多不少）② 无未报备新增 ③ 字节变化只落在 rels 上
    d = diff_parts(docx_path, tmp, allow_changed=frozenset())
    problems = []
    if set(d.lost) != set(orphans):
        problems.append(
            f"lost != 声明删除集: 多删 {sorted(set(d.lost) - set(orphans))} "
            f"/ 漏删 {sorted(set(orphans) - set(d.lost))}")
    if d.added:
        problems.append(f"未报备新增部件: {d.added}")
    bad_changed = [n for n in d.changed if not _RELS_RE.match(n)]
    if bad_changed:
        problems.append(f"rels 之外的部件被改写: {bad_changed}")
    if problems:
        tmp.unlink(missing_ok=True)
        raise PartIntegrityError(
            "strip_orphan_media 完整性校验未通过: " + "; ".join(problems))
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
    _cc.add_write_flags(parser, no_backup_help="inplace 模式下不留 .bak")
    parser.add_argument(
        "--deep", action="store_true",
        help="深扫: 只保留 document.xml/header*/footer*/footnotes/endnotes/"
             "comments 真正引用的 rId 对应 media; 其余即使 rels 还在 list 也算 orphan "
             "(split / table-extract 等 body 裁但 rels 未裁的场景必须开)",
    )
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 2

    inplace = args.output is None
    if not args.dry_run and inplace:
        occ = _cc.lsof_check(args.docx)
        if occ:
            print(f"[ERR] 文件被占用 (Word/WPS 在开?), 立即停止:\n{occ}", file=sys.stderr)
            return 3

    print(f"[INFO] 扫描 {args.docx.name}{' (deep)' if args.deep else ''}")
    scan = scan_orphans(args.docx, deep=args.deep)
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
        _cc.write_report(report, args.report, announce="[INFO] report -> {path}")
        return 0

    # dry-run: 列名
    if args.dry_run:
        print("[DRY-RUN] 将删除以下 orphan media:")
        for n in scan["orphans"][:50]:
            print(f"  - {n}")
        if scan["orphan_count"] > 50:
            print(f"  ... 共 {scan['orphan_count']} 个 (仅显示前 50)")
        print(f"[DRY-RUN] 预计释放 (compressed) {scan['orphan_compressed_bytes']} bytes")
        _cc.write_report(report, args.report, announce="[INFO] report -> {path}")
        return 0

    # 真写
    out_path = args.output if args.output else args.docx
    if inplace and not args.no_backup:
        bak = _cc.make_backup(args.docx)
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

    _cc.write_report(report, args.report, announce="[INFO] report -> {path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
