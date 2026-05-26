#!/usr/bin/env python3
# distilled from qual-supply/scripts/strip_revisions.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
strip_revisions.py
==================

单功能描述
----------
从 docx 移除所有「修订相关」残留, 让合稿前文档干净:
  1. 接受所有插入 ``<w:ins>`` (展开内部 run, 删 ins 包裹)
  2. 接受所有删除 ``<w:del>`` (整体移除 del 元素 + 内部 run, 即删除"被标记删除"的内容)
  3. 删 ``<w:commentRangeStart>`` / ``<w:commentRangeEnd>`` / ``<w:commentReference>``
     (commentReference 自身移除, 保留外层 run 防破坏段结构)
  4. 清空 ``word/comments.xml`` (留个 valid 空壳防 Word 抱怨; --keep-comments 关闭)
  5. 关 trackChanges (``word/settings.xml`` 移除 ``<w:trackChanges/>``)

不做
----
- 不删 rsid / rsidR / rsidRDefault 等修订追踪 ID 属性 (数量巨大, 不影响合稿)
- 不动 sectPr / pPr 中其他修订属性 (pPrChange / rPrChange 等暂忽略, 接受范围仅 ins/del/comments/trackChanges)
- 不裸解 zip 改 XML 字符串 — settings.xml / comments.xml 走 lxml.etree 解析改写

触发场景
--------
- 合稿前 docx 还残留 Word 修订模式产物 (ins/del/批注/trackChanges 开关)
- 需要 "扁平化" 文档让后续 normalize / renumber 等脚本不被 ins/del 包裹打乱段结构

CLI
---
    python3 scripts/strip_revisions.py <docx_path> \\
        [--mode accept-all] [--keep-comments] [--dry-run] [--no-backup] [--report <json>]

  --mode accept-all   接受所有修订 (ins 展开 + del 删除); 默认且当前唯一支持模式
  --keep-comments     保留 word/comments.xml 原内容 (默认清空)
  --dry-run           列每类计数不写
  --no-backup         不备份 .bak-N-YYYY-MM-DD.docx
  --report <json>     写 JSON 报告

Pipeline ready
--------------
- ``apply(doc, args) -> dict`` : 改 doc 内存对象 (body 内 ins/del/comment*)
- ``apply_path(docx_path, args) -> dict`` : 改 zip 内 word/comments.xml + word/settings.xml

约束
----
- python-docx + lxml, 禁裸 zip XML 字符串改 (settings/comments 走 etree.fromstring)
- 默认 --backup, 写前 lsof 自检 Word/WPS 占用 (占用立即退出)
- 一脚本一功能: 仅处理修订/批注/trackChanges, 不动样式/编号/段顺序
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path

from docx import Document
from lxml import etree

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{NS}}}"


# ---------------- 共用工具 ----------------

def find_next_backup(docx_path: Path) -> Path:
    today = date.today().isoformat()
    stem = docx_path.stem
    parent = docx_path.parent
    n = 1
    while True:
        cand = parent / f"{stem}.bak-{n}-{today}.docx"
        if not cand.exists():
            return cand
        n += 1


def lsof_check(docx_path: Path) -> str | None:
    """返回非空字符串 = 被占用, 返回 None = 可写"""
    try:
        out = subprocess.run(
            ["lsof", str(docx_path)],
            capture_output=True, text=True, timeout=5,
        )
        if out.returncode == 0 and out.stdout.strip():
            return out.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


# ---------------- 内存改 (body 内 ins/del/comment*) ----------------

def _strip_in_element(root) -> dict:
    """对一个 lxml element (body 或 header/footer root) 处理 ins/del/comment*"""
    ins_count = 0
    del_count = 0
    crs_count = 0
    cre_count = 0
    cref_count = 0

    # 1. <w:ins>: 展开内部子节点到父
    for ins in list(root.iter(f"{W}ins")):
        parent = ins.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ins)
        for child in list(ins):
            parent.insert(idx, child)
            idx += 1
        parent.remove(ins)
        ins_count += 1

    # 2. <w:del>: 整体删除 (内部 run 一起走)
    for del_el in list(root.iter(f"{W}del")):
        parent = del_el.getparent()
        if parent is None:
            continue
        parent.remove(del_el)
        del_count += 1

    # 3. comment 引用 三种全删
    for tag, counter_name in [
        (f"{W}commentRangeStart", "crs"),
        (f"{W}commentRangeEnd", "cre"),
        (f"{W}commentReference", "cref"),
    ]:
        for el in list(root.iter(tag)):
            parent = el.getparent()
            if parent is None:
                continue
            parent.remove(el)
            if counter_name == "crs":
                crs_count += 1
            elif counter_name == "cre":
                cre_count += 1
            else:
                cref_count += 1

    return {
        "ins_accepted": ins_count,
        "del_accepted": del_count,
        "comment_range_start_removed": crs_count,
        "comment_range_end_removed": cre_count,
        "comment_reference_removed": cref_count,
    }


def apply(doc, args=None) -> dict:
    """改 doc 内存对象的 body — pipeline 接口"""
    mode = getattr(args, "revision_mode", "accept-all") if args else "accept-all"
    body = doc.element.body
    result = _strip_in_element(body)
    result["mode"] = mode
    return result


# ---------------- zip 改 (comments.xml + settings.xml) ----------------

def apply_path(docx_path, args=None) -> dict:
    """改 word/comments.xml (清空, 除非 keep_comments) + word/settings.xml (删 trackChanges)"""
    keep_comments = bool(getattr(args, "keep_comments", False)) if args else False
    docx_path = Path(docx_path)
    tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")

    track_changes_removed = 0
    comments_emptied = False
    has_comments = False
    has_settings = False

    parser = etree.XMLParser(remove_blank_text=False)

    with zipfile.ZipFile(str(docx_path), 'r') as zin:
        with zipfile.ZipFile(str(tmp), 'w', zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == "word/settings.xml":
                    has_settings = True
                    try:
                        root = etree.fromstring(data, parser=parser)
                        for tc in root.findall(f".//{W}trackChanges"):
                            tc.getparent().remove(tc)
                            track_changes_removed += 1
                        data = etree.tostring(
                            root, xml_declaration=True,
                            encoding='UTF-8', standalone=True,
                        )
                    except etree.XMLSyntaxError:
                        pass
                elif it.filename == "word/comments.xml" and not keep_comments:
                    has_comments = True
                    empty_xml = (
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        b'<w:comments xmlns:w="' + NS.encode() + b'"/>'
                    )
                    data = empty_xml
                    comments_emptied = True
                zout.writestr(it, data)

    shutil.move(str(tmp), str(docx_path))
    return {
        "trackchanges_removed": track_changes_removed,
        "comments_emptied": comments_emptied,
        "has_comments_xml": has_comments,
        "has_settings_xml": has_settings,
    }


# ---------------- pre-scan (用 zip + lxml 不改, 算 before counts) ----------------

def scan_counts(docx_path: Path) -> dict:
    """只读扫 before 数量, 供报告用"""
    counts = {
        "ins": 0,
        "del": 0,
        "comment_range_start": 0,
        "comment_range_end": 0,
        "comment_reference": 0,
        "trackChanges": 0,
        "comments_xml_present": False,
    }
    try:
        with zipfile.ZipFile(str(docx_path)) as z:
            names = set(z.namelist())
            if "word/document.xml" in names:
                doc = etree.fromstring(z.read("word/document.xml"))
                counts["ins"] = len(doc.findall(f".//{W}ins"))
                counts["del"] = len(doc.findall(f".//{W}del"))
                counts["comment_range_start"] = len(doc.findall(f".//{W}commentRangeStart"))
                counts["comment_range_end"] = len(doc.findall(f".//{W}commentRangeEnd"))
                counts["comment_reference"] = len(doc.findall(f".//{W}commentReference"))
            if "word/settings.xml" in names:
                s = etree.fromstring(z.read("word/settings.xml"))
                counts["trackChanges"] = len(s.findall(f".//{W}trackChanges"))
            counts["comments_xml_present"] = "word/comments.xml" in names
    except (zipfile.BadZipFile, etree.XMLSyntaxError, KeyError):
        pass
    return counts


# ---------------- CLI ----------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("docx", type=Path)
    parser.add_argument(
        "--mode", choices=["accept-all"], default="accept-all",
        help="目前仅支持 accept-all (接受所有修订: ins 展开 + del 删除)",
    )
    parser.add_argument("--keep-comments", action="store_true",
                        help="保留 word/comments.xml (默认清空)")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-backup", action="store_true")
    parser.add_argument("--report", type=Path, default=None)
    args = parser.parse_args()

    # 兼容 pipeline adapter 名字
    args.revision_mode = args.mode

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 2

    if not args.dry_run:
        occ = lsof_check(args.docx)
        if occ:
            print(f"[ERR] 文件被占用 (Word/WPS 在开?), 立即停止:\n{occ}", file=sys.stderr)
            return 3

    before = scan_counts(args.docx)
    print(f"[INFO] 扫描 {args.docx.name}")
    print(f"  [before] ins={before['ins']} del={before['del']} "
          f"crs={before['comment_range_start']} cre={before['comment_range_end']} "
          f"cref={before['comment_reference']} trackChanges={before['trackChanges']} "
          f"comments_xml={'yes' if before['comments_xml_present'] else 'no'}")

    # ---- 1) 内存改 body
    doc = Document(str(args.docx))
    body_result = apply(doc, args)

    report = {
        "docx": str(args.docx.resolve()),
        "dry_run": args.dry_run,
        "mode": args.mode,
        "keep_comments": args.keep_comments,
        "backup": None,
        "wrote": False,
        "before": before,
        "body_changes": body_result,
        "zip_changes": None,
    }

    print(f"  [body] ins_accepted={body_result['ins_accepted']} "
          f"del_accepted={body_result['del_accepted']} "
          f"crs={body_result['comment_range_start_removed']} "
          f"cre={body_result['comment_range_end_removed']} "
          f"cref={body_result['comment_reference_removed']}")

    if args.dry_run:
        print("[DRY-RUN] 不写文件")
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(
                json.dumps(report, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"[INFO] report -> {args.report}")
        return 0

    # 是否要写? 只要 body 有改 或 settings/comments 需要清就写
    need_write_body = sum([
        body_result['ins_accepted'],
        body_result['del_accepted'],
        body_result['comment_range_start_removed'],
        body_result['comment_range_end_removed'],
        body_result['comment_reference_removed'],
    ]) > 0
    need_zip = (before['trackChanges'] > 0) or (
        before['comments_xml_present'] and not args.keep_comments
    )

    if not need_write_body and not need_zip:
        print("[INFO] 无需变更, 不写文件")
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(
                json.dumps(report, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"[INFO] report -> {args.report}")
        return 0

    if not args.no_backup:
        bak = find_next_backup(args.docx)
        shutil.copy2(args.docx, bak)
        report["backup"] = str(bak)
        print(f"[INFO] 备份 -> {bak.name}")

    # ---- 2) 写 body (即使 body 没变也 save 一次保证一致, 但只在 need_write_body 时 save)
    if need_write_body:
        doc.save(str(args.docx))

    # ---- 3) zip 改 comments.xml + settings.xml
    if need_zip:
        zip_result = apply_path(args.docx, args)
        report["zip_changes"] = zip_result
        print(f"  [zip] trackchanges_removed={zip_result['trackchanges_removed']} "
              f"comments_emptied={zip_result['comments_emptied']}")

    report["wrote"] = True
    print(f"[OK] 已清理修订残留, 写回 {args.docx.name}")

    # ---- 4) verify (after counts)
    after = scan_counts(args.docx)
    report["after"] = after
    print(f"  [after]  ins={after['ins']} del={after['del']} "
          f"crs={after['comment_range_start']} cre={after['comment_range_end']} "
          f"cref={after['comment_reference']} trackChanges={after['trackChanges']}")

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
