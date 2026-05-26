#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_images.py (2026-05-25 W1)
"""audit_images.py — audit docx 图片渲染状态 (audit-only, 不改 docx).

单功能: 对账 docx zip 内 word/media/* 与 document.xml 内 drawing/pict 引用,
找出 3 类问题:
  1. orphan media:  zip 里有图片二进制但没任何 rId 指向
  2. dangling rId:  drawing/pict 引用 rId 但 rels 缺失或目标文件不在 zip
  3. anchor 段位置: 每个 drawing/pict 所在段 idx + 段 text 片段, 便于定位
                     截图里"空框"对应哪段.

触发场景:
  用户报告 docx 里部分图片框是空的 (打开看是空白方框, 不是图).
  本脚本 audit-only 出 JSON 报告, 不修, 修复留给 W3 系列脚本.

CLI:
  python3 audit_images.py <docx_path> [--report <json_path>]

不改 docx: 全程只读 (zipfile.read + python-docx 解析), 不写回任何文件.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


def list_media(zf: zipfile.ZipFile) -> list[dict]:
    media = []
    for info in zf.infolist():
        if info.filename.startswith("word/media/") and not info.is_dir():
            name = info.filename[len("word/media/") :]
            if not name:
                continue
            ext = os.path.splitext(name)[1].lstrip(".").lower()
            media.append(
                {
                    "name": name,
                    "full_path": info.filename,
                    "size": info.file_size,
                    "ext": ext,
                    "referenced_by": [],
                }
            )
    return media


def parse_rels(zf: zipfile.ZipFile) -> dict[str, dict]:
    """返回 rId -> {target, type}"""
    try:
        data = zf.read("word/_rels/document.xml.rels")
    except KeyError:
        return {}
    tree = etree.fromstring(data)
    out = {}
    for rel in tree.findall("rels:Relationship", NS):
        rid = rel.get("Id")
        target = rel.get("Target") or ""
        rtype = rel.get("Type") or ""
        out[rid] = {"target": target, "type": rtype}
    return out


def snippet(text: str, n: int = 60) -> str:
    text = (text or "").strip().replace("\n", " ")
    if len(text) <= n:
        return text
    return text[:n] + "…"


def scan_body(doc) -> tuple[list[dict], list[dict], list[dict]]:
    """扫所有段, 返回 (drawings, picts, paragraph_index_with_text)"""
    drawings = []
    picts = []
    paragraphs = []

    body = doc.element.body
    para_idx = -1
    for child in body.iterchildren():
        tag = etree.QName(child).localname
        if tag == "p":
            para_idx += 1
            text = "".join(child.itertext())
            paragraphs.append({"idx": para_idx, "text": text})

            # drawings (DrawingML — image blip / chart / diagram)
            for d in child.findall(".//w:drawing", NS):
                rid = None
                subtype = "image"
                blip = d.find(".//a:blip", NS)
                if blip is not None:
                    rid = blip.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    ) or blip.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link"
                    )
                else:
                    # chart: <c:chart r:id="rIdN">
                    chart = d.find(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart"
                    )
                    if chart is not None:
                        rid = chart.get(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                        )
                        subtype = "chart"
                    else:
                        # diagram: <dgm:relIds r:dm="rIdN">
                        dgm = d.find(
                            ".//{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"
                        )
                        if dgm is not None:
                            rid = dgm.get(
                                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm"
                            )
                            subtype = "diagram"
                drawings.append(
                    {
                        "para_idx": para_idx,
                        "type": "drawing",
                        "subtype": subtype,
                        "rid": rid,
                        "para_text_snippet": snippet(text),
                    }
                )

            # legacy VML pict
            for p in child.findall(".//w:pict", NS):
                rid = None
                imgdata = p.find(".//v:imagedata", NS)
                if imgdata is not None:
                    rid = imgdata.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                    ) or imgdata.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}href"
                    )
                picts.append(
                    {
                        "para_idx": para_idx,
                        "type": "pict",
                        "rid": rid,
                        "para_text_snippet": snippet(text),
                    }
                )

        elif tag == "tbl":
            # 表格内段不递增主 idx, 但需扫 drawing/pict
            for tp in child.findall(".//w:p", NS):
                text = "".join(tp.itertext())
                for d in tp.findall(".//w:drawing", NS):
                    rid = None
                    blip = d.find(".//a:blip", NS)
                    if blip is not None:
                        rid = blip.get(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                        )
                    drawings.append(
                        {
                            "para_idx": para_idx,  # 表前最后一段 idx, 标识 in-table
                            "in_table": True,
                            "type": "drawing",
                            "rid": rid,
                            "para_text_snippet": snippet(text),
                        }
                    )
                for p in tp.findall(".//w:pict", NS):
                    rid = None
                    imgdata = p.find(".//v:imagedata", NS)
                    if imgdata is not None:
                        rid = imgdata.get(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                        )
                    picts.append(
                        {
                            "para_idx": para_idx,
                            "in_table": True,
                            "type": "pict",
                            "rid": rid,
                            "para_text_snippet": snippet(text),
                        }
                    )
    return drawings, picts, paragraphs


def resolve_target(rel_target: str) -> str:
    """rels Target 是相对 word/ 的, 拼成 zip 路径"""
    if not rel_target:
        return ""
    t = rel_target.lstrip("/")
    if t.startswith("media/"):
        return "word/" + t
    return "word/" + t


def audit(docx_path: Path) -> dict:
    with zipfile.ZipFile(docx_path, "r") as zf:
        zip_names = set(zf.namelist())
        media = list_media(zf)
        rels = parse_rels(zf)

    doc = Document(str(docx_path))
    drawings, picts, paragraphs = scan_body(doc)

    media_by_name = {m["name"]: m for m in media}

    def annotate(items: list[dict]) -> None:
        for it in items:
            rid = it.get("rid")
            if not rid:
                it["status"] = "no-rid"
                it["target"] = None
                continue
            rel = rels.get(rid)
            if rel is None:
                it["status"] = "dangling-no-rel"
                it["target"] = None
                continue
            target = rel["target"]
            it["target"] = target
            full = resolve_target(target)
            if full not in zip_names:
                it["status"] = "dangling-target-missing"
                continue
            # mark media referenced
            mname = os.path.basename(target)
            if mname in media_by_name:
                media_by_name[mname]["referenced_by"].append(rid)
            it["status"] = "ok"

    annotate(drawings)
    annotate(picts)

    orphan_media = [m["name"] for m in media if not m["referenced_by"]]
    dangling_rids = [
        {"type": x["type"], "para_idx": x["para_idx"], "rid": x["rid"], "status": x["status"]}
        for x in drawings + picts
        if x["status"] != "ok"
    ]

    # anchor 段: 找含"图"/"示意"/"布局"/"分布" 等图标题关键词的段, 列其相邻 drawing
    anchor_keywords = ["图", "示意", "布局", "分布", "structure", "map"]
    anchors = []
    drawing_by_para: dict[int, list[int]] = {}
    for i, d in enumerate(drawings):
        drawing_by_para.setdefault(d["para_idx"], []).append(i)
    for p in paragraphs:
        text = p["text"].strip()
        if not text:
            continue
        if any(k in text for k in anchor_keywords) and len(text) < 80:
            # 找紧邻的 drawing (该段或下一段)
            next_d = None
            for delta in (0, 1, 2, -1):
                cand = drawing_by_para.get(p["idx"] + delta, [])
                if cand:
                    next_d = {
                        "delta": delta,
                        "drawing_idx_in_list": cand[0],
                        "status": drawings[cand[0]]["status"],
                    }
                    break
            anchors.append(
                {
                    "para_idx": p["idx"],
                    "text": snippet(text, 80),
                    "nearest_drawing": next_d,
                }
            )

    summary = {
        "media_files_count": len(media),
        "drawings_count": len(drawings),
        "picts_count": len(picts),
        "orphan_media_count": len(orphan_media),
        "orphan_media": orphan_media,
        "dangling_rids_count": len(dangling_rids),
        "dangling_rids": dangling_rids,
        "issues_count": len(orphan_media) + len(dangling_rids),
    }

    return {
        "docx_path": str(docx_path),
        "summary": summary,
        "media_files": media,
        "drawings": drawings,
        "picts": picts,
        "anchor_paragraphs": anchors,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Audit docx images (audit-only).")
    ap.add_argument("docx_path", type=Path)
    ap.add_argument("--report", type=Path, default=None)
    args = ap.parse_args()

    if not args.docx_path.exists():
        print(f"ERR: not found: {args.docx_path}", file=sys.stderr)
        return 2

    report_path = args.report or Path(f"/tmp/audit-images-{args.docx_path.stem}.json")
    result = audit(args.docx_path)
    report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2))
    s = result["summary"]
    print(
        f"audit done -> {report_path}\n"
        f"  media={s['media_files_count']}  drawings={s['drawings_count']}  "
        f"picts={s['picts_count']}  orphan_media={s['orphan_media_count']}  "
        f"dangling={s['dangling_rids_count']}  issues={s['issues_count']}"
    )
    return 0


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline read-only: 走 zip 路径 audit"""
    return audit(Path(docx_path))


if __name__ == "__main__":
    sys.exit(main())
