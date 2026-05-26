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
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

# rels 文件可能在 word/_rels/ 下任意 *.xml.rels (document / header* / footer* /
# footnotes / endnotes / comments / numbering ...). 主 document 引用走
# document.xml.rels, 但 header/footer/footnotes 等各自带 rels 文件且各自的
# rId 命名空间独立 (rId1 在 document.xml.rels vs header1.xml.rels 可指不同 target).
_RELS_RE = re.compile(r"^word/_rels/(.+)\.xml\.rels$")
# 主 part XML (排除 _rels/ / theme/ / settings 等纯样式), 凡可能 embed image 的:
_PART_XML_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments|numbering)\.xml$"
)

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
    """[legacy] 返回 document.xml.rels 的 rId -> {target, type} (向后兼容).

    新代码应该用 parse_all_rels(), 该函数只覆盖主 document 的 rels, 不含
    header/footer/footnotes 等 part 各自的 rels. 保留只为兼容已有调用者.
    """
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


def parse_all_rels(zf: zipfile.ZipFile) -> dict[str, dict[str, dict]]:
    """扫所有 word/_rels/*.xml.rels, 返回 {part_name -> {rId -> {target, type}}}.

    part_name = rels 所属的 part 的 stem, 例: 'document', 'header1', 'footer3'.
    每个 part 的 rId 命名空间独立, 必须按 part 隔离查询.

    复用 strip_orphan_media._collect_referenced_media 的扫法 (但保留 rId 维度).
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    out: dict[str, dict[str, dict]] = {}
    for name in zf.namelist():
        m = _RELS_RE.match(name)
        if not m:
            continue
        part_name = m.group(1)  # 'document' / 'header1' / 'footer3' ...
        try:
            data = zf.read(name)
            tree = etree.fromstring(data, parser=parser)
        except (etree.XMLSyntaxError, KeyError):
            continue
        if tree is None:
            continue
        part_rels: dict[str, dict] = {}
        for rel in tree.findall("rels:Relationship", NS):
            rid = rel.get("Id")
            target = rel.get("Target") or ""
            rtype = rel.get("Type") or ""
            part_rels[rid] = {"target": target, "type": rtype}
        out[part_name] = part_rels
    return out


def collect_rid_refs_in_part(zf: zipfile.ZipFile, part_path: str) -> list[dict]:
    """扫一个 part XML 的 body, 返回 [{rid, kind}] (kind=drawing/pict/chart/diagram).

    drawing/pict 都查 r:embed / r:link / r:id 各种姿势.
    用于扫 header*/footer*/footnotes/endnotes/comments 的 body, 把它们
    引用的 rId 也算 referenced, 避免 audit-images 漏报这些 part 引用的 media.
    """
    refs: list[dict] = []
    try:
        data = zf.read(part_path)
    except KeyError:
        return refs
    try:
        root = etree.fromstring(data, parser=etree.XMLParser(recover=True))
    except etree.XMLSyntaxError:
        return refs
    if root is None:
        return refs
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    # DrawingML blip / chart / diagram
    for blip in root.findall(".//a:blip", NS):
        rid = blip.get(f"{R}embed") or blip.get(f"{R}link")
        if rid:
            refs.append({"rid": rid, "kind": "drawing"})
    for chart in root.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart"
    ):
        rid = chart.get(f"{R}id")
        if rid:
            refs.append({"rid": rid, "kind": "chart"})
    for dgm in root.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"
    ):
        rid = dgm.get(f"{R}dm")
        if rid:
            refs.append({"rid": rid, "kind": "diagram"})

    # legacy VML pict / imagedata
    for imgdata in root.findall(".//v:imagedata", NS):
        rid = imgdata.get(f"{R}id") or imgdata.get(f"{R}href")
        if rid:
            refs.append({"rid": rid, "kind": "pict"})

    # OLE object (oleObject r:id 也算引用, 走 embeddings 路径不算 media 但要登记)
    return refs


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
        rels = parse_rels(zf)  # legacy: 只 document.xml.rels
        all_rels = parse_all_rels(zf)  # 全 part rels (含 header/footer/footnotes/...)
        # 扫所有 part XML body 收集 rId 引用 (含 header/footer/footnotes/endnotes/comments/numbering)
        # 这是修正旧版漏扫的关键: 旧版只标 document.xml body 的 drawing/pict rId,
        # 导致 header/footer 里 imagedata 引用的 media 全被误判 orphan.
        part_refs: dict[str, list[dict]] = {}  # part_name -> [{rid, kind}]
        for zname in zf.namelist():
            pm = _PART_XML_RE.match(zname)
            if not pm:
                continue
            part_name = pm.group(1)
            part_refs[part_name] = collect_rid_refs_in_part(zf, zname)

    doc = Document(str(docx_path))
    drawings, picts, paragraphs = scan_body(doc)

    media_by_name = {m["name"]: m for m in media}

    def annotate(items: list[dict]) -> None:
        """标注 document.xml body 内 drawing/pict 的 status + 落 referenced_by.

        rId 解析走 document part 的 rels (`all_rels['document']`),
        与旧 `rels` 字典等价.
        """
        doc_rels = all_rels.get("document", rels)
        for it in items:
            rid = it.get("rid")
            if not rid:
                it["status"] = "no-rid"
                it["target"] = None
                continue
            rel = doc_rels.get(rid)
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
            # mark media referenced (带来源标注: "document:rIdN")
            mname = os.path.basename(target)
            if mname in media_by_name:
                media_by_name[mname]["referenced_by"].append(f"document:{rid}")
            it["status"] = "ok"

    annotate(drawings)
    annotate(picts)

    # 全 part 扫描: 把 header/footer/footnotes/... 引用的 media 也标 referenced.
    # 这是消除磐安 .bak-1 47 个 .wmf 误报为 orphan 的核心修复.
    for part_name, refs in part_refs.items():
        if part_name == "document":
            # document.xml body 已通过 annotate() 标过, 跳过避免重复 (但下面也兜底)
            pass
        part_rels = all_rels.get(part_name, {})
        for ref in refs:
            rid = ref["rid"]
            rel = part_rels.get(rid)
            if rel is None:
                continue
            target = rel["target"]
            full = resolve_target(target)
            if full not in zip_names:
                continue
            mname = os.path.basename(target)
            if mname in media_by_name:
                tag = f"{part_name}:{rid}"
                if tag not in media_by_name[mname]["referenced_by"]:
                    media_by_name[mname]["referenced_by"].append(tag)

    # 兜底: 严谨算法对账 — 若 rels 里 Target 直接指向 word/media/<name>,
    # 也算 referenced (与 strip_orphan_media.scan_orphans 语义对齐).
    # 这覆盖某些非常规 part 引用 / 模板残留 rels 仍真指 media 的边界情况.
    for part_name, part_rels in all_rels.items():
        for rid, rel in part_rels.items():
            target = (rel.get("target") or "").replace("\\", "/").lstrip("/")
            if not target:
                continue
            if target.startswith("media/") or "/media/" in target:
                mname = os.path.basename(target)
                if mname in media_by_name:
                    tag = f"rels[{part_name}]:{rid}"
                    if not any(
                        t.startswith(f"{part_name}:") or t == tag
                        for t in media_by_name[mname]["referenced_by"]
                    ):
                        media_by_name[mname]["referenced_by"].append(tag)

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
        # 修复 2026-05-26: 列出已扫的 rels parts (含 header/footer/footnotes/...),
        # 验证 orphan_media_count 与 strip_orphan_media 对齐.
        "rels_parts_scanned": sorted(all_rels.keys()),
        "part_xmls_scanned": sorted(part_refs.keys()),
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
