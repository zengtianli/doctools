#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_table_pairing.py (2026-05-25 W1)
"""audit_table_pairing.py — 只读 audit docx 中"表名段 ↔ <w:tbl>"配对状态.

检测 5 类问题:
  1. caption-name-content-mismatch  表名内容关键词与紧邻 tbl 首行列名语义不匹配
  2. orphan-caption-no-downstream-tbl  表名段下游 8 段内无 tbl
  3. orphan-tbl-no-upstream-caption    tbl 上游 8 段内无表名
  4. duplicate-caption-name            两个或多个表名段名字完全相同 (合并冲突)
  5. two-captions-compete-same-tbl     两个表名段都紧邻同一 tbl

输出 audit JSON 让人 (主会话/用户) 拍板 decision, 再喂给 pair_table_captions.py 改.

接口:
  python3 audit_table_pairing.py <docx> [--report <json>] [--quiet]

默认 stdout 打印 summary + issues; --report 写完整 JSON.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{W_NS}}}p"
W_TBL = f"{{{W_NS}}}tbl"
W_T = f"{{{W_NS}}}t"
W_PPR = f"{{{W_NS}}}pPr"
W_PSTYLE = f"{{{W_NS}}}pStyle"
W_VAL = f"{{{W_NS}}}val"
W_TR = f"{{{W_NS}}}tr"
W_TC = f"{{{W_NS}}}tc"

# 兼容两种表号: 扁平 "表3-1" 与中文章节式 "表3.1-1" (章.节-序, /docx renumber --cn-section 产出)。
# group(1)=章节号(3 或 3.1), group(2)=序号, group(3)=表名。向后兼容: (?:\.\d+)* 可选, 扁平号仍匹配。
CAP_PATTERN = re.compile(r"^\s*表\s*(\d+(?:\.\d+)*)\s*[-–—]\s*(\d+)\s*(.*)$")

# 关键词→同义词字典 (字面包含即视为命中)
KEYWORD_SYNONYMS: dict[str, list[str]] = {
    "成本": ["成本", "占比", "费用", "造价", "运营成本"],
    "价格": ["价格", "水价", "价差", "标准", "元/m³", "调价", "调整"],
    "目标": ["目标", "指标", "阶段", "任务"],
    "职责": ["职责", "部门", "分工", "主导"],
    "原则": ["原则", "措施", "核心要求"],
    "对比": ["对比", "占比", "嘉兴", "浙江省", "全省"],
    "格局": ["水源类型", "供水", "占比", "水质", "等级"],
    "政策": ["层级", "政策", "文件", "时间", "文号", "要求"],
    "实践": ["城市", "模式", "规模", "利用率", "置换"],
    "必要性": ["维度", "问题", "效果", "案例"],
    "前置": ["环节", "嵌入", "管控", "约束", "前置"],
    "差异化": ["用户类别", "水源类型", "价格", "价差", "适用"],
    "保障": ["保障类型", "措施", "实施", "依据", "主体"],
    "联动": ["阶段", "时间", "目标", "指标", "效益", "任务"],
    "征收": ["征收", "监督", "标准", "主体"],
}


def get_style_id(p) -> str:
    ppr = p.find(W_PPR)
    if ppr is None:
        return ""
    ps = ppr.find(W_PSTYLE)
    if ps is None:
        return ""
    return ps.get(W_VAL) or ""


def get_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(W_T)).strip()


def get_first_row_cells(tbl) -> list[str]:
    cells = []
    for tr in tbl.iter(W_TR):
        for tc in tr.iter(W_TC):
            txt = get_text(tc)
            cells.append(txt[:40])
        break
    return cells


def count_rows_cols(tbl) -> tuple[int, int]:
    rows = list(tbl.iter(W_TR))
    if not rows:
        return 0, 0
    cols = len(list(rows[0].iter(W_TC)))
    return len(rows), cols


def heuristic_match(caption_name: str, tbl_first_row: list[str]) -> tuple[float, str]:
    """对 caption 名字关键词与 tbl 首行 cell 命中度打分.

    返回 (score, method).
      score = 命中关键词数 / 名字关键词总数 (0.0 .. 1.0)
      method = "kw-match:k1+k2..." 或 "no-match"
    """
    if not caption_name:
        return 0.0, "no-caption-name"
    first_row_text = "|".join(tbl_first_row)
    hits = []
    total_kws = 0
    for kw, syns in KEYWORD_SYNONYMS.items():
        if kw in caption_name:
            total_kws += 1
            if any(s in first_row_text for s in syns):
                hits.append(kw)
    if total_kws == 0:
        return 0.0, "no-keyword-extracted"
    score = len(hits) / total_kws
    if hits:
        return score, f"kw-match:{'+'.join(hits)}"
    return 0.0, "no-match"


def _audit_from_doc(doc, docx_path_label: str = "") -> dict:
    elems = list(doc.element.body.iterchildren())

    # 收集 captions
    captions = []
    cap_counter = 0
    for i, e in enumerate(elems):
        if e.tag == W_P:
            text = get_text(e)
            m = CAP_PATTERN.match(text)
            if m:
                cap_counter += 1
                ch, num, name = m.group(1), m.group(2), m.group(3).strip()
                # 找紧邻下游"注:..."段 (style ZDWP, 文本以"注"开头)
                notes_idx = []
                j = i + 1
                while j < len(elems) and elems[j].tag == W_P:
                    nt = get_text(elems[j])
                    if nt.startswith(("注:", "注：", "注", "Note", "note")) and len(nt) < 200:
                        notes_idx.append(j)
                        j += 1
                    else:
                        break
                captions.append({
                    "id": f"cap-{cap_counter}",
                    "elem_idx": i,
                    "number": f"表{ch}-{num}",
                    "chapter": ch,  # 字符串: 兼容 cn-section "3.1" (旧扁平 "3" 亦为字符串)
                    "seq": int(num),
                    "name": name,
                    "style": get_style_id(e),
                    "raw_text": text,
                    "notes_idx": notes_idx,
                })

    # 收集 tbls
    tbls = []
    tbl_counter = 0
    for i, e in enumerate(elems):
        if e.tag == W_TBL:
            tbl_counter += 1
            cells = get_first_row_cells(e)
            rc, cc = count_rows_cols(e)
            tbls.append({
                "id": f"tbl-{tbl_counter}",
                "elem_idx": i,
                "first_row_cells": cells,
                "row_count": rc,
                "col_count": cc,
            })

    # 为每个 caption 算 nearest_downstream_tbl + heuristic_match
    for cap in captions:
        nxt = None
        for t in tbls:
            if t["elem_idx"] > cap["elem_idx"]:
                nxt = t
                break
        if nxt:
            cap["nearest_downstream_tbl"] = nxt["id"]
            cap["nearest_downstream_tbl_idx"] = nxt["elem_idx"]
            cap["distance"] = nxt["elem_idx"] - cap["elem_idx"]
            # 启发匹配下游 5 个候选 tbl 中得分最高的
            candidates = [t for t in tbls
                          if cap["elem_idx"] < t["elem_idx"] <= cap["elem_idx"] + 30]
            best_score, best_method, best_tid = 0.0, "no-match", nxt["id"]
            for c in candidates:
                s, m = heuristic_match(cap["name"], c["first_row_cells"])
                if s > best_score:
                    best_score, best_method, best_tid = s, m, c["id"]
            cap["heuristic_match"] = {
                "tbl_id": best_tid, "score": round(best_score, 3), "method": best_method,
            }
        else:
            cap["nearest_downstream_tbl"] = None
            cap["distance"] = None
            cap["heuristic_match"] = {"tbl_id": None, "score": 0.0, "method": "no-downstream"}

    # 为每个 tbl 算上游 8 段内的 captions
    for tbl in tbls:
        ups = [c["id"] for c in captions
               if tbl["elem_idx"] - 8 <= c["elem_idx"] < tbl["elem_idx"]]
        tbl["upstream_captions_within_8_paras"] = ups

    # 检测 issues
    issues = []

    # 1. orphan-caption-no-downstream-tbl  (孤儿表名: 与紧邻下游 tbl 距离 > 5)
    for cap in captions:
        if cap.get("distance") is None or cap["distance"] > 5:
            issues.append({
                "type": "orphan-caption-no-downstream-tbl",
                "caption_id": cap["id"],
                "caption_number": cap["number"],
                "caption_name": cap["name"],
                "elem_idx": cap["elem_idx"],
                "details": f"下游最近 tbl 距离 = {cap.get('distance')}",
            })

    # 2. orphan-tbl-no-upstream-caption (孤儿表: 上游 8 段无表名)
    for tbl in tbls:
        if not tbl["upstream_captions_within_8_paras"]:
            issues.append({
                "type": "orphan-tbl-no-upstream-caption",
                "tbl_id": tbl["id"],
                "elem_idx": tbl["elem_idx"],
                "first_row": tbl["first_row_cells"],
                "details": "上游 8 段内无表名段",
            })

    # 3. duplicate-caption-name (重名: 名字完全相同, 名字非空)
    from collections import defaultdict
    by_name = defaultdict(list)
    for cap in captions:
        if cap["name"]:
            by_name[cap["name"]].append(cap)
    for name, lst in by_name.items():
        if len(lst) > 1:
            issues.append({
                "type": "duplicate-caption-name",
                "name": name,
                "caption_ids": [c["id"] for c in lst],
                "details": f"{len(lst)} 个表名同名: " + ", ".join(
                    f"{c['id']}({c['number']})" for c in lst
                ),
            })

    # 4. two-captions-compete-same-tbl (两表名抢同一 tbl)
    by_tbl = defaultdict(list)
    for tbl in tbls:
        for cid in tbl["upstream_captions_within_8_paras"]:
            by_tbl[tbl["id"]].append(cid)
    for tid, caps in by_tbl.items():
        if len(caps) >= 2:
            issues.append({
                "type": "two-captions-compete-same-tbl",
                "tbl_id": tid,
                "competing_captions": caps,
                "details": f"tbl {tid} 上游 8 段内同时存在 {len(caps)} 个表名: " + ", ".join(caps),
            })

    # 5. caption-name-content-mismatch (启发分 = 0 且名字含已知关键词)
    for cap in captions:
        if not cap["name"]:
            continue
        hm = cap["heuristic_match"]
        if hm["score"] == 0.0 and hm["method"] == "no-match" and cap.get("distance") and cap["distance"] <= 5:
            # 名字含已知关键词但 tbl 首行不命中
            issues.append({
                "type": "caption-name-content-mismatch",
                "caption_id": cap["id"],
                "caption_number": cap["number"],
                "caption_name": cap["name"],
                "tbl_id": hm["tbl_id"],
                "details": f"表名关键词与下游 tbl 首行 cell 字面无交集",
            })

    # 6. empty-caption-name
    for cap in captions:
        if not cap["name"]:
            issues.append({
                "type": "empty-caption-name",
                "caption_id": cap["id"],
                "caption_number": cap["number"],
                "elem_idx": cap["elem_idx"],
                "details": "表名段编号后无标题文字",
            })

    summary = {
        "captions": len(captions),
        "tbls": len(tbls),
        "orphan_captions": sum(1 for i in issues if i["type"] == "orphan-caption-no-downstream-tbl"),
        "orphan_tbls": sum(1 for i in issues if i["type"] == "orphan-tbl-no-upstream-caption"),
        "duplicate_caption_names": sum(1 for i in issues if i["type"] == "duplicate-caption-name"),
        "competing_pairs": sum(1 for i in issues if i["type"] == "two-captions-compete-same-tbl"),
        "content_mismatches": sum(1 for i in issues if i["type"] == "caption-name-content-mismatch"),
        "empty_names": sum(1 for i in issues if i["type"] == "empty-caption-name"),
    }

    return {
        "docx_path": docx_path_label,
        "summary": summary,
        "captions": captions,
        "tbls": tbls,
        "issues": issues,
    }


def audit(docx_path: Path) -> dict:
    doc = Document(str(docx_path))
    return _audit_from_doc(doc, str(docx_path))


def apply(doc, args=None) -> dict:
    """pipeline read-only adapter"""
    label = str(getattr(args, "docx", "")) if args else ""
    return _audit_from_doc(doc, label)


def print_summary(audit_data: dict) -> None:
    s = audit_data["summary"]
    print(f"\n{'='*78}")
    print(f"audit: {audit_data['docx_path']}")
    print(f"{'='*78}")
    print(f"captions = {s['captions']} | tbls = {s['tbls']} | "
          f"差 = {s['captions'] - s['tbls']} (孤儿表名数)")
    print(f"orphan_captions  = {s['orphan_captions']}")
    print(f"orphan_tbls      = {s['orphan_tbls']}")
    print(f"duplicate_names  = {s['duplicate_caption_names']}")
    print(f"competing_pairs  = {s['competing_pairs']}")
    print(f"content_mismatch = {s['content_mismatches']}")
    print(f"empty_names      = {s['empty_names']}")
    print(f"\n{'─'*78}")
    print("issues:")
    for issue in audit_data["issues"]:
        head = f"  [{issue['type']}]"
        rest = " ".join(f"{k}={v}" for k, v in issue.items() if k != "type")
        print(f"{head} {rest}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("docx", help="输入 docx 路径")
    ap.add_argument("--report", help="完整 audit JSON 输出路径")
    ap.add_argument("--quiet", action="store_true", help="不打印 summary stdout")
    args = ap.parse_args(argv)

    src = Path(args.docx)
    if not src.exists():
        print(f"[error] 找不到 {src}", file=sys.stderr)
        return 2

    audit_data = audit(src)
    if not args.quiet:
        print_summary(audit_data)
    if args.report:
        Path(args.report).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report).write_text(
            json.dumps(audit_data, indent=2, ensure_ascii=False)
        )
        print(f"\n[report] {args.report}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
