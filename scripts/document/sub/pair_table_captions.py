#!/usr/bin/env python3
"""pair_table_captions.py — 按 decision JSON 修 docx 中"表名段 ↔ <w:tbl>"配对.

支持 5 个 op:
  - delete-caption        : 删指定 caption 段 (只删 <w:p>, 不动 tbl/正文)
  - rename-caption        : 改 caption 段文字 (new_number / new_name; 保留 run 级 bold/字号)
  - rename-orphan-tbl     : 在 tbl 紧前方插入新 caption (复用近邻 caption 段的 pPr + run rPr 当模板)
  - pair-caption-to-tbl   : 把 caption 段块(含 notes:开头注释段) 移到 tbl 紧前方
  - renumber-all-tables   : 按 body 物理顺序 + H1 章节, 重编 "表 X-Y" 编号

caption_id 体系: 在执行任何 op 前 snapshot 一次 body 中所有 "表 X-Y" 段, 按出现顺序
派 cap-1..cap-N; tbl_id 同理 tbl-1..tbl-M. 后续 op 通过 id 反查元素 (避开 idx 漂移).

接口:
  python3 pair_table_captions.py <docx> --decision <decision.json> [--dry-run]
                                  [--no-backup] [--report <out.json>]

默认 真改 + 自动备份 .bak-N-<date>.docx
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path
from typing import Optional

from docx import Document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{W_NS}}}p"
W_TBL = f"{{{W_NS}}}tbl"
W_R = f"{{{W_NS}}}r"
W_T = f"{{{W_NS}}}t"
W_PPR = f"{{{W_NS}}}pPr"
W_PSTYLE = f"{{{W_NS}}}pStyle"
W_RPR = f"{{{W_NS}}}rPr"
W_VAL = f"{{{W_NS}}}val"

CAP_PATTERN = re.compile(r"^\s*表\s*(\d+)\s*[-–—]\s*(\d+)\s*(.*)$")


def lsof_check(docx_path: Path) -> Optional[str]:
    try:
        out = subprocess.run(
            ["lsof", "--", str(docx_path)],
            capture_output=True, text=True, timeout=5,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None
    if out.returncode == 0 and out.stdout.strip():
        lines = out.stdout.strip().split("\n")
        if len(lines) > 1:
            return "\n".join(lines)
    return None


def get_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(W_T)).strip()


def get_style_id(p) -> str:
    ppr = p.find(W_PPR)
    if ppr is None:
        return ""
    ps = ppr.find(W_PSTYLE)
    if ps is None:
        return ""
    return ps.get(W_VAL) or ""


def is_h1(p) -> bool:
    sid = get_style_id(p)
    return sid in ("1", "Heading1", "heading1", "Heading 1", "h1")


def snapshot_ids(body) -> tuple[dict, dict]:
    """扫 body, 派 cap-N + tbl-N 稳定 ID, 返回 (id→element) 双映射."""
    cap_map = {}
    tbl_map = {}
    cap_n = 0
    tbl_n = 0
    for child in body:
        if child.tag == W_P:
            text = get_text(child)
            if CAP_PATTERN.match(text):
                cap_n += 1
                cap_map[f"cap-{cap_n}"] = child
        elif child.tag == W_TBL:
            tbl_n += 1
            tbl_map[f"tbl-{tbl_n}"] = child
    return cap_map, tbl_map


def find_h1_above(body, target_elem) -> Optional[int]:
    """找 target_elem 上方最近 H1 段的序数 (按 body 中 H1 出现顺序计 1, 2, ...).

    返回 None 表示 target 上方没 H1.
    """
    children = list(body)
    if target_elem not in children:
        return None
    target_idx = children.index(target_elem)
    h1_count = 0
    last_h1 = None
    for i, c in enumerate(children):
        if i >= target_idx:
            break
        if c.tag == W_P and is_h1(c):
            h1_count += 1
            last_h1 = h1_count
    return last_h1


# ---------- 5 个 op ----------

def op_delete_caption(body, cap_map: dict, op: dict, dry_run: bool) -> dict:
    cid = op["caption_id"]
    elem = cap_map.get(cid)
    if elem is None:
        return {"op": op["op"], "caption_id": cid, "status": "skip", "msg": "未找到"}
    text = get_text(elem)[:60]
    # dry-run 也真改内存 (但不落盘), 让后续 op (尤其 renumber) 看到正确状态
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)
    del cap_map[cid]
    return {"op": op["op"], "caption_id": cid,
            "status": "would-apply" if dry_run else "applied",
            "removed_text": text}


def op_rename_caption(body, cap_map: dict, op: dict, dry_run: bool) -> dict:
    cid = op["caption_id"]
    elem = cap_map.get(cid)
    if elem is None:
        return {"op": op["op"], "caption_id": cid, "status": "skip", "msg": "未找到"}
    new_number = op.get("new_number")
    new_name = op.get("new_name")
    old_text = get_text(elem)
    m = CAP_PATTERN.match(old_text)
    if not m:
        return {"op": op["op"], "caption_id": cid, "status": "skip", "msg": "不是 caption 形态"}
    cur_number = f"表{m.group(1)}-{m.group(2)}"
    cur_name = m.group(3).strip()
    final_number = new_number if new_number else cur_number
    final_name = new_name if new_name is not None else cur_name
    new_text = f"{final_number} {final_name}".strip()
    _replace_caption_text(elem, new_text)
    return {"op": op["op"], "caption_id": cid,
            "status": "would-apply" if dry_run else "applied",
            "old": old_text, "new": new_text}


def _replace_caption_text(p_elem, new_text: str) -> None:
    """把 caption 段所有 run 合并为单一 run, 用第一个 run 的 rPr (保 bold/字号),
    文本设为 new_text."""
    runs = p_elem.findall(W_R)
    if not runs:
        # 没有 run, 直接造一个
        from lxml import etree
        r = etree.SubElement(p_elem, W_R)
        t = etree.SubElement(r, W_T)
        t.text = new_text
        return
    # 保留第一个 run, 删后续 run; 第一个 run 内部清掉所有 <w:t> 后塞新 <w:t>
    first_run = runs[0]
    first_rpr = first_run.find(W_RPR)
    for r in runs[1:]:
        p_elem.remove(r)
    # 清第一个 run 内的非 rPr 子
    for child in list(first_run):
        if child.tag != W_RPR:
            first_run.remove(child)
    from lxml import etree
    t = etree.SubElement(first_run, W_T)
    t.text = new_text
    # 防止空格被压缩
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def op_rename_orphan_tbl(body, cap_map: dict, tbl_map: dict, op: dict, dry_run: bool) -> dict:
    tid = op["tbl_id"]
    tbl_elem = tbl_map.get(tid)
    if tbl_elem is None:
        return {"op": op["op"], "tbl_id": tid, "status": "skip", "msg": "未找到"}
    insert_spec = op["insert_caption_above"]
    number = insert_spec["number"]
    name = insert_spec["name"]
    style = insert_spec.get("style", "zdwp1")
    new_text = f"{number} {name}".strip()

    # 找模板: cap_map 里第一个 style==zdwp1 的 caption (留住格式)
    template_p = None
    for cid, ce in cap_map.items():
        if get_style_id(ce) == style:
            template_p = ce
            break
    if template_p is None and cap_map:
        # 没匹配 style 就拿任意 caption 当模板
        template_p = next(iter(cap_map.values()))

    # 复制模板段, 替换文字 (dry-run 也真改内存, 不落盘)
    if template_p is not None:
        new_p = copy.deepcopy(template_p)
        # 确保 style 正确
        ppr = new_p.find(W_PPR)
        if ppr is not None:
            ps = ppr.find(W_PSTYLE)
            if ps is not None:
                ps.set(W_VAL, style)
        _replace_caption_text(new_p, new_text)
    else:
        # 极端兜底: 造空段
        from lxml import etree
        new_p = etree.SubElement(body, W_P)
        ppr = etree.SubElement(new_p, W_PPR)
        ps = etree.SubElement(ppr, W_PSTYLE)
        ps.set(W_VAL, style)
        r = etree.SubElement(new_p, W_R)
        t = etree.SubElement(r, W_T)
        t.text = new_text
        body.remove(new_p)  # 取出来用 insert

    # 插到 tbl 之前
    parent = tbl_elem.getparent()
    pos = list(parent).index(tbl_elem)
    parent.insert(pos, new_p)
    # 加入 cap_map (派新 ID; 但我们不在中途重新 snapshot, 派 cap-new-N)
    cap_map[f"cap-new-{tid}"] = new_p

    return {"op": op["op"], "tbl_id": tid,
            "status": "would-apply" if dry_run else "applied",
            "inserted_text": new_text}


def op_pair_caption_to_tbl(body, cap_map: dict, tbl_map: dict, op: dict, dry_run: bool) -> dict:
    """把 caption 段块 (caption + 紧邻 '注:' 开头注释段) 移到 tbl 紧前方."""
    cid = op["caption_id"]
    tid = op["tbl_id"]
    cap_elem = cap_map.get(cid)
    tbl_elem = tbl_map.get(tid)
    if cap_elem is None or tbl_elem is None:
        return {"op": op["op"], "caption_id": cid, "tbl_id": tid,
                "status": "skip", "msg": "cap 或 tbl 未找到"}

    parent = cap_elem.getparent()
    cap_pos = list(parent).index(cap_elem)
    # 收集 cap + 紧邻 '注:' 段
    block = [cap_elem]
    for sib in list(parent)[cap_pos + 1:]:
        if sib.tag == W_P:
            sib_text = get_text(sib)
            if sib_text.startswith(("注:", "注：", "Note", "note")):
                block.append(sib)
            else:
                break
        else:
            break

    # detach 整块 (dry-run 也真改内存, 不落盘)
    for e in block:
        parent.remove(e)
    # insert before tbl
    tbl_parent = tbl_elem.getparent()
    tbl_pos = list(tbl_parent).index(tbl_elem)
    for i, e in enumerate(block):
        tbl_parent.insert(tbl_pos + i, e)
    return {"op": op["op"], "caption_id": cid, "tbl_id": tid,
            "status": "would-apply" if dry_run else "applied",
            "block_size": len(block)}


def op_renumber_all_tables(body, dry_run: bool) -> dict:
    """按 body 物理顺序 + H1 章节, 重编 "表 X-Y"."""
    children = list(body)
    chapter_counters = {}  # chapter → next_seq
    current_chapter = 0
    renames = []
    for child in children:
        if child.tag == W_P:
            if is_h1(child):
                current_chapter += 1
                continue
            text = get_text(child)
            m = CAP_PATTERN.match(text)
            if m:
                seq = chapter_counters.get(current_chapter, 0) + 1
                chapter_counters[current_chapter] = seq
                old_text = text
                name = m.group(3).strip()
                new_text = f"表{current_chapter}-{seq} {name}".strip()
                if old_text.strip() != new_text:
                    renames.append({"old": old_text[:80], "new": new_text[:80]})
                    _replace_caption_text(child, new_text)
    return {"op": "renumber-all-tables",
            "status": "would-apply" if dry_run else "applied",
            "renames": renames, "count": len(renames)}


# ---------- 主流程 ----------

OP_DISPATCH = {
    "delete-caption": op_delete_caption,
    "rename-caption": op_rename_caption,
}


def execute(doc, decision: dict, dry_run: bool) -> list[dict]:
    body = doc.element.body
    cap_map, tbl_map = snapshot_ids(body)
    results = []

    # 操作顺序: rename → delete → pair → rename-orphan → renumber
    # (renumber 必须最后, 因 caption 物理顺序变了)
    ops = decision.get("operations", [])
    order = ["rename-caption", "delete-caption", "pair-caption-to-tbl",
             "rename-orphan-tbl", "renumber-all-tables"]
    ops_sorted = sorted(ops, key=lambda o: order.index(o["op"]) if o["op"] in order else 99)

    for op in ops_sorted:
        kind = op["op"]
        if kind == "delete-caption":
            r = op_delete_caption(body, cap_map, op, dry_run)
        elif kind == "rename-caption":
            r = op_rename_caption(body, cap_map, op, dry_run)
        elif kind == "rename-orphan-tbl":
            r = op_rename_orphan_tbl(body, cap_map, tbl_map, op, dry_run)
        elif kind == "pair-caption-to-tbl":
            r = op_pair_caption_to_tbl(body, cap_map, tbl_map, op, dry_run)
        elif kind == "renumber-all-tables":
            r = op_renumber_all_tables(body, dry_run)
        else:
            r = {"op": kind, "status": "skip", "msg": "unknown op"}
        results.append(r)

    return results


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("docx", help="输入 docx 路径")
    ap.add_argument("--decision", required=True, help="decision JSON 路径")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    ap.add_argument("--report", help="结果 JSON 输出路径")
    args = ap.parse_args(argv)

    src = Path(args.docx)
    if not src.exists():
        print(f"[error] 找不到 {src}", file=sys.stderr)
        return 2

    # lsof 自检 (改前)
    if not args.dry_run:
        lsof = lsof_check(src)
        if lsof:
            print(f"[error] docx 被进程占用 (关 Word/WPS 后重试):\n{lsof}", file=sys.stderr)
            return 3

    decision = json.loads(Path(args.decision).read_text(encoding="utf-8"))
    # doctools v1 schema 校验 (best-effort)
    try:
        from lib.schemas import validate as _validate_schema
        _err = _validate_schema(decision, "decision")
        if _err:
            print(f"[error] decision schema 校验失败 (v1): {_err}", file=sys.stderr)
            return 2
    except Exception:
        pass
    backup_path = None
    if not args.dry_run and not args.no_backup:
        today = date.today().isoformat()
        n = 1
        while True:
            cand = src.with_name(f"{src.stem}.bak-{n}-{today}{src.suffix}")
            if not cand.exists():
                break
            n += 1
        shutil.copy2(src, cand)
        backup_path = cand
        print(f"[backup] {cand.name}")

    doc = Document(str(src))
    results = execute(doc, decision, args.dry_run)
    if not args.dry_run:
        doc.save(str(src))
        print(f"[saved] {src}")

    print(f"\n{'='*78}")
    print(f"{'dry-run' if args.dry_run else 'apply'} 结果 ({len(results)} ops):")
    for r in results:
        head = f"  [{r.get('status','?')}] {r.get('op','?')}"
        rest = " ".join(f"{k}={v}" for k, v in r.items()
                        if k not in ("status", "op", "renames"))
        print(f"{head} {rest}")
        if "renames" in r and r["renames"]:
            for rn in r["renames"]:
                print(f"      {rn['old']}  →  {rn['new']}")

    out = {
        "docx_path": str(src),
        "backup": str(backup_path) if backup_path else None,
        "dry_run": args.dry_run,
        "results": results,
    }
    if args.report:
        Path(args.report).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report).write_text(json.dumps(out, indent=2, ensure_ascii=False))
        print(f"\n[report] {args.report}")

    return 0


# ---------------- pipeline adapter ----------------
def apply(doc, args=None) -> dict:
    decision_path = getattr(args, "pair_decision", None) if args else None
    if not decision_path:
        return {"changed": 0, "skipped": "no pair_decision in args"}
    dry = bool(getattr(args, "dry_run", False)) if args else False
    try:
        decision = json.loads(Path(decision_path).read_text(encoding="utf-8"))
    except Exception as exc:
        return {"error": f"decision read failed: {exc}"}
    results = execute(doc, decision, dry)
    ok = sum(1 for r in results if r.get("status") == "ok")
    return {"changed": ok, "results_count": len(results), "ok": ok}


if __name__ == "__main__":
    sys.exit(main())
