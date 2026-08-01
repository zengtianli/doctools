#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""caption.py — 题注家族二合一（2026-07-31 家族折叠）

子命令 ↔ 原脚本（函数体逐字搬移；模块级 apply/main 改名 _<sub> 后缀，
lsof_check 两版语义不同各自保留（number 版返回 bool、pair 版返回 str），
.bak-N-日期 备份路径生成收敛到 _cli_common（机制同构，文案不变））：

    number  ← number_captions.py
    pair    ← pair_table_captions.py

各子命令 CLI 与原独立脚本逐字一致：python3 sub/caption.py <sub> <docx> …。
退役原件在 ~/.Trash/consolidation-20260731/caption/（含 MANIFEST.md）。
"""
from __future__ import annotations

# ── surgical 收口：python-docx 存盘只重写点名的部件（炸开面 60→1）─────────────
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
import docx_safe_save  # noqa: E402,F401  详见 lib/docx_safe_save.py
from cn_number import cn_to_int  # noqa: E402,F401  中文数字 SSOT
import caption_re  # noqa: E402  题注判据 SSOT

# sub/ 自身进 sys.path —— docx_cli 的 _dispatch 用 spec_from_file_location 加载,
# 不带脚本目录, 裸 import _cli_common 会 ImportError (append 不是 insert(0))
_sys.path.append(str(_Path(__file__).resolve().parent))
import _cli_common as _cc  # noqa: E402  备份路径机制 SSOT

import argparse  # noqa: E402
import copy  # noqa: E402
import datetime as _dt  # noqa: E402,F401  （原 number_captions 顶部导入，保留）
import json  # noqa: E402
import re  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402
import sys  # noqa: E402
from datetime import date  # noqa: E402,F401
from pathlib import Path  # noqa: E402
from typing import Optional  # noqa: E402

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402


# ══════════ number ← number_captions.py ══════════

# 中文数字：本文件原有一张「只到十五」的查表（CN_NUM），2026-08-01 换成
# lib/cn_number.cn_to_int（见文件顶部 import）。行为变化：第 16 章往后的中文章号
# 过去解析失败 → 章计数器不切换 → 表图继续按上一章编（表15-7、表15-8…）；现在正确切章。

# 样式名/关键词判据 2026-08-01 搬进 lib/caption_re.py。**没有取并集** —— 这里是
# 精确集合、audit 是关键词包含、shape_contract 是正则，三套各有对方缺的样式名，
# 合并 = 给本命令（写盘）扩范围，得单独一轮做。见 caption_re 模块 docstring。
CAPTION_STYLES = caption_re.CAPTION_STYLES_EXACT

# 本文件这 9 词是 styles.py 那份 13 词的**真子集**，故合成「核心 + 扩展」零行为变化。
FIG_KEYWORDS = caption_re.FIG_KEYWORDS_CORE

# 段开头已有编号则跳过。2026-08-01 判据下沉 lib/caption_re。行为变化：短横补齐五种，
# 旧的漏 U+2011/全角－ → `表3－1 xxx` 被当成「还没编号」再编一次（双重编号同一坑）。
RE_HAS_TABLE_NUM = caption_re.pattern(caption_re.HAS_NUM_TABLE)
RE_HAS_FIG_NUM = caption_re.pattern(caption_re.HAS_NUM_FIG)

# 句末标点 (中英文)
RE_SENTENCE_END = re.compile(r"[。；，.;,]$")

# H1 chapter heuristics
RE_CN_CHAPTER = re.compile(r"^([一二三四五六七八九十]+)、")
RE_AR_CHAPTER = re.compile(r"^(\d+)[\s　、.]")


# ---------- 辅助 ----------

def get_p_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def get_p_style(p) -> str | None:
    pStyle = p.find(".//" + qn("w:pStyle"))
    return pStyle.get(qn("w:val")) if pStyle is not None else None


def p_has_drawing(p) -> bool:
    return (next(p.iter(qn("w:drawing")), None) is not None
            or next(p.iter(qn("w:pict")), None) is not None)


def is_empty_p(p) -> bool:
    return get_p_text(p).strip() == "" and not p_has_drawing(p)


def parse_chapter(text: str) -> int | None:
    """从段文本解析 H1 章号; 失败返回 None。"""
    t = text.strip()
    m = RE_CN_CHAPTER.match(t)
    if m:
        return cn_to_int(m.group(1))
    m = RE_AR_CHAPTER.match(t)
    if m and len(t) < 30:
        return int(m.group(1))
    return None


def is_h1_chapter(text: str) -> bool:
    """是否是 H1 章节段 (短 + 章节编号开头)。"""
    t = text.strip()
    if len(t) > 30 or len(t) < 2:
        return False
    if RE_CN_CHAPTER.match(t):
        return True
    return False


def next_nonempty_idx(elements, start: int, lookahead: int = 3) -> int | None:
    """返回 start 之后 lookahead 范围内首个非空元素索引。"""
    for j in range(start, min(start + lookahead + 1, len(elements))):
        el = elements[j]
        tag = el.tag.split("}")[-1]
        if tag == "tbl":
            return j
        if tag == "p" and not is_empty_p(el):
            return j
        if tag == "p" and p_has_drawing(el):
            return j
    return None


def has_nearby_table(elements, i: int) -> bool:
    """段 i 后 1-2 个 body 元素中是否有 tbl (允许隔 1 空段)。"""
    for offset in (1, 2):
        j = i + offset
        if j >= len(elements):
            break
        el = elements[j]
        if el.tag.split("}")[-1] == "tbl":
            return True
        if el.tag.split("}")[-1] == "p" and not is_empty_p(el):
            return False
    return False


def has_nearby_drawing(elements, i: int) -> bool:
    """段 i 前后 1-3 元素中是否有含 drawing/pict 的段。"""
    for offset in (-1, -2, -3, 1, 2, 3):
        j = i + offset
        if not (0 <= j < len(elements)):
            continue
        el = elements[j]
        if el.tag.split("}")[-1] != "p":
            continue
        if p_has_drawing(el):
            # 中间不能隔实文本段 (避免误抓)
            step = 1 if offset > 0 else -1
            blocked = False
            for k in range(i + step, j, step):
                kel = elements[k]
                if kel.tag.split("}")[-1] == "p" and not is_empty_p(kel) and not p_has_drawing(kel):
                    blocked = True
                    break
            if not blocked:
                return True
    return False


def prepend_run_text(p, prefix: str) -> bool:
    """在段第一个含 <w:t> 的 run 的首个 <w:t> 前 prepend 文本。成功返回 True。"""
    for r in p.iter(qn("w:r")):
        t_elems = list(r.iter(qn("w:t")))
        if not t_elems:
            continue
        first_t = t_elems[0]
        old = first_t.text or ""
        first_t.text = prefix + old
        # 保 xml:space="preserve" 防全角空格被吞
        first_t.set(qn("xml:space"), "preserve")
        return True
    return False


def lsof_check_bool(path: Path) -> bool:
    """是否被打开。返回 True = 被占用。"""
    try:
        r = subprocess.run(["lsof", str(path)], capture_output=True, text=True, timeout=5)
        return bool(r.stdout.strip())
    except Exception:
        return False


def pick_backup_path(src: Path) -> Path:
    return _cc.find_next_backup(src)


# ---------- 主流程 ----------

def _process_doc(doc, dry_run: bool):
    body = doc.element.body
    elements = list(body.iterchildren())

    chapter = 0  # 当前 H1 章号
    tbl_y = 0    # 章内表序
    fig_y = 0    # 章内图序

    numbered: list[dict] = []
    manual_review: list[dict] = []
    chapters_detected: list[int] = []

    for i, el in enumerate(elements):
        tag = el.tag.split("}")[-1]
        if tag != "p":
            continue
        text = get_p_text(el)
        text_strip = text.strip()
        if not text_strip:
            continue

        # H1 章节检测
        if is_h1_chapter(text_strip):
            ch = parse_chapter(text_strip)
            if ch is not None and ch != chapter:
                chapter = ch
                tbl_y = 0
                fig_y = 0
                if ch not in chapters_detected:
                    chapters_detected.append(ch)
                continue

        # 已有编号 → 跳过
        if RE_HAS_TABLE_NUM.match(text_strip) or RE_HAS_FIG_NUM.match(text_strip):
            continue

        # caption 候选门槛: 短段 + 无标点结尾
        if len(text_strip) >= 60 or RE_SENTENCE_END.search(text_strip):
            continue

        style = get_p_style(el)

        # === 表名识别 ===
        is_table_cap = False
        if style in CAPTION_STYLES:
            # caption 类样式 + 紧邻 tbl
            if has_nearby_table(elements, i):
                is_table_cap = True
        elif has_nearby_table(elements, i):
            # 紧邻 tbl 的短段
            is_table_cap = True

        # === 图名识别 ===
        is_fig_cap = False
        if not is_table_cap:
            if has_nearby_drawing(elements, i):
                is_fig_cap = True
            elif any(kw in text_strip for kw in FIG_KEYWORDS) and len(text_strip) < 40:
                # 关键词强信号 (即使无邻接 drawing, 可能图被放在远处)
                # 仅当上下文没有 tbl 邻接时记 manual
                manual_review.append({
                    "idx": i,
                    "reason": "fig-keyword-no-adjacent-drawing",
                    "text_snippet": text_strip[:60],
                })
                continue

        if not (is_table_cap or is_fig_cap):
            continue

        # 章号未识别 → manual
        if chapter == 0:
            manual_review.append({
                "idx": i,
                "reason": "no-chapter-context",
                "text_snippet": text_strip[:60],
            })
            continue

        if is_table_cap:
            tbl_y += 1
            prefix = f"表 {chapter}-{tbl_y}　"
            cap_type = "table"
            number = f"表 {chapter}-{tbl_y}"
        else:
            fig_y += 1
            prefix = f"图 {chapter}-{fig_y}　"
            cap_type = "figure"
            number = f"图 {chapter}-{fig_y}"

        if not dry_run:
            ok = prepend_run_text(el, prefix)
            if not ok:
                manual_review.append({
                    "idx": i,
                    "reason": "no-run-with-text-to-prepend",
                    "text_snippet": text_strip[:60],
                })
                # 回滚计数
                if cap_type == "table":
                    tbl_y -= 1
                else:
                    fig_y -= 1
                continue

        numbered.append({
            "idx": i,
            "type": cap_type,
            "number": number,
            "text_after": prefix + text_strip,
        })

    summary = {
        "tables_numbered": sum(1 for x in numbered if x["type"] == "table"),
        "figures_numbered": sum(1 for x in numbered if x["type"] == "figure"),
        "manual_review_count": len(manual_review),
        "chapters_detected": chapters_detected,
    }
    return {
        "changed": summary["tables_numbered"] + summary["figures_numbered"],
        "summary": summary,
        "numbered": numbered,
        "manual_review": manual_review,
    }


def process(docx_path: Path, dry_run: bool, do_backup: bool, report_json: Path | None):
    if lsof_check_bool(docx_path):
        print(f"[ABORT] {docx_path} 被 Word/WPS 打开, 请先关闭。", file=sys.stderr)
        sys.exit(2)
    doc = Document(str(docx_path))
    result = _process_doc(doc, dry_run)
    numbered = result["numbered"]
    manual_review = result["manual_review"]
    summary = result["summary"]
    print(f"[summary] {json.dumps(summary, ensure_ascii=False)}")
    for entry in numbered:
        print(f"  +{entry['type']:7} idx={entry['idx']:3} → {entry['text_after']}")
    if manual_review:
        print("[manual_review]")
        for m in manual_review:
            print(f"  idx={m['idx']:3} reason={m['reason']:30} '{m['text_snippet']}'")
    if dry_run:
        print("[dry-run] no write")
    else:
        if do_backup:
            bak = pick_backup_path(docx_path)
            shutil.copy2(docx_path, bak)
            print(f"[backup] {bak}")
        doc.save(str(docx_path))
        print(f"[saved] {docx_path}")
    if report_json:
        report = {"summary": summary, "numbered": numbered, "manual_review": manual_review}
        report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[report] {report_json}")


def apply_number(doc, args=None) -> dict:
    """pipeline doc-based"""
    dry = bool(getattr(args, "dry_run", False)) if args else False
    return _process_doc(doc, dry)


def main_number():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1] if __doc__ else "")
    ap.add_argument("docx_path", type=Path)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    ap.add_argument("--report", type=Path, default=None, help="写 JSON 报告路径")
    args = ap.parse_args()

    if not args.docx_path.exists():
        print(f"[err] not found: {args.docx_path}", file=sys.stderr)
        sys.exit(1)

    process(args.docx_path, args.dry_run, not args.no_backup, args.report)


# ══════════ pair ← pair_table_captions.py ══════════

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{W_NS}}}p"
W_TBL = f"{{{W_NS}}}tbl"
W_R = f"{{{W_NS}}}r"
W_T = f"{{{W_NS}}}t"
W_PPR = f"{{{W_NS}}}pPr"
W_PSTYLE = f"{{{W_NS}}}pStyle"
W_RPR = f"{{{W_NS}}}rPr"
W_VAL = f"{{{W_NS}}}val"

# 2026-08-01：本处原是一份**与 audit.py 同名、语义相同、章号却只写 `(\d+)` 的更窄
# 拷贝** —— 上游 `audit table-pairing` 产出 `表3.1-1`，本处匹配不上，`caption pair`
# 对 --cn-section 文档的 decision.json 条目静默 no-op。现在 import 同一个 spec 对象。
CAP_SPEC = caption_re.TABLE_CAPTION_LINE
CAP_PATTERN = caption_re.pattern(CAP_SPEC)


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
    m = caption_re.parse(old_text, CAP_SPEC)
    if not m:
        return {"op": op["op"], "caption_id": cid, "status": "skip", "msg": "不是 caption 形态"}
    cur_number = f"表{m.section}-{m.seq}"
    cur_name = old_text[m.end:].strip()
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
            m = caption_re.parse(text, CAP_SPEC)
            if m:
                seq = chapter_counters.get(current_chapter, 0) + 1
                chapter_counters[current_chapter] = seq
                old_text = text
                name = text[m.end:].strip()
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


def main_pair(argv: list[str] | None = None) -> int:
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
        cand = _cc.make_backup(src)
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
def apply_pair(doc, args=None) -> dict:
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


# ──────────────────────────── 家族入口（子命令分发）────────────────────────────

SUBCOMMANDS = {
    "number": main_number,
    "pair": main_pair,
}


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in ("-h", "--help"):
        print("usage: caption.py {" + ",".join(SUBCOMMANDS) + "} <args…>\n"
              "每个子命令的参数与原独立脚本逐字一致：caption.py <sub> --help 查看。")
        return 0 if args else 2
    sub, rest = args[0], args[1:]
    fn = SUBCOMMANDS.get(sub)
    if fn is None:
        print(f"[caption] unknown subcommand: {sub!r}; choices={list(SUBCOMMANDS)}",
              file=sys.stderr)
        return 2
    saved = sys.argv[:]
    sys.argv = [sys.argv[0]] + rest
    try:
        rc = fn()
        return int(rc) if isinstance(rc, int) else 0
    finally:
        sys.argv = saved


if __name__ == "__main__":
    sys.exit(main())
