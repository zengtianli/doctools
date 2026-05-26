#!/usr/bin/env python3
# distilled from qual-supply/scripts/number_captions.py (2026-05-25 W1)
"""
number_captions.py — 给 docx 表名/图名段补章节序号 "表 X-Y" / "图 X-Y"

单功能描述:
  扫 docx body 全部 paragraph, 识别表名段 (紧邻 <w:tbl>) 和图名段 (紧邻含
  drawing/pict 的段), 段开头 prepend "表 X-Y　" 或 "图 X-Y　" (中文全角空格)。
  X = 当前 H1 章号, Y = 章内表/图递增序号 (跨 H1 重置)。

触发场景:
  整合后子报告 docx 表/图标题段缺章节编号, 验收前补齐。集团 /docx skill 无此能力。

CLI:
  python3 number_captions.py <docx_path> [--dry-run] [--no-backup] [--report <json>]

启发规则:
  表名候选:
    - 段非空, 段长 < 60 字, 不以 [。；，.;,] 结尾
    - 已有 "^表\\s*\\d+[-.]\\d+" 前缀 → 跳过
    - 紧邻 <w:tbl>: 后 1-2 个 body 元素中存在 tbl (允许中间隔空段)
    - 或 pStyle ∈ {zdwp表名, Caption, caption, 表题} 直接接受
  图名候选:
    - 段非空, 段长 < 60 字, 不以 [。；，.;,] 结尾
    - 已有 "^图\\s*\\d+[-.]\\d+" 前缀 → 跳过
    - 前 1-2 个或后 1-2 个 body 元素中有段含 drawing/pict
    - 或文本含图题特征关键词 (示意图/分布图/结构图/流程图/对比图/平面图)

H1 章节识别:
  - 中文数字开头: ^([一二三四五六七八九十]+)、 → 映射 1-10
  - 阿拉伯数字: ^(\\d+)[\\s　、.] (短段 < 30)

修改方式 (run 级):
  - 找段第一个含 <w:t> 的 run, 在其首个 <w:t> 文本前 prepend 编号
  - 不动其他 run, 保留 bold/字号

不许做:
  - 段级 paragraph.text = ... (会丢 run 格式)
  - 改样式 / 改表/drawing 本身
  - 跨 H1 不重置编号
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

# ---------- 常量 ----------

CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
          "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
          "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15}

CAPTION_STYLES = {"zdwp表名", "Caption", "caption", "表题", "图题"}

FIG_KEYWORDS = ("示意图", "分布图", "结构图", "流程图", "对比图",
                "平面图", "布局图", "线路图", "断面图")

# 段开头已有编号则跳过
RE_HAS_TABLE_NUM = re.compile(r"^\s*表\s*\d+\s*[-.–—]\s*\d+")
RE_HAS_FIG_NUM = re.compile(r"^\s*图\s*\d+\s*[-.–—]\s*\d+")

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
        return CN_NUM.get(m.group(1))
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


def lsof_check(path: Path) -> bool:
    """是否被打开。返回 True = 被占用。"""
    try:
        r = subprocess.run(["lsof", str(path)], capture_output=True, text=True, timeout=5)
        return bool(r.stdout.strip())
    except Exception:
        return False


def pick_backup_path(src: Path) -> Path:
    today = _dt.date.today().isoformat()
    parent = src.parent
    stem = src.stem
    suffix = src.suffix
    n = 1
    while True:
        cand = parent / f"{stem}.bak-{n}-{today}{suffix}"
        if not cand.exists():
            return cand
        n += 1


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
    if lsof_check(docx_path):
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


def apply(doc, args=None) -> dict:
    """pipeline doc-based"""
    dry = bool(getattr(args, "dry_run", False)) if args else False
    return _process_doc(doc, dry)


def main():
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


if __name__ == "__main__":
    main()
