#!/usr/bin/env python3
"""health.py — docx health diagnose / fix / full 编排层

8 病种诊断 (ThreadPoolExecutor 并发), 委托已有 sub/*.py 脚本，新增 2 个原生 check:
  heading-level-skew  (新写 ~15 行)
  heading-gap         (新写 ~20 行)
其余 6 种委托现有 audit_* / strip_* / apply_body_styles / renumber_headings。

CLI (via docx_cli.py health <subcommand>):
  health diagnose <docx> [--checks all|<list>] [--report path.json] [--html] [--workers N]
  health fix      <docx> [--auto <list>] [--plan report.json] [--dry-run] [--backup]
  health full     <docx> [--html] [--dry-run]

返回码: 0=全健康 / 1=有 warning / 2=有 error/High
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

from ._dispatch import exec_script, get_or_add_group, get_or_add_subparsers
from .docx_health_render_html import render_rich_html, render_simple_html

# ─── 常量 ────────────────────────────────────────────────────────────────────

SEVERITY = {
    "heading-level-skew":       "High",
    "heading-gap":              "Med",
    "caption-outline-pollution":"Med",
    "revision-tracking-residue":"High",
    "field-not-frozen":         "Med",
    "body-style-mess":          "Low",
    "duplicate-figure-numbers": "High",
    "heading-number-stale":     "Med",
}

SAFE_FIX = {
    "caption-outline-pollution": True,
    "revision-tracking-residue": True,
    "field-not-frozen":          True,
    "body-style-mess":           True,
    "heading-number-stale":      True,
    "heading-level-skew":        True,   # auto only when coverage >= 0.8
    "heading-gap":               False,
    "duplicate-figure-numbers":  False,
}

ALL_CHECKS = list(SEVERITY.keys())

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ─── 原生 check 函数 ──────────────────────────────────────────────────────────

_NUMBER_PREFIX = re.compile(r"^\s*(\d+(?:\.\d+)*)[\s\.、]")


def _text_prefix_depth(text: str) -> int | None:
    """抓数字前缀深度: '1.2.2 xxx' → 3, '2 河湖...' → 1, 无前缀 → None."""
    m = _NUMBER_PREFIX.match(text)
    if not m:
        return None
    return m.group(1).count(".") + 1


def _style_to_level_1based(style_name_or_id: str) -> int | None:
    """1-based level: 'Heading 4' → 4, 'heading 2' → 2, bare '4' → 4."""
    s = style_name_or_id.strip()
    m = re.search(r"[Hh]eading\s*(\d+)", s)
    if m:
        return int(m.group(1))
    if re.fullmatch(r"\d", s):
        return int(s)
    return None


def _para_text(para_xml: str) -> str:
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para_xml))


def _resolve_style_to_heading_level(styles_xml: str) -> dict[str, int]:
    """构 styleId → heading level(1-based) map.

    覆盖中文重命名（styleId='2' name='heading 2'）和原生 styleId='Heading1' 两种。
    """
    out: dict[str, int] = {}
    for m in re.finditer(r'<w:style[^>]+w:styleId="([^"]+)"', styles_xml):
        sid = m.group(1)
        block_start = m.start()
        block_end = styles_xml.find("</w:style>", block_start)
        if block_end < 0:
            continue
        block = styles_xml[block_start:block_end]
        name_m = re.search(r'<w:name w:val="([^"]+)"', block)
        name = name_m.group(1) if name_m else ""
        lvl = _style_to_level_1based(name) or _style_to_level_1based(sid)
        if lvl is not None:
            out[sid] = lvl
    return out


def check_heading_level_skew(doc_path: Path) -> dict:
    """检出：标题段「文本数字前缀深度」与「样式隐含级别」系统性偏差（磐安症状）。

    主算法：delta = style_level(1-based) − text_prefix_depth
      磐安：'1.2.2'(depth=3) 挂 Heading 4 → delta=+1（style 偏高 1 级）
      '2 河湖...'(depth=1) 挂 Heading 2 → delta=+1
    无数字前缀的标题（前言/结论/附录）跳过不计入。
    """
    try:
        with zipfile.ZipFile(doc_path) as z:
            with z.open("word/document.xml") as f:
                doc_xml = f.read().decode("utf-8")
            with z.open("word/styles.xml") as f:
                styles_xml = f.read().decode("utf-8")
    except Exception as e:
        return {"found": False, "error": str(e)}

    style_lvl_map = _resolve_style_to_heading_level(styles_xml)

    deltas: list[int] = []
    samples: list[dict] = []

    for para_xml in re.findall(r"<w:p[\s>].*?</w:p>", doc_xml, re.DOTALL):
        sm = re.search(r'<w:pStyle w:val="([^"]+)"', para_xml)
        if not sm:
            continue
        sid = sm.group(1)
        style_lvl = style_lvl_map.get(sid) or _style_to_level_1based(sid)
        if style_lvl is None:
            continue
        text = _para_text(para_xml)
        text_depth = _text_prefix_depth(text)
        if text_depth is None:
            continue
        delta = style_lvl - text_depth
        deltas.append(delta)
        if len(samples) < 5:
            samples.append({
                "text": text[:30],
                "style_lvl": style_lvl,
                "text_depth": text_depth,
                "delta": delta,
            })

    if not deltas:
        return {"found": False, "reason": "no headings with numeric prefix"}

    counter = Counter(deltas)
    dominant_delta, count = counter.most_common(1)[0]
    coverage = count / len(deltas)

    # coverage ≥ 0.7 = 系统性偏差（容忍少数前言/结论/附录段 delta=0）
    if coverage >= 0.7 and dominant_delta != 0:
        direction = "promote" if dominant_delta > 0 else "demote"
        # auto-fix 仍需 ≥ 0.8 严格阈值，0.7-0.8 段降为 plan-required
        safe = coverage >= 0.8
        return {
            "found": True,
            "delta": dominant_delta,
            "coverage": round(coverage, 3),
            "affected_count": count,
            "total_checked": len(deltas),
            "samples": samples,
            "deltas_distribution": dict(counter),
            "fix_hint": f"整体 {direction} {abs(dominant_delta)} 级 (style 比文本前缀{'高' if dominant_delta>0 else '低'} {abs(dominant_delta)})",
            "safe_fix": safe,
        }
    return {
        "found": False,
        "dominant_delta": dominant_delta,
        "coverage": round(coverage, 3),
        "deltas_distribution": dict(counter),
        "samples": samples,
    }


def check_outline_lvl_mismatch(doc_path: Path) -> dict:
    """副 check：outlineLvl XML 属性 vs style 隐含级别对账（旧逻辑保留）。

    两层任一命中即报：① paragraph 级 outlineLvl override；② styles.xml 级 outlineLvl 与 style 名不一致。
    """
    def style_to_implied_0based(s: str) -> int | None:
        lvl = _style_to_level_1based(s)
        return lvl - 1 if lvl is not None else None

    try:
        with zipfile.ZipFile(doc_path) as z:
            doc_xml = z.open("word/document.xml").read().decode("utf-8")
            styles_xml = z.open("word/styles.xml").read().decode("utf-8")
    except Exception as e:
        return {"found": False, "error": str(e)}

    # Layer 1: paragraph-level
    para_deltas: list[int] = []
    for para_xml in re.findall(r"<w:p[\s>].*?</w:p>", doc_xml, re.DOTALL):
        sm = re.search(r'<w:pStyle w:val="([^"]+)"', para_xml)
        if not sm:
            continue
        implied = style_to_implied_0based(sm.group(1))
        if implied is None:
            continue
        ppr_m = re.search(r"<w:pPr>(.*?)</w:pPr>", para_xml, re.DOTALL)
        if not ppr_m:
            continue
        ol_m = re.search(r'<w:outlineLvl w:val="(\d+)"', ppr_m.group(1))
        if not ol_m:
            continue
        para_deltas.append(int(ol_m.group(1)) - implied)

    if para_deltas:
        c = Counter(para_deltas)
        dd, cnt = c.most_common(1)[0]
        cov = cnt / len(para_deltas)
        if cov >= 0.8 and dd != 0:
            return {"found": True, "layer": "paragraph", "delta": dd,
                    "coverage": round(cov, 3), "affected_count": cnt}

    # Layer 2: style-level
    style_deltas: list[int] = []
    for m in re.finditer(r'<w:style[^>]+>', styles_xml):
        bs = m.start()
        be = styles_xml.find("</w:style>", bs)
        if be < 0:
            continue
        block = styles_xml[bs:be]
        name_m = re.search(r'<w:name w:val="([^"]+)"', block)
        implied = style_to_implied_0based(name_m.group(1) if name_m else "")
        if implied is None:
            continue
        ol_m = re.search(r'<w:outlineLvl w:val="(\d+)"', block)
        if not ol_m:
            continue
        style_deltas.append(int(ol_m.group(1)) - implied)

    if style_deltas:
        c = Counter(style_deltas)
        dd, cnt = c.most_common(1)[0]
        cov = cnt / len(style_deltas)
        if cov >= 0.8 and dd != 0:
            return {"found": True, "layer": "style", "delta": dd,
                    "coverage": round(cov, 3), "affected_count": cnt}

    return {"found": False}


def check_heading_gap(doc_path: Path) -> dict:
    """检出：标题级别跳级（H1 → H3 无中间 H2，H2 → H4 无中间 H3 等）。"""
    HEADING_RE = re.compile(r"[Hh]eading\s*(\d+)|^(\d)$")

    paragraphs: list[int] = []
    try:
        with zipfile.ZipFile(doc_path) as z:
            with z.open("word/document.xml") as f:
                doc_xml = f.read().decode("utf-8")
    except Exception as e:
        return {"found": False, "error": str(e)}

    for para_xml in re.findall(r"<w:p[\s>].*?</w:p>", doc_xml, re.DOTALL):
        sm = re.search(r'<w:pStyle w:val="([^"]+)"', para_xml)
        if not sm:
            continue
        m = HEADING_RE.search(sm.group(1))
        if not m:
            continue
        lvl_str = m.group(1) or m.group(2)
        if lvl_str:
            paragraphs.append(int(lvl_str))

    gaps: list[dict] = []
    for i in range(1, len(paragraphs)):
        prev, curr = paragraphs[i - 1], paragraphs[i]
        if curr > prev + 1:
            gaps.append({"para_idx": i, "from_level": prev, "to_level": curr, "skip": curr - prev - 1})

    if gaps:
        return {
            "found": True,
            "gap_count": len(gaps),
            "gaps": gaps[:20],  # cap at 20 samples
            "safe_fix": False,
            "fix_hint": "需人工判断是压缩标题还是补中间级别",
        }
    return {"found": False}


# ─── 委托 check 函数（调用现有 audit_* / strip_* sub scripts） ────────────────

def _run_script_json(script_name: str, argv: list[str], report_path: Path) -> dict:
    """执行 sub/<script_name>.py 并读回 JSON 报告。"""
    rc = exec_script(script_name, argv + ["--report", str(report_path)])
    if report_path.exists():
        try:
            return json.loads(report_path.read_text("utf-8"))
        except Exception:
            pass
    return {"_rc": rc, "_script": script_name}


def check_caption_outline(doc_path: Path, tmp_dir: Path) -> dict:
    rpt = tmp_dir / "caption_outline.json"
    data = _run_script_json("audit_caption_outline", [str(doc_path)], rpt)
    polluted = data.get("polluted_count", data.get("total_polluted", 0))
    if isinstance(polluted, int) and polluted > 0:
        return {"found": True, "polluted_count": polluted, "safe_fix": True,
                "fix_hint": "strip outlinelvl-captions"}
    # also check via keys
    if data.get("captions_with_outlinelvl") or data.get("found"):
        return {"found": True, "detail": data, "safe_fix": True}
    return {"found": False, "detail": data}


def check_revision_tracking(doc_path: Path) -> dict:
    """干跑 strip_revisions 检查残留 ins/del 数量。"""
    try:
        with zipfile.ZipFile(doc_path) as z:
            with z.open("word/document.xml") as f:
                doc_xml = f.read().decode("utf-8")
    except Exception as e:
        return {"found": False, "error": str(e)}
    ins_count = len(re.findall(r"<w:ins\b", doc_xml))
    del_count = len(re.findall(r"<w:del\b", doc_xml))
    total = ins_count + del_count
    if total > 0:
        return {"found": True, "ins_count": ins_count, "del_count": del_count,
                "total": total, "safe_fix": True,
                "fix_hint": "strip revisions"}
    return {"found": False}


def check_field_not_frozen(doc_path: Path, tmp_dir: Path) -> dict:
    rpt = tmp_dir / "fields.json"
    data = _run_script_json("audit_word_fields", [str(doc_path)], rpt)
    field_count = data.get("total_complex_fields", 0) + data.get("total_simple_fields", 0)
    unfrozen_types = data.get("field_type_counts", {})
    # Any TOC/PAGEREF/SEQ/REF = likely unfrozen
    hot_types = {k: v for k, v in unfrozen_types.items()
                 if k in ("TOC", "PAGEREF", "SEQ", "REF", "STYLEREF", "HYPERLINK")}
    if hot_types:
        return {"found": True, "field_types": hot_types, "total_fields": field_count,
                "safe_fix": True, "fix_hint": "freeze all-fields"}
    return {"found": False, "total_fields": field_count}


def check_body_style_mess(doc_path: Path) -> dict:
    """扫正文段中真正 BODY_LIKE (Normal/正文系) 段的直接 rPr/pPr 杂乱属性。

    Bug 修复 (2026-05-26 · W-fix-body-detector):
      原实现把 Heading 1-9 / Title / TOC* / 中文 N 级标题 / 图名 / 表格标题
      等 protected 段也算 mess (因为它们 style_id != "normal" 又不 startswith
      "heading")。结果跑 `style body` 修完(只动 Normal 段)detector 还报 14
      mess —— 因为它扫的是 protected 段,与 fix 范围根本无关。

      现实现:
        1. 复用 styles.py 的 _is_protected_paragraph / _is_body_like_paragraph
           (保证白名单与 `style body` 一致 —— 同源)
        2. 只对 BODY_LIKE 段扫直接 rPr/pPr 属性 (rFonts/color/sz/highlight/
           jc/ind/spacing 等非空 child)
        3. 阈值: body_mess_count > 20 才 found (兼容偶发遗留直接格式)
    """
    try:
        from docx import Document  # type: ignore
        from .styles import (  # type: ignore
            _is_protected_paragraph,
            _is_body_like_paragraph,
            _style_name_id_of,
            load_profile,
        )
    except Exception as e:
        return {"found": False, "error": f"import: {type(e).__name__}: {e}"}

    try:
        profile = load_profile("zdwp")  # 与 style body 默认 profile 一致
    except Exception as e:
        return {"found": False, "error": f"load_profile: {type(e).__name__}: {e}"}

    try:
        doc = Document(str(doc_path))
    except Exception as e:
        return {"found": False, "error": f"open: {type(e).__name__}: {e}"}

    W_NS_LOCAL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    # 视为"直接格式"的 pPr/rPr 子元素 (空 pPr/rPr 本身或仅含 pStyle/rStyle 不算)
    PPR_MESS_TAGS = {"jc", "ind", "spacing", "shd", "pBdr", "numPr", "framePr",
                     "outlineLvl", "tabs"}
    RPR_MESS_TAGS = {"rFonts", "color", "sz", "szCs", "highlight", "b", "bCs",
                     "i", "iCs", "u", "strike", "vertAlign", "shd"}

    def _has_direct_ppr(p_el) -> bool:
        pPr = p_el.find(f"{W_NS_LOCAL}pPr")
        if pPr is None:
            return False
        for child in pPr:
            tag = child.tag
            if not tag.startswith(W_NS_LOCAL):
                continue
            local = tag[len(W_NS_LOCAL):]
            if local in PPR_MESS_TAGS:
                return True
        return False

    def _has_direct_rpr(p_el) -> bool:
        for r in p_el.iter(f"{W_NS_LOCAL}r"):
            rPr = r.find(f"{W_NS_LOCAL}rPr")
            if rPr is None:
                continue
            for child in rPr:
                tag = child.tag
                if not tag.startswith(W_NS_LOCAL):
                    continue
                local = tag[len(W_NS_LOCAL):]
                if local in RPR_MESS_TAGS:
                    return True
        return False

    total = 0
    protected_skipped = 0
    unknown_skipped = 0
    body_total = 0
    body_mess_count = 0
    mess_by_style: Counter = Counter()
    mess_kind_counter: Counter = Counter()  # "rPr-only" / "pPr-only" / "both"

    for p in doc.paragraphs:
        total += 1
        text = (p.text or "").strip()
        if text == "":
            continue
        style_name, style_id = _style_name_id_of(p)
        if _is_protected_paragraph(style_name, style_id, profile):
            protected_skipped += 1
            continue
        if not _is_body_like_paragraph(style_name, style_id, profile):
            unknown_skipped += 1
            continue
        body_total += 1
        p_el = p._p
        has_ppr = _has_direct_ppr(p_el)
        has_rpr = _has_direct_rpr(p_el)
        if has_ppr or has_rpr:
            body_mess_count += 1
            key = f"{style_name or '?'}|{style_id or '?'}"
            mess_by_style[key] += 1
            if has_ppr and has_rpr:
                mess_kind_counter["both"] += 1
            elif has_ppr:
                mess_kind_counter["pPr-only"] += 1
            else:
                mess_kind_counter["rPr-only"] += 1

    THRESHOLD = 20  # 容忍偶发,> 20 才报
    base = {
        "total_paragraphs": total,
        "protected_skipped": protected_skipped,
        "unknown_style_skipped": unknown_skipped,
        "body_total": body_total,
        "body_mess_count": body_mess_count,
        "mess_kinds": dict(mess_kind_counter),
        "top_mess_styles": dict(mess_by_style.most_common(5)),
    }
    if body_mess_count > THRESHOLD:
        base.update({"found": True, "safe_fix": True, "fix_hint": "style body"})
        return base
    base["found"] = False
    return base


def check_duplicate_figures(doc_path: Path, tmp_dir: Path) -> dict:
    """扩展 audit_caption_outline 检测同章内图/表号重复。"""
    rpt = tmp_dir / "captions_dup.json"
    data = _run_script_json("audit_caption_outline", [str(doc_path)], rpt)
    # look for duplicates in caption_list
    captions = data.get("captions", data.get("caption_list", []))
    if not captions:
        return {"found": False, "reason": "no captions found by audit_caption_outline"}

    seen: dict[str, list] = {}
    for cap in captions:
        label = cap.get("label", cap.get("text", "")) if isinstance(cap, dict) else str(cap)
        key = re.sub(r"\s+", "", label)[:30]
        seen.setdefault(key, []).append(cap)
    dups = {k: v for k, v in seen.items() if len(v) > 1}
    if dups:
        return {"found": True, "duplicate_count": len(dups),
                "examples": list(dups.keys())[:5], "safe_fix": False,
                "fix_hint": "需人工确认正确序号后再修"}
    return {"found": False}


def check_heading_number_stale(doc_path: Path, tmp_dir: Path) -> dict:
    rpt = tmp_dir / "heading_audit.json"
    data = _run_script_json("audit_heading_numbers", [str(doc_path)], rpt)
    no_prefix = data.get("h_without_prefix", 0)
    with_prefix = data.get("h_with_prefix", 0)
    total = no_prefix + with_prefix
    if total == 0:
        return {"found": False, "reason": "no headings"}
    # If significant fraction lack prefix → stale numbering signal
    if no_prefix > 0 and (with_prefix == 0 or no_prefix / total > 0.3):
        return {"found": True, "no_prefix_count": no_prefix, "with_prefix_count": with_prefix,
                "safe_fix": True, "fix_hint": "renumber headings"}
    return {"found": False, "no_prefix_count": no_prefix, "with_prefix_count": with_prefix}


# ─── HealthChecker ────────────────────────────────────────────────────────────

class HealthChecker:
    def __init__(self, doc_path: Path, workers: int = 8):
        self.doc_path = doc_path
        self.workers = workers
        self._tmp_dir = Path("/tmp") / f"docx_health_{doc_path.stem[:20]}"
        self._tmp_dir.mkdir(parents=True, exist_ok=True)

    def run_all(self, checks: list[str] | None = None) -> dict[str, dict]:
        checks = checks or ALL_CHECKS
        results: dict[str, dict] = {}

        def run_one(check_id: str) -> tuple[str, dict]:
            try:
                r = self._run_check(check_id)
            except Exception as e:
                r = {"found": False, "error": f"{type(e).__name__}: {e}"}
            return check_id, r

        with ThreadPoolExecutor(max_workers=self.workers) as ex:
            futs = {ex.submit(run_one, c): c for c in checks}
            for fut in as_completed(futs):
                cid, res = fut.result()
                results[cid] = res

        return results

    def _run_check(self, check_id: str) -> dict:
        dp = self.doc_path
        td = self._tmp_dir
        if check_id == "heading-level-skew":
            return check_heading_level_skew(dp)
        if check_id == "heading-gap":
            return check_heading_gap(dp)
        if check_id == "caption-outline-pollution":
            return check_caption_outline(dp, td)
        if check_id == "revision-tracking-residue":
            return check_revision_tracking(dp)
        if check_id == "field-not-frozen":
            return check_field_not_frozen(dp, td)
        if check_id == "body-style-mess":
            return check_body_style_mess(dp)
        if check_id == "duplicate-figure-numbers":
            return check_duplicate_figures(dp, td)
        if check_id == "heading-number-stale":
            return check_heading_number_stale(dp, td)
        return {"found": False, "error": f"unknown check: {check_id}"}


# ─── HealthFixer ─────────────────────────────────────────────────────────────

SAFE_FIX_SCRIPTS: dict[str, list[str]] = {
    "caption-outline-pollution": ["strip_outlinelvl_from_captions"],
    "revision-tracking-residue": ["strip_revisions"],
    "field-not-frozen":          ["freeze_all_fields"],
    "heading-number-stale":      ["renumber_headings"],
}


class HealthFixer:
    def __init__(self, doc_path: Path, dry_run: bool = False, backup: bool = True):
        self.doc_path = doc_path
        self.dry_run = dry_run
        self.backup = backup

    def fix_safe(self, diagnose_results: dict[str, dict], auto_list: list[str] | None = None) -> dict:
        """串行执行 safe-fix 白名单中命中的病种修复。"""
        if auto_list is None:
            auto_list = [k for k, v in SAFE_FIX.items() if v]

        applied: list[str] = []
        skipped: list[str] = []
        plan_required: list[str] = []

        for check_id in auto_list:
            result = diagnose_results.get(check_id, {})
            if not result.get("found"):
                continue

            if not SAFE_FIX.get(check_id):
                plan_required.append(check_id)
                continue

            # heading-level-skew: only auto if coverage >= 0.8
            if check_id == "heading-level-skew":
                delta = result.get("delta", 0)
                coverage = result.get("coverage", 0)
                if abs(delta) == 0 or coverage < 0.8:
                    skipped.append(check_id)
                    continue
                sub = "promote-h1" if delta > 0 else "demote-h2"
                argv = [str(self.doc_path)]
                if self.dry_run:
                    argv.append("--dry-run")
                if not self.backup:
                    argv.append("--no-backup")
                rc = exec_script("outline", [sub] + argv)
                applied.append(f"{check_id}(outline {sub}, rc={rc})")
                continue

            # body-style-mess → style body via styles.py (style subcommand)
            if check_id == "body-style-mess":
                argv = [str(self.doc_path)]
                if self.dry_run:
                    argv.append("--dry-run")
                if not self.backup:
                    argv.append("--no-backup")
                rc = exec_script("styles", ["body"] + argv)
                applied.append(f"{check_id}(styles body, rc={rc})")
                continue

            scripts = SAFE_FIX_SCRIPTS.get(check_id, [])
            for script in scripts:
                argv = [str(self.doc_path)]
                if self.dry_run:
                    argv.append("--dry-run")
                if not self.backup:
                    argv.append("--no-backup")
                rc = exec_script(script, argv)
                applied.append(f"{check_id}({script}, rc={rc})")

        return {"applied": applied, "skipped": skipped, "plan_required": plan_required}


# ─── CLI handlers ─────────────────────────────────────────────────────────────

def _parse_checks(checks_str: str) -> list[str]:
    if checks_str == "all":
        return ALL_CHECKS
    return [c.strip() for c in checks_str.split(",") if c.strip()]


def _print_summary(results: dict[str, dict], doc_path: Path) -> int:
    """Print summary table; return exit code (0/1/2)."""
    max_rc = 0
    lines = [
        f"\n{'─'*70}",
        f"  docx health diagnose: {doc_path.name}",
        f"{'─'*70}",
        f"  {'Check ID':<32} {'Found':<6} {'Sev':<5} {'AutoFix':<8} {'Detail'}",
        f"  {'─'*32} {'─'*6} {'─'*5} {'─'*8} {'─'*20}",
    ]
    for cid in ALL_CHECKS:
        res = results.get(cid, {"found": False})
        found = res.get("found", False)
        sev = SEVERITY.get(cid, "?")
        safe = "✅" if SAFE_FIX.get(cid) else "❌ plan"
        if found:
            detail = res.get("fix_hint", res.get("error", ""))
            if "delta" in res:
                detail = f"delta={res['delta']:+d}, coverage={res.get('coverage', '?')}"
            lines.append(f"  {'⚠ ' + cid:<34} {'YES':<6} {sev:<5} {safe:<10} {detail}")
            if sev == "High":
                max_rc = max(max_rc, 2)
            elif sev in ("Med", "Low"):
                max_rc = max(max_rc, 1)
        else:
            err = res.get("error", "")
            detail = f"err: {err}" if err else "ok"
            lines.append(f"  {'  ' + cid:<34} {'no':<6} {sev:<5} {safe:<10} {detail}")
    lines.append(f"{'─'*70}")
    lines.append(f"  exit_code={max_rc}  (0=healthy / 1=warning / 2=error)\n")
    print("\n".join(lines))
    return max_rc


def cmd_diagnose(args) -> int:
    doc_path = Path(args.docx_path)
    if not doc_path.exists():
        print(f"[health] ERROR: file not found: {doc_path}", file=sys.stderr)
        return 2
    checks = _parse_checks(getattr(args, "checks", "all") or "all")
    workers = getattr(args, "workers", 8) or 8
    checker = HealthChecker(doc_path, workers=workers)
    results = checker.run_all(checks)

    rc = _print_summary(results, doc_path)

    report_path = getattr(args, "report", None)
    if report_path:
        rp = Path(report_path)
        payload = {
            "docx": str(doc_path),
            "checks": results,
            "exit_code": rc,
        }
        rp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[health] report saved → {rp}")

    if getattr(args, "html", False):
        style = getattr(args, "html_style", "rich") or "rich"
        _write_html_report(doc_path, results, rc, style=style,
                           tmp_dir=checker._tmp_dir)

    return rc


def cmd_fix(args) -> int:
    doc_path = Path(args.docx_path)
    if not doc_path.exists():
        print(f"[health] ERROR: file not found: {doc_path}", file=sys.stderr)
        return 2

    plan_path = getattr(args, "plan", None)
    if plan_path:
        data = json.loads(Path(plan_path).read_text("utf-8"))
        diagnose_results = data.get("checks", {})
    else:
        # Run diagnose first
        checker = HealthChecker(doc_path)
        diagnose_results = checker.run_all()

    auto_str = getattr(args, "auto", None)
    auto_list = _parse_checks(auto_str) if auto_str else None
    dry_run = getattr(args, "dry_run", False)
    backup = getattr(args, "backup", True)

    fixer = HealthFixer(doc_path, dry_run=dry_run, backup=backup)
    fix_result = fixer.fix_safe(diagnose_results, auto_list)

    print(f"\n[health fix] applied: {fix_result['applied']}")
    print(f"[health fix] skipped: {fix_result['skipped']}")
    print(f"[health fix] plan_required: {fix_result['plan_required']}")
    return 0


def cmd_full(args) -> int:
    """full = diagnose → fix safe → re-diagnose → print before/after."""
    doc_path = Path(args.docx_path)
    if not doc_path.exists():
        print(f"[health] ERROR: file not found: {doc_path}", file=sys.stderr)
        return 2
    dry_run = getattr(args, "dry_run", False)

    print("[health full] Phase 1: diagnose …")
    checker = HealthChecker(doc_path)
    before = checker.run_all()
    rc_before = _print_summary(before, doc_path)

    if not dry_run:
        print("\n[health full] Phase 2: auto-fix safe checks …")
        fixer = HealthFixer(doc_path, dry_run=False, backup=True)
        fix_result = fixer.fix_safe(before)
        print(f"  applied: {fix_result['applied']}")
        print(f"  plan_required: {fix_result['plan_required']}")

        print("\n[health full] Phase 3: re-diagnose …")
        checker2 = HealthChecker(doc_path)
        after = checker2.run_all()
        rc_after = _print_summary(after, doc_path)

        if getattr(args, "html", False):
            style = getattr(args, "html_style", "rich") or "rich"
            _write_html_report(doc_path, after, rc_after, before=before,
                               style=style, tmp_dir=checker2._tmp_dir)
        return rc_after
    else:
        print("[health full] --dry-run: skipping fix phases")
        if getattr(args, "html", False):
            style = getattr(args, "html_style", "rich") or "rich"
            _write_html_report(doc_path, before, rc_before, style=style,
                               tmp_dir=checker._tmp_dir)
        return rc_before


# ─── HTML 报告 ────────────────────────────────────────────────────────────────

def _write_html_report(doc_path: Path, results: dict, rc: int,
                       before: dict | None = None,
                       style: str = "rich",
                       tmp_dir: Path | None = None) -> None:
    """Write HTML report; style ∈ {'rich','simple'} (default rich).

    rich: 富 HTML (vault-citizen 范式) — TOC + 侧栏 + 8 病种卡 + caption 表 +
          orphan media + orphan tables + bookmarks + 修复 SOP。
    simple: 旧版单表 HTML (向后兼容)。
    """
    out = doc_path.parent / f"{doc_path.stem}_health_report.html"
    if style == "simple":
        render_simple_html(results, doc_path, out, rc, before=before)
    else:
        render_rich_html(results, doc_path, out, tmp_dir=tmp_dir, before=before)
    size = out.stat().st_size if out.exists() else 0
    print(f"[health] HTML report ({style}, {size:,} bytes) → {out}")


# ─── register() for docx_cli.py ──────────────────────────────────────────────

def register(subparsers) -> None:
    """Register `health <subcommand>` onto docx_cli.py's top-level subparsers."""
    p = get_or_add_group(
        subparsers, "health",
        help_text="docx health diagnose / fix / full (8 病种检查)",
    )
    sp = get_or_add_subparsers(p, dest="health_sub", metavar="<subcommand>")

    # diagnose
    diag = sp.add_parser("diagnose", help="诊断 8 病种，输出 summary + 可选 JSON/HTML 报告")
    diag.add_argument("docx_path", help="target docx")
    diag.add_argument("--checks", default="all",
                      help="逗号分隔病种 ID 或 'all' (default: all)")
    diag.add_argument("--report", help="输出 JSON 报告路径")
    diag.add_argument("--html", action="store_true", help="输出 HTML 报告")
    diag.add_argument("--html-style", choices=["rich", "simple"], default="rich",
                      help="HTML 风格: rich (默认富 HTML / vault-citizen) | simple (单表)")
    diag.add_argument("--workers", type=int, default=8, help="并发线程数 (default: 8)")
    diag.set_defaults(func=cmd_diagnose)

    # fix
    fix = sp.add_parser("fix", help="执行 safe-fix 白名单修复")
    fix.add_argument("docx_path", help="target docx")
    fix.add_argument("--auto", help="逗号分隔 check ID，默认全 safe 白名单")
    fix.add_argument("--plan", help="从 diagnose --report 输出的 JSON 读任务")
    fix.add_argument("--dry-run", action="store_true", help="不写文件，只打印 diff")
    fix.add_argument("--backup", action="store_true", default=True, help="自动备份 (default: on)")
    fix.set_defaults(func=cmd_fix)

    # full
    full = sp.add_parser("full", help="diagnose → fix safe → re-diagnose → HTML 报告")
    full.add_argument("docx_path", help="target docx")
    full.add_argument("--html", action="store_true", help="输出 HTML 报告")
    full.add_argument("--html-style", choices=["rich", "simple"], default="rich",
                      help="HTML 风格: rich (默认富 HTML / vault-citizen) | simple (单表)")
    full.add_argument("--dry-run", action="store_true", help="不写文件")
    full.set_defaults(func=cmd_full)
