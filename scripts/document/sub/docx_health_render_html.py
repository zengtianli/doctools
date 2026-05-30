#!/usr/bin/env python3
"""docx_health_render_html.py — 富 HTML 体检报告渲染器 (vault-citizen 范式).

distilled from taizhou-天台/bin/render_docx_health_html.py (2026-05-26).
通用化: 不依赖 /tmp 路径字面量,接受 dict 输入 + tmp_dir 路径,渲染单文件 HTML。

入口:
  render_rich_html(diagnose_result, docx_path, out_path, tmp_dir=None) -> None

设计:
  - HealthChecker 把 /tmp/docx_health_<stem>/ 下 5 个 audit JSON 写好 (heading/captions/
    fields/table-pairing/bookmarks/images),本渲染器读它们补充明细。
  - 8 病种 status 来自 diagnose_result (主表),明细数据来自 tmp_dir 下的 JSON。
  - 无明细 JSON 也能渲染 (graceful degrade 到只有 8 病种卡 + 摘要)。
  - 纯 stdlib,无 CDN,响应式 (≤1100px 单列)。

输出: 单文件 HTML 35-60 KB 含
  顶部 nav + 侧边 TOC + 主区 (摘要/结构/8 病种卡/captions 表/orphan media/orphan tables/
  bookmarks/修复 SOP) + 侧边 backlinks + footer。
"""
from __future__ import annotations

import datetime
import html as _html
import json
import re
from pathlib import Path
from typing import Any

# ─── 8 病种科普文案 (静态) ───────────────────────────────────────────────────

CHECKS_DOC: list[tuple[str, str, str, str, str, str, str]] = [
    ("heading-level-skew", "标题级别整体偏移",
     "全文标题被整体提/降了 N 级 (例 H2 当 H3 用),Word 导航栏层级错乱、TOC 缩进错位。",
     "扫所有 Heading 段比对实际级别 vs 章节序号深度,看是否整体差固定 delta。",
     "✅ safe auto-fix (coverage ≥ 0.8 时自动)",
     "High",
     "python3 docx_cli.py outline promote-h1 <file>  # 或 demote-h2"),
    ("heading-gap", "标题级别跳级",
     "H1 直接跳 H3、漏了 H2,导致 TOC 出现空层。",
     "扫所有 Heading 段,任意相邻对差 > 1 即报。",
     "❌ plan only (需用户决定补 H2 还是降 H3)",
     "Med",
     "python3 docx_cli.py outline normalize-arabic <file> --plan"),
    ("caption-outline-pollution", "图表名污染大纲",
     "图/表 caption 段被错配成 Heading 样式,Word 导航栏出现 \"图 2-1 XXX\" 当章节。",
     "扫 outlineLvl ≠ null 且段文本以 \"图|表\" 开头。",
     "✅ safe auto-fix (改回 caption 样式)",
     "Med",
     "python3 docx_cli.py strip outlinelvl <file>"),
    ("revision-tracking-residue", "修订标记残留",
     "合稿后仍含未接受的 track-changes (w:ins/w:del),交付件外人看到删除线 / 红字。",
     "扫 document.xml 找 w:ins / w:del / w:moveFrom / w:moveTo。",
     "✅ safe auto-fix (全部 accept)",
     "High",
     "python3 docx_cli.py strip revisions <file>"),
    ("field-not-frozen", "字段未冻结",
     "TOC / 交叉引用 / SEQ 编号仍是 field 代码,PDF 化或他人无 Word 时显示 { TOC \\o ... }。",
     "扫 w:fldChar / w:instrText 是否存在 TOC/PAGEREF/SEQ/REF/HYPERLINK 等 hot type。",
     "✅ safe auto-fix (转为纯文本)",
     "Med",
     "python3 docx_cli.py freeze all-fields <file>"),
    ("body-style-mess", "正文样式混乱",
     "Normal 段叠加大量 inline 字号/字体/段落格式,字号忽大忽小、行距不一致。",
     "统计正文段非 Heading/Caption 样式数量,> 3 种异常 styleId 即报。",
     "✅ safe auto-fix (套统一 body 模板)",
     "Low",
     "python3 docx_cli.py styles body <file>"),
    ("duplicate-figure-numbers", "图号重复",
     "多个 caption 共用同号 (图 3-1 出现两次),引用无法安全 remap。",
     "抽 caption 文本里的图号,数 frequency > 1。",
     "❌ plan only (需用户先消重再 renumber)",
     "High",
     "先人工消重 → python3 docx_cli.py renumber figures <file>"),
    ("heading-number-stale", "标题硬编码号过期",
     "手敲编号 (\"3.2 自然地理\") 与样式实际级别 (H4) 不匹配,改章节顺序后旧号留尸。",
     "解析标题文本前缀号深度 vs 样式级,无前缀 / 深度不一致占比 > 30% 即报。",
     "✅ safe auto-fix (按样式重排)",
     "Med",
     "python3 docx_cli.py renumber headings <file>"),
]

SEV_RANK = {"High": 0, "Med": 1, "Low": 2}

# ─── CSS (inline, 无 CDN) ─────────────────────────────────────────────────────

CSS = r"""
:root { --fg:#1f2937; --muted:#6b7280; --line:#e5e7eb; --bg:#fff;
  --accent:#2563eb; --ok:#16a34a; --warn:#ea580c; --bad:#dc2626;
  --sev-h:#dc2626; --sev-m:#ea580c; --sev-l:#ca8a04; }
*{box-sizing:border-box}
body{font-family:-apple-system,'PingFang SC','Helvetica Neue',sans-serif;
  color:var(--fg);background:#f9fafb;margin:0;line-height:1.55;font-size:14.5px}
.layout{display:grid;grid-template-columns:240px minmax(0,1fr) 220px;gap:24px;
  max-width:1400px;margin:0 auto;padding:24px}
nav.top{grid-column:1/-1;background:var(--bg);border:1px solid var(--line);
  border-radius:10px;padding:14px 20px;display:flex;justify-content:space-between;
  align-items:center;flex-wrap:wrap;gap:12px}
nav.top h1{margin:0;font-size:18px;font-weight:600}
nav.top .meta{color:var(--muted);font-size:13px}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;
  font-size:12px;font-weight:600;margin-left:8px}
.badge.warn{background:#fef3c7;color:#92400e}
.badge.ok{background:#d1fae5;color:#065f46}
.badge.bad{background:#fee2e2;color:#991b1b}
aside.toc,aside.backlinks{position:sticky;top:24px;align-self:start;
  background:var(--bg);border:1px solid var(--line);border-radius:10px;
  padding:14px 16px;font-size:13px}
aside.toc{max-height:calc(100vh - 80px);overflow-y:auto}
aside.toc h3,aside.backlinks h3{margin:0 0 8px;font-size:12px;color:var(--muted);
  text-transform:uppercase;letter-spacing:0.05em}
aside.toc a{display:block;color:var(--fg);text-decoration:none;
  padding:4px 0;border-left:2px solid transparent;padding-left:8px;margin-left:-8px}
aside.toc a:hover{color:var(--accent);border-color:var(--accent)}
aside.backlinks a{display:block;color:var(--accent);text-decoration:none;padding:3px 0}
main{min-width:0}
section{background:var(--bg);border:1px solid var(--line);border-radius:10px;
  padding:20px 24px;margin-bottom:18px;scroll-margin-top:20px}
section h2{margin:0 0 12px;font-size:18px;font-weight:600;
  padding-bottom:8px;border-bottom:1px solid var(--line)}
section h3{font-size:15px;margin:16px 0 8px}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.stat{padding:12px 14px;background:#f3f4f6;border-radius:8px}
.stat .n{font-size:22px;font-weight:600;display:block;line-height:1.2}
.stat .l{font-size:12px;color:var(--muted)}
.stat.warn{background:#fef3c7}
.stat.bad{background:#fee2e2}
.stat.ok{background:#d1fae5}
table{border-collapse:collapse;width:100%;font-size:13.5px;margin:6px 0}
th,td{border:1px solid var(--line);padding:6px 10px;text-align:left;vertical-align:top}
th{background:#f3f4f6;font-weight:600;font-size:12.5px}
tr.row-warn td{background:#fffbeb}
tr.row-bad td{background:#fef2f2}
tr.row-ok td{background:transparent}
.sev{font-weight:600;font-size:12px;padding:2px 8px;border-radius:10px}
.sev.High{background:#fee2e2;color:var(--sev-h)}
.sev.Med{background:#ffedd5;color:var(--sev-m)}
.sev.Low{background:#fef9c3;color:var(--sev-l)}
.status-ok{color:var(--ok);font-weight:600}
.status-warn{color:var(--warn);font-weight:600}
.fix-cmd{background:#0f172a;color:#e2e8f0;padding:8px 12px;border-radius:6px;
  font-family:'SF Mono',Menlo,monospace;font-size:12.5px;display:block;
  white-space:pre-wrap;margin-top:6px;overflow-x:auto}
.fix-cmd::before{content:"$ ";color:#64748b}
details{margin:8px 0;border:1px solid var(--line);border-radius:6px;
  padding:0 14px;background:#fafafa}
details summary{cursor:pointer;padding:10px 0;font-weight:600;font-size:13.5px}
details[open]{padding-bottom:14px}
.check-body{font-size:13.5px;color:#374151}
.check-body p{margin:6px 0}
.kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;font-size:13px;margin:8px 0}
.kv dt{color:var(--muted);font-weight:500}
.kv dd{margin:0}
.tag{display:inline-block;padding:1px 7px;background:#e5e7eb;border-radius:10px;
  font-size:11.5px;color:#374151;margin-right:4px}
code{background:#f3f4f6;padding:1px 5px;border-radius:3px;font-size:90%}
.empty-cell{color:#dc2626;font-style:italic;font-size:11.5px}
.scroll-x{overflow-x:auto;max-height:480px;overflow-y:auto}
footer{grid-column:1/-1;text-align:center;color:var(--muted);
  font-size:12px;margin-top:8px;padding-top:16px}
@media (max-width:1100px){.layout{grid-template-columns:1fr}
  aside.toc,aside.backlinks{position:static;max-height:none}}
"""


def _esc(s: Any) -> str:
    return _html.escape("" if s is None else str(s))


def _load_json(p: Path) -> dict:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _read_lines(p: Path) -> list[str]:
    try:
        return p.read_text(encoding="utf-8").splitlines()
    except Exception:
        return []


def _gather_supplemental(tmp_dir: Path | None, docx_path: Path) -> dict:
    """从 tmp_dir + /tmp/audit-images-<stem>.json 收集 5 类 audit 明细。

    HealthChecker 会写: heading_audit.json / caption_outline.json / fields.json /
                       captions_dup.json
    audit images 默认写: /tmp/audit-images-<stem>.json
    table-pairing / bookmarks 不在 health 主流程中预跑;若 tmp_dir 下 / 标准位置
    存在则读,否则空。
    """
    out: dict[str, Any] = {
        "heading": {},
        "caption_outline": {},
        "fields": {},
        "images": {},
        "table_pairing": {},
        "bookmarks": {},
    }
    if tmp_dir and tmp_dir.exists():
        out["heading"] = _load_json(tmp_dir / "heading_audit.json")
        out["caption_outline"] = _load_json(tmp_dir / "caption_outline.json")
        out["fields"] = _load_json(tmp_dir / "fields.json")
        # also try alternate names
        for name in ("table_pairing.json", "table-pairing.json"):
            p = tmp_dir / name
            if p.exists():
                out["table_pairing"] = _load_json(p)
                break
        for name in ("bookmarks.json", "audit_bookmarks.json"):
            p = tmp_dir / name
            if p.exists():
                out["bookmarks"] = _load_json(p)
                break

    # images default location
    img_p = Path("/tmp") / f"audit-images-{docx_path.stem}.json"
    if img_p.exists():
        out["images"] = _load_json(img_p)
    return out


def _render_check_card(check_tuple: tuple, found: bool, detail: dict) -> str:
    cid, name, what, how, autofix, sev, cmd = check_tuple
    icon = "⚠️ FOUND" if found else "✅ OK"
    status_cls = "warn" if found else "ok"
    detail_html = ""
    if found and detail:
        items = []
        for k, v in detail.items():
            if k in ("found", "safe_fix"):
                continue
            if isinstance(v, (dict, list)):
                v_str = json.dumps(v, ensure_ascii=False)[:300]
            else:
                v_str = str(v)[:300]
            items.append(f"<dt>{_esc(k)}</dt><dd><code>{_esc(v_str)}</code></dd>")
        if items:
            detail_html = f'<p><strong>本次检测明细</strong>:</p><dl class="kv">{"".join(items)}</dl>'
    return f"""
    <details {'open' if found else ''}>
      <summary>
        <span class="sev {sev}">{sev}</span>
        <code>{_esc(cid)}</code> · {_esc(name)} ·
        <span class="status-{status_cls}">{icon}</span>
      </summary>
      <div class="check-body">
        <p><strong>是什么</strong>: {_esc(what)}</p>
        <p><strong>怎么判定</strong>: {_esc(how)}</p>
        <p><strong>自动修复</strong>: {_esc(autofix)}</p>
        <p><strong>修复命令</strong>:</p>
        <code class="fix-cmd">{_esc(cmd)}</code>
        {detail_html}
      </div>
    </details>
    """


def _caption_rows(records: list[dict]) -> str:
    rows = []
    for r in records:
        idx = r.get("idx", "")
        style = r.get("style", "")
        text = r.get("text") or ""
        ol = r.get("outlineLvl")
        is_empty = not text.strip()
        is_normal = (style == "Normal")
        row_cls = "row-bad" if is_empty else ("row-warn" if is_normal else "row-ok")
        text_disp = ('<span class="empty-cell">(空 caption)</span>'
                     if is_empty else _esc(text))
        rows.append(
            f"<tr class='{row_cls}'><td>{_esc(idx)}</td>"
            f"<td><code>{_esc(style)}</code></td>"
            f"<td>{text_disp}</td>"
            f"<td>{_esc(ol) if ol is not None else '—'}</td></tr>"
        )
    return "".join(rows)


def _orphan_media_rows(orphans: list) -> str:
    return "".join(
        f"<tr class='row-warn'><td><code>{_esc(m)}</code></td></tr>"
        for m in orphans
    )


def _drawings_rows(drawings: list[dict]) -> str:
    return "".join(
        f"<tr><td>{_esc(d.get('para_idx'))}</td>"
        f"<td>{_esc(d.get('subtype', ''))}</td>"
        f"<td><code>{_esc(d.get('rid', ''))}</code></td>"
        f"<td><code>{_esc(d.get('target', ''))}</code></td>"
        f"<td>{_esc(d.get('status', ''))}</td></tr>"
        for d in drawings
    )


def _table_orphan_rows(tbl_data: dict) -> str:
    """table_pairing.json 结构: { issues: [ {type, ...} ], ... }"""
    issues = tbl_data.get("issues", []) if isinstance(tbl_data, dict) else []
    rows = []
    for it in issues:
        if it.get("type") != "orphan-tbl-no-upstream-caption":
            continue
        tid = it.get("tbl_id") or it.get("id") or ""
        idx = it.get("elem_idx") or it.get("idx") or ""
        first_row = it.get("first_row") or it.get("details") or ""
        if isinstance(first_row, str):
            disp = first_row[:80] + ("..." if len(first_row) > 80 else "")
        else:
            disp = str(first_row)[:80]
        rows.append(
            f"<tr class='row-warn'><td>{_esc(tid)}</td>"
            f"<td>{_esc(idx)}</td><td>{_esc(disp)}</td></tr>"
        )
    return "".join(rows)


def render_rich_html(
    diagnose_result: dict,
    docx_path: str | Path,
    out_path: str | Path,
    tmp_dir: str | Path | None = None,
    before: dict | None = None,
) -> Path:
    """渲染富 HTML 体检报告到 out_path,返回 out_path。

    Args:
      diagnose_result: HealthChecker.run_all() 返回的 {check_id: {found, ...}}
      docx_path:       被体检的 docx 路径 (用于显示)
      out_path:        HTML 输出路径
      tmp_dir:         HealthChecker._tmp_dir,内含 heading_audit.json 等明细
      before:          (optional) full 模式 phase1 的 diagnose_result,做对比展示
    """
    docx_path = Path(docx_path)
    out_path = Path(out_path)
    tmp_dir = Path(tmp_dir) if tmp_dir else None

    sup = _gather_supplemental(tmp_dir, docx_path)

    # 计 found_count / exit_code
    found_count = sum(1 for r in diagnose_result.values() if r.get("found"))
    rc = 0
    for cid in diagnose_result:
        if not diagnose_result[cid].get("found"):
            continue
        sev = next((c[5] for c in CHECKS_DOC if c[0] == cid), "Low")
        if sev == "High":
            rc = max(rc, 2)
        else:
            rc = max(rc, 1)

    # heading audit
    heading = sup["heading"]
    total_paras = heading.get("total_paragraphs", 0)
    h_count = heading.get("h_count", {})
    cap_records = heading.get("all_caption_records", [])
    cap_by_style = heading.get("captions_by_style", {})
    caption_styles_avail = heading.get("caption_styles_available", [])
    issues = heading.get("issues", {})
    empty_n = issues.get("empty_caption", 0)
    wrong_style_n = issues.get("wrong_style_count", 0)

    # images
    images = sup["images"]
    img_summary = images.get("summary", {}) if isinstance(images, dict) else {}
    orphan_media = img_summary.get("orphan_media", [])
    drawings_used = images.get("drawings", []) if isinstance(images, dict) else []

    # table pairing
    tbl = sup["table_pairing"]
    tbl_summary = tbl.get("summary", {}) if isinstance(tbl, dict) else {}
    orphan_tbl_n = tbl_summary.get("orphan_tbls", 0)

    # bookmarks
    bm = sup["bookmarks"]
    bm_total = bm.get("bookmark_count", 0)
    bm_by_prefix = bm.get("bookmark_by_prefix", {})
    bm_orphan_start = len(bm.get("orphan_starts", [])) if isinstance(bm, dict) else 0
    bm_orphan_end = len(bm.get("orphan_ends", [])) if isinstance(bm, dict) else 0

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── summary table ────────────────────────────────────────────────────────
    summary_rows = []
    for c in CHECKS_DOC:
        cid = c[0]
        sev = c[5]
        res = diagnose_result.get(cid, {})
        found = res.get("found", False)
        before_cell = ""
        if before is not None:
            bf = before.get(cid, {}).get("found", False)
            before_cell = f"<td>{'⚠ FOUND' if bf else 'ok'}</td>"
        row_cls = "row-warn" if found else "row-ok"
        safe = "✅ auto" if "safe auto-fix" in c[4] else "❌ plan"
        summary_rows.append(
            f"<tr class='{row_cls}'>"
            f"<td><code>{_esc(cid)}</code></td>{before_cell}"
            f"<td><span class='sev {sev}'>{sev}</span></td>"
            f"<td>{'⚠️ FOUND' if found else '✅ ok'}</td>"
            f"<td>{safe}</td>"
            f"<td>{_esc(c[1])}</td></tr>"
        )
    before_th = "<th>Before</th>" if before is not None else ""

    # ── check cards (sorted by sev then by found-first) ─────────────────────
    sorted_checks = sorted(
        CHECKS_DOC,
        key=lambda c: (0 if diagnose_result.get(c[0], {}).get("found") else 1,
                       SEV_RANK.get(c[5], 9))
    )
    cards_html = "".join(
        _render_check_card(c, diagnose_result.get(c[0], {}).get("found", False),
                           diagnose_result.get(c[0], {}))
        for c in sorted_checks
    )

    # ── document.write ───────────────────────────────────────────────────────
    badge_cls = "ok" if found_count == 0 else ("warn" if rc == 1 else "bad")

    cap_table_section = ""
    if cap_records:
        cap_table_section = f"""
<section id="captions">
  <h2>🏷️ Caption 完整数据 ({len(cap_records)} 条)</h2>
  <p style="color:var(--muted)">
    <span class='tag' style='background:#fee2e2;color:#991b1b'>红行 = 空 caption</span>
    <span class='tag' style='background:#fef3c7;color:#92400e'>黄行 = 样式 = Normal (应为 caption)</span>
  </p>
  <div class="scroll-x">
    <table>
      <tr><th>段 idx</th><th>style</th><th>文本</th><th>outlineLvl</th></tr>
      {_caption_rows(cap_records)}
    </table>
  </div>
</section>"""

    tbl_section = ""
    tbl_rows_html = _table_orphan_rows(tbl) if tbl else ""
    if tbl_rows_html:
        tbl_section = f"""
<section id="tables">
  <h2>📋 表格配对审计 (orphan = {orphan_tbl_n})</h2>
  <p style="color:var(--muted)">每行 = 一张表 + 表前 8 段内未找到表名段。</p>
  <div class="scroll-x">
    <table>
      <tr><th>表 ID</th><th>段 idx</th><th>表首行 (前 80 字)</th></tr>
      {tbl_rows_html}
    </table>
  </div>
</section>"""

    img_section = ""
    if images:
        img_section = f"""
<section id="images">
  <h2>🖼️ 图片审计</h2>
  <h3>已使用的图 ({len(drawings_used)} 张 drawings)</h3>
  <div class="scroll-x">
    <table>
      <tr><th>段 idx</th><th>类型</th><th>rId</th><th>target</th><th>状态</th></tr>
      {_drawings_rows(drawings_used)}
    </table>
  </div>
  <h3>孤儿媒体 ({len(orphan_media)} 个)</h3>
  <p style="color:var(--muted)">这些媒体压在 docx zip 里但没被任何段引用,多为模板 placeholder 残留。</p>
  <div class="scroll-x">
    <table>
      <tr><th>文件名</th></tr>
      {_orphan_media_rows(orphan_media)}
    </table>
  </div>
</section>"""

    bm_section = ""
    if bm:
        bm_rows = "".join(
            f"<dt>{_esc(k)}</dt><dd>{_esc(v)}</dd>"
            for k, v in bm_by_prefix.items()
        )
        bm_section = f"""
<section id="bookmarks">
  <h2>🔖 书签</h2>
  <div class="kv">
    <dt>总数</dt><dd>{bm_total}</dd>
    {bm_rows}
    <dt>orphan starts / ends</dt><dd>{bm_orphan_start} / {bm_orphan_end} {'✅' if bm_orphan_start + bm_orphan_end == 0 else '⚠️'}</dd>
  </div>
</section>"""

    # TOC
    toc_links = [
        '<a href="#summary">📊 体检摘要</a>',
        '<a href="#structure">🏗️ 文档结构</a>',
        '<a href="#checks">🩺 8 病种详情</a>',
    ]
    if cap_records:
        toc_links.append(f'<a href="#captions">🏷️ Caption 数据 ({len(cap_records)})</a>')
    if tbl_rows_html:
        toc_links.append(f'<a href="#tables">📋 表格配对 ({orphan_tbl_n})</a>')
    if images:
        toc_links.append('<a href="#images">🖼️ 图片审计</a>')
    if bm:
        toc_links.append('<a href="#bookmarks">🔖 书签</a>')
    toc_links.append('<a href="#fix-sop">🔧 修复 SOP</a>')

    body = f"""<!DOCTYPE html>
<html lang="zh"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>docx health · {_esc(docx_path.stem)}</title>
<style>{CSS}</style>
</head>
<body>
<div class="layout">

<nav class="top">
  <div>
    <h1>📋 DOCX 体检报告
      <span class="badge {badge_cls}">
        exit_code = {rc} · {found_count}/{len(CHECKS_DOC)} 病种检出
      </span>
    </h1>
    <div class="meta">
      <strong>文件</strong>: <code>{_esc(docx_path.name)}</code> ·
      <strong>生成</strong>: {now} ·
      <strong>引擎</strong>: <code>docx_cli.py health full</code>
    </div>
  </div>
</nav>

<aside class="toc">
  <h3>跳转</h3>
  {''.join(toc_links)}
</aside>

<main>

<section id="summary">
  <h2>📊 体检摘要</h2>
  <div class="summary-grid">
    <div class="stat"><span class="n">{total_paras}</span><span class="l">总段落数</span></div>
    <div class="stat"><span class="n">{sum(v for v in h_count.values() if isinstance(v, (int, float))) if h_count else 0}</span><span class="l">标题总数</span></div>
    <div class="stat"><span class="n">{len(cap_records)}</span><span class="l">Caption 总数</span></div>
    <div class="stat {'bad' if empty_n else 'ok'}"><span class="n">{empty_n}</span><span class="l">空 Caption</span></div>
    <div class="stat {'warn' if wrong_style_n else 'ok'}"><span class="n">{wrong_style_n}</span><span class="l">Caption 样式偏</span></div>
    <div class="stat"><span class="n">{img_summary.get('media_files_count', 0)}</span><span class="l">媒体文件</span></div>
    <div class="stat {'bad' if img_summary.get('orphan_media_count', 0) else 'ok'}"><span class="n">{img_summary.get('orphan_media_count', 0)}</span><span class="l">孤儿媒体</span></div>
    <div class="stat {'warn' if orphan_tbl_n else 'ok'}"><span class="n">{orphan_tbl_n}</span><span class="l">orphan 表</span></div>
  </div>

  <h3>8 病种体检结论</h3>
  <table>
    <tr><th>Check</th>{before_th}<th>严重度</th><th>状态</th><th>自动修</th><th>说明</th></tr>
    {''.join(summary_rows)}
  </table>
</section>

<section id="structure">
  <h2>🏗️ 文档结构</h2>
  <div class="kv">
    <dt>总段落</dt><dd>{total_paras}</dd>
    <dt>标题分布</dt><dd>{' · '.join(f'{_esc(k)}:{_esc(v)}' for k, v in (h_count or {{}}).items() if not isinstance(v, (dict, list, set))) or '—'}</dd>
    <dt>Caption 分布(按样式)</dt><dd>{' · '.join(f'{_esc(k)}:{_esc(v)}' for k, v in (cap_by_style or {{}}).items()) or '—'}</dd>
    <dt>可用 caption 样式</dt><dd>{' '.join(f'<span class=tag>{_esc(s)}</span>' for s in (caption_styles_avail or [])) or '—'}</dd>
  </div>
</section>

<section id="checks">
  <h2>🩺 8 病种详情</h2>
  <p style="color:var(--muted)">点击展开每条查看说明。<strong>FOUND</strong> 项默认展开,按 (检出优先 → 严重度) 排序。</p>
  {cards_html}
</section>

{cap_table_section}
{tbl_section}
{img_section}
{bm_section}

<section id="fix-sop">
  <h2>🔧 修复 SOP (按推荐顺序)</h2>
  <ol>
    <li><strong>先 dry-run 一次 full</strong> 看 phase1 状况,不要直接 fix:
      <code class="fix-cmd">python3 docx_cli.py health full {_esc(docx_path.name)} --dry-run --html</code>
    </li>
    <li><strong>safe 病种走 health fix</strong> (revision/field/outline-pollution/heading-number 自动修):
      <code class="fix-cmd">python3 docx_cli.py health fix {_esc(docx_path.name)}</code>
    </li>
    <li><strong>plan-only 病种</strong> (heading-gap / duplicate-figure-numbers) 看 diagnose JSON 后人工处理:
      <code class="fix-cmd">python3 docx_cli.py health diagnose {_esc(docx_path.name)} --report /tmp/h.json</code>
    </li>
    <li><strong>体积 / 孤儿媒体 / orphan 表名</strong>不在 8 病种里,需单独走 audit + caption 命令族。</li>
    <li><strong>完成后 re-diagnose</strong> 看 exit_code 是否归 0。</li>
  </ol>
</section>

</main>

<aside class="backlinks">
  <h3>引擎</h3>
  <a href="#">docx_cli.py health</a>
  <h3>明细 JSON 位置</h3>
  <p style='font-size:12px;color:var(--muted)'><code>{_esc(tmp_dir) if tmp_dir else '/tmp/docx_health_*/heading_audit.json'}</code></p>
  <p style='font-size:12px;color:var(--muted)'><code>/tmp/audit-images-{_esc(docx_path.stem)}.json</code></p>
</aside>

<footer>
  Generated by <code>docx_health_render_html.py</code> @ {now}
  · Engine: <code>doctools/scripts/document/sub/health.py</code>
</footer>

</div></body></html>
"""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(body, encoding="utf-8")
    return out_path


# ─── 兼容入口 (旧 simple HTML 风格) ──────────────────────────────────────────

def render_simple_html(
    diagnose_result: dict,
    docx_path: str | Path,
    out_path: str | Path,
    rc: int,
    before: dict | None = None,
) -> Path:
    """旧版简陋单表 HTML (向后兼容)。"""
    docx_path = Path(docx_path)
    out_path = Path(out_path)
    sev_color = {"High": "#e74c3c", "Med": "#e67e22", "Low": "#f1c40f"}
    rows = ""
    for c in CHECKS_DOC:
        cid, _, _, _, autofix, sev, _ = c
        res = diagnose_result.get(cid, {})
        found = res.get("found", False)
        color = sev_color.get(sev, "#aaa")
        badge = f'<span style="color:{color};font-weight:bold">{sev}</span>'
        safe = "✅ auto" if "safe auto-fix" in autofix else "❌ plan"
        status = ('<span style="color:#e74c3c">⚠ FOUND</span>' if found
                  else '<span style="color:#27ae60">ok</span>')
        before_cell = ""
        if before is not None:
            bf = before.get(cid, {}).get("found", False)
            before_cell = f'<td>{"⚠ FOUND" if bf else "ok"}</td>'
        detail = res.get("fix_hint", "") if found else ""
        rows += (
            f"<tr><td><code>{cid}</code></td>{before_cell}"
            f"<td>{status}</td><td>{badge}</td><td>{safe}</td>"
            f"<td style='font-size:0.85em;color:#555'>{_esc(detail)}</td></tr>\n"
        )
    before_col = "<th>Before</th>" if before is not None else ""
    html = f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="UTF-8">
<title>docx health: {_esc(docx_path.name)}</title>
<style>
  body{{font-family:system-ui,sans-serif;max-width:900px;margin:2em auto;padding:1em;}}
  table{{border-collapse:collapse;width:100%;}}
  th,td{{border:1px solid #ddd;padding:6px 10px;text-align:left;}}
  th{{background:#f5f5f5;}}
  tr:nth-child(even){{background:#fafafa;}}
  h2{{color:#2c3e50;}}
</style>
</head><body>
<h2>docx health report</h2>
<p><b>File:</b> {_esc(docx_path.name)}<br>
<b>Exit code:</b> {rc} (0=healthy / 1=warning / 2=error)</p>
<table>
<tr><th>Check</th>{before_col}<th>Status</th><th>Severity</th><th>AutoFix</th><th>Detail</th></tr>
{rows}
</table>
<p style="font-size:0.8em;color:#aaa;margin-top:2em">Generated by doctools health.py</p>
</body></html>"""
    out_path.write_text(html, encoding="utf-8")
    return out_path
