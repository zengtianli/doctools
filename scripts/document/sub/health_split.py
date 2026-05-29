"""health_split.py — one-shot docx 健康化 + 按 H1 切分 thin wrapper.

Distilled 2026-05-28 from `~/Work/projects/eco-flow/taizhou-天台/业务模板/SOP.md`
Step 1 + Step 2 (pipeline run --steps health-diagnose,audit-styleset-all,split-by-h1
   → 若不达锚: styleset restore --no-llm → 用新文件重跑 split).

This module is a thin orchestrator only — all real work delegates to:
  - pipeline_lib.run_pipeline (_BUILTIN_STEPS: health-diagnose / audit-styleset-all
    / split-by-h1)  — single parse, multi-step
  - fix_styleset.cmd_restore (--no-llm 强制, CC agent 内嵌套死锁)

Subcommand:
  docx_cli.py health-split <docx>
      [--out-dir <dir>]      默认 <docx-stem>-split/ (docx 同目录)
      [--backup-dir <dir>]   默认 ~/Archives/docx-backups/$(date +%F)/
      [--no-backup]          跳过备份 (危险)
      [--no-fix]             不达锚也不跑 styleset restore (危险)
      [--no-frontmatter]     split 时不输出 00-frontmatter.docx
      [--report <path>]      健康报告 HTML 输出, 默认 <out-dir>/_health-report.html

Health 判据 (磐安 v3 锚):
  audit-styleset-all 5 类 ≥ 4 PASS + ≤ 1 WARN + 0 FAIL
    AND health-diagnose overall_severity == "pass"
  → healthy, skip styleset restore.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import shutil
import sys
import time
from pathlib import Path

from .pipeline_lib import run_pipeline


def _judge_healthy(report: dict) -> tuple[bool, dict]:
    """Decide healthy from pipeline report.

    Returns (healthy, summary) where summary contains pass/warn/fail counts
    + health-diagnose severity for printing.
    """
    steps = report.get("steps", {})
    audit = steps.get("audit-styleset-all", {}) or {}
    health = steps.get("health-diagnose", {}) or {}

    checks = audit.get("checks", {}) or {}
    counts = {"pass": 0, "warn": 0, "fail": 0}
    for _name, r in checks.items():
        sev = (r or {}).get("severity", "pass")
        counts[sev] = counts.get(sev, 0) + 1

    health_sev = health.get("overall_severity", "fail")
    healthy = (
        counts["fail"] == 0 and counts["pass"] >= 4 and counts["pass"] + counts["warn"] >= 5 and health_sev == "pass"
    )
    return healthy, {
        "audit_counts": counts,
        "health_severity": health_sev,
    }


def _write_html_report(report: dict, summary: dict, out_path: Path) -> None:
    """Minimal self-contained HTML health report."""
    payload = {
        "timestamp": _dt.datetime.now().isoformat(timespec="seconds"),
        "summary": summary,
        "pipeline_report": report,
    }
    body = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    html = (
        "<!doctype html><html><head><meta charset='utf-8'>"
        f"<title>docx health-split report</title></head><body>"
        f"<h1>docx health-split report</h1>"
        f"<p>generated: {payload['timestamp']}</p>"
        f"<pre style='font:12px/1.4 monospace;background:#f6f8fa;"
        f"padding:12px;border-radius:6px'>{body}</pre>"
        "</body></html>"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")


def _run(args) -> int:
    src = Path(args.docx).expanduser().resolve()
    if not src.is_file():
        print(f"[health-split] ERROR: not a file: {src}", file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else src.parent / f"{src.stem}-split"
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. backup
    if not args.no_backup:
        bdir = (
            Path(args.backup_dir).expanduser()
            if args.backup_dir
            else Path.home() / "Archives" / "docx-backups" / _dt.date.today().isoformat()
        )
        bdir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, bdir / src.name)
        print(f"✓ 备份: {bdir / src.name}")

    t0 = time.perf_counter()

    # 2. speculative single-pass: run health+audit+split in one parse.
    #    若 audit/diagnose 判 healthy → 直接收工 (省 1 次 parse, SOP baseline 33s)
    #    若 unhealthy → 删除已写章节, 走 styleset restore → 用新文件重 split
    one_shot_ns = argparse.Namespace(
        docx=src,
        dry_run=False,
        no_backup=True,
        report=None,
        styleset_profile=None,
        split_out_dir=out_dir,
        split_name_pattern=None,
        include_frontmatter=not args.no_frontmatter,
        allow_no_h1=False,
        split_dry_run=False,
    )
    diag_report = run_pipeline(
        src,
        ["health-diagnose", "audit-styleset-all", "split-by-h1"],
        args=one_shot_ns,
        no_backup=True,
    )
    healthy, summary = _judge_healthy(diag_report)
    print(
        f"[health-split] health: audit={summary['audit_counts']} "
        f"diagnose={summary['health_severity']} → "
        f"{'healthy' if healthy else 'unhealthy'}"
    )

    target = src
    split_report = diag_report

    # 3. conditional styleset restore (unhealthy 才走第二轮)
    if not healthy and not args.no_fix:
        from . import fix_styleset

        # 丢弃 speculative split 产物, 用 restored docx 重切
        for f in out_dir.glob("*.docx"):
            f.unlink()
        new_name = f"{src.stem}-health-v3-{_dt.date.today().isoformat()}.docx"
        target = src.parent / new_name
        restore_ns = argparse.Namespace(
            docx_path=src,
            dry_run=False,
            inplace=False,
            no_backup=True,
            force=False,
            output=target,
            no_llm=True,
            yaml_profile=None,
            report=None,
        )
        rc = fix_styleset.cmd_restore(restore_ns)
        if rc != 0:
            print(f"[health-split] styleset restore FAIL rc={rc}", file=sys.stderr)
            return rc
        print(f"⚠ 不达锚, 已健康化 → {target.name}")
        split_ns = argparse.Namespace(
            docx=target,
            dry_run=False,
            no_backup=True,
            report=None,
            split_out_dir=out_dir,
            split_name_pattern=None,
            include_frontmatter=not args.no_frontmatter,
            allow_no_h1=False,
            split_dry_run=False,
        )
        split_report = run_pipeline(
            target,
            ["split-by-h1"],
            args=split_ns,
            no_backup=True,
        )
    elif not healthy:
        print("⚠ 不达锚 + --no-fix, 已 split 原 docx (用户自担风险)")

    # 5. report + summary
    report_path = Path(args.report).expanduser() if args.report else out_dir / "_health-report.html"
    merged = {
        "src": str(src),
        "target": str(target),
        "out_dir": str(out_dir),
        "diagnose": diag_report,
        "split": split_report,
    }
    _write_html_report(merged, summary, report_path)

    # 6. (optional) fold by chapter: 每个 docx 挪进同名子目录, 便于下游
    #    pipeline image-extract/table-extract 自动落到 <章>/images,tables/ 章内聚
    if args.fold_by_chapter:
        folded = 0
        for f in sorted(out_dir.glob("*.docx")):
            sub = out_dir / f.stem
            sub.mkdir(exist_ok=True)
            f.rename(sub / f.name)
            folded += 1
        print(f"[health-split] fold-by-chapter: {folded} 章已挪进同名子目录")

    wall = time.perf_counter() - t0
    split_step = split_report.get("steps", {}).get("split-by-h1") or {}
    n_chapters = split_step.get("slices_emitted") or len(split_step.get("emitted", []) or []) or "?"
    print(f"\n[health-split] DONE 章数={n_chapters} 墙钟={wall:.1f}s out_dir={out_dir} report={report_path}")
    return 0


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "health-split",
        help="docx 一键健康化 + 按 H1 切分 (backup + diagnose + audit + 条件 styleset restore + split)",
    )
    p.add_argument("docx", help="input docx path")
    p.add_argument("--out-dir", default=None, help="split 输出目录 (default: <docx-stem>-split/ 同目录)")
    p.add_argument("--backup-dir", default=None, help="备份目录 (default: ~/Archives/docx-backups/<today>/)")
    p.add_argument("--no-backup", action="store_true", help="跳过备份 (危险)")
    p.add_argument("--no-fix", action="store_true", help="health 不达锚也不跑 styleset restore (危险)")
    p.add_argument("--no-frontmatter", action="store_true", help="split 时不输出 00-frontmatter.docx")
    p.add_argument("--fold-by-chapter", dest="fold_by_chapter", action="store_true", default=True,
                   help="split 后每章 docx 挪进同名子目录 (便于下游 image/table extract 章内聚) [DEFAULT ON]")
    p.add_argument("--no-fold-by-chapter", dest="fold_by_chapter", action="store_false",
                   help="保留平铺产出 (老 SOP 兼容; 用 split/*.docx 一级 glob 的下游)")
    p.add_argument("--report", default=None, help="健康报告 HTML 输出路径 (default: <out-dir>/_health-report.html)")
    p.set_defaults(func=_run)
