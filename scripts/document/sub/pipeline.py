"""pipeline.py — docx batch pipeline driver (distilled from qual-supply · 2026-05-26)

暴露 `def run(args)` 供 docx_cli.py `pipeline run` 子命令调用。
逻辑上等价于 qual-supply `scripts/run_pipeline.py`，但以函数形式上提到总部。

用法（通过 docx_cli.py 调用）
-----------------------------
    # 单 docx 多 step（parse+save 各 1 次）
    docx_cli.py pipeline run docs/X.docx \\
        --steps freeze_heading_numbers,apply_caption_styles --step-dir scripts/

    # 多 docx 并行（ProcessPoolExecutor 满核）
    docx_cli.py pipeline run docs/A.docx docs/B.docx docs/C.docx \\
        --steps apply_body_styles,apply_table_styles --parallel --step-dir scripts/

    # dry-run 不写盘（用于墙钟基线）
    docx_cli.py pipeline run docs/*.docx --steps ... --dry-run --parallel --step-dir scripts/

注意：--step-dir 默认 cwd/scripts/（qual-supply 兼容），或用 --step-dir 指定任意项目目录。
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from .pipeline_lib import (
    format_timing_table,
    run_pipeline,
    run_pipeline_parallel,
)


def parse_steps(s: str) -> list[str]:
    out = [t.strip() for t in s.split(",") if t.strip()]
    if not out:
        raise argparse.ArgumentTypeError("--steps 不能为空")
    return out


def register(subparsers) -> None:
    """Register `pipeline run` subcommand onto docx_cli.py's top-level subparsers."""
    # pipeline group
    pipeline_p = subparsers.add_parser(
        "pipeline",
        help="docx batch pipeline (multi-step, multi-docx, parallel)",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    pipeline_sub = pipeline_p.add_subparsers(dest="pipeline_cmd", metavar="<pipeline-cmd>")
    pipeline_p.set_defaults(func=lambda a: (pipeline_p.print_help() or 0))

    # pipeline run
    run_p = pipeline_sub.add_parser(
        "run",
        help="顺序/并行执行指定 steps（parse+save 各 1 次）",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    run_p.add_argument("docx", nargs="+", help="目标 docx 路径列表")
    run_p.add_argument("--steps", type=parse_steps, required=True,
                       help="脚本名列表（逗号分隔，无 .py 后缀）")
    run_p.add_argument("--step-dir", type=Path, default=None,
                       help="step 脚本目录（默认 cwd/scripts/）")
    run_p.add_argument("--parallel", action="store_true",
                       help="跨 docx 并行（ProcessPoolExecutor）")
    run_p.add_argument("--max-workers", type=int, default=None,
                       help="并发数（默认 min(N_docx, cpu_count)）")
    run_p.add_argument("--dry-run", action="store_true",
                       help="不写盘（用于基线测时 / 验证）")
    run_p.add_argument("--no-backup", action="store_true",
                       help="跳过自动备份")
    run_p.add_argument("--report-dir", type=Path, default=None,
                       help="把每 docx 的 report JSON 落到此目录")
    # 破坏性 step 选项（透传给 step 的 args）
    run_p.add_argument("--delete-h1", default=None)
    run_p.add_argument("--delete-h1-text", default=None)
    run_p.add_argument("--relocate-plan", default=None)
    run_p.add_argument("--pair-decision", default=None)
    run_p.add_argument("--relink-source", default=None)
    run_p.add_argument("--freeze-levels", default=None)
    run_p.add_argument("--unlink-style", action="store_true", default=True)
    run_p.add_argument("--strip-styles", default=None)
    run_p.add_argument("--header", default="嘉兴市千岛湖引水受水区分质供水管理项目研究")
    run_p.add_argument("--footer-prefix", default="浙江省水利水电勘测设计院")
    run_p.add_argument("--page-number", action="store_true", default=True)
    run_p.set_defaults(func=run)


def run(args: argparse.Namespace) -> int:
    """Entry point for `pipeline run`. Called by docx_cli.py dispatcher."""
    docx_paths = [Path(d).resolve() for d in args.docx]
    missing = [p for p in docx_paths if not p.is_file()]
    if missing:
        print(f"ERROR: 文件不存在: {missing}", file=sys.stderr)
        return 2

    step_dir = getattr(args, "step_dir", None)
    if step_dir is None:
        # default: cwd/scripts/ (qual-supply compat)
        cwd_scripts = Path.cwd() / "scripts"
        if cwd_scripts.is_dir():
            step_dir = cwd_scripts

    print(f"[pipeline] {len(docx_paths)} docx × {len(args.steps)} steps "
          f"({'parallel' if args.parallel else 'serial'} mode, dry_run={args.dry_run})")
    print(f"[pipeline] steps: {' → '.join(args.steps)}")
    if step_dir:
        print(f"[pipeline] step-dir: {step_dir}")
    for p in docx_paths:
        print(f"  - {p.name}")

    t0 = time.perf_counter()

    parallel = getattr(args, "parallel", False)
    max_workers = getattr(args, "max_workers", None)
    dry_run = getattr(args, "dry_run", False)
    no_backup = getattr(args, "no_backup", False)

    if len(docx_paths) == 1 or not parallel:
        results = {}
        for p in docx_paths:
            ns_args = argparse.Namespace(
                docx=p, dry_run=dry_run, no_backup=True, report=None,
                delete_h1=getattr(args, "delete_h1", None),
                delete_h1_text=getattr(args, "delete_h1_text", None),
                relocate_plan=getattr(args, "relocate_plan", None),
                pair_decision=getattr(args, "pair_decision", None),
                relink_source=getattr(args, "relink_source", None),
                freeze_levels=getattr(args, "freeze_levels", None),
                unlink_style=getattr(args, "unlink_style", True),
                strip_styles=getattr(args, "strip_styles", None),
                header=getattr(args, "header", ""),
                footer_prefix=getattr(args, "footer_prefix", "浙江省水利水电勘测设计院"),
                page_number=getattr(args, "page_number", True),
            )
            try:
                rep = run_pipeline(
                    p, args.steps, args=ns_args,
                    dry_run=dry_run, no_backup=no_backup,
                    step_dir=step_dir,
                )
                results[str(p)] = rep
            except Exception as exc:
                results[str(p)] = {"error": repr(exc)}
    else:
        if any([
            getattr(args, "delete_h1", None),
            getattr(args, "delete_h1_text", None),
            getattr(args, "relocate_plan", None),
            getattr(args, "pair_decision", None),
            getattr(args, "relink_source", None),
        ]):
            print("[WARN] 破坏性 step 参数在并行模式下未透传; 建议用 serial 模式（去 --parallel）",
                  file=sys.stderr)
        results = run_pipeline_parallel(
            docx_paths, args.steps,
            max_workers=max_workers,
            dry_run=dry_run, no_backup=no_backup,
            step_dir=step_dir,
        )

    wall = time.perf_counter() - t0

    print()
    print(format_timing_table(results))
    print(f"\n[pipeline] 总墙钟: {wall:.3f}s")
    if parallel:
        ser_sum = sum(
            r.get("timing", {}).get("total", 0.0)
            for r in results.values() if isinstance(r, dict)
        )
        if wall > 0:
            print(f"[pipeline] 串行 work-time 总和: {ser_sum:.3f}s; "
                  f"并行加速比 ≈ {ser_sum / wall:.2f}x")

    # 写每 docx report
    report_dir = getattr(args, "report_dir", None)
    if report_dir:
        report_dir.mkdir(parents=True, exist_ok=True)
        for docx, rep in results.items():
            stem = Path(docx).stem
            out = report_dir / f"pipeline-{stem}.json"
            out.write_text(
                json.dumps(rep, ensure_ascii=False, indent=2, default=str),
                encoding="utf-8",
            )
        print(f"[pipeline] reports -> {report_dir}")

    if wall > 300:
        print(f"\n[FAIL] 墙钟 {wall:.1f}s > 300s 硬上限!", file=sys.stderr)
        return 1
    if wall > 30 and len(docx_paths) >= 4:
        print(f"\n[WARN] 墙钟 {wall:.1f}s > 30s 软目标（4 docx 期望）; 检查 step 选择 / 并行度",
              file=sys.stderr)

    any_err = any(
        isinstance(r, dict) and ("error" in r) for r in results.values()
    )
    return 1 if any_err else 0
