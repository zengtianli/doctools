#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_final.py — 标书终稿机械管线 driver（残留扫描 → 清理 → 身份门 → 付印门）。

用法: python3 bid_final.py <docx> --mode main|pei [--rules <yaml>] [--apply]

流程:
  ① bid_residue_scan（8 类残留只读扫描）
  ② --apply 时: bid_finalize_sweep --apply（确定性清理类1-7，备份+三护栏）→ residue 复扫须归零
  ③ subprocess 调 bid_identity_gate.py（第8类身份泄漏门）
  ④ subprocess 调 bid_print_ready.py（付印门）
  全绿 → 「终稿 PASS」+ sha256 前 12 位指纹；任一红 → exit 2 并汇总哪道门红。
③④ 引擎文件缺失 → 清晰报错 exit 1（用法/IO 错）。
exit 0 = PASS；exit 2 = 有门红；exit 1 = 用法/IO/引擎缺失。
"""
import argparse
import hashlib
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCAN = HERE / "bid_residue_scan.py"
SWEEP = HERE / "bid_finalize_sweep.py"
IDENTITY = HERE / "bid_identity_gate.py"
PRINT_READY = HERE / "bid_print_ready.py"


def run_engine(script, docx, mode, rules, extra=()):
    cmd = [sys.executable, str(script), str(docx), "--mode", mode]
    if rules:
        cmd += ["--rules", str(rules)]
    cmd += list(extra)
    print(f"── {script.name} {' '.join(cmd[3:])} " + "─" * max(4, 60 - len(script.name)))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.stdout:
        print(r.stdout.rstrip())
    if r.stderr:
        print(r.stderr.rstrip(), file=sys.stderr)
    return r.returncode


def main():
    ap = argparse.ArgumentParser(description="标书终稿机械管线 driver")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--mode", choices=["main", "pei"], default="pei")
    ap.add_argument("--rules", type=Path, default=None)
    ap.add_argument("--apply", action="store_true", help="执行确定性清理（否则只跑各道门）")
    args = ap.parse_args()

    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    if args.rules and not args.rules.is_file():
        print(f"错误: 规则文件不存在 {args.rules}", file=sys.stderr)
        return 1

    gates = {}  # 门名 → exit code

    # ① 残留扫描
    rc = run_engine(SCAN, args.docx, args.mode, args.rules)
    if rc == 1:
        print("residue_scan 用法/IO 错误，中止", file=sys.stderr)
        return 1

    # ② --apply: 确定性清理 + 复扫归零
    if args.apply:
        rc_sweep = run_engine(SWEEP, args.docx, args.mode, args.rules, extra=["--apply"])
        if rc_sweep == 1:
            print("finalize_sweep 用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates["finalize_sweep"] = rc_sweep
        rc = run_engine(SCAN, args.docx, args.mode, args.rules)  # 复扫须归零
        if rc == 1:
            print("residue_scan 复扫用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates["residue_rescan"] = rc
    else:
        gates["residue_scan"] = rc

    # ③④ 身份门 + 付印门（并行车道引擎；缺失 = 清晰错误 exit 1）
    for script, name in ((IDENTITY, "identity_gate"), (PRINT_READY, "print_ready")):
        if not script.is_file():
            print(f"错误: 引擎缺失 {script}（{name} 尚未就位，无法完成终稿门控）", file=sys.stderr)
            return 1
        rc = run_engine(script, args.docx, args.mode, args.rules)
        if rc == 1:
            print(f"{name} 用法/IO 错误，中止", file=sys.stderr)
            return 1
        gates[name] = rc

    # ── 汇总 ──
    red = [g for g, rc in gates.items() if rc != 0]
    print("═" * 68)
    print("门控汇总:", " · ".join(f"{g}={'✅' if rc == 0 else '❌'}" for g, rc in gates.items()))
    if red:
        print(f"红门: {', '.join(red)}")
        print(f"FAIL {len(red)} findings")
        return 2
    sha = hashlib.sha256(args.docx.read_bytes()).hexdigest()[:12]
    print(f"终稿 PASS · {args.docx.name} · sha256 {sha}")
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
