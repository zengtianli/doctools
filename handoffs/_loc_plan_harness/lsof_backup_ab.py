#!/usr/bin/env python3
"""lsof_backup_ab.py — W2 的等价对拍：10 份 lsof_check + 5 份备份路径实现是否真同义。

    python3 handoffs/_loc_plan_harness/lsof_backup_ab.py     # rc=0 即三项全过

三项判据（任何一项不过就别做 W2 的那一半）：
  1. 5 份 A 派 lsof + `_cli_common.lsof_check` → 输出必须只有 **1 种**
  2. 3 份 B 派 lsof → 输出必须只有 **1 种**，且 A.strip() == B（证明只差 strip，
     但**不许据此合并** —— B 多的那道 `len(lines) > 1` 守卫在这个样本上碰不到）
  3. 5 份备份路径实现 × 7 种文件名 × 3 轮 N 自增 → 0 mismatch

⚠ 必须 import 生产实现（`spec_from_file_location` 按路径加载），
禁自己按「我以为它是这么实现的」重写一个替身来测 —— 那测的是替身（铁律 #2）。
"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SUB = ROOT / "scripts" / "document" / "sub"

A_SITES = [("outline", "lsof_check"), ("styles", "lsof_check"),
           ("fix_styleset", "_lsof_check"), ("slim", "_lsof_check"),
           ("pipeline_lib", "lsof_check")]
B_SITES = [("add_header_footer", "lsof_check"), ("caption", "lsof_check"),
           ("blocks", "lsof_check")]
BACKUP_SITES = [("outline", "pick_backup_path"), ("styles", "pick_backup_path"),
                ("fix_styleset", "_pick_backup_path"), ("slim", "_find_next_backup"),
                ("pipeline_lib", "make_backup_path")]
NAMES = ["a.docx", "b.c.docx", "报告 v2.final.docx", "x.DOCX",
         "no-ext", ".hidden.docx", "x.tar.gz"]

def load(name: str):
    """按 `sub.<name>` 包路径导入。

    ⚠ 别用 `spec_from_file_location(name, path)` —— `fix_styleset.py` 等文件用的是
    相对导入（`from .shape_contract import …`），脱离包加载当场
    `ImportError: attempted relative import with no known parent package`。
    """
    return importlib.import_module(f"sub.{name}")


def main() -> int:
    sys.path.insert(0, str(ROOT / "scripts" / "document"))      # 让 `sub` 成为可导入的包
    sys.path.insert(0, str(SUB))
    sys.path.insert(0, str(ROOT / "lib"))
    import _cli_common as cc                                    # noqa: E402

    bad = 0

    # ── 3) 备份路径 ─────────────────────────────────────────────────
    mism = 0
    with tempfile.TemporaryDirectory() as d:
        for n in NAMES:
            p = Path(d) / n
            for _ in range(3):
                ref = cc.find_next_backup(p)
                for mod, fn in BACKUP_SITES:
                    got = getattr(load(mod), fn)(p)
                    if str(got) != str(ref):
                        mism += 1
                        print(f"MISMATCH backup {mod}.{fn} {n}: {got} != {ref}")
                ref.write_bytes(b"x")
    print(f"backup mismatches = {mism}   (必须 0)")
    bad += bool(mism)

    # ── 1)/2) lsof：真开一个 fd 让 lsof 有输出 ────────────────────────
    target = Path("/etc/hosts")
    with open(target):
        a_out = {str(cc.lsof_check(target))}
        for mod, fn in A_SITES:
            a_out.add(str(getattr(load(mod), fn)(target)))
        b_out = {str(getattr(load(mod), fn)(target)) for mod, fn in B_SITES}
    print(f"A distinct = {len(a_out)} | B distinct = {len(b_out)}   (各必须 1)")
    bad += (len(a_out) != 1) + (len(b_out) != 1)
    if len(a_out) == 1 and len(b_out) == 1:
        a, b = next(iter(a_out)), next(iter(b_out))
        same = a.strip() == b
        print(f"A.strip() == B : {same}   (必须 True；**但这不是合并 A/B 的依据**)")
        bad += not same

    print("✓ 全过" if not bad else f"⛔ {bad} 项不过")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
