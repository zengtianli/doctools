#!/usr/bin/env python3
"""allowed_deltas_ab.py — W3 的等价对拍：全零 allowed_deltas 是不是真 no-op。

    python3 handoffs/_loc_plan_harness/allowed_deltas_ab.py    # rc=0 即 24000 例全等价

判据：对随机快照对 (before, after)，`diff_structure(b, a)` 与
`diff_structure(b, a, allowed_deltas=<全零形状>)` 必须逐条相同。差一例就别删。

⚠ 必须 import 生产的 `sub/shape_contract.py`，禁按「我以为 diff_structure 是这么写的」
重写一个替身 —— 那测的是替身，会假绿（铁律 #2）。

⚠ 202 个 pytest 里**一条都没覆盖** shape_contract / fix_styleset，
所以「套件还是 202 passed」不能作为 W3 的验收证据。
"""
from __future__ import annotations

import importlib.util
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SUB = ROOT / "scripts" / "document" / "sub"

# fix_styleset.py 里 8 处 allowed_deltas 的三种键形状（5 / 6 / 7 键），值全 0
SHAPES = [
    {"para_count": 0, "table_count": 0, "drawings_count": 0, "section_count": 0,
     "heading_counts": 0, "figure_numbers": 0, "table_numbers": 0},
    {"para_count": 0, "table_count": 0, "drawings_count": 0,
     "heading_counts": 0, "figure_numbers": 0, "table_numbers": 0},
    {"para_count": 0, "table_count": 0, "heading_counts": 0,
     "figure_numbers": 0, "table_numbers": 0},
]
ROUNDS = 8000


def snap(rnd: random.Random) -> dict:
    return {
        "para_count": rnd.randint(0, 50),
        "table_count": rnd.randint(0, 9),
        "drawings_count": rnd.randint(0, 9),
        "section_count": rnd.randint(0, 4),
        "track_changes_count": rnd.randint(0, 3),
        "heading_counts": {f"H{i}": rnd.randint(0, 7) for i in range(1, rnd.randint(2, 5))},
        "figure_numbers": {f"图{i}-1" for i in range(rnd.randint(0, 6))},
        "table_numbers": {f"表{i}-1" for i in range(rnd.randint(0, 6))},
    }


def main() -> int:
    sys.path.insert(0, str(SUB))
    spec = importlib.util.spec_from_file_location("shape_contract", SUB / "shape_contract.py")
    sc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sc)                       # 生产实现，不是替身

    rnd = random.Random(0)
    diffs = 0
    for _ in range(ROUNDS):
        b, a = snap(rnd), snap(rnd)
        base = sc.diff_structure(b, a)
        for sh in SHAPES:
            if sc.diff_structure(b, a, allowed_deltas=sh) != base:
                diffs += 1
    total = ROUNDS * len(SHAPES)
    print(f"comparisons = {total} · differences = {diffs}   (必须 0)")
    if not total:                                     # fail-closed：空集不报绿
        print("⛔ 一例都没跑 —— 判据坏了", file=sys.stderr)
        return 2
    return 1 if diffs else 0


if __name__ == "__main__":
    sys.exit(main())
