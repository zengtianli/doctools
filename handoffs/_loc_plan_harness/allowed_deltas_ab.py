#!/usr/bin/env python3
"""allowed_deltas_ab.py — W3 的等价对拍：全零 allowed_deltas 是不是真 no-op。

    python3 handoffs/_loc_plan_harness/allowed_deltas_ab.py    # rc=0 即全部等价

判据：对随机快照对 (before, after)，`diff_structure(b, a)` 与
`diff_structure(b, a, allowed_deltas=<fix_styleset 里那 8 个真实形状>)` 必须逐条相同。

## 两处都不许手抄（2026-08-02 血的教训）

本脚本第一版**手抄了字段名**，写成 `para_count` / `figure_numbers` / `table_numbers`，
而生产字段其实叫 `paragraph_count` / `figure_number_set` / `table_number_set`。
后果：`before.get(f, 0)` 对每个标量都返 0 → 所有 delta 都是 0 → 天然无违规 →
**24000 例全绿，却什么都没测到**。比不测更坏，因为它看起来像测过了。

所以现在：
  ① 字段名从 `shape_contract._SCALAR_FIELDS` **取**（不是抄）
  ② 8 个 allowed_deltas 形状从 `fix_styleset.py` 源码 **ast 抽**（不是抄）
  ③ 断言快照真的产生了违规 —— 否则「两边都返空列表」是废等价，直接判红

⚠ 必须 import 生产的 `shape_contract`，禁按「我以为它是这么实现的」重写替身（铁律 #2）。
⚠ 202 个 pytest **一条都没覆盖** shape_contract / fix_styleset，
   所以「套件还是 202 passed」不能作为 W3 的验收证据。
"""
from __future__ import annotations

import ast
import importlib.util
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SUB = ROOT / "scripts" / "document" / "sub"
ROUNDS = 4000


def load_shape_contract():
    spec = importlib.util.spec_from_file_location("shape_contract", SUB / "shape_contract.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def real_shapes() -> list[dict]:
    """从 fix_styleset.py 源码里把 8 个 allowed_deltas 字面量抽出来。"""
    tree = ast.parse((SUB / "fix_styleset.py").read_text(encoding="utf-8"))
    out = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            for kw in node.keywords:
                if kw.arg == "allowed_deltas":
                    try:
                        out.append(ast.literal_eval(kw.value))
                    except (ValueError, SyntaxError):
                        pass
    return out


def main() -> int:
    sc = load_shape_contract()                       # 生产实现，不是替身
    fields = list(sc._SCALAR_FIELDS)                 # 字段名取自生产，不手抄
    shapes = real_shapes()

    if not fields:                                   # fail-closed
        print("⛔ 取不到 shape_contract._SCALAR_FIELDS —— 判据坏了", file=sys.stderr)
        return 2

    if not shapes:
        # 空集在这里有**确定含义**：W3 已经施工，8 个全零 kwarg 不存在了。
        # 但不能就这么报绿走人 —— 零容忍的保证已经转移给一条 pytest，
        # 那条不在就是真的没人守了。
        successor = ROOT / "scripts" / "document" / "tests" / "test_shape_contract_default.py"
        if not successor.is_file():
            print(f"⛔ fix_styleset 里已无全零 allowed_deltas，而接手的守卫 "
                  f"{successor} 不存在 —— 零容忍现在没人守", file=sys.stderr)
            return 2
        print("ℹ W3 已施工：fix_styleset 里不再有全零 allowed_deltas。")
        print(f"  零容忍的保证已转移给 {successor.relative_to(ROOT)}（跑 pytest 验它）。")
        return 0

    nonzero = [s for s in shapes if any(v for v in s.values())]
    if nonzero:
        print(f"⛔ {len(nonzero)} 个形状含非零值，前提不成立，不能删：{nonzero}", file=sys.stderr)
        return 2

    rnd = random.Random(0)

    def snap() -> dict:
        d = {f: rnd.randint(0, 40) for f in fields}
        d["heading_counts"] = {f"H{i}": rnd.randint(0, 7) for i in range(1, rnd.randint(2, 5))}
        d["figure_number_set"] = [f"图{i}-1" for i in range(rnd.randint(0, 6))]
        d["table_number_set"] = [f"表{i}-1" for i in range(rnd.randint(0, 6))]
        return d

    diffs = with_violations = 0
    for _ in range(ROUNDS):
        b, a = snap(), snap()
        base = sc.diff_structure(b, a)
        if base:
            with_violations += 1
        for sh in shapes:
            if sc.diff_structure(b, a, allowed_deltas=sh) != base:
                diffs += 1

    total = ROUNDS * len(shapes)
    print(f"字段 {len(fields)} 个 · 形状 {len(shapes)} 个 · 比对 {total} 次")
    print(f"其中 base 真的报出违规的轮次 = {with_violations}/{ROUNDS}   (必须 > 0，否则是废等价)")
    print(f"differences = {diffs}   (必须 0)")

    if not with_violations:
        print("⛔ 一轮违规都没造出来 —— 两边都返空列表的「等价」什么都没证明", file=sys.stderr)
        return 2
    return 1 if diffs else 0


if __name__ == "__main__":
    sys.exit(main())
