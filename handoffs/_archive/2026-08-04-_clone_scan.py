#!/usr/bin/env python3
"""clone_scan — 全仓函数级克隆检测（结构归一化后比对），给出重构上限的 ground truth。

判据：把每个函数体 ast.unparse 成规范形式，再把**标识符与字符串常量**统一替换成占位符，
剩下的就是「结构骨架」。骨架相同 = 结构级克隆（这正是「N 个分支只差数据」的形状）。

    python3 clone_scan.py            # 结构克隆分组
    python3 clone_scan.py --exact    # 只看逐字节相同的函数体
"""
from __future__ import annotations

import ast
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path("/Users/tianli/Dev/tools/doctools")
SCAN = [ROOT / "scripts", ROOT / "lib", ROOT / "tools"]
MIN_LINES = 6           # 5 行以下的函数抽出来不划算，直接不看


class Norm(ast.NodeTransformer):
    """把标识符和字面量抹平，只留结构。"""

    def visit_Name(self, n):
        return ast.copy_location(ast.Name(id="V", ctx=n.ctx), n)

    def visit_arg(self, n):
        n.arg = "V"
        n.annotation = None
        return n

    def visit_Attribute(self, n):
        self.generic_visit(n)
        n.attr = "A"
        return n

    def visit_Constant(self, n):
        return ast.copy_location(ast.Constant(value="C"), n)

    def visit_FunctionDef(self, n):
        self.generic_visit(n)
        n.name = "F"
        n.decorator_list = []
        n.returns = None
        return n


def bodies():
    for root in SCAN:
        for p in sorted(root.rglob("*.py")):
            if "__pycache__" in str(p):
                continue
            src = p.read_text(encoding="utf-8", errors="replace")
            try:
                tree = ast.parse(src)
            except SyntaxError:
                continue
            for node in ast.walk(tree):
                if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                span = (node.end_lineno or node.lineno) - node.lineno + 1
                if span < MIN_LINES:
                    continue
                yield p, node, span, src


def key_of(node, exact: bool) -> str:
    clone = ast.parse(ast.unparse(node))
    if not exact:
        clone = Norm().visit(clone)
        ast.fix_missing_locations(clone)
    return ast.unparse(clone)


def main() -> int:
    exact = "--exact" in sys.argv
    groups: dict[str, list[tuple[Path, str, int]]] = defaultdict(list)
    total = 0
    for p, node, span, _src in bodies():
        total += 1
        groups[key_of(node, exact)].append((p.relative_to(ROOT), node.name, span, node.lineno))
    dup = {k: v for k, v in groups.items() if len(v) >= 2}
    # 抽出来的净收益 ≈ (组内实例数 − 1) × 平均行数
    ranked = sorted(dup.items(), key=lambda kv: -(len(kv[1]) - 1) * (sum(x[2] for x in kv[1]) / len(kv[1])))
    saved = sum(int((len(v) - 1) * (sum(x[2] for x in v) / len(v))) for v in dup.values())
    print(f"函数总数(≥{MIN_LINES} 行) {total} · 克隆组 {len(dup)} · "
          f"理论净删上限 {saved} 行  [{'逐字节相同' if exact else '结构相同'}]")
    print()
    for k, v in ranked[:22]:
        avg = sum(x[2] for x in v) / len(v)
        print(f"── {len(v)} 处 × ~{avg:.0f} 行 → 省 ~{int((len(v)-1)*avg)}")
        for path, name, span, ln in v[:6]:
            print(f"     {path}:{ln}  {name}()  {span} 行")
        if len(v) > 6:
            print(f"     … 另 {len(v)-6} 处")
    return 0


if __name__ == "__main__":
    sys.exit(main())
