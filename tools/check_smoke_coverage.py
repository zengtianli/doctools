#!/usr/bin/env python3
"""check_smoke_coverage.py — 冒烟覆盖闸门：把 `tests/smoke/_verb_specs.py` 那张表
跟真实 CLI 子命令树对账。

存在意义与 `check_function_axis.py` 同构，也同样是 fail-closed：一张手维护的清单
只要没人跟 CLI 对账，它就会漂移，而漂移的那天没有任何人会发现 —— 表还在，
只是少了几条。少的那几条恰恰是新加的、最没跑过的。

    表里缺一条（CLI 有、smoke 表没有）  → exit 1，点名缺哪条
    表里多一条（表有、CLI 没有）        → exit 1，点名陈旧条目
    同一 parser 被写了两遍（别名）      → exit 1
    parser 建不起来 / 枚举为空          → exit 2（拒绝在空集上报绿）

别名不占表项：`read`=extract · `diff`=compare · `styleset *`=audit-styleset *
与被别名者是**同一个 parser 对象**，按 `id(parser)` 归组后共用一条契约。
（从真实 parser 派生，不是手抄的别名清单 —— 以后谁加别名都自动认领。）

    python3 tools/check_smoke_coverage.py            # 对账 + 打覆盖率
    python3 tools/check_smoke_coverage.py --list     # 逐条列出（含 skip 理由）
    python3 tools/check_smoke_coverage.py --md       # 贴进文档的表格
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import Optional

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
SMOKE = REPO / "tests" / "smoke"


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"[smoke-coverage] cannot spec {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _cli_leaves() -> "list[tuple[int, str]]":
    """枚举 CLI 全部叶子子命令 → [(parser 身份, 路径)]。

    复用 `cli_surface._load_cli()`，而不是自己再写一遍加载 —— 几个闸门看的必须是
    同一棵树，各造各的 loader 就会出现「这个闸门绿、那个闸门红」这种谁也说不清的分叉。
    """
    import argparse as _ap

    cs = _load("_cs_for_smoke_cov", HERE / "cli_surface.py")
    mod = cs._load_cli()
    build = getattr(mod, "build_parser", None) or getattr(mod, "_build_parser", None)
    if build is None:
        raise SystemExit("[smoke-coverage] docx_cli.py 没有 build_parser/_build_parser")
    root = build()

    out: list[tuple[int, str]] = []

    def walk(parser, path: list[str]) -> None:
        subs = [a for a in parser._actions if isinstance(a, _ap._SubParsersAction)]
        if not subs:
            if path:
                out.append((id(parser), " ".join(path)))
            return
        for a in subs:
            for name, sp in a.choices.items():
                walk(sp, path + [name])

    walk(root, [])
    return out


def _reconcile(specs_mod) -> "tuple[int, list[dict], list[str]]":
    """返回 (退出码, 已对上的行, 错误行)。"""
    leaves = _cli_leaves()
    if not leaves:
        return 2, [], ["CLI 子命令枚举为空 —— 空集与满集都『无差异』，这种绿不算通过。"]
    if not specs_mod.VERBS:
        return 2, [], ["_verb_specs 动词表为空 —— 同上，空表对上空集不算覆盖。"]

    groups: dict[int, list[str]] = {}
    for pid, path in leaves:
        groups.setdefault(pid, []).append(path)

    errors: list[str] = []
    matched: list[dict] = []
    used: set[str] = set()

    for _pid, members in groups.items():
        hits = [p for p in members if p in specs_mod.SPECS]
        if not hits:
            paths = " / ".join(sorted(members))
            errors.append(f"缺契约：`{paths}` 在 CLI 里可跑，但 _verb_specs 表里没有")
            continue
        if len(hits) > 1:
            dup = " / ".join(sorted(hits))
            errors.append(f"别名重复立契约：`{dup}` 是同一个 parser，只能有一条表项")
            # 这些键在 CLI 里确实有对应 parser，别再被下面的「陈旧条目」二次点名 ——
            # 一个手误报三条错，读的人会先去追那两条假的。
            used.update(hits)
            continue
        canon = hits[0]
        used.add(canon)
        spec = specs_mod.SPECS[canon]
        matched.append({
            "verb": canon,
            "aliases": sorted(p for p in members if p != canon),
            "argv": spec["argv"],
            "expect_rc": spec["expect_rc"],
            "mutates": spec["mutates"],
            "note": spec["note"],
            "skip_reason": spec["skip_reason"],
        })

    for verb in specs_mod.VERBS:
        if verb not in used:
            errors.append(f"陈旧条目：表里有 `{verb}`，CLI 里没有这条子命令")

    matched.sort(key=lambda r: r["verb"])
    return (1 if errors else 0), matched, errors


def _print_list(rows: list[dict]) -> None:
    for r in rows:
        flag = "SKIP" if r["skip_reason"] else ("写" if r["mutates"] else "读")
        alias = f"  (别名 {', '.join(r['aliases'])})" if r["aliases"] else ""
        print(f"  [{flag}] rc={r['expect_rc']}  {r['verb']}{alias}")
        if r["skip_reason"]:
            print(f"         理由: {r['skip_reason']}")


def _print_md(rows: list[dict]) -> None:
    print("| 动词 | argv | 预期 rc | 改源件 | 状态 |")
    print("|---|---|---|---|---|")
    for r in rows:
        argv = " ".join(r["argv"])
        state = "skip" if r["skip_reason"] else "跑"
        print(f"| `{r['verb']}` | `{argv}` | {r['expect_rc']} | "
              f"{'是' if r['mutates'] else '否'} | {state} |")


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="check_smoke_coverage.py",
        description="冒烟覆盖闸门：_verb_specs 表 vs 真实 CLI 子命令树（fail-closed）",
    )
    ap.add_argument("--list", action="store_true", help="逐条列出（含 skip 理由）")
    ap.add_argument("--md", action="store_true", help="输出 markdown 表格")
    args = ap.parse_args()

    if not SMOKE.is_dir():
        print(f"[smoke-coverage] 找不到 {SMOKE} —— 冒烟表整个不在了，拒绝报绿",
              file=sys.stderr)
        return 2

    try:
        specs = _load("_verb_specs_gate", SMOKE / "_verb_specs.py")
    except Exception as e:
        print(f"[smoke-coverage] 载入 _verb_specs.py 失败：{type(e).__name__}: {e}",
              file=sys.stderr)
        return 2

    try:
        rc, rows, errors = _reconcile(specs)
    except SystemExit:
        raise
    except Exception as e:
        print(f"[smoke-coverage] 枚举 CLI 失败：{type(e).__name__}: {e}", file=sys.stderr)
        return 2

    if rc == 2:
        for e in errors:
            print(f"[smoke-coverage] {e}", file=sys.stderr)
        return 2

    if errors:
        print(f"✗ 冒烟覆盖对账失败：{len(errors)} 处", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
        print("修法：在 tests/smoke/_verb_specs.py 的 _ROWS 里增/删对应行"
              "（新增的那条要先真跑一遍拿到 rc，别照抄邻居）。", file=sys.stderr)
        return 1

    n_total = len(rows)
    n_skip = sum(1 for r in rows if r["skip_reason"])
    n_run = n_total - n_skip
    n_alias = sum(len(r["aliases"]) for r in rows)

    if args.md:
        _print_md(rows)
    elif args.list:
        _print_list(rows)

    pct = 100.0 * n_run / n_total
    print(f"✓ 冒烟覆盖对账通过：真跑 {n_run} 条 / skip {n_skip} 条 / 共 {n_total} 条"
          f"（{pct:.1f}% 真跑；另 {n_alias} 条别名共用同一 parser，不单独占表项）",
          file=sys.stderr)
    if n_skip:
        for r in rows:
            if r["skip_reason"]:
                print(f"  skip: {r['verb']} —— {r['skip_reason']}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
