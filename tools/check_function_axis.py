#!/usr/bin/env python3
"""check_function_axis.py — 职能轴闸门：把 `sub/_function_axis.py` 那张表跟真实
CLI 子命令树对账。

存在意义：职能标签只要是「另一张手维护的清单」，它就会跟 CLI 漂移，而漂移的那天
没有任何人会发现 —— 表还在，只是少了几条、多了几条。所以对账做成 fail-closed：

    表里缺一条（CLI 有、表没有）        → exit 1，点名缺哪条
    表里多一条（表有、CLI 没有）        → exit 1，点名陈旧条目
    同一 parser 被打了两次标签（别名）  → exit 1
    parser 建不起来 / 枚举为空          → exit 2（拒绝在空集上报绿）

别名不占表项：`read`=extract · `diff`=compare · `styleset *`=audit-styleset *
与被别名者是**同一个 parser 对象**，按 `id(parser)` 归组后共用一条标签。
（这是从真实 parser 派生的，不是手抄的别名清单 —— 以后谁加别名都自动认领。）

    python3 tools/check_function_axis.py                # 对账
    python3 tools/check_function_axis.py --fn format    # 只列「只碰格式」的
    python3 tools/check_function_axis.py --md           # 贴进文档的表格
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import Optional

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
DOC = REPO / "scripts" / "document"


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"[function-axis] cannot spec {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _cli_leaves() -> "list[tuple[int, str, tuple[str, Optional[str]]]]":
    """枚举 CLI 的全部叶子子命令 → [(parser 身份, 路径, (顶层名, 动作))]。

    复用 `cli_surface._load_cli()` 而不是自己再写一遍加载 —— 两个闸门看的必须是
    同一棵树，各造各的 loader 就会出现「surface 绿、职能轴红」这种谁也说不清的分叉。
    """
    import argparse as _ap

    cs = _load("_cs_for_fn_axis", HERE / "cli_surface.py")
    mod = cs._load_cli()
    build = getattr(mod, "build_parser", None) or getattr(mod, "_build_parser", None)
    if build is None:
        raise SystemExit("[function-axis] docx_cli.py 没有 build_parser/_build_parser")
    root = build()

    out: list[tuple[int, str, tuple[str, Optional[str]]]] = []

    def walk(parser, path: list[str]) -> None:
        subs = [a for a in parser._actions if isinstance(a, _ap._SubParsersAction)]
        if not subs:
            if path:
                top = path[0]
                action = path[1] if len(path) > 1 else None
                out.append((id(parser), " ".join(path), (top, action)))
            return
        for a in subs:
            for name, sp in a.choices.items():
                walk(sp, path + [name])

    walk(root, [])
    return out


def _reconcile(axis_mod) -> "tuple[int, list[dict], list[str]]":
    """返回 (退出码, 已对上的行, 错误行)。"""
    leaves = _cli_leaves()
    if not leaves:
        return 2, [], ["CLI 子命令枚举为空 —— 空集与满集都『无差异』，这种绿不算通过。"]

    # 按 parser 身份归组：同一 parser 挂多个名字 = 别名
    groups: dict[int, list[tuple[str, tuple[str, Optional[str]]]]] = {}
    for pid, path, key in leaves:
        groups.setdefault(pid, []).append((path, key))

    errors: list[str] = []
    matched: list[dict] = []
    used_keys: set[tuple[str, Optional[str]]] = set()

    for pid, members in groups.items():
        hits = [(p, k) for p, k in members if k in axis_mod.AXIS]
        if not hits:
            paths = " / ".join(p for p, _ in members)
            errors.append(f"缺标签：`{paths}` 在 CLI 里可跑，但 _function_axis 表里没有")
            continue
        if len(hits) > 1:
            dup = " / ".join(p for p, _ in hits)
            errors.append(f"别名重复打标签：`{dup}` 是同一个 parser，只能有一条表项")
            # 这些键确实在 CLI 里有对应 parser，别再被下面的「陈旧条目」二次点名 ——
            # 一个手误报三条错，读的人会先去追那两条假的。
            used_keys.update(k for _, k in hits)
            continue
        canon_path, canon_key = hits[0]
        used_keys.add(canon_key)
        fn, note = axis_mod.AXIS[canon_key]
        matched.append({
            "path": canon_path,
            "top": canon_key[0],
            "action": canon_key[1],
            "fn": fn,
            "note": note,
            "aliases": sorted(p for p, _ in members if p != canon_path),
        })

    for key in axis_mod.AXIS:
        if key not in used_keys:
            errors.append(
                f"陈旧条目：表里有 `{axis_mod.path_of(*key)}`，CLI 里没有这条子命令"
            )

    matched.sort(key=lambda r: (list(axis_mod.FN_TAGS).index(r["fn"]), r["path"]))
    return (1 if errors else 0), matched, errors


def _print_md(rows: list[dict]) -> None:
    print("| 职能 | 子命令 | 别名 | 备注 |")
    print("|---|---|---|---|")
    for r in rows:
        alias = ", ".join(f"`{a}`" for a in r["aliases"]) or "—"
        print(f"| {r['fn']} | `{r['path']}` | {alias} | {r['note'] or '—'} |")


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="check_function_axis.py",
        description="职能轴闸门：_function_axis 表 vs 真实 CLI 子命令树（fail-closed）",
    )
    ap.add_argument("--fn", metavar="<职能>", help="只列该职能的子命令")
    ap.add_argument("--md", action="store_true", help="输出 markdown 表格")
    args = ap.parse_args()

    try:
        axis = _load("_function_axis_gate", DOC / "sub" / "_function_axis.py")
    except Exception as e:
        print(f"[function-axis] 载入 _function_axis.py 失败：{type(e).__name__}: {e}",
              file=sys.stderr)
        return 2

    if args.fn is not None and args.fn not in axis.FN_DEFINITIONS:
        print(f"[function-axis] 未知职能 {args.fn!r}；取值域 {', '.join(axis.FN_TAGS)}",
              file=sys.stderr)
        return 2

    try:
        rc, rows, errors = _reconcile(axis)
    except SystemExit:
        raise
    except Exception as e:
        print(f"[function-axis] 枚举 CLI 失败：{type(e).__name__}: {e}", file=sys.stderr)
        return 2

    if rc == 2:
        for e in errors:
            print(f"[function-axis] {e}", file=sys.stderr)
        return 2

    if errors:
        print(f"✗ 职能轴对账失败：{len(errors)} 处", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
        print("修法：在 scripts/document/sub/_function_axis.py 的 _ROWS 里增/删对应行。",
              file=sys.stderr)
        return 1

    shown = [r for r in rows if args.fn is None or r["fn"] == args.fn]
    if args.md:
        _print_md(shown)
    elif args.fn is not None:
        print(f"# {args.fn} —— {axis.FN_DEFINITIONS[args.fn]}")
        for r in shown:
            tail = f"  ({r['note']})" if r["note"] else ""
            print(f"  {r['path']}{tail}")
    else:
        by_fn: dict[str, int] = {t: 0 for t in axis.FN_TAGS}
        for r in rows:
            by_fn[r["fn"]] += 1
        for tag in axis.FN_TAGS:
            print(f"  {tag:<9}{by_fn[tag]:>4}  {axis.FN_DEFINITIONS[tag]}")

    n_alias = sum(len(r["aliases"]) for r in rows)
    print(f"✓ 职能轴对账通过：{len(rows)} 条子命令全部有标签"
          f"（另 {n_alias} 条别名共用同一 parser，不单独占表项）"
          f"{'；本次只列 ' + args.fn + ' ' + str(len(shown)) + ' 条' if args.fn else ''}",
          file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
