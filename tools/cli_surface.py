#!/usr/bin/env python3
"""cli_surface.py — 把 docx_cli 的全量 argparse 子命令树 dump 成稳定 JSON。

存在意义是**等价闸门**：把 23 个转发 shim 折成一张声明表时，唯一能证明「没弄丢
子命令、没改选项名、没改默认值」的东西，就是折之前和折之后这棵树逐字节相同。
读代码觉得等价不算数。

    python3 tools/cli_surface.py > before.json
    …重构…
    python3 tools/cli_surface.py > after.json && diff before.json after.json

fail-closed：解析不到 parser、或树里子命令数为 0，一律非零退出 —— 空树和满树
diff 起来一样「无差异」，那是最会骗人的一种绿。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DOC = HERE.parent / "scripts" / "document"


def _load_cli():
    sys.path.insert(0, str(DOC))
    spec = importlib.util.spec_from_file_location("_docx_cli_surface", DOC / "docx_cli.py")
    if spec is None or spec.loader is None:
        raise SystemExit("cannot spec docx_cli.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["_docx_cli_surface"] = mod
    spec.loader.exec_module(mod)
    return mod


def _norm(v: str) -> str:
    """把仓根绝对路径换成占位符。有几个默认值是 `HERE / "profiles" / …` 算出来的，
    在 git worktree 里跑基线时前缀不同 —— 那是运行位置的差异，不是接口的差异，
    留着它每次 diff 都要人眼分辨真假，迟早看漏一条真的。"""
    return v.replace(str(HERE.parent), "<REPO>")


def _action_sig(a: argparse.Action) -> dict:
    return {
        "opts": sorted(a.option_strings),
        "dest": a.dest,
        "nargs": a.nargs if not isinstance(a.nargs, int) else f"int:{a.nargs}",
        "default": _norm(repr(a.default)),
        "required": bool(a.required),
        "choices": sorted(map(str, a.choices)) if a.choices else None,
        "cls": type(a).__name__,
    }


def walk(parser: argparse.ArgumentParser, path: str = "") -> dict:
    node = {"opts": [], "subs": {}}
    for a in parser._actions:
        if isinstance(a, argparse._SubParsersAction):
            # 子命令组自身的 dest/required/metavar 也要进指纹。少了它，「把一个
            # 本来可省的子命令改成必填」这种改动闸门看不见 —— 而那正是重写
            # register() 时最容易顺手改掉的一项（required 的默认值是 False，
            # 手写时漏掉一个 required=True 不会有任何报错）。
            node["subparsers"] = {"dest": a.dest, "required": bool(a.required),
                                  "metavar": a.metavar}
            for name, sub in a.choices.items():
                node["subs"][name] = walk(sub, f"{path} {name}".strip())
        else:
            node["opts"].append(_action_sig(a))
    node["opts"].sort(key=lambda d: (d["dest"], ",".join(d["opts"])))
    # 互斥组也进指纹：--keep-shell-h1 / --no-keep-shell-h1 那种成对选项，
    # 拆出互斥关系之后两个都能同时传，argparse 不再拦。
    mx = [sorted(x.dest for x in g._group_actions)
          for g in getattr(parser, "_mutually_exclusive_groups", [])]
    if mx:
        node["mutually_exclusive"] = sorted(mx)
    return node


def count(node: dict) -> int:
    return len(node["subs"]) + sum(count(v) for v in node["subs"].values())


def main() -> int:
    mod = _load_cli()
    build = getattr(mod, "build_parser", None) or getattr(mod, "_build_parser", None)
    if build is None:
        print("docx_cli.py 没有 build_parser/_build_parser —— 无法取到 parser，判红。",
              file=sys.stderr)
        return 2
    tree = walk(build())
    n = count(tree)
    if n == 0:
        print("子命令树为空 —— 空树 diff 恒等于空树，这种绿不算通过。", file=sys.stderr)
        return 3
    print(json.dumps({"subcommands": n, "tree": tree}, ensure_ascii=False,
                     indent=1, sort_keys=True))
    print(f"[cli_surface] 子命令 {n} 个", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
