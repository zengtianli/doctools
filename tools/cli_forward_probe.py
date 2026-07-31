#!/usr/bin/env python3
"""cli_forward_probe.py — 录下每条子命令**转发给实现脚本的 argv**，供折叠前后对拍。

`cli_surface.py` 证明的是「接口长得一样」，证明不了「参数传过去还是一样」——
而把 18 个手写转发函数折成一张表，动的恰恰是后者。两者缺一不可：
接口对了 argv 错了 = 命令能敲、跑出来的东西不对，且不报错。

    python3 tools/cli_forward_probe.py > before.json     # 折叠前（git worktree 里跑）
    python3 tools/cli_forward_probe.py > after.json
    diff before.json after.json

实现方式：把 exec_script 换成录音机（不真跑脚本），然后逐条 parse_args + 调 func。
patch 的位置是 `sys.modules` 里**所有** sub.* 模块上的 exec_script 名字 ——
因为各模块是 `from ._dispatch import exec_script` 绑进自己命名空间的，
只 patch `_dispatch.exec_script` 对它们无效（那是最容易自欺的一种 patch）。

fail-closed：调用样本里只要有一条 parse 不出来或没有 func，整体非零退出。
静默跳过一条 = 这条的转发从此没人验，而它恰恰可能就是坏的那条。
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DOC = HERE.parent / "scripts" / "document"

D = "/tmp/probe.docx"          # 不会真打开——exec_script 被换掉了
G = "/tmp/golden.docx"

# 每条 = 一次完整调用。覆盖被折叠的全部组，重点在不规则的那几条：
# table borders --center（链式跑第二个脚本）、chrome --validate（短路）、
# section read --list（顶掉位置参数 query）、split body-replace（互斥默认真值）。
CASES: list[list[str]] = [
    ["audit", "headings", D, "--report", "/tmp/r.json"],
    ["audit", "bookmarks", D],
    ["audit", "table-pairing", D, "--dry-run"],
    ["freeze", "headings", D, "--no-backup"],
    ["freeze", "fields", D, "--dry-run", "--", "--types", "TOC,PAGEREF"],
    ["strip", "bookmarks", D, "--dry-run"],
    ["strip", "revisions", D, "--no-backup", "--report", "/tmp/r.json"],
    ["strip", "orphan-media", D, "-o", "/tmp/out.docx"],
    ["strip", "empty-captions", D, "--output", "/tmp/out.docx"],
    ["header-footer", "add", D, "--", "--header", "标题", "--footer-prefix", "院"],
    ["legacy", "fix-heading-disorder", D, "--dry-run"],
    ["blocks", "reorder", D],
    ["blocks", "relocate", D, "--plan", "/tmp/p.json"],
    ["chapter", "convert-arabic", D],
    ["chapter", "delete-empty-h1", D, "--dry-run"],
    ["chapter", "delete", D, "--h1", "3", "--dry-run"],
    ["chapter", "delete", D, "--prefix", "附录", "--no-backup"],
    ["caption", "number", D, "--dry-run"],
    ["caption", "pair", D, "--decision", "/tmp/d.json", "--dry-run"],
    ["renumber", "headings", D],
    ["image", "relink", D, "--source", G, "--dry-run"],
    ["image", "relink", D, "--apply-patch", "/tmp/patch.json"],
    ["image", "extract", D, "--out-dir", "/tmp/imgs", "--quiet"],
    ["seqdiff", "seq", "--src", D, "--dst", G, "--out", "/tmp/r.md"],
    ["seqdiff", "seq", "--src", D, "--dst", G, "--out", "/tmp/r.md", "--noise-len", "0"],
    ["seqdiff", "image", "--src", D, "--dst", G, "--out", "/tmp/r.md"],
    ["compare-ref", "ref", "--drafts-dir", "/tmp/md", "--ref", G, "--out", "/tmp/r.md"],
    ["compare-ref", "ref", "--drafts-dir", "/tmp/md", "--glob", "*.md"],
    ["revise-rules", "gen", "--drafts-dir", "/tmp/md", "--docx", D, "--out", "/tmp/r.json"],
    ["revise-rules", "gen", "--drafts-dir", "/tmp/md", "--no-title-skip", "--no-paren-strip"],
    ["section", "read", D, "--list"],
    ["section", "read", D, "第三章"],
    ["split", "by-h1", "--docx", D, "--out-dir", "/tmp/o", "--dry-run"],
    ["split", "by-h1", "--docx", D, "--out-dir", "/tmp/o",
     "--name-pattern", "{idx}.docx", "--include-frontmatter", "--allow-no-h1"],
    ["split", "body-replace", "--shell", D, "--content", G, "--out", "/tmp/o.docx"],
    ["split", "body-replace", "--shell", D, "--content", G, "--out", "/tmp/o.docx",
     "--no-keep-shell-h1", "--dry-run"],
    ["fonts", "normalize", D, "--dry-run"],
    ["fonts", "normalize", "--docx", D, "--body-cjk", "仿宋", "--latin", "Arial", "--no-backup"],
    ["table", "delete-rows", D, "--table-index", "0", "--rows", "1:3"],
    ["table", "delete-rows", D, "--table-index", "2", "--rows", "1:3",
     "--expected-first-col", "a,b", "--expected-residue", "x", "--expected-last", "y"],
    ["table", "extract", D, "--out-dir", "/tmp/t"],
    ["table", "extract", "--docx", D, "--out-dir", "/tmp/t",
     "--name-pattern", "{idx}.docx", "--dry-run"],
    ["table", "borders", D, "--dry-run"],
    ["table", "borders", D, "--val", "double", "--sz", "8", "--color", "FF0000",
     "--space", "2", "--keep-cell-borders", "--no-backup"],
    ["table", "borders", D, "--center", "--cell-center", "--dry-run"],
    ["table", "center", D, "--cell-center", "--no-backup"],
    ["chrome", "--raw", D, "--template", G, "--out", "/tmp/o.docx", "--county", "景宁县"],
    ["chrome", "--validate", "/tmp/o.docx", G],
    ["md-merge", "/tmp/a.md", D, "0", "12"],
    ["md-merge", "/tmp/a.md", D, "--start-anchor", "第三章", "--in-place", "--no-backup"],
]

_calls: list[dict] = []


def _recorder(script: str, argv: list[str]) -> int:
    _calls.append({"script": str(script), "argv": [str(x) for x in argv]})
    return 0


def _patch_everywhere() -> int:
    """把所有已加载的 sub.* 模块上的 exec_script 换成录音机。返回替换处数。"""
    n = 0
    for name, mod in list(sys.modules.items()):
        if not (name == "sub" or name.startswith("sub.") or name.startswith("_doctools_sub")):
            continue
        if getattr(mod, "exec_script", None) is not None:
            mod.exec_script = _recorder
            n += 1
        disp = getattr(mod, "_dispatch", None)
        if disp is not None and getattr(disp, "exec_script", None) is not _recorder:
            disp.exec_script = _recorder
            n += 1
    return n


def main() -> int:
    sys.path.insert(0, str(DOC))
    spec = importlib.util.spec_from_file_location("_probe_cli", DOC / "docx_cli.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["_probe_cli"] = mod
    spec.loader.exec_module(mod)
    build = getattr(mod, "build_parser", None) or getattr(mod, "_build_parser", None)
    if build is None:
        print("拿不到 build_parser，判红", file=sys.stderr)
        return 2

    parser = build()
    if _patch_everywhere() == 0:
        print("一个 exec_script 都没 patch 到 —— 那说明探针根本没生效，"
              "录出来的空结果 diff 起来永远相等。判红。", file=sys.stderr)
        return 3

    out = []
    bad = []
    for case in CASES:
        _calls.clear()
        try:
            args = parser.parse_args(case)
        except SystemExit as e:
            bad.append({"case": case, "err": f"parse 失败 exit={e.code}"})
            continue
        fn = getattr(args, "func", None)
        if fn is None:
            bad.append({"case": case, "err": "没有 func —— 这条子命令敲下去会得到 no handler"})
            continue
        try:
            rc = fn(args)
        except Exception as e:   # noqa: BLE001
            bad.append({"case": case, "err": f"{type(e).__name__}: {e}"})
            continue
        out.append({"case": case, "rc": rc, "calls": list(_calls)})

    print(json.dumps({"ok": out, "bad": bad}, ensure_ascii=False, indent=1))
    print(f"[probe] 正常 {len(out)} 条 · 异常 {len(bad)} 条", file=sys.stderr)
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
