"""_groups.py — docx_cli 子命令组的**声明表**，取代一批只做转发的 group 模块。

## 为什么有这个

2026-07-30 之前，`sub/` 下有 20 个文件只干一件事：声明 argparse 子命令，然后
`exec_script("<真实现>", argv)` 转发过去。它们不碰 docx、没有业务逻辑，加起来
2265 行，信息量等于一张「组名 / 子命令 / 实现脚本 / 选项」的表 —— 而这张表被
拆成 20 份手写的 `register()`，每份都自己重写一遍同样的四个标准选项。

后果不是「文件多」这么轻。是**同一件事写了 20 遍**：给全部子命令加一个
`--json` 要改 20 处；`_rest_argv` 的转发规则变一次要核对 20 处；新加一个组
只能靠复制粘贴上一个组（`strip.py` 的 `-o/--output` 分支就是这么长出来的）。

现在：**表在这里，构建器只有一个**（`register_all`）。加组 = 加一行数据。

## 边界：哪些没收进来，为什么

- **`chapters_sync` / `health_split` / `pipeline`** —— 名字像壳，其实有真逻辑
  （反向回写章节、一键健康化编排、批量 driver）。它们有 `register()` 不代表是壳。
- **手搓 argv 的那 12 个**（`table` / `split` / `images` / `fonts` …）—— 它们把
  自己的 argparse 选项翻译成实现脚本的 flag，翻译规则本身是数据，但每个组不同。
  收编要先给 opts 建模，属于下一步，不在本表。
- 收编进来的组，其**独立入口仍然可用**：`python3 sub/<script>.py <docx> …`。
  本表只接管 `docx_cli.py` 这条路。

## 等价性怎么保证

`tools/cli_surface.py` dump 全量子命令树（选项名/dest/nargs/默认值/required/choices）。
折之前 128 个子命令，折之后必须逐字节相同。读代码觉得等价不算数 —— 这类改动
出错的形状正是「少了一个子命令，谁也没发现」。
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field

from ._dispatch import exec_script, _rest_argv, get_or_add_group, get_or_add_subparsers


@dataclass(frozen=True)
class Group:
    """一个子命令组。字段全是数据，没有回调 —— 一旦允许挂回调，表就会重新长成代码。"""

    name: str                       # 组名，即 `docx_cli.py <name> <target>` 的 name
    help: str                       # 组的一行说明
    dest: str                       # argparse dest，存放选中的 target
    targets: dict[str, str]         # 子命令 → 实现脚本名（无 .py，住在 sub/）
    shared: bool = False            # True = 这个组名还有别的模块往里加 target，
                                    #        必须走 get_or_add_* 而不是 add_parser
    target_help: str = ""           # 子命令 help 模板，`{t}` 占位；空 = "<组名> <target>"
    rest_help: str = "extra args forwarded to underlying script"
    with_output: frozenset[str] = field(default_factory=frozenset)
                                    # 这些 target 额外认 -o/--output（写新文件，不动原件）
    deprecated: str = ""            # 非空 = 调用时先打这行警告


# ── 表 ────────────────────────────────────────────────────────────────────────
# 顺序 = 注册顺序 = `--help` 里的显示顺序。shared 组的顺序还决定谁先建组，
# 但 get_or_add_group 让先后无所谓（后来者复用先建的）。
GROUPS: list[Group] = [
    Group(
        "audit",
        "audit-only docx checks (headings/fields/captions/images/table-pairing/bookmarks)",
        "audit_target",
        {
            "headings":      "audit_heading_numbers",
            "fields":        "audit_word_fields",
            "captions":      "audit_caption_outline",
            "images":        "audit_images",
            "table-pairing": "audit_table_pairing",
            "bookmarks":     "audit_bookmarks",
        },
        target_help="audit {t} (read-only)",
    ),
    Group(
        "freeze",
        "freeze auto-computed Word elements (headings numbering / fields)",
        "freeze_target",
        {"headings": "freeze_heading_numbers", "fields": "freeze_all_fields"},
        rest_help="extra args (e.g. --types TOC,PAGEREF)",
    ),
    Group(
        "strip",
        "strip stale/polluting docx elements (outlinelvl / bookmarks / revisions / "
        "doc-protection / orphan-media / empty-captions)",
        "strip_target",
        {
            "outlinelvl":       "strip_outlinelvl_from_captions",
            "style-outlinelvl": "strip_style_outlinelvl",
            "bookmarks":        "strip_bookmarks",
            "revisions":        "strip_revisions",
            "doc-protection":   "strip_doc_protection",
            "orphan-media":     "strip_orphan_media",
            "empty-captions":   "strip_empty_captions",
        },
        with_output=frozenset({"orphan-media", "empty-captions"}),
    ),
    Group(
        "header-footer",
        "add/manage docx header & footer (院标准格式)",
        "hf_target",
        {"add": "add_header_footer"},
        target_help="header-footer {t}",
        rest_help="forwarded args (--header / --footer-prefix / --page-number / "
                  "--font-size / --gap-spaces / --header-align / --footer-align)",
    ),
    Group(
        "legacy",
        "DEPRECATED catch-all scripts (fix-heading-disorder)",
        "legacy_target",
        {"fix-heading-disorder": "fix_heading_disorder"},
        target_help="[DEPRECATED] {t}",
        rest_help="extra args forwarded",
        deprecated="[sub.legacy] WARNING: '{t}' is DEPRECATED; see sub/_groups.py "
                   "GROUPS 表注释 for the recommended single-purpose replacement scripts.",
    ),
    # ↓ 以下三个是 shared 组：同名组里还有别的模块注册的 target
    #   （caption pair ← captions.py · caption number-by-style ← styles.py ·
    #    chapter delete ← blocks.py · renumber h4-figures ← styles.py）
    Group(
        "caption",
        "caption ops (number / pair / number-by-style)",
        "caption_target",
        {"number": "number_captions"},
        shared=True,
    ),
    Group(
        "chapter",
        "chapter ops (convert-arabic / delete-empty-h1 / delete)",
        "chapter_target",
        {"convert-arabic": "convert_chapter_format", "delete-empty-h1": "delete_empty_h1"},
        shared=True,
    ),
    Group(
        "renumber",
        "renumber headings + caption numbers",
        "renumber_target",
        {"headings": "renumber_headings"},
        shared=True,
    ),
]

BY_NAME: dict[str, Group] = {g.name: g for g in GROUPS}

# 表自检：表是唯一事实源，它自己错了下面全错，所以 import 期就炸。
_seen: set[tuple[str, str]] = set()
for _g in GROUPS:
    if not _g.targets:
        raise SystemExit(f"[GROUPS 表错] {_g.name}: targets 为空 —— 空组注册出来是个"
                         f"点不动的子命令，比没有还坏")
    for _t in _g.targets:
        if (_g.name, _t) in _seen:
            raise SystemExit(f"[GROUPS 表错] 重复的子命令 {_g.name} {_t}")
        _seen.add((_g.name, _t))
    if not _g.with_output <= set(_g.targets):
        raise SystemExit(f"[GROUPS 表错] {_g.name}: with_output 里有不存在的 target "
                         f"{sorted(_g.with_output - set(_g.targets))}")
del _seen


# ── 构建器 ────────────────────────────────────────────────────────────────────

def _run(args) -> int:
    """所有组共用的转发入口。组名从 `_group_name` 拿（register 时 set_defaults 钉死），
    不靠猜 dest —— 猜的那版遇上 shared 组会拿错 target。"""
    g = BY_NAME.get(getattr(args, "_group_name", ""))
    if g is None:
        print(f"[sub._groups] 未知组: {getattr(args, '_group_name', None)!r}")
        return 2
    target = getattr(args, g.dest, None)
    script = g.targets.get(target)
    if script is None:
        # shared 组里别的模块注册的 target 走不到这里（它们自带 func），
        # 所以走到这里就是真的对不上。
        print(f"[sub.{g.name}] unknown target: {target}; choices={list(g.targets)}")
        return 2
    if g.deprecated:
        print(g.deprecated.format(t=target), flush=True)
    argv = _rest_argv(args)
    if target in g.with_output:
        out = getattr(args, "output", None)
        if out:
            argv.extend(["-o", str(out)])
    return exec_script(script, argv)


def _add_target(sp, g: Group, t: str) -> None:
    help_tmpl = g.target_help or (g.name + " {t}")
    spp = sp.add_parser(t, help=help_tmpl.format(t=t), add_help=False)
    spp.add_argument("docx_path", nargs="?", help="target docx path")
    spp.add_argument("--dry-run", action="store_true")
    spp.add_argument("--no-backup", action="store_true")
    spp.add_argument("--report", help="write JSON report to this path")
    if t in g.with_output:
        spp.add_argument("-o", "--output", default=None,
                         help="write to new path (do not modify original, no bak)")
    spp.add_argument("rest", nargs=argparse.REMAINDER, help=g.rest_help)
    spp.set_defaults(func=_run, _group_name=g.name)


def register_group(subparsers, name: str) -> None:
    """注册**单个**组。为什么要按名注册而不是一次性注册全部：顶层 `--help` 里组的
    先后 = 注册先后，而本表接管的 8 个组在原来的注册序列里是**散着**的（中间夹着
    audit_styleset / outline / blocks / images 这些没被接管的）。一次性注册会把它们
    挤成一坨，用户敲 `--help` 看到的顺序就变了 —— 那是没人要求过的改动。"""
    g = BY_NAME.get(name)
    if g is None:
        raise KeyError(f"GROUPS 表里没有组 {name!r}；有的是 {list(BY_NAME)}")
    _register_one(subparsers, g)


def register_all(subparsers) -> None:
    """把 GROUPS 全部注册上去（本仓不走这条，留给独立使用者）。"""
    for g in GROUPS:
        _register_one(subparsers, g)


def _register_one(subparsers, g: Group) -> None:
    if g.shared:
        p = get_or_add_group(subparsers, g.name, g.help)
        sp = get_or_add_subparsers(p, dest=g.dest)
        existing = getattr(sp, "choices", {}) or {}
    else:
        p = subparsers.add_parser(g.name, help=g.help)
        sp = p.add_subparsers(dest=g.dest, metavar="<target>", required=True)
        existing = {}
    for t in g.targets:
        if t in existing:
            continue            # 已被同组的另一个模块注册（shared 组的常态）
        _add_target(sp, g, t)
