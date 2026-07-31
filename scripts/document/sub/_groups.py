"""_groups.py — docx_cli 子命令的**声明表**，取代 18 个只做转发的 group 模块。

## 为什么有这个

2026-07-30 之前，`sub/` 下有 18 个文件只干一件事：声明 argparse 子命令，然后
`exec_script("<真实现>", argv)` 把参数转发过去。它们不碰 docx、没有业务逻辑，
加起来 2000+ 行，信息量等于一张「组名 / 子命令 / 实现脚本 / 选项 / 怎么转发」的表。

后果不是「文件多」这么轻。是**同一件事写了 18 遍**：

- 给全部子命令加一个 `--json` 要改 18 处；转发规则变一次要核对 18 处。
- 新加一个组只能复制粘贴上一个组 —— `strip.py` 的 `-o/--output` 分支就是这么长出来的。
- 而且**错了不会有人发现**：`table delete-rows` 从写下那天起就是坏的，它只
  `set_defaults(_sub_target=...)` 没设 `func=`，敲下去只得到「no handler for table」。
  旁边那个本该接住它的 `handle()` 从来没有任何人调用。18 份手写代码里藏一个死子命令，
  没有任何闸门看得见。收进表之后，每个 target 都必然拿到 func。

现在：**表在这里，构建器只有一个**。加子命令 = 加一行数据。

## 边界：哪些没收进来，为什么

- **`chapters_sync` / `health_split` / `pipeline`** —— 名字像壳，其实有真逻辑
  （反向回写章节、一键健康化编排、批量 driver）。有 `register()` 不代表是壳。
- **`outline` / `styles` / `health` / `slim` / `fix_styleset` / `docx_para` /
  `combine` / `md_merge_track` / `audit_styleset`** —— 自带业务实现，本来就不是转发层。
- 收编进来的组，其**独立入口仍然可用**：`python3 sub/<script>.py <docx> …`。
  本表只接管 `docx_cli.py` 这条路。

## 等价性怎么保证

`tools/cli_surface.py` dump 全量子命令树（选项名 / dest / nargs / 默认值 / required /
choices / 子命令组的 dest+required+metavar / 互斥组）。折之前 128 个子命令，折之后
必须逐字节相同。读代码觉得等价不算数 —— 这类改动出错的形状正是「少了一个子命令，
谁也没发现」。实测它抓到过一次：`header-footer` 的 metavar 被我从 `<action>` 写成
`<target>`，那个差异不报错、只在 `--help` 里显示不同，肉眼永远看不出来。
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from typing import Any

from ._dispatch import exec_script, get_or_add_group, get_or_add_subparsers


# ──────────────────────────────── 选项模型 ────────────────────────────────

@dataclass(frozen=True)
class Opt:
    """一个选项：既描述怎么**声明**给 argparse，也描述怎么**转发**给实现脚本。

    两件事必须写在一起。原来它们分居两处（`register()` 里声明、`_run()` 里转发），
    于是「加了选项忘了转发」是这类壳最常见的 bug —— 用户传了参数，脚本收不到，
    静默按默认值跑。
    """

    name: str                       # "--src"；不以 - 开头 = 位置参数
    alt: str = ""                   # 第二个写法（如 -o 的长形式 --output）
    dest: str = ""                  # 空 = 按 argparse 规则从 name 推
    help: str = ""
    action: str = "store"           # store | store_true | store_false
    default: Any = None
    nargs: Any = None
    type_int: bool = False
    required: bool = False
    metavar: tuple[str, ...] = ()
    mx: str = ""                    # 非空 = 归入这个名字的互斥组

    # ── 转发 ──
    forward: str = ""               # 转发时用的 flag；空 = 同 name（位置参数则直接发值）
    fwd: str = "auto"               # auto | never | when_false | short_circuit
    alias_of: str = ""              # 本项为空时回落到这个 dest 取值（--docx / 位置参数二选一）
    suppressed_by: str = ""         # 这个 dest 为真时本项不发（--list 顶掉 query）
    emit_zero: bool = False         # 用 `is not None` 而不是真值判断（0 也要发）
    must_have: str = ""             # 非空 = 取不到值就报这句并 rc=2


def _dest_of(o: Opt) -> str:
    if o.dest:
        return o.dest
    base = o.alt or o.name
    return base.lstrip("-").replace("-", "_")


def _is_pos(o: Opt) -> bool:
    return not o.name.startswith("-")


# 绝大多数 docx 动作的标准四件套 + rest。原来这 5 行在 18 个文件里各抄了一遍。
STD: tuple[Opt, ...] = (
    Opt("docx_path", nargs="?", help="target docx path"),
    Opt("--dry-run", action="store_true"),
    Opt("--no-backup", action="store_true"),
    Opt("--report", help="write JSON report to this path"),
)
OUT_OPT = Opt("-o", alt="--output", dest="output",
              help="write to new path (do not modify original, no bak)", forward="-o")


@dataclass(frozen=True)
class Target:
    """一个子命令。`impl` 是 sub/ 下的实现脚本名（无 .py）。

    `impl_argv`：转发时**前插**在用户 argv 之前的固定 token（家族折叠后的
    子命令名，如 ("bookmarks",)）。2026-07-31 家族折叠起：strip_*/audit_*/
    freeze_* 等旧脚本并成单文件子命令族，impl 指合并文件、impl_argv 选子命令。
    """

    impl: str
    impl_argv: tuple[str, ...] = ()
    help: str = ""
    opts: tuple[Opt, ...] = STD
    description: str = ""
    add_help: bool = False
    rest_help: str | None = "extra args forwarded to underlying script"
                                    # None = 本子命令不收 REMAINDER
    chain: tuple[str, str] = ()     # (dest, 同组另一个 target) —— 该 dest 为真时接着跑
    deprecated: str = ""


@dataclass(frozen=True)
class Group:
    name: str
    help: str
    dest: str                       # 空 = 扁平组（没有子命令，选项直接挂组上）
    targets: dict[str, Target | str] = field(default_factory=dict)
                                    # 值给 str = 用 STD 选项、impl 就是这个 str
    shared: bool = False            # True = 组名可能已被别的模块建过，走 get_or_add_*
    target_help: str = ""           # 子命令 help 模板，`{t}` 占位
    metavar: str = "<target>"       # 子命令组的 metavar；header-footer 原本是 "<action>"
    required: bool = True           # 子命令是否必填（section / split 原本不必填）
    rest_help: str = "extra args forwarded to underlying script"
    flat: Target | None = None      # 扁平组的唯一 Target


def _targets(g: Group) -> dict[str, Target]:
    out: dict[str, Target] = {}
    for name, t in g.targets.items():
        if isinstance(t, str):
            t = Target(t, help=(g.target_help or (g.name + " {t}")).format(t=name),
                       rest_help=g.rest_help)
        elif not t.help:
            t = Target(**{**t.__dict__,
                          "help": (g.target_help or (g.name + " {t}")).format(t=name)})
        out[name] = t
    return out


# ──────────────────────────────── 表 ────────────────────────────────
# 顺序 = 注册顺序 = 顶层 `--help` 的显示顺序。

GROUPS: list[Group] = [
    # ── 纯标准四件套的组 ───────────────────────────────────────────────────
    Group(
        "audit",
        "audit-only docx checks (headings/fields/captions/images/table-pairing/bookmarks)",
        "audit_target",
        {
            "headings":      Target("audit", impl_argv=("headings",)),
            "fields":        Target("audit", impl_argv=("fields",)),
            "captions":      Target("audit", impl_argv=("captions",)),
            "images":        Target("audit", impl_argv=("images",)),
            "table-pairing": Target("audit", impl_argv=("table-pairing",)),
            "bookmarks":     Target("audit", impl_argv=("bookmarks",)),
        },
        target_help="audit {t} (read-only)",
    ),
    Group(
        "freeze",
        "freeze auto-computed Word elements (headings numbering / fields)",
        "freeze_target",
        {"headings": Target("freeze", impl_argv=("headings",),
                            rest_help="extra args (e.g. --types TOC,PAGEREF)"),
         "fields": Target("freeze", impl_argv=("fields",),
                          rest_help="extra args (e.g. --types TOC,PAGEREF)")},
        rest_help="extra args (e.g. --types TOC,PAGEREF)",
    ),
    Group(
        "strip",
        "strip stale/polluting docx elements (outlinelvl / bookmarks / revisions / "
        "doc-protection / orphan-media / empty-captions)",
        "strip_target",
        {
            "outlinelvl":       Target("strip", impl_argv=("outlinelvl",)),
            "style-outlinelvl": Target("strip", impl_argv=("style-outlinelvl",)),
            "bookmarks":        Target("strip", impl_argv=("bookmarks",)),
            "revisions":        Target("strip", impl_argv=("revisions",)),
            "doc-protection":   Target("strip", impl_argv=("doc-protection",)),
            # 这两个额外认 -o/--output：写到新路径，原件不动
            "orphan-media":     Target("strip", impl_argv=("orphan-media",),
                                       opts=STD + (OUT_OPT,)),
            "empty-captions":   Target("strip", impl_argv=("empty-captions",),
                                       opts=STD + (OUT_OPT,)),
        },
    ),
    Group(
        "header-footer",
        "add/manage docx header & footer (院标准格式)",
        "hf_target",
        {"add": "add_header_footer"},
        target_help="header-footer {t}",
        metavar="<action>",
        rest_help="forwarded args (--header / --footer-prefix / --page-number / "
                  "--font-size / --gap-spaces / --header-align / --footer-align)",
    ),
    Group(
        "legacy",
        "DEPRECATED catch-all scripts (fix-heading-disorder)",
        "legacy_target",
        {"fix-heading-disorder": Target(
            "blocks", impl_argv=("fix-heading-disorder",), rest_help="extra args forwarded",
            deprecated="[sub.legacy] WARNING: 'fix-heading-disorder' is DEPRECATED; "
                       "拆分后的单功能替代见 sub/_groups.py 本条注释："
                       "outline normalize-arabic / promote-h1 / demote-h2 / "
                       "renumber headings / blocks reorder。")},
        target_help="[DEPRECATED] {t}",
    ),
    Group(
        "blocks",
        "paragraph-block structural ops (reorder/relocate)",
        "block_target",
        {
            "reorder": Target("blocks", impl_argv=("reorder",),
                              rest_help="extra args forwarded"),
            "relocate": Target("blocks", impl_argv=("relocate",), opts=STD + (
                Opt("--plan", help="plan JSON path (schemas/plan.schema.json v1)"),)),
        },
        shared=True,
        rest_help="extra args forwarded",
    ),

    # ── 有自己选项的组 ─────────────────────────────────────────────────────
    Group(
        "image",
        "image ops (relink media from source docx / extract images by caption)",
        "image_action",
        {
            "relink": Target(
                "image", impl_argv=("relink",),
                help="relink/re-embed images from a source docx",
                opts=(
                    Opt("target_docx", nargs="?", help="target docx (with dangling drawings)"),
                    Opt("--source", help="source docx to lift media from"),
                    Opt("--apply-patch", help="apply existing patch JSON (skip detect)"),
                    Opt("--report", help="write patch JSON to this path "
                                         "(use with --dry-run for detect-only)"),
                    Opt("--dry-run", action="store_true"),
                    Opt("--no-backup", action="store_true"),
                ),
                rest_help="extra args forwarded"),
            "extract": Target(
                "image", impl_argv=("extract",),
                help="extract images by neighboring caption text",
                opts=(
                    Opt("docx", nargs="?", help="source docx (read-only)"),
                    Opt("--out-dir", help="output directory"),
                    Opt("--quiet", action="store_true"),
                ),
                rest_help="extra args forwarded"),
        },
        metavar="<action>",
    ),
    Group(
        "seqdiff",
        "docx seqdiff ops: seq (逐段对照) / image (图片去重) — distilled from bid-diff-and-revise",
        "diff_action",
        {
            "seq": Target("biddiff", impl_argv=("seq",),
                          help="逐段 sequence diff: src vs dst → MD 报告", opts=(
                Opt("--src", help="源 docx（旧版）"),
                Opt("--dst", help="目标 docx（新版）"),
                Opt("--out", help="输出 MD 路径"),
                Opt("--noise-len", type_int=True, default=8, dest="noise_len",
                    help="短于 N 字符视为噪声（默认 8）", emit_zero=True),
            ), rest_help=""),
            "image": Target("image", impl_argv=("dedup",),
                            help="SHA256 图片去重: src vs dst → 重复清单 MD", opts=(
                Opt("--src", help="源 docx"),
                Opt("--dst", help="目标 docx"),
                Opt("--out", help="输出 MD 路径"),
            ), rest_help=""),
        },
        metavar="<action>",
    ),
    Group(
        "compare-ref",
        "compare-ref: 改动 MD 改为段 vs 参考 docx 雷同检查 — distilled from bid-diff-and-revise",
        "compare_action",
        {"ref": Target("biddiff", impl_argv=("ref",),
                       help="改动草稿 MD 改为段 vs 参考 docx 雷同检查", opts=(
            Opt("--drafts-dir", dest="drafts_dir", help="改动草稿 MD 所在目录"),
            Opt("--ref", help="参考 docx（主标或基准）"),
            Opt("--out", help="输出 MD 路径"),
            Opt("--glob", default="*改动草稿.md", help="MD 文件 glob 模式"),
        ), rest_help="")},
        metavar="<action>",
    ),
    Group(
        "revise-rules",
        "revision rules ops: gen (改动 MD → rules JSON for track-changes)",
        "revise_rules_action",
        {"gen": Target("biddiff", impl_argv=("gen",),
                       help="解析改动草稿 MD → rules JSON", opts=(
            Opt("--drafts-dir", dest="drafts_dir", help="改动草稿 MD 目录"),
            Opt("--docx", help="目标 docx（find 存在性校验 + 引号自适配）"),
            Opt("--out", help="输出 rules JSON 路径"),
            Opt("--glob", default="*-改动草稿.md", help="MD glob（默认 *-改动草稿.md）"),
            Opt("--no-title-skip", action="store_true", dest="no_title_skip",
                help="不跳过 title-only rule"),
            Opt("--no-paren-strip", action="store_true", dest="no_paren_strip",
                help="不剥除说明性括号"),
        ), rest_help="")},
        metavar="<action>",
    ),
    Group(
        "section",
        "section read/list ops on a DOCX",
        "section_action",
        {"read": Target(
            "section_read", help="read a named section (or list all headings)",
            add_help=True,
            opts=(
                Opt("docx_file", help="DOCX file path"),
                # 转发顺序是 docx → --list → query，而 --list 命中时 query 不发。
                # 声明顺序里位置参数的相对次序（docx_file 在 query 前）不受影响。
                Opt("--list", dest="list_headings", action="store_true",
                    help="list all headings with paragraph indices"),
                Opt("query", nargs="?", suppressed_by="list_headings",
                    help="heading keyword to match (fuzzy); omit with --list to enumerate"),
            ),
            rest_help="")},
        metavar="action",
        required=False,
    ),
    Group(
        "split",
        "split / body-replace docx (by-h1, body-replace)",
        "split_action",
        {
            "by-h1": Target("split", impl_argv=("by-h1",),
                            help="split docx by Heading 1 into N files",
                            add_help=True, rest_help="", opts=(
                Opt("--docx", required=True, help="input docx path"),
                Opt("--out-dir", dest="out_dir", required=True, help="output directory"),
                Opt("--name-pattern", dest="name_pattern",
                    help="filename pattern (default '{idx:02d}-{title}.docx')"),
                Opt("--include-frontmatter", action="store_true", dest="include_frontmatter",
                    help="emit content before first H1 as 00-frontmatter.docx"),
                Opt("--dry-run", action="store_true", help="print plan only, don't write files"),
                Opt("--allow-no-h1", action="store_true", dest="allow_no_h1",
                    help="suppress unhealthy-docx fail-fast (rare; default = "
                         "FAIL on 0 H1 and tell user to run /docx health first)"),
            )),
            "body-replace": Target(
                "split", impl_argv=("body-replace",),
                help="keep shell styles/cover/H1, replace body with content",
                add_help=True, rest_help="", opts=(
                    Opt("--shell", required=True,
                        help="shell docx (styles/numbering/cover/H1 source)"),
                    Opt("--content", required=True, help="content docx (body source)"),
                    Opt("--out", required=True, help="output docx path"),
                    # 这一对是互斥组：默认真值那支不转发，只有被显式关掉时才发反向 flag。
                    Opt("--keep-shell-h1", dest="keep_shell_h1", action="store_true",
                        default=True, mx="h1", fwd="never",
                        help="(default) keep shell's first H1, drop content's first H1"),
                    Opt("--no-keep-shell-h1", dest="keep_shell_h1", action="store_false",
                        mx="h1", fwd="when_false",
                        help="drop all shell body, take content from its first element"),
                    Opt("--dry-run", action="store_true", help="print plan only, don't write output"),
                )),
        },
        metavar="action",
        required=False,
    ),
    Group(
        "fonts",
        "字体/颜色规整 (normalize)",
        "fonts_target",
        {"normalize": Target(
            "normalize_fonts", help="统一中西文字体 + 标题颜色压黑（杀 pandoc 主题蓝）",
            add_help=True, rest_help=None,
            description=(
                "统一 docx 中西文字体并把标题颜色压成黑色。\n\n"
                "根因：pandoc 套 Word 内置 Heading 样式自带主题蓝（H1 无色继承蓝、\n"
                "H3/4/5=0F4761 青蓝），生成脚本只改字号没改色 → 标题蓝。本命令：\n"
                "  · 标题段（样式名含 heading/标题/title/subtitle/toc）：中文→黑体、\n"
                "    西文→Times New Roman、颜色→黑(000000)。\n"
                "  · 其余段（正文/图注/表格内）：中文→宋体、西文→Times New Roman；颜色不动。\n"
                "样式级 + run 级双保险（防 pandoc run 直接 rFonts 盖样式）。\n"
                "原地改 + 默认备份 .bak-时间戳。"),
            opts=(
                Opt("docx_path_pos", nargs="?", fwd="never",
                    help="(positional) docx 路径，等价 --docx"),
                Opt("--docx", dest="docx_kw", alias_of="docx_path_pos", forward="--docx",
                    must_have="[fonts normalize] missing docx (positional or --docx)",
                    help="目标 docx（原地修改）"),
                Opt("--body-cjk", default="宋体", help="正文中文字体，默认 宋体"),
                Opt("--heading-cjk", default="黑体", help="标题中文字体，默认 黑体"),
                Opt("--latin", default="Times New Roman", help="西文字体，默认 Times New Roman"),
                Opt("--heading-color", default="000000", help="标题颜色 hex，默认 000000"),
                Opt("--no-backup", action="store_true", help="不创建 .bak-时间戳"),
                Opt("--dry-run", action="store_true", help="只报告不写盘"),
            ))},
        shared=True,
    ),
    Group(
        "table",
        "表格操作 (delete-rows / extract / ...)",
        "table_target",
        {
            # 2026-07-30 修：这条从写下那天起没有 func=，敲下去只得到「no handler for table」。
            "delete-rows": Target(
                "table", impl_argv=("delete-rows",),
                help="删除指定表格的行范围（安全校验 + 原地保存）",
                add_help=True, rest_help=None,
                description=(
                    "删除 docx 中指定表（0-based 索引）的行范围 FROM:TO（闭区间 0-based）。\n"
                    "支持三段安全校验：保留行首列期望值、被删行关键字、删后末行期望值。\n"
                    "原地修改 docx，保存后重读验证。\n\n"
                    "用途：track-changes 不支持表行删除时，由脚本直接删除。"),
                opts=(
                    Opt("docx", forward="--docx", help="目标 docx（原地修改）"),
                    Opt("--table-index", dest="table_index", type_int=True, required=True,
                        help="表格索引（0-based）", emit_zero=True),
                    Opt("--rows", required=True, help="行范围 FROM:TO（闭区间 0-based）"),
                    Opt("--expected-first-col", default="",
                        help="删除前保留行第一列期望值（逗号分隔），安全校验；留空跳过"),
                    Opt("--expected-residue", default="",
                        help="被删起始行应含此关键字，确保删对行"),
                    Opt("--expected-last", default="",
                        help="删后末行第二列期望值，校验删后结果"),
                )),
            "extract": Target(
                "table", impl_argv=("extract",),
                help="抽 docx 内每张表为独立 docx（文件名=邻近 caption）",
                add_help=True, rest_help=None,
                description=(
                    "遍历 docx body 内每张表 <w:tbl>，往前 1-2 段（或往后 1 段）扫"
                    "「表…」开头的 caption 段，取文字作文件名 stem。\n"
                    "无 caption → fallback `table-{idx:02d}`；重名加 -2/-3。\n\n"
                    "输出策略：shutil.copy 源 docx → 删 body 内非目标元素（保留 sectPr）→ save，\n"
                    "保留完整 styles / numbering / media / rels（与 split_by_h1 同套路）。"),
                opts=(
                    Opt("docx_path_pos", nargs="?", fwd="never",
                        help="(positional) docx 路径，等价 --docx"),
                    Opt("--docx", dest="docx_kw", alias_of="docx_path_pos", forward="--docx",
                        must_have="[table extract] missing docx (positional or --docx)",
                        help="目标 docx 路径"),
                    Opt("--out-dir", dest="out_dir", required=True, help="输出目录（mkdir -p）"),
                    Opt("--name-pattern", dest="name_pattern", default="{stem}.docx",
                        help="文件名 pattern，默认 '{stem}.docx'（可用 {stem} {idx}）"),
                    Opt("--dry-run", action="store_true", help="只打印 plan，不写文件"),
                )),
            "borders": Target(
                "table", impl_argv=("borders",),
                help="把所有表格统一为满格实线（表级全 single + 清单元格 nil 覆盖）",
                add_help=True, rest_help=None, chain=("center", "center"),
                description=(
                    "把 docx 内所有表格（含嵌套表）统一为「满格实线」边框。\n\n"
                    "根因：表级 <w:tblBorders> 与单元格级 <w:tcBorders> 两级，后者优先级更高。\n"
                    "某表「看着不是全实线」常因表级没设边框、且部分单元格 tcBorders 把内部边\n"
                    "设成 val='nil'（无线）→ 内部竖/横线缺失。光设表级盖不住单元格 nil。\n\n"
                    "本命令两手抓：① 每表设表级 tblBorders 6 边全 single；\n"
                    "② 默认删每个单元格 tcBorders（让表级统一生效），或 --keep-cell-borders\n"
                    "时只把非 single 边改写为实线。原地改 + 默认备份 .bak-时间戳。"),
                opts=(
                    Opt("docx_path_pos", nargs="?", fwd="never",
                        help="(positional) docx 路径，等价 --docx"),
                    Opt("--docx", dest="docx_kw", alias_of="docx_path_pos", forward="--docx",
                        must_have="[table borders] missing docx (positional or --docx)",
                        help="目标 docx（原地修改）"),
                    Opt("--val", default="single", help="线型，默认 single 实线"),
                    Opt("--sz", type_int=True, default=4, help="线宽 1/8pt（4=0.5pt），默认 4",
                        emit_zero=True),
                    Opt("--color", default="auto", help="颜色，默认 auto（黑）"),
                    Opt("--space", type_int=True, default=0, help="边距，默认 0", emit_zero=True),
                    Opt("--keep-cell-borders", action="store_true", dest="keep_cell_borders",
                        help="保留 tcBorders，仅改写非实线边（默认=删 tcBorders）"),
                    # 这两个不转发给 set_table_borders：--center 触发 chain 去跑 center，
                    # --cell-center 是给 center 那一趟用的。
                    Opt("--center", action="store_true", fwd="never",
                        help="顺带把所有表格整体水平居中（= 再跑一次 table center）"),
                    Opt("--cell-center", action="store_true", dest="cell_center", fwd="never",
                        help="配合 --center：单元格内文字也水平+垂直居中"),
                    Opt("--no-backup", action="store_true", help="不创建 .bak-时间戳"),
                    Opt("--dry-run", action="store_true", help="只报告不写盘"),
                )),
            "center": Target(
                "table", impl_argv=("center",),
                help="把所有表格整体在页面水平居中（含嵌套表）",
                add_help=True, rest_help=None,
                description=(
                    "把 docx 内所有表格（含嵌套表）整体在页面水平居中——写表级\n"
                    "<w:tblPr>/<w:jc w:val='center'/>（表作为块左右居中），与单元格内文字\n"
                    "对齐是两回事。--cell-center 时同时把单元格内文字水平+垂直居中。\n"
                    "原地改 + 默认备份 .bak-时间戳。"),
                opts=(
                    Opt("docx_path_pos", nargs="?", fwd="never",
                        help="(positional) docx 路径，等价 --docx"),
                    Opt("--docx", dest="docx_kw", alias_of="docx_path_pos", forward="--docx",
                        must_have="[table center] missing docx (positional or --docx)",
                        help="目标 docx（原地修改）"),
                    Opt("--cell-center", action="store_true", dest="cell_center",
                        help="同时把单元格内文字水平+垂直居中（默认只表格整体居中）"),
                    Opt("--no-backup", action="store_true", help="不创建 .bak-时间戳"),
                    Opt("--dry-run", action="store_true", help="只报告不写盘"),
                )),
        },
        shared=True,
    ),

    # ── 扁平组：没有子命令，选项直接挂在组上 ─────────────────────────────
    Group(
        "chrome",
        "院报告版面装帧: 逐章分节+逐章页眉页脚水印+宽表横向节 (→ <raw>_chrome.docx)",
        "",
        flat=Target("docx_chrome", add_help=True, rest_help="", opts=(
            # --validate 命中时短路：只发 --validate 那两个值，其余一概不发。
            Opt("--validate", nargs=2, metavar=("OUT", "TEMPLATE"), fwd="short_circuit",
                help="diff 节结构对照范式而非构建"),
            Opt("--raw", help="输入 raw 合并 docx(正文已成型)"),
            Opt("--template", help="范式 docx(取页眉/页脚/水印部件)"),
            Opt("--out", help="输出(默认 <raw stem>_chrome.docx)"),
            Opt("--county", default="天台县", help="目标县名(替换范式县名)"),
        )),
    ),
    Group(
        "md-merge",
        "replace a DOCX section with MD content (table-safe; --in-place + .bak; "
        "--start/end-anchor)",
        "",
        flat=Target("md_merge_impl", add_help=True, rest_help="extra args forwarded "
                                                              "to underlying script", opts=(
            Opt("md_file", help="source Markdown file"),
            Opt("docx_file", help="target DOCX file"),
            Opt("start_idx", nargs="?", type_int=True, emit_zero=True,
                help="start paragraph index (heading; preserved + title updated). "
                     "可用 --start-anchor 替代"),
            Opt("end_idx", nargs="?", type_int=True, emit_zero=True,
                help="end paragraph index (exclusive; next section heading). "
                     "可用 --end-anchor 替代"),
            Opt("output_file", nargs="?",
                help="output path (default: <docx>-merged.docx; --in-place 时忽略)"),
            Opt("--start-anchor", help="按标题文本定位 start_idx (省掉手查索引)"),
            Opt("--end-anchor", help="按标题文本定位 end_idx (下一节标题)"),
            Opt("--in-place", action="store_true", help="原地改 + 自动 .bak-时间戳 (Work §1.5)"),
            Opt("--no-backup", action="store_true", help="配合 --in-place 跳过备份"),
        )),
    ),

    # ── shared 组：同名组里还有别的模块注册的 target ───────────────────────
    #   caption pair ← 本表 · caption number-by-style ← styles.py
    #   chapter delete ← 本表 · renumber h4-figures ← styles.py
    Group(
        "chapter",
        "chapter ops (convert-arabic / delete-empty-h1 / delete)",
        "chapter_target",
        {
            "convert-arabic": Target("chapter", impl_argv=("convert-arabic",)),
            "delete-empty-h1": Target("chapter", impl_argv=("delete-empty-h1",)),
            # 注意：delete 没有 --report（原样保持）
            "delete": Target("chapter", impl_argv=("delete",),
                             help="delete an entire H1 chapter",
                             rest_help="extra args forwarded", opts=(
                Opt("docx_path", nargs="?", help="target docx path"),
                Opt("--h1", help="H1 chapter number (e.g. 3)"),
                Opt("--prefix", help="H1 text prefix to match"),
                Opt("--dry-run", action="store_true"),
                Opt("--no-backup", action="store_true"),
            )),
        },
        shared=True,
    ),
    Group(
        "renumber",
        "renumber headings + caption numbers",
        "renumber_target",
        {"headings": Target("renumber", impl_argv=("headings",))},
        shared=True,
    ),
    Group(
        "caption",
        "caption ops (number / pair / number-by-style)",
        "caption_target",
        {
            "number": Target("caption", impl_argv=("number",)),
            "pair": Target("caption", impl_argv=("pair",),
                           help="apply decision.json to fix caption-table pairing",
                           rest_help="extra args forwarded", opts=STD + (
                Opt("--decision", help="decision JSON path (schemas/decision.schema.json v1)"),)),
        },
        shared=True,
    ),
]

BY_NAME: dict[str, Group] = {g.name: g for g in GROUPS}

# 表自检：表是唯一事实源，它自己错了下面全错，所以 import 期就炸。
for _g in GROUPS:
    if bool(_g.dest) == bool(_g.flat):
        raise SystemExit(f"[GROUPS 表错] {_g.name}: dest 与 flat 必须二选一")
    if not _g.dest and _g.targets:
        raise SystemExit(f"[GROUPS 表错] {_g.name}: 扁平组不该有 targets")
    if _g.dest and not _g.targets:
        raise SystemExit(f"[GROUPS 表错] {_g.name}: targets 为空 —— 空组注册出来是个"
                         f"点不动的子命令，比没有还坏")
    for _n, _t in _targets(_g).items():
        if _t.chain and _t.chain[1] not in _g.targets:
            raise SystemExit(f"[GROUPS 表错] {_g.name} {_n}: chain 指向不存在的 "
                             f"target {_t.chain[1]!r}")
        _d = [_dest_of(o) for o in _t.opts]
        if len(_d) != len(set(_d)) and not any(o.mx for o in _t.opts):
            raise SystemExit(f"[GROUPS 表错] {_g.name} {_n}: dest 重名 {_d}")


# ──────────────────────────────── 构建器 ────────────────────────────────

def _add_opt(parser, o: Opt) -> None:
    kw: dict[str, Any] = {}
    if o.help:
        kw["help"] = o.help
    if o.action != "store":
        kw["action"] = o.action
    if o.nargs is not None:
        kw["nargs"] = o.nargs
    if o.type_int:
        kw["type"] = int
    if o.metavar:
        kw["metavar"] = o.metavar
    if o.default is not None:
        kw["default"] = o.default
    if _is_pos(o):
        parser.add_argument(o.name, **kw)
        return
    if o.required:
        kw["required"] = True
    kw["dest"] = _dest_of(o)
    names = [o.name] + ([o.alt] if o.alt else [])
    parser.add_argument(*names, **kw)


def _add_target_parser(sp, name: str, t: Target):
    kw: dict[str, Any] = {"help": t.help, "add_help": t.add_help}
    if t.description:
        kw["description"] = t.description
    p = sp.add_parser(name, **kw)
    _fill(p, t)
    return p


def _fill(p, t: Target) -> None:
    mx_groups: dict[str, Any] = {}
    for o in t.opts:
        if o.mx:
            if o.mx not in mx_groups:
                mx_groups[o.mx] = p.add_mutually_exclusive_group()
            _add_opt(mx_groups[o.mx], o)
        else:
            _add_opt(p, o)
    if t.rest_help is not None:
        p.add_argument("rest", nargs=argparse.REMAINDER,
                       **({"help": t.rest_help} if t.rest_help else {}))


def _argv_for(t: Target, args) -> list[str] | None:
    """按 Target.opts 把 Namespace 翻译成实现脚本的 argv。返回 None = 缺必需值，已报错。"""
    # 短路项先看：命中就只发它（chrome --validate 是「验证模式」，不是普通选项）
    for o in t.opts:
        if o.fwd == "short_circuit":
            v = getattr(args, _dest_of(o), None)
            if v:
                return [o.forward or o.name] + [str(x) for x in v]

    argv: list[str] = []
    for o in t.opts:
        if o.fwd in ("never", "short_circuit"):
            continue
        if o.suppressed_by and getattr(args, o.suppressed_by, False):
            continue
        v = getattr(args, _dest_of(o), None)
        if o.fwd == "when_false":
            if v is False:
                argv.append(o.forward or o.name)
            continue
        if o.action in ("store_true", "store_false"):
            if v:
                argv.append(o.forward or o.name)
            continue
        if v is None and o.alias_of:
            v = getattr(args, o.alias_of, None)
        ok = (v is not None) if o.emit_zero else bool(v)
        if not ok:
            if o.must_have:
                print(o.must_have, flush=True)
                return None
            continue
        flag = o.forward or (o.name if not _is_pos(o) else "")
        if flag:
            argv.extend([flag, str(v)])
        else:
            argv.append(str(v))
    if t.rest_help is not None:
        argv.extend(str(x) for x in (getattr(args, "rest", None) or []))
    return argv


def _run(args) -> int:
    """所有组共用的转发入口。组名/子命令名由 register 时 set_defaults 钉死，不靠猜。"""
    g = BY_NAME.get(getattr(args, "_group_name", ""))
    if g is None:
        print(f"[sub._groups] 未知组: {getattr(args, '_group_name', None)!r}")
        return 2
    tname = getattr(args, "_target_name", "")
    t = g.flat if g.flat is not None else _targets(g).get(tname)
    if t is None:
        print(f"[sub.{g.name}] unknown target: {tname}; choices={list(g.targets)}")
        return 2
    if t.deprecated:
        print(t.deprecated, flush=True)
    argv = _argv_for(t, args)
    if argv is None:
        return 2
    rc = exec_script(t.impl, list(t.impl_argv) + argv)
    if rc == 0 and t.chain and getattr(args, t.chain[0], False):
        nxt = _targets(g)[t.chain[1]]
        nargv = _argv_for(nxt, args)
        if nargv is None:
            return 2
        rc = exec_script(nxt.impl, list(nxt.impl_argv) + nargv)
    return rc


def _register_one(subparsers, g: Group) -> None:
    if g.flat is not None:
        p = subparsers.add_parser(g.name, help=g.help)
        _fill(p, g.flat)
        p.set_defaults(func=_run, _group_name=g.name, _target_name="")
        return
    if g.shared:
        p = get_or_add_group(subparsers, g.name, g.help)
        sp = get_or_add_subparsers(p, dest=g.dest, metavar=g.metavar, required=g.required)
        existing = getattr(sp, "choices", {}) or {}
    else:
        p = subparsers.add_parser(g.name, help=g.help)
        sp = p.add_subparsers(dest=g.dest, metavar=g.metavar, required=g.required)
        existing = {}
    for name, t in _targets(g).items():
        if name in existing:
            continue            # 已被同组的另一个模块注册（shared 组的常态）
        tp = _add_target_parser(sp, name, t)
        tp.set_defaults(func=_run, _group_name=g.name, _target_name=name)


def register_group(subparsers, name: str) -> None:
    """注册**单个**组。为什么按名而不是一次性全注册：顶层 `--help` 里组的先后 =
    注册先后，而本表接管的组在原来的注册序列里是**散着**的（中间夹着 outline /
    styles / health 这些没被接管的）。一次性注册会把它们挤成一坨，用户敲
    `--help` 看到的顺序就变了 —— 那是没人要求过的改动。"""
    g = BY_NAME.get(name)
    if g is None:
        raise KeyError(f"GROUPS 表里没有组 {name!r}；有的是 {list(BY_NAME)}")
    _register_one(subparsers, g)


def register_all(subparsers) -> None:
    """把 GROUPS 全部注册上去（本仓不走这条，留给独立使用者）。"""
    for g in GROUPS:
        _register_one(subparsers, g)
