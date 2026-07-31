"""sub — docx 处理子命令模块包。

两类成员，别混：

1. **声明表**（`_groups.py`）—— 20 个原来只做 argparse 声明 + `exec_script` 转发的
   group 模块，2026-07-30 折成了一张 `GROUPS` 表。加子命令 = 加一行数据，不是加一个文件。
   被折掉的文件在 `~/.Trash/doctools-shim-fold-20260730/`（带 MANIFEST）。
2. **自带实现的模块** —— `outline` / `styles` / `health` / `slim` / `fix_styleset` /
   `docx_para` / `combine` / `md_merge_track` / `audit_styleset` / `chapters_sync` /
   `health_split` / `pipeline`。它们有真业务逻辑，各自 `register()`。

注册顺序 = 顶层 `--help` 的显示顺序，见下面 `register_all`。shared 组
（caption / chapter / renumber / table / fonts / blocks）由 `get_or_add_group`
让多个来源共用一个组父节点。

等价闸门（改这里之前先看）：
    python3 tools/cli_surface.py        # 全量子命令树指纹，128 个
    python3 tools/cli_forward_probe.py  # 每条子命令**转发出去的 argv**
两个都要跑：前者证明接口没变，后者证明参数传过去还是原样。只跑前者，
「命令能敲、跑出来的东西不对」这一类改动是看不见的。

每个实现脚本仍可独立运行：
    python3 sub/<script>.py <docx> [--dry-run] [--no-backup] [--report x.json]
"""

from . import _groups
from . import (
    audit_styleset,
    chapters_sync,
    combine,
    docx_para,
    fix_styleset,
    health,
    health_split,
    md_merge_track,
    outline,
    pipeline,
    slim,
    styles,
)

__all__ = [
    "audit",
    "audit_styleset",
    "blocks",
    "caption",
    "captions",
    "chapters_sync",
    "chrome",
    "chapter",
    "combine",
    "compare",
    "diff",
    "docx_para",
    "fix_styleset",
    "fonts",
    "freeze",
    "header_footer",
    "health",
    "health_split",
    "images",
    "legacy",
    "md_merge",
    "md_merge_track",
    "outline",
    "pipeline",
    "renumber",
    "revise_rules",
    "section",
    "slim",
    "split",
    "strip",
    "styles",
    "table",
]


def register_all(subparsers) -> None:
    """Convenience: register every group's subcommands onto a parent subparsers.

    Registration order matters for shared groups (caption / chapter / renumber):
    first-registrant defines the group parser; later modules add targets via
    `get_or_add_group` / `get_or_add_subparsers` helpers in _dispatch.py.

    Usage in docx_cli.py:
        from sub import register_all
        register_all(top_subparsers)
    """
    # 顺序 = 顶层 --help 里组的显示顺序，别随手调。
    # `_G("x")` = 由 _groups.py 的声明表注册的组（2026-07-30 起，取代 8 个只做转发的
    # group 模块）；`mod.register` = 还有自己逻辑、没被收编的模块。
    def _G(name):
        return lambda sp: _groups.register_group(sp, name)

    for reg in (
        _G("audit"), audit_styleset.register, _G("freeze"), _G("strip"),
        _G("header-footer"), outline.register, _G("blocks"), _G("image"),
        _G("legacy"),  # unique
        docx_para.register,                                                               # 段落级 查-改-验 工作台 (locate/inspect/edit/fix-ppr/scan-ppr/render, 2026-07-03)
        combine.register,                                                                 # combine N docx → 1 (docxcompose; inverse of split by-h1, 2026-06-07)
        chapters_sync.register,                                                           # 成品 docx 反向回写成品章节目录 (merge 的逆操作; govern 2026-06-08)
        _G("chrome"),                                                                  # 院报告版面装帧(逐章分节+横向节, distilled eco-flow 2026-06-04)
        _G("seqdiff"), _G("compare-ref"), _G("revise-rules"),                                             # distilled from bid-diff-and-revise
        health.register,                                                                  # health diagnose/fix/full
        pipeline.register,                                                                # pipeline driver
        _G("section"),                                                                 # section read/list (distilled from panan-rigid)
        _G("md-merge"),                                                                # md merge-into-docx (distilled from panan-rigid)
        md_merge_track.register,                                                          # md→track-changes 锚点前插 (上提 reclaim merge-tracked, 0-B 2026-05-29)
        _G("table"),                                                                   # table structural ops (delete-rows / borders / center, W4 2026-05-26)
        _G("fonts"),                                                                    # 字体/标题色规整 normalize (distilled reclaim 节水年会征文, 2026-06-21)
        _G("split"),                                                                   # split docx by-h1 (distilled from eco-flow/taizhou-天台, W1 2026-05-26)
        fix_styleset.register,                                                            # style-set fix family + shape_contract gate (W13 2026-05-26)
        health_split.register,                                                            # one-shot health + split thin wrapper (distilled from 业务模板 SOP, 2026-05-28)
        slim.register,                                                                    # docx-slim: safe ensemble + aggressive minimal skeleton (W docx-slim 2026-05-28)
        _G("chapter"), _G("renumber"), _G("caption"), styles.register,   # shared
    ):
        reg(subparsers)
