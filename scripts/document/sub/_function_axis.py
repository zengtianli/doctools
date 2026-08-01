#!/usr/bin/env python3
"""_function_axis.py — 子命令的**职能轴** SSOT（2026-08-01 立）。

目录轴（`sub/` 里文件怎么摆）回答的是「实现在哪」；职能轴回答的是
「这条命令碰的是**什么**」。两条轴正交，所以职能不落成目录、落成这张表 +
一道机检（`tools/check_function_axis.py`）。

    一条子命令一行数据。**加子命令 = 这里也加一行**，否则闸门转红。

职能枚举（严格按 `FN_DEFINITIONS`，别自创第七个值）：

    format   只碰版式格式：样式/字体/段落属性/页眉页脚/页面分节/编号与题注号/
             大纲层级/表格边框对齐/装帧/样式集
    content  碰正文写了什么：删段/改文字/合并拆分文档/回写章节/挪段落块/资源重挂
    review   审阅修订：w:ins/w:del/批注/track changes/版本对照
    inspect  只读检查审计，不写盘
    convert  跨格式转换
    dispatch 调度编排

键是 `(顶层名, 动作)`，扁平命令的动作是 `None`（`template` → `("template", None)`）。
**别名不进表**：`read`=extract · `diff`=compare · `styleset *`=audit-styleset *，
它们与被别名者是同一个 parser 对象，闸门按 `id(parser)` 归组后自动认领同一条标签。
在这里给别名再写一行 = 同一职能散在两处，闸门会判红（`别名重复打标签`）。

怎么查：

    python3 tools/check_function_axis.py            # 表与 CLI 对账（fail-closed）
    python3 tools/check_function_axis.py --fn format --md
    python3 scripts/document/docx_cli.py verbs --fn format
"""
from __future__ import annotations

from typing import Optional

# ─── 职能枚举（取值域 SSOT） ────────────────────────────────────────────
FN_DEFINITIONS: dict[str, str] = {
    "format": "只碰版式格式：样式/字体/段落属性/页眉页脚/页面分节/编号与题注号/"
              "大纲层级/表格边框对齐/装帧/样式集",
    "content": "碰正文写了什么：删段/改文字/合并拆分文档/回写章节/挪段落块/资源重挂",
    "review": "审阅修订：w:ins/w:del/批注/track changes/版本对照",
    "inspect": "只读检查审计，不写盘",
    "convert": "跨格式转换",
    "dispatch": "调度编排",
}
FN_TAGS: tuple[str, ...] = tuple(FN_DEFINITIONS)

# ─── 表本体 ────────────────────────────────────────────────────────────
# (顶层名, 动作 | None, 职能, 备注)
# 用 tuple-of-tuples 而不是 dict 字面量：dict 字面量里写重复键不会报错、
# 后一条静默覆盖前一条 —— 那正是这张表最该被拦住的一类手误。
_ROWS: tuple[tuple[str, Optional[str], str, str], ...] = (
    # ── format ──────────────────────────────────────────────────────
    ("style", "body", "format", ""),
    ("style", "table", "format", ""),
    ("style", "caption", "format", ""),
    ("fix", "clear-direct-format", "format", ""),
    ("fix", "role-fill", "format", ""),
    ("fix", "style-create", "format", ""),
    ("fix", "style-pane-filter", "format", ""),
    ("fix", "style-pool-cleanup", "format", ""),
    ("fix", "style-rebrand", "format", ""),
    ("fix", "style-rename", "format", ""),
    ("fonts", "normalize", "format", ""),
    ("template", None, "format", ""),
    ("format", None, "format", "版式复刻(docx_fmt clone)"),
    ("chrome", None, "format", "分节+页眉页脚+水印装帧"),
    ("header-footer", "add", "format", ""),
    ("text-fmt", None, "format", "灰区:改的是正文字符(引号/标点/单位),但目的是排版规范"),
    ("slim", None, "format", "灰区:瘦身删的是冗余样式/媒体,不是正文语义"),
    ("renumber", "headings", "format", ""),
    ("renumber", "h4-figures", "format", ""),
    ("renumber-fig", None, "format", ""),
    ("caption", "number", "format", ""),
    ("caption", "number-by-style", "format", ""),
    ("caption", "pair", "format", "题注与图表配对,判的是版式关系"),
    ("outline", "promote-h1", "format", ""),
    ("outline", "demote-h2", "format", ""),
    ("outline", "normalize-arabic", "format", ""),
    ("freeze", "headings", "format", "把自动编号冻成字面值"),
    ("freeze", "fields", "format", ""),
    ("chapter", "convert-arabic", "format", ""),
    ("image-caption", None, "format", "图片+图名样式"),
    ("fix-ref", None, "format", "上标引用的字符样式"),
    ("legacy", "fix-heading-disorder", "format", "DEPRECATED"),
    ("strip", "outlinelvl", "format", ""),
    ("strip", "style-outlinelvl", "format", ""),
    ("strip", "doc-protection", "format", "文档保护属性,不碰正文"),
    ("audit-styleset", "restore", "format", "族里唯一会写盘的:恢复样式集"),
    ("health", "fix", "format", "体检后修的是样式/版式问题"),
    ("health", "full", "format", "diagnose+fix 一条龙"),
    ("para", "fix-ppr", "format", "段落属性 = 版式"),
    ("table", "borders", "format", ""),
    ("table", "center", "format", ""),
    # ── content ─────────────────────────────────────────────────────
    ("strip", "bookmarks", "content", "书签是正文里的锚,删了引用会断"),
    ("strip", "orphan-media", "content", "删没人引用的媒体资源"),
    ("strip", "empty-captions", "content", "删空题注段落 = 删段"),
    ("table", "delete-rows", "content", ""),
    ("image", "relink", "content", "从源 docx 提媒体重嵌"),
    ("blocks", "reorder", "content", ""),
    ("blocks", "relocate", "content", ""),
    ("chapter", "delete", "content", ""),
    ("chapter", "delete-empty-h1", "content", ""),
    ("split", "by-h1", "content", ""),
    ("split", "body-replace", "content", ""),
    ("combine", None, "content", ""),
    ("chapters-sync", None, "content", "成品 docx 反向回写章节目录"),
    ("md-merge", None, "content", ""),
    ("para", "edit", "content", ""),
    # ── review ──────────────────────────────────────────────────────
    ("strip", "revisions", "review", "接受/清修订标记"),
    ("track", None, "review", ""),
    ("md-merge-track", None, "review", ""),
    ("revise-rules", "gen", "review", ""),
    ("seqdiff", "seq", "review", "逐段对照"),
    ("compare", None, "review", "两版对照"),
    # ── inspect ─────────────────────────────────────────────────────
    ("table", "extract", "inspect", ""),
    ("image", "extract", "inspect", "抽图出来,不改原件"),
    ("seqdiff", "image", "inspect", "图片去重,只读判定"),
    ("compare-ref", "ref", "inspect", "与参考件雷同检查"),
    ("audit", "bookmarks", "inspect", ""),
    ("audit", "captions", "inspect", ""),
    ("audit", "fields", "inspect", ""),
    ("audit", "headings", "inspect", ""),
    ("audit", "images", "inspect", ""),
    ("audit", "table-pairing", "inspect", ""),
    ("audit-styleset", "body-style-concentration", "inspect", ""),
    ("audit-styleset", "role-coverage", "inspect", ""),
    ("audit-styleset", "style-coherence", "inspect", ""),
    ("audit-styleset", "style-pane-filter", "inspect", ""),
    ("audit-styleset", "style-pool-cleanliness", "inspect", ""),
    ("check", None, "inspect", ""),
    ("snapshot", None, "inspect", ""),
    ("extract", None, "inspect", ""),
    ("section", "read", "inspect", ""),
    ("scan-sensitive", None, "inspect", ""),
    ("health", "diagnose", "inspect", ""),
    ("health", "gate", "inspect", ""),
    ("health-split", None, "inspect", ""),
    ("para", "locate", "inspect", ""),
    ("para", "inspect", "inspect", ""),
    ("para", "scan-ppr", "inspect", ""),
    ("para", "render", "inspect", ""),
    ("verbs", None, "inspect", "列出子命令职能(本表自己的读取入口)"),
    # ── convert ─────────────────────────────────────────────────────
    ("md-to-docx", None, "convert", ""),
    ("md", None, "convert", "md 工具子组"),
    # ── dispatch ────────────────────────────────────────────────────
    ("pipeline", "run", "dispatch", ""),
)


def _build() -> "dict[tuple[str, Optional[str]], tuple[str, str]]":
    """import 期自检：职能值必须在取值域内、键不许重复。

    fail-closed —— 表本身出错就让 import 炸，别让下游闸门拿着半张表报绿。
    """
    axis: dict[tuple[str, Optional[str]], tuple[str, str]] = {}
    for top, action, fn, note in _ROWS:
        if fn not in FN_DEFINITIONS:
            raise ValueError(
                f"_function_axis: 未知职能 {fn!r}（{top} {action or ''}）；"
                f"取值域只有 {', '.join(FN_TAGS)}"
            )
        key = (top, action)
        if key in axis:
            raise ValueError(f"_function_axis: 重复条目 {top} {action or ''}")
        axis[key] = (fn, note)
    return axis


AXIS: "dict[tuple[str, Optional[str]], tuple[str, str]]" = _build()


def path_of(top: str, action: Optional[str]) -> str:
    """`("audit","headings")` → `"audit headings"`；扁平命令只有顶层名。"""
    return f"{top} {action}".strip() if action else top


def lookup(top: str, action: Optional[str] = None) -> Optional[tuple[str, str]]:
    """查 `(职能, 备注)`；表里没有返回 None（调用方自己判 fail-closed）。"""
    return AXIS.get((top, action))


def rows(fn: Optional[str] = None) -> list[tuple[str, Optional[str], str, str]]:
    """按声明顺序列出 `(顶层名, 动作, 职能, 备注)`；`fn` 非空则只列该职能。"""
    if fn is not None and fn not in FN_DEFINITIONS:
        raise ValueError(f"未知职能 {fn!r}；取值域 {', '.join(FN_TAGS)}")
    return [r for r in _ROWS if fn is None or r[2] == fn]


def counts() -> dict[str, int]:
    """每个职能几条（含取值域里 0 条的那些，别让空职能从报告里消失）。"""
    out = {tag: 0 for tag in FN_TAGS}
    for _, _, fn, _ in _ROWS:
        out[fn] += 1
    return out
