#!/usr/bin/env python3
"""script_graph.py — 解出 doctools 仓内脚本的真实引用关系，出一张可点的全景页。

    python3 tools/script_graph.py            # 生成 reports/script-graph.html
    python3 tools/script_graph.py --open
    python3 tools/script_graph.py --json     # 只出数据

页面三视图（同一份数据，同一套筛选器）：**图谱**（力导向，可拖可缩）·
**清单**（99 个脚本，谁调谁）· **动词**（93 条 CLI 动词 → 落到哪个脚本）。

## 为什么不是 /module-map

`/module-map` 从 catalog.yaml 派生「工作区有哪些域/包」，doctools 的 catalog 只登记 3 个
key_scripts —— 它管的是**工作区拓扑**。这里要的是**仓内脚本之间谁调谁**，粒度差一个量级。
两者互补，不重复。

## 边是怎么解出来的（四种来源，全部来自读源码）

1. `import X` / `from X import` —— X 能在本仓解析到文件才算边
2. `_dispatch._load("foo.py")` / `_exec_script("foo.py")` —— 组模块转发到独立脚本
3. `subprocess` 里出现的 `foo.py`
4. 任何字面量 `"foo.py"` 且 foo.py 是本仓已知脚本 —— dispatcher 的命令表就是这么写的

宁可多算一条边，也不漏：**漏边会把活脚本误判成孤儿**，而孤儿名单是要拿来退役的。

## 孤儿判据（fail-closed）

入度为 0 **且** 不是入口（顶层 CLI / 有 `if __name__ == "__main__"` 且被文档提到）。
枚举为空 → rc=2，拒绝在空集上报「没有孤儿」。
"""
from __future__ import annotations

import argparse
import ast
import html
import json
import re
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN = [ROOT / "scripts", ROOT / "lib", ROOT / "tools", ROOT / "src"]

# 顶层 CLI = 人直接敲的入口，入度 0 也不算孤儿
ENTRIES = {"docx_cli", "pdf_cli", "doc_dispatch", "typeset_pipeline", "doc_gui_backend",
           "pptx_cli", "md_tools", "script_graph", "blast_radius",
           "check_docx_collar", "check_verbs_reachable", "_inventory",
           "cli",   # doctools/cli.py = console_script 入口（装出来的 `doctools` 命令）
           # 2026-07-30 登记：spec 引擎入口 + 两个等价闸门
           "typeset_apply", "cli_surface", "cli_forward_probe"}

FAMILY = [
    ("strip",    r"^strip_|^_?strip$"),
    ("audit",    r"^audit"),
    ("heading",  r"heading|outline|chapter|renumber|promote|demote"),
    ("caption",  r"caption"),
    ("table",    r"table"),
    ("image",    r"image|media|figure|fig_"),
    ("style",    r"style|font|restyle|chrome|line_spacing"),
    ("md",       r"^md_|markdown"),
    ("bid",      r"^bid_"),
    ("compare",  r"compare|diff|review|health|qa"),
    ("freeze",   r"freeze|bookmark|field"),
    ("split",    r"split|combine|merge|port_|body_replace|blocks"),
]


def family(stem: str) -> str:
    for name, pat in FAMILY:
        if re.search(pat, stem):
            return name
    return "其它"


def collect() -> dict[str, dict]:
    nodes: dict[str, dict] = {}
    for root in SCAN:
        if not root.is_dir():
            continue
        for p in sorted(root.rglob("*.py")):
            if "__pycache__" in str(p):
                continue
            src = p.read_text(encoding="utf-8", errors="replace")
            stem = p.stem
            if stem in nodes:            # 同名不同目录：用相对路径区分，保留先扫到的为主
                stem = str(p.relative_to(ROOT)).replace("/", ":")[:-3]
            nodes[stem] = {
                "id": stem,
                "path": str(p.relative_to(ROOT)),
                "lines": src.count("\n") + 1,
                "src": src,
                "family": family(p.stem),
                "is_test": "/tests/" in str(p) or p.name.startswith("test_"),
                "doc": (ast.get_docstring(ast.parse(src)) or "").strip().split("\n")[0][:120]
                       if p.suffix == ".py" else "",
            }
    return nodes


def engine_of(src: str) -> str:
    has_docx = re.search(r"^\s*(from docx[\w.]*\s+import|import docx\b)", src, re.M)
    if re.search(r"\bdocx_surgical\b|surgical_rewrite", src):
        return "surgical"
    if has_docx and ".save(" in src:
        return "python-docx+收口" if re.search(r"^\s*import docx_safe_save\b", src, re.M) \
            else "python-docx 裸用"
    if has_docx:
        return "python-docx 只读"
    if re.search(r"^\s*from docx_xml import|^\s*import docx_xml\b", src, re.M):
        return "docx_xml"
    if re.search(r"^\s*import zipfile\b", src, re.M) and re.search(r"^\s*(from lxml|import lxml)\b", src, re.M):
        return "裸 lxml+zipfile"
    if re.search(r"\bpandoc\b", src):
        return "pandoc"
    if re.search(r"\bsoffice\b|libreoffice", src):
        return "soffice"
    return "不碰 docx 内部"


def edges(nodes: dict[str, dict]) -> list[tuple[str, str, str]]:
    """→ [(from, to, 依据)]。四种来源见模块 docstring。"""
    known = {n["path"].rsplit("/", 1)[-1]: k for k, n in nodes.items()}   # foo.py -> id
    stems = {n["path"].rsplit("/", 1)[-1][:-3]: k for k, n in nodes.items()}
    out: list[tuple[str, str, str]] = []
    seen = set()

    def add(a: str, b: str, why: str) -> None:
        if a != b and (a, b) not in seen:
            seen.add((a, b))
            out.append((a, b, why))

    for nid, n in nodes.items():
        src = n["src"]
        try:
            tree = ast.parse(src)
        except SyntaxError:
            tree = None
        if tree is not None:
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    for a in node.names:
                        t = stems.get(a.name.split(".")[-1])
                        if t:
                            add(nid, t, "import")
                elif isinstance(node, ast.ImportFrom):
                    if node.module:
                        t = stems.get(node.module.split(".")[-1])
                        if t:
                            add(nid, t, "from-import")
                    # `from . import _groups` —— 相对导入时 node.module 是 None，
                    # 名字全在 node.names 里。漏掉这一支的表现是「被 __init__ 导入的
                    # 模块被判成孤儿」，而孤儿名单是拿来退役的：错杀一个还在用的模块，
                    # 比放过一个死模块贵得多。
                    for alias in node.names:
                        t = stems.get(alias.name.split(".")[-1])
                        if t:
                            add(nid, t, "from-import")
        # 字面量调用。两种写法都要认（2026-07-30 实测）：
        #   subprocess / _load 用带后缀的  "delete_chapter.py"
        #   exec_script 用**不带后缀的模块名**  exec_script("delete_chapter", argv)
        # 只认后缀那种时，81/151 被误判成孤儿 —— 而组模块（audit/strip/chapter…）
        # 恰恰全用不带后缀的写法，等于把整个分派层判没了。
        # 宁可多算一条边：孤儿名单是拿来退役的，错杀比放过贵。
        if tree is not None:
            for node in ast.walk(tree):
                if isinstance(node, ast.Constant) and isinstance(node.value, str):
                    s = node.value
                    t = known.get(s) or (stems.get(s) if len(s) >= 4 else None)
                    if t:
                        add(nid, t, "字面量调用")
    return out


DOC_ROOTS = [
    Path.home() / "Dev" / "tools" / "cc-home" / "skills",
    Path.home() / "Dev" / "tools" / "cc-home" / "commands",
    ROOT / "README.md", ROOT / "README_CN.md", ROOT / "CLAUDE.md",
]


def doc_refs(nodes: dict[str, dict]) -> dict[str, list[str]]:
    """谁被 skill / 命令文档点名 —— 这是第三种「还活着」的证据。

    代码里没人 import 不等于没人用：一多半脚本是**文档告诉我去敲**的
    （`/renumber` 的 SKILL.md 里就写着 `python3 sub/renumber_headings_seq.py …`）。
    少了这一路信号，会把一堆活脚本判成孤儿然后退役掉。

    模式要精确：裸搜 stem 会炸（"styles" 是常用词，裸搜命中 43 处文档，实际只有 1 处
    真在说 sub/styles.py）。所以只认 `<stem>.py` / `sub/<stem>` / `sub.<stem>`。
    """
    refs: dict[str, list[str]] = defaultdict(list)
    files: list[Path] = []
    for r in DOC_ROOTS:
        if r.is_dir():
            files += [p for p in r.rglob("*.md")]
        elif r.is_file():
            files.append(r)
    for k, n in nodes.items():
        stem = n["path"].rsplit("/", 1)[-1][:-3]
        pat = re.compile(rf"(?:^|[^\w]){re.escape(stem)}\.py|sub/{re.escape(stem)}\b"
                         rf"|sub\.{re.escape(stem)}\b")
        for f in files:
            try:
                if pat.search(f.read_text(encoding="utf-8", errors="replace")):
                    refs[k].append(str(f).replace(str(Path.home()), "~"))
            except OSError:
                continue
    return refs


# ─── 层（脚本住在哪一层）────────────────────────────────────────────────
LAYERS = [
    ("测试", r"/tests?/|/test_"),
    ("子命令实现", r"^scripts/document/sub/"),
    ("入口", r"^scripts/document/[^/]+\.py$"),
    ("数据转换", r"^scripts/data/"),
    ("公共库", r"^lib/"),
    ("闸门与工具", r"^tools/"),
]


def layer_of(rel: str) -> str:
    for name, pat in LAYERS:
        if re.search(pat, rel):
            return name
    return "其它"


# ─── 格式轴（这个脚本碰哪种文件格式）──────────────────────────────────────
# ⚠ 与「边」不同：边是 ast 解出来的**事实**，格式轴是**信号计数的启发式**。
# 页面上单独标注了这一点，别把两者当同一等级的证据。
# md 的阈值是 2：`.md` 在 docstring 里指 SKILL.md / README.md 的太多，
# 单次命中几乎全是「文档提到过」而不是「处理 md 文件」。
FORMAT_SIG = {
    "docx": (r"\.docx\b|python-docx|word/document\.xml|^\s*from docx[\s.]|^\s*import docx\b", 1),
    "pdf": (r"\.pdf\b|pymupdf|\bfitz\b|pdfplumber|PyPDF|pdf2image", 1),
    "pptx": (r"\.pptx\b|python-pptx|^\s*from pptx[\s.]|^\s*import pptx\b", 1),
    "xlsx": (r"\.xlsx\b|openpyxl|read_excel|\.xls\b", 1),
    "md": (r"\.md\b|markdown|marker_single", 2),
}


def formats_of(src: str) -> list[str]:
    out = []
    for fmt, (pat, need) in FORMAT_SIG.items():
        if len(re.findall(pat, src, re.M | re.I)) >= need:
            out.append(fmt)
    return out


# ─── 动词轴（93 条 CLI 动词 → 哪个脚本真的接住它）──────────────────────────
# 三种来源，全部读 SSOT 本体，不手抄：
#   ① 表驱动的组   → `sub/_groups.py` 的 `Target.impl`（+ `chain` 里的后续脚本）
#   ② 自带实现的组 → 走 parser 树取 `func.__module__`
#   ③ 扁平命令     → `docx_cli.CMD_TABLE` 里的 cmd_* 函数源码里 `_exec_script("X")`
# 职能标签来自 `sub/_function_axis.py`（同一张被闸门对账的表）。
_EXEC_CALL = re.compile(r'_?exec_script\(\s*["\']([\w./-]+)["\']')


def verb_map() -> tuple[list[dict], str]:
    """→ ([{path, fn, note, impl:[脚本…]}], 出错说明)。解不出来不静默，返回原因。"""
    import argparse as _ap
    import inspect

    sys.path.insert(0, str(ROOT / "scripts" / "document"))
    try:
        import docx_cli
        from sub import _function_axis, _groups
    except Exception as e:                                   # noqa: BLE001
        return [], f"动词轴载入失败：{type(e).__name__}: {e}"

    def impl_of(g, t) -> list[str]:
        """一个 Target 会落到哪几个脚本。

        `chain` 是 `(dest, 同组另一个 target 名)` —— **不是脚本名**。
        照字面收进来会把 `table borders` 的 chain=("center","center") 报成
        「实现脚本 center.py ×2」，而 center 只是同组的一个 target，
        它的实现照样是 table.py（2026-08-02 实测，本页第一版就是这么错的）。
        """
        if isinstance(t, str):
            return [t]
        out = [t.impl]
        if t.chain:
            nxt = (g.targets or {}).get(t.chain[1])
            if nxt is not None and not isinstance(nxt, str) and nxt.impl not in out:
                out.append(nxt.impl)
        return out

    tbl: dict[str, list[str]] = {}
    for g in _groups.GROUPS:
        # 扁平组（chrome / md-merge）的唯一 Target 挂在 `g.flat`，不在 `g.targets`。
        # 漏掉这一支的表现是它们退回 func.__module__ = sub._groups，
        # 报成「实现脚本 _groups.py」—— 看着像有答案，其实是壳的名字。
        if g.flat is not None:
            tbl[g.name] = impl_of(g, g.flat)
        for action, t in (g.targets or {}).items():
            tbl[f"{g.name} {action}"] = impl_of(g, t)

    def leaves(parser, prefix=""):
        subs = [a for a in parser._actions if isinstance(a, _ap._SubParsersAction)]
        if not subs:
            yield prefix.strip(), parser
            return
        seen = set()
        for a in subs:
            for name, sp in a.choices.items():
                if id(sp) in seen:      # 别名与本尊共用 parser 对象，只算一次
                    continue
                seen.add(id(sp))
                yield from leaves(sp, f"{prefix} {name}")

    rows: list[dict] = []
    for path, sp in leaves(docx_cli._build_parser()):
        fn_obj = sp.get_default("func") or docx_cli.CMD_TABLE.get(path)
        mod = getattr(fn_obj, "__module__", None) if fn_obj else None
        impl = tbl.get(path)
        if impl is None and mod:
            if mod.startswith("sub."):
                impl = [mod.split(".")[-1]]
            elif fn_obj is not None:
                try:
                    src = inspect.getsource(fn_obj)
                except (OSError, TypeError):
                    src = ""
                impl = [h.rsplit("/", 1)[-1].removesuffix(".py")
                        for h in _EXEC_CALL.findall(src)] or None
        top, _, action = path.partition(" ")
        hit = _function_axis.lookup(top, action or None)
        rows.append({
            "path": path,
            "fn": hit[0] if hit else "?",
            "note": hit[1] if hit else "",
            "impl": impl or ([mod] if mod == "docx_cli" else []),
        })
    unresolved = [r["path"] for r in rows if not r["impl"]]
    warn = f"{len(unresolved)} 条动词解不出实现脚本：{', '.join(unresolved)}" if unresolved else ""
    return rows, warn


def build() -> dict:
    nodes = collect()
    if not nodes:
        print("⛔ 一个脚本都没扫到 —— 判据坏了，拒绝出空图", file=sys.stderr)
        raise SystemExit(2)
    es = edges(nodes)
    drefs = doc_refs(nodes)
    fwd, bwd = defaultdict(list), defaultdict(list)
    for a, b, why in es:
        fwd[a].append({"id": b, "why": why})
        bwd[b].append({"id": a, "why": why})

    verbs, verb_warn = verb_map()
    if verb_warn:
        print(f"⚠ {verb_warn}", file=sys.stderr)
    by_impl: dict[str, list[dict]] = defaultdict(list)
    for v in verbs:
        for s in v["impl"]:
            by_impl[s].append({"path": v["path"], "fn": v["fn"]})

    for k, n in nodes.items():
        n["engine"] = engine_of(n["src"])
        n["layer"] = layer_of(n["path"])
        n["formats"] = formats_of(n["src"])
        n["verbs"] = sorted(by_impl.get(n["path"].rsplit("/", 1)[-1][:-3], []),
                            key=lambda x: x["path"])
        n["out"] = sorted(fwd[k], key=lambda x: x["id"])
        n["in"] = sorted(bwd[k], key=lambda x: x["id"])
        n["docs"] = sorted(drefs.get(k, []))
        # 孤儿 = 三种「还活着」的证据一条都没有：代码里没人引用、没有文档点名、不是入口。
        # __init__.py 是结构文件，天然没人 import 它的名字，不参与判定。
        n["orphan"] = (not n["in"] and not n["docs"] and not n["is_test"]
                       and n["path"].rsplit("/", 1)[-1] != "__init__.py"
                       and n["path"].rsplit("/", 1)[-1][:-3] not in ENTRIES)
        del n["src"]
    return {"nodes": nodes, "edges": [{"a": a, "b": b, "why": w} for a, b, w in es],
            "verbs": verbs, "verb_warn": verb_warn}


# ─── 渲染 ────────────────────────────────────────────────────────────────────
ENGINE_COLOR = {
    "surgical": "#0a7d55", "裸 lxml+zipfile": "#0a6b8a", "docx_xml": "#0a6b8a",
    "python-docx+收口": "#7a5c00", "python-docx 只读": "#5a6570",
    "python-docx 裸用": "#b3261e", "pandoc": "#6b4ea0", "soffice": "#6b4ea0",
    "不碰 docx 内部": "#8a8f98",
}

LAYER_COLOR = {
    "入口": "#b45309", "子命令实现": "#0a6b8a", "公共库": "#0a7d55",
    "闸门与工具": "#6b4ea0", "数据转换": "#a1554e", "测试": "#8a8f98", "其它": "#8a8f98",
}

FN_COLOR = {
    "format": "#0a6b8a", "content": "#b45309", "review": "#6b4ea0",
    "inspect": "#0a7d55", "convert": "#a1554e", "dispatch": "#5a6570", "?": "#b3261e",
}

# JS/CSS 大段花括号与 f-string 转义互相绞杀，所以模板走占位符替换，不走 f-string。
TEMPLATE = r"""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>doctools 全景图 · 脚本 / 动词 / 引用</title>
<style>
/* 通栏面必须同一个底色 —— 截屏门(html_verify)按「最右两列出现非背景像素」判 full-bleed，
   白 header + 浅灰右栏两种底色会直接判红。层次靠 --tint 卡片和 1px 描边做。 */
:root{--bg:#fff;--tint:#f7f5f2;--line:#e6e4e0;--ink:#22201d;--dim:#6b6560;--hl:#0a6b8a}
*{box-sizing:border-box}
body{margin:0;font:14px/1.6 -apple-system,"PingFang SC",system-ui,sans-serif;background:var(--bg);color:var(--ink)}
header{padding:16px 0 0;border-bottom:1px solid var(--line);background:var(--bg)}
/* 两侧留 16px 真背景边：截屏门按「最右两列出现非背景像素」判 full-bleed。
   ⚠ 约束必须落在 <header> 本身而不是 header>div —— header 的 border-bottom 是通栏的，
   那一行像素直接顶到 x=0 与 x=视口宽，判据当场判红（2026-08-02 实测 bbox=[0,22,1440,…]）。 */
header,.wrap{max-width:1600px;width:calc(100% - 32px);margin-left:auto;margin-right:auto}
h1{margin:0 0 5px;font-size:19px}
.sub{color:var(--dim);font-size:12.5px;line-height:1.55}
.stats{display:flex;gap:20px;margin:11px 0 0;flex-wrap:wrap}
.stat b{font-size:21px;display:block;line-height:1.2;font-variant-numeric:tabular-nums}
.stat span{color:var(--dim);font-size:11.5px}
.tabs{display:flex;gap:2px;margin-top:12px}
.tab{padding:7px 15px;border:1px solid var(--line);border-bottom:none;border-radius:7px 7px 0 0;
     cursor:pointer;font-size:13px;color:var(--dim);background:var(--tint);position:relative;top:1px}
.tab.on{background:var(--bg);color:var(--ink);font-weight:600;border-bottom:1px solid var(--bg)}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;padding:11px 0;border-bottom:1px solid var(--line)}
.bar input,.bar select{padding:6px 9px;border:1px solid var(--line);border-radius:6px;font:inherit;
     background:#fff;color:var(--ink)}
.bar input{width:210px}
.bar input[type=checkbox]{width:auto;margin:0}   /* 不写这条会被上一行撑成 210px 宽的空白 */
.bar label{font-size:12px;color:var(--dim);display:flex;align-items:center;gap:5px}
.hint{margin-left:auto;font-size:11.5px;color:var(--dim)}
.wrap{display:grid;grid-template-columns:1fr minmax(320px,400px);gap:0;align-items:start}
.main{min-width:0;padding:14px 18px 30px 0}
.side{border-left:1px solid var(--line);padding:18px 0 30px 20px;position:sticky;top:0;
      max-height:100vh;overflow-y:auto}
@media(max-width:1000px){.wrap{grid-template-columns:1fr}.side{position:static;max-height:none;
      border-left:none;border-top:1px solid var(--line);padding-left:0}}
svg{width:100%;height:700px;border:1px solid var(--line);border-radius:8px;background:var(--bg);
    display:block;cursor:grab}
svg.drag{cursor:grabbing}
.ln{stroke:#d8d5d0;stroke-width:1}
.ln.hot{stroke:var(--hl);stroke-width:1.8}
.nd{cursor:pointer;stroke:#fff;stroke-width:1.5}
.nd.dim{opacity:.13}
.nd.sel{stroke:#22201d;stroke-width:2.5}
.lb{font-size:9.5px;fill:var(--dim);pointer-events:none;user-select:none}
.lb.dim{opacity:.12}
/* 宽表必须在自己的容器里横向滚动 —— 否则窄屏(900px)下它顶到视口边，
   截屏门按「最右两列有非背景像素」判 full-bleed 直接红（2026-08-02 实测）。 */
#vlist,#vverb{overflow-x:auto}
table{border-collapse:collapse;width:100%;min-width:720px;font-size:13px}
th{text-align:left;font-weight:600;color:var(--dim);font-size:11.5px;letter-spacing:.04em;
   padding:7px 9px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--bg);
   cursor:pointer;white-space:nowrap}
th:hover{color:var(--ink)}
td{padding:6px 9px;border-bottom:1px solid #f2f0ed;vertical-align:top}
tbody tr{cursor:pointer}
tbody tr:hover{background:var(--tint)}
tbody tr.on{background:#e8f1f5;box-shadow:inset 3px 0 0 var(--hl)}
.mono{font-family:ui-monospace,Menlo,monospace;font-size:12px}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.pill{display:inline-block;font-size:11px;padding:1px 7px;border-radius:999px;border:1px solid var(--line);
      background:var(--tint);white-space:nowrap}
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;vertical-align:1px}
h2{font-size:16px;margin:0 0 3px}
.path{font-family:ui-monospace,Menlo,monospace;font-size:11.5px;color:var(--dim);margin-bottom:11px;
      word-break:break-all}
.badges{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:13px}
.box{border:1px solid var(--line);border-radius:8px;background:var(--tint);padding:12px;margin-bottom:13px}
.box h3{margin:0 0 8px;font-size:12px;color:var(--dim);font-weight:600}
.lnk{display:block;padding:3px 0;color:var(--hl);cursor:pointer;border-bottom:1px dotted transparent;font-size:13px}
.lnk:hover{border-bottom-color:var(--hl)}
.why{color:var(--dim);font-size:11px}
.empty{color:var(--dim);font-style:italic;font-size:12.5px}
.doc{background:var(--tint);border-left:3px solid var(--line);padding:9px 12px;border-radius:0 6px 6px 0;
     margin-bottom:13px;color:#3a3733;font-size:12.5px}
.note{padding:11px 13px;border:1px solid #f0d9d6;background:#fdf5f4;border-radius:8px;font-size:12.5px}
.warn{padding:9px 13px;border:1px solid #f0d9d6;background:#fdf5f4;border-radius:8px;font-size:12.5px;
      margin:11px 0}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:var(--dim);margin:9px 0 0}
.foot{color:var(--dim);font-size:11.5px;padding:16px 0 30px;border-top:1px solid var(--line);margin-top:18px}
.hide{display:none}
</style></head><body>
<header><div>
  <h1>doctools 全景图 · 脚本 / 动词 / 引用</h1>
  <div class="sub">扫 <code>scripts/ lib/ tools/</code> 全部 .py。<b>边是 ast 解出来的事实</b>（import + 字面量调用，dispatcher 的命令表就长这样）；
  <b>动词→脚本</b>读的是 <code>sub/_groups.py</code> 的 <code>Target.impl</code> + parser 树的 <code>func.__module__</code> + <code>CMD_TABLE</code> 源码，也不是手抄；
  <b>格式轴那一列是信号计数的启发式</b>，与前两者不同级，别当事实用。生成自 <code>tools/script_graph.py</code>，git __REV__。</div>
  <div class="stats">
    <div class="stat"><b>__NTOTAL__</b><span>脚本</span></div>
    <div class="stat"><b>__NEDGE__</b><span>引用关系</span></div>
    <div class="stat"><b>__NVERB__</b><span>CLI 动词</span></div>
    <div class="stat"><b style="color:__ORPHCOLOR__">__NORPH__</b><span>没人引用（退役候选）</span></div>
    <div class="stat"><b>__NLAYER__</b><span>层</span></div>
  </div>
  __WARNBOX__
  <div class="tabs">
    <div class="tab on" data-v="graph">图谱</div>
    <div class="tab" data-v="list">清单 · __NTOTAL__ 脚本</div>
    <div class="tab" data-v="verb">动词 · __NVERB__ 条</div>
  </div>
</div></header>
<div class="wrap">
  <div class="main">
    <div class="bar">
      <input id="q" placeholder="搜脚本名 / 功能 / 动词…">
      <select id="fl"><option value="">全部层</option>__LAYEROPTS__</select>
      <select id="fe"><option value="">全部底层引擎</option>__ENGINEOPTS__</select>
      <select id="ff"><option value="">全部格式</option>__FMTOPTS__</select>
      <select id="fn" class="hide"><option value="">全部职能</option>__FNOPTS__</select>
      <label><input type="checkbox" id="oo">只看没人引用的</label>
      <span class="hint" id="cnt"></span>
    </div>
    <div id="vgraph">
      <svg id="svg"><g id="cam"><g id="links"></g><g id="nodes"></g><g id="labels"></g></g></svg>
      <div class="legend" id="legend"></div>
      <div class="legend">滚轮缩放 · 拖背景平移 · 拖节点固定位置 · 点节点看双链 · 双击背景复位</div>
    </div>
    <div id="vlist" class="hide"></div>
    <div id="vverb" class="hide"></div>
    <div class="foot">孤儿判据是三条证据全无：代码里没人引用 + 没有文档点名 + 不是顶层入口。
      「代码里没人 import」<b>单独不算死</b> —— 一多半脚本是文档告诉我去敲的。</div>
  </div>
  <div class="side" id="detail"></div>
</div>
<script>
const N = __PAYLOAD__, VERBS = __VERBS__;
const EC = __EC__, LC = __LC__, FC = __FC__;
const $ = i => document.getElementById(i);
let cur = null, view = 'graph', sortKey = 'id', sortDir = 1;

/* ── 筛选（三视图共用同一套） ───────────────────────────────────── */
function pass(n){
  const q = $('q').value.trim().toLowerCase();
  return (!q || n.id.toLowerCase().includes(q) || (n.doc||'').toLowerCase().includes(q)
          || n.verbs.some(v => v.path.toLowerCase().includes(q)))
      && (!$('fl').value || n.layer === $('fl').value)
      && (!$('fe').value || n.engine === $('fe').value)
      && (!$('ff').value || n.formats.includes($('ff').value))
      && (!$('oo').checked || n.orphan);
}
function passVerb(v){
  const q = $('q').value.trim().toLowerCase();
  return (!q || v.path.toLowerCase().includes(q) || v.impl.join(' ').toLowerCase().includes(q))
      && (!$('fn').value || v.fn === $('fn').value);
}

/* ── 力导向图谱 ─────────────────────────────────────────────────── */
const svg = $('svg'), cam = $('cam');
const ids = Object.keys(N);
/* 不用 viewBox：viewBox 一上，屏幕坐标就 ≠ svg 用户坐标，拖节点 / 平移 / 缩放
   三处换算各写一遍，preserveAspectRatio 的信箱留白还会让换算再差一个偏移量。
   直接按元素像素宽布局 → 鼠标坐标 1:1，三处换算全省掉。 */
const H = 700;
const W = svg.clientWidth || 1100;
const P = {};                                   // id -> {x,y,vx,vy,pin}
ids.forEach((id,i) => {
  const a = i / ids.length * Math.PI * 2, r = 150 + (i % 7) * 42;
  P[id] = {x: W/2 + Math.cos(a)*r, y: H/2 + Math.sin(a)*r, vx:0, vy:0, pin:false};
});
const E = [];
ids.forEach(id => N[id].out.forEach(o => { if(N[o.id]) E.push([id, o.id, o.why]); }));
const rad = n => 4 + Math.min(9, Math.sqrt(n.lines) / 4.5);

function tick(){
  for(let a=0;a<ids.length;a++) for(let b=a+1;b<ids.length;b++){
    const p = P[ids[a]], q = P[ids[b]];
    let dx = q.x-p.x, dy = q.y-p.y, d2 = dx*dx+dy*dy || 1;
    if(d2 > 176400) continue;                   // 420px 外不算，省一半计算
    const f = 2400 / d2, d = Math.sqrt(d2);
    const fx = dx/d*f, fy = dy/d*f;
    p.vx -= fx; p.vy -= fy; q.vx += fx; q.vy += fy;
  }
  E.forEach(([a,b]) => {
    const p = P[a], q = P[b];
    const dx = q.x-p.x, dy = q.y-p.y, d = Math.hypot(dx,dy) || 1;
    const f = (d - 96) * 0.011;
    p.vx += dx/d*f; p.vy += dy/d*f; q.vx -= dx/d*f; q.vy -= dy/d*f;
  });
  ids.forEach(id => {
    const p = P[id];
    p.vx += (W/2 - p.x) * 0.0016; p.vy += (H/2 - p.y) * 0.0016;
    if(p.pin){ p.vx = p.vy = 0; return; }
    p.vx *= 0.86; p.vy *= 0.86;
    p.x += Math.max(-14, Math.min(14, p.vx));
    p.y += Math.max(-14, Math.min(14, p.vy));
  });
}

function paintGraph(){
  const vis = new Set(ids.filter(id => pass(N[id])));
  const hot = new Set(cur ? [cur, ...N[cur].in.map(x=>x.id), ...N[cur].out.map(x=>x.id)] : []);
  $('links').innerHTML = E.map(([a,b]) => {
    const on = cur && (a===cur || b===cur);
    return `<line class="ln${on?' hot':''}" x1="${P[a].x}" y1="${P[a].y}" x2="${P[b].x}" y2="${P[b].y}"
             opacity="${vis.has(a)&&vis.has(b) ? (cur ? (on?1:.22) : .5) : .05}"/>`;
  }).join('');
  $('nodes').innerHTML = ids.map(id => {
    const n = N[id];
    return `<circle class="nd${vis.has(id)?'':' dim'}${id===cur?' sel':''}" data-id="${id}"
             cx="${P[id].x}" cy="${P[id].y}" r="${rad(n)}"
             fill="${n.orphan ? '#b3261e' : (LC[n.layer]||'#999')}"><title>${id} · ${n.layer} · ${n.lines} 行</title></circle>`;
  }).join('');
  $('labels').innerHTML = ids.filter(id => vis.has(id) &&
      (cur ? hot.has(id) : N[id].lines > 260 || N[id].in.length + N[id].out.length >= 7))
    .map(id => `<text class="lb" x="${P[id].x + rad(N[id]) + 3}" y="${P[id].y + 3.5}">${id}</text>`).join('');
  $('nodes').querySelectorAll('.nd').forEach(c => {
    c.onclick = ev => { ev.stopPropagation(); show(c.dataset.id); };
    c.onmousedown = ev => startDrag(ev, c.dataset.id);
  });
}

let sim = 0;
function run(n){ sim = Math.max(sim, n); }
(function loop(){
  if(sim > 0){ sim--; for(let k=0;k<3;k++) tick(); if(view==='graph') paintGraph(); }
  requestAnimationFrame(loop);
})();
run(320);

/* 视口：缩放 + 平移 */
let vb = {x:0, y:0, k:1};
function applyCam(){ cam.setAttribute('transform', `translate(${vb.x},${vb.y}) scale(${vb.k})`); }
svg.addEventListener('wheel', e => {
  e.preventDefault();
  const r = svg.getBoundingClientRect(), mx = e.clientX-r.left, my = e.clientY-r.top;
  const k2 = Math.max(.25, Math.min(4, vb.k * (e.deltaY < 0 ? 1.12 : .89)));
  vb.x = mx - (mx - vb.x) * (k2/vb.k); vb.y = my - (my - vb.y) * (k2/vb.k); vb.k = k2;
  applyCam();
}, {passive:false});
let pan = null, drag = null;
svg.onmousedown = e => { if(!drag){ pan = {x:e.clientX-vb.x, y:e.clientY-vb.y}; svg.classList.add('drag'); } };
function startDrag(e, id){ e.stopPropagation(); drag = id; P[id].pin = true; }
window.onmousemove = e => {
  if(drag){
    const r = svg.getBoundingClientRect();
    P[drag].x = (e.clientX-r.left-vb.x)/vb.k; P[drag].y = (e.clientY-r.top-vb.y)/vb.k;
    paintGraph();
  } else if(pan){ vb.x = e.clientX-pan.x; vb.y = e.clientY-pan.y; applyCam(); }
};
window.onmouseup = () => { drag = null; pan = null; svg.classList.remove('drag'); };
svg.ondblclick = () => { vb = {x:0,y:0,k:1}; applyCam(); ids.forEach(i=>P[i].pin=false); run(160); };

/* ── 清单视图 ───────────────────────────────────────────────────── */
const COLS = [
  ['id','脚本',        n => `<span class="dot" style="background:${LC[n.layer]||'#999'}"></span><b>${n.id}</b>${n.orphan?' <span class="pill" style="border-color:#e5a9a4;color:#b3261e">孤</span>':''}`],
  ['layer','层',       n => `<span class="pill">${n.layer}</span>`],
  ['verbn','动词',     n => n.verbs.length ? `<span class="mono">${n.verbs.map(v=>v.path).join('<br>')}</span>` : '<span class="empty">—</span>'],
  ['engine','底层引擎', n => `<span style="color:${EC[n.engine]||'#333'}">${n.engine}</span>`],
  ['fmt','格式',       n => n.formats.length ? n.formats.map(f=>`<span class="pill">${f}</span>`).join(' ') : '<span class="empty">—</span>'],
  ['lines','行',       n => `<span class="num">${n.lines}</span>`],
  ['deg','入←→出',     n => `<span class="num">${n.in.length} ← → ${n.out.length}</span>`],
  ['doc','功能（首行 docstring）', n => n.doc ? n.doc : '<span class="empty">无 docstring</span>'],
];
const KEY = {id:n=>n.id, layer:n=>n.layer, verbn:n=>-n.verbs.length, engine:n=>n.engine,
             fmt:n=>n.formats.join(','), lines:n=>-n.lines, deg:n=>-(n.in.length+n.out.length), doc:n=>n.doc||''};
function paintList(){
  const rows = Object.values(N).filter(pass).sort((a,b) => {
    const ka = KEY[sortKey](a), kb = KEY[sortKey](b);
    return (typeof ka === 'number' ? ka-kb : String(ka).localeCompare(String(kb))) * sortDir;
  });
  $('vlist').innerHTML = `<table><thead><tr>${COLS.map(c =>
      `<th data-k="${c[0]}">${c[1]}${sortKey===c[0] ? (sortDir>0?' ▲':' ▼') : ''}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(n => `<tr class="${cur===n.id?'on':''}" data-id="${n.id}">${
      COLS.map(c => `<td>${c[2](n)}</td>`).join('')}</tr>`).join('')}</tbody></table>`
    || '<div class="empty">没有匹配的脚本</div>';
  $('vlist').querySelectorAll('th').forEach(th => th.onclick = () => {
    if(sortKey === th.dataset.k) sortDir *= -1; else { sortKey = th.dataset.k; sortDir = 1; }
    paintList();
  });
  $('vlist').querySelectorAll('tbody tr').forEach(tr => tr.onclick = () => show(tr.dataset.id));
}

/* ── 动词视图 ───────────────────────────────────────────────────── */
function paintVerb(){
  const rows = VERBS.filter(passVerb);
  $('vverb').innerHTML = `<table><thead><tr><th>动词</th><th>职能</th><th>实现脚本</th><th>说明</th></tr></thead>
    <tbody>${rows.map(v => `<tr data-id="${v.impl[0]||''}">
      <td class="mono"><b>docx_cli.py ${v.path}</b></td>
      <td><span class="pill" style="border-color:${FC[v.fn]||'#ccc'};color:${FC[v.fn]||'#333'}">${v.fn}</span></td>
      <td class="mono">${v.impl.length ? v.impl.map(s=>`<span class="lnk" data-go="${s}">${s}.py</span>`).join('') : '<span class="empty">解不出</span>'}</td>
      <td>${v.note || '<span class="empty">—</span>'}</td></tr>`).join('')}</tbody></table>`;
  $('vverb').querySelectorAll('[data-go]').forEach(a => a.onclick = ev => { ev.stopPropagation(); show(a.dataset.go); });
  $('vverb').querySelectorAll('tbody tr').forEach(tr => tr.onclick = () => tr.dataset.id && show(tr.dataset.id));
}

/* ── 右栏详情 ───────────────────────────────────────────────────── */
function show(id){
  const n = N[id]; if(!n){ return; }
  cur = id;
  const links = a => a.length
    ? a.map(x => `<span class="lnk" data-go="${x.id}">${x.id} <span class="why">· ${x.why}</span></span>`).join('')
    : '<div class="empty">没有</div>';
  $('detail').innerHTML = `
    <h2>${n.id}</h2>
    <div class="path">${n.path} · ${n.lines} 行</div>
    <div class="badges">
      <span class="pill" style="border-color:${LC[n.layer]||'#ccc'}">${n.layer}</span>
      <span class="pill" style="border-color:${EC[n.engine]||'#ccc'};color:${EC[n.engine]||'#333'}">${n.engine}</span>
      <span class="pill">${n.family} 族</span>
      ${n.formats.map(f=>`<span class="pill">${f}</span>`).join('')}
      ${n.orphan?'<span class="pill" style="border-color:#e5a9a4;color:#b3261e">没人引用 · 退役候选</span>':''}
      ${n.is_test?'<span class="pill">测试件</span>':''}
    </div>
    ${n.doc?`<div class="doc">${n.doc}</div>`:''}
    ${n.verbs.length ? `<div class="box"><h3>⌨️ 它接住的 CLI 动词（${n.verbs.length}）</h3>${
      n.verbs.map(v=>`<div class="mono" style="padding:2px 0">docx_cli.py ${v.path}
        <span class="pill" style="border-color:${FC[v.fn]||'#ccc'};color:${FC[v.fn]||'#333'}">${v.fn}</span></div>`).join('')}</div>` : ''}
    <div class="box"><h3>← 谁引用了它（${n.in.length}）</h3>${links(n.in)}</div>
    <div class="box"><h3>它引用了谁（${n.out.length}）→</h3>${links(n.out)}</div>
    <div class="box"><h3>📄 哪些 skill / 命令文档点了它的名（${n.docs.length}）</h3>
      ${n.docs.length ? n.docs.map(d=>`<div class="mono" style="padding:2px 0">${d}</div>`).join('')
                      : '<div class="empty">没有 —— 代码里也没人引用的话，它就是悬着的</div>'}</div>
    ${n.orphan?`<div class="note"><b>没人引用它</b>，也不是顶层入口。要么它是我手敲命令直接跑的（那就该登记进 docx_cli 的命令表或 CLAUDE.md 的独立入口表），要么它已经没人用了（那就该退役）。别让它悬着 —— 悬着的脚本会在下次横切改动时被漏掉。</div>`:''}`;
  $('detail').querySelectorAll('[data-go]').forEach(a => a.onclick = () => show(a.dataset.go));
  redraw();
}

/* ── 视图切换 / 重绘 ────────────────────────────────────────────── */
function redraw(){
  const nv = Object.values(N).filter(pass).length;
  $('cnt').textContent = view === 'verb'
    ? `${VERBS.filter(passVerb).length} / ${VERBS.length} 条动词`
    : `${nv} / ${Object.keys(N).length} 个脚本`;
  if(view === 'graph'){ paintGraph(); } else if(view === 'list'){ paintList(); } else { paintVerb(); }
}
document.querySelectorAll('.tab').forEach(t => t.onclick = () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('on', x === t));
  view = t.dataset.v;
  ['vgraph','vlist','vverb'].forEach(v => $(v).classList.toggle('hide', v !== 'v'+view));
  $('fn').classList.toggle('hide', view !== 'verb');
  ['fl','fe','ff'].forEach(i => $(i).classList.toggle('hide', view === 'verb'));
  redraw();
});
['q','fl','fe','ff','fn','oo'].forEach(i => $(i).addEventListener('input', () => { redraw(); }));
$('legend').innerHTML = Object.entries(LC).filter(([k]) => k !== '其它')
  .map(([k,v]) => `<span><span class="dot" style="background:${v}"></span>${k}</span>`).join('')
  + '<span><span class="dot" style="background:#b3261e"></span>没人引用</span><span>圆越大 = 行数越多</span>';
applyCam();
show(N['docx_cli'] ? 'docx_cli' : Object.keys(N).sort()[0]);
</script></body></html>"""


def render(data: dict) -> str:
    ns = data["nodes"]
    n_orphan = sum(1 for n in ns.values() if n["orphan"])
    engines = sorted({n["engine"] for n in ns.values()})
    layers = sorted({n["layer"] for n in ns.values()})
    fmts = sorted({f for n in ns.values() for f in n["formats"]})
    fns = sorted({v["fn"] for v in data["verbs"]})
    rev = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"],
                         capture_output=True, text=True).stdout.strip()

    def opts(vals):
        return "".join(f"<option>{html.escape(v)}</option>" for v in vals)

    warn = (f'<div class="warn">⚠ {html.escape(data["verb_warn"])} —— 动词视图不完整，'
            f"别拿它当全集。</div>") if data.get("verb_warn") else ""
    sub = {
        "__PAYLOAD__": json.dumps(ns, ensure_ascii=False),
        "__VERBS__": json.dumps(data["verbs"], ensure_ascii=False),
        "__EC__": json.dumps(ENGINE_COLOR, ensure_ascii=False),
        "__LC__": json.dumps(LAYER_COLOR, ensure_ascii=False),
        "__FC__": json.dumps(FN_COLOR, ensure_ascii=False),
        "__NTOTAL__": str(len(ns)),
        "__NEDGE__": str(len(data["edges"])),
        "__NVERB__": str(len(data["verbs"])),
        "__NORPH__": str(n_orphan),
        "__NLAYER__": str(len(layers)),
        "__ORPHCOLOR__": "#b3261e" if n_orphan else "#0a7d55",
        "__LAYEROPTS__": opts(layers),
        "__ENGINEOPTS__": opts(engines),
        "__FMTOPTS__": opts(fmts),
        "__FNOPTS__": opts(fns),
        "__WARNBOX__": warn,
        "__REV__": rev,
    }
    out = TEMPLATE
    for k, v in sub.items():
        out = out.replace(k, v)
    return out

def main() -> int:
    ap = argparse.ArgumentParser(description="doctools 仓内脚本引用关系 · 双链页面")
    ap.add_argument("--open", action="store_true")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("-o", "--out", default=str(ROOT / "reports" / "script-graph.html"))
    a = ap.parse_args()
    data = build()
    if a.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render(data), encoding="utf-8")
    orph = [k for k, v in data["nodes"].items() if v["orphan"]]
    print(f"{len(data['nodes'])} 个脚本 · {len(data['edges'])} 条引用 · "
          f"{len(orph)} 个没人引用 → {out}")
    if a.open:
        subprocess.run(["open", str(out)])
    return 0


if __name__ == "__main__":
    sys.exit(main())
