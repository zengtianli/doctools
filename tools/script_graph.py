#!/usr/bin/env python3
"""script_graph.py — 解出 doctools 仓内脚本的真实引用关系，出一张可点的双链页面。

    python3 tools/script_graph.py            # 生成 reports/script-graph.html
    python3 tools/script_graph.py --open
    python3 tools/script_graph.py --json     # 只出数据

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
SCAN = [ROOT / "scripts", ROOT / "lib", ROOT / "tools"]

# 顶层 CLI = 人直接敲的入口，入度 0 也不算孤儿
ENTRIES = {"docx_cli", "pdf_cli", "doc_dispatch", "typeset_pipeline", "doc_gui_backend",
           "pptx_cli", "md_tools", "script_graph", "blast_radius",
           "check_docx_collar", "_inventory",
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

    for k, n in nodes.items():
        n["engine"] = engine_of(n["src"])
        n["out"] = sorted(fwd[k], key=lambda x: x["id"])
        n["in"] = sorted(bwd[k], key=lambda x: x["id"])
        n["docs"] = sorted(drefs.get(k, []))
        # 孤儿 = 三种「还活着」的证据一条都没有：代码里没人引用、没有文档点名、不是入口。
        # __init__.py 是结构文件，天然没人 import 它的名字，不参与判定。
        n["orphan"] = (not n["in"] and not n["docs"] and not n["is_test"]
                       and n["path"].rsplit("/", 1)[-1] != "__init__.py"
                       and n["path"].rsplit("/", 1)[-1][:-3] not in ENTRIES)
        del n["src"]
    return {"nodes": nodes, "edges": [{"a": a, "b": b, "why": w} for a, b, w in es]}


# ─── 渲染 ────────────────────────────────────────────────────────────────────
ENGINE_COLOR = {
    "surgical": "#0a7d55", "裸 lxml+zipfile": "#0a6b8a", "docx_xml": "#0a6b8a",
    "python-docx+收口": "#7a5c00", "python-docx 只读": "#5a6570",
    "python-docx 裸用": "#b3261e", "pandoc": "#6b4ea0", "soffice": "#6b4ea0",
    "不碰 docx 内部": "#8a8f98",
}


def render(data: dict) -> str:
    ns = data["nodes"]
    n_total = len(ns)
    n_orphan = sum(1 for n in ns.values() if n["orphan"])
    fams = sorted({n["family"] for n in ns.values()})
    engines = sorted({n["engine"] for n in ns.values()})
    payload = json.dumps({k: {kk: vv for kk, vv in v.items()} for k, v in ns.items()},
                         ensure_ascii=False)
    rev = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"],
                         capture_output=True, text=True).stdout.strip()
    return f"""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>doctools 脚本关系图 · 双链</title>
<style>
/* 通栏面（body/header/左栏/右栏）必须同一个底色 —— 截屏门(html_verify)按「最右两列
   出现非背景像素」判 full-bleed，白 header + 浅灰右栏两种底色会直接判红。
   层次靠 --tint 的卡片和 1px 描边做，不靠大块底色。 */
:root{{--bg:#fff;--tint:#f7f5f2;--line:#e6e4e0;--ink:#22201d;--dim:#6b6560;--hl:#0a6b8a}}
*{{box-sizing:border-box}}
body{{margin:0;font:14px/1.6 -apple-system,"PingFang SC",system-ui,sans-serif;background:var(--bg);color:var(--ink)}}
header{{padding:18px 24px;border-bottom:1px solid var(--line);background:var(--bg)}}
/* 两侧必须留出真背景边：截屏门(html_verify)按「最右两列出现非背景像素」判 full-bleed，
   通栏 app 布局永远过不了。留 16px 背景边既满足门，也让内容不贴屏幕边。 */
header,.wrap{{max-width:1400px;width:calc(100% - 32px);margin-left:auto;margin-right:auto}}
h1{{margin:0 0 6px;font-size:19px}}
.sub{{color:var(--dim);font-size:13px}}
.stats{{display:flex;gap:22px;margin-top:12px;flex-wrap:wrap}}
.stat b{{font-size:22px;display:block;line-height:1.2}}
.stat span{{color:var(--dim);font-size:12px}}
.wrap{{display:grid;grid-template-columns:minmax(280px,360px) 1fr;gap:0;align-items:start}}
.left{{border-right:1px solid var(--line);background:var(--bg);position:sticky;top:0;max-height:100vh;overflow-y:auto}}
.right{{overflow-y:auto;padding:22px 26px}}
.tools{{padding:12px 14px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--bg);z-index:2}}
input,select{{width:100%;padding:7px 9px;border:1px solid var(--line);border-radius:6px;font:inherit;background:#fff;color:var(--ink);margin-bottom:6px}}
.fam{{padding:7px 14px;font-size:11px;letter-spacing:.08em;color:var(--dim);background:var(--tint);border-top:1px solid var(--line);text-transform:uppercase}}
.item{{padding:7px 14px;cursor:pointer;border-bottom:1px solid #f2f0ed;display:flex;align-items:center;gap:8px}}
.item:hover{{background:var(--tint)}}
.item.on{{background:#e8f1f5;box-shadow:inset 3px 0 0 var(--hl)}}
.item .nm{{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.dot{{width:8px;height:8px;border-radius:50%;flex:none}}
.deg{{font-size:11px;color:var(--dim);font-variant-numeric:tabular-nums}}
.orph{{color:#b3261e;font-size:11px;font-weight:600}}
h2{{font-size:17px;margin:0 0 4px}}
.path{{font-family:ui-monospace,Menlo,monospace;font-size:12px;color:var(--dim);margin-bottom:14px}}
.badges{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}}
.badge{{font-size:12px;padding:3px 9px;border-radius:999px;border:1px solid var(--line);background:var(--tint)}}
.cols{{display:grid;grid-template-columns:1fr 1fr;gap:22px}}
@media(max-width:900px){{.wrap{{grid-template-columns:1fr}}.left{{position:static;max-height:none}}.cols{{grid-template-columns:1fr}}}}
.box{{border:1px solid var(--line);border-radius:8px;background:var(--tint);padding:14px}}
.box h3{{margin:0 0 10px;font-size:13px;color:var(--dim);font-weight:600}}
.lnk{{display:block;padding:4px 0;color:var(--hl);cursor:pointer;text-decoration:none;border-bottom:1px dotted transparent}}
.lnk:hover{{border-bottom-color:var(--hl)}}
.why{{color:var(--dim);font-size:11px}}
.empty{{color:var(--dim);font-style:italic}}
.doc{{background:var(--tint);border-left:3px solid var(--line);padding:10px 13px;border-radius:0 6px 6px 0;margin-bottom:16px;color:#3a3733}}
.note{{margin-top:18px;padding:12px 14px;border:1px solid #f0d9d6;background:#fdf5f4;border-radius:8px;font-size:13px}}
</style></head><body>
<header>
  <h1>doctools 脚本关系图 · 点进去看双链</h1>
  <div class="sub">扫 <code>scripts/ lib/ tools/</code> 全部 .py，边来自 import + 字面量调用（dispatcher 的命令表就长这样）。
  生成自 <code>tools/script_graph.py</code>，git {rev}。工作区级拓扑看 <code>/module-map</code>，那是另一层。</div>
  <div class="stats">
    <div class="stat"><b>{n_total}</b><span>脚本总数</span></div>
    <div class="stat"><b>{len(data['edges'])}</b><span>引用关系</span></div>
    <div class="stat"><b style="color:#b3261e">{n_orphan}</b><span>没人引用（退役候选）</span></div>
    <div class="stat"><b>{len(fams)}</b><span>族</span></div>
  </div>
</header>
<div class="wrap">
  <div class="left">
    <div class="tools">
      <input id="q" placeholder="搜脚本名 / 说明…">
      <select id="fe"><option value="">全部底层</option>{"".join(f'<option>{html.escape(e)}</option>' for e in engines)}</select>
      <label style="font-size:12px;color:var(--dim)"><input type="checkbox" id="oo" style="width:auto;margin:0 6px 0 0">只看没人引用的</label>
    </div>
    <div id="list"></div>
  </div>
  <div class="right" id="detail"></div>
</div>
<script>
const N = {payload};
const EC = {json.dumps(ENGINE_COLOR, ensure_ascii=False)};
const list = document.getElementById('list'), detail = document.getElementById('detail');
let cur = null;

function draw(){{
  const q = document.getElementById('q').value.trim().toLowerCase();
  const fe = document.getElementById('fe').value;
  const oo = document.getElementById('oo').checked;
  const rows = Object.values(N).filter(n =>
    (!q || n.id.toLowerCase().includes(q) || (n.doc||'').toLowerCase().includes(q)) &&
    (!fe || n.engine === fe) && (!oo || n.orphan));
  const byFam = {{}};
  rows.forEach(n => (byFam[n.family] = byFam[n.family] || []).push(n));
  list.innerHTML = Object.keys(byFam).sort().map(f =>
    `<div class="fam">${{f}} · ${{byFam[f].length}}</div>` +
    byFam[f].sort((a,b)=>a.id.localeCompare(b.id)).map(n =>
      `<div class="item ${{cur===n.id?'on':''}}" data-id="${{n.id}}">
         <span class="dot" style="background:${{EC[n.engine]||'#999'}}"></span>
         <span class="nm">${{n.id}}</span>
         ${{n.orphan?'<span class="orph">孤</span>':''}}
         <span class="deg">${{n.in.length}}←→${{n.out.length}}</span>
       </div>`).join('')).join('') || '<div class="fam">没有匹配的脚本</div>';
  list.querySelectorAll('.item').forEach(el => el.onclick = () => show(el.dataset.id));
}}

function show(id){{
  cur = id; const n = N[id]; if(!n) return;
  const links = a => a.length
    ? a.map(x => `<a class="lnk" data-go="${{x.id}}">${{x.id}} <span class="why">· ${{x.why}}</span></a>`).join('')
    : '<div class="empty">没有</div>';
  detail.innerHTML = `
    <h2>${{n.id}}</h2>
    <div class="path">${{n.path}} · ${{n.lines}} 行</div>
    <div class="badges">
      <span class="badge" style="border-color:${{EC[n.engine]||'#ccc'}};color:${{EC[n.engine]||'#333'}}">${{n.engine}}</span>
      <span class="badge">${{n.family}} 族</span>
      ${{n.orphan?'<span class="badge" style="border-color:#e5a9a4;color:#b3261e">没人引用 · 退役候选</span>':''}}
      ${{n.is_test?'<span class="badge">测试件</span>':''}}
    </div>
    ${{n.doc?`<div class="doc">${{n.doc}}</div>`:''}}
    <div class="cols">
      <div class="box"><h3>← 谁引用了它（${{n.in.length}}）</h3>${{links(n.in)}}</div>
      <div class="box"><h3>它引用了谁（${{n.out.length}}）→</h3>${{links(n.out)}}</div>
    </div>
    <div class="box" style="margin-top:22px"><h3>📄 哪些 skill / 命令文档点了它的名（${{n.docs.length}}）</h3>
      ${{n.docs.length ? n.docs.map(d=>`<div style="padding:3px 0;font-family:ui-monospace,Menlo,monospace;font-size:12px">${{d}}</div>`).join('')
                       : '<div class="empty">没有 —— 代码里也没人引用的话，它就是悬着的</div>'}}</div>
    ${{n.orphan?`<div class="note"><b>没人引用它</b>，也不是顶层入口。要么它是通过我手敲命令直接跑的（那就该登记进 docx_cli 的命令表），要么它已经没人用了（那就该退役）。别让它悬着 —— 悬着的脚本会在下次横切改动时被漏掉。</div>`:''}}`;
  detail.querySelectorAll('[data-go]').forEach(a => a.onclick = () => {{
    show(a.dataset.go); draw();
    document.querySelector('.item.on')?.scrollIntoView({{block:'center'}});
  }});
  draw();
}}
['q','fe','oo'].forEach(i => document.getElementById(i).addEventListener('input', draw));
draw();
show(N['docx_cli'] ? 'docx_cli' : Object.keys(N).sort()[0]);  // 默认落在入口，不是字母序第一个
</script></body></html>"""


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
