#!/usr/bin/env python3
"""盘点所有和 docx 有关的脚本，按「底层用了什么」分类。判据全部来自读源码，不靠记忆。\n\n    python3 handoffs/_inventory.py --md   # 出清单（docx-scripts-inventory.md 的数字来源）\n    python3 handoffs/_inventory.py        # 出 JSON\n\n枚举为空一律 rc=2 —— 拒绝在空集上报「没有 docx 脚本」。\n"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOTS = [Path.home() / "Dev", Path.home() / "Work", Path.home() / "Apps"]
SKIP = re.compile(r"/(_archive|_scratch|\.venv|node_modules|site-packages|jobs)/"
                  r"|\.app/|/bak/|_bak/|_backup-pre-shim|_pre-shim-|_scratch-|/09-归档/"
                  r"|/\.git/")

# python-docx / 收口 两条判据**从守卫本体 import**，不在这里抄一份。
# 抄一份的代价 2026-08-01 实测到了：本文件曾写成 `from docx[\w.]*\s+import`，
# 于是 `from docx_parts import assert_parts_intact` 被判成「裸用 python-docx 写盘」
# —— pptx_cli.py（用的是 python-pptx，一行 python-docx 都没有）被误列进裸用名册。
# check_docx_collar.py 2026-07-31 修过同一条正则，这份没跟上：判据分居两处 = 必漂。
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from check_docx_collar import COLLAR as _COLLAR, IMPORTS_DOCX as _IMPORTS_DOCX  # noqa: E402

SIG = {
    "python-docx":  _IMPORTS_DOCX,
    "收口":         _COLLAR,
    "docx_surgical": re.compile(r"\bdocx_surgical\b|\bsurgical_rewrite"),
    "docx_xml":     re.compile(r"\bfrom docx_xml import|\bimport docx_xml\b"),
    "裸 zipfile":   re.compile(r"^\s*import zipfile\b", re.M),
    "裸 lxml":      re.compile(r"^\s*(from lxml|import lxml)\b", re.M),
    "docxcompose":  re.compile(r"\bdocxcompose\b"),
    "pandoc":       re.compile(r"[\"']pandoc[\"']|\bpandoc\b"),
    "soffice":      re.compile(r"\bsoffice\b|libreoffice"),
    "写盘":         re.compile(r"\.save\("),
}


def touches_docx(src: str, p: Path) -> bool:
    return (".docx" in src or SIG["python-docx"].search(src)
            or SIG["docx_xml"].search(src) or SIG["docx_surgical"].search(src)
            or "docx" in p.name)


def engine(f: dict) -> str:
    """一个脚本的「底层是什么」——按能力强弱定序，取最能说明问题的那一条。"""
    if f["docx_surgical"]:
        return "surgical（lxml+zipfile 只重写点名部件）"
    if f["python-docx"] and f["写盘"]:
        return "python-docx + 收口" if f["收口"] else "python-docx（裸，没收口）"
    if f["python-docx"]:
        return "python-docx（只读，不写盘）"
    if f["docx_xml"]:
        return "docx_xml（元素级遍历，lxml）"
    if f["裸 zipfile"] and f["裸 lxml"]:
        return "裸 lxml+zipfile"
    if f["pandoc"]:
        return "pandoc（外部进程）"
    if f["soffice"]:
        return "soffice（外部进程）"
    return "其它（不直接碰 docx 内部）"


def summarize(rows: list[dict]) -> None:
    """--md：按「区块 × 底层引擎」出清单，docx-scripts-inventory.md 的数字全部来自这里。"""
    import collections
    H = str(Path.home())

    def zone(p: str) -> str:
        if "/doctools/scripts/document/sub/" in p: return "doctools sub/（子命令实现）"
        if "/doctools/scripts/document/" in p: return "doctools document/（顶层工具）"
        if "/doctools/" in p: return "doctools 其它（lib/tools/tests）"
        if "/Dev/tools/dev/" in p: return "总部引擎 ~/Dev/tools/dev"
        if "/Dev/tools/cc-home/" in p: return "harness ~/.claude"
        if "/Work/" in p: return "~/Work 交付线"
        if "/Apps/" in p: return "~/Apps"
        return "~/Dev 其它"

    real = [r for r in rows if not r["engine"].startswith("其它")]
    print(f"# docx 脚本盘点\n\n共 {len(rows)} 个提到 docx，其中 {len(real)} 个真的动 docx 内部\n")
    print("## 按底层引擎")
    for e, n in collections.Counter(r["engine"] for r in rows).most_common():
        print(f"- {n} 个 · {e}")
    g = collections.defaultdict(list)
    for r in rows:
        g[(zone(r["path"]), r["engine"])].append(r["path"].replace(H, "~"))
    for k in sorted(g):
        print(f"\n### {k[0]} | {k[1]} | {len(g[k])} 个")
        for x in sorted(g[k]):
            print(f"- {x}")


def main() -> int:
    rows = []
    me = Path(__file__).resolve()   # 别把自己数进去：本文件的判据表里全是 docx 关键词
    for root in ROOTS:
        for p in root.rglob("*.py"):
            if SKIP.search(str(p)) or p.resolve() == me:
                continue
            try:
                src = p.read_text(encoding="utf-8")
            except (OSError, UnicodeDecodeError):
                continue
            if not touches_docx(src, p):
                continue
            f = {k: bool(v.search(src)) for k, v in SIG.items()}
            rows.append({"path": str(p), "engine": engine(f), **f})
    if not rows:
        print("枚举为空 —— 判据坏了，拒绝当成「没有 docx 脚本」", file=sys.stderr)
        return 2
    if "--md" in sys.argv:
        summarize(rows)
    else:
        print(json.dumps(rows, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
