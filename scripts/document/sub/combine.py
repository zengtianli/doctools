"""combine.py — group module: combine N docx into 1 (docxcompose, media-safe).

`combine` 是 `split by-h1` 的逆操作: 把多个章节 docx 按给定顺序合并成一个整本。
第一个输入领头(其 styles / sectPr / page-setup 为基)。媒体 / numbering / styles 由
docxcompose 安全合并(不会丢图、不串号)。

⚠️ 注意: 手工定稿成品里独有的封面 / 目录 / 附图 / 院版面装帧 **不在章节件里**, combine
只是把输入"原样按序拼接"——得到的是正文合并体, 不是定稿成品的逐位复刻。装帧 live 在
master 成品里(改 master 再 `split by-h1` 刷新章节, 是 2026-06-07 钦定的主方向)。

Standalone CLI:
    python3 sub/combine.py --docx 01.docx 02.docx ... 06.docx --out 整本.docx
docx_cli:
    docx_cli combine --docx 01.docx 02.docx ... 06.docx --out 整本.docx

Trigger scenarios:
  - 章节件分别编辑/审校后, 反向拼回整本(报告章↔成品双向转换的 compose 方向)
  - 多份独立 docx 顺序合并成一册
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    from docx import Document
    from docxcompose.composer import Composer
except ImportError:  # pragma: no cover
    Document = None
    Composer = None


def run_combine(inputs, out_path, verbose: bool = True) -> dict:
    """Merge `inputs` (ordered list of docx paths) → `out_path` via docxcompose.

    First input leads (styles/sectPr). Returns a report dict with exit_code.
    """
    if Document is None or Composer is None:
        return {"error": "python-docx / docxcompose 未安装 (pip install python-docx docxcompose)",
                "exit_code": 2}
    paths = [Path(p).expanduser().resolve() for p in inputs]
    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        return {"error": f"input(s) not found: {missing}", "exit_code": 2}
    if len(paths) < 2:
        return {"error": "combine 需 ≥2 个输入 docx", "exit_code": 2}

    out = Path(out_path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    master = Document(str(paths[0]))
    composer = Composer(master)
    for p in paths[1:]:
        composer.append(Document(str(p)))
    composer.save(str(out))

    nbytes = out.stat().st_size
    if verbose:
        print(f"[combine] {len(paths)} docx → {out}  ({nbytes:,} bytes)")
        for i, p in enumerate(paths):
            print(f"  [{i}] {p.name}")
    return {"inputs": [str(p) for p in paths], "out": str(out),
            "count": len(paths), "bytes": nbytes, "exit_code": 0}


def _run(args) -> int:
    rep = run_combine(args.docx, args.out, verbose=True)
    if rep.get("error"):
        print(f"[sub.combine] ERROR: {rep['error']}", file=sys.stderr)
    return rep.get("exit_code", 0)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "combine",
        help="combine N docx into 1 (docxcompose, media-safe; inverse of split by-h1)",
    )
    p.add_argument("--docx", required=True, nargs="+",
                   help="input docx paths IN ORDER (first leads styles/sectPr)")
    p.add_argument("--out", required=True, help="output combined docx path")
    p.set_defaults(func=_run)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Combine N docx into 1 (docxcompose, media-safe). Inverse of split by-h1.",
    )
    ap.add_argument("--docx", required=True, nargs="+",
                    help="input docx paths IN ORDER (first leads styles/sectPr)")
    ap.add_argument("--out", required=True, help="output combined docx path")
    args = ap.parse_args()
    rep = run_combine(args.docx, args.out, verbose=True)
    if rep.get("error"):
        print(f"ERROR: {rep['error']}", file=sys.stderr)
    return rep.get("exit_code", 0)


if __name__ == "__main__":
    sys.exit(main())
