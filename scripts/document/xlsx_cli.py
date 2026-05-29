#!/usr/bin/env python3
"""xlsx_cli.py — XLSX 处理统一 CLI (2026-05-28)

3 子命令:
  view    探查 sheet/列/sample；--out 写 HTML
  fig     pandas + matplotlib 出图 (svg + png + pdf)，配色源自 html-fig.ramps
  to-db   pandas → sqlite3 入库；--reconcile 全量重建

底座:
  pandas 3.0.2 / openpyxl 3.1.5 / sqlite3 (stdlib) / matplotlib 3.10.8

设计同源 `pdf_cli.py` 二级 dispatch,但所有实现单文件自包含 (~500 行)。
"""

from __future__ import annotations

import argparse
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Optional

# Lazy import pandas — heavy
try:
    import pandas as pd  # type: ignore
except ImportError:
    print("FATAL: pandas not installed", file=sys.stderr)
    sys.exit(2)


SCRIPT_VERSION = "1.0.0"
_HTMLFIG_LIB = (
    Path.home() / "Dev" / "tools" / "cc-home"
    / "skills" / "html-fig" / "lib" / "htmlfig.py"
)
_RSVG = "/opt/homebrew/bin/rsvg-convert"


# ═══════════════════════════════════════════════════════════════════════
# Shared helpers
# ═══════════════════════════════════════════════════════════════════════

def _detect_header(xlsx_path: Path, sheet: Any) -> Any:
    """启发式：行 0 多 NaN + 行 1 多字符串 → merged header → [0, 1]；否则 0。"""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None, nrows=5)
    except Exception:
        return 0
    if df.shape[0] < 2 or df.shape[1] == 0:
        return 0
    nan0 = df.iloc[0].isna().sum()
    str1 = df.iloc[1].apply(lambda x: isinstance(x, str)).sum()
    if nan0 > df.shape[1] / 3 and str1 > df.shape[1] / 2:
        return [0, 1]
    return 0


def _flatten_columns(df: "pd.DataFrame") -> "pd.DataFrame":
    """多级 header → 单层；去 'Unnamed:_x' 噪音。"""
    if df.columns.nlevels > 1:
        flat = []
        for c in df.columns:
            parts = [
                str(x) for x in c
                if not str(x).startswith("Unnamed") and str(x) != "nan"
            ]
            flat.append("_".join(parts).strip("_") or "col")
        df.columns = flat
    else:
        df.columns = [
            re.sub(r"^Unnamed:\s*\d+_level_\d+$", "col", str(c))
            for c in df.columns
        ]
    return df


def _sanitize_table_name(name: str) -> str:
    """sheet 名 → sqlite 表名 (中文/空格/特殊字符 → _)。"""
    s = re.sub(r"[^A-Za-z0-9_一-鿿]+", "_", str(name)).strip("_")
    if not s:
        s = "sheet"
    if s[0].isdigit():
        s = "t_" + s
    return s


def _read_sheet(xlsx_path: Path, sheet: Any) -> "pd.DataFrame":
    """读 sheet 并 flatten 列名 (header 自动 detect)。"""
    header = _detect_header(xlsx_path, sheet)
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=header)
    df = _flatten_columns(df)
    return df


# ═══════════════════════════════════════════════════════════════════════
# 1. view
# ═══════════════════════════════════════════════════════════════════════

def _df_md(df: "pd.DataFrame", max_col_width: int = 30) -> str:
    """简易 markdown 表格输出 (避免依赖 tabulate)。"""
    cols = [str(c)[:max_col_width] for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        rows.append([
            str("" if pd.isna(v) else v)[:max_col_width].replace("\n", " ")
            for v in r
        ])
    head = "| " + " | ".join(cols) + " |"
    sep = "|" + "|".join(["---"] * len(cols)) + "|"
    body = "\n".join("| " + " | ".join(r) + " |" for r in rows)
    return "\n".join([head, sep, body])


def _view_one(xlsx_path: Path, sheet: str, rows: int) -> str:
    """单 sheet 探查文本。"""
    df = _read_sheet(xlsx_path, sheet)
    lines = []
    lines.append(f"### sheet: {sheet}")
    lines.append(f"- 形状: {df.shape[0]} 行 × {df.shape[1]} 列")
    cols = list(df.columns)
    show = cols[:20]
    more = f" (+{len(cols)-20} more)" if len(cols) > 20 else ""
    lines.append("- 列名 (前 20): " + ", ".join(str(c) for c in show) + more)
    lines.append("")
    lines.append(f"前 {rows} 行 sample:")
    lines.append("")
    lines.append(_df_md(df.head(rows)))
    lines.append("")
    return "\n".join(lines)


def _view_html(xlsx_path: Path, sheets: list[str], rows: int, out: Path) -> None:
    css = (
        "body{font-family:-apple-system,sans-serif;max-width:1400px;"
        "margin:2em auto;padding:0 1em;color:#222;}"
        "h1{border-bottom:2px solid #333;}"
        "h2{margin-top:2em;color:#06c;}"
        "table{border-collapse:collapse;margin:1em 0;font-size:.85em;}"
        "th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;}"
        "th{background:#f4f4f4;}"
        ".meta{color:#666;font-size:.9em;}"
    )
    parts = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>xlsx view: {xlsx_path.name}</title>",
        f"<style>{css}</style></head><body>",
        f"<h1>{xlsx_path.name}</h1>",
        f"<p class='meta'>sheets: {len(sheets)}</p>",
    ]
    for s in sheets:
        try:
            df = _read_sheet(xlsx_path, s)
        except Exception as e:
            parts.append(f"<h2>{s}</h2><p>ERROR: {e}</p>")
            continue
        parts.append(f"<h2>{s}</h2>")
        parts.append(
            f"<p class='meta'>{df.shape[0]} 行 × {df.shape[1]} 列</p>"
        )
        parts.append(df.head(rows).to_html(index=False, na_rep=""))
    parts.append("</body></html>")
    out.write_text("\n".join(parts), encoding="utf-8")


def cmd_view(args: argparse.Namespace) -> int:
    xlsx_path = Path(args.xlsx).resolve()
    if not xlsx_path.is_file():
        print(f"ERROR: file not found: {xlsx_path}", file=sys.stderr)
        return 2

    xl = pd.ExcelFile(xlsx_path)
    sheets = [args.sheet] if args.sheet else xl.sheet_names

    if args.out:
        out = Path(args.out).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        _view_html(xlsx_path, sheets, args.rows, out)
        print(f"wrote HTML: {out}")
        return 0

    print(f"# {xlsx_path.name}\n")
    print(f"sheets ({len(sheets)}): {sheets}\n")
    for s in sheets:
        try:
            print(_view_one(xlsx_path, s, args.rows))
        except Exception as e:
            print(f"### sheet: {s}\nERROR: {e}\n")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# 2. fig — pandas DF → matplotlib SVG/PNG/PDF
# ═══════════════════════════════════════════════════════════════════════

def _load_htmlfig_ramps() -> dict:
    """importlib 加载 htmlfig.py 拿配色 (失败 → 内置 mako)。"""
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("htmlfig", _HTMLFIG_LIB)
        if spec and spec.loader:
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)  # type: ignore
            return getattr(mod, "RAMPS", {})
    except Exception:
        pass
    return {}


def _palette(n: int, ramp_name: str = "mako") -> list[str]:
    """从 html-fig RAMPS 取 n 个等距色 (失败 → matplotlib viridis)。"""
    ramps = _load_htmlfig_ramps()
    if ramp_name not in ramps:
        ramp_name = next(iter(ramps), "")
    if not ramp_name:
        try:
            import matplotlib.cm as cm
            return [
                "#%02x%02x%02x" % tuple(int(x * 255) for x in cm.viridis(t)[:3])
                for t in [i / max(1, n - 1) for i in range(n)]
            ]
        except Exception:
            return ["#06c"] * n
    stops = ramps[ramp_name]

    def _ramp(t):
        for i in range(len(stops) - 1):
            p0, c0 = stops[i]
            p1, c1 = stops[i + 1]
            if p0 <= t <= p1:
                f = (t - p0) / (p1 - p0 or 1)
                return tuple(c0[k] + (c1[k] - c0[k]) * f for k in range(3))
        return stops[-1][1]

    cols = []
    for i in range(n):
        t = i / max(1, n - 1)
        c = _ramp(t)
        cols.append("#%02x%02x%02x" % tuple(int(round(v)) for v in c))
    return cols


def cmd_fig(args: argparse.Namespace) -> int:
    xlsx_path = Path(args.xlsx).resolve()
    if not xlsx_path.is_file():
        print(f"ERROR: file not found: {xlsx_path}", file=sys.stderr)
        return 2

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        print("FATAL: matplotlib not installed", file=sys.stderr)
        return 2

    df = _read_sheet(xlsx_path, args.sheet)

    if args.filter:
        try:
            df = df.query(args.filter)
        except Exception as e:
            print(f"ERROR: --filter '{args.filter}' 失败: {e}", file=sys.stderr)
            return 2

    if args.x not in df.columns:
        print(f"ERROR: --x '{args.x}' 不在列里。可用: {list(df.columns)}",
              file=sys.stderr)
        return 2

    y_cols = [c.strip() for c in args.y.split(",") if c.strip()]
    miss = [c for c in y_cols if c not in df.columns]
    if miss:
        print(f"ERROR: --y 缺失 {miss}。可用: {list(df.columns)}",
              file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    # 数据落 CSV 副产物 (用户可独立用)
    csv_path = out_dir / "data.csv"
    df[[args.x, *y_cols]].to_csv(csv_path, index=False)

    palette = _palette(len(y_cols))
    title = args.title or f"{xlsx_path.stem} · {args.sheet}"

    fig, ax = plt.subplots(figsize=(10, 6), dpi=100)
    kind = args.kind

    x_vals = df[args.x].tolist()
    try:
        if kind == "line":
            for i, col in enumerate(y_cols):
                ax.plot(x_vals, df[col], marker="o", color=palette[i], label=col)
            ax.legend()
        elif kind == "bar":
            import numpy as np
            n = len(y_cols)
            w = 0.8 / max(1, n)
            xpos = np.arange(len(x_vals))
            for i, col in enumerate(y_cols):
                ax.bar(xpos + i * w - 0.4 + w / 2, df[col],
                       width=w, color=palette[i], label=col)
            ax.set_xticks(xpos)
            ax.set_xticklabels(x_vals, rotation=45, ha="right")
            if n > 1:
                ax.legend()
        elif kind == "scatter":
            for i, col in enumerate(y_cols):
                ax.scatter(x_vals, df[col], color=palette[i], label=col,
                           s=40, alpha=0.7)
            ax.legend()
        elif kind == "pie":
            col = y_cols[0]
            ax.pie(df[col], labels=x_vals, colors=_palette(len(x_vals)),
                   autopct="%1.1f%%")
        elif kind == "heat":
            import numpy as np
            mat = df[y_cols].to_numpy()
            im = ax.imshow(mat, aspect="auto", cmap="viridis")
            ax.set_yticks(range(len(x_vals)))
            ax.set_yticklabels(x_vals)
            ax.set_xticks(range(len(y_cols)))
            ax.set_xticklabels(y_cols, rotation=45, ha="right")
            fig.colorbar(im, ax=ax)
        else:
            print(f"ERROR: 未知 --kind {kind}", file=sys.stderr)
            return 2
    except Exception as e:
        print(f"ERROR: 作图失败 ({kind}): {e}", file=sys.stderr)
        return 2

    ax.set_title(title)
    if kind not in ("pie", "heat"):
        ax.set_xlabel(args.x)
        if len(y_cols) == 1:
            ax.set_ylabel(y_cols[0])
    fig.tight_layout()

    base = out_dir / f"fig-{int(time.time())}"
    svg = base.with_suffix(".svg")
    png = base.with_suffix(".png")
    pdf = base.with_suffix(".pdf")
    fig.savefig(svg, format="svg")
    fig.savefig(png, format="png", dpi=200)
    fig.savefig(pdf, format="pdf")
    plt.close(fig)

    print(f"data:  {csv_path}")
    print(f"svg:   {svg}")
    print(f"png:   {png}")
    print(f"pdf:   {pdf}")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# 3. to-db
# ═══════════════════════════════════════════════════════════════════════

def _write_one(df: "pd.DataFrame", table: str, conn: sqlite3.Connection,
               if_exists: str) -> tuple[int, float]:
    t0 = time.time()
    df.to_sql(table, conn, if_exists=if_exists, index=False)
    dt = time.time() - t0
    return len(df), dt


def cmd_to_db(args: argparse.Namespace) -> int:
    xlsx_path = Path(args.xlsx).resolve()
    if not xlsx_path.is_file():
        print(f"ERROR: file not found: {xlsx_path}", file=sys.stderr)
        return 2

    db_path = Path(args.db).resolve()
    db_path.parent.mkdir(parents=True, exist_ok=True)

    if_exists = "replace" if args.reconcile else args.if_exists

    if args.sheet and not args.table:
        # 单 sheet 默认表名 = sanitize(sheet)
        args.table = _sanitize_table_name(args.sheet)
    if args.table and not args.sheet:
        print("ERROR: --table 须配合 --sheet 使用", file=sys.stderr)
        return 2

    xl = pd.ExcelFile(xlsx_path)
    sheets = [args.sheet] if args.sheet else xl.sheet_names

    conn = sqlite3.connect(db_path)
    try:
        report = []
        for s in sheets:
            try:
                df = _read_sheet(xlsx_path, s)
            except Exception as e:
                print(f"[{s}] read FAIL: {e}", file=sys.stderr)
                report.append((s, "?", 0, 0.0, f"read FAIL: {e}"))
                continue
            table = args.table if (args.sheet and args.table) \
                else _sanitize_table_name(s)
            try:
                n, dt = _write_one(df, table, conn, if_exists)
                report.append((s, table, n, dt, "ok"))
                print(f"[{s}] → {table}: {n} 行, {dt*1000:.0f} ms")
            except ValueError as e:
                # if_exists=fail 已存在
                print(f"[{s}] → {table}: SKIP ({e})", file=sys.stderr)
                report.append((s, table, 0, 0.0, f"skip: {e}"))
            except Exception as e:
                print(f"[{s}] → {table}: FAIL {e}", file=sys.stderr)
                report.append((s, table, 0, 0.0, f"FAIL: {e}"))
        conn.commit()
    finally:
        conn.close()

    total_rows = sum(r[2] for r in report)
    print(f"\n汇总: db={db_path}, sheets={len(sheets)}, rows_total={total_rows}")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# argparse
# ═══════════════════════════════════════════════════════════════════════

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlsx_cli",
        description="XLSX 处理统一 CLI v%s (3 子命令: view/fig/to-db)"
        % SCRIPT_VERSION,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {SCRIPT_VERSION}"
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    # view
    p_view = sub.add_parser(
        "view", help="探查 sheet/列/sample (markdown 或 HTML)"
    )
    p_view.add_argument("xlsx")
    p_view.add_argument("--sheet", help="只列单 sheet (默认全部)")
    p_view.add_argument("--rows", type=int, default=10, help="sample 行数")
    p_view.add_argument("--out", help="写 HTML 到此路径")
    p_view.set_defaults(func=cmd_view)

    # fig
    p_fig = sub.add_parser(
        "fig", help="DataFrame → SVG/PNG/PDF (matplotlib, 配色源自 html-fig)"
    )
    p_fig.add_argument("xlsx")
    p_fig.add_argument("--sheet", required=True)
    p_fig.add_argument("--x", required=True, help="横轴列")
    p_fig.add_argument("--y", required=True, help="纵轴列 (逗号分隔多列)")
    p_fig.add_argument("--filter", help="pandas df.query() 表达式")
    p_fig.add_argument(
        "--kind", default="bar",
        choices=["line", "bar", "heat", "pie", "scatter"]
    )
    p_fig.add_argument("--out-dir", default=".", help="产物输出目录")
    p_fig.add_argument("--title", default="")
    p_fig.set_defaults(func=cmd_fig)

    # to-db
    p_db = sub.add_parser(
        "to-db", help="pandas → sqlite3 入库 (--reconcile = 全量重建)"
    )
    p_db.add_argument("xlsx")
    p_db.add_argument("--db", required=True, help="sqlite db 路径")
    p_db.add_argument("--sheet", help="只此 sheet (默认全部 sheet → 各自表)")
    p_db.add_argument("--table", help="目标表名 (须配 --sheet)")
    p_db.add_argument(
        "--reconcile", action="store_true",
        help="语义同 if-exists replace (全量重建)"
    )
    p_db.add_argument(
        "--if-exists", default="fail",
        choices=["fail", "replace", "append"]
    )
    p_db.set_defaults(func=cmd_to_db)

    return parser


def main() -> int:
    parser = build_parser()
    if len(sys.argv) < 2:
        parser.print_help()
        return 1
    args = parser.parse_args()
    return args.func(args) or 0


if __name__ == "__main__":
    sys.exit(main())
