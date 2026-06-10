#!/usr/bin/env python3
# distilled from eco-flow/taizhou-天台 govern 2026-06-08:
# "以前是 章节 组合 生成 成品、现在是 成品校核修改后 重新 反向 更新 章节。
#  得做 2 个 slash command 互相转换。以后的报告都可以复用"
r"""chapters_sync.py — 成品 docx 校核修改后, 反向回写/重生「成品章节」目录。

与 merge/combine (章节→成品, 正向) 互为逆操作:
    merge / combine:  章节 docx × N  ──组合──▶  成品 docx
    chapters-sync:    成品 docx      ──反向──▶  章节目录 (重生每章 docx)

动作 (execute 时):
    1. 备份: 现有章节目录里将被替换的章 docx / 已失效(stale)的章目录
       → mv 到 ``~/.Trash/<proj>-chapters-sync-<ts>/`` (禁 rm, 全程可回溯)
    2. 切分: 按 H1 切分成品 (复用 split_by_h1.plan_slices / write_slice)
    3. 媒体去冗余: 每章内置 orphan-media deep 清理 (write_slice prune_media=True,
       复用 strip_orphan_media) — 否则整本 27MB → 每章 27MB (历史踩坑, 硬需求)
    4. 语义命名: ``{章序:02d}-{章标题}`` (去标题前导阿拉伯编号 + 去空白),
       禁 idx 裸命名 — 与 02-工作区 折叠式章节目录 (NN-章名/NN-章名.docx) 对齐
    5. 章目录内非 docx 的工作材料 (CLAUDE.md / attachments / figs / xlsx)
       原地保留不动 — sync 只换章 docx 本体

默认 dry-run (打印计划不动盘), ``--execute`` 才真写 (work_ops 二段式风格)。

章节目录自动探测 (不传 --chapters-dir 时, 按序):
    1. <docx 同目录>/成品章节
    2. <docx 同目录>/<docx-stem>-split
    3. <docx 同目录>/../02-工作区   (须含 ≥1 个折叠式章节 NN-*/NN-*.docx)
    4. <docx 同目录>/../成品章节
    探测全失败 → 报错, 必须显式 --chapters-dir。

CLI (经 docx_cli 或独立):
    python3 scripts/document/docx_cli.py chapters-sync <成品.docx> \
        [--chapters-dir <dir>] [--execute] [--flat] \
        [--include-frontmatter] [--keep-all-media] [--report <json>]
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional

# 复用 split_by_h1 的切分引擎 (plan_slices / write_slice / sanitize_filename)。
# 三态 import 兼容: 包内(pipeline) / 脚本(sys.path[0]=sub/) / docx_cli runpy。
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
try:
    from . import split_by_h1 as _split  # type: ignore
except ImportError:
    import split_by_h1 as _split  # type: ignore


# 章节实体命名: NN-章名 (两位序号 + 连字符)
_CHAPTER_ENTRY_RE = re.compile(r"^\d{2}-")
# 标题前导阿拉伯章号: "1  编制目的与依据" / "2.河湖…" / "3、…"
_LEADING_NUM_RE = re.compile(r"^\d+[\s.、．]*")
_WS_RE = re.compile(r"\s+")


def clean_title(title: str) -> str:
    """章标题 → 语义文件名段: 去前导阿拉伯编号 + 去全部空白 + 去非法字符。

    '1  编制目的与依据' → '编制目的与依据' ; '前  言' → '前言'。
    """
    t = _LEADING_NUM_RE.sub("", title or "")
    t = _WS_RE.sub("", t)
    t = _split.sanitize_filename(t)
    return t or "untitled"


def detect_chapters_dir(docx: Path) -> Optional[Path]:
    """按约定自动探测「成品章节」目录; 找不到返回 None (调用方必须显式传参)。"""
    d = docx.parent
    candidates = [
        d / "成品章节",
        d / f"{docx.stem}-split",
        d.parent / "02-工作区",
        d.parent / "成品章节",
    ]
    for c in candidates:
        if not c.is_dir():
            continue
        if c.name == "02-工作区":
            # 须真含折叠式章节 NN-*/NN-*.docx 才认 (避免误伤别的工作区)
            folded = [
                e for e in c.iterdir()
                if e.is_dir() and _CHAPTER_ENTRY_RE.match(e.name)
                and (e / f"{e.name}.docx").exists()
            ]
            if not folded:
                continue
        return c.resolve()
    return None


def build_plan(
    docx: Path,
    chapters_dir: Path,
    include_frontmatter: bool = False,
    flat: bool = False,
) -> dict:
    """只读规划: 返回 {slices, targets, backups, stale, h1_count}。"""
    slices, _sect_idx, h1_count = _split.plan_slices(docx, include_frontmatter)
    if h1_count == 0:
        return {"error": "0 Heading-1 detected (docx unhealthy); run /docx health first",
                "exit_code": 3, "h1_count": 0}

    # 语义命名: frontmatter → 00-frontmatter; 第 k 章 (1-based) → {k:02d}-{clean_title}
    plan_slices = []
    ord_counter = 0
    for s in slices:
        if s["is_frontmatter"]:
            name = "00-frontmatter"
        else:
            ord_counter += 1
            name = f"{ord_counter:02d}-{clean_title(s['title'])}"
        if flat:
            target = chapters_dir / f"{name}.docx"
        else:
            target = chapters_dir / name / f"{name}.docx"
        plan_slices.append({
            "name": name,
            "title": s["title"],
            "start": s["start"],
            "end": s["end"],
            "target": str(target),
        })

    new_names = {p["name"] for p in plan_slices}

    # 备份计划: 只动章节实体 (^\d{2}- 前缀), 其余 (bin/figures/_audit/CLAUDE.md…) 不碰
    backups: list[dict] = []   # 被新章替换的旧 docx (章目录里其余工作材料原地保留)
    stale: list[dict] = []     # 不再对应任何新章的旧章实体 → 整体挪备份
    if chapters_dir.is_dir():
        for entry in sorted(chapters_dir.iterdir()):
            if not _CHAPTER_ENTRY_RE.match(entry.name):
                continue
            if entry.is_dir():
                if entry.name in new_names:
                    old_docx = entry / f"{entry.name}.docx"
                    if old_docx.exists():
                        backups.append({"src": str(old_docx),
                                        "rel": f"{entry.name}/{old_docx.name}"})
                else:
                    stale.append({"src": str(entry), "rel": entry.name})
            elif entry.suffix == ".docx":
                # 平铺旧章 docx: 同名被替换 / 不同名失效, 都挪备份
                backups.append({"src": str(entry), "rel": entry.name})

    return {
        "h1_count": h1_count,
        "docx": str(docx),
        "chapters_dir": str(chapters_dir),
        "layout": "flat" if flat else "folded",
        "slices": plan_slices,
        "backups": backups,
        "stale": stale,
    }


def run_sync(
    docx: Path,
    chapters_dir: Optional[Path] = None,
    execute: bool = False,
    flat: bool = False,
    include_frontmatter: bool = False,
    prune_media: bool = True,
) -> dict:
    """主入口: dry-run 默认只出计划; execute=True 才备份 + 切分 + 回写。"""
    src = Path(docx).expanduser().resolve()
    if not src.is_file():
        return {"error": f"input docx not found: {src}", "exit_code": 2}

    if chapters_dir is None:
        detected = detect_chapters_dir(src)
        if detected is None:
            return {
                "error": (
                    "无法自动探测「成品章节」目录 (试过: <docx目录>/成品章节, "
                    f"<docx目录>/{src.stem}-split, ../02-工作区, ../成品章节) — "
                    "必须显式传 --chapters-dir"
                ),
                "exit_code": 2,
            }
        chapters_dir = detected
    chapters_dir = Path(chapters_dir).expanduser().resolve()

    plan = build_plan(src, chapters_dir, include_frontmatter, flat)
    if plan.get("error"):
        return plan

    if not execute:
        plan["dry_run"] = True
        return plan

    # ① 备份 → ~/.Trash/<proj>-chapters-sync-<ts>/ (mv, 禁 rm)
    ts = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    proj = chapters_dir.parent.name or chapters_dir.name
    backup_root = Path.home() / ".Trash" / f"{proj}-chapters-sync-{ts}"
    moved = []
    for item in plan["backups"] + plan["stale"]:
        src_p = Path(item["src"])
        if not src_p.exists():
            continue
        dst_p = backup_root / item["rel"]
        dst_p.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src_p), str(dst_p))
        moved.append(item["rel"])

    # ② 切分 + ③ 每章 orphan-media 清理 (write_slice 内置 prune_media)
    chapters_dir.mkdir(parents=True, exist_ok=True)
    emitted, failed = [], []
    for s in plan["slices"]:
        dst = Path(s["target"])
        try:
            paras, nbytes = _split.write_slice(
                src, dst, s["start"], s["end"], prune_media=prune_media)
            emitted.append({"name": s["name"], "target": s["target"],
                            "paragraphs": paras, "bytes": nbytes})
        except Exception as e:
            failed.append({"name": s["name"], "target": s["target"],
                           "error": f"{type(e).__name__}: {e}"})

    # ④ 清 quarantine (best-effort)
    try:
        subprocess.run(
            ["xattr", "-dr", "com.apple.quarantine", str(chapters_dir)],
            capture_output=True, timeout=30)
    except Exception:
        pass

    plan.update({
        "dry_run": False,
        "backup_root": str(backup_root) if moved else None,
        "backed_up": moved,
        "emitted": emitted,
        "failed": failed,
        "exit_code": 0 if not failed else 1,
    })
    return plan


def _print_report(rep: dict, execute: bool) -> None:
    if rep.get("error"):
        print(f"ERROR: {rep['error']}", file=sys.stderr)
        return
    mode = "EXECUTE" if execute else "DRY-RUN (加 --execute 才动盘)"
    print(f"[chapters-sync] {mode}")
    print(f"  成品:     {rep['docx']}")
    print(f"  章节目录: {rep['chapters_dir']}  (layout={rep['layout']})")
    print(f"  H1 章数:  {rep['h1_count']}")
    print(f"  切分计划 ({len(rep['slices'])} 片):")
    for s in rep["slices"]:
        print(f"    {s['name']}  ← children[{s['start']}:{s['end']}]  (title={s['title']!r})")
        print(f"      → {s['target']}")
    if rep.get("backups"):
        print(f"  将备份旧章 docx ({len(rep['backups'])}):")
        for b in rep["backups"]:
            print(f"    {b['rel']}")
    if rep.get("stale"):
        print(f"  失效章实体整体挪备份 ({len(rep['stale'])}):")
        for b in rep["stale"]:
            print(f"    {b['rel']}/")
    if execute:
        if rep.get("backup_root"):
            print(f"  备份目录: {rep['backup_root']}  ({len(rep['backed_up'])} 项, mv 非 rm)")
        print(f"  已写 {len(rep['emitted'])} 章:")
        for e in rep["emitted"]:
            print(f"    {e['name']}  · {e['paragraphs']} paragraphs · {e['bytes']:,} bytes")
        for f in rep.get("failed", []):
            print(f"    FAILED {f['name']}: {f['error']}", file=sys.stderr)


def _run(args) -> int:
    rep = run_sync(
        docx=Path(args.docx),
        chapters_dir=Path(args.chapters_dir) if args.chapters_dir else None,
        execute=args.execute,
        flat=args.flat,
        include_frontmatter=args.include_frontmatter,
        prune_media=not args.keep_all_media,
    )
    _print_report(rep, args.execute)
    if getattr(args, "report", None):
        rp = Path(args.report).expanduser()
        rp.parent.mkdir(parents=True, exist_ok=True)
        rp.write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  report → {rp}")
    return int(rep.get("exit_code", 0))


def _add_args(p) -> None:
    p.add_argument("docx", help="成品 docx 路径 (不会被修改, 只读)")
    p.add_argument("--chapters-dir", default=None,
                   help="章节目录 (默认自动探测: 成品章节 / <stem>-split / ../02-工作区)")
    p.add_argument("--execute", action="store_true",
                   help="真执行 (默认 dry-run 只打印计划)")
    p.add_argument("--flat", action="store_true",
                   help="平铺产出 NN-章名.docx (默认折叠式 NN-章名/NN-章名.docx)")
    p.add_argument("--include-frontmatter", action="store_true",
                   help="首个 H1 前内容输出为 00-frontmatter (默认丢弃)")
    p.add_argument("--keep-all-media", action="store_true",
                   help="禁用每章 orphan-media 清理 (默认开; 关了每章扛整本媒体, 反模式)")
    p.add_argument("--report", default=None, help="JSON 报告输出路径")
    p.set_defaults(func=_run)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "chapters-sync",
        help="成品 docx 反向回写「成品章节」目录 (merge 的逆操作; "
             "备份→Trash + 按 H1 切分 + 每章 orphan-media 清理 + 语义命名)",
    )
    _add_args(p)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="成品 docx 校核修改后, 反向重生「成品章节」目录 (merge 的逆操作)",
    )
    _add_args(ap)
    args = ap.parse_args()
    return _run(args)


if __name__ == "__main__":
    sys.exit(main())
