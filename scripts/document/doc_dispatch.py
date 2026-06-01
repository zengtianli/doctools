#!/usr/bin/env python3
"""
文档/数据 统一调度器 —— 按文件后缀自动 route 到现有引擎(零重写,纯路由层)。

设计:命令只表达"动词",格式让本脚本运行时认。和 content-router 同一模式。
所有底层引擎(md_tools/pptx_tools/docx_*/data/convert 等)一个不改,subprocess 调用。

用法:
  doc_dispatch.py clean    <files...>                 规范化(docx 文本修复 / md format / pptx 全套)
  doc_dispatch.py convert  --to {md,word,xlsx,csv,txt} <files...>   转换(源自动认)
  doc_dispatch.py typeset  <files...>                 md/docx → 院模板成品 Word(套模板→修文本→图注)
  doc_dispatch.py merge    <files...>                 合并(md/txt→csv/xlsx)
  doc_dispatch.py split    <files...>                 拆分(md 按标题 / xlsx 按 sheet)
  doc_dispatch.py view     <files...>                 预览(md → HTML 浏览器)
  doc_dispatch.py scan     <dir>                      敏感词扫描(目录里 md/docx)

# 实现：doc_dispatch <verb> @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

DOC = Path(__file__).resolve().parent            # scripts/document
DATA = DOC.parent / "data"                       # scripts/data
PY = sys.executable                              # uv 环境的 python

# 兜底 PYTHONPATH:有些引擎(如 md_docx_template.py)只加了 doctools/lib、漏了 dev/lib，
# 直接子进程调用会 ModuleNotFoundError(file_ops 等)。这里统一补齐,覆盖所有引擎。
_LIBS = [
    str(DOC.parent.parent / "lib"),                              # doctools/lib
    str(Path.home() / "Dev/tools/dev/lib"),                      # dev/lib (file_ops 等 canonical)
    str(Path.home() / "Dev/stations/dockit/src"),               # dockit text utils
]
_ENV = {**os.environ, "PYTHONPATH": os.pathsep.join(_LIBS + [os.environ.get("PYTHONPATH", "")])}

GREEN, YELLOW, RED, DIM, RST = "\033[32m", "\033[33m", "\033[31m", "\033[2m", "\033[0m"


def _ext(f: str) -> str:
    return Path(f).suffix.lower().lstrip(".")


def _run(cmd: list[str], label: str) -> int:
    print(f"{DIM}  ↳ {label}{RST}")
    return subprocess.run(cmd, env=_ENV).returncode


def _py(engine: str, *args: str) -> list[str]:
    return [PY, str(DOC / engine), *args]


def _data(engine: str, *args: str) -> list[str]:
    return [PY, str(DATA / engine), *args]


def warn(msg: str) -> None:
    print(f"{YELLOW}  ⚠ {msg}{RST}")


# ───────────────────────────────────────────── 路由表

def route_clean(f: str) -> tuple[list[str], str] | None:
    e = _ext(f)
    if e == "docx":
        return _py("docx_text_formatter.py", f), "docx → 文本修复(引号/标点/单位)"
    if e == "md":
        return _py("md_tools.py", "format", f), "md → 格式标准化"
    if e in ("pptx",):
        return _py("pptx_tools.py", "all", f), "pptx → 字体+表格+文本 全套规范"
    if e in ("xlsx", "xlsm"):
        return _data("xlsx_lowercase.py", f), "xlsx → 英文小写整理"
    return None


def route_convert(f: str, target: str) -> tuple[list[str], str] | None:
    e = _ext(f)
    M = {
        ("docx", "md"): (_py("docx_to_md_sh", f), None),  # 占位,下面特判 .sh
        ("pptx", "md"): (_py("pptx_to_md.py", f), "pptx → Markdown"),
        ("ppt", "md"): (_py("pptx_to_md.py", f), "ppt → Markdown"),
        ("md", "word"): (_py("md_docx_template.py", f), "md → Word(套模板)"),
        ("docx", "word"): (_py("docx_apply_template.py", f), "docx → 套模板重排"),
        ("csv", "xlsx"): (_data("convert.py", "xlsx-from-csv", f), "csv → Excel"),
        ("txt", "xlsx"): (_data("convert.py", "xlsx-from-txt", f), "txt → Excel"),
        ("xls", "xlsx"): (_data("convert.py", "xlsx-from-xls", f), "老 xls → xlsx"),
        ("txt", "csv"): (_data("convert.py", "csv-from-txt", f), "txt → CSV"),
        ("xlsx", "csv"): (_data("convert.py", "xlsx-to-csv", f), "Excel → CSV"),
        ("csv", "txt"): (_data("convert.py", "csv-to-txt", f), "CSV → txt"),
        ("xlsx", "txt"): (_data("convert.py", "xlsx-to-txt", f), "Excel → txt"),
    }
    # docx→md 走 shell 引擎,特判
    if e == "docx" and target == "md":
        return ["bash", str(DOC / "docx_to_md.sh"), f], "docx → Markdown(markitdown)"
    hit = M.get((e, target))
    return (hit[0], hit[1]) if hit else None


def route_merge(f: str) -> tuple[list[str], str] | None:
    e = _ext(f)
    if e == "md":
        return None, "md"     # md 走批量(下面特判:一次传所有)
    if e == "txt":
        return None, "txt"
    if e in ("xlsx", "xlsm"):
        return None, "xlsx"
    return None


def route_split(f: str) -> tuple[list[str], str] | None:
    e = _ext(f)
    if e == "md":
        return _py("md_tools.py", "split", f), "md → 按标题拆分"
    if e in ("xlsx", "xlsm"):
        return _data("xlsx_splitsheets.py", f), "xlsx → 按 sheet 拆分"
    return None


# ───────────────────────────────────────────── 动词实现

def _per_file(files: list[str], router, verb: str) -> int:
    rc = 0
    for f in files:
        if not Path(f).exists():
            warn(f"文件不存在,跳过: {f}")
            continue
        hit = router(f)
        if not hit:
            warn(f"{Path(f).name}: 没有「{verb}」对应的 {_ext(f) or '?'} 引擎,跳过")
            continue
        cmd, label = hit
        print(f"{GREEN}● {Path(f).name}{RST}")
        rc |= _run(cmd, label)
    return rc


def do_clean(files):  return _per_file(files, route_clean, "规范化")
def do_split(files):  return _per_file(files, route_split, "拆分")


def do_convert(files, target):
    aliases = {"markdown": "md", "docx": "word", "excel": "xlsx", "text": "txt"}
    target = aliases.get(target, target)
    if target not in ("md", "word", "xlsx", "csv", "txt"):
        print(f"{RED}✖ 未知目标格式: {target}(支持 md/word/xlsx/csv/txt){RST}")
        return 2
    rc = 0
    for f in files:
        if not Path(f).exists():
            warn(f"文件不存在,跳过: {f}"); continue
        hit = route_convert(f, target)
        if not hit:
            warn(f"{Path(f).name}: {_ext(f) or '?'} → {target} 这条转换没引擎(不支持的组合),跳过")
            continue
        cmd, label = hit
        print(f"{GREEN}● {Path(f).name}{RST}")
        rc |= _run(cmd, label or f"→ {target}")
    return rc


def do_view(files):
    rc = 0
    for f in files:
        if _ext(f) != "md":
            warn(f"{Path(f).name}: 预览目前只支持 md,跳过"); continue
        print(f"{GREEN}● {Path(f).name}{RST}")
        rc |= _run(_py("md_tools.py", "to-html", f), "md → HTML 浏览器预览")
    return rc


def do_scan(target):
    d = target[0] if isinstance(target, list) else target
    print(f"{GREEN}● 扫描 {d}{RST}")
    return _run(_py("scan_sensitive_words.py", d), "敏感词扫描(竞品名/过硬措辞)")


def do_merge(files):
    """合并按类型分组:md→md_tools merge;txt→csv-merge-txt;xlsx→xlsx_merge_tables(需主/辅表参数)。"""
    groups: dict[str, list[str]] = {}
    for f in files:
        groups.setdefault(_ext(f), []).append(f)
    rc = 0
    for e, fs in groups.items():
        if e == "md":
            print(f"{GREEN}● 合并 {len(fs)} 个 md{RST}")
            rc |= _run(_py("md_tools.py", "merge", *fs), "md → 合并为一篇")
        elif e == "txt":
            print(f"{GREEN}● 合并 {len(fs)} 个 txt → CSV{RST}")
            rc |= _run(_data("convert.py", "csv-merge-txt", *fs), "txt 按列 → CSV")
        elif e in ("xlsx", "xlsm"):
            warn(f"xlsx 合并需指定主表/辅表/映射,非纯多选;请直接用引擎: "
                 f"python3 {DATA/'xlsx_merge_tables.py'} --master ... --aux ... --map ...")
        else:
            warn(f".{e} 没有合并引擎,跳过({len(fs)} 个)")
    return rc


# ───────────────────────────────────────────── CLI

def main() -> int:
    ap = argparse.ArgumentParser(description="文档/数据统一调度器")
    sub = ap.add_subparsers(dest="verb", required=True)
    for v in ("clean", "typeset", "merge", "split", "view", "scan"):
        p = sub.add_parser(v); p.add_argument("files", nargs="+")
    pc = sub.add_parser("convert")
    pc.add_argument("--to", required=True, dest="target")
    pc.add_argument("files", nargs="+")
    a = ap.parse_args()

    if a.verb == "clean":   return do_clean(a.files)
    if a.verb == "split":   return do_split(a.files)
    if a.verb == "view":    return do_view(a.files)
    if a.verb == "scan":    return do_scan(a.files)
    if a.verb == "merge":   return do_merge(a.files)
    if a.verb == "convert": return do_convert(a.files, a.target)
    if a.verb == "typeset": return do_typeset(a.files)
    return 1


# ───────────────────────────────────────────── typeset(md/docx → 院模板成品 Word)

def do_typeset(files) -> int:
    """复刻 md2word_pipeline:md→套模板 / docx→套模板,再文本修复,再图注居中,清中间文件。"""
    rc = 0
    for f in files:
        p = Path(f)
        if not p.exists():
            warn(f"文件不存在,跳过: {f}"); continue
        e = _ext(f)
        if e not in ("md", "docx"):
            warn(f"{p.name}: 套模板出 Word 只吃 md/docx,跳过"); continue
        print(f"{GREEN}● {p.name}{RST}")
        d, stem = p.parent, p.stem
        # Step 1: 转换/套模板
        if e == "md":
            if _run(_py("md_docx_template.py", str(p)), "1/3 md → Word(套模板)"): rc |= 1; continue
            step1 = d / f"{stem}.docx"; final = d / f"{stem}.docx"
        else:
            if _run(_py("docx_apply_template.py", str(p)), "1/3 docx 套模板重排"): rc |= 1; continue
            step1 = d / f"{stem}_styled.docx"; final = d / f"{stem}_styled.docx"
        if not step1.exists():
            warn(f"Step1 未产出 {step1.name},中止该文件"); rc |= 1; continue
        # Step 2: 文本修复 → <name>_fixed.docx
        if _run(_py("docx_text_formatter.py", str(step1)), "2/3 文本修复"): rc |= 1; continue
        step2 = step1.with_name(step1.stem + "_fixed.docx")
        if not step2.exists():
            warn(f"Step2 未产出 {step2.name},中止该文件"); rc |= 1; continue
        # Step 3: 图注居中(就地)
        _run(_py("docx_apply_image_caption.py", str(step2)), "3/3 图注样式")
        # 收尾:step2 → final,清中间
        step2.replace(final)
        if step1 != final and step1.exists():
            step1.unlink()
        bak = step2.with_name(step2.name + ".backup")
        if bak.exists():
            bak.unlink()
        print(f"{GREEN}  ✓ 成品 → {final.name}{RST}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
