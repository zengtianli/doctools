#!/usr/bin/env python3
"""
文档/数据 统一调度器 —— 按文件后缀自动 route 到现有引擎(零重写,纯路由层)。

设计:命令只表达"动词",格式让本脚本运行时认。和 content-router 同一模式。
所有底层引擎(md_tools/pptx_tools/docx_*/data/convert 等)一个不改,subprocess 调用。

用法:
  doc_dispatch.py clean    <files...>                 规范化(docx 文本修复 / md format / pptx 全套)
  doc_dispatch.py convert  --to {md,word,xlsx,csv,txt} <files...>   转换(源自动认;老 .doc 经 textutil 升级)
  doc_dispatch.py typeset  <files...>                 md/docx/doc → 院模板成品 Word(套模板→修文本→图注)
  doc_dispatch.py merge    <files...>                 合并(md/txt→csv/xlsx)
  doc_dispatch.py split    <files...>                 拆分(md 按标题 / xlsx 按 sheet)
  doc_dispatch.py view     <files...>                 预览(md → HTML 浏览器)
  doc_dispatch.py scan     <dir>                      敏感词扫描(目录里 md/docx)
  doc_dispatch.py renum    [--to all|tabfig|headings] <files...>   序号修正(docx 标题/图/表编号重排,产出 _序号修正.docx)

# 实现：doc_dispatch <verb> @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

DOC = Path(__file__).resolve().parent            # scripts/document
DATA = DOC.parent / "data"                       # scripts/data
PY = sys.executable                              # uv 环境的 python

# pdf 系引擎（pdf_to_docx.py）依赖 pdfplumber/pypdf，**只装在 homebrew python3**，
# 不在 ~/Dev/.venv —— 而 GUI 经 `uv run --project ~/Dev` 调本文件时 sys.executable
# 正是 .venv/bin/python。用 PY 跑会 ModuleNotFoundError，必须写死绝对路径
# （同 pdf_cli.py 的既定运行解释器）。
PY_PDF = "/opt/homebrew/bin/python3"

# 同理:GUI .app 不继承 shell PATH,spawn 的 CLI 一律解析成绝对路径再调
# (markitdown 是 uv tool,装在 ~/.local/bin,GUI 下裸名必 not found)。
_MARKITDOWN = shutil.which("markitdown") or str(Path.home() / ".local/bin/markitdown")
_PDFTOTEXT = shutil.which("pdftotext") or "/opt/homebrew/bin/pdftotext"

# 兜底 PYTHONPATH:有些引擎(如 md_docx_template.py)只加了 doctools/lib、漏了 dev/lib，
# 直接子进程调用会 ModuleNotFoundError(file_ops 等)。这里统一补齐,覆盖所有引擎。
_LIBS = [
    str(DOC.parent.parent / "lib"),                              # doctools/lib
    str(Path.home() / "Dev/tools/dev/lib"),                      # dev/lib (file_ops 等 canonical)
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


def _doc_to_docx(f: str) -> str | None:
    """老 .doc → .docx。优先 soffice(产完整 docx 含 styles.xml,下游套模板等引擎都吃),
    没装 LibreOffice 才 textutil 兜底(极简 docx,缺 styles.xml,套模板线会挂)。
    成功返回产出路径,失败返回 None。"""
    p = Path(f)
    out = p.with_suffix(".docx")
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(soffice).exists():
        cmd = [soffice, "--headless",
               "-env:UserInstallation=file:///tmp/lo_profile_doc_dispatch",
               "--convert-to", "docx", "--outdir", str(p.parent), str(p)]
        if _run(cmd, "老 doc → docx(soffice)") == 0 and out.exists():
            return str(out)
    if _run(["textutil", "-convert", "docx", str(p), "-output", str(out)], "老 doc → docx(textutil 兜底)"):
        return None
    return str(out) if out.exists() else None


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
    # pdf 系:结构提取路线(pdfplumber+python-docx),必须用 homebrew python3
    if e == "pdf":
        if target == "word":
            return [PY_PDF, str(DOC / "pdf_to_docx.py"), f], "PDF → Word(段落重组 + 真表格)"
        if target == "md":
            return [_MARKITDOWN, f, "-o", str(Path(f).with_suffix(".md"))], "PDF → Markdown"
        if target == "txt":
            return [_PDFTOTEXT, "-layout", f], "PDF → 纯文本(保排版)"
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


def _word_textfix(out: Path) -> int:
    """convert→word 收尾:自动文本修复(引号/标点/单位),成品就地替换,不留中间文件。
    用户钦定:转出来的 docx 直接就是 text_format 好的,不用再手动跑一遍规范化。"""
    if not out.exists():
        warn(f"未找到转换产出 {out.name},跳过文本修复")
        return 1
    if _run(_py("docx_text_formatter.py", str(out)), "文本修复(引号/标点/单位)"):
        return 1
    fixed = out.with_name(f"{out.stem}_fixed{out.suffix}")
    if not fixed.exists():
        warn(f"文本修复未产出 {fixed.name}")
        return 1
    fixed.replace(out)
    print(f"{GREEN}  ✓ 成品(已文本修复) → {out.name}{RST}")
    return 0


def _word_output(f: str) -> Path:
    """convert→word 各引擎的默认产出路径。"""
    p = Path(f)
    if _ext(f) == "docx":
        return p.with_name(f"{p.stem}_styled.docx")   # docx_apply_template.py
    return p.with_suffix(".docx")                     # md_docx_template.py / soffice


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
        # 老 .doc:textutil 直转(word/txt),或先升级成 docx 再走 docx 路由(md)
        if _ext(f) == "doc":
            if target == "txt":
                print(f"{GREEN}● {Path(f).name}{RST}")
                rc |= _run(["textutil", "-convert", "txt", f], "老 doc → txt(textutil)")
                continue
            if target in ("word", "md"):
                print(f"{GREEN}● {Path(f).name}{RST}")
                nf = _doc_to_docx(f)
                if nf is None:
                    rc |= 1; continue
                if target == "word":
                    rc |= _word_textfix(Path(nf))
                    continue
                f = nf  # md:继续按 docx → md 路由
            else:
                warn(f"{Path(f).name}: doc → {target} 这条转换没引擎(不支持的组合),跳过")
                continue
            hit = route_convert(f, target)
            if hit:
                rc |= _run(hit[0], hit[1] or f"→ {target}")
            continue
        hit = route_convert(f, target)
        if not hit:
            warn(f"{Path(f).name}: {_ext(f) or '?'} → {target} 这条转换没引擎(不支持的组合),跳过")
            continue
        cmd, label = hit
        print(f"{GREEN}● {Path(f).name}{RST}")
        r = _run(cmd, label or f"→ {target}")
        rc |= r
        if r == 0 and target == "word":
            rc |= _word_textfix(_word_output(f))
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
    pr = sub.add_parser("renum")
    pr.add_argument("--to", default="all", dest="target", choices=["all", "tabfig", "headings"])
    pr.add_argument("files", nargs="+")
    a = ap.parse_args()

    if a.verb == "clean":   return do_clean(a.files)
    if a.verb == "split":   return do_split(a.files)
    if a.verb == "view":    return do_view(a.files)
    if a.verb == "scan":    return do_scan(a.files)
    if a.verb == "merge":   return do_merge(a.files)
    if a.verb == "convert": return do_convert(a.files, a.target)
    if a.verb == "typeset": return do_typeset(a.files)
    if a.verb == "renum":   return do_renum(a.files, a.target)
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
        if e not in ("md", "docx", "doc"):
            warn(f"{p.name}: 套模板出 Word 只吃 md/docx/doc,跳过"); continue
        print(f"{GREEN}● {p.name}{RST}")
        if e == "doc":  # 老 doc 先升级成 docx 再走 docx 线
            nf = _doc_to_docx(str(p))
            if nf is None:
                rc |= 1; continue
            p, e = Path(nf), "docx"
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


# ───────────────────────────────────────────── renum(docx 序号修正:标题/图/表)

def _renum_has_en_figures(docx: str) -> bool:
    """document.xml 里有英文 Figure N 引用、且没有中文 图/表 X-Y 题注 → 走英文模式。"""
    import re as _re
    import zipfile as _zip
    try:
        xml = _zip.ZipFile(docx).read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return False
    cn = _re.search(r"[图表]\s*\d+(?:\.\d+)?\s*[-－—–]\s*\d+", xml)
    en = _re.search(r"(?:Figure|Fig\.?)\s*\d+", xml)
    return bool(en and not cn)


def do_renum(files, target: str = "all") -> int:
    """序号修正:docx 的标题号/图号/表号 断号·错号·缺号 一键重排,产出 <stem>_序号修正.docx(原件不动)。

    纯路由,组合单一职责引擎(引擎零改):
      标题号 = sub/renumber_headings.py       (按段落物理顺序重编 H1/H2/H3)
      图号   = docx_renumber_figures.py --cn-section --kind 图 (节内重排+补号+正文引用同步)
      表号   = 同上 --kind 表
      英文稿 = 无中文题注但有 Figure N → docx_renumber_figures.py 英文模式(1..N+引用同步)
    """
    rc = 0
    for f in files:
        p = Path(f)
        if not p.exists():
            warn(f"文件不存在,跳过: {f}"); continue
        if _ext(f) != "docx":
            warn(f"{p.name}: 序号修正只吃 docx,跳过"); continue
        print(f"{GREEN}● {p.name}{RST}")
        out = p.with_name(f"{p.stem}_序号修正.docx")
        shutil.copy2(p, out)
        failed = False
        if target in ("all", "headings"):
            if _run(_py("sub/renumber_headings_seq.py", str(out), "--no-backup"), "标题号重排(按现有深度)"):
                failed = True
        if not failed and target in ("all", "tabfig"):
            if _renum_has_en_figures(str(out)):
                if _run(_py("docx_renumber_figures.py", str(out), "--inplace"), "Figure 号重排(英文)"):
                    failed = True
            else:
                for kind in ("图", "表"):
                    if _run(_py("docx_renumber_figures.py", str(out),
                                "--cn-section", "--kind", kind, "--inplace"), f"{kind}号重排"):
                        failed = True; break
        bak = Path(str(out) + ".bak")
        if bak.exists():
            bak.unlink()  # 工作副本的中间备份,原件本身就是备份
        if failed:
            warn(f"{p.name}: 序号修正未全绿,产出保留供检查 → {out.name}")
            rc |= 1
        else:
            print(f"{GREEN}  ✓ 序号修正 → {out.name}{RST}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
