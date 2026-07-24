#!/usr/bin/env python3
"""
DocTools GUI 后端适配器 —— 给 doc_dispatch.py 套一层 JSON 信封,供 SwiftUI app 调用。

设计(Tier B · 零重写业务):
  · 复用 doc_dispatch 的 route_*/do_* 路由逻辑,一行不改原引擎。
  · 把它们「打文本到 stdout + 产出落盘」的终端 UX,翻译成纯 JSON stdout 信封。
  · 产出路径靠「跑前/跑后扫目录树」差分得到(确定性,不解析中文日志)。

信封协议(同 ssot-console/robinhood-desk):所有 gui-* 一律 exit 0;
  成功 {"ok": true, ...},失败 {"ok": false, "error": "人话"};snake_case。

子命令:
  gui-ops                        列出可用操作(给 UI 渲菜单)
  gui-run --op <verb> [--to T] --files <paths...>   跑一个操作,返回逐文件结果

原有终端用法(doc_dispatch.py <verb>)完全不受影响 —— 本文件只 import,不改它。
"""
from __future__ import annotations

import argparse
import contextlib
import io
import json
import re
import subprocess
import sys
from pathlib import Path

# 终端日志带 ANSI 色码(\x1b[..m),是裸控制字符;塞进 JSON 字符串会让解码器报
# "Invalid control character"。embed 前一律剥掉。
_ANSI = re.compile(r"\x1b\[[0-9;]*m")


def _strip_ansi(s: str) -> str:
    return _ANSI.sub("", s)

# 同目录引擎,直接 import 复用路由(零重写)。
import doc_dispatch as dd

# ───────────────────────────────────────────── 子进程输出收口
#
# doc_dispatch._run 用 subprocess.run(cmd, env=_ENV),子进程继承真 stdout fd,
# contextlib.redirect_stdout 拦不住(那只换 Python 的 sys.stdout 对象,不动 fd 1)。
# → 子进程的终端日志会糊进我们的 JSON stdout,Swift 解码必崩。
# 解 = monkeypatch dd._run,把子进程 stdout/stderr 收进 PIPE,攒进 _CAP 缓冲。
# 零改原引擎:只在本适配器进程里替换 dd 模块上的 _run 引用。

_CAP: list[str] = []


def _captured_run(cmd, label):
    """dd._run 的捕获版:子进程输出进 PIPE 不外泄,攒进 _CAP。返回 returncode(签名同原版)。"""
    _CAP.append(f"  ↳ {label}")
    proc = subprocess.run(cmd, env=dd._ENV, capture_output=True, text=True)
    if proc.stdout:
        _CAP.append(proc.stdout.rstrip())
    if proc.stderr:
        _CAP.append(proc.stderr.rstrip())
    return proc.returncode


dd._run = _captured_run  # noqa: SLF001 — 故意替换,把终端日志关进缓冲

# ───────────────────────────────────────────── 操作目录(SSOT,UI 从这渲菜单)

# 每个 op:verb=doc_dispatch 动词;exts=支持的源后缀(给 UI 做拖入校验提示);
# targets=convert 专用的目标格式列表;scan=是否吃目录(而非文件)。
OPS = [
    {
        "id": "clean",
        "verb": "clean",
        "title": "规范化",
        "subtitle": "引号/标点/单位/字体修复(docx·md·pptx·xlsx)",
        "icon": "wand.and.stars",
        "exts": ["docx", "md", "pptx", "xlsx", "xlsm"],
        "kind": "files",
    },
    {
        "id": "convert",
        "verb": "convert",
        "title": "格式转换",
        "subtitle": "源格式自动识别 → 目标格式（含 PDF → 可编辑 Word）",
        "icon": "arrow.triangle.2.circlepath",
        "exts": ["pdf", "docx", "doc", "pptx", "ppt", "md", "csv", "txt", "xls", "xlsx", "xlsm"],
        "kind": "files",
        "targets": [
            {"id": "md", "title": "Markdown"},
            {"id": "word", "title": "Word"},
            {"id": "xlsx", "title": "Excel"},
            {"id": "csv", "title": "CSV"},
            {"id": "txt", "title": "纯文本"},
        ],
    },
    {
        "id": "split",
        "verb": "split",
        "title": "拆分",
        "subtitle": "md 按标题 / xlsx 按 sheet",
        "icon": "scissors",
        "exts": ["md", "xlsx", "xlsm"],
        "kind": "files",
    },
    {
        "id": "merge",
        "verb": "merge",
        "title": "合并",
        "subtitle": "多个 md→一篇 / 多个 txt→CSV",
        "icon": "arrow.triangle.merge",
        "exts": ["md", "txt"],
        "kind": "files",
    },
    {
        "id": "typeset",
        "verb": "typeset",
        "title": "套模板成品 Word",
        "subtitle": "md/docx/doc → 院模板(套样式·修文本·图注居中)",
        "icon": "doc.richtext",
        "exts": ["md", "docx", "doc"],
        "kind": "files",
    },
    {
        "id": "renum",
        "verb": "renum",
        "title": "序号修正",
        "subtitle": "标题/图/表编号断号·错号·缺号一键重排(产出 _序号修正.docx,原件不动)",
        "icon": "list.number",
        "exts": ["docx"],
        "kind": "files",
        "targets": [
            {"id": "all", "title": "全部（标题+图+表）"},
            {"id": "tabfig", "title": "仅图表号"},
            {"id": "headings", "title": "仅标题号"},
        ],
    },
    {
        "id": "bidfinal",
        "verb": "bidfinal",
        "title": "标书终稿门检",
        "subtitle": "残留8类/身份泄漏/打印就绪 三道门干跑体检,只诊断不改文件;红门=半成品禁交付",
        "icon": "checkmark.seal",
        "exts": ["docx"],
        "kind": "files",
        "targets": [
            {"id": "pei", "title": "陪标·通用稿"},
            {"id": "main", "title": "主标·实名"},
        ],
    },
    {
        "id": "scan",
        "verb": "scan",
        "title": "敏感词扫描",
        "subtitle": "扫一个目录里的 md/docx(竞品名/过硬措辞)",
        "icon": "magnifyingglass",
        "exts": [],
        "kind": "dir",
    },
    {
        "id": "view",
        "verb": "view",
        "title": "预览",
        "subtitle": "md → HTML 浏览器预览(2026-06-14 从 raycast doc_preview 并入)",
        "icon": "eye",
        "exts": ["md"],
        "kind": "files",
    },
]

_OPS_BY_ID = {o["id"]: o for o in OPS}


# ───────────────────────────────────────────── 产出探测(目录树快照差分)

def _snapshot(roots: list[Path]) -> set[str]:
    """收集 roots 下(含子目录)所有现存路径,用于跑前/跑后差分。"""
    seen: set[str] = set()
    for r in roots:
        if not r.exists():
            continue
        seen.add(str(r))
        if r.is_dir():
            for p in r.rglob("*"):
                seen.add(str(p))
    return seen


def _scan_roots(files: list[str]) -> list[Path]:
    """每个输入文件所在目录 = 监控根(产出都落在源文件同级或其子目录)。"""
    roots: set[Path] = set()
    for f in files:
        p = Path(f)
        roots.add(p.parent if p.parent != Path("") else Path.cwd())
    return list(roots)


def _new_outputs(before: set[str], after: set[str], inputs: set[str]) -> list[str]:
    """跑后新增、且非输入本身的路径 = 产出。顶层去重(目录产出不再列其子项)。"""
    created = sorted(p for p in (after - before) if p not in inputs)
    top: list[str] = []
    for p in created:
        if any(p != q and p.startswith(q.rstrip("/") + "/") for q in created):
            continue  # 是某个新建目录的子项,不单列
        top.append(p)
    return top


# ───────────────────────────────────────────── gui-run

def _run_verb_capture(verb: str, files: list[str], target: str | None) -> tuple[int, str]:
    """调 doc_dispatch 的 do_* 实现,把所有日志(do_* 的 print + 子进程输出)关进缓冲,
    绝不让任何文本漏进真 stdout(那是 JSON 信封专用)。返回 (rc, captured_text)。
    两路收口:① redirect_stdout 接 do_* 自己的 print;② _CAP(monkeypatch 的 _run)接子进程。"""
    _CAP.clear()
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        if verb == "clean":
            rc = dd.do_clean(files)
        elif verb == "split":
            rc = dd.do_split(files)
        elif verb == "merge":
            rc = dd.do_merge(files)
        elif verb == "typeset":
            rc = dd.do_typeset(files)
        elif verb == "convert":
            rc = dd.do_convert(files, target or "md")
        elif verb == "renum":
            rc = dd.do_renum(files, target or "all")
        elif verb == "bidfinal":
            rc = dd.do_bidfinal(files, target or "pei")
        elif verb == "scan":
            rc = dd.do_scan(files)
        elif verb == "view":
            rc = dd.do_view(files)
        else:
            return 2, f"未知操作: {verb}"
    log = _strip_ansi("\n".join([buf.getvalue().rstrip(), *_CAP]).strip())
    return rc, log


def gui_run(op_id: str, files: list[str], target: str | None) -> dict:
    op = _OPS_BY_ID.get(op_id)
    if not op:
        return {"ok": False, "error": f"未知操作: {op_id}(支持 {', '.join(_OPS_BY_ID)})"}

    if op["kind"] == "dir":
        # scan:吃一个目录
        if not files:
            return {"ok": False, "error": "请选择一个目录"}
        d = files[0]
        if not Path(d).is_dir():
            return {"ok": False, "error": f"不是目录: {d}"}
        roots = [Path(d)]
        before = _snapshot(roots)
        rc, log = _run_verb_capture("scan", [d], None)
        return {
            "ok": rc == 0,
            "op": op_id,
            "results": [{
                "input": d,
                "name": Path(d).name,
                "ok": rc == 0,
                "outputs": _new_outputs(before, _snapshot(roots), {d}),
                "message": "扫描完成" if rc == 0 else "扫描出错(详见日志)",
            }],
            "log": log.strip(),
        } if rc == 0 else {"ok": False, "error": "敏感词扫描失败", "log": log.strip()}

    if op_id == "convert" and not target:
        return {"ok": False, "error": "格式转换需指定目标格式(--to)"}

    existing = [f for f in files if Path(f).exists()]
    missing = [f for f in files if not Path(f).exists()]
    if not existing:
        return {"ok": False, "error": "没有有效文件(全部不存在)"}

    # merge 是多对一,无法逐文件归因产出 → 整组跑一次,产出整体列出。
    if op_id == "merge":
        roots = _scan_roots(existing)
        before = _snapshot(roots)
        rc, log = _run_verb_capture("merge", existing, None)
        outs = _new_outputs(before, _snapshot(roots), set(existing))
        results = [{
            "input": " + ".join(Path(f).name for f in existing),
            "name": f"合并 {len(existing)} 个文件",
            "ok": rc == 0 and bool(outs),
            "outputs": outs,
            "message": ("已合并 → " + ", ".join(Path(o).name for o in outs)) if outs
                       else "未产出(可能格式不支持合并)",
        }]
        return _wrap(op_id, results, missing, log)

    # 其余动词:逐文件跑(才能逐文件归因产出)。
    results = []
    full_log: list[str] = []
    for f in existing:
        roots = _scan_roots([f])
        before = _snapshot(roots)
        rc, log = _run_verb_capture(op["verb"], [f], target)
        full_log.append(log.strip())
        outs = _new_outputs(before, _snapshot(roots), {f})
        ok = rc == 0 and bool(outs)
        results.append({
            "input": f,
            "name": Path(f).name,
            "ok": ok,
            "outputs": outs,
            "message": ("→ " + ", ".join(Path(o).name for o in outs)) if outs
                       else ("无对应引擎/未产出(可能此格式不支持该操作)" if rc == 0
                             else "处理失败(详见日志)"),
        })
    return _wrap(op_id, results, missing, "\n".join(full_log))


def _wrap(op_id: str, results: list[dict], missing: list[str], log: str) -> dict:
    out = {
        "ok": True,
        "op": op_id,
        "results": results,
        "succeeded": sum(1 for r in results if r["ok"]),
        "total": len(results),
        "log": log.strip(),
    }
    if missing:
        out["skipped_missing"] = [Path(m).name for m in missing]
    return out


# ───────────────────────────────────────────── CLI

def main() -> int:
    ap = argparse.ArgumentParser(description="DocTools GUI 后端(JSON 信封)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("gui-ops")

    pr = sub.add_parser("gui-run")
    pr.add_argument("--op", required=True)
    pr.add_argument("--to", dest="target", default=None)
    pr.add_argument("--files", nargs="+", required=True)

    a = ap.parse_args()

    try:
        if a.cmd == "gui-ops":
            payload = {"ok": True, "ops": OPS}
        elif a.cmd == "gui-run":
            payload = gui_run(a.op, a.files, a.target)
        else:
            payload = {"ok": False, "error": f"未知子命令: {a.cmd}"}
    except Exception as e:  # noqa: BLE001 — 任何异常都转人话信封,绝不让 Swift 见 traceback
        payload = {"ok": False, "error": f"后端异常: {type(e).__name__}: {e}"}

    print(json.dumps(payload, ensure_ascii=False))
    return 0  # gui-* 一律 exit 0(信封承载成败)


if __name__ == "__main__":
    sys.exit(main())
