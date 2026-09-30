#!/usr/bin/env python3
"""
DocKit 后端 —— GUI 的 JSON 适配器,也是 agent 用的 `dockit` 命令行(同一个业务层)。

设计(Tier B · 零重写业务):
  · 复用 doc_dispatch 的 route_*/do_* 路由逻辑;校验、外发同意门、产出归因只在 gui_run 一处。
  · 把它们「打文本到 stdout + 产出落盘」的终端 UX,翻译成纯 JSON stdout 信封。
  · 产出路径靠「跑前/跑后扫目录树」差分得到(确定性,不解析中文日志)。

GUI 契约(DocKit.app 经 `uv run --project ~/Dev` 调用):所有 gui-* 一律 exit 0;
  成功 {"ok": true, ...},失败 {"ok": false, "error": "人话", "error_code": "..."};snake_case。
  gui-ops                        列出可用操作(给 UI 渲菜单)
  gui-run --op <verb> [--to T] [--opt K=V]... --files <paths...>   跑一个操作,返回逐文件结果

agent 契约(App 包内 Contents/Resources/bin/dockit → 本文件;退出码承载成败):
  ops [<op>] [--json]            操作目录 / 单个操作的目标、选项与默认值
  run <op> [--to T] [--opt K=V]... [--dry-run] [--yes] [--open] [--background] [--json] <paths...>
                                 与 GUI「执行」同一条 gui_run 路径;--background 返回 job id
  status [<job>] [--json]        后台任务的状态与结果
  doctor [--json]                依赖就绪检查(对应 GUI 的「已就绪 / 后端不可达」)
  退出码 0 成功 · 1 业务失败 · 2 用法/校验错误 · 75 同一目录树另有 DocKit 操作在跑
         · 128+N 被信号 N 中断(引擎子进程已一并终止)

原有终端用法(doc_dispatch.py <verb>)不受影响。
"""
from __future__ import annotations

import argparse
import contextlib
import fcntl
import hashlib
import importlib.util
import io
import json
import os
import re
import secrets
import shutil
import signal
import subprocess
import sys
import tempfile
import time
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
# 每次子进程的原始 rc/stdout:do_* 会把 rc 折成 0/1(bidfinal 的红门 2 与 IO 错 1 分不开),
# scan --json 的发现项也只在 stdout 里 —— 需要原值的判定从这里取,不解析中文日志。
_RUNS: list[dict] = []
# 正在跑的引擎子进程。每个引擎单独一个进程组(start_new_session):本进程被 agent 的超时
# SIGTERM/SIGINT 杀掉时,信号处理把整组(含 soffice 这类孙进程)一起终止,
# 不会留下没人管、还在往用户目录写的引擎,而目录锁却已经放掉。
_CURRENT: list[subprocess.Popen] = []


def _captured_run(cmd, label):
    """dd._run 的捕获版:子进程输出进 PIPE 不外泄,攒进 _CAP。返回 returncode(签名同原版)。
    stdin 给 /dev/null:引擎一律非交互,agent 调用时不会卡在等输入。"""
    _CAP.append(f"  ↳ {label}")
    proc = subprocess.Popen(cmd, env=dd._ENV, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, start_new_session=True)
    _CURRENT.append(proc)
    try:
        out, err = proc.communicate()
    finally:
        _CURRENT.remove(proc)
    if out:
        _CAP.append(out.rstrip())
    if err:
        _CAP.append(err.rstrip())
    _RUNS.append({"rc": proc.returncode, "stdout": out or ""})
    return proc.returncode


dd._run = _captured_run  # noqa: SLF001 — 故意替换,把终端日志关进缓冲


class _Interrupted(BaseException):
    """收到 SIGTERM/SIGINT/SIGHUP:引擎进程组已终止,沿调用栈退出(途经 finally 放掉目录锁)。"""

    def __init__(self, signum: int):
        super().__init__(signum)
        self.signum = signum


_STOPPING: list[int] = []


def _terminate_engines(grace: float = 5.0) -> None:
    procs = [p for p in _CURRENT if p.poll() is None]
    for p in procs:
        with contextlib.suppress(ProcessLookupError, PermissionError):
            os.killpg(p.pid, signal.SIGTERM)
    deadline = time.monotonic() + grace
    for p in procs:
        try:
            p.wait(timeout=max(0.1, deadline - time.monotonic()))
        except subprocess.TimeoutExpired:
            with contextlib.suppress(ProcessLookupError, PermissionError):
                os.killpg(p.pid, signal.SIGKILL)


def _on_signal(signum, _frame):
    if _STOPPING:          # 收尾期间再来的信号不重入
        return
    _STOPPING.append(signum)
    _terminate_engines()
    raise _Interrupted(signum)


def _install_signal_handlers() -> None:
    for s in (signal.SIGTERM, signal.SIGINT, signal.SIGHUP):
        signal.signal(s, _on_signal)

# ───────────────────────────────────────────── 操作目录(SSOT,UI 从这渲菜单)

# 每个 op:verb=doc_dispatch 动词;exts=支持的源后缀(给 UI 做拖入校验提示);
# targets=convert 专用的目标格式列表;scan=是否吃目录(而非文件)。

# 产出名与源同名换后缀(b.md → b.docx)、merged.md、按 sheet 命名的操作:目标已存在时
# 引擎会直接覆盖且不留备份。gui_run 默认拒绝(would_overwrite),人要覆盖就勾这一项,
# agent 用 --yes(等价 --opt output.overwrite=1)。GUI 按 type 泛化渲染,不认识这个 id。
OVERWRITE_OPT = "output.overwrite"
_OVERWRITE_GROUP = {"id": "output", "title": "已有同名文件", "danger": True}
_OVERWRITE_OPTION = {"id": OVERWRITE_OPT, "group": "output", "type": "bool", "default": False,
                     "title": "允许覆盖已有同名文件",
                     "note": "默认遇到同名文件就停下;勾选后直接覆盖,不留备份"}

OPS = [
    {
        "id": "clean",
        "aliases": "normalize clean tidy format punctuation quotes units",
        "verb": "clean",
        "title": "规范化",
        "subtitle": "纯文本修复:引号+标点+单位一起做(docx/md 另存 _fixed;pptx 原地改写并留备份)",
        "icon": "wand.and.stars",
        "exts": ["docx", "md", "pptx"],
        "kind": "files",
        # 勾选项 SSOT。Swift 侧只按 type 泛化渲染(bool→Toggle),不认识任何 option id ——
        # 以后加旋钮只改这里,不重编 .app。applies_to 用于「该格式不吃这项」时灰显。
        "option_groups": [
            {"id": "rule", "title": "修哪些内容"},
            {"id": "scope", "title": "改哪些范围", "applies_to": ["docx"]},
        ],
        "options": [
            {"id": "rule.quotes", "group": "rule", "type": "bool", "default": True,
             "title": "引号统一为中文引号"},
            {"id": "rule.punct", "group": "rule", "type": "bool", "default": True,
             "title": "英文标点转中文", "note": "含 , : ; ! ? ( )"},
            {"id": "rule.units", "group": "rule", "type": "bool", "default": True,
             "title": "中文单位转标准符号", "note": "平方米→m²(仅数字后)"},
            {"id": "rule.quote_font", "group": "rule", "type": "bool", "default": True,
             "title": "引号设为宋体", "note": "会把引号拆成独立 run"},
            {"id": "scope.body", "group": "scope", "type": "bool", "default": True, "title": "正文"},
            {"id": "scope.table", "group": "scope", "type": "bool", "default": True, "title": "表格"},
            {"id": "scope.revision", "group": "scope", "type": "bool", "default": True,
             "title": "审阅修订", "note": "w:ins/w:del 插入与删除态"},
            {"id": "scope.comments", "group": "scope", "type": "bool", "default": True, "title": "批注"},
            {"id": "scope.notes", "group": "scope", "type": "bool", "default": True, "title": "脚注/尾注"},
            {"id": "scope.headers", "group": "scope", "type": "bool", "default": True, "title": "页眉页脚"},
        ],
    },
    {
        "id": "quotes",
        "aliases": "quotes quotation curly smart 引号",
        "verb": "quotes",
        "title": "引号统一",
        "subtitle": "只把引号统一成中文弯引号,不碰标点/单位(docx·md)",
        "icon": "quote.opening",
        "exts": ["docx", "md"],
        "kind": "files",
        "option_groups": [
            {"id": "rule", "title": "怎么处理"},
            {"id": "scope", "title": "改哪些范围", "applies_to": ["docx"]},
        ],
        "options": [
            {"id": "rule.quote_font", "group": "rule", "type": "bool", "default": True,
             "title": "引号设为宋体", "note": "会把引号拆成独立 run"},
            {"id": "scope.body", "group": "scope", "type": "bool", "default": True, "title": "正文"},
            {"id": "scope.table", "group": "scope", "type": "bool", "default": True, "title": "表格"},
            {"id": "scope.revision", "group": "scope", "type": "bool", "default": True,
             "title": "审阅修订", "note": "w:ins/w:del 插入与删除态"},
            {"id": "scope.comments", "group": "scope", "type": "bool", "default": True, "title": "批注"},
            {"id": "scope.notes", "group": "scope", "type": "bool", "default": True, "title": "脚注/尾注"},
            {"id": "scope.headers", "group": "scope", "type": "bool", "default": True, "title": "页眉页脚"},
        ],
    },
    {
        "id": "fontunify",
        "aliases": "font unify typeface pptx 字体",
        "verb": "fontunify",
        "title": "字体统一",
        "subtitle": "pptx 全篇(含母版/版式)统一字体 ⚠原地覆写原文件(留 .backup)",
        "icon": "textformat",
        "exts": ["pptx"],
        "kind": "files",
        "danger": True,
    },
    {
        "id": "lowercase",
        "aliases": "lowercase lower case downcase xlsx 小写",
        "verb": "lowercase",
        "title": "英文小写整理",
        "subtitle": "xlsx 数据行 / docx 正文里的英文转小写(语义级数据改写)",
        "icon": "textformat.abc",
        "exts": ["xlsx", "xlsm", "docx"],
        "kind": "files",
        "danger": True,
    },
    {
        "id": "stripchrome",
        "aliases": "strip header footer chrome remove 页眉 页脚",
        "verb": "stripchrome",
        "title": "清页眉页脚",
        "subtitle": "删除 docx 的页眉页脚引用,一个字不改 ⚠不可撤销(产出 _fixed 副本)",
        "icon": "rectangle.topthird.inset.filled",
        "exts": ["docx"],
        "kind": "files",
        "danger": True,
    },
    {
        # 2026-07-27 从「TL 代笔台」收编(该 app 实测 18 天开过 1 次共 0 分钟,
        # 唯一动作就是调 docx_fmt.py clone,原 docx_format_clone.py)。能力留在这里,壳退役。
        # 引擎是 HQ SSOT(/docx format 也在用),此处只做编排。
        "id": "formatclone",
        "aliases": "format clone style template reference 版式 复刻 公文",
        "verb": "formatclone",
        "title": "公文版式复刻",
        "subtitle": "拿一份范式件当格式源,把内容刷成同款版式(原件不动,产出 _成品.docx,永不覆盖)",
        "icon": "doc.on.doc",
        "exts": ["docx", "md"],
        "kind": "files",
        "option_groups": [
            {"id": "src", "title": "格式来源(必选)"},
            {"id": "opt", "title": "选项"},
        ],
        "options": [
            # type=file:声明式契约的第二个类型(2026-07-27 加)。Swift 按 type 泛化渲染成
            # 文件选择器,同样不出现任何 option id 字面量。
            {"id": "ref", "group": "src", "type": "file", "exts": ["docx"], "required": True,
             "title": "范式 docx", "note": "格式从这份抽(标题/正文/落款的直排版一并复刻)"},
            {"id": "signature", "group": "opt", "type": "bool", "default": False,
             "title": "保留落款署名", "note": "默认不含 —— 复刻件多数要另行署名"},
        ],
    },
    {
        "id": "convert",
        "aliases": "convert transform export pdf word markdown 转换",
        "verb": "convert",
        "title": "格式转换",
        "subtitle": "源格式自动识别 → 目标格式（含 PDF → 可编辑 Word）",
        "icon": "arrow.triangle.2.circlepath",
        "exts": ["pdf", "docx", "doc", "pptx", "ppt", "md", "csv", "txt", "xls", "xlsx", "xlsm"],
        "kind": "files",
        "option_groups": [_OVERWRITE_GROUP],
        "options": [_OVERWRITE_OPTION],
        # 其余带 targets 的操作不给 --to 时取第一项;convert 没有合理默认,必须显式选。
        "target_required": True,
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
        "aliases": "split divide separate chapters sheets 拆分",
        "verb": "split",
        "title": "拆分",
        "subtitle": "md 按标题 / xlsx 按 sheet",
        "icon": "scissors",
        "exts": ["md", "xlsx", "xlsm"],
        "kind": "files",
        "option_groups": [_OVERWRITE_GROUP],
        "options": [_OVERWRITE_OPTION],
    },
    {
        "id": "merge",
        "aliases": "merge combine join concat 合并",
        "verb": "merge",
        "title": "合并",
        "subtitle": "多个 md → merged.md / 多个 txt 按列 → merged.csv(产出在第一份旁)",
        "icon": "arrow.triangle.merge",
        "exts": ["md", "txt"],
        "kind": "files",
        "option_groups": [_OVERWRITE_GROUP],
        "options": [_OVERWRITE_OPTION],
    },
    {
        "id": "typeset",
        "aliases": "typeset template word 套模板 成品",
        "verb": "typeset",
        "title": "套模板成品 Word",
        "subtitle": "md/docx/doc → 院模板(套样式·修文本·图注居中)",
        "icon": "doc.richtext",
        "exts": ["md", "docx", "doc"],
        "kind": "files",
        "option_groups": [_OVERWRITE_GROUP],
        "options": [_OVERWRITE_OPTION],
    },
    {
        "id": "renum",
        "aliases": "renumber renum numbering figures tables headings 序号",
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
        "aliases": "bid final gate check residue identity print 标书 门检",
        "verb": "bidfinal",
        "title": "标书终稿门检",
        "subtitle": "残留8类/身份泄漏/打印就绪 三道门干跑体检,只诊断不改文件;红门=半成品禁交付",
        "icon": "checkmark.seal",
        "exts": ["docx"],
        "kind": "files",
        # 诊断型:不写产出文件,成败看门检退出码(0 全绿 · 2 有红门 · 其余=出错),不看「有无新产出」。
        "produces": "verdict",
        "targets": [
            {"id": "pei", "title": "陪标·通用稿"},
            {"id": "main", "title": "主标·实名"},
        ],
    },
    {
        "id": "scan",
        "aliases": "scan sensitive detect audit 敏感词",
        "verb": "scan",
        "title": "敏感词扫描",
        # 引擎只收集目录内 .md(scan_sensitive_words.py 不读 docx),说明按实际范围写。
        "subtitle": "云端扫描目录内 .md：文件名与内容发送至 Claude，需先明确同意",
        "icon": "magnifyingglass",
        "exts": [],
        "kind": "dir",
        "option_groups": [{"id": "privacy", "title": "外发授权"}],
        "options": [
            {"id": "privacy.cloud_consent", "group": "privacy", "type": "bool",
             "default": False, "title": "同意将所选目录内文件名与内容发送至 Claude",
             "note": "敏感词扫描使用云端模型；规范化与引号统一在本机处理。"},
        ],
    },
    {
        "id": "view",
        "aliases": "view preview render html 预览",
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

def _snapshot(roots: list[Path]) -> dict[str, tuple]:
    """收集 roots 下(含子目录)所有现存路径 → (mtime_ns, size),用于跑前/跑后差分。

    带指纹而非只带路径:同名产出已存在时(同一文件跑第二遍,_fixed.docx 覆盖旧的),
    纯路径差分为空 → GUI 会把成功的跑报成「未产出」。"""
    seen: dict[str, tuple] = {}
    for r in roots:
        if not r.exists():
            continue
        for p in [r, *(r.rglob("*") if r.is_dir() else [])]:
            try:
                st = p.stat()
                seen[str(p)] = (st.st_mtime_ns, st.st_size)
            except OSError:
                seen[str(p)] = ()
    return seen


def _scan_roots(files: list[str]) -> list[Path]:
    """每个输入文件所在目录 = 监控根(产出都落在源文件同级或其子目录)。"""
    roots: set[Path] = set()
    for f in files:
        p = Path(f)
        roots.add(p.parent if p.parent != Path("") else Path.cwd())
    return list(roots)


def _new_outputs(before: dict[str, tuple], after: dict[str, tuple], inputs: set[str]) -> list[str]:
    """跑后**新增或被覆盖**、且非输入本身的路径 = 产出。顶层去重(目录产出不再列其子项)。

    目录的 mtime 会因内部新建文件而变,所以「被覆盖」只认文件,不认目录。"""
    created = sorted(
        p for p, fp in after.items()
        if p not in inputs and (p not in before or (before[p] != fp and not Path(p).is_dir()))
    )
    top: list[str] = []
    for p in created:
        if any(p != q and p.startswith(q.rstrip("/") + "/") for q in created):
            continue  # 是某个新建目录的子项,不单列
        top.append(p)
    return top


# ───────────────────────────────────────────── 目录互斥
#
# 产出归因靠对源文件所在目录**整棵子树**做跑前/跑后快照差分。GUI 与 CLI(或两个 CLI)同时
# 在同一目录、或一个在父目录一个在子目录里跑时,一方的产出会被记到另一方头上。
# 所以锁是分层的:本次监控根拿 LOCK_EX,它的每一级上级目录拿 LOCK_SH ——
#   父目录的操作(EX 父)与子目录的操作(SH 父)互斥;兄弟目录的操作(SH 同一父)互不影响。
# 全部非阻塞:抢不到就报 busy,不排队、不猜。锁文件放 App 自己的缓存目录,不往用户文档目录写;
# 释放时确认没人再持有才删掉锁文件(删前核对 inode,拿锁后也核对,防删与开之间的竞态)。

CACHE_DIR = Path(os.environ.get("DOCKIT_CACHE_DIR")
                 or Path.home() / "Library" / "Caches" / "cyou.tianli.DocTools")
LOCK_DIR = CACHE_DIR / "locks"
JOB_DIR = CACHE_DIR / "jobs"


class _Busy(Exception):
    """某个监控根(或其上级/下级目录)已被另一个 DocKit 操作占用。"""


def lock_path(root: Path | str) -> Path:
    key = os.path.realpath(str(root))
    return LOCK_DIR / (hashlib.sha256(key.encode("utf-8")).hexdigest()[:24] + ".lock")


def _lock_plan(roots) -> list[tuple[str, int]]:
    """[(目录实路径, LOCK_EX|LOCK_SH)],按路径排序。本身是监控根的目录只拿 EX。"""
    exclusive = {os.path.realpath(str(r)) for r in roots}
    shared: set[str] = set()
    for key in exclusive:
        parent = os.path.dirname(key)
        while parent and parent not in shared:
            shared.add(parent)
            if parent == os.path.dirname(parent):
                break
            parent = os.path.dirname(parent)
    shared -= exclusive
    plan = [(k, fcntl.LOCK_EX) for k in exclusive] + [(k, fcntl.LOCK_SH) for k in shared]
    return sorted(plan)


def _acquire(key: str, mode: int) -> int:
    path = lock_path(key)
    for _ in range(5):
        fd = os.open(path, os.O_RDWR | os.O_CREAT, 0o600)
        try:
            fcntl.flock(fd, mode | fcntl.LOCK_NB)
        except BlockingIOError:
            os.close(fd)
            raise _Busy(key) from None
        try:
            same = os.fstat(fd).st_ino == os.stat(path).st_ino
        except FileNotFoundError:
            same = False
        if same:
            return fd
        os.close(fd)   # 拿到的是刚被释放方删掉的旧文件,重开
    raise _Busy(key)


def _release(key: str, fd: int) -> None:
    path = lock_path(key)
    try:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)   # 能独占 = 没有别人持有 → 可以删
        except BlockingIOError:
            return
        with contextlib.suppress(FileNotFoundError):
            if os.fstat(fd).st_ino == os.stat(path).st_ino:
                os.unlink(path)
    finally:
        with contextlib.suppress(OSError):
            fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


@contextlib.contextmanager
def _dir_locks(roots):
    held: list[tuple[str, int]] = []
    try:
        LOCK_DIR.mkdir(parents=True, exist_ok=True)
        for key, mode in _lock_plan(roots):
            held.append((key, _acquire(key, mode)))
        yield
    finally:
        for key, fd in reversed(held):
            _release(key, fd)


# ───────────────────────────────────────────── 错误码(GUI 读 error 文本,CLI 按 error_code 定退出码)

EXIT_OK, EXIT_FAILED, EXIT_USAGE, EXIT_BUSY = 0, 1, 2, 75
USAGE_ERRORS = frozenset({
    "usage", "unknown_op", "unknown_option", "bad_option_value", "bad_opt_form",
    "option_file_missing", "missing_required", "consent_required", "confirm_required",
    "missing_target", "bad_target", "no_input", "not_a_dir", "no_valid_files",
    "would_overwrite", "unknown_job",
})
CONSENT_OPT = "privacy.cloud_consent"
_TRUE = ("1", "true", "yes", "on")


def _err(code: str, message: str, **extra) -> dict:
    return {"ok": False, "error": message, "error_code": code, **extra}


def _busy(b: _Busy) -> dict:
    return _err("busy", f"同一目录(或其上级/下级目录)正有另一个 DocKit 操作在运行({b}),等它完成后重试")


def _log_hint(log: str) -> str:
    """日志里最能说明问题的一行:优先最后一条告警/失败/跳过,否则最后一行。"""
    lines = [ln.strip() for ln in log.splitlines() if ln.strip()]
    for markers in (("已存在",), ("⚠", "❌", "✖", "失败", "Error", "error"), ("跳过",)):
        for ln in reversed(lines):
            if any(m in ln for m in markers):
                return ln[:160]
    return lines[-1][:160] if lines else ""


# ───────────────────────────────────────────── gui-run

def _run_verb_capture(verb: str, files: list[str], target: str | None,
                      opts: dict | None = None, *, view_dir: Path | None = None,
                      open_viewer: bool = True, scan_json: bool = False) -> tuple[int, str]:
    """调 doc_dispatch 的 do_* 实现,把所有日志(do_* 的 print + 子进程输出)关进缓冲,
    绝不让任何文本漏进真 stdout(那是 JSON 信封专用)。返回 (rc, captured_text)。
    两路收口:① redirect_stdout 接 do_* 自己的 print;② _CAP(monkeypatch 的 _run)接子进程。"""
    _CAP.clear()
    _RUNS.clear()
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        if verb == "clean":
            rc = dd.do_clean(files, opts)
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
            rc = dd.do_scan(files, json_out=scan_json)
        elif verb == "view":
            rc = dd.do_view(files, output_dir=str(view_dir) if view_dir else None,
                            open_browser=open_viewer)
        elif verb == "quotes":
            rc = dd.do_quotes(files, opts)
        elif verb == "fontunify":
            rc = dd.do_fontunify(files)
        elif verb == "lowercase":
            rc = dd.do_lowercase(files)
        elif verb == "stripchrome":
            rc = dd.do_stripchrome(files)
        elif verb == "formatclone":
            rc = dd.do_formatclone(files, opts)
        else:
            return 2, f"未知操作: {verb}"
    log = _strip_ansi("\n".join([buf.getvalue().rstrip(), *_CAP]).strip())
    return rc, log


def _parse_findings(stdout: str) -> list | None:
    """scan_sensitive_words --json 的 stdout 末尾是 JSON 数组;前面可能夹着 show_* 的告警行。"""
    text = stdout.strip()
    candidates = [text]
    cut = text.rfind("\n[")
    if cut >= 0:
        candidates.append(text[cut + 1:])
    for c in candidates:
        try:
            value = json.loads(c)
        except ValueError:
            continue
        if isinstance(value, list):
            return value
    return None


def _supports(op: dict, f: str, target: str | None) -> tuple[bool, str]:
    """该输入有没有对应引擎(与 do_* 的路由判断同源,不执行)。dry-run 与「未产出」的说明都用它。"""
    e = dd._ext(f)  # noqa: SLF001
    if op["id"] == "convert":
        ok = target in ("txt", "word", "md") if e == "doc" else dd.route_convert(f, target) is not None
        return ok, (f"将转换为 {target}" if ok else f"{e or '?'} → {target} 没有转换引擎")
    if op["id"] == "merge":
        ok = e in ("md", "txt")
        return ok, ("将参与合并" if ok else f".{e or '?'} 没有合并引擎")
    ok = e in op.get("exts", [])
    return ok, ("将处理" if ok else f"{op['title']}不支持 .{e or '?'}")


def _overwrite_clashes(op: dict, existing: list[str], target: str | None) -> dict[str, list[str]]:
    """{输入: [会被覆盖的已有文件]}。产出命名来自 doc_dispatch(与路由同源),只列真的已存在、
    又不是本次输入的路径;DocKit 自己后缀的产出(_fixed 等)重跑覆盖是既定行为,不在此列。"""
    if OVERWRITE_OPT not in {o["id"] for o in op.get("options", [])}:
        return {}
    inputs = {os.path.realpath(f) for f in existing}
    if op["verb"] == "merge":
        plan = {existing[0]: dd.merge_outputs(existing)}
    else:
        plan = {f: dd.planned_outputs(op["verb"], f, target) for f in existing}
    clashes = {}
    for f, outs in plan.items():
        hit = [str(p) for p in outs if p.exists() and os.path.realpath(p) not in inputs]
        if hit:
            clashes[f] = hit
    return clashes


def _plan(op: dict, existing: list[str], missing: list[str], target: str | None,
          opts: dict | None, clashes: dict[str, list[str]]) -> dict:
    results = []
    for f in existing:
        ok, why = _supports(op, f, target)
        row = {"input": f, "name": Path(f).name, "ok": ok, "outputs": [], "message": why}
        if ok and clashes.get(f):
            row["overwrites"] = clashes[f]
            row["message"] = why + ";将覆盖已有 " + ", ".join(Path(p).name for p in clashes[f])
        results.append(row)
    out = {"ok": True, "dry_run": True, "op": op["id"], "target": target, "options": dict(opts or {}),
           "results": results, "succeeded": sum(1 for r in results if r["ok"]),
           "total": len(results), "log": ""}
    if missing:
        out["skipped_missing"] = [Path(m).name for m in missing]
    return out


def _verdict_row(f: str, outs: list[str]) -> dict:
    """诊断型操作(produces=verdict)的逐文件结果:看引擎原始退出码,不看有无产出。"""
    gate = _RUNS[-1]["rc"] if _RUNS else None
    if gate is None:
        verdict, ok, hard, msg = None, False, 0, "无对应引擎(可能此格式不支持该操作)"
    elif gate == 0:
        verdict, ok, hard, msg = "pass", True, 0, "四门全绿,可交付"
    elif gate == 2:
        verdict, ok, hard, msg = "red", False, 0, "有红门:半成品,禁交付(详见日志)"
    else:
        verdict, ok, hard, msg = "error", False, gate, "门检出错(详见日志)"
    return {"_rc": hard, "input": f, "name": Path(f).name, "ok": ok, "outputs": outs,
            "message": msg, "verdict": verdict}


def _md_files(d: str) -> int:
    """scan 引擎收集的范围:目录内(含子目录)全部 .md。"""
    return sum(1 for p in Path(d).glob("**/*.md") if p.is_file())


def gui_run(op_id: str, files: list[str], target: str | None, opts: dict | None = None, *,
            open_viewer: bool = True, scan_findings: bool = False, dry_run: bool = False) -> dict:
    """GUI「执行」与 `dockit run` 的唯一执行路径。

    open_viewer   view 是否打开浏览器(GUI 打开;CLI 默认不打开,不抢焦点)
    scan_findings scan 透传引擎 --json,把发现项放进 results[0].findings
    dry_run       走完全部校验后只报告将处理哪些输入、会覆盖哪些已有文件;不跑引擎、不取锁、不写盘

    信封的 ok 是**请求级**:请求被处理了就是 true,哪怕个别输入失败。逐个输入看 results[].ok;
    all_ok = 全部输入成功且没有被跳过的输入(即 CLI 退出码为 0)。
    """
    payload = _gui_run(op_id, files, target, opts, open_viewer=open_viewer,
                       scan_findings=scan_findings, dry_run=dry_run)
    payload["all_ok"] = exit_code(payload) == EXIT_OK
    return payload


def _gui_run(op_id: str, files: list[str], target: str | None, opts: dict | None, *,
             open_viewer: bool, scan_findings: bool, dry_run: bool) -> dict:
    op = _OPS_BY_ID.get(op_id)
    if not op:
        return _err("unknown_op", f"未知操作: {op_id}(支持 {', '.join(_OPS_BY_ID)})")
    # 相对路径一律按当前目录展开:快照差分的键是绝对路径,相对输入会被误当成「新产出」
    files = [os.path.abspath(os.path.expanduser(f)) for f in files]

    # 选项校验:未知 key / 非布尔值一律走信封 ok:false,**不许 argparse exit 2**
    # (GUI 契约:所有 gui-* 一律 exit 0,成败只由信封承载)
    declared = {o["id"] for o in op.get("options", [])}
    if opts:
        bad = sorted(set(opts) - declared)
        if bad:
            return _err("unknown_option",
                        f"未知选项: {', '.join(bad)}"
                        + (f"(该操作可用 {', '.join(sorted(declared))})" if declared
                           else "(该操作不接受选项)"))
        _types = {o["id"]: o.get("type", "bool") for o in op.get("options", [])}
        for k, v in opts.items():
            t = _types.get(k, "bool")
            if t == "file":
                # file 型的值是绝对路径。空值放行(由必填检查或动词自己报),
                # 但给了就必须真存在 —— 否则错误会在子进程里以退出码形式出现,人看不懂。
                if str(v).strip() and not Path(str(v)).expanduser().exists():
                    return _err("option_file_missing", f"选项 {k} 指向的文件不存在: {v}")
                continue
            if str(v).strip().lower() not in ("0", "1", "true", "false", "yes", "no", "on", "off"):
                return _err("bad_option_value", f"选项 {k} 的值 {v!r} 不是布尔(用 1/0)")

    # 必填选项(2026-07-27 立):缺了必须在契约层报人话。
    # 踩过才加 —— formatclone 不选「范式 docx」时,动词只能 warn 后跳过,信封却报
    # ok:true + 逐文件「无对应引擎/未产出(可能此格式不支持该操作)」,把「你没选格式来源」
    # 说成「你的文件格式不对」。必填是契约属性,该在契约层判,不该漏到执行层去猜。
    for o in op.get("options", []):
        if o.get("required") and not str((opts or {}).get(o["id"], "")).strip():
            return _err("missing_required", f"请先选择「{o['title']}」—— {op['title']} 没有它无法进行")

    # 在枚举目录、读取内容或启动扫描引擎之前，要求单独的明确外发授权。
    # GUI 与 dockit run 共用这一处；终端 doc_dispatch scan / doctools scan-sensitive 的原契约保持不变。
    if CONSENT_OPT in declared and str((opts or {}).get(CONSENT_OPT, "0")).strip().lower() not in _TRUE:
        return _err("consent_required",
                    "敏感词扫描会将所选目录内文件名与内容发送至 Claude；请先明确勾选外发授权"
                    f"(命令行:--opt {CONSENT_OPT}=1)。")

    # 目标单选槽:只收声明过的 id。renum 收到非法范围会产出一份未改动的副本并报成功,
    # bidfinal 会把 argparse 报错说成「有红门」—— 所以在契约层拦,不交给执行层。
    target_ids = [t["id"] for t in op.get("targets", [])]
    if target_ids:
        if not target:
            if op.get("target_required"):
                return _err("missing_target",
                            f"{op['title']}需指定目标(--to):{' / '.join(target_ids)}")
            target = target_ids[0]
        elif target not in target_ids:
            return _err("bad_target",
                        f"{op['title']}不支持目标 {target!r}(可选 {', '.join(target_ids)})")
    elif target:
        return _err("bad_target", f"{op['title']}没有目标可选,不接受 --to")

    if op["kind"] == "dir":
        return _run_scan(op_id, files, opts, scan_findings=scan_findings, dry_run=dry_run)

    existing = [f for f in files if Path(f).exists()]
    missing = [f for f in files if not Path(f).exists()]
    if not existing:
        return _err("no_valid_files", "没有有效文件(全部不存在)")

    # 覆盖门:引擎会把与源同名换后缀的产出(b.md → b.docx)、merged.md 等**直接覆盖且不留备份**。
    # 已存在就停下,要人明确同意 —— 产出只增不改,覆盖是人的决定,不是工具的默认。
    clashes = _overwrite_clashes(op, existing, target)
    if clashes and str((opts or {}).get(OVERWRITE_OPT, "0")).strip().lower() not in _TRUE:
        hit = [p for ps in clashes.values() for p in ps]
        return _err("would_overwrite",
                    f"{op['title']}会覆盖已有文件且不留备份:{', '.join(Path(p).name for p in hit)}。"
                    "要覆盖:勾选「允许覆盖已有同名文件」(命令行 --yes);不想覆盖就先把它移走或改名。",
                    overwrites=hit)

    if dry_run:
        return _plan(op, existing, missing, target, opts, clashes)

    try:
        with _dir_locks(_scan_roots(existing)):
            if op_id == "merge":
                return _run_merge(op, existing, missing)
            return _run_each(op, existing, missing, target, opts, open_viewer)
    except _Busy as b:
        return _busy(b)


def _run_scan(op_id: str, files: list[str], opts: dict | None, *, scan_findings: bool,
              dry_run: bool) -> dict:
    # scan:吃一个目录。能走到这里 = 外发同意已给。
    if not files:
        return _err("no_input", "请选择一个目录")
    d = files[0]
    if not Path(d).is_dir():
        return _err("not_a_dir", f"不是目录: {d}")
    n = _md_files(d)
    row = {"input": d, "name": Path(d).name, "ok": True, "outputs": [], "md_files": n}
    if dry_run:
        row["message"] = (f"将扫描 {n} 个 .md 文件(文件名与正文发送至 Claude)" if n
                          else "目录中没有 .md 文件,不会发送任何内容")
        return {"ok": True, "dry_run": True, "op": op_id, "target": None, "options": dict(opts or {}),
                "results": [row], "succeeded": 1, "total": 1, "log": ""}
    if n == 0:
        # 没东西可扫:不启动引擎,也不外发。findings 给空数组,与「输出解析失败」(null)区分开。
        row["message"] = "目录中没有 .md 文件,未扫描、未发送任何内容"
        if scan_findings:
            row["findings"] = []
        return {"ok": True, "op": op_id, "results": [row], "succeeded": 1, "total": 1, "log": ""}
    roots = [Path(d)]
    try:
        with _dir_locks(roots):
            before = _snapshot(roots)
            rc, log = _run_verb_capture("scan", [d], None, scan_json=scan_findings)
            row["outputs"] = _new_outputs(before, _snapshot(roots), {d})
    except _Busy as b:
        return _busy(b)
    if rc != 0:
        return _err("failed", f"敏感词扫描失败({_log_hint(log) or '详见日志'})", op=op_id, log=log.strip())
    row["message"] = "扫描完成"
    if scan_findings:
        found = _parse_findings(_RUNS[-1]["stdout"]) if _RUNS else None
        row["findings"] = found
        row["message"] = ("扫描完成(未取得结构化结果,详见日志)" if found is None
                          else f"发现 {len(found)} 个可疑敏感词" if found
                          else "未发现新的可疑敏感词")
    return {"ok": True, "op": op_id, "results": [row], "succeeded": 1, "total": 1, "log": log.strip()}


def _run_merge(op: dict, existing: list[str], missing: list[str]) -> dict:
    # merge 是多对一,无法逐文件归因产出 → 整组跑一次,产出整体列出。
    roots = _scan_roots(existing)
    before = _snapshot(roots)
    rc, log = _run_verb_capture("merge", existing, None)
    outs = _new_outputs(before, _snapshot(roots), set(existing))
    ok = rc == 0 and bool(outs)
    results = [{
        "_rc": rc,
        "input": " + ".join(Path(f).name for f in existing),
        "name": f"合并 {len(existing)} 个文件",
        "ok": ok,
        "outputs": outs,
        "message": ("已合并 → " + ", ".join(Path(o).name for o in outs)) if outs
                   else _no_output_message(op, existing[0], None, rc, log),
    }]
    return _wrap(op["id"], results, missing, log)


def _no_output_message(op: dict, f: str, target: str | None, rc: int, log: str) -> str:
    """跑完没有新产出时的说明:区分「格式不支持」与「引擎跑了但没写出东西」(并给日志线索)。"""
    if rc != 0:
        hint = _log_hint(log)
        return "处理失败" + (f"({hint})" if hint else "(详见日志)")
    supported, why = _supports(op, f, target)
    if not supported:
        return f"无对应引擎:{why}"
    hint = _log_hint(log)
    return "引擎没有写出新文件" + (f"(日志:{hint})" if hint else "")


def _run_each(op: dict, existing: list[str], missing: list[str], target: str | None,
              opts: dict | None, open_viewer: bool) -> dict:
    # 其余动词:逐文件跑(才能逐文件归因产出)。
    diagnostic = op.get("produces") == "verdict"
    results = []
    full_log: list[str] = []
    for f in existing:
        # view:HTML 渲染进本文件专用目录并纳入快照,才能报告产出(引擎默认的临时目录在快照外)
        view_dir = Path(tempfile.mkdtemp(prefix="dockit-view-")) if op["id"] == "view" else None
        roots = _scan_roots([f]) + ([view_dir] if view_dir else [])
        before = _snapshot(roots)
        rc, log = _run_verb_capture(op["verb"], [f], target, opts,
                                    view_dir=view_dir, open_viewer=open_viewer)
        full_log.append(log.strip())
        after = _snapshot(roots)
        outs = _new_outputs(before, after, {f})
        # 原地改写(fontunify;clean 处理 pptx):输入自身的指纹变了。
        # 单列出来,agent 才知道「原件已经不是原来那份了」,备份在 outputs 里。
        in_place = [f] if f in before and before.get(f) != after.get(f) else []
        if view_dir and not outs:
            with contextlib.suppress(OSError):
                view_dir.rmdir()
        if diagnostic:
            results.append(_verdict_row(f, outs))
            continue
        ok = rc == 0 and bool(outs or in_place)
        if in_place:
            message = f"原地改写 {Path(f).name}" + (
                (";备份 → " + ", ".join(Path(o).name for o in outs)) if outs else "(没有留备份)")
        elif outs:
            message = "→ " + ", ".join(Path(o).name for o in outs)
        else:
            message = _no_output_message(op, f, target, rc, log)
        row = {"_rc": rc, "input": f, "name": Path(f).name, "ok": ok, "outputs": outs, "message": message}
        if in_place:
            row["modified_in_place"] = in_place
        results.append(row)
    return _wrap(op["id"], results, missing, "\n".join(full_log))


def _wrap(op_id: str, results: list[dict], missing: list[str], log: str) -> dict:
    # ok 曾经写死 True —— 于是「每个文件都失败」也报成功(2026-07-27 修)。
    # 判据:有文件成功 → ok;一个都没成但子进程也没报错(诊断型动词的正常形态)→ 仍 ok;
    # 一个都没成且确有子进程非零退出 → ok:false,并带 error(GUI 只显示 error,缺了就只剩
    # 「后端返回失败（未给出原因）」)。
    _succ = sum(1 for r in results if r["ok"])
    _hard_fail = any(r.get("_rc", 0) != 0 for r in results)
    out = {
        "ok": bool(_succ) or not _hard_fail,
        "op": op_id,
        "results": results,
        "succeeded": _succ,
        "total": len(results),
        "log": log.strip(),
    }
    if missing:
        out["skipped_missing"] = [Path(m).name for m in missing]
    if not out["ok"]:
        tail = _log_hint(log)
        title = _OPS_BY_ID.get(op_id, {}).get("title", op_id)
        out["error"] = (f"{title}:{len(results)} 个输入全部处理失败"
                        + (f"(日志:{tail})" if tail else ""))
        out["error_code"] = "failed"
    for r in results:
        r.pop("_rc", None)          # 内部字段,不进对外信封
    return out


# ───────────────────────────────────────────── doctor(对应 GUI 的「已就绪 / 后端不可达」)

_EXTRA_BIN = ("/opt/homebrew/bin", str(Path.home() / ".local/bin"), "/usr/local/bin")


def _which(name: str) -> str | None:
    """GUI 进程只继承 launchd 的极简 PATH,BackendClient 会前插这几处;这里按同样范围找。"""
    found = shutil.which(name)
    if found:
        return found
    for d in _EXTRA_BIN:
        p = Path(d) / name
        if os.access(p, os.X_OK):
            return str(p)
    return None


def _is_exe(path: str | None) -> bool:
    return bool(path) and os.path.isfile(path) and os.access(path, os.X_OK)


def doctor_checks() -> list[dict]:
    """只做存在/可导入检查:不启动 soffice、不调用 Claude、不读任何用户文档。"""
    checks: list[dict] = []

    def add(cid: str, ok, detail: str, affects: list[str]) -> None:
        checks.append({"id": cid, "ok": bool(ok), "detail": detail, "affects": affects})

    add("python", True, f"{sys.executable} · Python {sys.version.split()[0]}", ["全部操作"])
    add("operations", OPS and len(_OPS_BY_ID) == len(OPS), f"{len(OPS)} 个操作可加载", ["全部操作"])
    libs = [p for p in dd._LIBS if not Path(p).is_dir()]  # noqa: SLF001
    add("engine_libs", not libs, "doctools/lib 与 ~/Dev/tools/dev/lib 可用" if not libs
        else "缺目录: " + ", ".join(libs), ["全部操作"])
    pkgs = {"docx": "python-docx", "pptx": "python-pptx", "openpyxl": "openpyxl",
            "lxml": "lxml", "yaml": "PyYAML"}
    lost = [name for mod, name in pkgs.items() if importlib.util.find_spec(mod) is None]
    add("python_packages", not lost, ("、".join(pkgs.values()) + " 可导入") if not lost
        else "缺: " + ", ".join(lost), ["docx / pptx / xlsx 各操作"])
    uv = _which("uv")
    add("uv", uv, uv or "未找到 uv", ["DocKit.app 经 uv run 启动后端"])
    so = dd.find_soffice()
    add("soffice", so, so or "未找到 LibreOffice(soffice)",
        ["convert / typeset 的老 .doc 升级(缺时退回 textutil,套模板可能失败)"])
    add("textutil", _is_exe("/usr/bin/textutil"), "/usr/bin/textutil", ["convert .doc → txt", ".doc 升级兜底"])
    add("markitdown", _is_exe(dd._MARKITDOWN), dd._MARKITDOWN, ["convert docx / pdf → md"])  # noqa: SLF001
    add("pdftotext", _is_exe(dd._PDFTOTEXT), dd._PDFTOTEXT, ["convert pdf → txt"])  # noqa: SLF001
    if _is_exe(dd.PY_PDF):
        try:
            probe = subprocess.run([dd.PY_PDF, "-c", "import pdfplumber, pypdf"],
                                   capture_output=True, text=True, timeout=60)
            pdf_ok = probe.returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            pdf_ok = False
        add("pdf_python", pdf_ok, f"{dd.PY_PDF}(pdfplumber、pypdf {'可导入' if pdf_ok else '导入失败'})",
            ["convert pdf → word"])
    else:
        add("pdf_python", False, f"{dd.PY_PDF} 不存在", ["convert pdf → word"])
    claude = _which("claude")
    add("claude", claude, claude or "未找到 claude CLI", ["scan(云端敏感词扫描)"])
    return checks


# ───────────────────────────────────────────── 文本输出(agent 默认也能读;--json 给稳定对象)

def _width(s: str) -> int:
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in s)


def _pad(s: str, width: int) -> str:
    return s + " " * max(0, width - _width(s))


def _op_tags(op: dict) -> list[str]:
    tags = []
    if op.get("danger"):
        tags.append("破坏性,需 --yes")
    if CONSENT_OPT in {o["id"] for o in op.get("options", [])}:
        tags.append(f"外发 Claude,需 --opt {CONSENT_OPT}=1")
    if op.get("produces") == "verdict":
        tags.append("只诊断,不写文件")
    if OVERWRITE_OPT in {o["id"] for o in op.get("options", [])}:
        tags.append("覆盖已有同名文件需 --yes")
    return tags


def _example(op: dict) -> str:
    parts = ["dockit", "run", op["id"]]
    targets = op.get("targets", [])
    if targets:
        parts += ["--to", targets[0]["id"]]
    for o in op.get("options", []):
        if o.get("required"):
            parts += ["--opt", f"{o['id']}=<{'/'.join(o.get('exts') or ['文件'])} 路径>"]
        elif o["id"] == CONSENT_OPT:
            parts += ["--opt", f"{CONSENT_OPT}=1"]
    if op.get("danger"):
        parts.append("--yes")
    parts.append("<目录>" if op["kind"] == "dir" else "<文件>...")
    return " ".join(parts)


def ops_text() -> str:
    lines = [f"DocKit 操作 {len(OPS)} 个(dockit ops <操作> 看目标与选项;dockit run <操作> … 执行)"]
    for op in OPS:
        src = "<目录>" if op["kind"] == "dir" else " ".join(op.get("exts", []))
        tags = _op_tags(op)
        lines.append("  " + _pad(op["id"], 12) + _pad(op["title"], 18) + _pad(src, 29) + " "
                     + (f"[{';'.join(tags)}]" if tags else "").rstrip())
    return "\n".join(line.rstrip() for line in lines)


def op_text(op: dict) -> str:
    lines = [f"{op['id']} · {op['title']}", f"  {op['subtitle']}"]
    if op["kind"] == "dir":
        lines.append("  输入:一个目录")
    else:
        lines.append("  输入:文件(可多个) · 源格式 " + " ".join(op.get("exts", [])))
    for tag in _op_tags(op):
        lines.append(f"  注意:{tag}")
    targets = op.get("targets", [])
    if targets:
        rule = "必填" if op.get("target_required") else f"不给则 {targets[0]['id']}"
        lines.append(f"  目标 --to({rule}):"
                     + " | ".join(f"{t['id']}({t['title']})" for t in targets))
    options = op.get("options", [])
    if options:
        lines.append("  选项 --opt K=V(布尔用 1/0;不传 = 保持默认):")
        for g in op.get("option_groups", []):
            items = [o for o in options if o.get("group") == g["id"]]
            if not items:
                continue
            scope = f"(仅 {' '.join(g['applies_to'])})" if g.get("applies_to") else ""
            lines.append(f"    {g['title']}{scope}")
            for o in items:
                if o.get("type") == "file":
                    key = f"{o['id']}=<{'/'.join(o.get('exts') or ['文件'])} 路径>"
                else:
                    key = f"{o['id']}={1 if o.get('default', True) else 0}"
                flag = " [必填]" if o.get("required") else ""
                note = f" — {o['note']}" if o.get("note") else ""
                lines.append(f"      {_pad(key, 26)}{o['title']}{flag}{note}")
    if op.get("aliases"):
        lines.append(f"  别名:{op['aliases']}")
    lines.append(f"  示例:{_example(op)}")
    return "\n".join(lines)


def run_text(payload: dict) -> str:
    lines = []
    if payload.get("dry_run"):
        lines.append("[dry-run] 只做校验,没有执行任何引擎、没有写文件")
    for r in payload.get("results", []):
        verdict = f" [{r['verdict']}]" if r.get("verdict") else ""
        lines.append(f"{'✓' if r.get('ok') else '✗'} {r.get('name')}{verdict}:{r.get('message')}")
        for o in r.get("modified_in_place", []):
            lines.append(f"    ✎ {o}(原地改写)")
        for o in r.get("outputs", []):
            lines.append(f"    → {o}")
        for o in r.get("overwrites", []):
            lines.append(f"    ! 将覆盖 {o}")
        for fnd in r.get("findings") or []:
            where = Path(str(fnd.get("file", ""))).name
            lines.append(f"    · {fnd.get('word')} [{fnd.get('category')}/{fnd.get('severity')}]"
                         f" {fnd.get('reason', '')}" + (f"({where})" if where else ""))
    summary = f"成功 {payload.get('succeeded', 0)}/{payload.get('total', 0)}"
    if payload.get("skipped_missing"):
        summary += " · 跳过不存在 " + ", ".join(payload["skipped_missing"])
    lines.append(summary)
    return "\n".join(lines)


# ───────────────────────────────────────────── CLI

class _Parser(argparse.ArgumentParser):
    """用法错误:命令行带 --json 时同样给 JSON 信封(stdout)+ exit 2。"""

    def error(self, message):
        if "--json" in sys.argv[1:]:
            _emit(_err("usage", f"{self.prog}: {message}"))
            sys.exit(EXIT_USAGE)
        super().error(message)


def _emit(payload: dict) -> None:
    print(json.dumps(payload, ensure_ascii=False))


def exit_code(payload: dict) -> int:
    """run 信封 → 退出码:ok:false 按 error_code 分 2/75/128+N/1;ok:true 但有输入没成或被跳过 → 1。"""
    if not payload.get("ok"):
        code = payload.get("error_code")
        if code in USAGE_ERRORS:
            return EXIT_USAGE
        if code == "interrupted":
            return 128 + int(payload.get("signal") or signal.SIGTERM)
        return EXIT_BUSY if code == "busy" else EXIT_FAILED
    if payload.get("skipped_missing") or any(not r.get("ok") for r in payload.get("results", [])):
        return EXIT_FAILED
    return EXIT_OK


def _interrupted(i: _Interrupted) -> dict:
    return _err("interrupted", f"被信号 {i.signum} 中断,已终止正在运行的引擎(已写出的产出保留在源文件旁)",
                signal=i.signum)


def _fail(a, payload: dict) -> int:
    if getattr(a, "json", False):
        _emit(payload)
    else:
        print(f"dockit: {payload.get('error')}", file=sys.stderr)
        if getattr(a, "verbose", False) and payload.get("log"):
            print(payload["log"], file=sys.stderr)
    return exit_code(payload)


def _parse_opts(pairs: list[str]) -> tuple[dict, list[str]]:
    bad = [x for x in pairs if "=" not in x]
    opts = {}
    for kv in pairs:
        if "=" in kv:
            k, v = kv.split("=", 1)
            opts[k.strip()] = v.strip()
    return opts, bad


def cmd_ops(a) -> int:
    if a.op_id:
        op = _OPS_BY_ID.get(a.op_id)
        if op is None:
            return _fail(a, _err("unknown_op", f"未知操作: {a.op_id}(支持 {', '.join(_OPS_BY_ID)})"))
        if a.json:
            _emit({"ok": True, "op": op})
        else:
            print(op_text(op))
        return EXIT_OK
    if a.json:
        _emit({"ok": True, "ops": OPS})
    else:
        print(ops_text())
    return EXIT_OK


def cmd_run(a) -> int:
    opts, bad = _parse_opts(a.opt)
    if bad:
        return _fail(a, _err("bad_opt_form", f"--opt 需要 K=V 形式: {', '.join(bad)}"))
    op = _OPS_BY_ID.get(a.op_id)
    if a.open and a.op_id != "view":
        return _fail(a, _err("usage", "--open 只用于 view(在浏览器打开预览)"))
    if a.background and a.dry_run:
        return _fail(a, _err("usage", "--background 与 --dry-run 不能同用(dry-run 本身立即返回)"))
    # GUI 的「破坏性」红标 + 人点执行 = 命令行的 --yes;dry-run 不执行,不需要确认。
    if op and op.get("danger") and not (a.yes or a.dry_run):
        return _fail(a, _err("confirm_required",
                             f"「{op['title']}」是破坏性操作:{op['subtitle']}。确认后加 --yes;"
                             "只想看会处理哪些文件先用 --dry-run"))
    # --yes 也是「允许覆盖已有同名文件」的确认(GUI 里是一个勾选项);显式 --opt 优先。
    if a.yes and op and OVERWRITE_OPT in {o["id"] for o in op.get("options", [])}:
        opts.setdefault(OVERWRITE_OPT, "1")
    for o in (op or {}).get("options", []):
        if o.get("type") == "file" and str(opts.get(o["id"], "")).strip():
            opts[o["id"]] = os.path.abspath(os.path.expanduser(opts[o["id"]]))
    paths = [os.path.abspath(os.path.expanduser(p)) for p in a.paths]
    if a.background:
        return _start_job(a, opts, paths)
    payload = gui_run(a.op_id, paths, a.target, opts, open_viewer=a.open,
                      scan_findings=True, dry_run=a.dry_run)
    if not payload.get("ok"):
        return _fail(a, payload)
    if a.json:
        _emit(payload)
    else:
        print(run_text(payload))
        if a.verbose and payload.get("log"):
            print(payload["log"], file=sys.stderr)
    return exit_code(payload)


# ───────────────────────────────────────────── 后台任务(run --background / status)
#
# typeset、soffice 冷启动、整批 convert 可能跑好几分钟,比 agent 单次工具调用的超时还长。
# --background 先在前台走完全部校验(与 --dry-run 同一条 gui_run),再起一个脱离会话的
# `dockit run … --json` 子进程,立刻返回 job id。子进程在跑时一直持有 <id>.lock 的 flock
# (从父进程经 pass_fds 继承),status 用非阻塞 flock 判断「还在跑」,不靠 pid(会被复用)。
# 结果就是子进程 stdout 那一个 JSON 信封;退出码由信封按 exit_code() 还原。

JOB_KEEP_DAYS = 7
_JOB_ID = re.compile(r"\d{8}-\d{6}-[0-9a-f]{6}")


def _job_files(job_id: str) -> dict[str, Path]:
    return {ext: JOB_DIR / f"{job_id}.{ext}" for ext in ("json", "out", "err", "lock")}


def _job_running(lock: Path) -> bool:
    try:
        fd = os.open(lock, os.O_RDONLY)
    except FileNotFoundError:
        return False
    try:
        fcntl.flock(fd, fcntl.LOCK_SH | fcntl.LOCK_NB)
        return False
    except BlockingIOError:
        return True
    finally:
        os.close(fd)


def _prune_jobs() -> None:
    cutoff = time.time() - JOB_KEEP_DAYS * 86400
    for meta in JOB_DIR.glob("*.json"):
        with contextlib.suppress(OSError):
            if meta.stat().st_mtime > cutoff:
                continue
            files = _job_files(meta.stem)
            if _job_running(files["lock"]):
                continue
            for p in files.values():
                with contextlib.suppress(FileNotFoundError):
                    p.unlink()


def _start_job(a, opts: dict, paths: list[str]) -> int:
    plan = gui_run(a.op_id, paths, a.target, opts, open_viewer=a.open, dry_run=True)
    if not plan.get("ok"):
        return _fail(a, plan)
    JOB_DIR.mkdir(parents=True, exist_ok=True)
    _prune_jobs()
    job_id = time.strftime("%Y%m%d-%H%M%S") + "-" + secrets.token_hex(3)
    files = _job_files(job_id)
    argv = ["run", a.op_id, *(["--to", a.target] if a.target else []),
            *[x for k, v in opts.items() for x in ("--opt", f"{k}={v}")],
            *(["--yes"] if a.yes else []), *(["--open"] if a.open else []), "--json", *paths]
    lock_fd = os.open(files["lock"], os.O_RDWR | os.O_CREAT, 0o600)
    try:
        fcntl.flock(lock_fd, fcntl.LOCK_EX)
        with open(files["out"], "w", encoding="utf-8") as out, \
                open(files["err"], "w", encoding="utf-8") as err:
            proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), *argv],
                                    stdin=subprocess.DEVNULL, stdout=out, stderr=err,
                                    start_new_session=True, pass_fds=(lock_fd,))
    finally:
        os.close(lock_fd)   # 子进程继承的那份一直持有锁,直到它退出
    job = {"id": job_id, "op": a.op_id, "target": a.target, "inputs": paths, "pid": proc.pid,
           "argv": ["dockit", *argv], "started": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
           "status": "running"}
    tmp = JOB_DIR / f".{job_id}.json.tmp"
    tmp.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    tmp.replace(files["json"])
    payload = {"ok": True, "job": job, "status_cmd": f"dockit status {job_id}"}
    if a.json:
        _emit(payload)
    else:
        print(f"已在后台启动 {a.op_id}:job {job_id}(pid {proc.pid})。查询:dockit status {job_id}")
    return EXIT_OK


def job_state(job_id: str) -> dict | None:
    files = _job_files(job_id)
    if not _JOB_ID.fullmatch(job_id) or not files["json"].is_file():
        return None
    job = json.loads(files["json"].read_text(encoding="utf-8"))
    if _job_running(files["lock"]):
        job["status"] = "running"
        return job
    result = None
    with contextlib.suppress(OSError, ValueError):
        value = json.loads(files["out"].read_text(encoding="utf-8").strip() or "null")
        result = value if isinstance(value, dict) and isinstance(value.get("ok"), bool) else None
    if result is None:
        job["status"] = "lost"
        job["exit_code"] = EXIT_FAILED
        job["error"] = "进程已结束,但没有留下结果(可能被强制结束)"
        with contextlib.suppress(OSError):
            tail = [ln for ln in files["err"].read_text(encoding="utf-8").splitlines() if ln.strip()]
            if tail:
                job["error"] += f";stderr 末行:{tail[-1][:300]}"
        return job
    job["status"] = "interrupted" if result.get("error_code") == "interrupted" else "done"
    job["exit_code"] = exit_code(result)
    job["finished"] = time.strftime("%Y-%m-%dT%H:%M:%S%z", time.localtime(files["out"].stat().st_mtime))
    job["result"] = result
    return job


def cmd_status(a) -> int:
    if a.job_id:
        job = job_state(a.job_id)
        if job is None:
            return _fail(a, _err("unknown_job", f"没有这个后台任务: {a.job_id}(dockit status 列出最近的)"))
        if a.json:
            _emit({"ok": True, "job": job})
        else:
            print(f"job {job['id']} · {job['op']} · {job['status']}"
                  + (f" · 退出码 {job['exit_code']}" if "exit_code" in job else "")
                  + (f" · pid {job['pid']}" if job["status"] == "running" else ""))
            res = job.get("result")
            if res:
                print(run_text(res) if res.get("ok") else f"dockit: {res.get('error')}")
            elif job.get("error"):
                print(job["error"])
        return EXIT_OK if job["status"] == "running" else job["exit_code"]
    jobs = []
    if JOB_DIR.is_dir():
        for meta in sorted(JOB_DIR.glob("*.json"), reverse=True)[:20]:
            job = job_state(meta.stem)
            if job:
                res = job.pop("result", None) or {}
                job.update({k: res[k] for k in ("succeeded", "total", "error_code") if k in res})
                jobs.append(job)
    if a.json:
        _emit({"ok": True, "jobs": jobs})
    else:
        for j in jobs:
            print(f"{j['id']}  {_pad(j['op'], 12)}{_pad(j['status'], 12)}"
                  + (f"退出码 {j['exit_code']}" if "exit_code" in j else ""))
        if not jobs:
            print(f"没有后台任务(保留最近 {JOB_KEEP_DAYS} 天)")
    return EXIT_OK


def cmd_doctor(a) -> int:
    checks = doctor_checks()
    ready = sum(1 for c in checks if c["ok"])
    payload = {"ok": ready == len(checks), "ready": ready, "total": len(checks), "checks": checks}
    if a.json:
        _emit(payload)
    else:
        for c in checks:
            line = f"{'✓' if c['ok'] else '✗'} {_pad(c['id'], 16)}{c['detail']}"
            if not c["ok"]:
                line += f"(影响:{'、'.join(c['affects'])})"
            print(line)
        print(f"就绪 {ready}/{len(checks)}")
    return EXIT_OK if payload["ok"] else EXIT_FAILED


def _gui_main(a) -> int:
    try:
        if a.cmd == "gui-ops":
            payload = {"ok": True, "ops": OPS}
        else:
            _opts, _bad_form = _parse_opts(a.opt)
            if _bad_form:
                payload = _err("bad_opt_form", f"--opt 需要 K=V 形式: {', '.join(_bad_form)}")
            else:
                payload = gui_run(a.op, a.files, a.target, _opts)
    except _Interrupted as i:
        payload = _interrupted(i)
    except Exception as e:  # noqa: BLE001 — 任何异常都转人话信封,绝不让 Swift 见 traceback
        payload = _err("backend_exception", f"后端异常: {type(e).__name__}: {e}")
    print(json.dumps(payload, ensure_ascii=False))
    return 0  # gui-* 一律 exit 0(信封承载成败)


_DESCRIPTION = """DocKit 命令行:与 DocKit.app 同一个业务层(操作目录、校验、外发同意门、覆盖门、产出归因都在这里)。
GUI 给人用,dockit 给 agent 用;拖拽、⌘K 面板、在 Finder 中显示这类纯界面动作只在 App 里。"""

_EPILOG = """示例:
  dockit ops                                   列出 14 个操作
  dockit ops convert --json                    convert 的目标与选项(稳定 JSON)
  dockit run clean --opt scope.table=0 --json /abs/报告.docx
  dockit run convert --to md /abs/a.docx /abs/b.pdf
  dockit run convert --to word --dry-run /abs/b.md   先看会不会覆盖已有的 b.docx
  dockit run formatclone --opt ref=/abs/范式.docx /abs/草稿.md
  dockit run fontunify --dry-run /abs/汇报.pptx     破坏性操作先看计划,确认后 --yes
  dockit run typeset --background --json /abs/报告.md   长任务:立刻返回 job id
  dockit status <job> --json                   查询后台任务;结束后附完整结果信封
  dockit run scan --opt privacy.cloud_consent=1 /abs/标书目录   外发 Claude,须本人同意后才加
  dockit doctor

退出码:0 成功 · 1 业务失败(有输入失败/被跳过、门检有红门、doctor 有缺项)
        2 用法或校验错误(未知操作/选项/目标、缺必填、缺外发同意、破坏性操作缺 --yes、
          会覆盖已有同名文件却没给 --yes、没有有效文件)
        75 同一目录树(含上级/下级目录)正有另一个 DocKit 操作在运行,稍后重试
        128+N 被信号 N 中断(引擎进程组一并终止) · 69 包装脚本找不到 ~/Dev/.venv 或后端
--json 信封里 ok 是请求级(请求被处理即 true);逐个输入看 results[].ok,整体看 all_ok 或退出码。
产出写在源文件旁:_fixed / _styled / _序号修正 / _lower 等 DocKit 后缀的产出重跑时覆盖上一次;
与源同名换后缀(b.md → b.docx)、merged.md、按 sheet 命名的产出已存在时默认拒绝(would_overwrite),
确认覆盖加 --yes。fontunify 与 pptx 的 clean 原地改写(已有 .backup 时另存编号备份,不顶掉原件)。
run 默认同步:typeset 或 soffice 冷启动可能要数分钟,给命令留 10 分钟超时,或用 --background。
被 SIGTERM/SIGINT 结束时会连带终止引擎子进程,不会留下无人看管的写盘进程。"""


def build_parser() -> argparse.ArgumentParser:
    fmt = argparse.RawDescriptionHelpFormatter
    ap = _Parser(prog="dockit", description=_DESCRIPTION, epilog=_EPILOG, formatter_class=fmt)
    sub = ap.add_subparsers(dest="cmd", required=True, metavar="<command>")

    p = sub.add_parser("ops", help="列出操作;给出 <op> 时显示它的目标、选项和默认值",
                       description="列出 GUI 侧栏里的全部操作,或单个操作的输入、目标、选项、默认值与示例。")
    p.add_argument("op_id", nargs="?", metavar="op", help="操作 id,如 clean / convert / scan")
    p.add_argument("--json", action="store_true", help='输出 JSON:{"ok":true,"ops":[...]} 或 {"ok":true,"op":{...}}')

    p = sub.add_parser("run", help="执行一个操作(与 GUI「执行」同一条路径)",
                       description="执行一个操作,返回逐文件结果与产出路径。"
                                   "校验、外发同意门、覆盖门、产出归因与 GUI 完全相同。",
                       epilog="退出码:0 全部成功 · 1 有输入失败/被跳过/门检红门 · 2 用法或校验错误(含 would_overwrite)"
                              " · 75 目录忙 · 128+N 被信号中断\n"
                              "--json 的 ok 是请求级;逐个输入看 results[].ok,整体看 all_ok 或退出码。",
                       formatter_class=fmt)
    p.add_argument("op_id", metavar="op", help="操作 id(dockit ops 查看)")
    p.add_argument("paths", nargs="+", metavar="path", help="输入文件;scan 为一个目录")
    p.add_argument("--to", dest="target", default=None, metavar="T",
                   help="目标单选槽:convert 目标格式(必填)/ renum 范围 / bidfinal 模式")
    p.add_argument("--opt", action="append", default=[], metavar="K=V",
                   help="操作选项,可重复;可用键见 dockit ops <op>")
    p.add_argument("--dry-run", action="store_true",
                   help="只做校验并列出将处理的输入与将覆盖的已有文件,不执行、不写盘")
    p.add_argument("--yes", action="store_true",
                   help="确认破坏性操作(fontunify / lowercase / stripchrome)与覆盖已有同名文件")
    p.add_argument("--open", action="store_true", help="view 专用:渲染后在浏览器打开(默认不打开)")
    p.add_argument("--background", action="store_true",
                   help="校验通过后在后台运行,立刻返回 job id;用 dockit status <id> 查询")
    p.add_argument("--json", action="store_true", help="输出与 GUI 相同的 JSON 结果信封")
    p.add_argument("-v", "--verbose", action="store_true", help="文本模式下把引擎日志打到 stderr")

    p = sub.add_parser("status", help="查询后台任务(run --background)的状态与结果",
                       description="不带参数列出最近 20 个后台任务;带 job id 给出该任务状态"
                                   "(running / done / interrupted / lost),结束后附完整结果信封。",
                       epilog="退出码:running 为 0;结束后与该任务同步运行时的退出码相同;lost 为 1;未知 id 为 2。",
                       formatter_class=fmt)
    p.add_argument("job_id", nargs="?", metavar="job", help="run --background 返回的 id")
    p.add_argument("--json", action="store_true",
                   help='输出 JSON:{"ok":true,"job":{...}} 或 {"ok":true,"jobs":[...]}')

    p = sub.add_parser("doctor", help="检查依赖是否就绪(venv、uv、soffice、markitdown、pdftotext、claude 等)",
                       description="只做存在与可导入检查,不启动转换、不联网、不读文档。有缺项 exit 1。")
    p.add_argument("--json", action="store_true", help='输出 JSON:{"ok","ready","total","checks":[...]}')

    sub.add_parser("gui-ops", help="GUI 内部契约:{ok, ops} JSON,总是 exit 0")
    p = sub.add_parser("gui-run", help="GUI 内部契约:成败只在 JSON 信封里,总是 exit 0(agent 用 run)")
    p.add_argument("--op", required=True)
    p.add_argument("--to", dest="target", default=None)
    p.add_argument("--opt", action="append", default=[], metavar="K=V",
                   help="per-op 选项,可重复(见 gui-ops 的 options 声明)")
    p.add_argument("--files", nargs="+", required=True)
    return ap


def main() -> int:
    a = build_parser().parse_args()
    if a.cmd in ("gui-run", "run"):
        _install_signal_handlers()
    if a.cmd in ("gui-ops", "gui-run"):
        return _gui_main(a)
    try:
        if a.cmd == "ops":
            return cmd_ops(a)
        if a.cmd == "run":
            return cmd_run(a)
        if a.cmd == "status":
            return cmd_status(a)
        if a.cmd == "doctor":
            return cmd_doctor(a)
    except _Interrupted as i:
        return _fail(a, _interrupted(i))
    except Exception as e:  # noqa: BLE001 — agent 拿信封与退出码,不拿 traceback
        return _fail(a, _err("backend_exception", f"后端异常: {type(e).__name__}: {e}"))
    return EXIT_USAGE


if __name__ == "__main__":
    sys.exit(main())
