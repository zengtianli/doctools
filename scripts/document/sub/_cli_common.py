#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_cli_common.py — strip_*/audit_* 家族 main() 样板 SSOT (P4, 2026-07-31)

这里只收「机制」: lsof 占用检查 / .bak-N-YYYY-MM-DD 备份 / 标准 flag 声明 /
report JSON 落盘 / stdout JSON 打印。**不收**各脚本的报错文案与退出码 ——
家族里 missing-file 有 return 1 / return 2 / sys.exit(1) / sys.exit(2) 四种,
busy 有 exit(2) / return 3 两种, 统一它们 = 行为变更, 须另立任务用户拍板。

语义抄 canonical 实现 pipeline_lib.lsof_check / make_backup_path (不 import
pipeline_lib —— 它会拖起 docx / dataclass 等重依赖, 本模块保持零副作用)。

铁则: 本模块禁 import docx、禁任何 import 期写盘副作用
(check_docx_collar 枚举全仓, 惊动它 = 假红)。

被三种加载模式消费, 脚本侧 import 姿势必须带 self-dir append
(docx_cli → _dispatch._load 用 spec_from_file_location, 不把 sub/ 放 sys.path):

    import sys as _sys
    from pathlib import Path as _Path
    _sys.path.append(str(_Path(__file__).resolve().parent))
    import _cli_common as _cc  # noqa: E402
"""
from __future__ import annotations

import datetime
import json
import shutil
import subprocess
from pathlib import Path


# ---------------- lsof / backup ----------------

def lsof_check(p: Path) -> str | None:
    """被占用返回 lsof stdout, 空闲返回 None。

    5s timeout; lsof 不存在 / 超时一律当空闲 (返回 None)。
    报错文案与退出方式留在调用处 (各脚本不一致, 是契约不是债)。
    """
    try:
        r = subprocess.run(
            ["lsof", str(p)], capture_output=True, text=True, timeout=5
        )
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def find_next_backup(p: Path) -> Path:
    """<stem>.bak-N-YYYY-MM-DD<suffix>, N 自增。只算路径不 copy。"""
    today = datetime.date.today().isoformat()
    n = 1
    while True:
        cand = p.with_name(f"{p.stem}.bak-{n}-{today}{p.suffix}")
        if not cand.exists():
            return cand
        n += 1


def make_backup(p: Path) -> Path:
    """find_next_backup + copy2, 返回 bak 路径。"""
    b = find_next_backup(p)
    shutil.copy2(p, b)
    return b


# ---------------- 标准 flag 声明 ----------------

def _kw(help_str):
    # help=None 时不传 help 参数 (帮助行只显示 flag 本身, 与现状一致)
    return {} if help_str is None else {"help": help_str}


def add_write_flags(ap, *, dry_run_help=None, no_backup_help=None,
                    report_help=None) -> None:
    """strip 家族标准写盘三件套: --dry-run / --no-backup / --report。"""
    ap.add_argument("--dry-run", action="store_true", **_kw(dry_run_help))
    ap.add_argument("--no-backup", action="store_true", **_kw(no_backup_help))
    ap.add_argument("--report", type=Path, default=None, **_kw(report_help))


def add_report_flag(ap, *, help=None) -> None:  # noqa: A002 - 契约签名
    """audit 家族只加 --report。"""
    ap.add_argument("--report", type=Path, default=None, **_kw(help))


# ---------------- report 落盘 / stdout ----------------

def write_report(report: dict, path, *, mkdir: bool = True,
                 announce: str | None = None) -> None:
    """path 为 None/falsy 直接返回。mkdir=False 保真「父目录不存在时报错」派。

    announce 非 None 时 print(announce.format(path=path)) —— 逐字传原打印串,
    如 "[INFO] report -> {path}" / "[report] {path}"。
    """
    if not path:
        return
    p = Path(path)
    if mkdir:
        p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                 encoding="utf-8")
    if announce is not None:
        print(announce.format(path=path))


def print_json(obj) -> None:
    print(json.dumps(obj, ensure_ascii=False, indent=2))
