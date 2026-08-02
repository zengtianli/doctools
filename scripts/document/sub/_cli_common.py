#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_cli_common.py — strip_*/audit_* 家族 main() 样板 SSOT (P4, 2026-07-31)

`family_main()` (2026-08-02 下沉) 收的是 12 个家族文件那段逐字相同的子命令
分发器本体; 以下几段说的是它之外的「机制」层。

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
import sys
from pathlib import Path


# ---------------- 家族 main() 分发器 ----------------

def family_main(subcommands, argv=None, *, file, usage_args="<args…>") -> int:
    """12 个 sub/ 家族文件 (audit/biddiff/blocks/caption/chapter/freeze/image/
    renumber/split/strip/table/typeset_ops) 逐字相同的那段 19 行分发器。

    调用姿势 (家族文件末尾)::

        def main(argv: list[str] | None = None) -> int:
            return _cc.family_main(SUBCOMMANDS, argv, file=__file__,
                                   usage_args="<docx> [flags…]")

    行为逐字保真, 三件事必须知道:

    ① **本次唯一扩大的攻击面**: 为了裸 `import _cli_common`, 每个家族文件都要把
       `sub/` 自己 append 进 sys.path。原来只有 7 个文件这么干, 现在 12 个都干 ——
       `sub/` 被更多模块推进 sys.path, 而 `lib/styles.py` 与 `sub/styles.py` **同名**。
       今天安全: `grep -rn '^ *import styles\\|^ *from styles' scripts/ lib/ tools/`
       零命中, 三处要 lib/styles 的都走 `spec_from_file_location` 按绝对路径加载,
       不吃 sys.path。**将来谁写下裸 `import styles` 就会中招** —— 一律 append
       (不是 insert(0)), 让 sys.path[0] = 脚本自身目录这条规则继续兜底。

    ② 族名从 `file` (传 `__file__`) 的 `Path(...).stem` 派生, usage 串与错误前缀
       都用它。**文件改名 = usage 串跟着变**, 那是 `--help` 输出的一部分;
       家族文件重命名前先想清楚外部有没有人在 grep 这行。

    ③ `_cli_common` 铁则照旧: 禁 import docx、禁任何 import 期副作用。本函数只碰
       `sys.argv` / `sys.stderr`, 不写盘、不引重依赖。

    `usage_args` 尾巴当前只有两种取值 —— audit / freeze / strip 传
    `"<docx> [flags…]"`, 其余 9 个用默认 `"<args…>"`。**弄反了 `--help` 输出就变,
    而 cli_surface / cli_forward_probe / collar / axis 全部照样绿**
    (那几道门看 argv 与接口形状, 进不到家族 main), 只有
    `handoffs/_loc_plan_harness/family_ab.py` 的 144 例对拍抓得住。
    """
    fam = Path(file).stem
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in ("-h", "--help"):
        print(f"usage: {fam}.py {{" + ",".join(subcommands) + f"}} {usage_args}\n"
              f"每个子命令的参数与原独立脚本逐字一致：{fam}.py <sub> --help 查看。")
        return 0 if args else 2
    sub, rest = args[0], args[1:]
    fn = subcommands.get(sub)
    if fn is None:
        print(f"[{fam}] unknown subcommand: {sub!r}; choices={list(subcommands)}",
              file=sys.stderr)
        return 2
    saved = sys.argv[:]
    sys.argv = [sys.argv[0]] + rest
    try:
        rc = fn()
        return int(rc) if isinstance(rc, int) else 0
    finally:
        sys.argv = saved


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
