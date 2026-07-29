#!/usr/bin/env python3
"""blast_radius.py — 量一个 docx 工具的「炸开面」:跑完之后包里有多少部件被改写。

## 为什么有这个 (2026-07-28 立)

python-docx 与 surgical(lxml+zipfile) 的真实差别,**光数 zip 条目数看不出来** ——
实测一份 301 条目 / 75 个嵌入公式的真报告:

    python-docx 打开即存回(什么都没改)  -> 条目 301 一个没少, 但 **60 个部件被重写**
                                          (含 [Content_Types].xml / _rels/.rels /
                                           28 个 header + 14 个 footer)
    surgical 解析再写回                  -> **1 个部件**被重写(就是点名那个)

条目数相同会让人以为「没损失」,于是把 python-docx 判成安全 —— 那是拿错尺子量。
真正该量的是**这次改动到底碰了多少个不该碰的文件**:改一个字号重写 60 个部件,
任何一个的命名空间声明/属性序/空白差异都可能让 Word 渲染不同。

灾难性的那一档另算:`Document()` 新建再搬段落 -> 301 条目掉到 17、75 个公式对象归零。

## 用法

    # 量一条命令的炸开面
    blast_radius.py run <fixture.docx> -- python3 sub/normalize_fonts.py {docx}
    # {docx} 会被替换成工作副本路径;原夹具永不被改

    # 对拍:同一夹具跑两条命令(老实现 vs 新实现),比较炸开面与正文等价性
    blast_radius.py diff <fixture.docx> \
        --old "python3 sub/normalize_fonts.py {docx}" \
        --new "python3 sub/normalize_fonts.py {docx}"

退出码: 0 通过 / 1 有差异(diff 模式下 = 迁移不等价) / 2 真故障(夹具不存在/命令跑不起来)
"""

from __future__ import annotations

import argparse
import hashlib
import shlex
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# 这些部件被改写是「必然且无害」的:改正文本来就要动 document.xml。
# 其余任何一个被改写都要报出来 —— 它是这次改动的**意外**炸开面。
EXPECTED = {"word/document.xml"}


def part_hashes(docx: Path) -> dict[str, str]:
    with zipfile.ZipFile(docx) as z:
        return {n: hashlib.sha1(z.read(n)).hexdigest() for n in z.namelist()}


def run_on_copy(fixture: Path, cmd: str | list[str], workdir: Path) -> tuple[Path, int, str]:
    """把夹具复制一份、在副本上跑命令。返回 (副本路径, 退出码, 输出)。原件不动。

    cmd 是 list → 不过 shell 直接传 argv。`run` 子命令拿到的 REMAINDER 本来就已经被
    shell 切好了,再 " ".join() 回去会把 `-c "import docx; ..."` 里的引号丢光
    (首版实测:sh 报 syntax error near unexpected token `(')。
    cmd 是 str → 过 shell(diff 子命令的 --old/--new 是整条命令串)。
    """
    work = workdir / fixture.name
    shutil.copy2(fixture, work)
    if isinstance(cmd, list):
        argv = [c.replace("{docx}", str(work)) for c in cmd]
        p = subprocess.run(argv, capture_output=True, text=True)
    else:
        p = subprocess.run(cmd.replace("{docx}", shlex.quote(str(work))),
                           shell=True, capture_output=True, text=True)
    return work, p.returncode, (p.stdout + p.stderr)[-2000:]


def radius(before: dict[str, str], after: dict[str, str]) -> dict[str, list[str]]:
    gone = sorted(set(before) - set(after))
    added = sorted(set(after) - set(before))
    changed = sorted(n for n in before if n in after and before[n] != after[n])
    return {"删掉": gone, "新增": added, "改写": changed}


def report(tag: str, before: dict, after: dict) -> dict:
    r = radius(before, after)
    unexpected = [n for n in r["改写"] if n not in EXPECTED]
    print(f"\n[{tag}] 原包 {len(before)} 个部件 → 现在 {len(after)} 个")
    print(f"  删掉 {len(r['删掉'])} · 新增 {len(r['新增'])} · 改写 {len(r['改写'])}"
          f"（其中意料之外 {len(unexpected)}）")
    for k in ("删掉", "新增"):
        for n in r[k][:8]:
            print(f"    {k}: {n}")
    for n in unexpected[:12]:
        print(f"    意外改写: {n}")
    if len(unexpected) > 12:
        print(f"    …另有 {len(unexpected) - 12} 个")
    return {**r, "意外": unexpected}


def cmd_run(a) -> int:
    fx = Path(a.fixture).expanduser().resolve()
    if not fx.is_file():
        print(f"夹具不存在: {fx}", file=sys.stderr)
        return 2
    before = part_hashes(fx)
    with tempfile.TemporaryDirectory() as td:
        work, rc, out = run_on_copy(fx, list(a.cmd), Path(td))
        if rc != 0:
            print(f"命令退出码 {rc}:\n{out}", file=sys.stderr)
            return 2
        r = report("炸开面", before, part_hashes(work))
    # 删/增部件一律判红;意外改写只报数不判红(某些工具确实该动 styles.xml)
    return 1 if (r["删掉"] or r["新增"]) else 0


def cmd_diff(a) -> int:
    fx = Path(a.fixture).expanduser().resolve()
    if not fx.is_file():
        print(f"夹具不存在: {fx}", file=sys.stderr)
        return 2
    before = part_hashes(fx)
    with tempfile.TemporaryDirectory() as td:
        wd = Path(td)
        (wd / "old").mkdir(); (wd / "new").mkdir()
        o, orc, oout = run_on_copy(fx, a.old, wd / "old")
        n, nrc, nout = run_on_copy(fx, a.new, wd / "new")
        if orc != 0 or nrc != 0:
            print(f"老实现 rc={orc} 新实现 rc={nrc}\n--- old ---\n{oout}\n--- new ---\n{nout}",
                  file=sys.stderr)
            return 2
        ha, hb = part_hashes(o), part_hashes(n)
        ro = report("老实现", before, ha)
        rn = report("新实现", before, hb)

        # 正文等价性:两条实现改完之后 document.xml 该是同一份
        same_doc = ha.get("word/document.xml") == hb.get("word/document.xml")
        print(f"\n正文 document.xml 两边一致: {'是' if same_doc else '否'}")
        shrink = len(ro["意外"]) - len(rn["意外"])
        print(f"意外炸开面: 老 {len(ro['意外'])} → 新 {len(rn['意外'])}"
              f"（{'收窄 %d' % shrink if shrink > 0 else '没有收窄'}）")

        bad = []
        if not same_doc:
            bad.append("正文结果不一致 —— 迁移改变了行为，不是等价替换")
        if rn["删掉"] or rn["新增"]:
            bad.append(f"新实现动了部件集合：删 {len(rn['删掉'])} 增 {len(rn['新增'])}")
        if shrink <= 0:
            bad.append("新实现没有收窄炸开面 —— 迁移没带来收益，先查是不是没真走 surgical")
        for b in bad:
            print(f"  ✗ {b}")
        if not bad:
            print("  ✓ 等价且炸开面收窄")
        return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser(description="量 docx 工具的炸开面 / 迁移前后对拍")
    # dest 必须叫 mode 不能叫 cmd:run 子命令的 REMAINDER 也叫 cmd,同名会互相覆盖,
    # 于是 a.cmd 拿到的是命令列表、`a.cmd == "run"` 永远为假(首版实测)。
    sub = ap.add_subparsers(dest="mode", required=True)
    r = sub.add_parser("run", help="量一条命令的炸开面")
    r.add_argument("fixture")
    r.add_argument("cmd", nargs=argparse.REMAINDER,
                   help="-- 之后是命令，{docx} 会被替换成工作副本路径")
    d = sub.add_parser("diff", help="同一夹具跑老/新两条实现并对拍")
    d.add_argument("fixture")
    d.add_argument("--old", required=True)
    d.add_argument("--new", required=True)
    a = ap.parse_args()
    if a.mode == "run":
        while a.cmd and a.cmd[0] == "--":
            a.cmd.pop(0)
        if not a.cmd:
            print("没给命令 —— 拒绝在空命令上报通过", file=sys.stderr)
            return 2
        return cmd_run(a)
    return cmd_diff(a)


if __name__ == "__main__":
    sys.exit(main())
