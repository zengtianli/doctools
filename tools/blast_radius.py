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
        3 没测到(老实现在本夹具上一个部件都没改) —— 既不算通过也不算判红
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

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
from docx_surgical import canonical  # noqa: E402

# 这些部件被改写是「必然且无害」的:改正文本来就要动 document.xml。
# 其余任何一个被改写都要报出来 —— 它是这次改动的**意外**炸开面。
EXPECTED = {"word/document.xml"}


def read_doc_xml(docx: Path) -> bytes:
    """读 word/document.xml。缺了直接退出,不返回 b"" —— 两份都缺时空串会互相相等,
    于是「正文一致」报绿,那是假绿里最难发现的一种。"""
    with zipfile.ZipFile(docx) as z:
        try:
            return z.read("word/document.xml")
        except KeyError:
            raise SystemExit(f"⛔ {docx} 里没有 word/document.xml —— 不是可用的 Word 文件")


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

        # 正文等价性必须比**规范化后**的 XML，不能比 sha1(2026-07-30 修)。
        # 拿 sha1 比会得出「迁移改变了行为」的假红:新实现在正文语义没变时会把原件
        # 字节整个还原(那正是它的价值),而老实现留下的是 python-docx 重新序列化过的
        # 同一份内容 —— 声明写成 <?xml version='1.0'?> 而不是 "1.0"、换行 \n 而不是
        # \r\n,字节就不同了。实测 number_captions / convert_chapter_format 双双因此
        # 被判红,而 canonical 比对两边完全一致。
        # 这跟「数 zip 条目数」是同一类错误:拿看得见的尺子量,量的不是要判的那件事。
        same_doc = canonical(read_doc_xml(o)) == canonical(read_doc_xml(n))
        print(f"\n正文 document.xml 两边语义一致（C14N 比对）: {'是' if same_doc else '否'}")
        shrink = len(ro["意外"]) - len(rn["意外"])
        print(f"意外炸开面: 老 {len(ro['意外'])} → 新 {len(rn['意外'])}"
              f"（{'收窄 %d' % shrink if shrink > 0 else '没有收窄'}）")

        # 老实现一个部件都没动 = 这条命令在本夹具上什么都没干(没找到可改的东西,或
        # 另存到别的文件去了)。此时「炸开面没收窄」是必然的,判红是假红;但也**不能报
        # 通过** —— 什么都没测到就宣布等价是假绿。单独一档 rc=3。
        if not (ro["删掉"] or ro["新增"] or ro["改写"]):
            print("  ⊘ 老实现在本夹具上一个部件都没改 —— 闸门没测到东西，不算通过。"
                  "换一份能触发它的夹具，或给它该有的参数（--plan/--decision…）")
            return 3

        # ── 判据的那根轴（2026-07-30 立，被三个 bug 连着教会的）──────────────────
        # 本子命令叫 diff，它判的是**等价迁移**，所以每一条判据都必须是「老 vs 新」的
        # 相对比较，**不能是「新 vs 我以为的理想值」**。三个 bug 全是同一个形状：
        #   ① 正文用 sha1 比绝对字节  → 新实现还原原件字节(那正是收益)被判成「改变了行为」
        #   ② 没收窄就判红            → 老实现本来啥都没改，没有东西可收窄
        #   ③ 新实现增删部件就判红    → add_header_footer 加页眉页脚是本职，老新都增同样 10 个
        # 再加判据前先问：这条是在和老实现比，还是在和我脑内的理想值比？后者一律是假红。
        bad = []
        if not same_doc:
            bad.append("正文语义不一致 —— 迁移改变了行为，不是等价替换")
        # 判据是「迁移**改变了**部件集合的增删」,不是「新实现增删了部件」(2026-07-30 修)。
        # 有些脚本增删部件就是它的本职:add_header_footer 给 17 个节加页眉页脚,老新两边
        # 都新增同样的 10 个部件 —— 拿绝对值判会把它判红,而它恰恰是一次合格的等价迁移
        # (炸开面 59→26)。要抓的是迁移**引入**的增删差异。
        if set(rn["删掉"]) != set(ro["删掉"]) or set(rn["新增"]) != set(ro["新增"]):
            bad.append(
                f"迁移改变了部件集合：删 老{len(ro['删掉'])}/新{len(rn['删掉'])} · "
                f"增 老{len(ro['新增'])}/新{len(rn['新增'])}")
        if shrink <= 0:
            bad.append("新实现没有收窄炸开面 —— 迁移没带来收益，先查是不是没真走 surgical")
        for b in bad:
            print(f"  ✗ {b}")
        if not bad:
            print("  ✓ 语义等价且炸开面收窄")
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
