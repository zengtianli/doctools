#!/usr/bin/env python3
"""check_docx_collar — 任何用 python-docx 存盘的脚本，必须挂上 surgical 收口。

    python3 tools/check_docx_collar.py          # 判红即 exit 1
    python3 tools/check_docx_collar.py --list   # 只列名册

## 为什么要这个守卫

`lib/docx_safe_save.py` 把 python-docx 的存盘炸开面从 60 个部件收到 1 个，但它得**被
import** 才生效。光在 CLAUDE.md 里写「记得 import」不是约束 —— 下一个新脚本照样会裸
存盘，而且不会有任何症状：文件能打开、Word 也不报错，只是 60 个部件被无声重写了一遍。

判据不写脚本名单（名单会漂）：**打开每个 .py 自己看** —— import 了 python-docx 且调
了 `.save()`，就必须能找到 `import docx_safe_save`。

fail-closed：枚举为空 / 扫描根不存在，一律非 0 退出。空集上报绿是这类守卫最常见的死法。
"""
from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
# 判据 1 的扫描根**必须含 lib/**（2026-08-02 补）。原来只有 scripts/，于是 lib/ 下
# 「import python-docx + .save() + 无收口」这一格是个三不管：判据 1 扫不到、
# 判据 2 只管 ZipFile、判据 3 又主动让开（见下面那句 continue）。实测注入
# lib/_zz_libsaver.py（带 docx import 的裸存盘）守卫报 rc=0 —— 而 lib/ 正是本仓
# CLAUDE.md 钦定的公共模块层，「把存盘抽进公共层」恰恰是最该被守住的动作。
SCAN_ROOTS = [ROOT / "scripts", ROOT / "lib", ROOT / "src"]
# `src/doctools/` 是 2026-08-02 新建的可安装包壳。它当时落在**所有**闸门的扫描根之外，
# 往里放一个裸存盘 docx 的脚本五道门全绿 —— 「壳里不放业务逻辑」那条规矩
# 当时只写在 docstring 里，没有任何机器层兜着（铁律 #13）。
SCAN = SCAN_ROOTS[0]        # 兼容既有引用

# 只认 python-docx 本尊：`import docx` / `from docx import` / `from docx.xxx import`。
# 禁写成 docx[\w.]*——那会把自家 docx_parts/docx_xml/docx_safe_save 的 import 也判成
# python-docx 存盘脚本（2026-07-31 实证：pptx_cli 挂部件断言反被本守卫误判红）。
IMPORTS_DOCX = re.compile(r"^\s*(from docx(?:\.[\w.]+)?\s+import|import docx\b)", re.M)
COLLAR = re.compile(r"^\s*import docx_safe_save\b", re.M)

# ── 第二判据（2026-07-31 加）────────────────────────────────────────────
# 上面那条只管住 python-docx 那一半。surgical 改法（zipfile+lxml 手工重打包）
# 根本不 import docx，docx_safe_save 是 monkey-patch OpcPackage.save，
# 在这条路上一个字节都管不到 —— 于是丢部件无人可挡，实测两次：
#   162 部件 → 74（11 个原生图表 + 40 header + 18 footer 全丢）
#   137 部件 → 35（对同一文件开两个 ZipFile 句柄，逐部件复制被截断）
# 两次文件都照样能打开、Word 也不报错。
# 判据：自己开 ZipFile 写 docx 的脚本，必须能找到部件完整性断言。
WRITES_ZIP = re.compile(r"ZipFile\s*\([^)]*?['\"]w['\"]|ZIP_DEFLATED", re.S)
TOUCHES_DOCX = re.compile(r"\.docx\b|word/document\.xml")

# 保留给 ast 解析不了的文件（语法错/非 UTF-8）当兜底；正常路径走 has_part_assert()。
PART_ASSERT = re.compile(r"\b(assert_parts_intact|diff_parts)\s*\(")

# ── 判据必须落在**真调用**上，不能落在字面量出现上（2026-08-01 修）──────────
# 原来 PART_ASSERT 是裸 `\b(assert_parts_intact|diff_parts)\b`，于是这两样都能喂饱它：
#   ① 一条已经没人用的 `from docx_parts import assert_parts_intact` 死 import
#   ② 一句提到它的**注释**（fix_styleset.py:1350 就有一句）
# 实测：把死 import 删掉之后，往 fix_styleset.py 追加一段裸 ZipFile 写 docx，
# 守卫**照样报「17 个全挂 ✓」** —— 只因为那句注释还在。哑掉的守卫和它要防的 bug
# 是同一类东西，所以判据改成 ast：只认「函数调用」这一种形态，注释和 import 不算。
ASSERT_NAMES = {"assert_parts_intact", "diff_parts"}

# ── 第三判据（2026-08-02 加）· 判据要跟着调用链走，不能跟着 import 字面量走 ──────
# 前两条判据都问「这个文件自己 import 没 import python-docx」。于是有一个洞：
# **把存盘逻辑抽成公共函数，该函数所在的模块往往根本不 import docx**
# （它只是收一个 doc 参数然后 `doc.save(...)`），于是：
#   ① 新模块 IMPORTS_DOCX=False → 不进名册 → 不要求收口
#   ② 原来的调用方 `.save(` 没了 → 也退出名册
# 结果是**这次存盘从此没有任何守卫**，而守卫照样报绿。
# 2026-08-02 实测：sub/styles.py 与 sub/outline.py **各只有 1 处 `.save(`**，
# 都在各自的 `_save_with_backup` 里 —— 一次「抽公共函数」的普通重构就能同时
# 把这两个文件从名册上摘掉，五道闸门全绿。
#
# 所以第三判据按**反向 import 图**判：一个文件只要自己会 `.save(`，
# 且有任何（传递地）import 它的文件用了 python-docx，它就要挂收口。
# 宁可多要求（多一行 import，对非 docx 的存盘是无害 no-op），不可漏 —— fail-closed。
SAVE_CALL = re.compile(r"\.save\s*\(")
IMPORT_ANY = re.compile(r"^\s*(?:from\s+([\w.]+)\s+import|import\s+([\w.]+))", re.M)


def has_part_assert(src: str) -> bool:
    """源码里有没有**真的调用**部件完整性断言。"""
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return bool(PART_ASSERT.search(src))      # 解析不了就退回正则，不静默放行
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        fn = node.func
        name = (fn.id if isinstance(fn, ast.Name)
                else fn.attr if isinstance(fn, ast.Attribute) else None)
        if name in ASSERT_NAMES:
            return True
    return False

# 测试件豁免：它们在 tmp_path 里造玩具 docx，不碰交付件；而且好几个测试的断言就是
# 「裸 python-docx 会怎样」，挂上收口反而测不到要测的东西。
EXEMPT_DIR = re.compile(r"/tests?/|/_backup[-\w]*/|/_archive/|/archives?/")


def offenders() -> tuple[list[Path], list[Path]]:
    missing = [r for r in SCAN_ROOTS if not r.is_dir()]
    if missing:
        print(f"⛔ 扫描根不存在: {missing} —— 拒绝在空集上报通过", file=sys.stderr)
        raise SystemExit(2)
    need, bad = [], []
    for p in sorted(q for r in SCAN_ROOTS for q in r.rglob("*.py")):
        if EXEMPT_DIR.search(str(p)) or p.name.startswith("test_"):
            continue
        src = p.read_text(encoding="utf-8", errors="replace")
        if not (IMPORTS_DOCX.search(src) and ".save(" in src):
            continue
        need.append(p)
        if not COLLAR.search(src):
            bad.append(p)
    if not need:
        print(f"⛔ 在 {SCAN} 下一个用 python-docx 存盘的脚本都没扫到 —— 判据肯定坏了，"
              f"拒绝报绿", file=sys.stderr)
        raise SystemExit(2)
    return need, bad


def _repo_files(roots: list[Path]) -> list[Path]:
    """要检查的全部文件。

    ⚠ **按路径列全，不能按 stem 去重**（2026-08-02 修）。原来是
    `out.setdefault(p.stem, p)`，于是同名文件只留先扫到的那个：
    `lib/styles.py`（被 `sub/styles.py` 顶掉）、`lib/docx_revise.py`（被
    `scripts/document/docx_revise.py` 顶掉）、`lib/__init__.py` 三个文件
    **整个从判据 3 的视野里消失** —— 往 lib/styles.py 里加一处裸 `.save()`
    守卫照样报绿。stem 去重只该用在「解析 import 名字」那一步，不该用在「谁要被检查」。
    """
    out: list[Path] = []
    for root in roots:
        if not root.is_dir():
            continue
        for p in sorted(root.rglob("*.py")):
            if EXEMPT_DIR.search(str(p)) or p.name.startswith("test_"):
                continue
            out.append(p)
    return out


# ── 判据 3 的接收者溯源（2026-08-26 加）────────────────────────────────
# 为什么需要：判据 3 靠「有 .save( 且被 docx 使用方传递地 import」抓人，这对
# 「helper 收个 doc 参数、自己不提 docx 却存了 docx」是对的，但会误伤纯 PDF/图片
# 工具 —— pdf_bind.py 全文 docx 出现 0 次，5 处 .save() 全是 reportlab canvas，
# 只因被 pdf_cli.py import、而 pdf_cli 那条链上有人 import docx 就被判红。
#
# 判别点是**接收者是谁**，不是文件里提没提 docx。用 ast 溯源每个 .save() 的接收
# 变量：**全部**能追到已知非-docx 构造器才放行；有一个追不到就仍然拦
# （fail-closed —— 宁可误报也不许把真的 docx 存盘放过去）。
NON_DOCX_CTORS = (
    'canvas.Canvas', 'rl_canvas.Canvas', 'Canvas',        # reportlab
    'PdfWriter', 'PdfMerger', 'pypdf.PdfWriter',          # pypdf
    'Workbook', 'openpyxl.Workbook',                      # openpyxl
    'Image.new', 'Image.open',                            # PIL
    'Presentation',                                       # python-pptx
)


def _ctor_name(node) -> str:
    """Call 节点 → 'a.b.C' 形式的构造器名（取不出返回 ''）。"""
    f = getattr(node, 'func', None)
    parts = []
    while isinstance(f, ast.Attribute):
        parts.append(f.attr); f = f.value
    if isinstance(f, ast.Name):
        parts.append(f.id)
    return '.'.join(reversed(parts))


def saves_only_non_docx(src: str) -> bool:
    """文件里每一处 `<x>.save(` 的 <x> 都能追到已知非-docx 构造器 → True。"""
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return False                       # 解析不了就别放行
    assigned = {}                          # 变量名 → 构造器名
    for n in ast.walk(tree):
        if isinstance(n, ast.Assign) and isinstance(n.value, ast.Call):
            name = _ctor_name(n.value)
            for t in n.targets:
                if isinstance(t, ast.Name):
                    assigned[t.id] = name
    recvs = []
    for n in ast.walk(tree):
        if (isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == 'save'):
            v = n.func.value
            recvs.append(v.id if isinstance(v, ast.Name) else None)
    if not recvs:
        return False
    return all(r is not None and any(assigned.get(r, '').endswith(c) or assigned.get(r, '') == c
                                     for c in NON_DOCX_CTORS)
               for r in recvs)


def chain_offenders(roots: list[Path]) -> tuple[list[Path], list[Path]]:
    """第三判据：会 `.save(` 且被 python-docx 使用方（传递地）import 的文件，必须挂收口。

    与 offenders() 不同，这里**不要求非空**：一个仓里可以确实没有这种「被 docx
    脚本调用的存盘 helper」，那是正常状态。
    """
    files = _repo_files(roots)
    src_of = {p: p.read_text(encoding="utf-8", errors="replace") for p in files}

    # stem → **全部**同名文件（不是只留一个）。import 名字解析必然有歧义，
    # 遇到歧义就把边加给所有候选 —— 宁可多连一条边，也不许把某个文件从图里抹掉。
    by_stem: dict[str, list[Path]] = {}
    for p in files:
        by_stem.setdefault(p.stem, []).append(p)

    # 反向：谁被谁 import
    rev: dict[Path, set[Path]] = {p: set() for p in files}
    for p, src in src_of.items():
        for m in IMPORT_ANY.finditer(src):
            name = (m.group(1) or m.group(2) or "").split(".")[-1]
            for target in by_stem.get(name, ()):
                if target != p:
                    rev[target].add(p)

    def reached_by_docx(start: Path) -> bool:
        """自己或任何（传递的）上游使用方 import 了 python-docx。"""
        seen, stack = {start}, [start]
        while stack:
            cur = stack.pop()
            if IMPORTS_DOCX.search(src_of[cur]):
                return True
            for up in rev.get(cur, ()):
                if up not in seen:
                    seen.add(up)
                    stack.append(up)
        return False

    need, bad = [], []
    for p in files:
        src = src_of[p]
        if not SAVE_CALL.search(src):
            continue
        if IMPORTS_DOCX.search(src):
            continue                      # 已由第一判据管着（判据 1 现在也扫 lib/）
        if not reached_by_docx(p):
            continue                      # 存的不是 docx（xlsx/图片/json…），不管
        if not TOUCHES_DOCX.search(src) and saves_only_non_docx(src):
            continue                      # 全文不提 docx，且每处 .save() 都追到非-docx 构造器
        need.append(p)
        if not COLLAR.search(src):
            bad.append(p)
    return need, bad


def zip_offenders(roots: list[Path]) -> tuple[list[Path], list[Path]]:
    """第二判据：自己开 ZipFile 写 docx 的脚本，必须挂部件完整性断言。

    与 offenders() 不同，这里**不要求非空** —— 一个仓里可以确实没有 surgical 脚本，
    那是正常状态，不是判据坏了。
    """
    need, bad = [], []
    for root in roots:
        if not root.is_dir():
            continue
        for p in sorted(root.rglob("*.py")):
            if EXEMPT_DIR.search(str(p)) or p.name.startswith("test_"):
                continue
            if p.name == "docx_parts.py":       # 断言本体
                continue
            src = p.read_text(encoding="utf-8", errors="replace")
            if not (WRITES_ZIP.search(src) and TOUCHES_DOCX.search(src)):
                continue
            need.append(p)
            if not has_part_assert(src):
                bad.append(p)
    return need, bad


def main() -> int:
    need, bad = offenders()
    # 第二判据的扫描范围：默认全仓（scripts/ + lib/）+ 命令行额外指定的目录
    # （项目侧的一次性 surgical 脚本不在本仓，可显式传：
    #   python3 check_docx_collar.py ~/Work/projects/qual-supply/scripts）
    # 历史：判据 2026-07-31 立时存量 27 个未接线，默认只扫显式目录以免守卫第一天长红；
    # 同日存量批量接线清零（25 脚本 + 2 engine，各带 stub/反向验证）后翻成默认全扫——
    # fail-closed 从此不再靠 --all 自觉。--all 保留为兼容 no-op。
    extra = [Path(a).expanduser() for a in sys.argv[1:] if not a.startswith("-")]
    scan_roots = [SCAN, ROOT / "lib"] + extra
    z_need, z_bad = zip_offenders(scan_roots)
    c_need, c_bad = chain_offenders(scan_roots)
    if "--list" in sys.argv:
        for p in need:
            mark = "✗" if p in bad else "✓"
            print(f"{mark} [python-docx] {p.relative_to(ROOT)}")
        for p in z_need:
            mark = "✗" if p in z_bad else "✓"
            print(f"{mark} [zipfile]     {p}")
        for p in c_need:
            mark = "✗" if p in c_bad else "✓"
            print(f"{mark} [调用链]      {p}")
        return 0
    if c_bad:
        print(f"⛔ {len(c_bad)}/{len(c_need)} 个文件会 .save() 且被 python-docx 使用方"
              f"（传递地）import，但没挂 surgical 收口：", file=sys.stderr)
        for p in c_bad:
            print(f"    {p}", file=sys.stderr)
        print("\n为什么算它：判据跟着**调用链**走，不跟着 import 字面量走。"
              "\n把存盘抽成公共函数时，新模块往往根本不 import docx（只收一个 doc 参数），"
              "\n于是前两条判据同时看不见它、原调用方也退出名册 —— 这次存盘从此无人守。"
              "\n修法与下面一致：文件顶部 import docx_safe_save。"
              "\n（确实存的不是 docx？那把它挪出被 docx 脚本 import 的链，或说明白为什么。）",
              file=sys.stderr)
        return 1
    if z_bad:
        print(f"⛔ {len(z_bad)}/{len(z_need)} 个脚本自己开 ZipFile 写 docx，"
              f"但没挂部件完整性断言：", file=sys.stderr)
        for p in z_bad:
            print(f"    {p}", file=sys.stderr)
        print(f"\n为什么要挂：docx_safe_save 只 monkey-patch python-docx 的存盘路径，"
              f"zipfile 手工重打包它管不到。实测两次事故 162→74、137→35 部件，"
              f"文件照样能打开、Word 不报错。\n"
              f"修法：存盘后加\n"
              f"    sys.path.append('{ROOT / 'lib'}')\n"
              f"    from docx_parts import assert_parts_intact\n"
              f"    assert_parts_intact(src, dst, allow_added={{'word/comments.xml'}})\n"
              f"（本意就要减部件的场景改用 diff_parts + 自定白名单断言）",
              file=sys.stderr)
        return 1
    if bad:
        print(f"⛔ {len(bad)}/{len(need)} 个脚本用 python-docx 存盘但没挂 surgical 收口"
              f"（炸开面 ~60 个部件 vs 1 个）：", file=sys.stderr)
        for p in bad:
            print(f"    {p.relative_to(ROOT)}", file=sys.stderr)
        depth = "  # parents[N] 里 N = 该文件到仓根的层数"
        print(f"\n修法：在文件顶部加\n"
              f'    import sys as _sys\n'
              f'    from pathlib import Path as _Path\n'
              f'    _sys.path.append(str(_Path(__file__).resolve().parents[N] / "lib")){depth}\n'
              f"    import docx_safe_save  # noqa: E402,F401\n"
              f"（append 不是 insert(0)：lib/ 和 sub/ 有同名模块，插 0 位会顶掉脚本自己那份）\n"
              f"确实该裸存盘（从零造新文件）也照样加 —— 收口对 Document() 新建文档自动"
              f"不介入，不会碍事。", file=sys.stderr)
        return 1
    print(f"✓ {len(need)} 个用 python-docx 存盘的脚本全部挂了 surgical 收口")
    print(f"✓ {len(z_need)} 个自己开 ZipFile 写 docx 的脚本全部挂了部件完整性断言")
    print(f"✓ {len(c_need)} 个「自己不 import docx、但被 docx 使用方调用」的存盘 helper "
          f"全部挂了收口")
    return 0


if __name__ == "__main__":
    sys.exit(main())
