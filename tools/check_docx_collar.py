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

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN = ROOT / "scripts"

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
PART_ASSERT = re.compile(r"\b(assert_parts_intact|diff_parts)\b")

# 测试件豁免：它们在 tmp_path 里造玩具 docx，不碰交付件；而且好几个测试的断言就是
# 「裸 python-docx 会怎样」，挂上收口反而测不到要测的东西。
EXEMPT_DIR = re.compile(r"/tests?/|/_backup[-\w]*/|/_archive/|/archives?/")


def offenders() -> tuple[list[Path], list[Path]]:
    if not SCAN.is_dir():
        print(f"⛔ 扫描根不存在: {SCAN} —— 拒绝在空集上报通过", file=sys.stderr)
        raise SystemExit(2)
    need, bad = [], []
    for p in sorted(SCAN.rglob("*.py")):
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
            if not PART_ASSERT.search(src):
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
    if "--list" in sys.argv:
        for p in need:
            mark = "✗" if p in bad else "✓"
            print(f"{mark} [python-docx] {p.relative_to(ROOT)}")
        for p in z_need:
            mark = "✗" if p in z_bad else "✓"
            print(f"{mark} [zipfile]     {p}")
        return 0
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
