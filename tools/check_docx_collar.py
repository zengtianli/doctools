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

IMPORTS_DOCX = re.compile(r"^\s*(from docx[\w.]*\s+import|import docx\b)", re.M)
COLLAR = re.compile(r"^\s*import docx_safe_save\b", re.M)

# 测试件豁免：它们在 tmp_path 里造玩具 docx，不碰交付件；而且好几个测试的断言就是
# 「裸 python-docx 会怎样」，挂上收口反而测不到要测的东西。
EXEMPT_DIR = re.compile(r"/tests?/")


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


def main() -> int:
    need, bad = offenders()
    if "--list" in sys.argv:
        for p in need:
            mark = "✗" if p in bad else "✓"
            print(f"{mark} {p.relative_to(ROOT)}")
        return 0
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
