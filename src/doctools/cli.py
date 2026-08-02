"""console_script 入口 —— `doctools …` ≡ `python3 scripts/document/docx_cli.py …`。

装包**不是替代**绝对路径调用：~/Work 有 130 处 `python3 <abs>/docx_cli.py …` 在跑，
它们必须继续有效。所以这里不复刻任何解析/分发逻辑，只按文件路径把 `docx_cli.py`
载进来调它的 `main()` —— 两条入口共用同一个 argv 语义，不可能漂移。

定位方式按顺序试**两个根**，两者的内部结构完全一样（`scripts/` `lib/` `config/`
`schemas/` 四个兄弟目录），所以选中哪个之后，下游代码一行都不用分叉：

  ① **工作树根** `<repo>/`（本文件在 `<repo>/src/doctools/`，`parents[2]` 即仓根）。
     editable 安装走这条 —— 改完源码立刻生效，这正是 editable 的意义。
  ② **包内副本** `<site-packages>/doctools/_bundled/`。wheel 安装走这条。
     由 pyproject 的 `force-include` 在构建时镜像进来，工作树里没有这份目录。

为什么「两份副本」不会漂：**仓库里根本没有第二份**。`_bundled/` 只存在于构建产物
里，是构建那一刻从工作树 checkout 出来的逐字节快照（实测其文件集 == `git ls-files
scripts lib config schemas`）。所以两条分支加载的永远是同一批字节，差别只是「这个
wheel 是什么时候打的」—— 那是发布物的正常语义，不是双份维护。
机器判据见 `tools/check_wheel_selfcontained.py`（逐文件 sha256 比对 + clean venv 实跑）。

fail-closed：两个根都找不到 docx_cli.py 一律 exit 2，并把**试过的路径全列出来** ——
一个「装上了但其实指不到实现」的 CLI，报错必须一眼能看出是安装形态的问题。
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any, Optional

# src-layout：本文件在 `<repo>/src/doctools/cli.py`，parents[2] 才是仓根。
# 为什么是 src-layout：包直接放仓根时，hatchling 的 **editable** 安装会往
# site-packages 写一个指向**项目根**的 .pth，于是共享的 ~/Dev/.venv 里凭空多出
# lib / scripts / tools / config … 八个顶层可导入名（实测 `import lib` 当场
# 解析到 doctools/lib/__init__.py），而 tools/dev、tools/mactools 共用那个 venv。
# `dev-mode-exact = true` 能只映射一个名字，但它依赖运行时包 editables，
# venv 里没有 → 每次启动 python 都报 .pth 错误，比原问题更糟。src-layout 无此依赖。
_REPO = Path(__file__).resolve().parents[2]

# 实现根候选，**顺序即优先级**：工作树优先，包内副本兜底。
# 两者结构相同，所以「选根」是唯一的分叉点，选完之后没有第二处 if。
_IMPL_ROOTS = (
    _REPO,                                   # editable：<repo>/
    Path(__file__).resolve().parent / "_bundled",  # wheel：<site-packages>/doctools/_bundled/
)

_MOD_NAME = "doctools._docx_cli_entry"


def _find_impl() -> Path:
    """返回第一个真实存在的 `docx_cli.py`；两个根都没有则 fail-closed。"""
    tried = []
    for root in _IMPL_ROOTS:
        cand = root / "scripts" / "document" / "docx_cli.py"
        tried.append(cand)
        if cand.is_file():
            return cand
    listing = "\n".join(f"    - {p}" for p in tried)
    print(
        f"[doctools] FATAL: 找不到实现入口，按顺序试过：\n{listing}\n"
        f"  ① 是工作树（editable 安装才有），② 是包内副本（wheel 应当自带）。\n"
        f"  两个都没有 = 这个 wheel 打漏了 `force-include`，或包被裁剪过。\n"
        f"  重新构建请跑：python3 tools/check_wheel_selfcontained.py",
        file=sys.stderr,
    )
    raise SystemExit(2)


def _load_docx_cli() -> Any:
    cached = sys.modules.get(_MOD_NAME)
    if cached is not None:
        return cached
    _DOCX_CLI = _find_impl()
    spec = importlib.util.spec_from_file_location(_MOD_NAME, _DOCX_CLI)
    if spec is None or spec.loader is None:
        print(f"[doctools] FATAL: cannot spec {_DOCX_CLI}", file=sys.stderr)
        raise SystemExit(2)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[_MOD_NAME] = mod
    spec.loader.exec_module(mod)
    return mod


def main(argv: Optional[list[str]] = None) -> int:
    return int(_load_docx_cli().main(argv))


if __name__ == "__main__":
    sys.exit(main())
