"""console_script 入口 —— `doctools …` ≡ `python3 scripts/document/docx_cli.py …`。

装包**不是替代**绝对路径调用：~/Work 有 130 处 `python3 <abs>/docx_cli.py …` 在跑，
它们必须继续有效。所以这里不复刻任何解析/分发逻辑，只按文件路径把 `docx_cli.py`
载进来调它的 `main()` —— 两条入口共用同一个 argv 语义，不可能漂移。

定位方式：本包在仓根（`<repo>/doctools/`），`parents[1]` 即仓根，
`<repo>/scripts/document/docx_cli.py` 就是那个唯一实现。uv workspace 的 member
默认 editable 安装，所以装完之后这条路径指回真实工作树。

fail-closed：找不到 docx_cli.py 一律 exit 2 并打印它到底找了哪个路径 —— 一个
「装上了但其实指不到实现」的 CLI，报错必须一眼能看出是安装形态的问题。
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any, Optional

_REPO = Path(__file__).resolve().parents[1]
_DOCX_CLI = _REPO / "scripts" / "document" / "docx_cli.py"

_MOD_NAME = "doctools._docx_cli_entry"


def _load_docx_cli() -> Any:
    cached = sys.modules.get(_MOD_NAME)
    if cached is not None:
        return cached
    if not _DOCX_CLI.is_file():
        print(
            f"[doctools] FATAL: 找不到实现入口 {_DOCX_CLI}\n"
            f"  本包只是 console_script 壳，实现在仓库工作树里。\n"
            f"  非 editable 安装（wheel 只带 doctools/ 一个包）会走到这里 —— "
            f"请在 ~/Dev 用 `uv sync --all-packages` 装 workspace member。",
            file=sys.stderr,
        )
        raise SystemExit(2)
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
