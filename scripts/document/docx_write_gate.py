#!/usr/bin/env python3
"""docx 原地写回并发门（SSOT）— 写回前 md5/mtime 基线比对。

Why（2026-07-13 立）：多会话 / 用户 WPS 并行改同一 docx 是常态，天台报告两会话
靠手工 md5 才没互相覆盖（lines-plan 复发摩擦「docx 并行编辑撞车」）。此门 =
结构性防线：引擎读入时 capture 基线，原地写回前 assert_unchanged——文件在此
期间被第三方改过 → 拒写并退出(3)，防静默覆盖别人的改动。

用法（引擎侧，只对「写回源文件」路径生效；另存新文件无冲突不需要门）：

    from docx_write_gate import WriteGate
    gate = WriteGate(path)          # 读入源文件之前 capture
    ...                             # 解析/处理（可能耗时）
    gate.assert_unchanged()         # 原地写回前一刻

逃生：env `DOCX_GATE_OK=1`（打印 warning 后放行，语义同 GIT_ADD_ALL_OK 家族）。
"""
from __future__ import annotations

import hashlib
import os
import sys
from pathlib import Path


def _md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class WriteGate:
    """capture 时记 mtime+md5；check 时 mtime 未动走快路径，动了再比 md5。"""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        st = self.path.stat()
        self.mtime = st.st_mtime
        self.md5 = _md5(self.path)

    def changed(self) -> bool:
        try:
            st = self.path.stat()
        except FileNotFoundError:
            return True
        if st.st_mtime == self.mtime:
            return False
        return _md5(self.path) != self.md5

    def assert_unchanged(self) -> None:
        if not self.changed():
            return
        if os.environ.get("DOCX_GATE_OK") == "1":
            print(f"⚠ DOCX_GATE_OK=1 → 跳过并发写回门（{self.path} 确已被外部修改，继续覆盖）",
                  file=sys.stderr)
            return
        print(
            f"⛔ 写回被拒：{self.path.name} 在本工具读入后被其他进程修改过"
            f"（用户 WPS / 另一 CC 会话?）。先核对两边改动再重跑；"
            f"确认要覆盖 → 前缀 DOCX_GATE_OK=1 重跑。",
            file=sys.stderr,
        )
        raise SystemExit(3)
