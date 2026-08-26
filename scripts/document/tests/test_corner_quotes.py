"""tests for 直角引号（corner brackets）归一 —— 2026-08-25 桐乡 ztl.docx 实测出的 2 条缺口。

规则来源：用户 2026-08-22 /govern 钦定「外发中文件一律弯引号 “”‘’，禁直角引号 「」『』」。
本文件守 `text_fixes.fix_quotes` 这一侧（`docx_fmt.py text --quotes` / `md_tools.py format
--quotes` / `pptx_cli` 共用它），专项工具 `docx_quotes.py` 是另一条通路。

1. 『』根本没被处理 —— 旧 QUOTE_PATTERN 只含 「」，单引号形态整个漏网
2. 「」的开闭被 counter 奇偶算错 —— 「」方向自带，不该进奇偶配对：
   前文出现一个孤立直立引号，就把后面整串「」翻成 ”…“（旧实现实测）
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parents[3] / "lib"
sys.path.insert(0, str(_LIB))

spec = importlib.util.spec_from_file_location("text_fixes_uut", str(_LIB / "text_fixes.py"))
assert spec and spec.loader
tf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tf)
fix_quotes = tf.fix_quotes


def test_corner_double_to_curly():
    assert fix_quotes("他说「过程决定结果」。")[0] == "他说“过程决定结果”。"


def test_corner_single_to_curly():
    """『』→ ‘’：旧实现完全不碰，是漏网的那一半。"""
    assert fix_quotes("又说『小引号』。")[0] == "又说‘小引号’。"


def test_corner_mixed():
    assert fix_quotes("混排「甲」和『乙』")[0] == "混排“甲”和‘乙’"


def test_corner_direction_not_flipped_by_stray_quote():
    """回归：孤立直立引号在前，旧实现把「」整串翻转成 ”…“。"""
    got, _, _ = fix_quotes('前面有个孤引号" 然后「过程决定结果」')
    assert got == '前面有个孤引号“ 然后“过程决定结果”'


def test_reversed_curly_still_repaired_by_counter():
    """方向不明的弯引号仍靠 counter 奇偶修正——这条能力不能被上面的改动搞丢。"""
    assert fix_quotes("说”过程决定结果“。")[0] == "说“过程决定结果”。"


def test_correct_quotes_not_counted():
    got, n, _ = fix_quotes("已经是“对的”了。")
    assert got == "已经是“对的”了。"
    assert n == 0


def test_corner_counted_as_changes():
    assert fix_quotes("他说「过程决定结果」。")[1] == 2


def test_cross_run_counter_still_pairs():
    """docx 场景：引号被拆在两个 run 里，counter 跨 run 传递。"""
    a, _, c = fix_quotes('他说"', 0)
    b, _, _ = fix_quotes('过程决定结果"', c)
    assert a + b == "他说“过程决定结果”"
