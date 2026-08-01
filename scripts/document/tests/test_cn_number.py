"""tests for lib/cn_number.py —— 中文数字 → int 的 SSOT 回归门。

2026-08-01 起因：全仓有 5 处各写各的中文数字转换（chapter / outline / blocks /
caption / styles），分三档能力，同一个输入给三种答案 —— 「十六、xxx」在 caption 侧
解析失败会让章计数器不切换，第 16 章往后的表图继续按上一章编号；styles 侧更隐蔽，
解析失败时上层拿「上一章+1」静默顶上。5 处合并到 lib/cn_number.py，本文件是它的门。

红线（改坏了必须红）：
  · 两个 API 的**异常契约不能统一成一种** —— chapter/outline 三处靠 `except ValueError`
    控流，blocks/caption/styles 三处靠 None 分支。只留一个会把另一半改坏（收 None 的那侧
    会拿着 None 往下算，写出「None、标题」这种脏文本且不报错）。
  · 能力上界：十 / 百 / 千 + 「十X」省略一 + 〇/两 + 阿拉伯直通，缺一即回退到旧的某一档。
  · 「万」**不支持**是有意的（三套旧实现没有一套支持），哪天有人加了，这里要显式改。
  · 畸形串的宽容度不许收紧：`十一二`→12、`十十`→20 是 chapter/outline 今天的行为，
    收紧会改掉 outline promote-h1 的匹配结果。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parents[3] / "lib"))

from cn_number import (  # noqa: E402
    CN_CHARS,
    CN_DIGIT,
    CN_UNIT,
    chinese_to_arabic,
    cn_to_int,
)


# ─────────────────────────────────────────────────────────────────────────────
# 能力上界：三套旧实现能认的，这里必须全认
# ─────────────────────────────────────────────────────────────────────────────
# 备注列 = 合并前哪一档认得它（chapter/outline=C·O，blocks=B，caption/styles=查表）
OK_CASES = [
    # (输入, 期望)          # 合并前谁认得
    ("一", 1),              # 全部
    ("三", 3),              # 全部
    ("九", 9),              # 全部
    ("十", 10),             # 全部
    ("十一", 11),           # 全部
    ("十五", 15),           # 全部（查表的上界，正好卡在这）
    ("十六", 16),           # 只有 C·O + B —— 查表那两处过去返 None
    ("十九", 19),
    ("二十", 20),
    ("二十三", 23),         # 只有 C·O + B
    ("三十五", 35),
    ("九十九", 99),         # B 的上界（4 条硬编码形态分支到此为止）
    ("一百", 100),          # 只有 C·O —— blocks 过去返 None
    ("一百零三", 103),
    ("一百零五", 105),      # 只有 C·O
    ("一百二十三", 123),
    ("一千", 1000),         # 只有 C·O
    ("两千零五", 2005),     # 谁都不全认：C·O 认千但对「两」抛错，B 认「两」但不认千
    # 零 / 〇 / 两 单字
    ("零", 0),
    ("〇", 0),              # 过去只有 B 认；C·O 抛 ValueError
    ("两", 2),              # 同上
    ("十零", 10),
    # 阿拉伯直通
    ("5", 5),
    ("42", 42),
    ("0", 0),
    ("105", 105),
    # 两侧空白照旧 strip
    ("  二十三  ", 23),
]


@pytest.mark.parametrize("text,expected", OK_CASES)
def test_strict_api_parses(text, expected):
    assert chinese_to_arabic(text) == expected


@pytest.mark.parametrize("text,expected", OK_CASES)
def test_lenient_api_agrees_with_strict(text, expected):
    """两个 API 语义必须逐值一致 —— 分叉过一次就够了。"""
    assert cn_to_int(text) == expected


# ─────────────────────────────────────────────────────────────────────────────
# 异常契约：严格版抛 / 宽松版返 None，两者都不许改
# ─────────────────────────────────────────────────────────────────────────────
BAD_CASES = ["", "   ", "abc", "第五", "五、", "十5", "二十 三"]


@pytest.mark.parametrize("text", BAD_CASES)
def test_strict_api_raises_valueerror(text):
    """chapter.parse_h1 / outline 三处全靠 `except ValueError` 控流。"""
    with pytest.raises(ValueError):
        chinese_to_arabic(text)


@pytest.mark.parametrize("text", BAD_CASES)
def test_lenient_api_returns_none(text):
    """blocks / caption / styles 全靠 None 分支；这里抛异常它们会直接崩。"""
    assert cn_to_int(text) is None


def test_lenient_api_does_not_swallow_type_errors():
    """只吞 ValueError：非 str 入参是调用方传错类型，不该被静默成「解析不出」。"""
    with pytest.raises(AttributeError):
        cn_to_int(None)  # type: ignore[arg-type]


# ─────────────────────────────────────────────────────────────────────────────
# 边界：畸形串的宽容度（沿用 chapter/outline 旧行为，不许收紧）
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("text,expected", [
    ("十一二", 12),   # 数字后又跟数字：后者覆盖前者，不报错
    ("十十", 20),     # 两套旧算法在这个输入上巧合一致，保持
    ("二十百", 120),  # 位权后又跟位权
])
def test_malformed_inputs_stay_tolerant(text, expected):
    assert chinese_to_arabic(text) == expected


def test_wan_is_deliberately_unsupported():
    """「万」不在能力范围内 —— 三套旧实现没有一套支持它。

    要加得改累加器结构（万以下整段乘 10000），那是新增能力不是合并，
    改的时候顺手把这条测试一起改，别让它悄悄通过。
    """
    assert "万" not in CN_UNIT
    with pytest.raises(ValueError):
        chinese_to_arabic("一万")


# ─────────────────────────────────────────────────────────────────────────────
# 常量表：调用点的正则字符类以后要引用它
# ─────────────────────────────────────────────────────────────────────────────
def test_char_table_covers_all_variants():
    """数字表必须是 5 处旧实现的并集（〇/两 来自 blocks，其余四处没有）。"""
    for ch in "零〇一二两三四五六七八九":
        assert ch in CN_DIGIT
    assert CN_UNIT == {"十": 10, "百": 100, "千": 1000}


def test_cn_chars_is_exactly_what_parser_accepts():
    """CN_CHARS 是给正则字符类用的，必须与解析器实际认得的字符集严格同步。"""
    for ch in CN_CHARS:
        assert chinese_to_arabic(ch) is not None  # 单字都能解析，不抛
    assert set(CN_CHARS) == set(CN_DIGIT) | set(CN_UNIT)


# ─────────────────────────────────────────────────────────────────────────────
# 调用点：确认 5 处真的 import 了同一份（不是各留一份副本）
# ─────────────────────────────────────────────────────────────────────────────
def test_no_local_reimplementation_left_in_sub():
    """抽 SSOT = 调用点改为 import 同一份，不是复制粘贴到新地方再留旧的。"""
    sub = Path(__file__).resolve().parents[1] / "sub"
    offenders = []
    for py in sub.glob("*.py"):
        src = py.read_text(encoding="utf-8")
        # 只看代码行，注释里提到旧名字（迁移记录）不算
        code = "\n".join(ln for ln in src.splitlines() if not ln.lstrip().startswith("#"))
        for token in ("_CN_DIGIT", "_CN_UNIT", "CN_NUM ="):
            if token in code:
                offenders.append(f"{py.name}: {token}")
    assert not offenders, f"sub/ 下仍有中文数字的本地实现: {offenders}"
