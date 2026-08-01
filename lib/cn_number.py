#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cn_number.py — 中文数字 → int 的**唯一**实现（2026-08-01 立）。

纯 stdlib，零第三方依赖 —— 它会被「自己开 zipfile 写 docx」那条路上的脚本 import，
那条路必须能用系统 python3 直接跑，不能要求先进 venv。

──────────────────────────────────────────────────────────────────────────────
为什么统一
──────────────────────────────────────────────────────────────────────────────
合并前全仓有 **5 处** 各写各的中文数字转换，分三档能力，同一个输入给三种答案：

    输入        chapter   outline   blocks    caption   styles
    十五          15        15        15        15        15
    十六          16        16        16      None      None
    二十三        23        23        23      None      None
    一百零五     105       105      None      None      None
    一千        1000      1000      None      None      None
    十一二        12        12      None      None      None
    〇 / 两    ValueError ValueError  0 / 2    None      None
    无法解析   ValueError ValueError  None      None      None

后果不是学术问题：`caption number` 的章计数器碰上「十六、xxx」解析失败就**不换章**，
第 16 章往后的表图继续编成「表15-7、表15-8…」；`styles` 那侧更隐蔽——解析失败时
上层写的是 `_parse_chapter_from_text(t) or (chapter + 1)`，**静默拿「上一章+1」顶上**，
章号跳号时就分叉。三个入口（chapter.apply_convert_arabic / blocks.apply_fix_heading_disorder
/ caption.apply_number）都挂在 `typeset_apply.py` 的步骤表里，所以 /typeset 一条龙每次都在跑。

──────────────────────────────────────────────────────────────────────────────
统一到哪一档：三套的**能力上界**
──────────────────────────────────────────────────────────────────────────────
算法取 chapter/outline 的累加器那套（当时最强的一档）：

  · 阿拉伯数字串直通 `int(s)`
  · 「十」= 10、「十X」省略一（十五 = 15）
  · 支持 十 / 百 / 千 的位段累加（一百零五 = 105、两千零五 = 2005）
  · 数字表并入 blocks 独有的 `〇`(0) 与 `两`(2)（chapter/outline 旧实现对这两个字抛错，
    但它们的正则字符类根本产不出这两个字，所以并进来是纯增量）

**不支持「万」及以上** —— 三套旧实现没有一套支持，累加器要支持「万」需要额外的
section 结构（万以下的整段乘 10000），那是新增能力不是合并，别在这次悄悄加。

──────────────────────────────────────────────────────────────────────────────
为什么必须给两个 API（严格 / 宽松）
──────────────────────────────────────────────────────────────────────────────
5 个调用点里 **2 个吃异常、3 个吃 None**，只留一个会把另一半改坏：

  · `chapter.parse_h1` / `outline` 三处 —— 靠 `except ValueError` 控流；
    若改成收 None，会拿着 None 继续往下算，写出「None、标题」这种脏文本，**不报错**。
  · `blocks.detect_heading_form` / `caption.parse_chapter` / `styles._parse_chapter_from_text`
    —— 靠 None 分支；若改成抛异常，直接崩。

所以：`chinese_to_arabic()` 严格（抛 ValueError），`cn_to_int()` 宽松（返 None），
后者就是前者外面包一层 try/except，**语义永远一致，不会再分叉**。

──────────────────────────────────────────────────────────────────────────────
本模块**不管**正则
──────────────────────────────────────────────────────────────────────────────
调用点各自的章标题正则（`^([一二三四五六七八九十]+)、` 之类）字符类**没有统一**，
是另一根轴：动它会同时放大「哪些段落被认成章标题」的匹配面，得单独一轮做、单独对拍。
`CN_CHARS` 摆在这里供以后引用，**当前没有任何调用点用它** —— 这是有意的。
具体影响：caption / styles 侧的字符类不含 `百`/`零`，所以「一百零五、」压根匹配不到，
本次合并对它们真正解锁的是「十六…九十九」这一段。

用法::

    from cn_number import chinese_to_arabic, cn_to_int
    chinese_to_arabic("二十三")   # 23；解析不了抛 ValueError
    cn_to_int("二十三")           # 23；解析不了返 None
"""
from __future__ import annotations

from typing import Optional

__all__ = ["CN_DIGIT", "CN_UNIT", "CN_CHARS", "chinese_to_arabic", "cn_to_int"]

#: 单字数字。`〇` `两` 来自 blocks.py 那一档（其余四处旧实现没有）。
CN_DIGIT = {
    "零": 0, "〇": 0,
    "一": 1,
    "二": 2, "两": 2,
    "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9,
}

#: 位权。三套旧实现的上界就到 `千`，见模块 docstring「不支持万」。
CN_UNIT = {"十": 10, "百": 100, "千": 1000}

#: 本模块认得的全部中文字符，供正则字符类引用（`f"[{CN_CHARS}]+"`），
#: 省得各文件再手抄一遍。**当前无调用点使用**，见模块 docstring「本模块不管正则」。
CN_CHARS = "".join(CN_DIGIT) + "".join(CN_UNIT)


def chinese_to_arabic(s: str) -> int:
    """中文数字串转 int（**严格版**：解析不了抛 ValueError）。

    支持「一」「十」「十一」「二十」「三十五」「一百零三」「两千零五」等；
    `s` 已是阿拉伯数字串则直接 `int(s)`。

    宽容边界（沿用 chapter/outline 旧行为，不做收紧——收紧会改掉今天 outline
    promote-h1 的匹配结果）：畸形串不报错，`十一二` → 12、`十十` → 20。

    Raises:
        ValueError: 空串 / 含本模块不认得的字符。
    """
    s = s.strip()
    if not s:
        raise ValueError("empty numeral")
    if s.isdigit():
        return int(s)

    total = 0
    current = 0  # 当前累积位段
    for ch in s:
        if ch in CN_DIGIT:
            current = CN_DIGIT[ch]
        elif ch in CN_UNIT:
            unit = CN_UNIT[ch]
            if current == 0:
                current = 1  # 「十X」开头省略一: 十 = 10
            total += current * unit
            current = 0
        else:
            raise ValueError(f"unrecognized char in numeral: {ch!r}")

    total += current
    return total


def cn_to_int(s: str) -> Optional[int]:
    """中文数字串转 int（**宽松版**：解析不了返 None，不抛）。

    语义与 :func:`chinese_to_arabic` **完全一致**，只是把 ValueError 换成 None ——
    给那些「匹配到了就用、没匹配就跳过」的调用点用（blocks / caption / styles）。

    只吞 ValueError：非 str 入参照旧炸 AttributeError（与旧 blocks.cn_to_int 一致，
    那是调用方传错了类型，不该被静默成「解析不出」）。
    """
    try:
        return chinese_to_arabic(s)
    except ValueError:
        return None
