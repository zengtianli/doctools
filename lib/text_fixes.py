"""中文文本规范化原语(引号/标点/单位)。
源自 dockit.text —— 2026-06-01 解耦:doctools 不再依赖 stations/dockit(切断 doctools⇄dockit 循环+反向跨层)。
dockit(在线转换 web app)保留自己那份;此为 doctools 的 canonical 副本(3 个稳定纯函数)。
"""
"""Text formatting rules — quotes, punctuation, and unit normalization.

Pure functions with no I/O dependencies. Designed for CJK (Chinese) document processing.

Usage:
    from dockit.text import fix_all

    result, stats, counter = fix_all("这是一段测试文本,包含英文标点.")
    # result: "这是一段测试文本，包含英文标点。"
    # stats: {"quotes": 0, "punctuation": 2, "units": 0}
"""

import re

# -- Punctuation mapping (English -> Chinese) ---------------------------------

PUNCTUATION_MAP = {
    ",": "\uff0c",
    ":": "\uff1a",
    ";": "\uff1b",
    "!": "\uff01",
    "?": "\uff1f",
    "(": "\uff08",
    ")": "\uff09",
}

# -- Unit mapping (Chinese text -> standard symbols) ---------------------------
# Sorted by key length descending to ensure longest match first.

UNITS_MAP = {
    # Area
    "\u5e73\u65b9\u516c\u91cc": "km\u00b2",
    "\u5e73\u65b9\u5343\u7c73": "km\u00b2",
    "\u5e73\u65b9\u5398\u7c73": "cm\u00b2",
    "\u5e73\u65b9\u6beb\u7c73": "mm\u00b2",
    "\u5e73\u65b9\u7c73": "m\u00b2",
    # Volume
    "\u7acb\u65b9\u516c\u91cc": "km\u00b3",
    "\u7acb\u65b9\u5343\u7c73": "km\u00b3",
    "\u7acb\u65b9\u5398\u7c73": "cm\u00b3",
    "\u7acb\u65b9\u6beb\u7c73": "mm\u00b3",
    "\u7acb\u65b9\u7c73": "m\u00b3",
    # Length
    "\u516c\u91cc": "km",
    "\u5343\u7c73": "km",
    "\u5398\u7c73": "cm",
    "\u6beb\u7c73": "mm",
    "\u5fae\u7c73": "\u03bcm",
    "\u7eb3\u7c73": "nm",
    # Mass
    "\u516c\u65a4": "kg",
    "\u5343\u514b": "kg",
    "\u6beb\u514b": "mg",
    "\u5fae\u514b": "\u03bcg",
    # Volume (liquid)
    "\u6beb\u5347": "mL",
    "\u5fae\u5347": "\u03bcL",
    # Time
    "\u5c0f\u65f6": "h",
    "\u5206\u949f": "min",
    "\u79d2\u949f": "s",
    # Temperature
    "\u6444\u6c0f\u5ea6": "\u2103",
    "\u534e\u6c0f\u5ea6": "\u2109",
    # Superscript normalization
    "km2": "km\u00b2",
    "km3": "km\u00b3",
    "m2": "m\u00b2",
    "m3": "m\u00b3",
}

_UNITS_SORTED = sorted(UNITS_MAP.items(), key=lambda x: len(x[0]), reverse=True)

# 单位替换必须带边界,否则裸 str.replace 会改坏普通词汇(2026-07-26 实测:
# 小时候→h候 · 毫米波→mm波 · 秒钟表→s表 · 分钟表→min表 · Item2→Item²)。
#   中文单位:只在**紧跟数字**时替换(允许中间空格) —— "5 平方米"✓ / "小时候"✗
#   ASCII 上标(m2/km3):前面不许是字母、后面不许是字母数字 —— "5m2"✓ / "Item2"✗
_UNIT_RULES = [
    (
        re.compile(r"(?<=[0-9０-９])(\s*)" + re.escape(k))
        if not k.isascii()
        else re.compile(r"(?<![A-Za-z])(" + re.escape(k) + r")(?![A-Za-z0-9])"),
        v,
    )
    for k, v in _UNITS_SORTED
]

# -- Quote pattern -------------------------------------------------------------

# \u65b9\u5411\u4e0d\u660e\u3001\u53ea\u80fd\u9760\u5947\u5076\u914d\u5bf9\u5b9a\u5f00\u95ed\u7684\u5f15\u53f7\uff1a\u76f4\u7acb " \u4e0e\u4e24\u4e2a\u5f2f\u5f15\u53f7\uff08\u53ef\u80fd\u672c\u6765\u5c31\u88c5\u53cd\u4e86\uff09
QUOTE_PATTERN = '["\u201c\u201d]'

# \u76f4\u89d2\u5f15\u53f7\uff08corner brackets\uff09\u300c\u300d\u300e\u300f\u2014\u2014\u65b9\u5411\u81ea\u5e26\uff0c\u786e\u5b9a\u6027\u6620\u5c04\uff0c\u4e0d\u8fdb counter\u3002
# \u8d70 counter \u4f1a\u88ab\u5947\u5076\u7b97\u9519\uff1a\u6587\u4e2d\u51fa\u73b0\u5b64\u7acb " \u65f6\u6574\u4e32\u5f00\u95ed\u7ffb\u8f6c\uff082026-08-25 \u6850\u4e61\u5b9e\u6d4b\uff09\u3002
CORNER_QUOTES = str.maketrans({
    "\u300c": "\u201c", "\u300d": "\u201d",   # \u300c\u300d\u2192 \u201c\u201d
    "\u300e": "\u2018", "\u300f": "\u2019",   # \u300e\u300f\u2192 \u2018\u2019\uff08\u5185\u5c42\u5f15\u53f7\uff09
})
CORNER_QUOTE_PATTERN = re.compile("[\u300c\u300d\u300e\u300f]")


def fix_quotes(text: str, counter: int = 0) -> tuple[str, int, int]:
    """Replace all double quotes with paired Chinese quotes (\u201c \u201d).

    \u76f4\u89d2\u5f15\u53f7\u300c\u300d\u300e\u300f\u4e00\u5e76\u5f52\u4e00\u4e3a\u4e2d\u6587\u5f2f\u5f15\u53f7\uff08\u7528\u6237 2026-08-22 /govern \u94a6\u5b9a\uff1a
    \u5916\u53d1\u4e2d\u6587\u4ef6\u7981\u76f4\u89d2\u5f15\u53f7\uff09\uff0c\u4f46**\u4e0d**\u53c2\u4e0e counter \u5947\u5076\u914d\u5bf9\u2014\u2014\u5b83\u81ea\u5e26\u65b9\u5411\u3002

    Args:
        text: Input text.
        counter: External counter for cross-run quote pairing (DOCX scenarios).

    Returns:
        (result_text, replacement_count, updated_counter)
    """
    count = 0

    def _replace(match):
        nonlocal counter, count
        counter += 1
        new = "\u201c" if counter % 2 == 1 else "\u201d"
        if new != match.group(0):      # \u53ea\u6570\u771f\u6539\u52a8:\u672c\u6765\u5c31\u662f\u6b63\u786e\u4e2d\u6587\u5f15\u53f7\u7684\u4e0d\u8ba1
            count += 1
        return new

    result = re.sub(QUOTE_PATTERN, _replace, text)
    n_corner = len(CORNER_QUOTE_PATTERN.findall(result))
    if n_corner:
        result = result.translate(CORNER_QUOTES)
        count += n_corner
    return result, count, counter


# 半角→全角必须绕开的「机器读的」片段。2026-07-26 实测旧实现的破坏：
#   https://example.com/a?b=1;c=2  →  https：//example.com/a？b=1；c=2   (链接失效)
#   df.groupby(['a','b'])          →  df.groupby（['a'，'b']）           (代码跑不了)
#   Total 1,234 items              →  Total 1，234 items                (千分位坏)
_PROTECTED = re.compile(
    r"""(
        [a-zA-Z][a-zA-Z0-9+.\-]*://\S+       # URL(http/https/ftp/file…)
      | (?:www\.|mailto:)\S+                 # www. 开头 / mailto:
      | [\w.\-+]+@[\w.\-]+\.\w+              # 邮箱
      | ```[\s\S]*?```                       # markdown 围栏代码块(md 是整篇进来的)
      | `[^`\n]+`                            # markdown 行内代码
      | \d{1,3}(?:,\d{3})+(?:\.\d+)?         # 千分位数字 1,234 / 1,234,567.89
    )""",
    re.VERBOSE,
)


def fix_punctuation(text: str) -> tuple[str, int]:
    """Convert English punctuation to Chinese equivalents.

    URL / 邮箱 / 行内代码 / 千分位数字整段跳过 —— 那些半角符号是给机器读的，
    转成全角等于把内容改坏（链接点不开、代码跑不了）。

    Returns:
        (result_text, replacement_count)
    """
    total = 0

    def _convert(chunk: str) -> str:
        nonlocal total
        for eng, chn in PUNCTUATION_MAP.items():
            n = chunk.count(eng)
            if n:
                total += n
                chunk = chunk.replace(eng, chn)
        return chunk

    out = []
    pos = 0
    for m in _PROTECTED.finditer(text):
        out.append(_convert(text[pos:m.start()]))
        out.append(m.group(0))          # 受保护片段原样放回
        pos = m.end()
    out.append(_convert(text[pos:]))
    return "".join(out), total


def fix_units(text: str) -> tuple[str, int]:
    """Convert Chinese unit names to standard symbols (e.g. \u5e73\u65b9\u7c73 -> m\u00b2).

    Returns:
        (result_text, replacement_count)
    """
    result = text
    total = 0
    for pattern, unit_sym in _UNIT_RULES:
        result, n = pattern.subn(lambda m, s=unit_sym: (m.group(1) if m.group(1).isspace() else "") + s, result)
        total += n
    return result, total


def fix_all(text: str, counter: int = 0) -> tuple[str, dict[str, int], int]:
    """Apply all text fixes: quotes, punctuation, and units.

    Convenience wrapper that applies all three transformations in sequence.

    Args:
        text: Input text.
        counter: External quote counter for cross-run pairing.

    Returns:
        (result_text, stats_dict, updated_counter)
        where stats_dict has keys: "quotes", "punctuation", "units"
    """
    result, quote_count, counter = fix_quotes(text, counter)
    result, punct_count = fix_punctuation(result)
    result, unit_count = fix_units(result)
    return result, {"quotes": quote_count, "punctuation": punct_count, "units": unit_count}, counter
