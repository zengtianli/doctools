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

# -- Quote pattern -------------------------------------------------------------

QUOTE_PATTERN = '[""\u201c\u201d\u300c\u300d]'


def fix_quotes(text: str, counter: int = 0) -> tuple[str, int, int]:
    """Replace all double quotes with paired Chinese quotes (\u201c \u201d).

    Args:
        text: Input text.
        counter: External counter for cross-run quote pairing (DOCX scenarios).

    Returns:
        (result_text, replacement_count, updated_counter)
    """
    count = len(re.findall(QUOTE_PATTERN, text))

    def _replace(match):
        nonlocal counter
        counter += 1
        return "\u201c" if counter % 2 == 1 else "\u201d"

    result = re.sub(QUOTE_PATTERN, _replace, text)
    return result, count, counter


def fix_punctuation(text: str) -> tuple[str, int]:
    """Convert English punctuation to Chinese equivalents.

    Returns:
        (result_text, replacement_count)
    """
    result = text
    total = 0
    for eng, chn in PUNCTUATION_MAP.items():
        escaped = re.escape(eng)
        n = len(re.findall(escaped, result))
        total += n
        result = re.sub(escaped, chn, result)
    return result, total


def fix_units(text: str) -> tuple[str, int]:
    """Convert Chinese unit names to standard symbols (e.g. \u5e73\u65b9\u7c73 -> m\u00b2).

    Returns:
        (result_text, replacement_count)
    """
    result = text
    total = 0
    for unit_cn, unit_sym in _UNITS_SORTED:
        n = result.count(unit_cn)
        total += n
        result = result.replace(unit_cn, unit_sym)
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
