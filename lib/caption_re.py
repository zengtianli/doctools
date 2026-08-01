#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""caption_re.py — 题注／图表编号判据的**唯一**实现（2026-08-01 立）。

纯 stdlib，零第三方依赖 —— 它会被「自己开 zipfile 写 docx」那条路上的脚本 import
（`renum.py` / `bid_gate.py` / `shape_contract.py`），那条路必须能用系统 python3
直接跑，不能要求先进 venv。

══════════════════════════════════════════════════════════════════════════════
为什么统一
══════════════════════════════════════════════════════════════════════════════
合并前全仓 **12 个文件、~30 条 compiled pattern、20 个互相竞争的判据点**，没有任何
两处口径相同。光是「短横」就有 **8 种互不相同的字符子集**，没有一处是全集：

    [-.–—]  [-–—]  [-—–]  [-‑–—]  [-－—–]  [.．-]  [-—]  [-－]

三条实测出来的后果（都是拿真 fixture 走完整 CLI 跑出来的，不是读代码推的）：

① `renum.py figures` **写出双重编号**。renum 的字符类含全角 `－` 但不含 U+2011，
   于是把已编号的 `图1‑2` 当成「无号题注」prepend 了新号，落盘成
   `图1-3 图1‑2 非断行连字符图`。更糟的是自检 `_verify_cn` 用**同一条瞎掉的正则**
   读回，对着被写坏的文档打印「✓ 每节连续」—— fail-open 的守卫。
   而 `bid_gate` 的字符类恰好相反（含 U+2011、不含全角 `－`），两个门互为盲区。

② `docx_cli health diagnose` 对合法文档报 High + exit 2。`附图1` 被非锚定 search
   截成 `图1`，与真 `图1` 撞 key 成「重复图号」。仓内另外四处都显式把附图/附表当
   独立编号族，只有这一处没有。

③ `audit table-pairing` → `caption pair` 这条链在章节式表号上是断的。上游产出
   `number="表3.1-1"`，下游**同名常量** `CAP_PATTERN` 章号只写了 `(\d+)`，匹配不上
   → decision.json 里的条目静默修不了。扁平 `表3-1` 两边都过，所以洞只在
   `--cn-section` 产出的文档上现形。

══════════════════════════════════════════════════════════════════════════════
统一到什么形状：**一套字符类 + 一个构造器 + 若干具名 spec**
══════════════════════════════════════════════════════════════════════════════
**不合成一条正则。** 20 处判据的差异有一半是**有意的**，合成一条会同时制造漏判和
误判：

  · `bid_gate` 的右界断言 `(?=[　\s）】]|$)` —— 防把正文内联「见图3-12的说明」误吃成
    题注。合并时最容易丢的就是这条。
  · `bid_residue` 的非锚定 —— 它扫的是正文交叉引用，不是题注段。
  · `typeset_ops` / `styles._CAPTION_PATTERN` 强制三段 —— 前者在 PDF 文本层只找正文
    图，后者是 `style caption` 的写盘范围，放宽 = 给既有动词偷偷加破坏性范围。
  · `table.py` / `image.py` 只看首字 + `len<80` —— 有意的兜底，要给**无编号题注**
    命名。收紧成要求编号，抽图/切表会大面积掉回 `table-{idx:02d}` fallback，违反
    CLAUDE.md §抽取类工具默认契约第 ① 条。

所以真正统一的是**字符类**（那层是纯 bug），结构差异改为**显式声明**在 `CaptionSpec`
里 —— 从「谁记得写就有」变成「写在 spec 上、看得见、测得到」。

══════════════════════════════════════════════════════════════════════════════
本模块**不合并**什么（有意留多份）
══════════════════════════════════════════════════════════════════════════════
**样式名判据三套并存**，只搬进来共处一室，不取并集：
`is_fig_caption_style`（shape_contract，中英双线正则）· `is_caption_family_style`
（audit，含 `0图/0表/zdwp` 院模板变体）· `CAPTION_STYLES_EXACT`（caption，5 个字面量）。
取并集会让 `caption number` 给更多段落写编号 —— 那是**写盘动词的范围扩大**，不是
判据统一。三者各有对方缺的样式名是事实，合并它们得单独一轮、单独对拍。

`FIG_KEYWORDS_CORE` ⊂ `FIG_KEYWORDS_EXT` 是**真子集关系**（9 词全在 13 词里），所以
这一对合并成「核心 + 扩展」零行为变化，是本轮唯一动了的词表。

用法::

    from caption_re import parse, finditer, pattern, BID_STRICT_CAPTION
    n = parse("图3-1 灌区分布图", RENUM_CN_CAPTION.for_kind("图"))
    n.section, n.seq            # ("3", 1)
    pattern(PREFIX_STRIP_FIG).sub("", text)     # 直接拿 compiled 用 sub/match
"""
from __future__ import annotations

import re
from dataclasses import dataclass, replace
from functools import lru_cache
from typing import Iterator, Optional

__all__ = [
    "DASH_CHARS", "DOT_CHARS", "CN_SEQ_CHARS", "DASH", "DOT", "SEP",
    "CaptionSpec", "CaptionNum", "pattern", "parse", "finditer",
    "md_token_pattern", "en_caption_pattern", "en_citation_pattern",
    "NUM_TOKEN_ANY", "NUM_SPLIT_LOOSE", "CAPTION_ANY_CN",
    "TABLE_CAPTION_LINE", "TABLE_CAPTION_APPENDIX", "TABLE_CAPTION_FLAT_SEC",
    "HAS_NUM_FIG", "HAS_NUM_TABLE", "FIG_APPENDIX_PREFIX",
    "SECTIONED_CAPTION", "TYPESET_FIG_SECTIONED", "TYPESET_FIG_APPENDIX",
    "TABLE_NAME_HEURISTIC", "PREFIX_STRIP_FIG", "PREFIX_STRIP_TBL",
    "IMG_FIG_PREFIX", "IMG_TBL_PREFIX", "KIND_PREFIX_FIG", "KIND_PREFIX_TABLE",
    "BID_STRICT_CAPTION", "BID_REF_LOOSE",
    "SHAPE_FIG_NUM", "SHAPE_TBL_NUM", "RENUM_CN_CAPTION",
    "RENUM_CN_REF", "APPENDIX_PREFIX_ANY",
    "FIG_KEYWORDS_CORE", "FIG_KEYWORDS_EXT", "TABLE_KEYWORDS",
    "CAPTION_STYLES_EXACT", "CAPTION_FAMILY_KEYWORDS",
    "is_caption_family_style", "is_fig_caption_style", "is_table_caption_style",
]

# ═══════════════════════════════════════════════════════════════════════════
# 字符类 —— 唯一真正被强制统一的一层
# ═══════════════════════════════════════════════════════════════════════════

#: 五种短横，全收。合并前 8 种子集没有一处是全集，实测双重编号与假重复都出在这里。
#: `-`=ASCII HYPHEN-MINUS · `‑`=U+2011 NON-BREAKING HYPHEN · `–`=U+2013 EN DASH
#: `—`=U+2014 EM DASH · `－`=U+FF0D FULLWIDTH HYPHEN-MINUS
DASH_CHARS = "-\u2011\u2013\u2014\uff0d"

#: 半角 + 全角句点（章节号内分隔 `3.1` / `3．1`）。
DOT_CHARS = ".\uff0e"

#: 编号里允许的中文数字（`表一-1`）。**只有 blocks / styles 两处表名启发式用**。
#: 有意**不**复用 `cn_number.CN_CHARS`（那份含 `〇两千`）—— 与 cn_number 模块
#: docstring「别在这次悄悄加」同一纪律：这里取的是两处旧实现的并集，不做增量。
CN_SEQ_CHARS = "一二三四五六七八九十百零"

DASH = f"[{re.escape(DASH_CHARS)}]"
DOT = f"[{re.escape(DOT_CHARS)}]"
#: 「点或短横」的合并字符类。编号内分隔符不区分点/横的那些判据用它
#: （audit 的 `[\.\-]`、bid_residue 的 `[\.\-－]` 原本各写各的子集）。
SEP = f"[{re.escape(DOT_CHARS + DASH_CHARS)}]"

_EN_FIG = r"Figure|Fig\.?"
_EN_TBL = r"Table"

#: 右界断言：题注编号后面必须是空白/中文空格/收尾括号/行尾。
#: 有意为之 —— 防把正文内联「见图3-12的说明」误吃成题注（bid_gate 独有能力）。
_RIGHT_BOUNDARY = r"(?=[\u3000\s）】]|$)"


# ═══════════════════════════════════════════════════════════════════════════
# spec
# ═══════════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class CaptionSpec:
    """一个调用点的判据形状。**差异写在这里，不写在各自的正则字面量里。**

    Attributes:
        kinds: 认哪些中文种类字，`("图",)` / `("表",)` / `("图","表")`。
        mode: ``"split"`` 拆出 章号/序号（要结构化结果的用它）·
            ``"flat"`` 只认「数字(分隔符数字)*」这一串不拆（判真假/取 token/剥前缀用）·
            ``"prefix"`` 只认种类字本身，完全不看编号（table/image 的宽松兜底）。
        anchored: 是否 ``^\\s*`` 锚定。非锚定 = 扫正文引用。
        allow_en: 是否同时认 ``Figure/Fig./Table``。
        appendix: ``"exclude"`` 附图/附表不认（并加 ``(?<!附)`` 防非锚定 search 截半）·
            ``"include"`` 附图 走独立扁平分支 · ``"only"`` 只认附图。
        sec_min_depth / sec_max_depth: split 模式下章号几段（``2,2`` = 强制 ``3.1``）。
            ``sec_max_depth=0`` 表示不限。
        require_seq: split 模式下是否必须有「短横 + 序号」。
        cn_digits: 编号是否允许中文数字。
        right_boundary: 是否加右界断言（见 ``_RIGHT_BOUNDARY``）。
        trailing: 尾随消费，``""`` / ``r"\\s*"``（剥前缀）/ ``r"\\s"``（表名启发式）。
        flat_min_seg / flat_max_seg: flat 模式下「分隔符+数字」重复几次。
        appendix_sectioned: 附图分支也走「章号-序号」结构（默认 False = 扁平单号）。
            只有 ``BID_REF_LOOSE`` 打开 —— 标书正文写的是「见附图3-1」，扁平分支
            只能吃到 `附图3`，分组标签会串成「图 None-」且丢掉 -1/-3 的断号。
            默认关是因为其余附图消费点（shape_contract 快照 / renum 排除闸 /
            typeset 附图）本来就把附图当扁平单号，改它们要单独一轮。
    """

    kinds: tuple = ("图", "表")
    mode: str = "split"
    anchored: bool = True
    allow_en: bool = False
    appendix: str = "exclude"
    sec_min_depth: int = 1
    sec_max_depth: int = 0
    require_seq: bool = True
    cn_digits: bool = False
    right_boundary: bool = False
    trailing: str = ""
    flat_min_seg: int = 0
    flat_max_seg: int = 0
    appendix_sectioned: bool = False
    flags: int = 0

    def for_kind(self, kind: str) -> "CaptionSpec":
        """派生出只认单一种类字的 spec（`renum` 按 ``--kind 图|表`` 现造）。"""
        return replace(self, kinds=(kind,))


@dataclass(frozen=True)
class CaptionNum:
    """一次命中的结构化结果。

    ``section`` / ``seq`` 只有 split 模式给得出；flat / prefix 模式恒为 ``None``，
    这两档的调用点本来就只用 ``raw``（当 token / 当布尔 / 拿去 sub）。
    """

    kind: str                 # "图" | "表"（英文 Figure/Fig./Table 已归一到中文）
    lang: str                 # "cn" | "en"
    appendix: bool            # 附图/附表 → True，此时 section 恒为 None
    section: Optional[str]    # "3" | "3.1" | "3.1.2"；扁平/单号时 None
    seq: Optional[int]
    raw: str                  # 原文片段。回写定位用它，别拿归一化串去 replace
    start: int
    end: int


# ═══════════════════════════════════════════════════════════════════════════
# 构造
# ═══════════════════════════════════════════════════════════════════════════

def _num_atom(spec: CaptionSpec) -> str:
    chars = r"\d" + (CN_SEQ_CHARS if spec.cn_digits else "")
    return f"[{chars}]+"


def _rep(lo: int, hi: Optional[int]) -> str:
    """重复次数量词。``hi=None`` = 不限。

    ⚠ 上界必须用 ``None`` 表示「不限」，**不能用 0** —— spec 里 ``sec_max_depth=0``
    是「不限」的哨兵，减 1 之后也是 0，两种含义会撞在一起：``sec_max_depth=1``
    （只许单段章号、不许小数点）会被静默编译成不限。变异测试实证抓到过。
    """
    if hi is None:
        return "*" if lo == 0 else f"{{{lo},}}"
    return f"{{{lo},{hi}}}"


def _kind_atom(spec: CaptionSpec, group: str) -> str:
    alts = list(spec.kinds)
    if spec.allow_en:
        # prefix 模式后面不跟数字，英文词必须自带右词界，否则 `Tablet` / `Figment` 命中
        wb = r"\b" if spec.mode == "prefix" else ""
        if "图" in spec.kinds:
            alts.append(f"(?:{_EN_FIG}){wb}")
        if "表" in spec.kinds:
            alts.append(f"{_EN_TBL}{wb}")
    return f"(?P<{group}>{'|'.join(alts)})"


def _build(spec: CaptionSpec) -> str:
    head = r"^\s*" if spec.anchored else ""
    num = _num_atom(spec)
    branches = []

    # ── 附图/附表 分支：扁平单号，无章节概念 ──────────────────────────────
    if spec.appendix in ("include", "only"):
        akind = _kind_atom(spec, "akind")
        # prefix 模式只认「附+种类字」本身，不要求跟编号（renum 的附图排除闸就是这档）
        if spec.mode == "prefix":
            branches.append(f"附{akind}")
        elif spec.appendix_sectioned and spec.mode == "split":
            asec = f"(?P<asec>{num}(?:{DOT}{num})*)"
            branches.append(f"附{akind}\\s*{asec}\\s*{DASH}\\s*(?P<aseq>{num})")
        else:
            branches.append(f"附{akind}\\s*(?P<aseq>{num})")

    # ── 正常分支 ────────────────────────────────────────────────────────
    if spec.appendix != "only":
        # 附 前缀排除：非锚定 search 时防把「附图1」截成「图1」（health 假重复的根因）
        guard = "(?<!附)"
        kind = _kind_atom(spec, "kind")
        if spec.mode == "prefix":
            body = f"{guard}{kind}"
        elif spec.mode == "flat":
            sep = f"(?:{DOT}|{DASH})"
            rep = _rep(spec.flat_min_seg,
                       None if spec.flat_max_seg == 0 else spec.flat_max_seg)
            body = (f"{guard}{kind}\\s*(?P<flat>{num}"
                    f"(?:\\s*{sep}\\s*{num}){rep})")
        else:  # split
            sec_rep = _rep(spec.sec_min_depth - 1,
                           None if spec.sec_max_depth == 0
                           else spec.sec_max_depth - 1)
            sec = f"(?P<sec>{num}(?:{DOT}{num}){sec_rep})"
            tail = f"\\s*{DASH}\\s*(?P<seq>{num})"
            body = (f"{guard}{kind}\\s*{sec}"
                    + (tail if spec.require_seq else f"(?:{tail})?"))
        branches.append(body)

    pat = head + ("(?:" + "|".join(branches) + ")" if len(branches) > 1
                  else branches[0])
    if spec.right_boundary:
        pat += _RIGHT_BOUNDARY
    return pat + spec.trailing


@lru_cache(maxsize=None)
def pattern(spec: CaptionSpec) -> re.Pattern:
    """spec → compiled pattern（缓存）。直接拿去 ``.match`` / ``.sub`` / ``.findall``。"""
    return re.compile(_build(spec), spec.flags)


def _to_num(m: re.Match, spec: CaptionSpec) -> CaptionNum:
    gd = m.groupdict()
    if gd.get("akind") is not None:
        raw_kind = gd["akind"]
        seq = gd.get("aseq")
        asec = gd.get("asec")
        return CaptionNum(
            kind=_norm_kind(raw_kind), lang=_lang(raw_kind), appendix=True,
            section=asec.replace("\uff0e", ".") if asec else None,
            seq=_int_or_none(seq), raw=m.group(0),
            start=m.start(), end=m.end())
    raw_kind = gd.get("kind") or ""
    sec = gd.get("sec")
    return CaptionNum(
        kind=_norm_kind(raw_kind), lang=_lang(raw_kind), appendix=False,
        section=sec.replace("\uff0e", ".") if sec else None,
        seq=_int_or_none(gd.get("seq")), raw=m.group(0),
        start=m.start(), end=m.end())


def _norm_kind(raw: str) -> str:
    low = raw.lower().rstrip(".")
    if low in ("figure", "fig"):
        return "图"
    if low == "table":
        return "表"
    return raw


def _lang(raw: str) -> str:
    return "cn" if raw in ("图", "表") else "en"


def _int_or_none(s) -> Optional[int]:
    if s is None:
        return None
    try:
        return int(s.replace("\uff0e", ".").rstrip("."))
    except (ValueError, AttributeError):
        return None


def parse(text: str, spec: CaptionSpec) -> Optional[CaptionNum]:
    """锚定解析（``.match`` 语义）。题注段用。解析不出返 ``None``。"""
    m = pattern(spec).match(text)
    return _to_num(m, spec) if m else None


def finditer(text: str, spec: CaptionSpec) -> Iterator[CaptionNum]:
    """非锚定扫描（``.finditer`` 语义）。正文交叉引用用。"""
    for m in pattern(spec).finditer(text):
        yield _to_num(m, spec)


# ═══════════════════════════════════════════════════════════════════════════
# 具名 spec —— 一个 spec 对应一类调用点，差异在这里被**声明**出来
# ═══════════════════════════════════════════════════════════════════════════

#: health `duplicate-figure-numbers` 判重 key。中英双线 + 单号「图3」 + 附图独立族。
NUM_TOKEN_ANY = CaptionSpec(
    mode="flat", anchored=False, allow_en=True, appendix="include",
    flat_min_seg=0, flags=re.I)

#: health 跳号分析：从 shape_contract 产出的编号集里拆 (章号, 序号)。
NUM_SPLIT_LOOSE = CaptionSpec(mode="split", anchored=False)

#: audit `captions`：题注段布尔判据。点与短横同类，允许 1–2 段。
CAPTION_ANY_CN = CaptionSpec(
    mode="flat", flat_min_seg=1, flat_max_seg=2)

#: audit `table-pairing` 产出 + caption `pair` 回读 —— **同一个对象**，
#: 合并前是同名常量各写各的（下游更窄，章节式表号整条链断掉）。
TABLE_CAPTION_LINE = CaptionSpec(kinds=("表",))
TABLE_CAPTION_APPENDIX = CaptionSpec(kinds=("表",), appendix="only")

#: caption `pair` 的 **renumber-all-tables 专用**：章号只认单段（`表3-1` 认、
#: `表3.1-1` 不认）。为什么不跟着 TABLE_CAPTION_LINE 一起放宽：那是个**写盘**动词，
#: 放宽 = 给既有动词偷偷扩大改写范围。实测（2026-08-01，fixture 两条 `表3.1-N`）：
#: 用放宽的 spec 时 dry-run 从 `count=0` 变成 `count=2`，把 `表3.1-1 → 表1-1`
#: 整体压平 —— 对 `--cn-section` 产出的章节式文档就是把编号体系冲掉，且默认 --apply 落盘。
#: 读侧（rename-caption / audit table-pairing 回读）该放宽，写侧不该，两者就此解耦。
TABLE_CAPTION_FLAT_SEC = replace(TABLE_CAPTION_LINE, sec_max_depth=1)

#: caption `number` / styles `number-by-style`：段首已有编号则跳过。
HAS_NUM_FIG = CaptionSpec(kinds=("图",), mode="flat", flat_min_seg=1, flat_max_seg=1)
HAS_NUM_TABLE = CaptionSpec(kinds=("表",), mode="flat", flat_min_seg=1, flat_max_seg=1)

#: styles `number-by-style`：附图 = 独立扁平族，直接 return None。
FIG_APPENDIX_PREFIX = CaptionSpec(kinds=("图",), appendix="only")

#: styles `style caption`：**有意只认章节式三段**。放宽 = 给写盘动词偷加范围。
SECTIONED_CAPTION = CaptionSpec(sec_min_depth=2, sec_max_depth=2)

#: typeset `shot-center-images`：PDF 文本层挑含图页，同样有意只认三段。
TYPESET_FIG_SECTIONED = CaptionSpec(
    kinds=("图",), anchored=False, sec_min_depth=2, sec_max_depth=2)
TYPESET_FIG_APPENDIX = CaptionSpec(kinds=("图",), anchored=False, appendix="only")

#: blocks `detect_heading_form` + styles `style body`：表名启发式（认中文数字号）。
TABLE_NAME_HEURISTIC = CaptionSpec(
    kinds=("表",), cn_digits=True, require_seq=False, trailing=r"\s")

#: styles `renumber h4-figures`：剥前缀。零段也认（`图3 灌区图`）。
PREFIX_STRIP_FIG = CaptionSpec(
    kinds=("图",), mode="flat", cn_digits=True, trailing=r"\s*")
PREFIX_STRIP_TBL = CaptionSpec(
    kinds=("表",), mode="flat", cn_digits=True, trailing=r"\s*")

#: image `extract --by-caption`：剥前缀 + `has_fig_prefix` 判题注。
#: **必须有短横+序号**（比 PREFIX_STRIP_* 窄）—— 它同时决定「哪个段算图题」，
#: 放宽到零段会把 `图书馆…` 之外的任何「图+数字」段都收进来。
IMG_FIG_PREFIX = CaptionSpec(kinds=("图",), trailing=r"\s*")
IMG_TBL_PREFIX = CaptionSpec(kinds=("表",), trailing=r"\s*")

#: table `split` / image `extract`：**只看种类字**，完全不看编号。
#: 有意的宽松兜底 —— 收紧会让无编号题注掉回 `table-{idx:02d}` fallback。
KIND_PREFIX_FIG = CaptionSpec(kinds=("图",), mode="prefix")
KIND_PREFIX_TABLE = CaptionSpec(kinds=("表",), mode="prefix", allow_en=True, flags=re.I)

#: bid_gate fatal③ 题注编号断裂：**唯一带右界断言的一档**，防吃正文内联引用。
BID_STRICT_CAPTION = CaptionSpec(right_boundary=True)

#: bid_residue cat6 断裂引用：**有意非锚定** —— 扫的是正文交叉引用不是题注段。
#: `appendix="include"` 是必须的：默认 `"exclude"` 会带上 `(?<!附)`，于是正文里的
#: 「见附图3-1…见附图3-3」整条看不见 —— 实测这道交付门因此从 FAIL 3 findings 掉成
#: FAIL 2，**类别6 检出直接消失**（2026-08-01 前后对拍）。include 让附图自成一组做
#: 连续性核：既不再把 `附图3-1` 截成 `图3-1` 混进正图组，也不丢掉附图自己的断号。
BID_REF_LOOSE = CaptionSpec(anchored=False, appendix="include", appendix_sectioned=True)

#: shape_contract 快照：题注文本 → 编号集合（fix_styleset 前后对账的基准）。
SHAPE_FIG_NUM = CaptionSpec(kinds=("图",), anchored=False, appendix="include")
SHAPE_TBL_NUM = CaptionSpec(kinds=("表",), anchored=False, appendix="include")

#: renum `figures --cn-section`：写入端 / 回写端 / 自检端**共用**它（用 `for_kind`）。
#: 合并前这三端各写一条 rf-string，字符类漏 U+2011 → 已编号的 `图1‑2` 被当无号题注
#: prepend 出 `图1-3 图1‑2 …`，而自检因同样看不见 U+2011 而打印「✓ 每节连续」。
RENUM_CN_CAPTION = CaptionSpec()

#: renum 正文交叉引用（非锚定，同一 spec 的 loose 版）。
RENUM_CN_REF = CaptionSpec(anchored=False)

#: renum 排除附图/附表（它们是独立扁平族，不属 `图X.Y-N` 范畴）。
APPENDIX_PREFIX_ANY = CaptionSpec(mode="prefix", appendix="only")


# ═══════════════════════════════════════════════════════════════════════════
# 非 spec 形状的两份（分组语义被调用点按位置消费，单列）
# ═══════════════════════════════════════════════════════════════════════════

@lru_cache(maxsize=None)
def md_token_pattern() -> re.Pattern:
    """renum `tabfig`：md 侧「表/图 章号-序号」token。

    **4 个位置分组**（kind, 空白, 章号, 短横+序号）被 `align_text` 的 `sub` 回调
    按位置解包重组，所以不能改成具名组的 spec 形状 —— 单列一份，但字符类照样
    走本模块的 `DASH` / `DOT`。
    """
    num = r"\d+"
    return re.compile(
        f"([表图])(\\s*)({num}(?:{DOT}{num})*)({DASH}{num})")


@lru_cache(maxsize=None)
def en_caption_pattern(prefix: str) -> re.Pattern:
    """renum `figures` 英文线题注：`Figure 3` / `Fig. 3`，扁平单号无章节概念。

    英文线与中文线**零共享**是有意的：混排文档里两套编号各管各的，合流会出现
    同一段被两套逻辑同时认领。
    """
    return re.compile(rf"^\s*(?:{prefix}|Fig\.?)\s*(\d+)\b", re.I)


@lru_cache(maxsize=None)
def en_citation_pattern(prefix: str) -> re.Pattern:
    """renum `figures` 英文线正文引用：`Figures 3–5`、`Figs. 3, 4 and 5`。"""
    return re.compile(
        rf"(?:{prefix}s?|Figs?\.?)\s*\d+(?:\s*(?:{DASH}|,|and)\s*\d+)*", re.I)


# ═══════════════════════════════════════════════════════════════════════════
# 词表 / 样式名 —— 见模块 docstring「本模块不合并什么」
# ═══════════════════════════════════════════════════════════════════════════

#: caption `number` 的关键词兜底（9 词）。
FIG_KEYWORDS_CORE = (
    "示意图", "分布图", "结构图", "流程图", "对比图",
    "平面图", "布局图", "线路图", "断面图",
)

#: styles `number-by-style` 的关键词兜底 = CORE 的**真超集**（多 4 词）。
#: 合并前是两份手抄，CORE 的 9 词全在这 13 词里，所以这一对合成「核心+扩展」
#: 零行为变化 —— 本模块唯一动了的词表。
FIG_KEYWORDS_EXT = FIG_KEYWORDS_CORE + ("布置图", "范围图", "管线", "输水")

#: 表关键词。全仓唯一一份（caption 侧没有对应物，判表名只能靠邻接 tbl）。
TABLE_KEYWORDS = (
    "统计表", "配置表", "余缺水量表", "比选表", "拟定表",
    "计算结果表", "结果表", "情况", "对照表", "汇总表",
)

#: caption `number` 认的样式名（**精确集合**，不是关键词包含）。
CAPTION_STYLES_EXACT = {"zdwp表名", "Caption", "caption", "表题", "图题"}

#: audit 认的 caption 族样式名（关键词包含，含院模板 `0图/0表/zdwp` 变体）。
CAPTION_FAMILY_KEYWORDS = (
    "caption", "表名", "图名", "0图", "0表", "zdwp 表名", "zdwp 图名")

_FIG_CAPTION_NAME_RE = re.compile(
    r"(图\s*名|图名称|figure\s*caption|image\s*caption|图\s*注|^figure$)", re.I)
_TABLE_CAPTION_NAME_RE = re.compile(
    r"(表\s*名|表名称|表格标题|table\s*caption|表\s*注|表\s*题|^table$)", re.I)


def is_caption_family_style(style_name: str) -> bool:
    """audit 档：关键词包含式。**与下面两个不是同一口径，有意不合并。**"""
    if not style_name:
        return False
    low = style_name.lower().strip()
    return any(k in low for k in CAPTION_FAMILY_KEYWORDS)


def is_fig_caption_style(style_name: str) -> bool:
    """shape_contract 档：按样式名中英双线正则判图题样式。"""
    return bool(style_name) and bool(_FIG_CAPTION_NAME_RE.search(style_name))


def is_table_caption_style(style_name: str) -> bool:
    """shape_contract 档：按样式名中英双线正则判表题样式。"""
    return bool(style_name) and bool(_TABLE_CAPTION_NAME_RE.search(style_name))
