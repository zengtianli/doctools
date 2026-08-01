"""tests for lib/caption_re.py —— 题注/图表编号判据的 SSOT 回归门。

合并前全仓 12 个文件、~30 条 compiled pattern、20 个互相竞争的判据点，光短横就有
**8 种互不相同的字符子集**，没有一处是全集。三条实测后果：renum 写出双重编号（且
自检 fail-open 打印「✓」）· health 对合法文档报 High + exit 2 · audit→caption 这条链
在章节式表号上静默断掉。本文件是它的门。

**这个门管的是判据，`cli_surface` / `cli_forward_probe` 管不到** —— 那两个只看 argv，
正则怎么变它们都是绿的。
"""
import re
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "lib"))

import caption_re as C  # noqa: E402


# ═══════════════════════════════════════════════════════════════════════════
# 字符类：唯一被强制统一的一层
# ═══════════════════════════════════════════════════════════════════════════

#: 五种短横。合并前没有任何一处认全（bid_gate 有 U+2011 无全角，renum 反过来）。
ALL_DASHES = ["-", "‑", "–", "—", "－"]


def test_dash_class_covers_all_five():
    assert set(ALL_DASHES) == set(C.DASH_CHARS), "短横字符类必须是五种的全集"


def test_dot_class_covers_halfwidth_and_fullwidth():
    assert set(C.DOT_CHARS) == {".", "．"}


@pytest.mark.parametrize("dash", ALL_DASHES)
def test_every_dash_recognized_by_split_specs(dash):
    """五种短横在**每个** split 档 spec 下结论必须一致 —— 这正是合并前不成立的那条。"""
    text = f"图3{dash}1 灌区分布图"
    for name in ("RENUM_CN_CAPTION", "BID_STRICT_CAPTION", "SHAPE_FIG_NUM"):
        spec = getattr(C, name)
        n = next(C.finditer(text, spec), None) if not spec.anchored \
            else C.parse(text, spec)
        assert n is not None, f"{name} 认不出 {text!r}"
        assert (n.section, n.seq) == ("3", 1), f"{name} 解析错 {text!r}"


@pytest.mark.parametrize("dash", ALL_DASHES)
def test_every_dash_stripped_completely_by_prefix_specs(dash):
    """剥前缀必须**整段剥掉**。旧实现漏 U+2011 时剥出 `‑1 灌区分布图` 这种垃圾。"""
    text = f"图3{dash}1 灌区分布图"
    assert C.pattern(C.PREFIX_STRIP_FIG).sub("", text) == "灌区分布图"
    assert C.pattern(C.IMG_FIG_PREFIX).sub("", text) == "灌区分布图"


# ═══════════════════════════════════════════════════════════════════════════
# 附图/附表 = 独立扁平编号族（health 假重复 + exit 2 的根因）
# ═══════════════════════════════════════════════════════════════════════════

def test_appendix_not_truncated_into_plain_kind():
    """`附图2` 非锚定 search 绝不能被截成 `图2` 去和真 `图2` 撞 key。"""
    n = next(C.finditer("附图2 总平面布置图", C.NUM_TOKEN_ANY))
    assert n.appendix is True
    assert n.raw == "附图2", "旧实现在这里给出 '图2' → 与真图2 撞 key 报假重复"
    assert n.section is None, "附图是扁平族，没有章号"
    assert n.seq == 2


def test_appendix_and_plain_get_distinct_keys():
    keys = {next(C.finditer(t, C.NUM_TOKEN_ANY)).raw
            for t in ("附图2 总平面布置图", "图2 系统概化图")}
    assert keys == {"附图2", "图2"}, "两者必须是两个 key，否则 health 报假重复"


def test_appendix_excluded_where_it_should_be():
    """排除档必须真排除 —— `(?<!附)` 在非锚定下也得挡住。"""
    assert C.parse("附图2 总平面布置图", C.RENUM_CN_CAPTION) is None
    assert next(C.finditer("附表3 设备清单", C.NUM_SPLIT_LOOSE), None) is None


# ═══════════════════════════════════════════════════════════════════════════
# 断掉的那条链：audit table-pairing → caption pair
# ═══════════════════════════════════════════════════════════════════════════

def test_audit_and_caption_share_the_same_spec_object():
    """结构性断言：两处不是「碰巧写得一样」，是**同一个对象**。

    合并前它们是同名常量各写各的，下游章号只写 `(\\d+)`，`表3.1-1` 匹配不上。
    """
    sys.path.insert(0, str(ROOT / "scripts" / "document"))
    from sub import audit, caption
    assert audit.CAP_SPEC is caption.CAP_SPEC is C.TABLE_CAPTION_LINE


@pytest.mark.parametrize("text,section,seq", [
    ("表3-1 水量统计表", "3", 1),
    ("表 3.1-2 成本对照表", "3.1", 2),
    ("表3.1.2-4 深层表", "3.1.2", 4),
])
def test_sectioned_table_numbers_parse_at_any_depth(text, section, seq):
    n = C.parse(text, C.TABLE_CAPTION_LINE)
    assert n is not None, f"{text!r} 匹配不上 → caption pair 会静默 no-op"
    assert (n.section, n.seq) == (section, seq)


def test_caption_name_extraction_matches_legacy_tail_group():
    """名字靠 `text[n.end:]` 取，必须与旧实现的 `(.*)$` 分组逐字节一致。"""
    for text, want in [("表3-1 水量统计表", "水量统计表"),
                       ("表 3.1-2 成本对照表", "成本对照表"),
                       ("表7-2  双空格名", "双空格名"),
                       ("表3-1", "")]:
        n = C.parse(text, C.TABLE_CAPTION_LINE)
        assert text[n.end:].strip() == want


# ═══════════════════════════════════════════════════════════════════════════
# 有意保留的窄档 —— 放宽了就是给写盘动词偷加范围，这些 assert 是防止那件事
# ═══════════════════════════════════════════════════════════════════════════

def test_sectioned_only_specs_still_reject_flat_numbers():
    """`style caption` / `shot-center-images` **有意只认三段**，不许被「顺手放宽」。"""
    assert C.parse("表3-1 水量统计表", C.SECTIONED_CAPTION) is None
    assert C.parse("表3.1-2 成本对照表", C.SECTIONED_CAPTION) is not None
    assert C.pattern(C.TYPESET_FIG_SECTIONED).search("图3-1 灌区图") is None
    assert C.pattern(C.TYPESET_FIG_SECTIONED).search("图3.1-2 灌区图") is not None


def test_loose_prefix_specs_still_accept_unnumbered_captions():
    """table split / image extract 的兜底：**无编号题注也要认**，否则命名掉回 fallback。"""
    assert C.pattern(C.KIND_PREFIX_TABLE).match("表 主要设备清单") is not None
    assert C.pattern(C.KIND_PREFIX_FIG).match("图 灌区总体布置") is not None
    assert C.pattern(C.KIND_PREFIX_TABLE).match("Table of contents") is not None


def test_loose_prefix_still_word_bounded_for_english():
    """`Table\\b` 的词界不能在搬家时丢 —— 否则 `Tablet` 被当表题。"""
    assert C.pattern(C.KIND_PREFIX_TABLE).match("Tablet 电脑清单") is None


def test_bid_gate_right_boundary_preserved():
    """全仓唯一的右界断言：防把正文内联引用误吃成题注。合并时最容易丢的能力。

    ⚠ 判据必须用**段首就是编号、但后面紧跟正文**的串（`图3-12的说明`）。
    拿 `见图3-12的说明` 测是无效的 —— 那条被 `^\\s*` 锚定挡掉，右界断言删了照样过，
    变异测试实证：删掉 right_boundary 时该 assert 仍然绿。
    """
    assert C.parse("图3-12的说明见附录", C.BID_STRICT_CAPTION) is None, "右界断言丢了"
    assert C.parse("图3-12）", C.BID_STRICT_CAPTION) is not None, "收尾全角括号是合法右界"
    assert C.parse("图3-12　灌区图", C.BID_STRICT_CAPTION) is not None, "中文空格是合法右界"
    assert C.parse("图3-12 灌区分布图", C.BID_STRICT_CAPTION) is not None
    assert C.parse("图3-12", C.BID_STRICT_CAPTION) is not None
    assert C.parse("见图3-12的说明", C.BID_STRICT_CAPTION) is None  # 锚定挡掉


def test_bid_ref_loose_stays_non_anchored():
    """bid_residue 有意非锚定 —— 它扫的是正文交叉引用，不是题注段。"""
    hits = list(C.finditer("详见图3-1 与 表4-2 的对比", C.BID_REF_LOOSE))
    assert [(h.kind, h.section, h.seq) for h in hits] == [("图", "3", 1), ("表", "4", 2)]


def test_english_line_stays_separate():
    """中英线零共享是有意的：合流会让同一段被两套编号逻辑同时认领。"""
    assert C.en_caption_pattern("Figure").match("Figure 3 Layout") is not None
    assert C.en_caption_pattern("Figure").match("图3-1 灌区图") is None
    assert C.parse("Figure 3 Layout", C.RENUM_CN_CAPTION) is None


# ═══════════════════════════════════════════════════════════════════════════
# 干扰项：统一字符类不许把这些变成命中（放宽最容易在这里翻车）
# ═══════════════════════════════════════════════════════════════════════════

NOISE = ["图书馆藏书情况统计", "表面粗糙度检测结果", "2026-07-31 完成初稿",
         "图纸会审记录", "表决结果公示"]


@pytest.mark.parametrize("text", NOISE)
def test_noise_never_matches_numbering_specs(text):
    for name in ("RENUM_CN_CAPTION", "TABLE_CAPTION_LINE", "SECTIONED_CAPTION",
                 "BID_STRICT_CAPTION", "CAPTION_ANY_CN", "HAS_NUM_FIG",
                 "HAS_NUM_TABLE", "IMG_FIG_PREFIX"):
        assert C.parse(text, getattr(C, name)) is None, f"{name} 误吃干扰项 {text!r}"


@pytest.mark.parametrize("text", NOISE)
def test_noise_never_matches_loose_scan(text):
    assert list(C.finditer(text, C.BID_REF_LOOSE)) == []
    assert list(C.finditer(text, C.NUM_SPLIT_LOOSE)) == []


# ═══════════════════════════════════════════════════════════════════════════
# 中文数字 / 表名启发式
# ═══════════════════════════════════════════════════════════════════════════

def test_table_name_heuristic_accepts_cn_digits_and_dotted_sections():
    pat = C.pattern(C.TABLE_NAME_HEURISTIC)
    assert pat.match("表一-1 水量统计表 ") is not None
    assert pat.match("表 3.1-2 成本对照表") is not None, \
        "旧 blocks.RE_TABLE 章号不许带点 → --cn-section 表名判成未知形态"
    assert pat.match("表3 概况 ") is not None, "序号可缺省"


def test_depth_bound_sentinel_does_not_collapse_to_unlimited():
    """``sec_max_depth=1`` 必须真的只许单段 —— 别和「不限」的哨兵 0 撞在一起。

    构造器最初 `_rep(lo, hi)` 拿 `hi==0` 同时表示「不限」和「减 1 后为 0」，于是
    `sec_max_depth=1` 被静默编译成不限。变异测试抓到的，这条 assert 把它钉死。
    """
    from dataclasses import replace
    narrow = replace(C.TABLE_CAPTION_LINE, sec_max_depth=1)
    assert C.parse("表3-1 水量统计表", narrow) is not None
    assert C.parse("表3.1-2 成本对照表", narrow) is None, "上界失效，退化成不限"
    # 哨兵 0 仍然是「不限」
    assert C.parse("表3.1.2-4 深层表", C.TABLE_CAPTION_LINE) is not None


def test_cn_seq_chars_not_silently_widened_to_cn_number_module():
    """有意不复用 cn_number.CN_CHARS（那份含 〇/两/千）—— 别在这一轮悄悄加。"""
    assert "〇" not in C.CN_SEQ_CHARS
    assert "千" not in C.CN_SEQ_CHARS


# ═══════════════════════════════════════════════════════════════════════════
# 词表：唯一动了的一处，必须是真子集关系（零行为变化的前提）
# ═══════════════════════════════════════════════════════════════════════════

def test_fig_keywords_core_is_strict_subset_of_ext():
    assert set(C.FIG_KEYWORDS_CORE) < set(C.FIG_KEYWORDS_EXT)
    assert len(C.FIG_KEYWORDS_CORE) == 9 and len(C.FIG_KEYWORDS_EXT) == 13


def test_style_name_predicates_deliberately_not_unioned():
    """三套样式名判据**有意不合并** —— 合并会给 caption number 写盘扩范围。

    这条 assert 的作用是：以后谁想「顺手统一」它们，先在这里失败一次并读注释。
    """
    assert C.is_fig_caption_style("图名") and not C.is_fig_caption_style("0图")
    assert C.is_caption_family_style("0图") and not C.is_caption_family_style("图题")
    assert "图题" in C.CAPTION_STYLES_EXACT and "图名" not in C.CAPTION_STYLES_EXACT


# ═══════════════════════════════════════════════════════════════════════════
# renum 自检必须 fail-closed（旧实现对着写坏的文档打印「✓ 每节连续」）
# ═══════════════════════════════════════════════════════════════════════════

def test_double_numbering_is_detectable_by_orthogonal_check():
    """连续性判据在结构上看不见双重编号 —— 必须靠「一段里 ≥2 个编号」这条正交判据。"""
    corrupted = "图1-3 图1‑2 非断行连字符图"
    ref = C.pattern(C.RENUM_CN_REF.for_kind("图"))
    assert len(ref.findall(corrupted)) == 2, "正交判据抓不到双重编号 = 自检又 fail-open"
    # 而写入端判据只看得见开头那个，所以它自己永远发现不了
    head = C.parse(corrupted, C.RENUM_CN_CAPTION.for_kind("图"))
    assert (head.section, head.seq) == ("1", 3)


def test_renum_verify_cn_reports_double_numbering(tmp_path):
    """端到端：把老实现写坏的那种段落放回去，确认新自检真判红。"""
    import importlib.util
    import zipfile
    spec = importlib.util.spec_from_file_location(
        "renum_under_test", ROOT / "scripts" / "document" / "renum.py")
    renum = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(renum)

    src = _make_docx(tmp_path, ["图1-3 图1‑2 非断行连字符图"])
    by_sec, ok, doubled = renum._verify_cn(str(src), "图")
    assert by_sec == {"1": [3]}
    assert doubled, "双重编号没被抓到 —— 自检回到 fail-open"
    assert ok is False

    clean = _make_docx(tmp_path.joinpath("c"), ["图1-1 常规图"])
    by_sec, ok, doubled = renum._verify_cn(str(clean), "图")
    assert doubled == [] and ok is True, "干净文档被误判 = 假红"
    assert zipfile.is_zipfile(clean)


def _make_docx(dirpath: Path, captions: list[str]) -> Path:
    """最小 docx fixture（走 python-docx，收口由 lib/docx_safe_save 自动介入）。"""
    sys.path.insert(0, str(ROOT / "lib"))
    import docx_safe_save  # noqa: F401
    from docx import Document
    dirpath.mkdir(parents=True, exist_ok=True)
    doc = Document()
    for c in captions:
        doc.add_paragraph(c)
    out = dirpath / "fixture.docx"
    doc.save(str(out))
    return out


def test_module_is_pure_stdlib():
    """它会被「自己开 zipfile 写 docx」那条路 import，必须能用系统 python3 裸跑。"""
    import ast
    src = (ROOT / "lib" / "caption_re.py").read_text(encoding="utf-8")
    # 走 ast 而不是正则 —— docstring 里的用法示例含 `from caption_re import ...`，
    # 正则会把它当成真 import（本测试第一版就是这么假红的）。
    mods = set()
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Import):
            mods.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0:
            mods.add((node.module or "").split(".")[0])
    allowed = {"re", "dataclasses", "functools", "typing", "__future__"}
    assert mods <= allowed, f"引入了非 stdlib 依赖: {mods - allowed}"
