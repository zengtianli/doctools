"""tests for 规范化的**安全红线** —— 2026-07-26 审计实测出的 4 条,每条都真出过错。

1. 制表符原位:run 内 `[t, tab, t]` 改写后 tab 不许被挤到末尾(目录条/签署行会错位)
2. 页眉页脚默认保留:clean 不再删 headerReference/footerReference,要删得显式 --strip-headers
   (旧默认删掉了一切,连 typeset 套完院模板的 14/13 个引用也一并没了)
3. 单位词边界:小时候 ✗→ h候 · 毫米波 ✗→ mm波 · Item2 ✗→ Item²
4. 引号计数是真改动数,不是匹配数(本来就是中文引号的不该报"替换了 N 个")
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

_DOC = Path(__file__).resolve().parent.parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

spec = importlib.util.spec_from_file_location("dtf_safety_uut", str(_DOC / "docx_text_formatter.py"))
assert spec and spec.loader
dtf = importlib.util.module_from_spec(spec)
sys.modules["dtf_safety_uut"] = dtf
spec.loader.exec_module(dtf)

sys.path.insert(0, str(_DOC.parent.parent / "lib"))
from text_fixes import fix_quotes, fix_units  # noqa: E402


def _stats():
    return {"quotes": 0, "punctuation": 0, "units": 0, "scopes": {}}


def _kids(r):
    return [c.tag.split("}")[1] for c in r]


# ── 1. 制表符原位 ────────────────────────────────────────────────

def test_tab_stays_between_text():
    """run 子节点顺序 [t, tab, t] 必须原样保留 —— 合并文本会把 tab 挤到末尾。"""
    d = Document()
    p = d.add_paragraph()
    p._p.append(parse_xml(
        f'<w:r {nsdecls("w")}><w:t>第一项,</w:t><w:tab/><w:t>第二项(A)</w:t></w:r>'))
    r = p._p.findall(qn("w:r"))[0]
    assert _kids(r) == ["t", "tab", "t"]

    dtf.process_paragraph_element(p._p, _stats())

    r = p._p.findall(qn("w:r"))[0]
    assert _kids(r) == ["t", "tab", "t"], "tab 被挪位 = 目录条/签署行对齐错乱"
    assert [t.text for t in r.findall(qn("w:t"))] == ["第一项，", "第二项（A）"]


def test_br_also_preserved():
    d = Document()
    p = d.add_paragraph()
    p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>上行,</w:t><w:br/><w:t>下行;</w:t></w:r>'))
    dtf.process_paragraph_element(p._p, _stats())
    assert _kids(p._p.findall(qn("w:r"))[0]) == ["t", "br", "t"]


# ── 2. 页眉页脚默认保留 ──────────────────────────────────────────

@pytest.fixture
def docx_with_header(tmp_path: Path) -> Path:
    d = Document()
    sec = d.sections[0]
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "页眉:标题(H)"
    sec.footer.is_linked_to_previous = False
    sec.footer.paragraphs[0].text = "页脚:第1页"
    d.add_paragraph("正文,内容(A)")
    src = tmp_path / "withheader.docx"
    d.save(src)
    return src


def _header_refs(docx: Path) -> tuple[int, int]:
    d = Document(docx)
    h = sum(len(s._sectPr.findall(qn("w:headerReference"))) for s in d.sections)
    f = sum(len(s._sectPr.findall(qn("w:footerReference"))) for s in d.sections)
    return h, f


def _run(src: Path) -> Path:
    assert dtf.process_docx(str(src)) is True
    return src.with_name(f"{src.stem}_fixed{src.suffix}")


def test_headers_kept_by_default(docx_with_header: Path):
    """默认跑一遍规范化,页眉页脚引用必须一个不少。"""
    before = _header_refs(docx_with_header)
    assert before == (1, 1)
    assert dtf.STRIP_HEADERS is False
    assert _header_refs(_run(docx_with_header)) == before


def test_header_text_normalized(docx_with_header: Path):
    """保留的同时,页眉页脚里的半角标点也要被规范化。"""
    out = _run(docx_with_header)
    d = Document(out)
    assert "（H）" in d.sections[0].header.paragraphs[0].text


def test_strip_headers_opt_in(docx_with_header: Path):
    """显式 --strip-headers 时才删。"""
    dtf.STRIP_HEADERS = True
    try:
        assert _header_refs(_run(docx_with_header)) == (0, 0)
    finally:
        dtf.STRIP_HEADERS = False


# ── 3. 单位词边界 ────────────────────────────────────────────────

@pytest.mark.parametrize("text", [
    "小时候在毫米波实验室",
    "秒钟表和分钟表",
    "Item2 Form3",
    "毫升装置",
])
def test_units_not_touched_without_number(text: str):
    """没有数字打头的中文词 / 英文单词内的 m2 一律不动。"""
    assert fix_units(text) == (text, 0)


@pytest.mark.parametrize("text,expect", [
    ("面积5平方米", "面积5m²"),
    ("体积3 立方米", "体积3 m³"),
    ("长2公里", "长2km"),
    ("5m2", "5m²"),
    ("温度25摄氏度", "温度25℃"),
])
def test_units_still_convert_after_number(text: str, expect: str):
    assert fix_units(text)[0] == expect


# ── 4. 引号计数 = 真改动数 ───────────────────────────────────────

def test_quote_count_is_real_changes():
    assert fix_quotes("已是中文引号“甲”")[1] == 0
    assert fix_quotes('英文"甲"')[1] == 2
