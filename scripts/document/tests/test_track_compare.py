"""`track compare` 回归门 —— 段落级 docx 对比 → w:ins/w:del 修订件。

2026-08-03 起因：这条动词从落地那天起就是个假装成功的空桩
（`print("compare 功能将在 v2 实现。")`），rc 还是 0 —— 敲下去像成功了、
一件事没干。本文件把它换成真实现之后的**该抓的东西**全部钉死。

红线（改坏了必须红）：
  · 删除态的文本载体必须是 `w:delText`。写成 `w:t` = 把删掉的字变回正文，
    Word 里"拒绝修订"之后原文就没了，而肉眼看输出一切正常。
  · 空集不许报绿：两份都没有顶层段落时 difflib 安静地返回「无差异」——
    那是最会骗人的一种 rc=0，必须抛错。
  · 「段落级无差异」不许产出一个没有修订标记的副本冒充成功（out is None）。
  · 范围外差异（表格 / 页眉页脚 / 改稿新增图片）不许被静默吞掉，必须进
    out_of_scope 清单（CLI 侧 → rc=3）。
  · 两个输入件一个字节都不许变（compare 是只读输入 + 写第三份）。
  · 非 document.xml 的部件必须逐字节 verbatim（assert_parts_intact 在引擎里断言，
    这里再从产物侧独立验一遍 media 的字节）。
"""
from __future__ import annotations

import hashlib
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

sys.path.append(str(Path(__file__).resolve().parents[1] / "sub"))
from docx_track import CompareError, compare_docx  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _p(text: str, style: str | None = None) -> str:
    st = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return (f"<w:p>{st}<w:r><w:rPr><w:sz w:val=\"24\"/></w:rPr>"
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def _img_p(rid: str = "rId9") -> str:
    """带图片引用的段落（r:embed 指向 rels 里的一条关系）。"""
    return (f'<w:p><w:r><w:drawing><a:blip xmlns:a="x" r:embed="{rid}"/>'
            f"</w:drawing></w:r></w:p>")


def _tbl(cell: str) -> str:
    return (f"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>{cell}</w:t></w:r></w:p>"
            f"</w:tc></w:tr></w:tbl>")


def _docx(path: Path, body_xml: str, media: bytes = b"PNGDATA") -> Path:
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document xmlns:w="{WNS}" xmlns:r="{RNS}"><w:body>'
           + body_xml + "</w:body></w:document>")
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
          'package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxml'
          'formats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{RNS.replace("officeDocument/2006", "package/2006")}">'
            f'<Relationship Id="rId1" Type="{RNS}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId9" Type="{RNS}/image" Target="media/a.png"/>'
            "</Relationships>")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns='
                   '"http://schemas.openxmlformats.org/package/2006/relationships"/>')
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/styles.xml",
                   f'<w:styles xmlns:w="{WNS}"><w:style w:styleId="2"/></w:styles>')
        z.writestr("word/media/a.png", media)
    return path


BASE = _p("第1章 概述", style="2") + _p("原文甲") + _p("原文乙") + _p("原文丙")


@pytest.fixture()
def src(tmp_path):
    return _docx(tmp_path / "src.docx", BASE)


def _root(path: Path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


# ── 主路径 ───────────────────────────────────────────────────────────────
def test_modify_add_delete_produces_ins_and_del(src, tmp_path):
    """改一段 + 增一段 + 删一段 → w:ins / w:del 都真的出现在产物里。"""
    dst = _docx(tmp_path / "dst.docx",
                _p("第1章 概述", style="2") + _p("改后甲") + _p("原文乙")
                + _p("末尾新增段"))          # 丙被删、甲被改、末尾新增
    out = tmp_path / "out.docx"
    st = compare_docx(str(src), str(dst), str(out))

    assert out.exists(), "有差异却没产出文件"
    r = _root(out)
    ins, dele = list(r.iter(W + "ins")), list(r.iter(W + "del"))
    assert ins and dele, f"产物里没有修订标记：ins={len(ins)} del={len(dele)}"

    del_text = "".join(t.text or "" for t in r.iter(W + "delText"))
    ins_text = "".join(t.text or "" for t in r.iter(W + "t")
                       if any(a.tag == W + "ins" for a in t.iterancestors()))
    assert "原文甲" in del_text and "原文丙" in del_text, del_text
    assert "改后甲" in ins_text and "末尾新增段" in ins_text, ins_text
    assert "原文乙" not in del_text and "原文乙" not in ins_text, "没改的段被动了"
    assert st["deleted_paras"] == 2 and st["inserted_paras"] == 2, st
    assert st["out_of_scope"] == [], st["out_of_scope"]


def test_deleted_text_uses_delText_not_t(src, tmp_path):
    """红线：w:del 内一律 w:delText。写成 w:t = 删掉的字变回正文（拒绝修订后原文丢失）。"""
    dst = _docx(tmp_path / "dst.docx", _p("第1章 概述", style="2") + _p("原文乙"))
    out = tmp_path / "out.docx"
    compare_docx(str(src), str(dst), str(out))
    r = _root(out)
    stray = [t for t in r.iter(W + "t")
             if any(a.tag == W + "del" for a in t.iterancestors())]
    assert not stray, f"{len(stray)} 个 w:del 内的文本用了 w:t"
    assert list(r.iter(W + "delText")), "一个 w:delText 都没有"


def test_paragraph_marks_are_marked(src, tmp_path):
    """整段增删必须连段落标记一起标，否则 Word 里接受修订后会剩下空行。"""
    dst = _docx(tmp_path / "dst.docx",
                _p("第1章 概述", style="2") + _p("原文乙") + _p("新增"))
    out = tmp_path / "out.docx"
    compare_docx(str(src), str(dst), str(out))
    r = _root(out)
    marks = [(p.find(W + "pPr") is not None
              and p.find(W + "pPr").find(W + "rPr") is not None
              and p.find(W + "pPr").find(W + "rPr").find(W + tag) is not None)
             for tag in ("ins", "del") for p in r.iter(W + "p")]
    assert any(marks), "没有任何段落标记被标成修订态"
    ins_marks = sum(1 for p in r.iter(W + "p")
                    if (ppr := p.find(W + "pPr")) is not None
                    and (rpr := ppr.find(W + "rPr")) is not None
                    and rpr.find(W + "ins") is not None)
    del_marks = sum(1 for p in r.iter(W + "p")
                    if (ppr := p.find(W + "pPr")) is not None
                    and (rpr := ppr.find(W + "rPr")) is not None
                    and rpr.find(W + "del") is not None)
    assert ins_marks == 1 and del_marks == 2, (ins_marks, del_marks)


def test_inputs_are_never_touched(src, tmp_path):
    """compare 只读两个输入、写第三份。输入件哪怕一个字节都不许变。"""
    dst = _docx(tmp_path / "dst.docx", BASE + _p("新增"))
    before = (_md5(src), _md5(dst))
    compare_docx(str(src), str(dst), str(tmp_path / "out.docx"))
    assert (_md5(src), _md5(dst)) == before, "输入件被改写了"


def test_non_document_parts_are_verbatim(src, tmp_path):
    """除 word/document.xml 外逐字节 verbatim（丢 chart / 剥 OLE 那类事故的门）。"""
    dst = _docx(tmp_path / "dst.docx", BASE + _p("新增"))
    out = tmp_path / "out.docx"
    compare_docx(str(src), str(dst), str(out))
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        assert set(a.namelist()) == set(b.namelist())
        for name in a.namelist():
            if name == "word/document.xml":
                continue
            assert a.read(name) == b.read(name), f"部件被重写：{name}"


# ── fail-closed 出口 ─────────────────────────────────────────────────────
def test_identical_docs_produce_no_file(src, tmp_path):
    """无差异 → 不产出文件（禁「产一个没有修订的副本」冒充成功）。"""
    same = _docx(tmp_path / "same.docx", BASE)
    out = tmp_path / "out.docx"
    st = compare_docx(str(src), str(same), str(out))
    assert st["out"] is None and not out.exists()
    assert st["deleted_paras"] == 0 and st["inserted_paras"] == 0


def test_empty_body_raises(src, tmp_path):
    """空集不许报绿：0 段落的一边必须抛错，不许安静地「无差异」。"""
    empty = _docx(tmp_path / "empty.docx", "")
    with pytest.raises(CompareError):
        compare_docx(str(src), str(empty), str(tmp_path / "out.docx"))
    with pytest.raises(CompareError):
        compare_docx(str(empty), str(src), str(tmp_path / "out.docx"))


def test_output_over_input_refused(src, tmp_path):
    dst = _docx(tmp_path / "dst.docx", BASE + _p("新增"))
    with pytest.raises(CompareError):
        compare_docx(str(src), str(dst), str(src))
    with pytest.raises(CompareError):
        compare_docx(str(src), str(dst), str(dst))


def test_missing_input_raises(src, tmp_path):
    with pytest.raises(CompareError):
        compare_docx(str(src), str(tmp_path / "nope.docx"), str(tmp_path / "o.docx"))


# ── 范围外差异：不许静默吞 ───────────────────────────────────────────────
def test_table_only_diff_is_reported_not_swallowed(tmp_path):
    a = _docx(tmp_path / "a.docx", _p("正文") + _tbl("100"))
    b = _docx(tmp_path / "b.docx", _p("正文") + _tbl("999"))
    st = compare_docx(str(a), str(b), str(tmp_path / "out.docx"))
    assert st["out"] is None, "表格差异不该产出一个没有修订标记的文件"
    assert any("表格" in n for n in st["out_of_scope"]), st["out_of_scope"]


def test_foreign_rel_in_inserted_para_is_stripped_and_reported(src, tmp_path):
    """改稿新增段里的图引用的是**改稿的** rels；照搬进原稿包 = Word 开门就报损坏。
    摘掉可以，静默摘掉不行。"""
    dst = _docx(tmp_path / "dst.docx", BASE + _img_p("rIdNOTEXIST"))
    out = tmp_path / "out.docx"
    st = compare_docx(str(src), str(dst), str(out))
    assert out.exists()
    assert any("摘除" in n for n in st["out_of_scope"]), st["out_of_scope"]
    r = _root(out)
    assert not list(r.iter("{x}blip")), "悬空关系引用被原样搬进了产物"


def test_unknown_style_in_inserted_para_is_reported(src, tmp_path):
    dst = _docx(tmp_path / "dst.docx", BASE + _p("新增", style="NoSuchStyle"))
    st = compare_docx(str(src), str(dst), str(tmp_path / "out.docx"))
    assert any("样式" in n for n in st["out_of_scope"]), st["out_of_scope"]


# ── 已有修订态的稿子 ─────────────────────────────────────────────────────
def test_existing_deletions_do_not_participate(tmp_path):
    """原稿里别人未接受的 w:del 文本不参与比对 —— 否则会被当成「改稿删掉了它」再删一遍。"""
    a = _docx(tmp_path / "a.docx",
              "<w:p><w:r><w:t>留下的</w:t></w:r>"
              "<w:del w:id=\"1\" w:author=\"甲\" w:date=\"2026-01-01T00:00:00Z\">"
              "<w:r><w:delText>已删的</w:delText></w:r></w:del></w:p>")
    b = _docx(tmp_path / "b.docx", "<w:p><w:r><w:t>留下的</w:t></w:r></w:p>")
    st = compare_docx(str(a), str(b), str(tmp_path / "out.docx"))
    assert st["out"] is None, "把别人已删的字又比出一次差异"
