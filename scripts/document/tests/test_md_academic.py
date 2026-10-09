"""Behavioral coverage for template-based academic MD conversion."""
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import unittest
import tempfile

ROOT = Path(__file__).resolve().parents[3]
DOC = ROOT / "scripts/document"
sys.path.insert(0, str(DOC))
spec = importlib.util.spec_from_file_location("md_tools_test", DOC / "md_tools.py")
md = importlib.util.module_from_spec(spec)
spec.loader.exec_module(md)


def template(tmp_path):
    path = tmp_path / "template.docx"
    Document().save(path)
    return path


def convert(tmp_path, text, academic=True):
    source = tmp_path / "source.md"
    source.write_text(text)
    styles, config = md.extract_styles_xml(template(tmp_path), tmp_path)
    target = tmp_path / "out.docx"
    md.convert_md_to_docx(source, styles, target, config, academic=academic)
    with ZipFile(target) as z:
        assert z.testzip() is None
        root = etree.fromstring(z.read("word/document.xml"))
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    return source, target, root, rels


def visible(p):
    return "".join(p.xpath(".//w:t/text()", namespaces=md.NSMAP))


def case_four_levels_references_and_exact_prose(tmp_path):
    text = """# 第一章 研究背景与文献综述

## 1.1 研究背景

### 1.1.1 复杂特性

#### 1.1.1.1 数学条件

作者（2024）指出，流量不能等同水位。[1](https://doi.org/10.1000/example)

# 参考文献

[1] Author. A **real title**[J]. Journal, 2024: 1-3. [DOI: 10.1000/example](https://doi.org/10.1000/example).

[2] 机构. 文件[EB/OL]. (2023-05-25)[2026-10-09]. [官方原文](https://example.org/policy).
"""
    source, target, root, rels = convert(tmp_path, text)
    assert source.read_text() == text
    ps = root.xpath("w:body/w:p", namespaces=md.NSMAP)
    headings = [(visible(p), p.xpath("string(w:pPr/w:outlineLvl/@w:val)", namespaces=md.NSMAP)) for p in ps if p.xpath("string(w:pPr/w:pStyle/@w:val)", namespaces=md.NSMAP).startswith("Heading")]
    assert [x[1] for x in headings] == ["0", "1", "2", "3", "0"]
    assert len(headings) == 5
    body = next(p for p in ps if visible(p).startswith("作者"))
    assert visible(body) == "作者（2024）指出，流量不能等同水位。[1]"
    citation = body.xpath('.//w:r[w:t="[1]"]', namespaces=md.NSMAP)[0]
    assert citation.xpath("string(w:rPr/w:vertAlign/@w:val)", namespaces=md.NSMAP) == "superscript"
    refs = [p for p in ps if visible(p).startswith("[")]
    assert len(refs) == 2
    assert "Author. A real title[J]. Journal, 2024: 1-3." in visible(refs[0])
    assert "官方原文: https://example.org/policy" in visible(refs[1])
    assert refs[0].xpath("string(w:pPr/w:ind/@w:hangingChars)", namespaces=md.NSMAP) == "200"
    assert not refs[0].xpath(".//w:vertAlign", namespaces=md.NSMAP)
    urls = {r.get("Target") for r in rels if r.get("Type", "").endswith("/hyperlink")}
    assert urls == {"https://doi.org/10.1000/example", "https://example.org/policy"}
    assert all(r.get(qn("w:val")) == "000000" for r in root.xpath(".//w:rPr/w:color", namespaces=md.NSMAP))
    assert root.xpath("w:body/w:p[w:r/w:t='参考文献']/w:pPr/w:pageBreakBefore", namespaces=md.NSMAP)
    assert not root.xpath(".//w:numPr", namespaces=md.NSMAP)
    with ZipFile(target) as z:
        assert b' PAGE ' in z.read("word/footer1.xml")
        styles = etree.fromstring(z.read("word/styles.xml"))
        assert styles.xpath('string(w:style[@w:styleId="Normal"]/w:rPr/w:sz/@w:val)', namespaces=md.NSMAP) == "24"
        assert styles.xpath('string(w:style[@w:styleId="参考文献"]/w:rPr/w:sz/@w:val)', namespaces=md.NSMAP) == "21"


def case_default_keeps_markdown_emphasis_and_resolves_links(tmp_path):
    _, target, root, _ = convert(tmp_path, "# Header\n\n**Bold** and *italic* [link](https://example.org).\n", academic=False)
    text = "".join(root.xpath(".//w:t/text()", namespaces=md.NSMAP))
    assert text == "HeaderBold and italic link."
    assert not root.xpath(".//w:vertAlign", namespaces=md.NSMAP)
    assert root.xpath(".//w:b", namespaces=md.NSMAP)
    assert root.xpath(".//w:i", namespaces=md.NSMAP)
    with ZipFile(target) as z:
        assert "word/footer1.xml" not in z.namelist()


def case_link_with_balanced_parentheses_and_safe_text(tmp_path):
    _, _, root, rels = convert(tmp_path, "# 一级\n\n见[文献](https://doi.org/10.1000/(ASCE).example)及[深层文献](https://example.org/a((b)c))。\n")
    assert "见文献及深层文献。" in "".join(root.xpath(".//w:t/text()", namespaces=md.NSMAP))
    assert any(r.get("Target") == "https://doi.org/10.1000/(ASCE).example" for r in rels)
    assert any(r.get("Target") == "https://example.org/a((b)c)" for r in rels)


def case_emphasis_around_and_inside_links(tmp_path):
    _, _, root, _ = convert(tmp_path, "# 一级\n\n**见[文献](https://example.org)**与[**加粗标签**](https://example.org/title)及*[斜体](https://example.org/italic)*。\n", academic=False)
    assert "见文献与加粗标签及斜体。" in "".join(root.xpath(".//w:t/text()", namespaces=md.NSMAP))
    links = root.xpath(".//w:hyperlink", namespaces=md.NSMAP)
    assert len(links) == 3
    assert all(link.xpath(".//w:b", namespaces=md.NSMAP) for link in links[:2])
    assert links[2].xpath(".//w:i", namespaces=md.NSMAP)


def case_academic_refuses_finder_fallback():
    p = subprocess.run([sys.executable, str(DOC / "md_tools.py"), "md2docx", "--academic"], capture_output=True, text=True)
    assert p.returncode != 0
    assert "显式提供" in p.stderr


def case_dispatch_rejects_existing_docx_before_touch(tmp_path):
    path = template(tmp_path)
    before = path.read_bytes()
    p = subprocess.run([sys.executable, str(DOC / "doc_dispatch.py"), "typeset", "--academic", str(path)], capture_output=True, text=True)
    assert p.returncode == 2
    assert path.read_bytes() == before


class AcademicConversionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_four_levels_and_references(self):
        case_four_levels_references_and_exact_prose(self.path)

    def test_default_and_emphasis(self):
        case_default_keeps_markdown_emphasis_and_resolves_links(self.path)

    def test_balanced_link(self):
        case_link_with_balanced_parentheses_and_safe_text(self.path)

    def test_emphasized_links(self):
        case_emphasis_around_and_inside_links(self.path)

    def test_no_finder(self):
        case_academic_refuses_finder_fallback()

    def test_existing_docx_unchanged(self):
        case_dispatch_rejects_existing_docx_before_touch(self.path)


if __name__ == "__main__":
    unittest.main()
