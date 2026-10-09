"""Academic formatting for newly generated MD documents; no renderer or writer."""
import importlib.util
from pathlib import Path

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree


def _profile():
    # Explicit owner module avoids the unrelated sub/styles.py name collision.
    path = Path(__file__).with_name("styles.py")
    spec = importlib.util.spec_from_file_location("_academic_styles_owner", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.load_profile("academic").FORMAT


def _child(parent, tag):
    node = parent.find(qn(tag))
    if node is None:
        node = OxmlElement(tag)
        parent.append(node)
    return node


def _remove(parent, tags):
    for tag in tags:
        for node in list(parent.findall(qn(tag))):
            parent.remove(node)


def _font(rpr, cfg, face=None, size=None):
    fonts = _child(rpr, "w:rFonts")
    fonts.attrib.clear()
    for slot in ("ascii", "hAnsi", "cs"):
        fonts.set(qn("w:" + slot), cfg["latin_font"])
    fonts.set(qn("w:eastAsia"), face or cfg["cjk_font"])
    color = _child(rpr, "w:color")
    color.attrib.clear()
    color.set(qn("w:val"), cfg["color"])
    _remove(rpr, ("w:u", "w:shd", "w:highlight"))
    if size is not None:
        for tag in ("w:sz", "w:szCs"):
            _child(rpr, tag).set(qn("w:val"), str(round(size * 2)))


def _text(node):
    return "".join(t.text or "" for t in node.iter(qn("w:t")))


def apply_academic(doc, elements, body_style, references):
    """Format an already generated document, preserving text and MD heading levels.

    references is the existing superscript module, not a second citation parser.
    Caller owns save/part-integrity contract. Tables/captions retain template styles.
    """
    cfg = _profile()
    for style in doc.styles:
        face = cfg["heading_font"] if "heading" in style.name.lower() or "标题" in style.name else cfg["cjk_font"]
        _font(_child(style.element, "w:rPr"), cfg, face)
        ppr = style.element.find(qn("w:pPr"))
        if ppr is not None:
            _remove(ppr, ("w:numPr", "w:pBdr"))
    defaults = _child(doc.styles.element, "w:docDefaults")
    _font(_child(_child(defaults, "w:rPrDefault"), "w:rPr"), cfg)
    for name in ["参考文献", *[f"Heading {n}" for n in range(1, 5)]]:
        if name not in doc.styles:
            doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    doc.styles["参考文献"].base_style = doc.styles[body_style]
    roles = [(body_style, cfg["body_pt"], None), ("参考文献", cfg["reference_pt"], None)]
    roles += [(f"Heading {n}", cfg["heading_pt"][n - 1], n) for n in range(1, 5)]
    for name, size, level in roles:
        style = doc.styles[name]
        _font(_child(style.element, "w:rPr"), cfg,
              cfg["heading_font"] if level else cfg["cjk_font"], size)
        style.font.bold = bool(level)
        fmt = style.paragraph_format
        fmt.line_spacing = cfg["line_spacing"]
        fmt.keep_with_next = bool(level)
        fmt.keep_together = bool(level)
        fmt.widow_control = True
        fmt.space_before = Pt(12 if level and level != 1 else 0)
        fmt.space_after = Pt(6 if level else 0)
        reference = name == "参考文献"
        hanging = cfg["reference_pt"] * cfg["reference_hanging_chars"]
        fmt.left_indent = Pt(hanging if reference else 0)
        fmt.first_line_indent = Pt(-hanging if reference else 0 if level else cfg["body_pt"] * cfg["first_line_chars"])
        ind = _child(_child(style.element, "w:pPr"), "w:ind")
        for key in ("firstLineChars", "hangingChars"):
            ind.attrib.pop(qn("w:" + key), None)
        if reference:
            ind.set(qn("w:hangingChars"), str(round(cfg["reference_hanging_chars"] * 100)))
        elif not level:
            ind.set(qn("w:firstLineChars"), str(round(cfg["first_line_chars"] * 100)))
    nodes = [n for n in doc.element.body if n.tag in (qn("w:p"), qn("w:tbl"))]
    if len(nodes) != len(elements):
        raise ValueError("MD elements and generated DOCX structure differ")
    from docx.text.paragraph import Paragraph
    for node, element in zip(nodes, elements):
        if node.tag != qn("w:p") or element["type"] in ("figure_title", "table_title"):
            continue
        paragraph = Paragraph(node, doc._body)
        heading = element["type"] == "heading"
        reference = references.is_ref_list_line(_text(node))
        level = element.get("level") if heading else None
        if heading and level not in (1, 2, 3, 4):
            raise ValueError("Academic mode supports MD H1-H4")
        paragraph.style = f"Heading {level}" if heading else "参考文献" if reference else body_style
        size = cfg["heading_pt"][level - 1] if heading else cfg["reference_pt"] if reference else cfg["body_pt"]
        face = cfg["heading_font"] if heading else cfg["cjk_font"]
        fmt = paragraph.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT if heading else WD_ALIGN_PARAGRAPH.JUSTIFY
        fmt.line_spacing = cfg["line_spacing"]
        fmt.widow_control = True
        fmt.keep_with_next = heading
        fmt.keep_together = heading
        fmt.page_break_before = heading and element["text"] == "参考文献"
        fmt.space_before = Pt(12 if heading and level != 1 else 0)
        fmt.space_after = Pt(6 if heading else 0)
        hanging = cfg["reference_pt"] * cfg["reference_hanging_chars"]
        fmt.left_indent = Pt(hanging if reference else 0)
        fmt.right_indent = Pt(0)
        fmt.first_line_indent = Pt(-hanging if reference else 0 if heading else cfg["body_pt"] * cfg["first_line_chars"])
        ppr = paragraph._p.get_or_add_pPr()
        _remove(ppr, ("w:numPr", "w:pBdr", "w:shd"))
        ind = _child(ppr, "w:ind")
        for key in ("firstLineChars", "hangingChars", "leftChars", "rightChars"):
            ind.attrib.pop(qn("w:" + key), None)
        if reference:
            ind.set(qn("w:hangingChars"), str(round(cfg["reference_hanging_chars"] * 100)))
        elif not heading:
            ind.set(qn("w:firstLineChars"), str(round(cfg["first_line_chars"] * 100)))
        _child(ppr, "w:outlineLvl").set(qn("w:val"), str(level - 1) if heading else "9")
        for run in node.iter(qn("w:r")):
            if _text(run):
                rpr = _child(run, "w:rPr")
                _font(rpr, cfg, face, size)
                if heading:
                    _child(rpr, "w:b").set(qn("w:val"), "1")
    for section in doc.sections:
        section.page_width, section.page_height = map(Cm, cfg["page_cm"])
        for name, value in cfg["margins_cm"].items():
            setattr(section, name + "_margin", Cm(value))
        footer = section.footer
        p = footer.paragraphs[0]
        p.clear()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(0)
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), " PAGE ")
        field.set(qn("w:dirty"), "true")
        run = OxmlElement("w:r")
        _font(_child(run, "w:rPr"), cfg, size=cfg["footer_pt"])
        _child(run, "w:t").text = "1"
        field.append(run)
        p._p.append(field)
    _child(doc.settings.element, "w:updateFields").set(qn("w:val"), "true")
    # Theme/effect parts can otherwise reintroduce DengXian in Word/WPS.
    for part in doc.part.package.parts:
        name = str(part.partname)
        if name == "/word/stylesWithEffects.xml":
            root = etree.fromstring(part.blob)
            for fonts in root.iter(qn("w:rFonts")):
                face = fonts.get(qn("w:eastAsia"))
                if face not in (cfg["cjk_font"], cfg["heading_font"]):
                    face = cfg["cjk_font"]
                _font(fonts.getparent(), cfg, face)
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif name == "/word/theme/theme1.xml":
            root = etree.fromstring(part.blob)
            for font in root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}fontScheme/*/*"):
                tag = etree.QName(font).localname
                if tag in ("latin", "cs", "ea") or font.get("script") == "Hans":
                    font.set("typeface", cfg["cjk_font"] if tag == "ea" or font.get("script") == "Hans" else cfg["latin_font"])
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    references.apply(doc)
