"""
Word XML namespace 常量与工具函数

供 docx_track_changes / docx_extract / md_docx_template 共用

⚠ 「全文改写」类工具必读（2026-07-26）：python-docx 的 `Document.paragraphs` 只认
   `w:body/./w:p`、`Paragraph.runs` 只认 `./w:r`（`CT_P.r = ZeroOrMore("w:r")`），
   **静默漏掉审阅修订（`w:ins`/`w:del`）、`w:hyperlink`、`w:smartTag`、文本框
   （`w:txbxContent`）、嵌套表格里的 run**；批注（comments.xml）/脚注/尾注更是整个
   part 都碰不到。凡「把全文某类字符改一遍」的工具都必须走本模块 §run 遍历，
   否则带修订的稿子看着跑绿了，修订段其实一个字没改。
"""

# ── WordprocessingML 命名空间 ─────────────────────────────────

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

W = NSMAP["w"]
R_NS = NSMAP["r"]

# 关系类型常量
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"


def qn(tag: str) -> str:
    """将 'w:t' 转换为 '{namespace}t' 格式"""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


# ── run / 文本元素级遍历（绕开 python-docx 漏修订的坑）────────────

TEXT_TAGS = (qn("w:t"), qn("w:delText"))   # 正常文本 / 修订删除态文本
_NESTED = (qn("w:p"),)                     # 嵌套段落（文本框、嵌套表）自成一段


def iter_paragraphs(root):
    """root 下所有 w:p，文档顺序（含表格 / 文本框 / 修订块里的嵌套段落）。"""
    return root.iter(qn("w:p"))


def para_own_runs(p) -> list:
    """段落 p 自己的 w:r 列表（文档顺序）。

    含：直接子 run + `w:ins`/`w:del`/`w:hyperlink`/`w:smartTag`/`w:sdt` 等包装层里的 run。
    不含：嵌套 `w:p`（文本框 `w:txbxContent`、嵌套表格单元格）里的 run —— 那些由
    `iter_paragraphs` 各自迭代到，在此剪掉才不会被处理两遍（引号配对会算错）。
    """
    out: list = []

    def walk(el):
        for child in el:
            if child.tag in _NESTED:
                continue
            if child.tag == qn("w:r"):
                out.append(child)
            walk(child)

    walk(p)
    return out


_DEL_ANCESTORS = (qn("w:del"), qn("w:moveFrom"))   # 删除态 / 移动走的原位置


def in_deleted(el) -> bool:
    """el 位于删除态（`w:del` 审阅删除 / `w:moveFrom` 审阅移出）内。

    删除态的文本载体是 `w:delText` 而非 `w:t` —— 写成 `w:t` 会让删除的文字变成正文。
    """
    return any(a.tag in _DEL_ANCESTORS for a in el.iterancestors())


def in_inserted(el) -> bool:
    """el 位于 `w:ins`（审阅插入态）内。"""
    return any(a.tag == qn("w:ins") for a in el.iterancestors())


_REVISION_ANCESTORS = (qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"))


def in_revision(el) -> bool:
    """el 位于任一审阅修订标记内（插入 / 删除 / 移入 / 移出）。"""
    return any(a.tag in _REVISION_ANCESTORS for a in el.iterancestors())


def in_table(el) -> bool:
    """el 位于表格单元格内（含嵌套表）。"""
    return any(a.tag == qn("w:tbl") for a in el.iterancestors())


def scope_of(el) -> str:
    """正文 part 内一个 run 归属哪个可勾选的域。

    修订优先于表格：表格里的修订段勾「审阅修订」就该动，不该被「表格」勾选状态左右。
    """
    if in_revision(el):
        return "revision"
    if in_table(el):
        return "table"
    return "body"


def text_elems(r) -> list:
    """run 的文本子元素（`w:t` / `w:delText`），文档顺序。"""
    return [c for c in r if c.tag in TEXT_TAGS]


def run_text(r) -> str:
    """run 的文本（含删除态）。不含 `w:instrText`（域代码）/ `w:tab` / `w:br`。"""
    return "".join(c.text or "" for c in text_elems(r))


def text_groups(r) -> list[list]:
    """run 内**可安全合并处理**的文本段分组。

    被 `w:tab` / `w:br` / `w:drawing` 等非文本节点隔开的文本元素分属不同组：
    合并跨越它们会把中间那个节点挤到末尾 —— 实测 `[t, tab, t]` 改写后变成
    `[t, tab]`，"第一项⇥第二项" 成了 "第一项第二项⇥"，目录条与签署行的
    制表位对齐当场错位。组内相邻文本元素合并处理（跨元素的"平方米"仍能命中）。
    """
    groups: list[list] = []
    cur: list = []
    for c in r:
        if c.tag in TEXT_TAGS:
            cur.append(c)
        elif c.tag == qn("w:rPr"):
            continue          # rPr 是属性不是内容，不切断文本
        elif cur:
            groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)
    return groups


def group_text(g: list) -> str:
    return "".join(e.text or "" for e in g)


def set_group_text(g: list, s: str) -> None:
    """把整组文本写回组内首个元素，删掉组内其余元素（组外节点原位不动）。"""
    g[0].text = s
    _mark_space(g[0])
    for extra in g[1:]:
        extra.getparent().remove(extra)


def has_other_content(r) -> bool:
    """run 里除 rPr 与文本外还有别的内容（tab / br / drawing / 脚注引用…）。

    这类 run 不能做"按引号拆 run"——拆出来的副本会丢失或复制这些一次性载荷。
    """
    return any(c.tag not in TEXT_TAGS and c.tag != qn("w:rPr") for c in r)


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _mark_space(el) -> None:
    """首尾空白必须 `xml:space="preserve"`，否则 Word 吃掉空格。"""
    s = el.text or ""
    if s != s.strip():
        el.set(XML_SPACE, "preserve")


def set_run_text(r, s: str) -> None:
    """写回 run 文本。

    保持原文本元素类型（`w:delText` 不退化成 `w:t`，否则修订删除态会变成正文）
    与同级非文本子节点（`w:br`/`w:tab`/`w:drawing` 原位保留）—— 这两点是
    python-docx `Run.text` setter 做不到的（它把 run 内容清空重建成单个 `w:t`）。
    """
    elems = text_elems(r)
    if not elems:
        if not s:
            return
        r.append(new_text_elem(r, s))
        return
    elems[0].text = s
    _mark_space(elems[0])
    for extra in elems[1:]:        # 多 w:t 的 run：文本已合并进首个，其余删掉
        r.remove(extra)


def text_tag_for(r) -> str:
    """该 run 应该用哪种文本元素：优先看它现有的（`w:delText` 就继续 delText），
    否则按修订态推断。`w:moveFrom`/`w:del` 里写成 `w:t` = 把删掉的字变回正文。"""
    elems = text_elems(r)
    if elems:
        return "w:delText" if elems[0].tag == qn("w:delText") else "w:t"
    return "w:delText" if in_deleted(r) else "w:t"


def new_text_elem(r, s: str):
    """按 run 所处修订态建对应文本元素（删除态 → `w:delText`）。"""
    from docx.oxml import OxmlElement  # 惰性导入：zip 级工具不必装 python-docx

    el = OxmlElement(text_tag_for(r))
    el.text = s
    _mark_space(el)
    return el


PART_SCOPES = {"comments": "批注", "notes": "脚注/尾注", "headers": "页眉页脚"}
BODY_SCOPES = {"body": "正文", "table": "表格", "revision": "审阅修订"}
ALL_SCOPES = {**BODY_SCOPES, **PART_SCOPES}


def iter_text_roots(doc, *, include_headers: bool = True, parts: set[str] | None = None):
    """遍历一个 python-docx Document 里**所有承载正文文本的 XML part**。

    yield `(label, root, flush)`：
      - `label` 中文域名（正文 / 批注 / 脚注 / 尾注 / 页眉 / 页脚），用于分域统计
      - `root`  该 part 的 lxml 根（正文给 `w:body`）
      - `flush` 收尾回写回调；`None` 表示该 part 由 python-docx 自己序列化

    `parts` 给 `PART_SCOPES` 的 key 子集（comments / notes / headers）时只产出被选中的
    part；`None` = 全给。正文 part 永远产出（正文内部的 body/table/revision 三个域是
    **run 级**过滤，由调用方按 `scope_of()` 决定，不在这里剪）。

    脚注 / 尾注在 python-docx 里是**没有 `.element` 的通用 Part**（只有 `_blob` 字节），
    所以必须解析 blob 再回写；批注是 `CommentsPart`（有 `.element`）直接改。
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.oxml import serialize_part_xml
    from docx.oxml import parse_xml

    yield ("正文", doc.element.body, None)

    want = set(PART_SCOPES) if parts is None else set(parts)
    if not include_headers:
        want.discard("headers")

    wanted = {}
    if "comments" in want:
        wanted[RT.COMMENTS] = "批注"
    if "notes" in want:
        wanted[RT.FOOTNOTES] = "脚注"
        wanted[RT.ENDNOTES] = "尾注"
    if "headers" in want:
        wanted[RT.HEADER] = "页眉"
        wanted[RT.FOOTER] = "页脚"

    for rel in doc.part.rels.values():
        if rel.is_external:
            continue
        label = wanted.get(rel.reltype)
        if not label:
            continue
        part = rel.target_part
        elm = getattr(part, "element", None)
        if elm is not None:
            yield (label, elm, None)
            continue
        root = parse_xml(part.blob)

        def _flush(_part=part, _root=root):
            _part._blob = serialize_part_xml(_root)

        yield (label, root, _flush)
