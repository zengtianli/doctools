"""docx surgical 惯用法公共库 — 从 restyle.py / center_images.py / docx_renumber_figures.py 抽取。

收口「一次解压、只重写 word/document.xml、其余 zip 项 verbatim」这套复制粘贴 ≥10 处的
surgical 惯用法。只供新工具（docx_para.py）用；存量脚本不强制回迁（防 balloon）。

import 路径:本文件在 doctools/lib 下,与 docx_xml 同层 —— 插同层目录后直接 import。
"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))  # 同层 docx_xml
from docx_xml import NSMAP, W, qn  # noqa: E402  复用既有 namespace 常量,不重造

__all__ = [
    "NSMAP",
    "W",
    "qn",
    "RepackError",
    "read_document_xml",
    "parse_document",
    "iter_paras",
    "para_text",
    "body_start_idx",
    "make_backup",
    "surgical_rewrite",
    "surgical_rewrite_parts",
    "read_part",
    "parse_part",
    "list_parts",
    "serialize",
    "PPR_ORDER",
    "ensure_ppr_child",
    "word_lock_file",
    "verify_repacked",
]

DOCUMENT_XML = "word/document.xml"


class RepackError(Exception):
    """重打包后自检失败(zip CRC 坏 / document.xml 无法 parse)。已尽量回滚。"""


# ── 读 ────────────────────────────────────────────────────────────────────
def read_document_xml(docx: Path) -> bytes:
    """zip 内只读 word/document.xml 字节。"""
    with zipfile.ZipFile(docx) as z:
        return z.read(DOCUMENT_XML)


def parse_document(docx: Path):
    """read + etree.fromstring → root。"""
    return etree.fromstring(read_document_xml(docx))


def read_part(docx: Path, name: str) -> bytes:
    """zip 内只读任意一个部件的字节(word/styles.xml、word/header1.xml …)。"""
    with zipfile.ZipFile(docx) as z:
        return z.read(name)


def parse_part(docx: Path, name: str):
    """read_part + etree.fromstring → root。部件不存在直接抛,不返回 None ——
    调用方拿到 None 往往会当成「这个部件是空的」继续跑,那是假绿。"""
    return etree.fromstring(read_part(docx, name))


def list_parts(docx: Path, pattern: str | None = None) -> list[str]:
    """列出包内部件名;pattern 给正则则只留匹配的(如 r"word/(header|footer)\\d*\\.xml")。"""
    with zipfile.ZipFile(docx) as z:
        names = z.namelist()
    if pattern is None:
        return names
    rx = re.compile(pattern)
    return [n for n in names if rx.fullmatch(n)]


def iter_paras(root) -> list:
    """body 直接 + 嵌套(表格内)的所有 <w:p>,文档序。"""
    return list(root.iter(qn("w:p")))


_WS = re.compile(r"\s+")


def para_text(p, *, normalize: bool = False) -> str:
    """拼段内所有 w:t 文本。normalize=去全部空白(CJK 报告匹配用,防 run-split 假阴性)。"""
    s = "".join(t.text or "" for t in p.iter(qn("w:t")))
    return _WS.sub("", s) if normalize else s


def body_start_idx(paras) -> int:
    """正文起始段索引 = 目录(TOC)字段之后(封面/批准/落款/目录都在 TOC 之前)。

    抽自 docx_renumber_figures.py:199 / center_images.py:41(逐字一致)。
    判据 = 最后一个 TOC / PAGEREF _Toc 字段段之后;无 TOC 则返 0(保持旧行为)。
    """
    last = -1
    for idx, p in enumerate(paras):
        instr = "".join(n.text or "" for n in p.iter(qn("w:instrText")))
        if "TOC" in instr or "PAGEREF _Toc" in instr:
            last = idx
    return last + 1


# ── 写 ────────────────────────────────────────────────────────────────────
def make_backup(docx: Path) -> Path:
    """.bak-%Y%m%d-%H%M%S,shutil.copy2(保留 mtime)。返回 bak 路径。"""
    bak = docx.with_suffix(docx.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
    shutil.copy2(docx, bak)
    return bak


def serialize(root) -> bytes:
    """etree.tostring — 与存量惯用法(restyle/center_images)逐字一致:
    xml_declaration + UTF-8 + standalone=True + 不 pretty_print(防 Word 拒读)。"""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def surgical_rewrite(docx: Path, new_document_xml: bytes, *, backup: bool = True) -> Path | None:
    """verbatim repack:只替换 word/document.xml,其余 zip 项(媒体/embeddings/OLE/公式)原样搬运。

    = `surgical_rewrite_parts(docx, {DOCUMENT_XML: ...})` 的单部件快捷式,签名与语义
    保持不变(存量 docx_para.py 直接依赖它)。
    """
    return surgical_rewrite_parts(docx, {DOCUMENT_XML: new_document_xml}, backup=backup)


def surgical_rewrite_parts(docx: Path, parts: dict[str, bytes], *,
                           backup: bool = True) -> Path | None:
    """verbatim repack:替换 `parts` 里点名的若干 zip 部件,其余全部原样搬运。

    2026-07-28 加。原来只能改 `word/document.xml`,而真正要迁的存量脚本里 12/16 还要动
    `word/styles.xml`(字体/样式)、`word/numbering.xml`(编号)、`word/header*.xml`
    `word/footer*.xml`(页眉页脚) —— 缺这个能力,那些脚本就只能继续用 python-docx 整包重存,
    而整包重存正是剥掉 OLE/公式/嵌入对象的那个动作(磐安母本 47 个公式对象 → 产物 0 个)。

    **只改不增删**:parts 里出现的部件必须已存在于原包。想新增部件(如首次加页眉)需要连
    `[Content_Types].xml` 和 `.rels` 一起改,那是另一件事,本函数拒绝——静默忽略一个
    写不进去的部件,和它要防的数据丢失是同一类东西。

    写 .tmp 后 tmp.replace(docx) 原子换 → verify_repacked 自检(逐个 parse 改过的部件)。
    坏 → 用 bak 回滚 + raise RepackError。返回 bak 路径(backup=False 则 None)。
    """
    if not parts:
        raise ValueError("surgical_rewrite_parts: parts 为空 —— 拒绝做一次无意义的重打包")
    with zipfile.ZipFile(docx) as z:
        existing = set(z.namelist())
    missing = sorted(set(parts) - existing)
    if missing:
        raise RepackError(
            f"这些部件不在原包里,本函数只改不增: {missing}。"
            f"新增部件需同时改 [Content_Types].xml 与 .rels")

    bak = make_backup(docx) if backup else None
    tmp = docx.with_suffix(docx.suffix + ".tmp")
    try:
        with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = parts.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)  # 复用 zin.infolist() 的 ZipInfo 保留每项 compress_type/日期
        tmp.replace(docx)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    # 写后自检 → 坏则回滚
    try:
        verify_repacked(docx, extra_parts=list(parts))
    except Exception as e:
        if bak is not None and bak.exists():
            shutil.copy2(bak, docx)
            raise RepackError(f"repack 自检失败,已从 {bak.name} 回滚: {e}") from e
        raise RepackError(f"repack 自检失败(无备份可回滚): {e}") from e
    return bak


def verify_repacked(docx: Path, *, extra_parts: list[str] | None = None) -> None:
    """重开 zip + testzip(全项 CRC) + parse document.xml 及所有被改过的部件。坏 → raise。

    改过哪个就 parse 哪个:只 parse document.xml 而放过刚写坏的 styles.xml,
    等于自检报了个假绿。
    """
    with zipfile.ZipFile(docx) as z:
        bad = z.testzip()
        if bad is not None:
            raise ValueError(f"zip CRC 坏: {bad}")
        for name in dict.fromkeys([DOCUMENT_XML, *(extra_parts or [])]):
            if not name.endswith(".xml") and not name.endswith(".rels"):
                continue          # 非 XML 部件(媒体)无从 parse
            try:
                etree.fromstring(z.read(name))
            except Exception as e:
                raise ValueError(f"{name} 写坏了,parse 失败: {e}") from e


# ── pPr schema 序 ─────────────────────────────────────────────────────────
# CT_PPr 子元素 schema 合法顺序 — 抽自 center_images.py:61-70。
# jc/ind 等必须按此插入,否则 Word 拒读。
PPR_ORDER: list[str] = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]
_ORDER_IDX = {qn("w:" + t): i for i, t in enumerate(PPR_ORDER)}


def ensure_ppr_child(pPr, tag: str):
    """取/建 pPr 下 <w:{tag}> 子元素,按 schema 顺序就位。返回该元素。

    抽自 center_images._ensure_in_order。tag = 局部名(如 'jc' / 'ind')。
    """
    full = qn("w:" + tag)
    el = pPr.find(full)
    if el is not None:
        return el
    el = etree.Element(full)
    my = _ORDER_IDX.get(full, 999)
    pos = len(pPr)
    for i, ch in enumerate(pPr):
        if _ORDER_IDX.get(ch.tag, 999) > my:
            pos = i
            break
    pPr.insert(pos, el)
    return el


# ── 防呆 ──────────────────────────────────────────────────────────────────
def word_lock_file(docx: Path) -> Path | None:
    """Word/WPS 打开文档时的 owner 锁文件。存在即返回(说明文档正被打开)。

    命名规则:短名 = '~$' + 全名;长名 = Word 用 '~$' 替换前两个字符('~$' + name[2:])。
    两式都查(实测天台报告走后者:天台县… → ~$县…)。
    """
    for cand in (docx.parent / ("~$" + docx.name), docx.parent / ("~$" + docx.name[2:])):
        if cand.exists():
            return cand
    return None
