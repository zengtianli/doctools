"""docx surgical 惯用法公共库 — 从 restyle.py / center_images.py / docx_renumber_figures.py
（现 renum.py figures，2026-07-31 家族折叠时已回迁改用本库）抽取。

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
from io import BytesIO
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))  # 同层 docx_xml
from docx_xml import NSMAP, W, qn  # noqa: E402  复用既有 namespace 常量,不重造
from docx_parts import DEFAULT_ALLOW_CHANGED, assert_parts_intact, diff_parts  # noqa: E402

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
    "canonical",
    "graft_unchanged",
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

    抽自 docx_renumber_figures.py:199 / center_images.py:41(逐字一致；前者 2026-07-31
    折进 renum.py figures 后已直接引本函数,景宁 0313 坑注释见 renum._body_start_idx)。
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
        # 部件完整性断言(replace 前:此刻 docx 仍是未动源件=天然基线;炸则 tmp 被清、源件无损)。
        # 白名单 = DEFAULT + 本次点名改写的 parts —— 强制把改动面显式报备。
        assert_parts_intact(docx, tmp,
                            allow_changed=set(DEFAULT_ALLOW_CHANGED) | set(parts),
                            verbose=False)
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


def canonical(data: bytes) -> bytes:
    """XML 规范化(C14N):抹掉属性序 / 命名空间声明位置 / 无意义空白,只留语义。
    非 XML(媒体/OLE 二进制)原样返回 —— 它们本来就该逐字节比。"""
    try:
        tree = etree.parse(BytesIO(data))
    except Exception:
        return data
    out = BytesIO()
    tree.write_c14n(out)
    return out.getvalue()


def graft_unchanged(original: Path, modified: Path, *, on_missing: str = "error") -> dict:
    """把 `modified` 里**语义没变**的部件,按 `original` 的原始字节还原。原地改写 modified。

    ## 为什么需要它 (2026-07-28 立,实测数据)

    一份 301 部件 / 75 个嵌入公式的真报告,python-docx **什么都不改**地打开再存回:

        字节层面变了 60 个部件(含 [Content_Types].xml / _rels/.rels / 28 个 header / 14 个 footer)
        语义层面变了  1 个部件([Content_Types].xml:两条 <Default> 被拆成逐文件 <Override>)

    也就是说 59/60 是纯重新序列化 —— 内容一模一样,只是被 python-docx 的序列化器重写了一遍。
    改一个字号,炸开面 60 倍;而每一次重写都是一次「Word 可能渲染得不一样」的机会。

    有了这个函数,存量那 16 个 python-docx 脚本**一行都不用改**:让它照常跑完,然后把没真变的
    部件还原回去,炸开面立刻从 60 收到 1。不需要把 8500 行重写成手搓 XML。

    ## on_missing —— 部件在 modified 里消失了怎么办

    `Document()` 新建再搬内容那种写法会真的丢部件(实测 301 → 17、75 个公式对象归零)。
    默认 "error":直接抛。**不默认悄悄补回来** —— 补回来的部件没有任何东西引用它,
    得到的是一份「看起来完整、其实公式已经不在正文里」的文件,比直接报错危险得多。
      "error"   丢了就抛 RepackError(默认)
      "restore" 用 original 的字节补回(只在你确认调用方本就该保留全部部件时用)
      "ignore"  接受丢失(极少;比如 pdf_to_docx 这种本来就在造新文件)

    返回 {"还原": [...], "真变了": [...], "新增": [...], "丢失": [...]}。
    """
    if on_missing not in ("error", "restore", "ignore"):
        raise ValueError(f"on_missing 只能是 error/restore/ignore,给的是 {on_missing!r}")

    with zipfile.ZipFile(original) as z:
        orig = {i.filename: (z.read(i.filename), i) for i in z.infolist()}
    with zipfile.ZipFile(modified) as z:
        mod = {i.filename: (z.read(i.filename), i) for i in z.infolist()}

    lost = sorted(set(orig) - set(mod))
    if lost and on_missing == "error":
        raise RepackError(
            f"这个工具把 {len(lost)} 个部件整个弄丢了,不是「改」是「重造」:{lost[:8]}"
            f"{' …' if len(lost) > 8 else ''}。确认无害请显式传 on_missing=")

    kept, real, added = [], [], sorted(set(mod) - set(orig))
    out_parts: list[tuple[zipfile.ZipInfo, bytes]] = []
    for name, (mdata, minfo) in mod.items():
        if name in orig:
            odata, oinfo = orig[name]
            if odata == mdata or canonical(odata) == canonical(mdata):
                kept.append(name)
                out_parts.append((oinfo, odata))       # 连 ZipInfo 一起用原件的
                continue
            real.append(name)
        out_parts.append((minfo, mdata))
    if lost and on_missing == "restore":
        for name in lost:
            odata, oinfo = orig[name]
            out_parts.append((oinfo, odata))

    # [Content_Types].xml:部件集合一个没增没减时,原来那张表必然仍然正确 —— 还原它。
    # python-docx 每次存盘都会重建这张表(实测把 <Default Extension="bin"/> 和
    # <Default Extension="vsdx"/> 拆成 122 条逐文件 <Override>),那是它按自己认得的
    # 部件重新推导的结果 —— 它认不得的东西就靠 <Default> 兜底,而 <Default> 恰好被它删了。
    # 部件集合一变(新增/丢失)就不能还原:新部件需要新的类型声明。
    CT = "[Content_Types].xml"
    if not added and not lost and CT in real and CT in orig:
        real.remove(CT)
        kept.append(CT)
        odata, oinfo = orig[CT]
        out_parts = [(oinfo, odata) if i.filename == CT else (i, d) for i, d in out_parts]

    tmp = modified.with_suffix(modified.suffix + ".graft")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info, data in out_parts:
                zout.writestr(info, data)
        tmp.replace(modified)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    verify_repacked(modified, extra_parts=real)
    result = {"还原": sorted(kept), "真变了": sorted(real), "新增": added, "丢失": lost}
    if on_missing in ("restore", "ignore"):
        # 非 fail-closed 分支不设静态白名单断言(graft 本身就是完整性机制的实现层,
        # 多部件真变是常态) —— 但要留一份部件 diff 记录给上层,丢失/新增不静默。
        result["部件diff"] = diff_parts(original, modified).summary()
    return result


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
