#!/usr/bin/env python3
"""
docx_format_clone.py — 提取 / 复刻 docx 版式（格式 SSOT）

为什么单独一个工具：
  - pandoc --reference-doc 只按「样式名」映射，抓不到「直排版」（run 直接 rPr，
    如公文标题 方正小标宋简体 二号居中 不走 heading 样式）→ 复刻不忠实。
  - docx_apply_template.py apply -t 只换 styles.xml，同样丢直排版。
  本工具按「角色（标题/正文/落款）」抽出 **有效格式**（直排版 + 样式叠加后的真实呈现），
  并用「段落外壳克隆」法复刻 —— 克隆范式件的段落 XML 外壳、只换文字，逐位忠实。

典型场景：人大代表/基层立法联系点 公文式交付件（标题 + 逐条正文 + 落款），
  范式件一份 → 所有同类交付件 apply 出同款版式，格式漂移结构上不可能。

子命令：
  extract <ref.docx> [-o profile.json]
      解析范式件，按角色抽有效格式 → 输出 format-profile.json + 人读报告。

  apply <content.md|.docx> --ref <ref.docx> [-o out.docx]
      用范式件的 标题壳/正文壳/落款壳 装载 content 文字 → 与范式逐位一致的 docx。
      content 解析：第一个 # 标题（或首段）= 标题；尾部无句末标点的连续段 = 落款；
      其余非空段 = 正文（逐条）。保留 ref 的 styles/fonts/numbering/sectPr。

用法：
  python docx_format_clone.py extract 范式.docx -o profile.json
  python docx_format_clone.py apply 草稿.md --ref 范式.docx -o 成品.docx
"""
import argparse
import copy
import json
import os
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))

try:
    from lxml import etree
except ImportError:
    print("❌ 需要 lxml: pip install lxml", file=sys.stderr)
    sys.exit(1)

from docx_xml import NSMAP, qn  # noqa: E402

try:
    from file_ops import clear_quarantine
except Exception:  # pragma: no cover
    def clear_quarantine(path):  # type: ignore
        os.system(f'xattr -d com.apple.quarantine "{path}" 2>/dev/null; true')

W = NSMAP["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"  # xml:space（内建命名空间）


# ─── XML 读取辅助 ────────────────────────────────────────────────────
def _read_document_xml(docx_path: str) -> etree._Element:
    with zipfile.ZipFile(docx_path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _body(root: etree._Element) -> etree._Element:
    return root.find(qn("w:body"))


def _para_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _first_rpr(p: etree._Element):
    """段落里第一个 run 的 rPr（run 直排版）。"""
    for r in p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            return rpr
    return None


def _eff_font(p: etree._Element) -> str | None:
    rpr = _first_rpr(p)
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for attr in ("eastAsia", "ascii", "hAnsi"):
            v = rfonts.get(qn(f"w:{attr}"))
            if v and v != "zh-CN":
                return v
    return None


def _eff_size_halfpt(p: etree._Element) -> int | None:
    rpr = _first_rpr(p)
    if rpr is None:
        return None
    sz = rpr.find(qn("w:sz"))
    if sz is not None:
        v = sz.get(qn("w:val"))
        return int(v) if v and v.isdigit() else None
    return None


def _ppr_attr(p: etree._Element):
    """(jc, firstLineChars or firstLine) 段落直排版。"""
    ppr = p.find(qn("w:pPr"))
    jc_val = None
    ind_chars = None
    if ppr is not None:
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            jc_val = jc.get(qn("w:val"))
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            flc = ind.get(qn("w:firstLineChars"))
            fl = ind.get(qn("w:firstLine"))
            if flc:
                ind_chars = int(flc) / 100.0
            elif fl:
                ind_chars = int(fl) / 100.0  # 近似（twip 时另算，公文件多用 Chars）
    return jc_val, ind_chars


_SENT_END = "。！？；…"


# ─── 角色分类（范式件） ───────────────────────────────────────────────
def classify_shells(root: etree._Element) -> dict:
    """返回 {'title': <w:p>, 'body': <w:p>, 'signature': <w:p>}（代表性外壳）。

    规则：
      正文主导字号 = 非空段里出现最多的字号（公文件正文段最多）
      title  = 起始 居中 或 字号大于主导字号 的段（标题可能拆成多行，取首行壳）
      signature = 尾部 无句末标点 的非空段（落款）
      body   = 主导字号、非标题、非落款、且优先「无 pStyle（纯 Normal）」的代表段
    """
    from collections import Counter
    paras = [p for p in _body(root).findall(qn("w:p"))]
    nonempty = [p for p in paras if _para_text(p).strip()]
    if not nonempty:
        raise ValueError("范式件无文字段落")

    sizes = [(_eff_size_halfpt(p) or 0) for p in nonempty]
    cnt = Counter(s for s in sizes if s)
    dominant = cnt.most_common(1)[0][0] if cnt else 0

    def is_titleish(p) -> bool:
        jc, _ = _ppr_attr(p)
        sz = _eff_size_halfpt(p) or 0
        return jc == "center" or (dominant > 0 and sz > dominant)

    # 标题：首个 titleish 段
    title_p = next((p for p in nonempty if is_titleish(p)), nonempty[0])

    # 落款：尾段无句末标点 → 落款；否则退而取尾段
    last_p = nonempty[-1]
    lt = _para_text(last_p).strip()
    sig_p = last_p if (lt and lt[-1] not in _SENT_END) else last_p

    # 正文：主导字号 + 非标题 + 非落款，优先无 pStyle 的纯 Normal 段
    cands = [p for p in nonempty
             if not is_titleish(p) and p is not sig_p
             and (_eff_size_halfpt(p) or 0) == dominant]
    body_p = None
    for p in cands:
        ppr = p.find(qn("w:pPr"))
        if ppr is None or ppr.find(qn("w:pStyle")) is None:
            body_p = p
            break
    if body_p is None:
        body_p = cands[0] if cands else next(
            (p for p in nonempty if p is not title_p and p is not sig_p), title_p)

    return {"title": title_p, "body": body_p, "signature": sig_p}


def shell_profile(p: etree._Element, role: str) -> dict:
    jc, ind = _ppr_attr(p)
    sz = _eff_size_halfpt(p)
    return {
        "role": role,
        "font_cn": _eff_font(p),
        "size_pt": (sz / 2.0) if sz else None,
        "size_hao": _halfpt_to_hao(sz) if sz else None,
        "align": jc,
        "first_line_indent_chars": ind,
    }


_HAO = {84: "初号", 72: "小初", 52: "一号", 44: "二号", 36: "小二",
        32: "三号", 30: "小三", 28: "四号", 24: "小四", 21: "五号"}


def _halfpt_to_hao(halfpt: int) -> str | None:
    return _HAO.get(halfpt)


# ─── extract ────────────────────────────────────────────────────────
def cmd_extract(args) -> int:
    ref = args.ref or args.input
    if not ref or not os.path.exists(ref):
        print(f"❌ 范式件不存在: {ref}", file=sys.stderr)
        return 1
    root = _read_document_xml(ref)
    shells = classify_shells(root)
    profile = {
        "source": os.path.basename(ref),
        "title": shell_profile(shells["title"], "title"),
        "body": shell_profile(shells["body"], "body"),
        "signature": shell_profile(shells["signature"], "signature"),
    }
    out = args.output or (os.path.splitext(ref)[0] + ".format-profile.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)
    print(f"✅ 格式 profile → {out}")
    for role in ("title", "body", "signature"):
        r = profile[role]
        print(f"  [{role:9}] {r['font_cn']} · {r['size_hao'] or r['size_pt']} · "
              f"对齐={r['align']} · 首行缩进={r['first_line_indent_chars']}字符")
    return 0


# ─── content 解析 ────────────────────────────────────────────────────
def parse_content(path: str) -> dict:
    """→ {'title': str, 'body': [str], 'signature': [str]}"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".md":
        blocks = _md_blocks(path)
    elif ext == ".docx":
        blocks = _docx_blocks(path)
    else:
        raise ValueError(f"content 须为 .md 或 .docx: {path}")
    if not blocks:
        raise ValueError("content 无文字")

    # 标题：首块（去 # 标记）
    title = re.sub(r"^#+\s*", "", blocks[0]).strip()
    rest = blocks[1:]

    # 落款：尾部连续 无句末标点 的块
    sig: list[str] = []
    while rest and rest[-1].strip() and rest[-1].strip()[-1] not in _SENT_END:
        sig.insert(0, rest.pop().strip())
    body = [b.strip() for b in rest if b.strip()]
    return {"title": title, "body": body, "signature": sig}


def _md_blocks(path: str) -> list[str]:
    raw = Path(path).read_text(encoding="utf-8")
    # 空行分块；块内多行折成一段（去硬换行）
    blocks = []
    for chunk in re.split(r"\n\s*\n", raw):
        line = " ".join(s.strip() for s in chunk.splitlines() if s.strip())
        line = re.sub(r"\s+", "", line) if _is_cjk_heavy(line) else line
        if line.strip():
            blocks.append(line)
    return blocks


def _is_cjk_heavy(s: str) -> bool:
    cjk = sum(1 for c in s if "一" <= c <= "鿿")
    return cjk > len(s) * 0.3


def _docx_blocks(path: str) -> list[str]:
    root = _read_document_xml(path)
    out = []
    for p in _body(root).findall(qn("w:p")):
        t = _para_text(p).strip()
        if t:
            out.append(t)
    return out


# ─── apply（外壳克隆） ────────────────────────────────────────────────
def _clone_with_text(shell: etree._Element, text: str) -> etree._Element:
    """克隆 shell 段落，保留 pPr + 首个 run 的 rPr，正文换为 text（单 run）。"""
    newp = copy.deepcopy(shell)
    # 取首 run 的 rPr 作模板
    rpr_tmpl = None
    src_rpr = _first_rpr(shell)
    if src_rpr is not None:
        rpr_tmpl = copy.deepcopy(src_rpr)
    # 删所有 run（保 pPr 等）
    for r in newp.findall(qn("w:r")):
        newp.remove(r)
    # 也删 hyperlink 内 run（简化：不处理超链接）
    r = etree.SubElement(newp, qn("w:r"))
    if rpr_tmpl is not None:
        r.append(rpr_tmpl)
    t = etree.SubElement(r, qn("w:t"))
    t.set(XML_SPACE, "preserve")
    t.text = text
    return newp


def _blank_like(shell: etree._Element) -> etree._Element:
    """造一个与 shell 同 pPr 的空段（间距用）。"""
    newp = copy.deepcopy(shell)
    for r in newp.findall(qn("w:r")):
        newp.remove(r)
    return newp


def cmd_apply(args) -> int:
    content_path = args.input
    ref = args.ref
    if not content_path or not os.path.exists(content_path):
        print(f"❌ content 不存在: {content_path}", file=sys.stderr)
        return 1
    if not ref or not os.path.exists(ref):
        print(f"❌ 范式件不存在: {ref}", file=sys.stderr)
        return 1

    content = parse_content(content_path)
    root = _read_document_xml(ref)
    shells = classify_shells(root)
    body_el = _body(root)
    sectPr = body_el.find(qn("w:sectPr"))

    # 重建 body：清空旧段，按 标题 / 正文 / 空行 / 落款 重填
    for child in list(body_el):
        if child.tag == qn("w:sectPr"):
            continue
        body_el.remove(child)

    new_children: list[etree._Element] = []
    if content["title"]:
        new_children.append(_clone_with_text(shells["title"], content["title"]))
    for para in content["body"]:
        new_children.append(_clone_with_text(shells["body"], para))
    emit_sig = content["signature"] and not getattr(args, "no_signature", False)
    if emit_sig:
        new_children.append(_blank_like(shells["body"]))  # 落款前空一行
        for line in content["signature"]:
            new_children.append(_clone_with_text(shells["signature"], line))

    # 插到 sectPr 之前
    if sectPr is not None:
        for el in new_children:
            sectPr.addprevious(el)
    else:
        for el in new_children:
            body_el.append(el)

    # 输出：复制 ref（保 styles/fonts/numbering/media），仅换 document.xml
    out = args.output or (os.path.splitext(content_path)[0] + "_styled_fixed.docx")
    new_doc_xml = etree.tostring(root, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(ref) as zin:
        names = zin.namelist()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name == "word/document.xml":
                    data = new_doc_xml
                zout.writestr(name, data)
    clear_quarantine(out)
    print(f"✅ 复刻 → {out}")
    print(f"   标题: {content['title'][:40]}")
    print(f"   正文: {len(content['body'])} 段 · 落款: {len(content['signature'])} 行")
    return 0


# ─── main ────────────────────────────────────────────────────────────
def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="docx_format_clone.py",
                                description="提取/复刻 docx 版式")
    sub = p.add_subparsers(dest="cmd")

    pe = sub.add_parser("extract", help="提取格式 → profile.json")
    pe.add_argument("input", help="范式 docx")
    pe.add_argument("--ref", help="（同 input）")
    pe.add_argument("-o", "--output", help="输出 profile.json")
    pe.set_defaults(func=cmd_extract)

    pa = sub.add_parser("apply", help="复刻格式（外壳克隆）")
    pa.add_argument("input", help="content（.md 或 .docx）")
    pa.add_argument("--ref", required=True, help="范式 docx（格式来源）")
    pa.add_argument("-o", "--output", help="输出 docx")
    pa.add_argument("--no-signature", action="store_true",
                    help="不发射落款段（交付件留空让对方自填署名）")
    pa.set_defaults(func=cmd_apply)

    args = p.parse_args(argv if argv is not None else sys.argv[1:])
    if not getattr(args, "func", None):
        p.print_help()
        return 0
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
