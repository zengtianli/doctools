#!/usr/bin/env python3
"""restyle — 按同源 golden 把段落 pStyle 精确移植回「样式被剥光」的稿子（surgical · OLE 安全）。

Why（2026-06-25 天台 0624 vs 0625）：
  用户把 0624 的院样式(ZDWP*)全扒光做成扁平稿（3215 段 pStyle 全空，靠直接格式硬撑），
  要求「排版成 0625 那样」。golden 0625 与 0624 内容 99% 同源、且套了完整 ZDWP 样式。
  → 按段落文本把 golden 的 pStyle 逐段搬回，比 styles.py 启发式猜样式可靠得多。
  装帧引擎 chrome 靠 pStyle 找章边界，restyle 是它的前置必需步。

策略（保守 · 只补不覆盖）：
  · 从 golden 建 文本→pStyle 映射（同文本多 pStyle 取众数，记歧义）。
  · 遍历目标件正文段：仅当 ① 该段当前**无 pStyle** ② golden 有同文本且该文本 pStyle 唯一/众数明确
    → 套上。已有 pStyle 的段**跳过**（不覆盖人工样式）。
  · 文本归一：strip + 压缩内部连续空白（caption 三空格/双空格差异不影响匹配）。

surgical：只重写 word/document.xml，其余 zip 项（媒体/embeddings/OLE/OMML 公式）verbatim。

Usage:
  python3 restyle.py target.docx --ref golden.docx --check          # 只读：报可套/歧义/无源
  python3 restyle.py target.docx --ref golden.docx --apply          # 套样式 + .bak
  python3 restyle.py target.docx --ref golden.docx --apply --no-backup
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def q(tag: str) -> str:
    return W + tag


_WS = re.compile(r"\s+")


def _norm(s: str) -> str:
    return _WS.sub(" ", s).strip()


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(q("t")))


def _pstyle(p):
    pPr = p.find(q("pPr"))
    if pPr is None:
        return None
    st = pPr.find(q("pStyle"))
    return st.get(q("val")) if st is not None else None


def _body_root(path: Path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _ref_map(ref: Path):
    """golden 文本→pStyle 票数。多 pStyle 同文本 → Counter（取众数 + 标歧义）。"""
    root = _body_root(ref)
    body = root.find(q("body"))
    votes = defaultdict(Counter)
    for p in body.iter(q("p")):
        txt = _norm(_para_text(p))
        st = _pstyle(p)
        if txt and st:
            votes[txt][st] += 1
    return votes


def _resolve(votes_for_text: Counter):
    """众数 pStyle；并列(歧义) → None。"""
    if not votes_for_text:
        return None
    top = votes_for_text.most_common(2)
    if len(top) >= 2 and top[0][1] == top[1][1]:
        return None  # 票数并列 = 歧义，不套
    return top[0][0]


def _scan(target: Path, votes):
    """返回 (待套[(p,text,style)], 已有样式数, 歧义数, 无源数)。"""
    root = _body_root(target)
    body = root.find(q("body"))
    todo, kept, ambig, nosrc = [], 0, 0, 0
    for p in body.iter(q("p")):
        txt = _norm(_para_text(p))
        if not txt:
            continue
        if _pstyle(p) is not None:
            kept += 1
            continue
        vt = votes.get(txt)
        if not vt:
            nosrc += 1
            continue
        st = _resolve(vt)
        if st is None:
            ambig += 1
            continue
        todo.append((p, txt, st))
    return root, todo, kept, ambig, nosrc


def cmd_check(target: Path, ref: Path) -> int:
    votes = _ref_map(ref)
    _, todo, kept, ambig, nosrc = _scan(target, votes)
    print(f"[restyle 机检 · 对照 {ref.name}] {target.name}")
    print(f"  golden 文本→pStyle 映射条目: {len(votes)}")
    print(f"  目标件已有 pStyle（跳过）  : {kept}")
    print(f"  可套样式（无→有, 唯一源）  : {len(todo)}")
    print(f"  歧义（同文本多 pStyle 并列）: {ambig}")
    print(f"  无源（golden 无同文本）    : {nosrc}")
    if todo:
        dist = Counter(s for _, _, s in todo)
        print(f"  将套样式分布: {dict(dist.most_common())}")
        print("✗ 有正文段待套样式（样式被剥）")
        return 2
    print("✓ 无待套样式段")
    return 0


def cmd_apply(target: Path, ref: Path, no_backup: bool) -> int:
    votes = _ref_map(ref)
    root, todo, kept, ambig, nosrc = _scan(target, votes)
    if not todo:
        print(f"[restyle] {target.name}: 无需修改（{kept} 段已有样式 / "
              f"歧义{ambig} / 无源{nosrc}）")
        return 0

    for p, _txt, st in todo:
        pPr = p.find(q("pPr"))
        if pPr is None:
            pPr = etree.Element(q("pPr"))
            p.insert(0, pPr)
        ps = etree.Element(q("pStyle"))
        ps.set(q("val"), st)
        pPr.insert(0, ps)  # pStyle 必须是 pPr 首子元素

    if not no_backup:
        bak = target.with_suffix(target.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(target, bak)
        print(f"  备份 → {bak.name}")

    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    with zipfile.ZipFile(target) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc
            zout.writestr(item, data)
    tmp.replace(target)
    dist = Counter(s for _, _, s in todo)
    print(f"[restyle] {target.name}: 对照 golden 套样式 {len(todo)} 段 {dict(dist.most_common())}"
          f"（保留已有 {kept} / 歧义跳过 {ambig} / 无源 {nosrc}）")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="按同源 golden 移植 pStyle（surgical）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--ref", type=Path, required=True, help="同源 golden（pStyle 源）")
    ap.add_argument("--check", action="store_true", help="只读机检, exit2=有待套")
    ap.add_argument("--apply", action="store_true", help="套样式")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1
    if not a.ref.exists():
        print(f"找不到参照: {a.ref}", file=sys.stderr)
        return 1
    if a.apply:
        return cmd_apply(a.docx, a.ref, a.no_backup)
    return cmd_check(a.docx, a.ref)


if __name__ == "__main__":
    sys.exit(main())
