#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""renum.py — 编号/题注位移与重排族（原 3 个入口脚本 2026-07-31 合并为子命令）。

子命令（CLI 表面与原脚本零改动，banner/usage/退出码逐字保留）:
  chapter <chapters.yaml> [--apply]                原 chapter_renumber.py（md 侧章号位移引擎，config 驱动）
  tabfig  <chapters.yaml|章节目录> [--apply|--check]  原 tabfig_align.py（md 侧 表/图 题注号对齐章号）
  figures <docx> [...]                             原 docx_renumber_figures.py（docx 侧图号重排+引用同步）

exit 契约: chapter 0/3 · tabfig 0/1/2/3（--check 漂移=2）· figures 0/2/3。
**3 = 枚举为空**（2026-08-03 加，见下面 EMPTY_RC 一节）：一个对象都没枚举到时
旧实现三条路径全打通过语 + exit 0，是铁律「拒绝在空集上报绿」的直接违例。
2 的语义原样不动（真发现了问题），所以老的 `rc != 0 = 有问题` 判断不会被 3 骗到，
但把 3 当失败的调用方要自己放行 —— 本仓两处已改（doc_dispatch.do_renum /
typeset_pipeline 步骤③ 的 keep_rcs）。
figures 的 WriteGate 并发门与 assert_parts_intact 部件完整性断言原样保留；
_body_start_idx / serialize 惯用法 2026-07-31 起改用 lib/docx_surgical 的 SSOT 版
（body_start_idx 抽取时即与本脚本逐字一致，见该文件自注）。
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_write_gate import WriteGate  # noqa: E402  原地写回并发门（同目录 SSOT）

sys.path.append(str(Path(__file__).resolve().parents[2] / "lib"))
import caption_re  # noqa: E402  题注判据 SSOT
from chapter_numbering import ChapterNumbering  # noqa: E402
from docx_parts import assert_parts_intact  # noqa: E402
from docx_surgical import body_start_idx as _lib_body_start_idx  # noqa: E402
from docx_surgical import serialize as _serialize  # noqa: E402

# ════════════════════════════════════════════════════════════════════════════
# 空集判据（2026-08-03 立 · 三个子命令共用）
#
# 铁律：**枚举扫描类检查,枚举为空一律 exit != 0,拒绝在空集上报绿**。本文件原来
# 三条 figures 路径 + tabfig + chapter 全都在「一个对象都没枚举到」时打通过语并
# exit 0 —— 因为连续性判据在空 list 上**恒真**：
#     [] == list(range(1, 1))            → 默认(英文)模式打「✓ 连续 1..N」
#     all(v == … for v in {}.values())   → --cn-section 打「✓ 每节连续 1..k」
#     any(iss[k] for k in (...)) == False → --check 打「✓ 图序号与居中均合规」
# 于是「什么都没量到」被自检当成「量过了,全合规」发合规证。指错目录 / 指错 --kind /
# 判据字符类漏了一种短横,三种事故的表现全是这一句绿。
#
# ── 退出码为什么选 3 不选 2 ──
# 本文件的 **2 已经被三种「真发现了问题」占着**：--check 报断号/重号/未居中 ·
# 检测到重复图号 · 重编号后仍不连续。空集是**没能做出任何判定**,与之正交。
# 混进 2 会让上游把「这份文档没有图题」误读成「图号有问题」,实测两处会坏：
#   · doc_dispatch.do_renum  ——「图」那轮返 2 就 break,后面的「--kind 表」整轮不跑
#   · typeset_pipeline 步骤③ —— rc!=0 即回滚,--fix-center 做的居中会被一起撤销
# 3 同时对齐本仓既有先例：`para scan-ppr` 用 rc=3 表示「非错误、但需要人看」。
EMPTY_RC = 3


def _empty_set_exit(what: str, hint: str = ""):
    """枚举为空 → 打非通过语 + EMPTY_RC 退出。**禁在这里 return 0**。"""
    print(f"✗ 未发现任何{what} —— 枚举为空,未做任何判定"
          f"(空集不报绿,exit {EMPTY_RC})", file=sys.stderr)
    if hint:
        print(f"  {hint}", file=sys.stderr)
    sys.exit(EMPTY_RC)


# ════════════════════════════════════════════════════════════════════════════
# chapter — 原 chapter_renumber.py：章号统一位移引擎(通用·config 驱动)
#
# 改章号 = 只改项目 chapters.yaml 的 number_base 一个数,跑本引擎 --apply;
# 源码/正文不用手改。按「当前号→目标号」整数映射,位移所有引用点的前导章号:
#   ① 章节 md: H1 首行章号 + 图题注号(默认 【图 X-Y】)
#   ② 章节文件名: ch<号>-<slug>.md
#   ③ 成图 PNG: 图<号>-<k>_<名>.png   (两段式改名防撞号)
#   ④ FACTS.md facts-machine 块的章号键 "X"/"X.Y"
#   ⑤ 目录大纲: 第N章 / chN / 列表与标题前导号 / N—M 区间
# 目标路径来自 chapters.yaml 的 renumber_targets(缺省=标书/报告约定,见
# chapter_numbering.DEFAULT_TARGETS)。幂等: 磁盘已达标 → no-op。
# 位移后需项目侧重跑 number_headings(重派子号)+ 重渲。
# chapters.yaml 最小 schema 见 lib/chapter_numbering.py 顶部。
# ════════════════════════════════════════════════════════════════════════════

FACTS_FENCE = re.compile(r"(```yaml[ \t]+facts-machine[ \t]*\n)(.*?)(\n```)", re.DOTALL)


def find_config(argv):
    for a in argv:
        if a.endswith(".yaml") or a.endswith(".yml"):
            return Path(a).resolve()
    for cand in (Path("技术标/chapters.yaml"), Path("chapters.yaml")):
        if cand.exists():
            return cand.resolve()
    sys.exit("错误: 未找到 chapters.yaml,请显式传入路径。")


def remap_num(numstr, imap):
    head, dot, tail = numstr.partition(".")
    if not head.isdigit() or int(head) not in imap:
        return numstr
    return str(imap[int(head)]) + (dot + tail if dot else "")


def xform_chapter_md(text, imap, caption_prefix):
    n = [0]

    def h1_sub(m):
        new = remap_num(m.group(2), imap)
        if new == m.group(2):
            return m.group(0)
        n[0] += 1
        return f"{m.group(1)}{new}{m.group(3)}"

    def cap_sub(m):
        new = remap_num(m.group(2), imap)
        if new == m.group(2):
            return m.group(0)
        n[0] += 1
        return f"{m.group(1)}{new}{m.group(3)}"

    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if ln.startswith("# "):  # 只动 H1;H2+ 由 number_headings 从新 base 派生
            lines[i] = re.sub(r"^(#\s+)(\d+(?:\.\d+)*)(\s.*)$", h1_sub, ln)
    text = "\n".join(lines)
    text = re.sub(
        r"(" + re.escape(caption_prefix) + r"\s*)(\d+(?:\.\d+)*)(-\d+)", cap_sub, text
    )
    return text, n[0]


def xform_facts(text, imap):
    def block_sub(bm):
        body = re.sub(
            r'"(\d+(?:\.\d+)*)"(\s*):',
            lambda m: f'"{remap_num(m.group(1), imap)}"{m.group(2)}:',
            bm.group(2),
        )
        return bm.group(1) + body + bm.group(3)

    return FACTS_FENCE.sub(block_sub, text)


def xform_outline(text, imap):
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if ln.startswith(">"):  # 引语/说明行(语义)人工维护,引擎不碰
            continue
        s = ln
        s = re.sub(r"第(\d+)章", lambda m: f"第{remap_num(m.group(1), imap)}章", s)
        s = re.sub(r"ch(\d+(?:\.\d+)*)", lambda m: f"ch{remap_num(m.group(1), imap)}", s)
        s = re.sub(
            r"^(\s*(?:#{2,}\s+|-\s+))(\d+(?:\.\d+)*)(\s|　)",
            lambda m: f"{m.group(1)}{remap_num(m.group(2), imap)}{m.group(3)}",
            s,
        )
        s = re.sub(
            r"(\d+)—(\d+)",
            lambda m: f"{remap_num(m.group(1), imap)}—{remap_num(m.group(2), imap)}",
            s,
        )
        lines[i] = s
    return "\n".join(lines)


def new_leading(name_pat, name, imap):
    m = re.match(name_pat, name)
    if not m:
        return None
    new = remap_num(m.group(1), imap)
    if new == m.group(1):
        return None
    return new, m


def chapter_main(argv):
    apply = "--apply" in argv
    cfg_path = find_config(argv)
    cn = ChapterNumbering(cfg_path)
    root = cn.root
    t = cn.targets()
    ch_dir = (root / t["chapters_glob"]).parent
    imap = cn.integer_map(ch_dir)
    identity = all(k == v for k, v in imap.items())
    print(f"=== 章号位移引擎(通用) ({'APPLY' if apply else 'DRY-RUN'}) · {cfg_path} ===")
    print(f"number_base={cn.load()['number_base']} · 整数章号映射: {dict(sorted(imap.items()))}")
    # 空集：sequence 为空 → imap={} → `all(())` 恒真 → 旧实现打「磁盘已与 config 一致」
    # 报绿。(sequence 非空但磁盘缺章文件那条已由 chapter_numbering.integer_map
    # SystemExit 挡住,是本仓少数原本就 fail-closed 的枚举点。)
    if not imap:
        _empty_set_exit("章", f"chapters.yaml 的 sequence 为空: {cfg_path}")
    if identity:
        print("磁盘已与 config 一致,无需位移 (no-op)。")
        return

    # ① + ② 章节 md 内容 + 文件重命名
    md_files = sorted(ch_dir.glob(Path(t["chapters_glob"]).name))
    print(f"\n[md] {len(md_files)} 个章节文件:")
    for f in md_files:
        text = f.read_text(encoding="utf-8")
        new_text, nchg = xform_chapter_md(text, imap, t["caption_prefix"])
        r = new_leading(r"^ch(\d+(?:\.\d+)*)-", f.name, imap)
        nn = (f"ch{r[0]}-" + f.name[r[1].end():]) if r else None
        print(f"  {f.name}  H1/题注×{nchg}" + (f"  → {nn}" if nn else ""))
        if apply:
            if new_text != text:
                f.write_text(new_text, encoding="utf-8")
            if nn and nn != f.name:
                subprocess.run(["git", "mv", f.name, nn], cwd=str(ch_dir), capture_output=True)
                if (ch_dir / f.name).exists():
                    os.rename(ch_dir / f.name, ch_dir / nn)

    # ③ 成图 PNG 两段式改名
    png_dir = (root / t["figure_png_glob"]).parent
    pngs = sorted(png_dir.glob(Path(t["figure_png_glob"]).name)) if png_dir.is_dir() else []
    renames = []
    for p in pngs:
        r = new_leading(r"^图(\d+(?:\.\d+)*)(-\d+_)", p.name, imap)
        if r:
            renames.append((p, f"图{r[0]}" + p.name[r[1].start(2):]))
    print(f"\n[png] {len(renames)}/{len(pngs)} 个成图改名:")
    for p, nn in renames:
        print(f"  {p.name}  → {nn}")
    if apply and renames:
        for p, _ in renames:
            os.rename(p, p.with_name("__tmp__" + p.name))
        for p, nn in renames:
            os.rename(p.with_name("__tmp__" + p.name), png_dir / nn)

    # ④ FACTS
    facts = root / t["facts_file"]
    if facts.exists():
        ftxt = facts.read_text(encoding="utf-8")
        fnew = xform_facts(ftxt, imap)
        print(f"\n[FACTS] facts-machine 键位移: {'有改动' if fnew != ftxt else '无'}")
        if apply and fnew != ftxt:
            facts.write_text(fnew, encoding="utf-8")

    # ⑤ 目录大纲
    for outline in sorted(root.glob(t["outline_glob"])):
        otxt = outline.read_text(encoding="utf-8")
        onew = xform_outline(otxt, imap)
        print(f"[大纲] {outline.name} 章号位移: {'有改动' if onew != otxt else '无'}")
        if apply and onew != otxt:
            outline.write_text(onew, encoding="utf-8")

    print(f"\n{'✅ 已执行' if apply else '（干跑,加 --apply 执行）'}")
    print("后续: 项目侧 number_headings.py --apply(重派子号) → 重渲 docx。")
    print("提示: 引擎不碰散文类 SSOT(CLAUDE.md/总纲/大纲引语行),需人工同步语义。")


# ════════════════════════════════════════════════════════════════════════════
# tabfig — 原 tabfig_align.py：表/图题注号与所在章号对齐
# ════════════════════════════════════════════════════════════════════════════

# 原 tabfig_align.py 的模块 docstring —— 无参时打印的 usage 文案，逐字保留（对拍契约）。
TABFIG_DOC = """tabfig_align — 表/图题注号与所在章号对齐（单一职责,报告/标书通用）。

问题域: 章号位移(chapter_renumber)后,正文里的 `表 9.2-1　…` / `图 8-1` 类编号
前缀还是旧章号。本脚本只干一件事: **让每个 表/图 编号的章号段 = 它所在
章文件的章号**(从文件名 ch<N>-*.md 取),自愈式、幂等、与位移映射解耦——
不管中间改过几轮章号,跑一次就对。

不做的事(各归各的脚本): 章文件名/H1/PNG/FACTS 位移=chapter_renumber.py;
子标题派生=项目 number_headings.py;渲染=gen_bid_docx.py。

用法:
  python3 tabfig_align.py <chapters.yaml|章节目录> [--apply|--check]
    (默认)   干跑,列出待改项,exit 0
    --apply  写回文件
    --check  机检门: 有漂移 exit 2,干净 exit 0
"""

# 2026-08-01 判据下沉 lib/caption_re.md_token_pattern()（4 个**位置分组**被 align_text
# 的 sub 回调按位置解包重组，所以单列一份，但字符类走同一套）。行为变化：短横补齐
# 五种、章节点允许全角．—— 旧的只认 ASCII，`图3—1`/`图3－1` 在 md 侧一律漏检，
# `renum.py tabfig --check` 这道机检门对它们是瞎的。
TOKEN_RE = caption_re.md_token_pattern()
CH_RE = re.compile(r"^ch(\d+(?:\.\d+)*)-")


def chapter_files(arg: Path):
    if arg.is_dir():
        return sorted(arg.glob("ch*.md"))
    # chapters.yaml → 走总部 lib 的 targets(chapters_glob 相对 config 目录)
    cn = ChapterNumbering(arg)
    return sorted(cn.root.glob(cn.targets()["chapters_glob"]))


def align_text(text: str, ch_num: str):
    """返回 (新文本, [(旧token, 新token), ...])。只改章号段,保留序号与空白。"""
    changes = []

    def sub(m):
        kind, sp, num, tail = m.groups()
        if num == ch_num:
            return m.group(0)
        new = f"{kind}{sp}{ch_num}{tail}"
        changes.append((m.group(0), new))
        return new

    return TOKEN_RE.sub(sub, text), changes


def tabfig_main(argv):
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(TABFIG_DOC)
        return 1
    apply_ = "--apply" in argv
    check = "--check" in argv
    src = Path(args[0]).expanduser().resolve()

    # 空集①：一个 ch<N>-*.md 都没枚举到（指错目录 / 指错 chapters.yaml 的典型表现）。
    # 旧实现直接落到 total==0 打「✅ 全部对齐」—— 对着零个文件发合规证。
    scanned = [(f, m) for f in chapter_files(src) if (m := CH_RE.match(f.name))]
    if not scanned:
        _empty_set_exit("ch<N>-*.md 章节文件", f"扫描根: {src}")

    total, tokens = 0, 0
    for f, m in scanned:
        text = f.read_text(encoding="utf-8")
        tokens += len(TOKEN_RE.findall(text))
        new_text, changes = align_text(text, m.group(1))
        if not changes:
            continue
        total += len(changes)
        print(f"{f.name}  ({len(changes)} 处)")
        for old, new in changes:
            print(f"  {old}  →  {new}")
        if apply_:
            f.write_text(new_text, encoding="utf-8")

    # 空集②：文件枚举到了,但一个 表/图 编号 token 都没有 → 同样没有对象可对齐。
    if tokens == 0:
        _empty_set_exit("表/图 编号 token",
                        f"已扫 {len(scanned)} 个章节文件,判据 = lib/caption_re.md_token_pattern()")

    if total == 0:
        print(f"✅ 表/图编号与章号全部对齐,无需改动"
              f"（已扫 {len(scanned)} 个章节文件 · {tokens} 个编号 token）")
        return 0
    if apply_:
        print(f"✅ 已写回 {total} 处")
        return 0
    print(f"⚠ 共 {total} 处待对齐(干跑未写)。--apply 执行,--check 作机检门")
    return 2 if check else 0


# ════════════════════════════════════════════════════════════════════════════
# figures — 原 docx_renumber_figures.py：按文档出现顺序重排 Figure 编号
#           + 中文章节式 图X.Y-N 节内重排/补号 + 同步全部正文引用
#
# 为什么不是简单 str.replace：
#   1. **跨 run 分裂**：Word 常把 "Figure " 和 "23" 拆进相邻 w:r/w:t，朴素正则
#      按单个 w:t 改会漏掉数字在独立节点的 caption。按"段落级 concat 文本 +
#      字符偏移定位 → 写回覆盖该偏移所在 w:t"，跨 run 也能改。
#   2. **轮转/置换防碰撞**：重排常是置换（如 28→20, 20→21…27→28）。逐个 token
#      读旧值、原子写新值，绝不串改。
#   3. **排除 w:del**：track-changes 删除态(w:del/delText)里的旧文本不能算数、不改。
#   4. ⚠ **python-docx 陷阱**：`Paragraph.text` 静默漏掉 `w:ins`（修订插入）里的
#      run 文本 → 直接走 lxml 遍历 w:t（含 w:ins、排除 w:del）。
#
# 退出码：0 成功且重编号后 captions 连续 1..N；2 检测到重复图号（引用无法安全remap）。
# ════════════════════════════════════════════════════════════════════════════

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _under_del(t):
    return any(a.tag in (f"{W}del", f"{W}delText") for a in t.iterancestors())


def _visible_t_nodes(p):
    """段落内可见 w:t（含 w:ins，排除 w:del）按文档顺序。"""
    return [t for t in p.iter(f"{W}t") if not _under_del(t)]


def _apply_edits(nodes, edits):
    """把 (start,end,repl) 编辑（基于 nodes 拼接串的字符偏移，互不重叠）落到各 w:t。
    repl 仅在编辑**起点**所在节点写入；跨节点的尾部节点只做删除，避免 repl 重复。"""
    spans, pos = [], 0
    for n in nodes:
        tx = n.text or ""
        spans.append((pos, pos + len(tx)))
        pos += len(tx)
    edits = sorted(edits)
    for n, (ns, ne) in zip(nodes, spans):
        tx = n.text or ""
        local = []
        for s, e, repl in edits:
            if e <= ns or s >= ne:
                continue
            local.append((max(s, ns) - ns, min(e, ne) - ns, repl if s >= ns else ""))
        if local:
            local.sort()
            out, last = [], 0
            for cs, ce, repl in local:
                out.append(tx[last:cs]); out.append(repl); last = ce
            out.append(tx[last:])
            n.text = "".join(out)


def renumber(docx_path, prefix="Figure", dry_run=False):
    """返回 (root, remap, caption_order, ok)。dry_run 不修改 root（仍可读 remap）。"""
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))

    def ptext(p):
        return "".join(n.text or "" for n in _visible_t_nodes(p))

    cap_re = caption_re.en_caption_pattern(prefix)
    caption_order = []  # 现号，按物理顺序
    for p in paras:
        m = cap_re.match(ptext(p).strip())
        if m:
            caption_order.append(int(m.group(1)))

    # 重复图号 → 引用无法安全 remap
    if len(caption_order) != len(set(caption_order)):
        dup = [x for x in caption_order if caption_order.count(x) > 1]
        return root, {}, caption_order, False, sorted(set(dup))

    remap = {old: i + 1 for i, old in enumerate(caption_order)}
    if dry_run:
        return root, remap, caption_order, True, []

    # 引用匹配：Figure(s)/Fig. + 数字 + 可选 范围/列举（–,-,—,，and）
    cit = caption_re.en_citation_pattern(prefix)
    num = re.compile(r'\d+')
    for p in paras:
        nodes = _visible_t_nodes(p)
        full = "".join(n.text or "" for n in nodes)
        if not re.search(rf'{prefix}|Fig', full, re.I):
            continue
        edits = []
        for m in cit.finditer(full):
            for nm in num.finditer(m.group(0)):
                v = int(nm.group(0))
                if v in remap and remap[v] != v:
                    edits.append((m.start() + nm.start(), m.start() + nm.end(), str(remap[v])))
        if edits:
            _apply_edits(nodes, edits)
    return root, remap, caption_order, True, []


def _write(src_docx, root, out_path):
    new_xml = _serialize(root)   # lib/docx_surgical.serialize —— 与旧内联 tostring 逐字一致
    zin = zipfile.ZipFile(src_docx)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = new_xml if it.filename == "word/document.xml" else zin.read(it.filename)
            zout.writestr(it, data)
    # 部件完整性断言（fail-closed，一处覆盖全部调用点）：基线 = src_docx
    # （inplace 路是改前 .bak 副本，非 inplace 路是原件），只换 document.xml
    # （默认白名单），其余部件必须逐字节 verbatim。verbose=False 不动 stdout 契约。
    assert_parts_intact(src_docx, out_path, verbose=False)


def _verify(out_path, prefix):
    """重读输出确认 captions 连续 1..N。"""
    root = etree.fromstring(zipfile.ZipFile(out_path).read("word/document.xml"))
    cap_re = caption_re.en_caption_pattern(prefix)
    nums = []
    for p in root.iter(f"{W}p"):
        txt = "".join(n.text or "" for n in _visible_t_nodes(p)).strip()
        m = cap_re.match(txt)
        if m:
            nums.append(int(m.group(1)))
    return nums, nums == list(range(1, len(nums) + 1))


def _para_style(p):
    pPr = p.find(f"{W}pPr")
    if pPr is None:
        return None
    ps = pPr.find(f"{W}pStyle")
    return ps.get(f"{W}val") if ps is not None else None


def _has_drawing(p):
    return p.find(f".//{W}drawing") is not None or p.find(f".//{W}pict") is not None


def _ptext(p):
    return "".join(n.text or "" for n in _visible_t_nodes(p))


def _caption_content_len(s):
    """题注「内容」长度 = 去掉所有空白后的字符数。
    Word 常把题注用 tab/大段空格填满整行（表3.1-1 + 几十个空格 + 标题 + 单位），
    raw len 可达 125 → 误超「题注 <80」阈值被漏检（节内 gap 假阳性）。按内容长度量，
    既挡住真正的长正文段（首词碰巧是 图X.Y-N 的整句），又不冤枉空格填充的真题注。"""
    return len(re.sub(r"\s+", "", s))


def _center_style_ids(docx_path):
    """从 styles.xml 取「有效 jc=center」的段落样式 id 集合（含 basedOn 继承链）。"""
    try:
        sroot = etree.fromstring(zipfile.ZipFile(docx_path).read("word/styles.xml"))
    except Exception:
        return set()
    jc, based = {}, {}
    for st in sroot.iter(f"{W}style"):
        if st.get(f"{W}type") != "paragraph":
            continue
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        ppr = st.find(f"{W}pPr")
        j = ppr.find(f"{W}jc") if ppr is not None else None
        if j is not None:
            jc[sid] = j.get(f"{W}val")
        b = st.find(f"{W}basedOn")
        if b is not None:
            based[sid] = b.get(f"{W}val")

    def eff(sid, seen=None):
        seen = seen or set()
        if sid is None or sid in seen:
            return None
        seen.add(sid)
        return jc[sid] if sid in jc else eff(based.get(sid), seen)

    return {sid for sid in set(list(jc) + list(based)) if eff(sid) == "center"}


def _para_centered(p, center_ids):
    """段落有效对齐是否 center：显式 jc 优先，否则看 pStyle 是否继承 center。"""
    pPr = p.find(f"{W}pPr")
    jc = pPr.find(f"{W}jc") if pPr is not None else None
    if jc is not None:
        return jc.get(f"{W}val") == "center"
    return _para_style(p) in center_ids


def _body_start_idx(paras):
    """正文起始段索引 = 目录(TOC)字段之后。实现 = lib/docx_surgical.body_start_idx
    （抽取时与本脚本逐字一致，2026-07-31 起直接引 SSOT 版）。

    封面/批准页/落款/目录都在 TOC 之前；其中院 logo 图旁的文本（日期/署名/编制单位）
    **不是图题**，必须排除出 caption 采集——否则补号会把「二○二六年二月」误标成
    图X.Y-N（景宁 0313 成品踩坑：日期紧跟院 logo 图、又恰好同款 caption 样式）。
    判据 = 最后一个 TOC 字段 / PAGEREF 锚点段之后；无 TOC 则返 0（不排除，保持旧行为）。
    """
    return _lib_body_start_idx(paras)


def _collect_captions(paras, kind):
    """返回 (numbered, unnumbered, cap_style_ids)。
    numbered  : [(idx, sec, old_n)]  行首匹配 图X.Y-N 的题注
    unnumbered: [idx]                紧跟图片、属 caption 样式、无号的题注段（待补号）
    caption 样式 = numbered 题注所用 pStyle 的并集（据此识别同款无号题注，排除封面 logo 的非题注后段）。
    **前置区(封面/落款/目录)整段排除**（见 _body_start_idx）——图题只在正文。
    """
    body0 = _body_start_idx(paras)
    # 2026-08-01 判据下沉 lib/caption_re.RENUM_CN_CAPTION —— 写入端/回写端/自检端
    # 从三条各写各的 rf-string 收成同一个 spec。行为变化：短横补齐五种（旧的漏
    # U+2011 → 已编号的 `图1‑2` 被当无号题注 prepend 出双重编号，实测）、章号放开到
    # 任意深度（旧的最多两段，`图3.1.2-4` 同样被当无号）。
    cap_re = caption_re.pattern(caption_re.RENUM_CN_CAPTION.for_kind(kind))
    numbered, cap_styles = [], set()
    for idx, p in enumerate(paras):
        if idx < body0:
            continue
        s = _ptext(p).strip()
        if 0 < _caption_content_len(s) < 80:
            m = cap_re.match(s)
            if m:
                numbered.append((idx, m.group("sec"), int(m.group("seq"))))
                st = _para_style(p)
                if st:
                    cap_styles.add(st)
    # 附图/附表 = 独立扁平编号体系（附图1、附图2…），不属 图X.Y-N 范畴 → 排除，否则误判为无号
    appendix_re = caption_re.pattern(caption_re.APPENDIX_PREFIX_ANY)
    unnumbered = []
    for idx, p in enumerate(paras):
        if idx < body0 or idx == 0 or not _has_drawing(paras[idx - 1]):
            continue
        s = _ptext(p).strip()
        if not s or _caption_content_len(s) >= 80 or cap_re.match(s) or appendix_re.match(s):
            continue
        if cap_styles and _para_style(p) in cap_styles:
            unnumbered.append(idx)
    return numbered, unnumbered, cap_styles


def _prepend_caption_number(p, text):
    """在 caption 段首插入一个 run「图X.Y-N 」，rPr 克隆自原首 run（字体一致）。"""
    import copy
    first_r = p.find(f"{W}r")
    new_r = etree.Element(f"{W}r")
    if first_r is not None:
        rpr = first_r.find(f"{W}rPr")
        if rpr is not None:
            new_r.append(copy.deepcopy(rpr))
        first_r.addprevious(new_r)
    else:
        p.append(new_r)
    t = etree.SubElement(new_r, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def check_cn_section(docx_path, kind="图", check_center=True):
    """只读机检：返回 issues dict（全空=干净）。供 gate / report figs 调。
    {unnumbered:[题注文本], gaps:{sec:[缺号]}, duplicates:{sec:[重号]}, uncentered:[题注/段号],
     n_captions:int}

    ⚠ `n_captions`（= 已编号 + 无号题注的枚举总数）**不是 issue,是空集判据**：
    四个 issue 桶全空有两种截然不同的成因 —— 「量过了,全合规」与「一个题注都没量到」。
    调用方必须先看 n_captions 再看四个桶,否则就是在空集上报绿（本文件旧实现的原病）。
    """
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))
    numbered, unnumbered, cap_styles = _collect_captions(paras, kind)

    by_sec = {}
    for _, sec, n in numbered:
        by_sec.setdefault(sec, []).append(n)
    gaps, dups = {}, {}
    for sec, ns in by_sec.items():
        seen, d = set(), []
        for n in ns:
            if n in seen:
                d.append(n)
            seen.add(n)
        if d:
            dups[sec] = sorted(set(d))
        miss = sorted(set(range(1, max(ns) + 1)) - set(ns))
        if miss:
            gaps[sec] = miss

    uncentered = []
    if check_center and cap_styles:   # 仅对「有题注的报告」查居中，零题注文档跳过
        body0 = _body_start_idx(paras)   # 前置(封面/落款 logo)不算正文图，不查居中
        center_ids = _center_style_ids(docx_path)
        for idx, p in enumerate(paras):
            if idx < body0 or not _has_drawing(p) or _para_centered(p, center_ids):
                continue
            cap = _ptext(paras[idx + 1]).strip()[:30] if idx + 1 < len(paras) else ""
            uncentered.append(cap or f"段{idx}")

    return {
        "unnumbered": [_ptext(paras[i]).strip()[:30] for i in unnumbered],
        "gaps": gaps,
        "duplicates": dups,
        "uncentered": uncentered,
        "n_captions": len(numbered) + len(unnumbered),
    }


def center_figure_paragraphs(root, center_ids):
    """把所有含图段落的有效对齐改为 center（已居中的跳过）。返回改动数。
    前置(封面/落款 logo)不碰——只居中正文图（见 _body_start_idx）。"""
    n = 0
    paras = list(root.iter(f"{W}p"))
    body0 = _body_start_idx(paras)
    for idx, p in enumerate(paras):
        if idx < body0 or not _has_drawing(p) or _para_centered(p, center_ids):
            continue
        pPr = p.find(f"{W}pPr")
        if pPr is None:
            pPr = etree.Element(f"{W}pPr")
            p.insert(0, pPr)
        jc = pPr.find(f"{W}jc")
        if jc is None:
            jc = etree.SubElement(pPr, f"{W}jc")
        jc.set(f"{W}val", "center")
        n += 1
    return n


def renumber_cn_section(docx_path, kind="图", dry_run=False, supplement=True, fix_center=False):
    """中文章节式 图{X.Y}-{N} / 表{X.Y}-{N} 按**节内物理顺序**重排 + 补号 + 同步正文引用。

    编号是「节内」的（同 X.Y 前缀各自从 1 递增）：断号(图2.1-2 后直接 2.1-4)、重复号
    (两个图2.2-1) 按物理顺序顺排修正。**supplement=True 时**，紧跟图片、与已编号题注同
    样式、却无号的题注段（如「狮子口水库溢洪道出口段」）会被**补号**（节号据物理上最近的
    已编号题注推断，前优先后兜底；定位不到则记 warning 不补）。重复号有正文引用 → warning
    不动引用、caption 仍按物理位置改对。fix_center=True 同时把含图段落居中。

    kind: '图' 或 '表'。返回 (root, plan, caps, ok, warnings, n_found)。
    plan: [(para_idx, typ, sec, old_n, new_n)]；typ='num' 改号 / 'new' 补号。dry_run 不改 root。

    ⚠ `n_found` = `_collect_captions` 枚举到的题注总数（已编号 + 无号），2026-08-03 加。
    **不能拿 `len(plan)` 当空集判据**：`--no-supplement` 时无号题注不进 plan、`supplement`
    定位不到节号的也会被 skip 掉,于是「枚举到了但没排进计划」与「一个题注都没有」
    在 plan 上长得一模一样。空集必须问枚举端,不能问计划端。
    """
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))
    numbered, unnumbered, _ = _collect_captions(paras, kind)

    num_idx_sec = [(idx, sec) for idx, sec, _ in numbered]

    def infer_sec(idx):
        prev = [s for i, s in num_idx_sec if i < idx]
        if prev:
            return prev[-1]
        nxt = [s for i, s in num_idx_sec if i > idx]
        return nxt[0] if nxt else None

    items = sorted([(idx, "num", sec, old_n) for idx, sec, old_n in numbered]
                   + [(idx, "new", None, None) for idx in (unnumbered if supplement else [])])

    sec_ctr, plan, skipped = {}, [], []
    for idx, typ, sec, old_n in items:
        if typ == "new":
            sec = infer_sec(idx)
            if sec is None:
                skipped.append(idx)
                continue
        sec_ctr[sec] = sec_ctr.get(sec, 0) + 1
        plan.append((idx, typ, sec, old_n, sec_ctr[sec]))

    # (sec,old_n)->new_n 映射（仅已编号题注，供正文引用 remap）；同 key 多 new_n = 重复号冲突
    remap, conflict = {}, set()
    for idx, typ, sec, old_n, new_n in plan:
        if typ != "num":
            continue
        k = (sec, old_n)
        if k in remap and remap[k] != new_n:
            conflict.add(k)
        remap.setdefault(k, new_n)

    caps_compat = list(numbered)  # 兼容旧返回签名（仅已编号题注）
    n_found = len(numbered) + len(unnumbered)  # 空集判据,见 docstring

    if dry_run:
        return root, plan, caps_compat, True, sorted(conflict), n_found

    para_by_idx = dict(enumerate(paras))
    # 回写端与写入端（_collect_captions 的 cap_re）**同一个 spec 对象**，named group
    # 'seq' 定位序号。合并前这两端各写一条 rf-string，字符类可以各自漂移。
    n_after_re = caption_re.pattern(caption_re.RENUM_CN_CAPTION.for_kind(kind))

    # 1) caption 改写
    for idx, typ, sec, old_n, new_n in plan:
        p = para_by_idx[idx]
        if typ == "num":
            if old_n == new_n:
                continue
            nodes = _visible_t_nodes(p)
            full = "".join(n.text or "" for n in nodes)
            m2 = n_after_re.match(full)
            if m2:
                _apply_edits(nodes, [(*m2.span("seq"), str(new_n))])
        else:  # 补号
            _prepend_caption_number(p, f"{kind}{sec}-{new_n} ")

    # 2) 正文引用改写（排除**所有** caption 段：已编号 + 补号）；重复号有引用 → warning 不动
    #    ⚠ 必须含补号段——否则刚 prepend 的「图X.Y-N」会被本循环当正文引用再 remap 一次（曾踩）。
    cap_ids = {row[0] for row in plan}
    ref_re = caption_re.pattern(caption_re.RENUM_CN_REF.for_kind(kind))
    warnings = []
    for idx, p in enumerate(paras):
        if idx in cap_ids:
            continue
        nodes = _visible_t_nodes(p)
        full = "".join(n.text or "" for n in nodes)
        if kind not in full:
            continue
        edits = []
        for m in ref_re.finditer(full):
            k = (m.group("sec"), int(m.group("seq")))
            if k in conflict:
                warnings.append(f"{kind}{k[0]}-{k[1]}（重复号，引用需人工确认）")
                continue
            if k in remap and remap[k] != k[1]:
                edits.append((*m.span("seq"), str(remap[k])))
        if edits:
            _apply_edits(nodes, edits)

    if fix_center:
        n_c = center_figure_paragraphs(root, _center_style_ids(docx_path))
        if n_c:
            warnings.append(f"已居中 {n_c} 个含图段落")
    if skipped:
        warnings.append(f"{len(skipped)} 个无号题注无法定位章节（无相邻已编号题注），未补号")
    return root, plan, caps_compat, True, warnings, n_found


def _verify_cn(out_path, kind):
    """重读输出确认每个 sec 内 n 连续 1..k。返回 ({sec:[n...]}, all_ok, [双重编号段])。

    ⚠ **自检不能只复用写入端的连续性判据** —— 2026-08-01 实测：写入端字符类漏
    U+2011 时，已编号的 `图1‑2` 被当无号题注 prepend 成 `图1-3 图1‑2 …`，而自检用
    同一条瞎正则只看得见开头的 `图1-3`，于是对着被写坏的文档打印「✓ 每节连续」。
    连续性 check 在结构上看不见双重编号（叠上去的号本身就是连续的），所以这里额外
    扫「同一段里出现 ≥2 个题注编号」—— 这条判据与写入端**正交**，写入端漏了什么
    它照样抓得到，是 fail-closed 的那一半。
    """
    root = etree.fromstring(zipfile.ZipFile(out_path).read("word/document.xml"))
    cap_re = caption_re.pattern(caption_re.RENUM_CN_CAPTION.for_kind(kind))
    ref_re = caption_re.pattern(caption_re.RENUM_CN_REF.for_kind(kind))
    by_sec = {}
    doubled = []
    for p in root.iter(f"{W}p"):
        s = "".join(n.text or "" for n in _visible_t_nodes(p)).strip()
        if _caption_content_len(s) >= 80:
            continue
        m = cap_re.match(s)
        if m:
            by_sec.setdefault(m.group("sec"), []).append(int(m.group("seq")))
            if len(ref_re.findall(s)) > 1:
                doubled.append(s[:60])
    ok = all(v == list(range(1, len(v) + 1)) for v in by_sec.values()) and not doubled
    return by_sec, ok, doubled


def figures_main(argv):
    ap = argparse.ArgumentParser(prog="docx_renumber_figures.py",
                                 description="按出现顺序重排 docx 图号 + 同步正文引用")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--prefix", default="Figure")
    ap.add_argument("--cn-section", action="store_true",
                    help="中文章节式 图X.Y-N / 表X.Y-N 节内重排（修断号+重复号）")
    ap.add_argument("--kind", default="图", choices=["图", "表"],
                    help="--cn-section 模式下的题注类型，默认 图")
    ap.add_argument("--check", action="store_true",
                    help="只读机检（配 --cn-section）：报无序号题注/断号/重复号/未居中，"
                         "有问题 exit 2；一个题注都没枚举到 exit 3（空集不报绿）")
    ap.add_argument("--no-supplement", action="store_true",
                    help="--cn-section 重排时不给无号题注补号（默认补号）")
    ap.add_argument("--fix-center", action="store_true",
                    help="--cn-section 重排时顺带把含图段落居中")
    a = ap.parse_args(argv)
    write_gate = WriteGate(a.docx) if a.inplace else None  # 读入前 capture 基线

    # 只读机检模式（gate / report figs 调用入口）
    if a.check:
        kind = a.kind if a.cn_section else "图"
        iss = check_cn_section(a.docx, kind)
        bad = any(iss[k] for k in ("unnumbered", "gaps", "duplicates", "uncentered"))
        print(f"[{kind}题机检] {a.docx}")
        print(f"  无序号题注: {iss['unnumbered'] or '无'}")
        print(f"  节内断号  : {iss['gaps'] or '无'}")
        print(f"  重复号    : {iss['duplicates'] or '无'}")
        print(f"  未居中图片: {iss['uncentered'] or '无'}")
        # 空集先判：四个桶全空 ≠ 合规,也可能是一个题注都没枚举到（bad 恒 False）。
        if iss["n_captions"] == 0:
            _empty_set_exit(f"{kind}题注", f"文档: {a.docx}（--kind 选错 / 文档确实无{kind}）")
        print("✗ 发现问题，需重排/补号/居中" if bad else "✓ 图序号与居中均合规")
        sys.exit(2 if bad else 0)

    if a.cn_section:
        root, plan, caps, ok, warns, n_found = renumber_cn_section(
            a.docx, a.kind, dry_run=a.dry_run,
            supplement=not a.no_supplement, fix_center=a.fix_center)
        changes = [(f"{a.kind}{s}-{o or '—'}", f"{a.kind}{s}-{n}")
                   for _, typ, s, o, n in plan if typ == "new" or o != n]
        print(f"{a.kind}题数: {len(caps)} 已编号 + {sum(1 for _,t,*_ in plan if t=='new')} 补号（节内分组）")
        print(f"变动 (现→新): {changes or '无（已连续）'}")
        if warns:
            print(f"⚠ 提示: {warns}")
        if a.dry_run:
            print("[dry-run] 未写文件")
            if n_found == 0:
                _empty_set_exit(f"{a.kind}题注", f"文档: {a.docx}")
            return
        # 空集且没别的活可干 → 连输出文件都不产（产一份逐字节复制的 .renumbered.docx
        # 只会让下游以为「重排过了」）。--fix-center 例外：居中与题注枚举正交,那是
        # 真活,必须先写盘再报空集,否则 typeset 步骤③ 的居中成果会随非 0 rc 被回滚。
        if n_found == 0 and not a.fix_center:
            _empty_set_exit(f"{a.kind}题注", f"文档: {a.docx}（--kind 选错 / 文档确实无{a.kind}）")
        out = a.docx if a.inplace else (a.output or re.sub(r'\.docx$', '.renumbered.docx', a.docx))
        if a.inplace:
            write_gate.assert_unchanged()  # 源被 WPS/其他会话改过 → 拒写(逃生 DOCX_GATE_OK=1)
            shutil.copy2(a.docx, a.docx + ".bak")
            _write(a.docx + ".bak", root, out)   # 从 .bak 读、写回原文件，避免读写同路径截断
        else:
            _write(a.docx, root, out)
        by_sec, seq, doubled = _verify_cn(out, a.kind)
        print(f"已写: {out}")
        print(f"验证: 各节 {a.kind}号 = {by_sec}")
        if doubled:
            # fail-closed：同一段出现 ≥2 个题注编号 = 号被叠在原号前面（补号误判）
            print(f"✗ {len(doubled)} 个题注出现**双重编号**（补号叠在原号上）:",
                  file=sys.stderr)
            for t in doubled[:5]:
                print(f"    {t}", file=sys.stderr)
        elif n_found == 0:
            # 走到这里只可能是 --fix-center：盘已写（居中生效），但题注枚举是空的,
            # 不许打「✓ 每节连续 1..k」——`{}` 上 all() 恒真,那句话什么都没证明。
            _empty_set_exit(f"{a.kind}题注", f"已写 {out}（仅 --fix-center 的居中生效）")
        elif seq:
            print("✓ 每节连续 1..k")
        else:
            print("✗ 重编号后仍不连续，请检查")
        sys.exit(0 if seq else 2)

    root, remap, order, ok, dup = renumber(a.docx, a.prefix, dry_run=a.dry_run)
    if not ok:
        print(f"✗ 检测到重复图号 {dup}，正文引用无法安全 remap。先消重再跑。", file=sys.stderr)
        sys.exit(2)

    changes = {o: n for o, n in remap.items() if o != n}
    print(f"图题数: {len(order)}  现号顺序: {order}")
    print(f"变动 (现→新): {changes or '无（已连续）'}")

    if a.dry_run:
        print("[dry-run] 未写文件")
        if not order:
            _empty_set_exit(f" {a.prefix} 题注", f"文档: {a.docx}")
        return

    # 空集：`[] == list(range(1, 1))` 恒真 → 旧实现在零题注文档上打「✓ 连续 1..N」
    # 并 exit 0。英文线没有 --fix-center 这类正交副作用,直接在写盘前退出。
    if not order:
        _empty_set_exit(f" {a.prefix} 题注",
                        f"文档: {a.docx}（中文题注 图X.Y-N 走 --cn-section,不归本模式）")

    out = a.docx if a.inplace else (a.output or re.sub(r'\.docx$', '.renumbered.docx', a.docx))
    if a.inplace:
        write_gate.assert_unchanged()  # 源被 WPS/其他会话改过 → 拒写(逃生 DOCX_GATE_OK=1)
        shutil.copy2(a.docx, a.docx + ".bak")
        _write(a.docx + ".bak", root, out)   # 从 .bak 读、写回原文件，避免读写同路径截断
    else:
        _write(a.docx, root, out)
    nums, seq = _verify(out, a.prefix)
    print(f"已写: {out}")
    print(f"验证: captions = {nums}")
    print("✓ 连续 1..N" if seq else "✗ 重编号后仍不连续，请检查")
    sys.exit(0 if seq else 2)


# ════════════════════════════════════════════════════════════════════════════
# 首 token 分发
# ════════════════════════════════════════════════════════════════════════════

USAGE = """renum.py — 编号/题注位移与重排族

子命令:
  chapter <chapters.yaml> [--apply]                  章号位移引擎（原 chapter_renumber.py）
  tabfig  <chapters.yaml|章节目录> [--apply|--check]   表/图题注号对齐章号（原 tabfig_align.py）
  figures <docx> [--dry-run|--inplace|-o OUT] [--cn-section --kind 图|表 --check ...]
                                                     docx 图号重排+引用同步（原 docx_renumber_figures.py）
"""


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        print(USAGE)
        return 1
    if argv[0] in ("-h", "--help"):
        print(USAGE)
        return 0
    cmd, rest = argv[0], argv[1:]
    if cmd == "chapter":
        rc = chapter_main(rest)
        return 0 if rc is None else rc
    if cmd == "tabfig":
        return tabfig_main(rest)
    if cmd == "figures":
        rc = figures_main(rest)
        return 0 if rc is None else rc
    print(USAGE)
    print(f"未知子命令: {cmd}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
