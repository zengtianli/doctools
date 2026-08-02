#!/usr/bin/env python3
"""docx track-changes — 读取/写入修订标记和批注。

docx_tools.py（1705 行单体）2026-07-31 P2 拆解产物之一：本模块是 track-changes 段的
实现层，也是全家唯一写盘路径（DocxReviewer.save 走 zipfile 重打包 + WriteGate 并发门）。
docx_tools.py 仍是组合入口 + library re-export 面，CLI 契约声明在本模块
add_track_parser() 只写一遍。review 的 --no-include-ins / --no-strict 默认全开是
~/Work/CLAUDE.md L161 的机器强制条款，禁改。也可独立敲：

    python3 sub/docx_track.py track-changes read input.docx [--format md|json]
    python3 sub/docx_track.py track-changes review input.docx -o out.docx -r rules.json
"""

import argparse
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import UTC, datetime
from pathlib import Path

# ── surgical 收口：python-docx 存盘只重写点名的部件（炸开面 60→1）─────────────
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
import docx_safe_save  # noqa: E402,F401  详见 lib/docx_safe_save.py

# file_ops 的 canonical 家在 dev/lib（doctools/lib/file_ops.py 2026-05-21 已删；
# docx_cli / doc_dispatch 同样以 dev/lib 为准）。append 不 insert(0)。
_sys.path.append(str(_Path.home() / "Dev" / "tools" / "dev" / "lib"))
# docx_write_gate 在 scripts/document/（同 renum.py figures / docx_fmt.py text 共用的 SSOT）
_sys.path.append(str(_Path(__file__).resolve().parents[1]))

from lxml import etree  # noqa: E402

from docx_parts import assert_parts_intact  # noqa: E402  surgical 部件完整性断言
from docx_xml import R_NS, REL_COMMENTS, W, qn  # noqa: E402
from file_ops import clear_quarantine  # noqa: E402

from docx_write_gate import WriteGate  # noqa: E402  原地写回并发门（同目录 SSOT）


def read_track_changes(docx_path: str, fmt: str = "md") -> str:
    """读取 .docx 中的修订和批注，返回 Markdown 或 JSON"""
    changes = []
    comments = []

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
        tree = etree.fromstring(doc_xml)

        for ins in tree.iter(qn("w:ins")):
            author = ins.get(qn("w:author"), "未知")
            date = ins.get(qn("w:date"), "")
            text = _tc_extract_text(ins)
            if text.strip():
                changes.append({"type": "insert", "author": author, "date": date, "text": text})

        for dl in tree.iter(qn("w:del")):
            author = dl.get(qn("w:author"), "未知")
            date = dl.get(qn("w:date"), "")
            text = _tc_extract_del_text(dl)
            if text.strip():
                changes.append({"type": "delete", "author": author, "date": date, "text": text})

        if "word/comments.xml" in zf.namelist():
            comments_xml = zf.read("word/comments.xml")
            ctree = etree.fromstring(comments_xml)
            for comment in ctree.iter(qn("w:comment")):
                comments.append(
                    {
                        "id": comment.get(qn("w:id"), ""),
                        "author": comment.get(qn("w:author"), "未知"),
                        "date": comment.get(qn("w:date"), ""),
                        "text": _tc_extract_text(comment),
                    }
                )

    if fmt == "json":
        return json.dumps({"changes": changes, "comments": comments}, ensure_ascii=False, indent=2)

    # Markdown 格式
    lines = []
    if changes:
        lines.append("## 修订记录\n")
        for i, c in enumerate(changes, 1):
            icon = "插入" if c["type"] == "insert" else "删除"
            date_str = c["date"][:10] if c["date"] else ""
            lines.append(f"{i}. **{icon}** | {c['author']} | {date_str}")
            lines.append(f"   > {c['text']}\n")
    else:
        lines.append("## 修订记录\n\n无修订。\n")

    if comments:
        lines.append("## 批注\n")
        for c in comments:
            date_str = c["date"][:10] if c["date"] else ""
            lines.append(f"- **[{c['id']}]** {c['author']} ({date_str}):")
            lines.append(f"  > {c['text']}\n")

    return "\n".join(lines)


def _tc_extract_text(node) -> str:
    """从 XML 节点中提取所有 <w:t> 文本"""
    return "".join(t.text for t in node.iter(qn("w:t")) if t.text)


def _tc_extract_del_text(node) -> str:
    """从删除标记中提取 <w:delText> 文本"""
    parts = [t.text for t in node.iter(qn("w:delText")) if t.text]
    if not parts:
        parts = [t.text for t in node.iter(qn("w:t")) if t.text]
    return "".join(parts)


class DocxReviewer:
    """对 .docx 文件应用替换规则，生成带修订标记的新文件。保持原文格式不变。"""

    def __init__(
        self, docx_path: str, author: str = "CC审阅", include_ins: bool = True
    ):
        self.docx_path = docx_path
        self._write_gate = WriteGate(docx_path)  # 读入时 capture,save 回源文件前 assert
        self.author = author
        # include_ins：允许命中 <w:ins> 内的文字。**默认 True**。
        # 目标 docx 在修订模式协作、正文大量插入未接受时，关掉它规则一条都匹配不到
        # （2026-07-29 缙云 v5 实证：2960 处 w:ins，目标文本全在别人的 ins 里，
        #  默认 False 时静默报"0 处替换"）。关掉 = 制造假阴性，故逃生开关是 --no-include-ins。
        # w:del 内的文字任何情况下都跳过——那是已删除的字。
        self.include_ins = include_ins
        self.date = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.comment_id_counter = 0
        self.comments = []
        self.revision_id_counter = 100

        self.tmpdir = tempfile.mkdtemp(prefix="docx_review_")
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(self.tmpdir)

        doc_path = os.path.join(self.tmpdir, "word", "document.xml")
        self.doc_tree = etree.parse(doc_path)
        self.doc_root = self.doc_tree.getroot()
        self._init_comment_ids()

    def _init_comment_ids(self):
        comments_path = os.path.join(self.tmpdir, "word", "comments.xml")
        if os.path.exists(comments_path):
            ctree = etree.parse(comments_path)
            for c in ctree.getroot().iter(qn("w:comment")):
                cid = int(c.get(qn("w:id"), "0"))
                if cid >= self.comment_id_counter:
                    self.comment_id_counter = cid + 1

    def _next_rid(self) -> str:
        self.revision_id_counter += 1
        return str(self.revision_id_counter)

    def _next_comment_id(self) -> int:
        cid = self.comment_id_counter
        self.comment_id_counter += 1
        return cid

    def apply_rules(self, rules: list[dict]) -> int:
        """应用替换规则列表。返回成功替换的数量。"""
        count = 0
        for rule in rules:
            replace = rule.get("replace", rule["find"])
            # comment_only：只挂批注、不产生修订标记。显式声明优先；
            # 未声明时 find == replace 自动视为 comment_only（替换成同样的字没有意义）。
            comment_only = rule.get("comment_only", replace == rule["find"])
            n = self._apply_one_rule(
                rule["find"], replace, rule.get("comment"), comment_only,
                rule.get("para_index"),
            )
            count += n
        return count

    def _apply_one_rule(
        self,
        find: str,
        replace: str,
        comment: str | None,
        comment_only: bool = False,
        para_index=None,
    ) -> int:
        """para_index：可选，限定只在指定段落序号（body 内 w:p 的 0-based 文档序）
        上生效。int 或 int 列表。用于 find 文本本身不唯一（如多处同名"小结"标题）
        而必须逐处区分的场景。省略 = 全文命中，与历史行为一致。"""
        body = self.doc_root.find(qn("w:body"))
        if body is None:
            return 0
        if para_index is None:
            scope = None
        elif isinstance(para_index, int):
            scope = {para_index}
        else:
            scope = set(para_index)
        count = 0
        for pi, para in enumerate(body.iter(qn("w:p"))):
            if scope is not None and pi not in scope:
                continue
            while True:
                result = self._find_in_paragraph(para, find)
                if result is None:
                    break
                self._replace_in_paragraph(
                    para, result, find, replace, comment, comment_only
                )
                count += 1
                if comment_only:
                    # 原文保留在段内，再找必然命中同一处 → 死循环。
                    # comment_only 只标注每段首次出现。
                    break
        return count

    def _find_in_paragraph(self, para, find_text: str):
        """在段落中跨 run 搜索文本（跳过已有修订标记内的 run）"""
        runs = list(para.iter(qn("w:r")))
        if not runs:
            return None

        active_runs = []
        for r in runs:
            parent = r.getparent()
            if parent is not None and parent.tag == qn("w:del"):
                continue  # 已删除的文字永不参与匹配
            if (
                not self.include_ins
                and parent is not None
                and parent.tag == qn("w:ins")
            ):
                continue
            active_runs.append(r)
        if not active_runs:
            return None

        run_texts = []
        for r in active_runs:
            t_elem = r.find(qn("w:t"))
            run_texts.append(t_elem.text if t_elem is not None and t_elem.text else "")

        full_text = "".join(run_texts)
        idx = full_text.find(find_text)
        if idx == -1:
            return None

        start_pos, end_pos = idx, idx + len(find_text)
        cumulative = 0
        start_run_idx = end_run_idx = None
        start_offset = end_offset = 0

        for i, text in enumerate(run_texts):
            run_start = cumulative
            run_end = cumulative + len(text)
            if start_run_idx is None and run_end > start_pos:
                start_run_idx = i
                start_offset = start_pos - run_start
            if run_end >= end_pos:
                end_run_idx = i
                end_offset = end_pos - run_start
                break
            cumulative = run_end

        if start_run_idx is None or end_run_idx is None:
            return None

        return {
            "runs": active_runs[start_run_idx : end_run_idx + 1],
            "start_offset": start_offset,
            "end_offset": end_offset,
        }

    def _replace_in_paragraph(
        self, para, match, find_text, replace_text, comment_text, comment_only=False
    ):
        """拆分 run，插入 del/ins 标记，保持原有格式。

        comment_only=True 时不生成 del/ins，原文原样保留，只包上批注范围——
        用于给既有文字挂溯源批注（不改字）。
        """
        runs = match["runs"]
        start_offset = match["start_offset"]
        end_offset = match["end_offset"]

        # 继承第一个 run 的格式
        rpr_template = runs[0].find(qn("w:rPr"))
        if rpr_template is not None:
            rpr_template = copy.deepcopy(rpr_template)

        # 前缀文本（第一个 run 中匹配之前的部分）
        first_t = runs[0].find(qn("w:t"))
        first_text = first_t.text if first_t is not None and first_t.text else ""
        prefix_text = first_text[:start_offset]

        # 后缀文本（最后一个 run 中匹配之后的部分）
        last_t = runs[-1].find(qn("w:t"))
        last_text = last_t.text if last_t is not None and last_t.text else ""
        suffix_text = last_text[end_offset:]

        # 记录插入位置，删除原 run
        parent = runs[0].getparent()
        insert_pos = list(parent).index(runs[0])
        for r in runs:
            r.getparent().remove(r)

        # 构建替换节点
        nodes = []

        if prefix_text:
            nodes.append(self._make_run(prefix_text, rpr_template))

        # 批注起始
        comment_id = None
        if comment_text:
            comment_id = self._next_comment_id()
            cs = etree.Element(qn("w:commentRangeStart"))
            cs.set(qn("w:id"), str(comment_id))
            nodes.append(cs)

        if comment_only:
            # 只挂批注：原文原样放回，不产生任何修订标记
            nodes.append(self._make_run(find_text, rpr_template))
        else:
            # <w:del>
            rid = self._next_rid()
            del_node = etree.Element(qn("w:del"))
            del_node.set(qn("w:id"), rid)
            del_node.set(qn("w:author"), self.author)
            del_node.set(qn("w:date"), self.date)
            del_node.append(self._make_del_run(find_text, rpr_template))
            nodes.append(del_node)

            # <w:ins>
            ins_node = etree.Element(qn("w:ins"))
            ins_node.set(qn("w:id"), self._next_rid())
            ins_node.set(qn("w:author"), self.author)
            ins_node.set(qn("w:date"), self.date)
            ins_node.append(self._make_run(replace_text, rpr_template))
            nodes.append(ins_node)

        # 批注结束 + 引用
        if comment_text and comment_id is not None:
            ce = etree.Element(qn("w:commentRangeEnd"))
            ce.set(qn("w:id"), str(comment_id))
            nodes.append(ce)

            ref_run = etree.Element(qn("w:r"))
            ref_rpr = etree.SubElement(ref_run, qn("w:rPr"))
            ref_style = etree.SubElement(ref_rpr, qn("w:rStyle"))
            ref_style.set(qn("w:val"), "CommentReference")
            ref_elem = etree.SubElement(ref_run, qn("w:commentReference"))
            ref_elem.set(qn("w:id"), str(comment_id))
            nodes.append(ref_run)

            self.comments.append(
                {
                    "id": comment_id,
                    "author": self.author,
                    "date": self.date,
                    "text": comment_text,
                }
            )

        if suffix_text:
            nodes.append(self._make_run(suffix_text, rpr_template))

        if parent.tag == qn("w:ins") and not comment_only:
            # 命中的原文本身是一处未接受的插入（w:ins）。
            # w:ins 不能嵌套 w:ins，故：删除标记留在原 w:ins 内（w:ins>w:del = 插入后又删除，
            # 这是 Word 的标准表示），新增文字提到 w:p 层级、紧跟在原 w:ins 之后。
            inner = [n for n in nodes if n.tag != qn("w:ins")]
            outer = [n for n in nodes if n.tag == qn("w:ins")]
            for i, node in enumerate(inner):
                parent.insert(insert_pos + i, node)
            grandparent = parent.getparent()
            gi = list(grandparent).index(parent)
            for j, node in enumerate(outer):
                grandparent.insert(gi + 1 + j, node)
        else:
            for i, node in enumerate(nodes):
                parent.insert(insert_pos + i, node)

    def _make_run(self, text: str, rpr=None) -> etree._Element:
        """创建 <w:r>，继承原有 rPr 格式"""
        run = etree.Element(qn("w:r"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        t = etree.SubElement(run, qn("w:t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return run

    def _make_del_run(self, text: str, rpr=None) -> etree._Element:
        """创建 <w:del> 内部的 run（用 <w:delText>）"""
        run = etree.Element(qn("w:r"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        dt = etree.SubElement(run, qn("w:delText"))
        dt.text = text
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return run

    def _write_comments_xml(self):
        if not self.comments:
            return

        comments_path = os.path.join(self.tmpdir, "word", "comments.xml")
        if os.path.exists(comments_path):
            ctree = etree.parse(comments_path)
            croot = ctree.getroot()
        else:
            croot = etree.Element(qn("w:comments"), nsmap={"w": W, "r": R_NS})

        for c in self.comments:
            ce = etree.SubElement(croot, qn("w:comment"))
            ce.set(qn("w:id"), str(c["id"]))
            ce.set(qn("w:author"), c["author"])
            ce.set(qn("w:date"), c["date"])
            ce.set(qn("w:initials"), c["author"][:2])

            p = etree.SubElement(ce, qn("w:p"))
            etree.SubElement(p, qn("w:pPr"))
            r = etree.SubElement(p, qn("w:r"))
            rpr = etree.SubElement(r, qn("w:rPr"))
            rs = etree.SubElement(rpr, qn("w:rStyle"))
            rs.set(qn("w:val"), "CommentReference")
            etree.SubElement(r, qn("w:annotationRef"))

            r2 = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(r2, qn("w:t"))
            t.text = c["text"]
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        with open(comments_path, "wb") as f:
            f.write(etree.tostring(croot, xml_declaration=True, encoding="UTF-8", standalone=True))
        self._ensure_content_type("comments")
        self._ensure_rels("comments")

    def _ensure_content_type(self, part: str):
        ct_path = os.path.join(self.tmpdir, "[Content_Types].xml")
        ct_tree = etree.parse(ct_path)
        ct_root = ct_tree.getroot()
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        if part == "comments":
            part_name = "/word/comments.xml"
            ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
            for o in ct_root.iter(f"{{{ct_ns}}}Override"):
                if o.get("PartName") == part_name:
                    return
            o = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
            o.set("PartName", part_name)
            o.set("ContentType", ct)
            with open(ct_path, "wb") as f:
                f.write(etree.tostring(ct_tree, xml_declaration=True, encoding="UTF-8", standalone=True))

    def _ensure_rels(self, part: str):
        rels_path = os.path.join(self.tmpdir, "word", "_rels", "document.xml.rels")
        rels_tree = etree.parse(rels_path)
        rels_root = rels_tree.getroot()
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        if part == "comments":
            for rel in rels_root.iter(f"{{{rels_ns}}}Relationship"):
                if rel.get("Type") == REL_COMMENTS:
                    return
            max_id = 0
            for rel in rels_root.iter(f"{{{rels_ns}}}Relationship"):
                m = re.search(r"(\d+)", rel.get("Id", "rId0"))
                if m:
                    max_id = max(max_id, int(m.group(1)))
            new_rel = etree.SubElement(rels_root, f"{{{rels_ns}}}Relationship")
            new_rel.set("Id", f"rId{max_id + 1}")
            new_rel.set("Type", REL_COMMENTS)
            new_rel.set("Target", "comments.xml")
            with open(rels_path, "wb") as f:
                f.write(etree.tostring(rels_tree, xml_declaration=True, encoding="UTF-8", standalone=True))

    def save(self, output_path: str):
        """保存修改后的 .docx"""
        doc_path = os.path.join(self.tmpdir, "word", "document.xml")
        with open(doc_path, "wb") as f:
            f.write(etree.tostring(self.doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True))
        self._write_comments_xml()

        output_path = os.path.abspath(output_path)
        baseline = os.path.abspath(self.docx_path)
        snap = None
        if output_path == os.path.abspath(self.docx_path):
            self._write_gate.assert_unchanged()  # 原地写回:源被 WPS/其他会话改过 → 拒写(逃生 DOCX_GATE_OK=1)
            # 原地覆写会毁掉基线——先把改前源件快照到 tmpdir **之外**
            # （绝不能放 self.tmpdir：下面 os.walk 会把它打进包里）
            fd, snap = tempfile.mkstemp(prefix="docx_track_baseline_", suffix=".docx")
            os.close(fd)
            shutil.copy2(self.docx_path, snap)
            baseline = snap
        try:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, _dirs, files in os.walk(self.tmpdir):
                    for fn in files:
                        abs_path = os.path.join(root, fn)
                        zf.write(abs_path, os.path.relpath(abs_path, self.tmpdir))
            clear_quarantine(output_path)
            # 部件完整性断言（fail-closed）：extractall→os.walk 全目录重打包与
            # 137→35 截断事故同形状；tmpdir 里混入的多余文件也会被 unexpected-added 抓住。
            # comments.xml 仅在首次加批注时新建（_ensure_content_type/_ensure_rels 已报备）。
            assert_parts_intact(baseline, output_path,
                                allow_added={"word/comments.xml"}, verbose=False)
        except Exception:
            if snap is not None and os.path.exists(snap):
                shutil.copy2(snap, output_path)  # 原地写坏 → 从快照复原源件再抛
            raise
        finally:
            if snap is not None and os.path.exists(snap):
                os.unlink(snap)
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def cleanup(self):
        if os.path.exists(self.tmpdir):
            shutil.rmtree(self.tmpdir, ignore_errors=True)


def review_docx(
    input_path: str,
    output_path: str,
    rules: list[dict],
    author: str = "CC审阅",
    include_ins: bool = True,
    strict: bool = True,
) -> int:
    """便捷函数：对 .docx 应用替换规则，输出带修订标记的新文件。

    strict **默认 True**：任一规则 0 命中即抛错（fail-closed）——避免"静默 0 处替换"
    被当成执行成功（铁律 #2「守卫必须 fail-closed，禁静默跳过」）。
    确实允许部分规则落空时才传 strict=False / CLI --no-strict。
    """
    reviewer = DocxReviewer(input_path, author=author, include_ins=include_ins)
    try:
        if strict:
            misses = []
            count = 0
            for rule in rules:
                n = reviewer.apply_rules([rule])
                if n == 0:
                    misses.append(rule["find"][:60])
                count += n
            if misses:
                raise ValueError(
                    "以下规则 0 命中（strict 模式）：\n  - " + "\n  - ".join(misses)
                )
        else:
            count = reviewer.apply_rules(rules)
        reviewer.save(output_path)
        return count
    except Exception:
        reviewer.cleanup()
        raise


def cmd_track_changes(args):
    """track-changes 子命令入口"""
    if args.tc_command == "read":
        print(read_track_changes(args.input, args.format))
    elif args.tc_command == "review":
        with open(args.rules, encoding="utf-8") as f:
            rules = json.load(f)
        try:
            count = review_docx(
                args.input,
                args.output,
                rules,
                author=args.author,
                include_ins=bool(getattr(args, "include_ins", True)),
                strict=bool(getattr(args, "strict", True)),
            )
        except ValueError as e:
            print(f"错误：{e}", file=sys.stderr)
            print(
                "\n排查顺序："
                "\n  1. 目标文本是否被 run 拆散或含全角/半角差异 → 用 track-changes read 看原文"
                "\n  2. 目标文本是否在 <w:del> 里（已删除的字，任何情况都不命中）"
                "\n  3. 确实允许部分规则落空 → 加 --no-strict",
                file=sys.stderr,
            )
            sys.exit(1)
        print(f"完成：{count} 处替换已写入 {args.output}")
    elif args.tc_command == "compare":
        print("compare 功能将在 v2 实现。")
        sys.exit(1)
    else:
        print("请指定子命令：read、review 或 compare")
        sys.exit(1)


def add_track_parser(sub):
    """把 track-changes 子命令声明挂到组合入口的 subparsers（声明只写这一遍）。"""
    tc = sub.add_parser("track-changes", help="读取/写入修订标记和批注")
    tc_sub = tc.add_subparsers(dest="tc_command", help="track-changes 子命令")

    rp = tc_sub.add_parser("read", help="读取修订和批注")
    rp.add_argument("input", help="输入 .docx 文件")
    rp.add_argument("--format", "-f", choices=["md", "json"], default="md", help="输出格式 (默认: md)")

    wp = tc_sub.add_parser("review", help="写入修订标记")
    wp.add_argument("input", help="输入 .docx 文件")
    wp.add_argument("--output", "-o", required=True, help="输出 .docx 文件")
    wp.add_argument("--rules", "-r", required=True, help="替换规则 JSON 文件")
    wp.add_argument("--author", "-a", default="CC审阅", help="作者名")
    # 下面两项默认开启（主流用法），逃生开关是 --no-* 形式
    wp.add_argument("--no-include-ins", dest="include_ins", action="store_false",
                    default=True,
                    help="不命中 <w:ins> 内的文字。默认命中——目标 docx 在修订模式协作、"
                         "正文大量插入未接受时，关掉则规则一条都匹配不到（假阴性）")
    wp.add_argument("--no-strict", dest="strict", action="store_false",
                    default=True,
                    help="允许部分规则 0 命中。默认 fail-closed：任一规则 0 命中即报错退出，"
                         "防「静默 0 处替换」被当成执行成功")

    tcp = tc_sub.add_parser("compare", help="对比生成修订 (v2)")
    tcp.add_argument("original", help="原始 .docx")
    tcp.add_argument("modified", help="修改后 .docx")
    tcp.add_argument("--output", "-o", required=True, help="输出 .docx")

    return tc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx-track", description="docx track-changes — 修订标记读写（docx_tools 拆解实现模块）"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    add_track_parser(sub)
    args = parser.parse_args(argv)
    cmd_track_changes(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
