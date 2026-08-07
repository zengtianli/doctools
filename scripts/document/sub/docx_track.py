#!/usr/bin/env python3
"""docx track-changes — 读取/写入修订标记和批注。

docx_tools.py（1705 行单体）2026-07-31 P2 拆解产物之一：本模块是 track-changes 段的
实现层，也是全家唯一写盘路径（DocxReviewer.save 走 zipfile 重打包 + WriteGate 并发门）。
docx_tools.py 仍是组合入口 + library re-export 面，CLI 契约声明在本模块
add_track_parser() 只写一遍。review 的 --no-include-ins / --no-strict 默认全开是
~/Work/CLAUDE.md L161 的机器强制条款，禁改。也可独立敲：

    python3 sub/docx_track.py track-changes read input.docx [--format md|json]
    python3 sub/docx_track.py track-changes review input.docx -o out.docx -r rules.json
    python3 sub/docx_track.py track-changes compare 原.docx 改.docx -o 修订.docx

三条分支的分工：read=只读回显 · review=按规则表注入 · compare=拿两份文档比出修订。
compare 2026-08-03 落地（此前是 `print("compare 功能将在 v2 实现。")` 的空桩，
rc 还是 0 —— 敲下去像成功了、一件事没干）。粒度**段落级**，见下方 compare 段注释。
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
        self._reject_self_containing(rules)
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

    @staticmethod
    def _reject_self_containing(rules: list[dict]) -> None:
        """fail-closed：replace 原样包含 find ⇒ _apply_one_rule 的 while 会永远重新命中
        自己写进去的那段文字，进程 100% CPU 挂死、无任何报错（2026-08-06 缙云典型案例
        实测，一条 append 式规则空转 3 分钟才被发现）。这类规则没有任何合法用途，
        直接在入口拒绝，并给出可照做的改法。"""
        bad = []
        for i, rule in enumerate(rules, 1):
            replace = rule.get("replace", rule["find"])
            comment_only = rule.get("comment_only", replace == rule["find"])
            if not comment_only and rule["find"] in replace:
                bad.append((i, rule["find"], replace))
        if not bad:
            return
        lines = [
            "规则自包含：replace 里原样含有 find，替换后会再次命中自己 → 死循环挂死。",
        ]
        for i, find, replace in bad:
            lines.append(f"  规则 #{i}: find={find[:40]!r} 被完整包含在 replace={replace[:60]!r} 中")
        lines.append(
            "改法：让 replace 不再原样包含 find（改一两个字即可，如"
            "「…奠定坚实基础」→「…奠定了坚实基础。<后接新增内容>」）；"
            "只想挂批注不改字则显式写 \"comment_only\": true。"
        )
        raise ValueError("\n".join(lines))

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


# ── compare：两份 docx → 带 Word 修订标记的第三份 ─────────────────────────
#
# 粒度是**段落级**（difflib 对齐 body 顶层 w:p）：改过的段整段 w:del + 整段 w:ins，
# 不做 run 级字符粒度 diff。做法是 surgical：以**原稿**的 zip 为底，只重写
# word/document.xml，其余部件逐字节 verbatim + assert_parts_intact 断言。
#
# 为什么不用 python-docx 重建：它会剥 OLE / 原生图表（lib/docx_revise.py 开头记的
# 丢 11 个 chart 事故），而 compare 的输入正是「一份完整的真报告」。
#
# 范围外的差异（表格内、页眉页脚、文本框、改稿新增的图片）不会被静默吞掉：
# 一律进 out_of_scope 清单 → 打到 stderr + rc=3。见 _EXIT_* 常量。

_EXIT_OK = 0
_EXIT_IDENTICAL = 1        # 段落级无差异：不产出文件（禁「产一个没有修订的副本」冒充成功）
_EXIT_USAGE = 2            # 输入不合法 / 空文档
_EXIT_PARTIAL = 3          # 产出了修订件，但存在本引擎范围外的差异（未标注）

_REL_ATTR_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class CompareError(SystemExit):
    """输入问题一律带上下文抛出 —— 绝不静默产出一个「没有修订标记的副本」。"""


def _cmp_body(zf: zipfile.ZipFile):
    root = etree.fromstring(zf.read("word/document.xml"))
    body = root.find(qn("w:body"))
    if body is None:
        raise CompareError("✗ word/document.xml 里没有 w:body")
    return root, body


def _cmp_pstyle(p) -> str:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return ""
    ps = ppr.find(qn("w:pStyle"))
    return (ps.get(qn("w:val")) or "") if ps is not None else ""


def _cmp_text(el) -> str:
    """可见文本：跳过 w:del / w:moveFrom 子树 —— 那是已删除的字，不参与比对
    （否则原稿里别人未接受的删除会被当成「改稿删掉了它」再删一遍）。"""
    from docx_xml import in_deleted  # noqa: PLC0415  局部导入：只有 compare 用得到
    return "".join(t.text or "" for t in el.iter(qn("w:t")) if not in_deleted(t))


def _cmp_key(p) -> tuple[str, str]:
    """段落比对键。带上 pStyle：文本相同但一个是标题一个是正文 = 真的不同。
    文本不做空白归一化 —— 归一化会把「只差空格」判成相同，那是假绿。"""
    return (_cmp_pstyle(p), _cmp_text(p))


def _cmp_blocks(body) -> list[tuple[str, str]]:
    """body 顶层里**不参与段落 diff** 的块（表格等）的文本签名，用于范围外差异检测。"""
    out = []
    for k in body:
        ln = etree.QName(k).localname
        if ln in ("p", "sectPr"):
            continue
        out.append((ln, _cmp_text(k)))
    return out


def _cmp_part_texts(zf: zipfile.ZipFile) -> dict[str, str]:
    """页眉/页脚/脚注/尾注等旁路 part 的文本 —— 本引擎不碰它们，但差异要报出来。"""
    out = {}
    for name in zf.namelist():
        if re.fullmatch(r"word/(header\d*|footer\d*|footnotes|endnotes)\.xml", name):
            try:
                out[name] = _cmp_text(etree.fromstring(zf.read(name)))
            except etree.XMLSyntaxError:
                out[name] = "<unparsable>"
    return out


def _style_ids(zf: zipfile.ZipFile) -> set[str]:
    if "word/styles.xml" not in zf.namelist():
        return set()
    root = etree.fromstring(zf.read("word/styles.xml"))
    return {s.get(qn("w:styleId")) for s in root.iter(qn("w:style"))}


def _sanitize_inserted(p, orig_style_ids: set[str], same_numbering: bool,
                       notes: list[str]) -> None:
    """把改稿段落搬进原稿的 zip 之前，掐断它对**改稿 rels** 的引用。

    r:id / r:embed 指向的是改稿自己的 document.xml.rels；原稿包里没有那条关系，
    直接搬过来 = Word 开门就报「文件已损坏」。所以：
      w:hyperlink → 只摘掉 r:id（文字留着，链接失效）
      其余（图片 w:drawing / OLE w:object …）→ 连所在 w:r 一起摘掉
    两种都记进 notes（→ rc=3），绝不静默丢内容。
    """
    for el in list(p.iter()):
        rattrs = [k for k in el.attrib if k.startswith(_REL_ATTR_NS)]
        if not rattrs:
            continue
        if el.getroottree().getroot() is not p.getroottree().getroot():
            continue  # 已被上一轮连根摘走
        if el.tag == qn("w:hyperlink"):
            for k in rattrs:
                del el.attrib[k]
            notes.append(f"新增段里的超链接已降级为纯文字（原稿包内没有该关系）：{_cmp_text(el)[:24]!r}")
            continue
        holder = el
        while holder is not None and holder.tag != qn("w:r"):
            holder = holder.getparent()
            if holder is p:
                holder = None
                break
        victim = holder if holder is not None else el
        if victim.getparent() is not None:
            victim.getparent().remove(victim)
        notes.append(f"新增段里的 {etree.QName(el).localname} 已摘除"
                     f"（引用改稿 rels，搬进原稿包会开不了门）")
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return
    ps = ppr.find(qn("w:pStyle"))
    if ps is not None and orig_style_ids and ps.get(qn("w:val")) not in orig_style_ids:
        notes.append(f"新增段引用了原稿没有的样式 {ps.get(qn('w:val'))!r}（Word 会退回默认样式）")
    if ppr.find(qn("w:numPr")) is not None and not same_numbering:
        notes.append("新增段带自动编号(w:numPr)，而两份 numbering.xml 不同 —— 编号可能对不上")


def _mark_para_inserted(p, ids, meta: dict) -> int:
    """整段标插入：全部 run 包 w:ins + 段落标记也标 w:ins（Word 里能一键整段接受/拒绝）。"""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, qn("w:rPr"))
    if rpr.find(qn("w:ins")) is None:
        mark = etree.Element(qn("w:ins"))
        mark.set(qn("w:id"), str(next(ids)))
        mark.set(qn("w:author"), meta["author"])
        mark.set(qn("w:date"), meta["date"])
        rpr.insert(0, mark)
    n = 0
    for r in list(p.iter(qn("w:r"))):
        parent = r.getparent()
        if parent is None:
            continue
        if any(a.tag in (qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"))
               for a in r.iterancestors()):
            continue  # 改稿里本就带修订标记的 run，原样搬（别人的修订不改写）
        pos = list(parent).index(r)
        ins = etree.Element(qn("w:ins"))
        ins.set(qn("w:id"), str(next(ids)))
        ins.set(qn("w:author"), meta["author"])
        ins.set(qn("w:date"), meta["date"])
        parent.remove(r)
        ins.append(r)
        parent.insert(pos, ins)
        n += 1
    return n


def compare_docx(original: str, modified: str, output: str,
                 author: str = "CC对比", date: str | None = None) -> dict:
    """段落级对比 original vs modified，产出带 w:ins/w:del 的第三份 docx。

    返回 stats dict；不产出文件的两种情形（无差异 / 只有范围外差异）也如实返回，
    由调用方决定退出码。硬错误一律抛 CompareError（fail-closed）。
    """
    import difflib  # noqa: PLC0415

    from docx_revise import tracked_delete_runs  # noqa: PLC0415  修订注入引擎(总部 SSOT)

    src, dst, out = Path(original), Path(modified), Path(output)
    for label, p_ in (("原稿", src), ("改稿", dst)):
        if not p_.exists():
            raise CompareError(f"✗ 找不到{label}：{p_}")
    if out.resolve() in (src.resolve(), dst.resolve()):
        raise CompareError(f"✗ 输出路径与输入相同：{out} —— 拒绝覆盖输入件")

    meta = {"author": author,
            "date": date or datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}

    zs, zd = zipfile.ZipFile(src), zipfile.ZipFile(dst)
    try:
        root, body = _cmp_body(zs)
        _mroot, mbody = _cmp_body(zd)

        o_paras = [k for k in body if k.tag == qn("w:p")]
        m_paras = [k for k in mbody if k.tag == qn("w:p")]
        # 空集判红：0 段落的两份文件 difflib 会安静地报「无差异」——那是最会骗人的绿
        if not o_paras:
            raise CompareError(f"✗ 原稿 body 里一个顶层段落都没有：{src}")
        if not m_paras:
            raise CompareError(f"✗ 改稿 body 里一个顶层段落都没有：{dst}")

        used = [int(v) for e in root.iter() for v in [e.get(qn("w:id"))] if v and v.isdigit()]
        ids = iter(range(max(used, default=0) + 1000, 10 ** 8))

        # ── 范围外差异（本引擎不标注，但必须报出来）──────────────────────
        notes: list[str] = []
        if _cmp_blocks(body) != _cmp_blocks(mbody):
            notes.append("表格等非段落顶层块存在差异（段落级引擎不标注表格改动）")
        op, mp = _cmp_part_texts(zs), _cmp_part_texts(zd)
        for name in sorted(set(op) | set(mp)):
            if op.get(name) != mp.get(name):
                notes.append(f"旁路部件 {name} 文本有差异（页眉/页脚/脚注不参与段落 diff）")
        same_numbering = (("word/numbering.xml" in zs.namelist()) ==
                          ("word/numbering.xml" in zd.namelist())) and (
            "word/numbering.xml" not in zs.namelist()
            or zs.read("word/numbering.xml") == zd.read("word/numbering.xml"))
        o_styles = _style_ids(zs)

        # ── 段落级对齐 ─────────────────────────────────────────────────
        sm = difflib.SequenceMatcher(
            None, [_cmp_key(p) for p in o_paras], [_cmp_key(p) for p in m_paras],
            autojunk=False)
        n_del = n_ins = n_ins_runs = 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            if tag in ("delete", "replace"):
                for p_ in o_paras[i1:i2]:
                    tracked_delete_runs(p_, ids, meta, include_nontext=True)
                    n_del += 1
            if tag in ("insert", "replace"):
                new_ps = [copy.deepcopy(p_) for p_ in m_paras[j1:j2]]
                for np_ in new_ps:
                    _sanitize_inserted(np_, o_styles, same_numbering, notes)
                    n_ins_runs += _mark_para_inserted(np_, ids, meta)
                    n_ins += 1
                # 落位：原稿 o_paras[i1] 之前；i1 越界（改稿尾部新增）则贴在
                # 最后一个顶层段落之后（sectPr 之前——sectPr 必须留在 body 末尾）
                if i1 < len(o_paras):
                    for np_ in new_ps:
                        o_paras[i1].addprevious(np_)
                else:
                    tail = o_paras[-1]
                    for np_ in new_ps:
                        tail.addnext(np_)
                        tail = np_

        stats = {"deleted_paras": n_del, "inserted_paras": n_ins,
                 "inserted_runs": n_ins_runs, "out_of_scope": notes,
                 "orig_paras": len(o_paras), "mod_paras": len(m_paras),
                 "out": None}
        if n_del == 0 and n_ins == 0:
            return stats   # 段落级无差异 → 一个字节都不写（由调用方判非 0 退出）

        if out.exists():
            bak = out.with_suffix(out.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
            shutil.copy2(out, bak)
            stats["backup"] = str(bak)
        out.parent.mkdir(parents=True, exist_ok=True)
        new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                 standalone=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zs.infolist():
                zo.writestr(item, new_doc if item.filename == "word/document.xml"
                            else zs.read(item.filename))
        clear_quarantine(str(out))
        # 部件完整性断言（fail-closed）：底件是原稿，除 document.xml 外任何部件
        # 变动/增删都判红 —— 丢 chart / 截断部件那类事故的结构性防复发
        assert_parts_intact(str(src), str(out), verbose=False)
        stats["out"] = str(out)
        return stats
    finally:
        zs.close()
        zd.close()


def cmd_track_compare(args) -> int:
    """compare 分支：返回退出码（0 正常 / 1 无差异 / 2 输入不合法 / 3 有范围外差异）。"""
    for attr in ("original", "modified", "output"):
        if getattr(args, attr, None) is None:
            print(f"错误：compare 缺少参数 {attr}", file=sys.stderr)
            return _EXIT_USAGE
    try:
        st = compare_docx(args.original, args.modified, args.output,
                          author=getattr(args, "author", "CC对比"))
    except CompareError as e:
        print(str(e.code) if e.code else "✗ compare 失败", file=sys.stderr)
        return _EXIT_USAGE
    for n in st["out_of_scope"]:
        print(f"⚠ 范围外差异：{n}", file=sys.stderr)
    if st["out"] is None:
        print(f"无差异：两份文档在段落级完全相同"
              f"（原稿 {st['orig_paras']} 段 / 改稿 {st['mod_paras']} 段），"
              f"未产出文件 {args.output}")
        if st["out_of_scope"]:
            print("⚠ 但存在上列范围外差异 —— 本引擎标不了，别当成「两份一样」",
                  file=sys.stderr)
            return _EXIT_PARTIAL
        return _EXIT_IDENTICAL
    if st.get("backup"):
        print(f"旧件已备份 → {st['backup']}")
    print(f"完成：{st['deleted_paras']} 段标删除 / {st['inserted_paras']} 段标新增"
          f"（{st['inserted_runs']} 个 run 包 w:ins）→ {st['out']}")
    if st["out_of_scope"]:
        print(f"⚠ 上列 {len(st['out_of_scope'])} 处范围外差异未写进修订标记（rc=3）",
              file=sys.stderr)
        return _EXIT_PARTIAL
    return _EXIT_OK


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
        rc = cmd_track_compare(args)
        if rc:
            sys.exit(rc)   # 成功路径不 sys.exit：batch 的 _run_one 只 except Exception，
                           # SystemExit(0) 会穿过它把整个 batch worker 掀了
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

    tcp = tc_sub.add_parser(
        "compare", help="对比两份 docx，生成带 Word 修订标记的第三份（段落级粒度）",
        description="原稿 vs 改稿 → 带 w:ins/w:del 的第三份，可在 Word 里逐条接受/拒绝。",
        epilog=(
            "粒度 = **段落级**：改过的段整段标删除 + 整段标新增，不做 run 级字符 diff。\n"
            "范围 = body 顶层段落；表格内 / 页眉页脚 / 脚注 / 文本框的差异标不了，\n"
            "  但会打到 stderr 并让退出码变 3（绝不静默吞掉）。\n"
            "退出码：0 已产出修订件 · 1 段落级无差异(不产文件) · 2 输入不合法/空文档 ·\n"
            "  3 已产出，但存在上述范围外差异。"),
        formatter_class=argparse.RawDescriptionHelpFormatter)
    tcp.add_argument("original", help="原始 .docx")
    tcp.add_argument("modified", help="修改后 .docx")
    tcp.add_argument("--output", "-o", required=True, help="输出 .docx")
    tcp.add_argument("--author", "-a", default="CC对比", help="修订标记作者名")

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
