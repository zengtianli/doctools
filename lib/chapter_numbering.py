#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""章号 SSOT 解析器(通用) —— 报告/标书「章号从几开始」config 的读取与派生。

一个项目建一份 chapters.yaml(见 chapter_renumber.py 顶部 schema),声明 number_base + 章序;
本 lib 把它派生成: 显示章号 / 合并序列 / 章头块 / 宽图关键词 / 整数章号映射 / 位移目标路径。
报告与标书通用: 项目的 merge/gen 脚本 import 本 lib 取派生值,renumber 引擎 import 本 lib 做位移,
章号只存 chapters.yaml 一处,散在多处(文件名/H1/图号/FACTS/大纲)全派生 —— 铁律#5 SSOT。

用法(项目侧):
    from chapter_numbering import ChapterNumbering
    cn = ChapterNumbering("技术标/chapters.yaml")
    cn.sequence(); cn.canonical_titles(); cn.header_block(); cn.wide_figure_keywords()
"""
from pathlib import Path

import yaml

# 位移目标默认值(标书/报告约定,均相对 config 所在目录)。项目可在 chapters.yaml
# 的 renumber_targets: 下逐键覆盖。
DEFAULT_TARGETS = {
    "chapters_glob": "chapters/ch*.md",       # 章节 md(文件名带 ch<号>- + H1 首行带号)
    "figure_png_glob": "素材图片/成图/图*.png",  # 成图 PNG(文件名带 图<号>-)
    "facts_file": "FACTS.md",                  # facts-machine 机检块(键=章号)
    "outline_glob": "03-*.md",                # 目录大纲(第N章/chN/列表前导号)
    "caption_prefix": "【图",                   # 正文图题注前缀
}


class ChapterNumbering:
    def __init__(self, config_path):
        self.config_path = Path(config_path).resolve()
        self.root = self.config_path.parent  # 目标路径相对此目录

    def load(self):
        with open(self.config_path, encoding="utf-8") as f:
            return yaml.safe_load(f)

    def targets(self):
        cfg = self.load()
        t = dict(DEFAULT_TARGETS)
        t.update(cfg.get("renumber_targets") or {})
        return t

    def resolved(self, cfg=None):
        """[(kind, num_str, slug, title)]。kind='file' 有独立文件; 'header' 为含 subs 章的章头。"""
        cfg = cfg or self.load()
        base = int(cfg["number_base"])
        out = []
        for i, ch in enumerate(cfg["sequence"]):
            num = base + i
            subs = ch.get("subs")
            if subs:
                out.append(("header", str(num), ch["slug"], ch["title"]))
                for j, sub in enumerate(subs, 1):
                    out.append(("file", f"{num}.{j}", sub["slug"], sub["title"]))
            else:
                out.append(("file", str(num), ch["slug"], ch["title"]))
        return out

    def canonical_titles(self):
        return {num: title for _k, num, _s, title in self.resolved()}

    def sequence(self):
        return [(kind, num) for kind, num, _s, _t in self.resolved()]

    def header_block(self):
        cfg = self.load()
        base = int(cfg["number_base"])
        for i, ch in enumerate(cfg["sequence"]):
            if ch.get("subs"):
                return f"# {base + i} {ch['title']}\n\n{cfg.get('header_body', '').strip()}"
        return ""

    def wide_figure_keywords(self):
        return list(self.load().get("wide_figure_keywords", []))

    def integer_map(self, chapters_dir):
        """{当前整数章号: 目标整数章号}。当前号从磁盘 ch<N>-<slug>.md 取(含 subs 章用首 sub)。"""
        import re
        cfg = self.load()
        base = int(cfg["number_base"])
        chapters_dir = Path(chapters_dir)
        imap = {}
        for i, ch in enumerate(cfg["sequence"]):
            target = base + i
            slug = ch["subs"][0]["slug"] if ch.get("subs") else ch["slug"]
            cur = None
            for p in sorted(chapters_dir.glob(f"ch*-{slug}.md")):
                m = re.match(r"ch(\d+)", p.name)
                if m:
                    cur = int(m.group(1))
                    break
            if cur is None:
                raise SystemExit(f"错误: 未找到 slug={slug} 的章节文件,无法定位现号")
            imap[cur] = target
        return imap
