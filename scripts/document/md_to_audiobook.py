#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = ["edge-tts"]
# ///
"""
md_to_audiobook.py — Markdown 转有声书

格式：
  html         自包含 HTML 播放器 — 文字 + 音频 + 逐句高亮同步
  epub         EPUB3 + Media Overlay — 边听边看，文字逐句高亮
  m4b          M4B 有声书 — 纯音频，章节导航

用法：
  uv run md_to_audiobook.py input.md                        # → input.html
  uv run md_to_audiobook.py input.md --format m4b           # → input.m4b
  uv run md_to_audiobook.py input.md -o out.html            # 指定输出
  uv run md_to_audiobook.py input.md --voice zh-CN-YunxiNeural
  uv run md_to_audiobook.py input.md --heading-level 3      # 按 ### 拆章节
"""

import argparse
import asyncio
import base64
import html
import json
import re
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path

# ── 配置 ──────────────────────────────────────────────
DEFAULT_VOICE = "zh-CN-XiaoxiaoNeural"
# ─────────────────────────────────────────────────────


# ── MD 解析 ──────────────────────────────────────────

def split_chapters(md_text: str, level: int = 2) -> list[dict]:
    """按指定标题级别拆分章节"""
    prefix = "#" * level
    splits = []
    for m in re.finditer(rf"^{prefix}\s+(.+)$", md_text, re.MULTILINE):
        splits.append((m.start(), m.group(1).strip()))

    if not splits:
        title_match = re.match(r"^#\s+(.+)$", md_text, re.MULTILINE)
        title = title_match.group(1).strip() if title_match else "全文"
        return [{"title": title, "body": md_text}]

    chapters = []
    for i, (pos, title) in enumerate(splits):
        end = splits[i + 1][0] if i + 1 < len(splits) else len(md_text)
        chapters.append({"title": title, "body": md_text[pos:end]})
    return chapters


def clean_text(md_body: str) -> str:
    """清洗 Markdown 为可朗读纯文本（M4B/EPUB 用）"""
    text = md_body
    text = re.sub(r"^\|.*$", "", text, flags=re.MULTILINE)       # 表格
    text = re.sub(r"^-{3,}$", "", text, flags=re.MULTILINE)      # 水平线
    text = re.sub(r"^>\s*", "", text, flags=re.MULTILINE)         # 引用
    text = re.sub(r"^#{1,6}\s+", "", text, flags=re.MULTILINE)   # 标题标记
    text = re.sub(r"\*{1,2}(.+?)\*{1,2}", r"\1", text)           # 加粗/斜体
    text = re.sub(r"^[\s]*[-*+]\s+", "", text, flags=re.MULTILINE)  # 无序列表
    text = re.sub(r"^[\s]*\d+\.\s+", "", text, flags=re.MULTILINE)  # 有序列表
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)          # 链接
    text = re.sub(r"`([^`]+)`", r"\1", text)                      # 行内代码
    text = re.sub(r"```[\s\S]*?```", "", text)                    # 代码块
    text = re.sub(r"\n{3,}", "\n\n", text)                        # 多余空行
    return text.strip()


def _clean_inline(text: str) -> str:
    """清洗行内 Markdown 格式，保留纯文本"""
    text = re.sub(r"\*{1,2}(.+?)\*{1,2}", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.strip()


def parse_chapter_blocks(md_body: str) -> list[dict]:
    """将章节 MD 解析为结构化块（HTML 用）。

    返回 list of:
      {'type': 'heading', 'level': int, 'text': str}
      {'type': 'text', 'text': str}
    """
    blocks = []
    current_lines: list[str] = []
    in_code_block = False
    in_table = False

    def flush_text():
        if current_lines:
            joined = _clean_inline(" ".join(current_lines))
            if joined:
                blocks.append({"type": "text", "text": joined})
            current_lines.clear()

    for line in md_body.split("\n"):
        stripped = line.strip()

        # 代码块
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            flush_text()
            continue
        if in_code_block:
            continue

        # 表格（连续 | 开头的行，或分隔线 |---|---|）
        if stripped.startswith("|") or re.match(r"^[-:|]+$", stripped):
            flush_text()
            in_table = True
            continue
        if in_table and not stripped:
            in_table = False

        # 水平线
        if re.match(r"^[-*_]{3,}$", stripped):
            flush_text()
            continue

        # 空行 → 段落结束
        if not stripped:
            flush_text()
            continue

        # 标题
        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            flush_text()
            blocks.append({
                "type": "heading",
                "level": len(hm.group(1)),
                "text": _clean_inline(hm.group(2)),
            })
            continue

        # 引用
        cleaned = re.sub(r"^>\s*", "", stripped)
        # 列表项
        cleaned = re.sub(r"^[-*+]\s+", "", cleaned)
        cleaned = re.sub(r"^\d+\.\s+", "", cleaned)
        current_lines.append(cleaned)

    flush_text()
    return blocks


# ── 音频生成 ──────────────────────────────────────────

async def generate_audio_simple(text: str, output_path: Path, voice: str) -> float:
    """生成音频（M4B 用），返回时长秒"""
    import edge_tts
    communicate = edge_tts.Communicate(text, voice)
    await communicate.save(str(output_path))
    return _get_duration(output_path)


async def generate_audio_with_sync(
    text: str, output_path: Path, voice: str
) -> tuple[list[dict], float]:
    """生成音频 + 收集句子级时间戳（EPUB 用）。
    返回 (sentences_with_timestamps, duration_seconds)
    """
    import edge_tts

    communicate = edge_tts.Communicate(text, voice)
    audio_data = bytearray()
    sentences = []

    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            audio_data.extend(chunk["data"])
        elif chunk["type"] == "SentenceBoundary":
            sentences.append({
                "text": chunk["text"],
                "start": chunk["offset"] / 10_000_000,
                "end": (chunk["offset"] + chunk["duration"]) / 10_000_000,
            })

    output_path.write_bytes(bytes(audio_data))
    duration = _get_duration(output_path)

    # 修正最后一句的 end 不超过实际时长
    if sentences:
        sentences[-1]["end"] = min(sentences[-1]["end"], duration)

    return sentences, duration


def _get_duration(path: Path) -> float:
    r = subprocess.run(
        ["ffprobe", "-v", "quiet", "-show_entries", "format=duration",
         "-of", "csv=p=0", str(path)],
        capture_output=True, text=True,
    )
    return float(r.stdout.strip())


# ── M4B 构建 ──────────────────────────────────────────

def build_m4b(chapter_files: list[dict], output_path: Path, title: str) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        concat_file = tmpdir / "concat.txt"
        with open(concat_file, "w") as f:
            for ch in chapter_files:
                safe = str(ch["path"]).replace("'", "'\\''")
                f.write(f"file '{safe}'\n")

        merged = tmpdir / "merged.mp3"
        subprocess.run(
            ["ffmpeg", "-y", "-f", "concat", "-safe", "0",
             "-i", str(concat_file), "-c", "copy", str(merged)],
            capture_output=True, check=True,
        )

        meta = tmpdir / "metadata.txt"
        with open(meta, "w") as f:
            f.write(f";FFMETADATA1\ntitle={title}\nartist=edge-tts\n\n")
            t = 0
            for ch in chapter_files:
                s, e = int(t), int(t + ch["duration"] * 1000)
                f.write(f"[CHAPTER]\nTIMEBASE=1/1000\nSTART={s}\nEND={e}\ntitle={ch['title']}\n\n")
                t = e

        subprocess.run(
            ["ffmpeg", "-y", "-i", str(merged), "-i", str(meta),
             "-map_metadata", "1", "-c:a", "aac", "-b:a", "64k",
             "-movflags", "+faststart", str(output_path)],
            capture_output=True, check=True,
        )


# ── EPUB3 构建 ──────────────────────────────────────────

def _fmt_smil_time(seconds: float) -> str:
    """秒 → SMIL clock value"""
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds % 60
    return f"{h}:{m:02d}:{s:06.3f}"


def _fmt_duration(seconds: float) -> str:
    """秒 → OPF duration"""
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def _make_chapter_xhtml(title: str, sentences: list[dict], ch_idx: int) -> str:
    """生成章节 XHTML，每句一个 span"""
    lines = [f'    <h2>{html.escape(title)}</h2>']

    # 按句子分段，遇到文本中的换行拆分 <p>
    current_spans = []
    for i, sent in enumerate(sentences):
        sid = f"s{ch_idx:02d}_{i:04d}"
        text = sent["text"]

        # 如果句子内有换行，拆成多个 p
        parts = text.split("\n")
        for j, part in enumerate(parts):
            part = part.strip()
            if not part:
                if current_spans:
                    lines.append(f'    <p>{" ".join(current_spans)}</p>')
                    current_spans = []
                continue
            if j == 0:
                current_spans.append(f'<span id="{sid}">{html.escape(part)}</span>')
            else:
                # 换行后的部分归入同一个 span 时间段但新起一段
                if current_spans:
                    lines.append(f'    <p>{" ".join(current_spans)}</p>')
                    current_spans = []
                current_spans.append(f'<span id="{sid}">{html.escape(part)}</span>')

    if current_spans:
        lines.append(f'    <p>{" ".join(current_spans)}</p>')

    body_html = "\n".join(lines)
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head>
  <title>{html.escape(title)}</title>
  <link rel="stylesheet" type="text/css" href="style.css"/>
</head>
<body>
{body_html}
</body>
</html>'''


def _make_chapter_smil(sentences: list[dict], ch_idx: int, xhtml_file: str, audio_file: str) -> str:
    """生成章节 SMIL"""
    pars = []
    for i, sent in enumerate(sentences):
        sid = f"s{ch_idx:02d}_{i:04d}"
        pars.append(
            f'    <par id="par_{sid}">\n'
            f'      <text src="{xhtml_file}#{sid}"/>\n'
            f'      <audio src="{audio_file}" '
            f'clipBegin="{_fmt_smil_time(sent["start"])}" '
            f'clipEnd="{_fmt_smil_time(sent["end"])}"/>\n'
            f'    </par>'
        )

    body = "\n".join(pars)
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<smil xmlns="http://www.w3.org/ns/SMIL" xmlns:epub="http://www.idpf.org/2007/ops" version="3.0">
<body>
  <seq id="seq_{ch_idx:02d}" epub:textref="{xhtml_file}">
{body}
  </seq>
</body>
</smil>'''


def build_epub3(chapters: list[dict], output_path: Path, title: str) -> None:
    """打包 EPUB3 with Media Overlay"""
    book_id = str(uuid.uuid4())
    total_dur = sum(ch["duration"] for ch in chapters)

    css = '''body {
  font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
  line-height: 1.8;
  padding: 1em 1.5em;
  color: #333;
  max-width: 40em;
  margin: 0 auto;
}
h2 { color: #1a5276; border-bottom: 2px solid #aed6f1; padding-bottom: 0.3em; }
p { margin: 0.8em 0; text-indent: 2em; }
span { transition: background-color 0.3s; }
.-epub-media-overlay-active {
  background-color: #fff3cd;
  border-radius: 3px;
  padding: 1px 2px;
}'''

    items_opf = []
    spine_refs = []
    duration_meta = []

    with zipfile.ZipFile(output_path, "w") as zf:
        zf.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)

        zf.writestr("META-INF/container.xml", '''<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>''')

        zf.writestr("OEBPS/style.css", css)
        items_opf.append('<item id="css" href="style.css" media-type="text/css"/>')

        for i, ch in enumerate(chapters):
            ch_id = f"ch{i:02d}"
            xhtml_name = f"chapter_{i:03d}.xhtml"
            smil_name = f"chapter_{i:03d}.smil"
            audio_name = f"audio/chapter_{i:03d}.mp3"

            xhtml = _make_chapter_xhtml(ch["title"], ch["sentences"], i)
            zf.writestr(f"OEBPS/{xhtml_name}", xhtml)

            smil = _make_chapter_smil(ch["sentences"], i, xhtml_name, audio_name)
            zf.writestr(f"OEBPS/{smil_name}", smil)

            zf.writestr(
                f"OEBPS/{audio_name}",
                ch["audio_path"].read_bytes(),
                compress_type=zipfile.ZIP_STORED,
            )

            items_opf.append(
                f'<item id="{ch_id}" href="{xhtml_name}" '
                f'media-type="application/xhtml+xml" media-overlay="{ch_id}_overlay"/>'
            )
            items_opf.append(
                f'<item id="{ch_id}_overlay" href="{smil_name}" '
                f'media-type="application/smil+xml"/>'
            )
            items_opf.append(
                f'<item id="{ch_id}_audio" href="{audio_name}" '
                f'media-type="audio/mpeg"/>'
            )
            spine_refs.append(f'<itemref idref="{ch_id}"/>')
            duration_meta.append(
                f'<meta property="media:duration" refines="#{ch_id}_overlay">'
                f'{_fmt_duration(ch["duration"])}</meta>'
            )

        # Navigation
        nav_items = "\n".join(
            f'        <li><a href="chapter_{i:03d}.xhtml">{html.escape(ch["title"])}</a></li>'
            for i, ch in enumerate(chapters)
        )
        toc_xhtml = f'''<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>{html.escape(title)}</title></head>
<body>
  <nav epub:type="toc">
    <h1>目录</h1>
    <ol>
{nav_items}
    </ol>
  </nav>
</body>
</html>'''
        zf.writestr("OEBPS/toc.xhtml", toc_xhtml)
        items_opf.append(
            '<item id="toc" href="toc.xhtml" media-type="application/xhtml+xml" properties="nav"/>'
        )

        items_str = "\n    ".join(items_opf)
        spine_str = "\n    ".join(spine_refs)
        dur_str = "\n    ".join(duration_meta)
        content_opf = f'''<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">urn:uuid:{book_id}</dc:identifier>
    <dc:title>{html.escape(title)}</dc:title>
    <dc:language>zh-CN</dc:language>
    <dc:creator>edge-tts</dc:creator>
    <meta property="dcterms:modified">2026-01-01T00:00:00Z</meta>
    <meta property="media:duration">{_fmt_duration(total_dur)}</meta>
    {dur_str}
    <meta property="media:active-class">-epub-media-overlay-active</meta>
  </metadata>
  <manifest>
    {items_str}
  </manifest>
  <spine>
    {spine_str}
  </spine>
</package>'''
        zf.writestr("OEBPS/content.opf", content_opf)


# ── HTML 播放器构建 ──────────────────────────────────

def _map_sentences_to_blocks(
    text_blocks: list[dict], sentences: list[dict],
) -> list[list[dict]]:
    """将 TTS 返回的句子映射回各自所属的段落块"""
    full_text = "\n\n".join(b["text"] for b in text_blocks)
    # 每个 block 在 full_text 中的 [start, end)
    offset = 0
    ranges = []
    for b in text_blocks:
        s = offset
        e = offset + len(b["text"])
        ranges.append((s, e))
        offset = e + 2  # "\n\n"

    result: list[list[dict]] = [[] for _ in text_blocks]
    search_from = 0
    for sent in sentences:
        pos = full_text.find(sent["text"], search_from)
        if pos == -1:
            pos = full_text.find(sent["text"])
        if pos == -1:
            # 找不到 → 归到最后一个块
            result[-1].append(sent)
            continue
        search_from = pos + len(sent["text"])
        assigned = False
        for i, (bs, be) in enumerate(ranges):
            if pos >= bs and pos < be + 2:
                result[i].append(sent)
                assigned = True
                break
        if not assigned:
            result[-1].append(sent)
    return result


def build_html(chapters: list[dict], output_path: Path, title: str) -> None:
    """生成自包含 HTML 播放器：文字 + 内嵌音频 + 逐句高亮"""

    toc_items = []
    chapter_sections = []
    audio_elements = []
    sync_data: dict = {}
    sent_counter = 0  # 全局句子计数器

    for i, ch in enumerate(chapters):
        toc_items.append(
            f'<li><a href="#" data-ch="{i}">{html.escape(ch["title"])}</a></li>'
        )

        # 音频 base64
        audio_b64 = base64.b64encode(ch["audio_path"].read_bytes()).decode()
        audio_elements.append(
            f'<audio id="audio-{i}" preload="auto" '
            f'src="data:audio/mpeg;base64,{audio_b64}"></audio>'
        )

        # 按 block 渲染 HTML
        blocks = ch["blocks"]
        text_blocks = [b for b in blocks if b["type"] == "text"]
        block_sents = _map_sentences_to_blocks(text_blocks, ch["sentences"])

        body_parts = []
        text_idx = 0
        ch_sync = []

        for blk in blocks:
            if blk["type"] == "heading":
                lvl = min(blk["level"] + 1, 6)  # 章节标题已是 h2，子标题往下推
                body_parts.append(
                    f'<h{lvl}>{html.escape(blk["text"])}</h{lvl}>'
                )
            else:
                sents = block_sents[text_idx]
                text_idx += 1
                if not sents:
                    body_parts.append(
                        f'<p>{html.escape(blk["text"])}</p>'
                    )
                    continue
                spans = []
                for s in sents:
                    sid = f"s{i}_{sent_counter}"
                    spans.append(
                        f'<span id="{sid}" data-ch="{i}" '
                        f'data-start="{s["start"]:.3f}" '
                        f'data-end="{s["end"]:.3f}">'
                        f'{html.escape(s["text"])}</span>'
                    )
                    ch_sync.append({
                        "id": sid,
                        "start": round(s["start"], 3),
                        "end": round(s["end"], 3),
                    })
                    sent_counter += 1
                body_parts.append(f'<p>{"".join(spans)}</p>')

        sync_data[i] = ch_sync

        section = (
            f'<section id="ch{i}" class="chapter" data-ch="{i}">\n'
            f'<h2>{html.escape(ch["title"])}</h2>\n'
            + "\n".join(body_parts)
            + "\n</section>"
        )
        chapter_sections.append(section)

    toc_html = "\n".join(toc_items)
    chapters_html = "\n".join(chapter_sections)
    audios_html = "\n".join(audio_elements)
    sync_json = json.dumps(sync_data, ensure_ascii=False)
    n_ch = len(chapters)

    page = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html.escape(title)}</title>
<style>
:root {{
  --c-primary: #1a5276;
  --c-primary-light: #2980b9;
  --c-bg: #fafbfc;
  --c-surface: #fff;
  --c-border: #e8eaed;
  --c-text: #2c3e50;
  --c-text-secondary: #7f8c8d;
  --c-highlight: #fef9c3;
  --c-highlight-border: #facc15;
  --player-h: 72px;
  --sidebar-w: 260px;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
  font-family: -apple-system, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei",
               "Segoe UI", system-ui, sans-serif;
  background: var(--c-bg); color: var(--c-text); line-height: 1.9;
  font-size: 16px;
}}

/* ── 布局 ── */
#app {{ display: flex; height: 100vh; flex-direction: column; }}
#body-wrap {{ display: flex; flex: 1; overflow: hidden; }}

/* ── 侧边栏 ── */
#sidebar {{
  width: var(--sidebar-w); min-width: var(--sidebar-w);
  background: var(--c-surface); border-right: 1px solid var(--c-border);
  display: flex; flex-direction: column; overflow: hidden;
}}
#sidebar-header {{
  padding: 20px 20px 12px; border-bottom: 1px solid var(--c-border);
  position: relative;
}}
#sidebar-header h1 {{
  font-size: 15px; font-weight: 700; color: var(--c-primary);
  line-height: 1.4; padding-right: 32px;
}}
#sidebar-toggle {{
  position: absolute; top: 16px; right: 10px;
  background: none; border: none; cursor: pointer;
  font-size: 18px; color: var(--c-text-secondary); padding: 4px;
  border-radius: 4px; transition: background .15s;
}}
#sidebar-toggle:hover {{ background: #edf2f7; }}
.sidebar-hidden #sidebar {{ display: none; }}
#sidebar-show {{
  display: none; position: fixed; top: 12px; left: 12px; z-index: 50;
  background: var(--c-surface); border: 1px solid var(--c-border);
  border-radius: 8px; padding: 6px 10px; cursor: pointer;
  font-size: 18px; color: var(--c-text-secondary);
  box-shadow: 0 2px 8px rgba(0,0,0,.08);
}}
.sidebar-hidden #sidebar-show {{ display: block; }}
#toc {{
  list-style: none; flex: 1; overflow-y: auto; padding: 8px 10px;
}}
#toc li {{ margin-bottom: 2px; }}
#toc a {{
  display: block; padding: 8px 12px; border-radius: 8px;
  color: var(--c-text-secondary); text-decoration: none; font-size: 14px;
  line-height: 1.5; transition: all .15s;
}}
#toc a:hover {{ background: #edf2f7; color: var(--c-text); }}
#toc a.active {{
  background: var(--c-primary); color: #fff; font-weight: 500;
}}

/* ── 内容区 ── */
#content-wrap {{
  flex: 1; overflow-y: auto; padding-bottom: var(--player-h);
}}
#content {{
  max-width: 960px; margin: 0 auto;
  padding: 32px 40px 48px;
}}
.chapter {{ display: none; }}
.chapter.visible {{ display: block; }}
.chapter h2 {{
  font-size: 24px; font-weight: 700; color: var(--c-primary);
  margin-bottom: 24px; padding-bottom: 12px;
  border-bottom: 3px solid #d4e6f1;
}}
.chapter h3 {{
  font-size: 19px; font-weight: 600; color: #34495e;
  margin: 28px 0 12px; padding-left: 12px;
  border-left: 4px solid var(--c-primary-light);
}}
.chapter h4 {{
  font-size: 17px; font-weight: 600; color: #4a6274;
  margin: 20px 0 8px;
}}
.chapter h5, .chapter h6 {{
  font-size: 16px; font-weight: 600; color: #5a6c7d;
  margin: 16px 0 6px;
}}
.chapter p {{
  margin: 12px 0; text-indent: 2em;
  color: var(--c-text); line-height: 2;
}}
.chapter span[data-start] {{
  transition: background-color .25s, box-shadow .25s;
  border-radius: 3px; padding: 1px 0; cursor: pointer;
}}
.chapter span[data-start]:hover {{
  background-color: #f0f4f8;
}}
.chapter span.highlight {{
  background-color: var(--c-highlight);
  box-shadow: 0 0 0 2px var(--c-highlight-border);
}}

/* ── 底部播放栏 ── */
#player {{
  position: fixed; bottom: 0; left: 0; right: 0;
  height: var(--player-h); background: var(--c-surface);
  border-top: 1px solid var(--c-border);
  display: flex; align-items: center; padding: 0 24px; gap: 16px;
  box-shadow: 0 -2px 12px rgba(0,0,0,.06);
  z-index: 100;
}}

/* 播放按钮组 */
.ctrl-group {{ display: flex; align-items: center; gap: 6px; }}
.ctrl-btn {{
  width: 36px; height: 36px; border-radius: 50%; border: none;
  background: transparent; cursor: pointer; font-size: 16px;
  color: var(--c-text); display: flex; align-items: center;
  justify-content: center; transition: background .15s;
}}
.ctrl-btn:hover {{ background: #edf2f7; }}
#play-btn {{
  width: 44px; height: 44px; font-size: 20px;
  background: var(--c-primary); color: #fff; border-radius: 50%;
  border: none; cursor: pointer; display: flex;
  align-items: center; justify-content: center;
  transition: background .15s;
}}
#play-btn:hover {{ background: var(--c-primary-light); }}

/* 进度区 */
.progress-area {{ flex: 1; display: flex; flex-direction: column; gap: 4px; min-width: 0; }}
.progress-info {{
  display: flex; justify-content: space-between; align-items: center;
  font-size: 13px; color: var(--c-text-secondary);
}}
#ch-title {{
  font-weight: 500; color: var(--c-text);
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  max-width: 400px;
}}
#progress-wrap {{
  width: 100%; height: 6px; background: #e2e8f0;
  border-radius: 3px; cursor: pointer; position: relative;
  transition: height .15s;
}}
#progress-wrap:hover {{ height: 10px; }}
#progress-bar {{
  height: 100%; background: var(--c-primary); border-radius: 3px;
  width: 0; transition: width .15s linear;
  position: relative;
}}
#progress-bar::after {{
  content: ''; position: absolute; right: -6px; top: 50%;
  transform: translateY(-50%); width: 14px; height: 14px;
  background: var(--c-primary); border: 2px solid #fff;
  border-radius: 50%; opacity: 0; transition: opacity .15s;
  box-shadow: 0 1px 4px rgba(0,0,0,.2);
}}
#progress-wrap:hover #progress-bar::after {{ opacity: 1; }}

/* 倍速 */
.speed-group {{ display: flex; gap: 4px; }}
.speed-btn {{
  padding: 4px 10px; border-radius: 16px; border: 1px solid var(--c-border);
  background: transparent; cursor: pointer; font-size: 13px;
  color: var(--c-text-secondary); transition: all .15s; white-space: nowrap;
}}
.speed-btn:hover {{ border-color: var(--c-primary); color: var(--c-primary); }}
.speed-btn.active {{
  background: var(--c-primary); color: #fff;
  border-color: var(--c-primary);
}}

/* ── 响应式 ── */
@media (max-width: 768px) {{
  #sidebar {{ display: none; }}
  #content {{ padding: 20px 16px 40px; }}
  .speed-group {{ display: none; }}
  #ch-title {{ max-width: 200px; }}
}}
</style>
</head>
<body>
<div id="app">
  <div id="body-wrap">
    <button id="sidebar-show" title="显示目录">&#9776;</button>
    <nav id="sidebar">
      <div id="sidebar-header">
        <h1>{html.escape(title)}</h1>
        <button id="sidebar-toggle" title="收起目录">&#9776;</button>
      </div>
      <ul id="toc">{toc_html}</ul>
    </nav>
    <div id="content-wrap">
      <div id="content">{chapters_html}</div>
    </div>
  </div>

  <div id="player">
    <div class="ctrl-group">
      <button class="ctrl-btn" id="prev-btn" title="上一章">&#9198;</button>
      <button class="ctrl-btn" id="back-btn" title="-15s">&#8634;</button>
      <button id="play-btn" title="播放/暂停">&#9654;&#65039;</button>
      <button class="ctrl-btn" id="fwd-btn" title="+15s">&#8635;</button>
      <button class="ctrl-btn" id="next-btn" title="下一章">&#9197;</button>
    </div>
    <div class="progress-area">
      <div class="progress-info">
        <span id="ch-title"></span>
        <span id="time">00:00 / 00:00</span>
      </div>
      <div id="progress-wrap"><div id="progress-bar"></div></div>
    </div>
    <div class="speed-group">
      <button class="speed-btn" data-speed="0.75">0.75</button>
      <button class="speed-btn active" data-speed="1">1x</button>
      <button class="speed-btn" data-speed="1.25">1.25</button>
      <button class="speed-btn" data-speed="1.5">1.5</button>
      <button class="speed-btn" data-speed="2">2x</button>
    </div>
  </div>
</div>

{audios_html}

<script>
(function() {{
  const N = {n_ch};
  const SYNC = {sync_json};
  const TITLES = {json.dumps([ch["title"] for ch in chapters], ensure_ascii=False)};
  let cur = 0, playing = false;

  const $ = s => document.querySelector(s);
  const $$ = s => document.querySelectorAll(s);
  const audio = ch => document.getElementById('audio-' + ch);

  function fmt(s) {{
    if (!s || isNaN(s)) return '00:00';
    return String(Math.floor(s/60)).padStart(2,'0') + ':' + String(Math.floor(s%60)).padStart(2,'0');
  }}

  function go(idx, t) {{
    if (idx < 0 || idx >= N) return;
    const was = playing;
    const a = audio(cur);
    if (a) {{ a.pause(); a.currentTime = 0; }}
    cur = idx;
    $$('.chapter').forEach(s => s.classList.remove('visible'));
    const sec = document.getElementById('ch' + idx);
    if (sec) sec.classList.add('visible');
    $$('#toc a').forEach(a => a.classList.remove('active'));
    const link = document.querySelector('#toc a[data-ch="' + idx + '"]');
    if (link) {{ link.classList.add('active'); link.scrollIntoView({{block:'nearest'}}); }}
    $('#ch-title').textContent = TITLES[idx] || '';
    const na = audio(idx);
    if (na && typeof t === 'number') na.currentTime = t;
    if (was && na) na.play();
    clearHL();
    save();
  }}

  function clearHL() {{ $$('.highlight').forEach(s => s.classList.remove('highlight')); }}

  let lastHL = null;
  function highlight(a) {{
    const t = a.currentTime;
    const ss = SYNC[cur];
    if (!ss) return;
    for (const s of ss) {{
      if (t >= s.start && t < s.end) {{
        if (lastHL === s.id) return;
        clearHL();
        const el = document.getElementById(s.id);
        if (el) {{
          el.classList.add('highlight');
          el.scrollIntoView({{behavior:'smooth', block:'center'}});
        }}
        lastHL = s.id;
        return;
      }}
    }}
    // 句间空隙，不清除
  }}

  function onTime() {{
    const a = audio(cur);
    if (!a) return;
    highlight(a);
    const pct = a.duration ? (a.currentTime / a.duration * 100) : 0;
    $('#progress-bar').style.width = pct + '%';
    $('#time').textContent = fmt(a.currentTime) + ' / ' + fmt(a.duration);
  }}

  for (let i = 0; i < N; i++) {{
    const a = audio(i);
    if (!a) continue;
    a.addEventListener('timeupdate', onTime);
    a.addEventListener('ended', () => {{
      if (i + 1 < N) {{ go(i + 1, 0); audio(i+1).play(); }}
      else {{ playing = false; $('#play-btn').innerHTML = '&#9654;&#65039;'; }}
    }});
  }}

  $('#play-btn').addEventListener('click', () => {{
    const a = audio(cur);
    if (!a) return;
    if (playing) {{
      a.pause(); playing = false;
      $('#play-btn').innerHTML = '&#9654;&#65039;';
    }} else {{
      a.play(); playing = true;
      $('#play-btn').innerHTML = '&#9208;&#65039;';
    }}
  }});

  $('#progress-wrap').addEventListener('click', e => {{
    const a = audio(cur);
    if (!a || !a.duration) return;
    const r = $('#progress-wrap').getBoundingClientRect();
    a.currentTime = ((e.clientX - r.left) / r.width) * a.duration;
  }});

  $('#back-btn').addEventListener('click', () => {{
    const a = audio(cur); if (a) a.currentTime = Math.max(0, a.currentTime - 15);
  }});
  $('#fwd-btn').addEventListener('click', () => {{
    const a = audio(cur); if (a) a.currentTime = Math.min(a.duration||0, a.currentTime + 15);
  }});
  $('#prev-btn').addEventListener('click', () => go(cur - 1, 0));
  $('#next-btn').addEventListener('click', () => go(cur + 1, 0));

  $$('[data-speed]').forEach(btn => {{
    btn.addEventListener('click', () => {{
      const sp = parseFloat(btn.dataset.speed);
      for (let i = 0; i < N; i++) {{ const a = audio(i); if(a) a.playbackRate = sp; }}
      $$('[data-speed]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    }});
  }});

  $$('#toc a').forEach(a => {{
    a.addEventListener('click', e => {{ e.preventDefault(); go(parseInt(a.dataset.ch), 0); }});
  }});

  $('#content').addEventListener('click', e => {{
    const span = e.target.closest('span[data-start]');
    if (!span) return;
    const ch = parseInt(span.dataset.ch);
    const t = parseFloat(span.dataset.start);
    if (ch !== cur) go(ch, t);
    else {{ const a = audio(cur); if(a) a.currentTime = t; }}
    if (!playing) {{
      const a = audio(cur);
      if(a) {{ a.play(); playing = true; $('#play-btn').innerHTML = '&#9208;&#65039;'; }}
    }}
  }});

  const KEY = 'audiobook_' + {json.dumps(title)};
  function save() {{
    try {{
      const a = audio(cur);
      localStorage.setItem(KEY, JSON.stringify({{ch:cur, t: a ? a.currentTime : 0}}));
    }} catch(_) {{}}
  }}
  setInterval(save, 3000);

  function restore() {{
    try {{
      const s = JSON.parse(localStorage.getItem(KEY));
      if (s && typeof s.ch === 'number') {{ go(s.ch, s.t || 0); return; }}
    }} catch(_) {{}}
    go(0, 0);
  }}

  // 侧边栏 toggle
  const bodyWrap = $('#body-wrap');
  $('#sidebar-toggle').addEventListener('click', () => bodyWrap.classList.add('sidebar-hidden'));
  $('#sidebar-show').addEventListener('click', () => bodyWrap.classList.remove('sidebar-hidden'));

  document.addEventListener('keydown', e => {{
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    const a = audio(cur);
    if (e.code === 'Space') {{ e.preventDefault(); $('#play-btn').click(); }}
    else if (e.code === 'ArrowLeft') {{ if(a) a.currentTime = Math.max(0, a.currentTime - 5); }}
    else if (e.code === 'ArrowRight') {{ if(a) a.currentTime = Math.min(a.duration||0, a.currentTime + 5); }}
  }});

  restore();
}})();
</script>
</body>
</html>'''

    output_path.write_text(page, encoding="utf-8")


# ── 主流程 ──────────────────────────────────────────

async def main():
    parser = argparse.ArgumentParser(description="Markdown → 有声书（EPUB3/M4B）")
    parser.add_argument("input", help="Markdown 文件路径")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument("--format", choices=["html", "epub", "m4b"], default="html",
                        help="输出格式（默认 html）")
    parser.add_argument("--voice", default=DEFAULT_VOICE,
                        help=f"音色（默认 {DEFAULT_VOICE}）")
    parser.add_argument("--heading-level", type=int, default=2,
                        help="章节拆分标题级别（默认 2 = ##）")
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"错误：文件不存在 {input_path}")
        return

    suffix = {"html": ".html", "epub": ".epub", "m4b": ".m4b"}[args.format]
    output_path = (
        Path(args.output).resolve() if args.output
        else input_path.with_suffix(suffix)
    )

    md_text = input_path.read_text(encoding="utf-8")
    book_title_m = re.match(r"^#\s+(.+)$", md_text, re.MULTILINE)
    book_title = book_title_m.group(1).strip() if book_title_m else input_path.stem

    chapters = split_chapters(md_text, args.heading_level)
    print(f"📖 {book_title}")
    print(f"共 {len(chapters)} 章：")
    for i, ch in enumerate(chapters, 1):
        print(f"  {i}. {ch['title']}")
    print()

    if args.format == "html":
        await _build_html_flow(chapters, output_path, book_title, args.voice)
    elif args.format == "epub":
        await _build_epub_flow(chapters, output_path, book_title, args.voice)
    else:
        await _build_m4b_flow(chapters, output_path, book_title, args.voice)

    size_mb = output_path.stat().st_size / 1024 / 1024
    print(f"\n✅ {output_path}")
    print(f"   大小: {size_mb:.1f} MB")


async def _build_epub_flow(chapters, output_path, title, voice):
    """EPUB3 流程：逐章生成音频 + 句子时间戳 → 打包"""
    enriched = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        for i, ch in enumerate(chapters, 1):
            cleaned = clean_text(ch["body"])
            if not cleaned:
                print(f"  [{i}/{len(chapters)}] {ch['title']} — 跳过")
                continue

            mp3_path = tmpdir / f"chapter_{i:03d}.mp3"
            print(f"  [{i}/{len(chapters)}] {ch['title']} ({len(cleaned)} 字) ...", end="", flush=True)

            sentences, duration = await generate_audio_with_sync(cleaned, mp3_path, voice)
            print(f" {duration:.0f}s ({len(sentences)} 句)")

            enriched.append({
                "title": ch["title"],
                "sentences": sentences,
                "audio_path": mp3_path,
                "duration": duration,
            })

        if not enriched:
            print("错误：没有可合成的内容")
            return

        total = sum(c["duration"] for c in enriched)
        print(f"\n打包 EPUB3 ... 总时长 {int(total//60)}分{int(total%60)}秒")
        build_epub3(enriched, output_path, title)


async def _build_html_flow(chapters, output_path, title, voice):
    """HTML 播放器流程：解析 MD 结构 → 逐章生成音频 + 时间戳 → 打包 HTML"""
    enriched = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        for i, ch in enumerate(chapters, 1):
            blocks = parse_chapter_blocks(ch["body"])
            text_blocks = [b for b in blocks if b["type"] == "text"]
            tts_text = "\n\n".join(b["text"] for b in text_blocks)

            if not tts_text.strip():
                print(f"  [{i}/{len(chapters)}] {ch['title']} — 跳过")
                continue

            mp3_path = tmpdir / f"chapter_{i:03d}.mp3"
            print(f"  [{i}/{len(chapters)}] {ch['title']} ({len(tts_text)} 字) ...", end="", flush=True)

            sentences, duration = await generate_audio_with_sync(tts_text, mp3_path, voice)
            print(f" {duration:.0f}s ({len(sentences)} 句)")

            enriched.append({
                "title": ch["title"],
                "blocks": blocks,
                "sentences": sentences,
                "audio_path": mp3_path,
                "duration": duration,
            })

        if not enriched:
            print("错误：没有可合成的内容")
            return

        total = sum(c["duration"] for c in enriched)
        print(f"\n生成 HTML 播放器 ... 总时长 {int(total//60)}分{int(total%60)}秒")
        build_html(enriched, output_path, title)


async def _build_m4b_flow(chapters, output_path, title, voice):
    """M4B 流程：逐章生成音频 → 合并"""
    chapter_files = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        for i, ch in enumerate(chapters, 1):
            cleaned = clean_text(ch["body"])
            if not cleaned:
                print(f"  [{i}/{len(chapters)}] {ch['title']} — 跳过")
                continue

            mp3_path = tmpdir / f"chapter_{i:03d}.mp3"
            print(f"  [{i}/{len(chapters)}] {ch['title']} ({len(cleaned)} 字) ...", end="", flush=True)
            duration = await generate_audio_simple(cleaned, mp3_path, voice)
            print(f" {duration:.0f}s")

            chapter_files.append({"title": ch["title"], "path": mp3_path, "duration": duration})

        if not chapter_files:
            print("错误：没有可合成的内容")
            return

        total = sum(c["duration"] for c in chapter_files)
        print(f"\n合并 M4B ... 总时长 {int(total//60)}分{int(total%60)}秒")
        build_m4b(chapter_files, output_path, title)


if __name__ == "__main__":
    asyncio.run(main())
