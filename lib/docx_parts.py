#!/usr/bin/env python3
"""docx_parts — surgical 改写的部件完整性校验（zipfile 路线专用）。

## 为什么有这个（2026-07-31 立，两次真实事故）

`docx_safe_save` 只收口 **python-docx** 的存盘路径。但 surgical 改法（`zipfile` +
`lxml` 手工重打包）根本不走 python-docx，**它一个字节都管不到**。于是这条路上丢部件
无人可挡，实测两次：

    成果2 修改稿   162 部件 → 74    11 个原生图表 + 40 header + 18 footer 全丢
    成果3 修改稿   137 部件 → 35    同上，外加 2 个 themeOverride

第二次是我自己的脚本干的：对同一个文件开了**两个** `ZipFile` 句柄，一个读
`comments.xml`、一个在 `for item in z.infolist()` 里逐部件复制，复制被截断，
只写出 35/137 个部件就收工了。**文件照样能打开，Word 也不报错** —— 图表和页眉
就是没了，肉眼翻到那一页才发现。

比丢部件更隐蔽的是**验证时比错基线**：当时我报「137→137 零丢失」，
因为拿 `修改稿 ↔ 修改稿` 比。基线必须是**未被改动的源件**，这是本模块
`assert_parts_intact(src, dst)` 强制两个入参的原因 —— 让「拿什么当基线」
成为显式决定，而不是随手取一个手边的文件。

## 用法（surgical 脚本存盘后加一行）

    from docx_parts import assert_parts_intact
    assert_parts_intact(src_path, dst_path, allow_added={'word/comments.xml'})

丢部件 / 非白名单内的部件字节被改 → 抛 `PartIntegrityError`，不静默。

只想看不想抛：

    report = diff_parts(src, dst)      # -> PartDiff
    print(report.summary())

## 设计取舍

- **默认最严**：除 `word/document.xml` 外任何部件字节变化都算异常。
  surgical 的定义就是「只重写点名的部件」，多改一个就该解释清楚。
  确实要改多个部件（如同时改 comments.xml）就显式列进 `allow_changed`。
- **fail-closed**：源件读不到 / 枚举为空 → 抛，不返回「看起来没问题」。
"""
from __future__ import annotations

import hashlib
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

__all__ = ['PartIntegrityError', 'PartDiff', 'diff_parts', 'assert_parts_intact',
           'media_census', 'assert_media_intact']

# surgical 改写中「本来就会变」的部件：正文永远要改；
# 改批注/新增部件时会连带动 rels 与 content-types 注册表。
DEFAULT_ALLOW_CHANGED = frozenset({
    'word/document.xml',
    'word/comments.xml',
    'word/_rels/document.xml.rels',
    '[Content_Types].xml',
})


class PartIntegrityError(RuntimeError):
    """部件丢失，或白名单外的部件被改写。"""


@dataclass
class PartDiff:
    src: Path
    dst: Path
    lost: list[str] = field(default_factory=list)
    added: list[str] = field(default_factory=list)
    changed: list[str] = field(default_factory=list)
    unchanged: int = 0

    @property
    def ok(self) -> bool:
        return not self.lost and not self.changed

    def summary(self) -> str:
        head = f'{self.src.name} → {self.dst.name}'
        bits = [f'保持原样 {self.unchanged}']
        if self.changed:
            bits.append(f'改写 {len(self.changed)}')
        if self.added:
            bits.append(f'新增 {len(self.added)}')
        if self.lost:
            bits.append(f'**丢失 {len(self.lost)}**')
        out = [f'{head}：' + ' / '.join(bits)]
        if self.lost:
            out.append('  丢失：' + _brief(self.lost))
        if self.changed:
            out.append('  改写：' + _brief(self.changed))
        if self.added:
            out.append('  新增：' + _brief(self.added))
        return '\n'.join(out)


def _brief(names: list[str], n: int = 6) -> str:
    shown = ', '.join(sorted(names)[:n])
    return shown + (f' …共 {len(names)} 个' if len(names) > n else '')


def _digests(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as z:
        # zip 目录条目（`word/`、`_rels/` 这类以 / 结尾的零长条目）不是 OOXML 部件。
        # WPS 写出的 docx 常带这类占位条目，python 侧重打包时不会复现——
        # 若当部件比对，会把「目录占位没了」误报成「丢失 9 个部件」，属假阳性。
        names = [n for n in z.namelist() if not n.endswith('/')]
        if not names:
            raise PartIntegrityError(f'{path} 里一个部件都没有——拒绝在空集上报绿')
        return {n: hashlib.sha256(z.read(n)).hexdigest() for n in names}


# ── 图片守卫 ─────────────────────────────────────────────────────────────
# 部件集合比对看不见的一种丢失：media 部件与 rels 都还在，正文里的 w:drawing 没了
# （python-docx 删段/重建后就是这个形状——zip 条目一个没少，图却不在文档里）。
# 所以另量一把尺：图片部件数 + 正文类部件里的图引用数 + 悬空的图片关系。
_STORY_RE = re.compile(
    r'^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$')
_REF_RE = re.compile(rb'<w:(?:drawing|pict|object)[\s>/]')
_RID_RE = re.compile(rb'<(?:a:blip|v:imagedata|asvg:svgBlip)\b[^>]*?'
                     rb'\br:(?:embed|link|id)="([^"]+)"')
_REL_RE = re.compile(rb'<Relationship\b[^>]*>')
_ATTR_RE = re.compile(rb'\b(Id|Target|TargetMode)="([^"]*)"')


def _rels_of(z: zipfile.ZipFile, part: str, names: set) -> dict:
    d, b = posixpath.split(part)
    rp = f'{d}/_rels/{b}.rels'
    if rp not in names:
        return {}
    out = {}
    for m in _REL_RE.finditer(z.read(rp)):
        a = {k.decode(): v.decode('utf-8', 'replace')
             for k, v in _ATTR_RE.findall(m.group(0))}
        if 'Id' in a:
            out[a['Id']] = (a.get('Target', ''), a.get('TargetMode', ''))
    return out


def media_census(path) -> dict:
    """数一份 docx 里的图：`media` 图片部件数 / `refs` 正文类部件里的图引用数 /
    `dangling` 指不到部件的图片关系（形如 ``word/document.xml#rId7``）。

    纯 stdlib（zipfile + 正则），不依赖 python-docx。读不了 → 抛，不返回全 0 装没事。
    """
    path = Path(path)
    try:
        z = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        raise PartIntegrityError(f'图片普查失败：{path} 读不了（{e}）') from e
    with z:
        names = {n for n in z.namelist() if not n.endswith('/')}
        if 'word/document.xml' not in names:
            raise PartIntegrityError(f'图片普查失败：{path} 里没有 word/document.xml')
        media = sum(1 for n in names if n.startswith('word/media/'))
        refs, dangling = 0, []
        for part in sorted(n for n in names if _STORY_RE.match(n)):
            xml = z.read(part)
            refs += len(_REF_RE.findall(xml))
            rids = {r.decode() for r in _RID_RE.findall(xml)}
            if not rids:
                continue
            rels = _rels_of(z, part, names)
            for rid in sorted(rids):
                if rid not in rels:
                    dangling.append(f'{part}#{rid}')
                    continue
                target, mode = rels[rid]
                if mode == 'External':
                    continue                       # 链接图：目标本来就不在包里
                full = (target.lstrip('/') if target.startswith('/') else
                        posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
                if full not in names:
                    dangling.append(f'{part}#{rid}')
    return {'media': media, 'refs': refs, 'dangling': dangling}


def assert_media_intact(before: dict, after: dict, *, allow_loss: bool = False,
                        label: str = '') -> None:
    """改后的图不许比改前少。`before`/`after` 是 `media_census()` 的返回。

    - 新出现的悬空图片关系 → **总是**抛（Word 里就是一个红叉，没有「有意为之」）
    - `media` 或 `refs` 变少 → 抛；确属有意减图传 `allow_loss=True`
    - 变多不算错。
    """
    bad = []
    new_dangling = sorted(set(after['dangling']) - set(before['dangling']))
    if new_dangling:
        bad.append(f'新增 {len(new_dangling)} 处悬空图片引用：{_brief(new_dangling)}')
    if not allow_loss:
        if after['media'] < before['media']:
            bad.append(f"图片部件 {before['media']} → {after['media']}")
        if after['refs'] < before['refs']:
            bad.append(f"正文图引用 {before['refs']} → {after['refs']}")
    if bad:
        raise PartIntegrityError(
            '\n'.join([f'图片守卫未通过{("（" + label + "）") if label else ""}：']
                      + [f'  · {x}' for x in bad]
                      + ['  图变少通常意味着「重建文档」而不是「改文档」——内嵌图不会跟着文字搬过去。',
                         '  确属有意减图 → python-docx 存盘包 `with docx_safe_save.allow_media_loss():`，'
                         'surgical 校验传 `allow_media_loss=True`。']))


def diff_parts(src, dst, allow_changed=DEFAULT_ALLOW_CHANGED) -> PartDiff:
    """比对源件与产物的部件集合与逐部件字节。src 必须是**未被改动的源件**。"""
    src, dst = Path(src), Path(dst)
    for p in (src, dst):
        if not p.exists():
            raise PartIntegrityError(f'比对失败：{p} 不存在')
    if src.resolve() == dst.resolve():
        raise PartIntegrityError(
            f'源件与产物是同一个文件（{src}）——这样比等于没比。'
            f'原地改写时请先把改前的副本留出来当基线。')

    a, b = _digests(src), _digests(dst)
    d = PartDiff(src=src, dst=dst)
    d.lost = sorted(set(a) - set(b))
    d.added = sorted(set(b) - set(a))
    for n in sorted(set(a) & set(b)):
        if a[n] == b[n]:
            d.unchanged += 1
        elif n not in allow_changed:
            d.changed.append(n)
    return d


def assert_parts_intact(src, dst, allow_changed=DEFAULT_ALLOW_CHANGED,
                        allow_added=frozenset(), verbose: bool = True,
                        allow_media_loss: bool | None = None) -> PartDiff:
    """surgical 存盘后调它。丢部件或白名单外部件被改 → 抛 PartIntegrityError。

    allow_added: 本次有意新增的部件（如首次加批注的 word/comments.xml）。
                 新增不算错，但**必须报备**，否则一样抛 —— 防的是
                 「悄悄多塞了个部件进去」。
    allow_media_loss: 正文图引用（w:drawing/w:pict/w:object）变少怎么办。
                 None（默认）= 打印告警不抛；False = 抛；True = 有意减图，不吭声。
                 新增悬空图片关系与图片部件丢失不看这个参数，一律抛。
                 默认先不抛的原因：本函数在存盘**之后**才被调用（拦不住覆盖），且有
                 按章裁剪 document.xml 的调用方，减图对它们是正常结果。
    """
    allow_changed = set(allow_changed) | set(allow_added)
    d = diff_parts(src, dst, allow_changed)
    src, dst = d.src, d.dst   # diff_parts 已归一成 Path，后面报错要用 .name
    bad = []
    if d.lost:
        bad.append(f'丢失 {len(d.lost)} 个部件：{_brief(d.lost)}')
    if d.changed:
        bad.append(f'{len(d.changed)} 个未报备的部件被改写：{_brief(d.changed)}')
    unexpected = [n for n in d.added if n not in allow_added]
    if unexpected:
        bad.append(f'{len(unexpected)} 个未报备的新增部件：{_brief(unexpected)}')
    before, after = media_census(src), media_census(dst)
    try:
        assert_media_intact(before, after, allow_loss=allow_media_loss is not False,
                            label=f'{src.name} → {dst.name}')
    except PartIntegrityError as e:
        bad.append(str(e))
    if allow_media_loss is None and after['refs'] < before['refs']:
        print(f"  ⚠ 正文图引用 {before['refs']} → {after['refs']}（{src.name} → {dst.name}）"
              f"——有意减图请传 allow_media_loss=True", file=sys.stderr)
    if bad:
        raise PartIntegrityError(
            '\n'.join([f'surgical 完整性校验未通过（{src.name} → {dst.name}）：']
                      + [f'  · {x}' for x in bad]
                      + ['  修法：逐部件复制时复用同一个 ZipFile 句柄，'
                         '别对同一文件开第二个句柄读别的部件；',
                         '  确属有意改动 → 显式列进 allow_changed / allow_added。']))
    if verbose:
        print(f'  ✓ 部件完整性：{d.summary()}')
    return d


if __name__ == '__main__':
    import argparse

    ap = argparse.ArgumentParser(description='比对两个 docx 的部件完整性')
    ap.add_argument('src', help='未被改动的源件（基线）')
    ap.add_argument('dst', help='surgical 改写后的产物')
    ap.add_argument('--strict', action='store_true', help='有问题则 exit 1')
    a = ap.parse_args()
    diff = diff_parts(a.src, a.dst)
    print(diff.summary())
    if a.strict and not diff.ok:
        raise SystemExit(1)
