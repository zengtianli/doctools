"""Word 插图（流程图、框架图、组织架构图）统一样式：坐标单位 cm，按版宽 15.5cm 实尺绘制，300dpi 输出。
字号为印刷实际磅值（正文不小于 8.5pt）。配色为蓝色系，ACCENT 橙色只用于公式和关键指标。

字体（2026-09-18 海宁标实测，macOS）：
- matplotlib 里的 "Songti SC" 只注册到 Black(900) 字重，画出来像粗黑体，不能当宋体正文用；
  正文用 SimSong（常规宋体），标题用 STHeiti（华文黑体）。
- "PingFang SC" 在本机 matplotlib 字体表里没有（只有 PingFang HK），写了会静默回退。
用法：sys.path 加入本目录后 `from figstyle import *`；工程线条元件见同目录 figeng.py。
"""
import math
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle, Polygon

HEI = 'STHeiti'          # 标题（华文黑体）
SONG = 'SimSong'         # 正文（常规宋体）
plt.rcParams['font.sans-serif'] = [HEI, SONG]
plt.rcParams['axes.unicode_minus'] = False

DARK = '#2F5597'    # 深蓝：层级标题、主框表头
MID = '#4A74CC'     # 中蓝：阶段条、箭头
BAR = '#6B8FD6'     # 浅阶段条（同技术路线图）
LIGHT = '#EAF0FA'   # 框体填充
LINE = '#8FAADC'    # 框线
ACCENT = '#C55A11'  # 强调（公式、结果），少用
ACC_BG = '#FBEEE6'
INK = '#1F1F1F'
GRAY = '#595959'
PT = 0.03528        # 1pt = 0.03528cm
W = 15.5            # 版宽 cm


def new(h, w=W):
    fig = plt.figure(figsize=(w / 2.54, h / 2.54), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, w); ax.set_ylim(0, h); ax.axis('off')
    return fig, ax


def save(fig, path, margin_cm=0.15):
    """保存并裁去上下多余白边（左右保持版宽，便于 Word 中统一按 15.5cm 排版）。"""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=300, facecolor='white')
    plt.close(fig)
    from PIL import Image, ImageChops
    im = Image.open(path).convert('RGB')
    bbox = ImageChops.difference(im, Image.new('RGB', im.size, 'white')).getbbox()
    if bbox:
        m = int(margin_cm / 2.54 * 300)
        im.crop((0, max(0, bbox[1] - m), im.size[0], min(im.size[1], bbox[3] + m))).save(path, dpi=(300, 300))


def wrap(text, width_cm, size):
    """按字宽折行：汉字宽≈字号，ASCII≈0.55字号。保留显式换行。"""
    out = []
    for para in str(text).split('\n'):
        line, lw = '', 0.0
        for ch in para:
            cw = size * PT * (1.0 if ord(ch) > 0x2E80 else 0.55)
            if lw + cw > width_cm and line:
                out.append(line); line, lw = '', 0.0
            line += ch; lw += cw
        out.append(line)
    return out


def text(ax, x, y, s, size=9, color=INK, font=SONG, weight='normal', ha='center', va='center', width=None, lh=1.35):
    """在 (x,y) 写多行文字；width 给定时自动折行。返回文字块高度(cm)。"""
    lines = wrap(s, width, size) if width else str(s).split('\n')
    step = size * PT * lh
    total = step * len(lines)
    if va == 'center':
        y0 = y + total / 2 - step / 2
    elif va == 'top':
        y0 = y - step / 2
    else:
        y0 = y + total - step / 2
    for i, ln in enumerate(lines):
        ax.text(x, y0 - i * step, ln, fontsize=size, color=color, family=font, weight=weight, ha=ha, va='center')
    return total


def text_h(s, width, size=9, lh=1.35):
    return len(wrap(s, width, size)) * size * PT * lh


def _body_lines(body, width, size):
    """返回 [(行文字, 缩进cm)]；列表项加“· ”并悬挂缩进。"""
    items = body if isinstance(body, list) else [body]
    bullet = isinstance(body, list) and len(items) > 1
    out = []
    for it in items:
        if bullet:
            ind = size * PT * 1.0
            ls = wrap(it, width - ind, size)
            out.append(('· ' + ls[0], 0.0))
            out += [(l, ind) for l in ls[1:]]
        else:
            out += [(l, 0.0) for l in wrap(it, width, size)]
        out.append((None, 0.0))       # 项间距标记
    return out[:-1]


def box_height(w, title=None, body=None, title_size=9.5, body_size=8.5, lh=1.4, pad=0.18):
    h = 0.0
    if title:
        h += max(0.62, text_h(title, w - 0.3, title_size) + 0.22)
    if body:
        n = sum(1 for l, _ in _body_lines(body, w - 2 * pad, body_size) if l is not None)
        gaps = sum(1 for l, _ in _body_lines(body, w - 2 * pad, body_size) if l is None)
        h += n * body_size * PT * lh + gaps * 0.06 + 0.32
    return h


def box(ax, x, y=None, w=None, h=None, title=None, body=None, title_size=9.5, body_size=8.5, head=DARK, fill=LIGHT,
        edge=LINE, head_h=None, body_align='left', round_=0.08, top=None, lh=1.4):
    """带表头的框。给 (x, y=左下角, h) 或 (x, top=顶边, h=None 自动高)。返回 (x, 底y, w, h)。
    body：字符串或列表（列表项自动加“·”并悬挂缩进）。"""
    pad = 0.18
    if h is None:
        h = box_height(w, title, body, title_size, body_size, lh, pad)
    if top is not None:
        y = top - h
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f'round,pad=0,rounding_size={round_}',
                                fc=fill, ec=edge, lw=0.8))
    cur_top = y + h
    if title:
        hh = head_h or max(0.62, text_h(title, w - 0.3, title_size) + 0.22)
        ax.add_patch(FancyBboxPatch((x, cur_top - hh), w, hh, boxstyle=f'round,pad=0,rounding_size={round_}',
                                    fc=head, ec=head, lw=0.8))
        ax.add_patch(Rectangle((x, cur_top - hh), w, hh / 2, fc=head, ec=head, lw=0))
        text(ax, x + w / 2, cur_top - hh / 2, title, size=title_size, color='white', font=HEI, weight='bold',
             width=w - 0.3)
        cur_top -= hh
    if body:
        lines = _body_lines(body, w - 2 * pad, body_size)
        step = body_size * PT * lh
        n = sum(1 for l, _ in lines if l is not None); g = sum(1 for l, _ in lines if l is None)
        blk = n * step + g * 0.06
        avail = cur_top - y
        cy = cur_top - max(0.16, (avail - blk) / 2) - step / 2 if body_align == 'center' else cur_top - 0.16 - step / 2
        for l, ind in lines:
            if l is None:
                cy -= 0.06; continue
            if body_align == 'center':
                ax.text(x + w / 2, cy, l, fontsize=body_size, color=INK, family=SONG, ha='center', va='center')
            else:
                ax.text(x + pad + ind, cy, l, fontsize=body_size, color=INK, family=SONG, ha='left', va='center')
            cy -= step
    return (x, y, w, h)


def pill(ax, x, y, w, h, s, size=9, fc=MID, color='white', font=HEI, weight='bold'):
    """实心圆角条（阶段名、层级名）。"""
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0,rounding_size=0.08', fc=fc, ec=fc, lw=0.8))
    text(ax, x + w / 2, y + h / 2, s, size=size, color=color, font=font, weight=weight, width=w - 0.2)


def formula(ax, x, y, w, h, s, size=9):
    """公式框：浅橙底、橙边，居中。"""
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0,rounding_size=0.08', fc=ACC_BG, ec=ACCENT, lw=0.9))
    text(ax, x + w / 2, y + h / 2, s, size=size, color=ACCENT, font=HEI, weight='bold', width=w - 0.3)


def arrow(ax, p1, p2, color=MID, lw=1.2, style='-|>', ls='-', label=None, label_size=8, label_off=(0, 0.18), rad=0.0):
    a = FancyArrowPatch(p1, p2, arrowstyle=style, mutation_scale=9, color=color, lw=lw, linestyle=ls,
                        connectionstyle=f'arc3,rad={rad}', shrinkA=0, shrinkB=0)
    ax.add_patch(a)
    if label:
        mx, my = (p1[0] + p2[0]) / 2 + label_off[0], (p1[1] + p2[1]) / 2 + label_off[1]
        text(ax, mx, my, label, size=label_size, color=GRAY)


def poly_arrow(ax, pts, color=MID, lw=1.2, ls='-'):
    """折线箭头：最后一段带箭头。"""
    for a, b in zip(pts[:-2], pts[1:-1]):
        ax.plot([a[0], b[0]], [a[1], b[1]], color=color, lw=lw, ls=ls, solid_capstyle='butt')
    arrow(ax, pts[-2], pts[-1], color=color, lw=lw, ls=ls)


def title_bar(ax, y, s, h=0.6, size=10.5, x=0.2, w=None):
    w = w or (W - 0.4)
    ax.add_patch(Rectangle((x, y), w, h, fc=BAR, ec=BAR))
    text(ax, x + w / 2, y + h / 2, s, size=size, color=INK, font=HEI, weight='bold')


def note(ax, x, y, s, size=8, width=None, ha='left', color=GRAY):
    return text(ax, x, y, s, size=size, color=color, ha=ha, va='top', width=width)


def row_height(w, items, title_size=9.5, body_size=8.5):
    """同一行若干框取统一高度：items 为 [(title, body), ...]。"""
    return max(box_height(w, t, b, title_size, body_size) for t, b in items)
