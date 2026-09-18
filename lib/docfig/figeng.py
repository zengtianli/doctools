"""工程示意图（管道、流量计、阀门、构筑物、尺寸线）的简化线条元件，配合 figstyle 使用，坐标单位 cm。"""
import numpy as np
from matplotlib.patches import Circle, Wedge, Arc, Rectangle, FancyBboxPatch, Polygon
from figstyle import *

PIPE_FILL = '#F3F6FC'
SONGX = ['Times New Roman', SONG]   # 宋体正文中的数字、字母用 Times New Roman（与 Word 正文一致）
OUT = DARK          # 设备轮廓线
LW = 0.9


def hpipe(ax, x1, x2, yc, r, fill=PIPE_FILL, lw=LW):
    ax.add_patch(Rectangle((x1, yc - r), x2 - x1, 2 * r, fc=fill, ec='none', zorder=2))
    ax.plot([x1, x2], [yc + r, yc + r], color=OUT, lw=lw, zorder=3)
    ax.plot([x1, x2], [yc - r, yc - r], color=OUT, lw=lw, zorder=3)


def vpipe(ax, xc, y1, y2, r, fill=PIPE_FILL, lw=LW):
    ax.add_patch(Rectangle((xc - r, y1), 2 * r, y2 - y1, fc=fill, ec='none', zorder=2))
    ax.plot([xc - r, xc - r], [y1, y2], color=OUT, lw=lw, zorder=3)
    ax.plot([xc + r, xc + r], [y1, y2], color=OUT, lw=lw, zorder=3)


def flange_v(ax, x, yc, r, t=0.09, ext=0.12):
    """竖向法兰（水平管上）。"""
    ax.add_patch(Rectangle((x - t / 2, yc - r - ext), t, 2 * (r + ext), fc=MID, ec=OUT, lw=0.6, zorder=5))


def flange_h(ax, xc, y, r, t=0.09, ext=0.12):
    ax.add_patch(Rectangle((xc - r - ext, y - t / 2), 2 * (r + ext), t, fc=MID, ec=OUT, lw=0.6, zorder=5))


def meter_h(ax, x1, x2, yc, r, head=True, head_w=0.62, head_h=0.42, body_ext=0.16):
    """水平管上的电磁流量计（侧视）：两端法兰 + 测量管本体 + 顶部转换器表头。返回表头顶 y。"""
    ax.add_patch(FancyBboxPatch((x1 + 0.05, yc - r - body_ext), x2 - x1 - 0.1, 2 * (r + body_ext),
                                boxstyle='round,pad=0,rounding_size=0.08', fc='white', ec=OUT, lw=LW, zorder=4))
    # 线圈/电极位置示意（中线）
    ax.plot([x1 + 0.2, x2 - 0.2], [yc, yc], color=LINE, lw=0.6, ls=(0, (3, 2)), zorder=5)
    ax.add_patch(Circle(((x1 + x2) / 2, yc), 0.06, fc=DARK, ec='none', zorder=6))
    flange_v(ax, x1, yc, r); flange_v(ax, x2, yc, r)
    top = yc + r + body_ext
    if head:
        cx = (x1 + x2) / 2
        ax.add_patch(Rectangle((cx - 0.07, top), 0.14, 0.12, fc=LIGHT, ec=OUT, lw=0.6, zorder=4))
        ax.add_patch(FancyBboxPatch((cx - head_w / 2, top + 0.12), head_w, head_h,
                                    boxstyle='round,pad=0,rounding_size=0.05', fc=LIGHT, ec=OUT, lw=LW, zorder=4))
        ax.add_patch(Rectangle((cx - head_w / 2 + 0.1, top + 0.12 + 0.1), head_w - 0.2, head_h - 0.2,
                               fc='white', ec=LINE, lw=0.5, zorder=5))
        top += 0.12 + head_h
    return top


def meter_v(ax, xc, y1, y2, r, body_ext=0.16, head=True, head_w=0.42, head_h=0.62):
    """竖直管上的电磁流量计，表头在右侧。"""
    ax.add_patch(FancyBboxPatch((xc - r - body_ext, y1 + 0.05), 2 * (r + body_ext), y2 - y1 - 0.1,
                                boxstyle='round,pad=0,rounding_size=0.08', fc='white', ec=OUT, lw=LW, zorder=4))
    ax.add_patch(Circle((xc, (y1 + y2) / 2), 0.06, fc=DARK, ec='none', zorder=6))
    flange_h(ax, xc, y1, r); flange_h(ax, xc, y2, r)
    if head:
        cy = (y1 + y2) / 2; right = xc + r + body_ext
        ax.add_patch(Rectangle((right, cy - 0.07), 0.12, 0.14, fc=LIGHT, ec=OUT, lw=0.6, zorder=4))
        ax.add_patch(FancyBboxPatch((right + 0.12, cy - head_h / 2), head_w, head_h,
                                    boxstyle='round,pad=0,rounding_size=0.05', fc=LIGHT, ec=OUT, lw=LW, zorder=4))


def gate_valve_h(ax, xc, yc, r, w=0.56, wheel=True):
    """水平管上的阀门（蝶形符号 + 阀杆手轮）。"""
    hh = r + 0.12
    ax.add_patch(Polygon([(xc - w / 2, yc + hh), (xc, yc), (xc - w / 2, yc - hh)], closed=True,
                         fc='white', ec=OUT, lw=LW, zorder=5))
    ax.add_patch(Polygon([(xc + w / 2, yc + hh), (xc, yc), (xc + w / 2, yc - hh)], closed=True,
                         fc='white', ec=OUT, lw=LW, zorder=5))
    flange_v(ax, xc - w / 2, yc, r); flange_v(ax, xc + w / 2, yc, r)
    top = yc
    if wheel:
        ax.plot([xc, xc], [yc, yc + hh + 0.38], color=OUT, lw=LW, zorder=4)
        ax.plot([xc - 0.24, xc + 0.24], [yc + hh + 0.38, yc + hh + 0.38], color=OUT, lw=1.6, zorder=4)
        top = yc + hh + 0.38
    return top


def elbow(ax, cx, cy, R, r, a1, a2, fill=PIPE_FILL):
    """圆弧弯头：中心 (cx,cy)，中线半径 R，管半径 r，角度 a1→a2（度）。"""
    ax.add_patch(Wedge((cx, cy), R + r, a1, a2, width=2 * r, fc=fill, ec='none', zorder=2))
    ax.add_patch(Arc((cx, cy), 2 * (R + r), 2 * (R + r), theta1=a1, theta2=a2, color=OUT, lw=LW, zorder=3))
    ax.add_patch(Arc((cx, cy), 2 * (R - r), 2 * (R - r), theta1=a1, theta2=a2, color=OUT, lw=LW, zorder=3))


def flow(ax, x1, y1, x2, y2, lw=1.4, color=MID, zorder=8):
    """流向箭头（可画在管内，位于管道填充之上）。"""
    from matplotlib.patches import FancyArrowPatch
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle='-|>', mutation_scale=9, color=color, lw=lw,
                                 shrinkA=0, shrinkB=0, zorder=zorder))


def dim_h(ax, x1, x2, y, label=None, size=8.5, color=GRAY, ext_from=None, label_dy=0.25, font=HEI, lcolor=None):
    """水平尺寸线：两端箭头 + 竖向引出线（从 ext_from 起）。"""
    if ext_from is not None:
        for x in (x1, x2):
            ax.plot([x, x], [ext_from, y + 0.05 * np.sign(y - ext_from)], color=color, lw=0.5, zorder=3)
    arrow(ax, (x1, y), (x2, y), color=color, lw=0.6, style='<|-|>')
    if label:
        text(ax, (x1 + x2) / 2, y + label_dy, label, size=size, color=lcolor or color, font=font)


def dim_v(ax, x, y1, y2, label=None, size=8.5, color=GRAY, ext_from=None, label_dx=0.2, ha='left'):
    if ext_from is not None:
        for y in (y1, y2):
            ax.plot([ext_from, x + 0.1 * np.sign(x - ext_from)], [y, y], color=color, lw=0.5, zorder=3)
    arrow(ax, (x, y1), (x, y2), color=color, lw=0.6, style='<|-|>')
    if label:
        text(ax, x + label_dx, (y1 + y2) / 2, label, size=size, color=color, ha=ha)


def leader(ax, p, t, s, size=8.5, ha='left', color=INK, lcolor=GRAY, font=SONGX, dot=True, width=None):
    """引出线：从被注对象点 p 到文字位置 t。"""
    ax.plot([p[0], t[0]], [p[1], t[1]], color=lcolor, lw=0.5, zorder=7)
    if dot:
        ax.add_patch(Circle(p, 0.035, fc=lcolor, ec='none', zorder=7))
    off = 0.06 if ha == 'left' else (-0.06 if ha == 'right' else 0)
    text(ax, t[0] + off, t[1], s, size=size, color=color, ha=ha, font=font, width=width)


def callout(ax, x, y, n, r=0.15, size=7.5):
    ax.add_patch(Circle((x, y), r, fc='white', ec=DARK, lw=0.8, zorder=9))
    ax.text(x, y, str(n), fontsize=size, color=DARK, family=HEI, weight='bold', ha='center', va='center', zorder=10)


def callout_at(ax, p, c, n, r=0.15):
    """带引出线的编号：p 为被注对象点，c 为编号圆心。"""
    ax.plot([p[0], c[0]], [p[1], c[1]], color=GRAY, lw=0.5, zorder=8)
    ax.add_patch(Circle(p, 0.03, fc=GRAY, ec='none', zorder=8))
    callout(ax, c[0], c[1], n, r=r, size=7.5)


def check(ax, x, y, s=0.16, color=MID, lw=1.6):
    ax.plot([x - s, x - s * 0.3, x + s], [y, y - s * 0.7, y + s * 0.8], color=color, lw=lw,
            solid_capstyle='round', zorder=9)


def cross(ax, x, y, s=0.13, color=ACCENT, lw=1.6):
    ax.plot([x - s, x + s], [y - s, y + s], color=color, lw=lw, solid_capstyle='round', zorder=9)
    ax.plot([x - s, x + s], [y + s, y - s], color=color, lw=lw, solid_capstyle='round', zorder=9)


def ground(ax, x1, x2, y, hatch=True, color=GRAY):
    ax.plot([x1, x2], [y, y], color=color, lw=1.0, zorder=3)
    if hatch:
        xs = np.arange(x1 + 0.1, x2, 0.2)
        for x in xs:
            ax.plot([x, x - 0.12], [y, y - 0.12], color=color, lw=0.5, zorder=3)


def hatch_rect(ax, x, y, w, h, fc='#E7E2DA', ec=GRAY, step=0.14, hatch_color='#B8AFA3', lw=0.7, zorder=3):
    """带斜线填充的构件截面（砖墙、混凝土等）。"""
    ax.add_patch(Rectangle((x, y), w, h, fc=fc, ec=ec, lw=lw, zorder=zorder))
    import matplotlib.patches as mp
    clip = mp.Rectangle((x, y), w, h, transform=ax.transData)
    for k in np.arange(-h, w, step):
        ln, = ax.plot([x + k, x + k + h], [y, y + h], color=hatch_color, lw=0.4, zorder=zorder)
        ln.set_clip_path(clip)


def legend_list(ax, x, top, w, items, size=8.5, lh=1.4, gap=0.06, nr=0.15):
    """带圆圈编号的说明列表，返回底 y。items: [(n, 文字)]。"""
    step = size * PT * lh
    y = top - step / 2
    ind = 2 * nr + 0.12
    for n, s in items:
        lines = wrap(s, w - ind, size)
        callout(ax, x + nr, y, n, r=nr, size=7.5)
        for ln in lines:
            ax.text(x + ind, y, ln, fontsize=size, color=INK, family=SONGX, ha='left', va='center')
            y -= step
        y -= gap
    return y + step / 2 + gap


def legend_height(w, items, size=8.5, lh=1.4, gap=0.06, nr=0.15):
    step = size * PT * lh
    ind = 2 * nr + 0.12
    n = sum(len(wrap(s, w - ind, size)) for _, s in items)
    return n * step + gap * (len(items) - 1)
