# docfig · Word 插图 matplotlib 样式库

给标书、报告画流程图、框架图、组织架构图、设备安装示意图。按 Word 版宽实尺（cm）绘制，300dpi 出 PNG，放进 docx 后字号就是印刷磅值。

| 文件 | 内容 |
|---|---|
| `figstyle.py` | 画布 `new(h)`、自动裁白边 `save()`、带表头自动高的框 `box(..., top=, h=None)`、同行等高 `row_height()`、`pill`、`formula`、`arrow`/`poly_arrow`、`title_bar`、`note` |
| `figeng.py` | 管道、法兰、电磁流量计、闸阀、弯头、尺寸线、引线标注、编号圆圈、地面/回填剖面、图例列表 |

```python
import sys; sys.path.insert(0, '/Users/tianli/Dev/tools/doctools/lib/docfig')
from figstyle import *
fig, ax = new(10.0)          # 高 10cm，宽固定 15.5cm
box(ax, 0.3, top=9.5, w=4.5, title='班组自验', body=['数据完整性', '记录规范性'])
save(fig, 'out.png')
```

画完用图片查看逐张看：文字不出框、不压线，内容与所在节正文逐项一致。换进 docx 用
`python3 ~/Dev/tools/doctools/scripts/document/bid_body.py swapfig <docx> --plan plan.json [--apply]`（按图号定位，允许比例变化，可改图题）。

实例：`~/Work/shared/bids/海宁-再生水利用试点-2026/src/figs7/f7_*.py`（第7章 10 张）。
