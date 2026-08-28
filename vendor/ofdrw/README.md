# vendor/ofdrw —— OFD → PDF 转换后端

2026-08-27 立。被 `scripts/document/ofd_ops.py to-pdf` 调用，不单独使用。

## 是什么

[ofdrw](https://github.com/ofdrw/ofdrw) 的 converter 模块（Apache-2.0）+ 两个自写的 CLI 壳：

| 文件 | 作用 |
|---|---|
| `src/OfdToImages.java` | **在用** —— OFD 逐页渲染成 PNG（`ImageMaker`） |
| `src/OfdToPdf.java` | **留证不用** —— PDF 直转（`ConvertHelper.toPdf`），见下面「为什么不用」 |
| `lib/` | maven 拉下来的依赖 jar（41 个，24 MB，不进 git） |
| `classes/` | 编译产物（不进 git） |

## 重建（换机器 / lib 丢了）

```bash
brew install maven
cd ~/Dev/tools/doctools/vendor/ofdrw
mvn -q dependency:copy-dependencies -DoutputDirectory=lib
javac -cp "lib/*" -d classes src/OfdToImages.java src/OfdToPdf.java
```

## 两个踩过的坑

**① ofdrw 2.0.10 在 macOS 上必崩。** `FontLoader.init()` 按 "宋体/楷体/仿宋" 找默认字体，
macOS `/System/Library/Fonts` 里一个都没有 → `getReplaceSimilarFontPath` 返回 null →
`loadAsDefaultFont(null)` → `Paths.get(null)` NPE。**用 2.4.0**（已修）。
另：`loadAsDefaultFont` 只吃 `.ttf/.otf`，**`.ttc` 一律加载失败**（macOS 的中文字体几乎全是 ttc），
所以 fallback 只能是 `/Library/Fonts/Arial Unicode.ttf`。可用 `OFD_DEFAULT_FONT` 覆盖。

**② 为什么不用 PDF 直转（`ConvertHelper.toPdf`）。** 2026-08-27 磐安花溪取水许可证实测：
直转出的 PDF 有 11 页、抽得到 1938 字符、国徽红章边框二维码俱全，**但字段值视觉上是散的** ——
「单位名称 磐 …… 安 …… 自」，中间的字被 DeltaX 推到页外（页右缘挂着一串孤字）。
文本层完整，所以页数检查、字数检查、体积检查**全部报绿**；只有把页面渲染成图看一眼才发现证件是残的。
根因：fallback 字体的字宽与 OFD 内嵌字体不符，偏移逐字累积。
同一份走 `ImageMaker` 逐页渲染则完全正确 —— 所以 `to-pdf` 走图片路线。

**`ImageMaker(reader, int)` 的第二参是 ppm（像素/毫米），不是 dpi。** 传 150 会得到
44556×31485 的巨图并 OOM。200dpi ≈ 7.87 px/mm → 取 8。

## 代价

产物是**图片 PDF，没有文字层**。要检索用 `ofd_ops.py read` 落的 `.正文摘录.txt`；
**要验签必须回 OFD 原件** —— 转换后签章只剩图像，原件别删。
