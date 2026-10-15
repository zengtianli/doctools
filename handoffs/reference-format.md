# 按参考角色套用 Word 格式

2026-10-14 新增 `lib/docx_reference.py`，由 Stemma 的格式参考 CLI 消费。它按参考件有效属性继承（含部分 rFonts 属性继承）给目标段落套角色，保留目标表格、分节、页眉页脚、媒体及其他 ZIP 部件；写到新候选，实际只改 `word/document.xml`。合段仅允许显式的简单段落区间，表格、批注、域、书签和未接受修订不被扁平化。

`lib/soffice.py` 本轮在原先未提交的转换实现上增补：读取用户/系统报告字体、复制字体字节而非 macOS 保护标志、在隔离配置内设置 GB2312 字体名到同类字体的兼容映射、识别小标宋的 PostScript 名。没有改 GUI 或安装包。

新引擎 3 项、原 PDF 转换 4 项，共 7 项通过。Stemma 真实庆元两份候选均保留四张表、284 条出处且无新增问题；PDF 实际字体小标宋/黑体/楷体/仿宋，全部内嵌，13 页预览。证据在 `/Users/tianli/Apps/hydro-assistant/stemma/perf/reference-cli/`。

原有 soffice 等未提交改动不是本轮全部所有；不得整文件盲提交。GUI 的 vendor/装机留后续授权迭代。
