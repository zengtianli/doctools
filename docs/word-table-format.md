# Word 表格一键整理

脚本原版：[word_table_format.command](../scripts/document/word_table_format.command)。适用于 macOS 的 Microsoft Word 和已保存在本机的 `.docx` 文档，不需要安装 Python 包或 Word 加载项。

## 使用

1. 在 Word 中打开要处理的文档；打开多份文件时，先切换到目标文档。
2. 在 Finder 中双击 `word_table_format.command`。
3. 等待终端显示完成信息。脚本会保存当前编辑内容，在原文档旁生成 `.bak-年月日-时分秒` 备份，再整理并保存原文档。

首次运行如果 macOS 询问是否允许终端控制 Microsoft Word，允许后即可通过 Word 文档接口处理。

每次运行会统一已有表格的三项格式：

- 单元格、段落和文字底纹设为“无颜色”。
- 单元格内容水平、垂直居中。
- 外框和内部横竖线设为 **0.5 磅黑色单实线**，包括原先缺少的边框。

脚本遍历正文、页眉页脚、文本框、脚注等区域，并递归处理嵌套表格。它保留文字、图片、字体、列宽及表格样式名称。以后新增表格或再次粘贴带格式的表格，重新运行即可；脚本不会作为后台服务自动监控文档。

也可以在终端运行，并传入路径校验当前处理对象：

```bash
/Users/tianli/Dev/tools/doctools/scripts/document/word_table_format.command "/完整路径/文档.docx"
```

不传路径时处理 Word 当前文档。`--help` 仅显示帮助，不修改文档。脚本使用 Word 文档对象接口，不模拟鼠标键盘操作，不使用剪贴板或切换窗口。

需要恢复时，先关闭对应 Word 文档，将备份复制为一个以 `.docx` 结尾的文件再打开核对。
