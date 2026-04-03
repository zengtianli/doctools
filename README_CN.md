**中文** | [English](README.md)

# doctools

文档处理与数据转换工具集，主要通过 Raycast 调用。

## 文档处理 (document/)

| 脚本 | 功能 |
|------|------|
| `bullet_to_paragraph.py` | 要点转公文段落/表格（AI） |
| `chart.py` | 数据驱动图表生成（JSON → PNG） |
| `docx_apply_template.py` | Word 文档样式套模板 + 清理 |
| `docx_text_formatter.py` | 文本格式自动修复（DOCX） |
| `docx_tools.py` | Word 文档工具集 |
| `md_docx_template.py` | Markdown 转 Docx（样式复刻） |
| `md_tools.py` | Markdown 工具集 |
| `pptx_to_md.py` | PPTX 转 Markdown |
| `pptx_tools.py` | PPTX 文档标准化工具集 |
| `report_quality_check.py` | 报告/标书质量检查 + 自动修复 |
| `scan_sensitive_words.py` | AI 敏感词扫描器 |

## 数据转换 (data/)

| 脚本 | 功能 |
|------|------|
| `convert.py` | 数据格式转换统一工具（8 种格式互转） |
| `xlsx_merge_tables.py` | Excel 多表合并（AI 智能匹配） |
| `xlsx_splitsheets.py` | Excel 工作表拆分 |

## 安装

```bash
pip3 install -r requirements.txt
```
