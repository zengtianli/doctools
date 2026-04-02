"""
Word XML namespace 常量与工具函数

供 docx_track_changes / docx_extract / md_docx_template 共用
"""

# ── WordprocessingML 命名空间 ─────────────────────────────────

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

W = NSMAP["w"]
R_NS = NSMAP["r"]

# 关系类型常量
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"


def qn(tag: str) -> str:
    """将 'w:t' 转换为 '{namespace}t' 格式"""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"
