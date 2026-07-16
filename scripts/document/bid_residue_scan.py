#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_residue_scan.py — 标书终稿 8 类残留【只读】扫描器。

逐段（lxml 解 word/document.xml，段=w:p，段文本=拼接 w:t）匹配内置 taxonomy
+ --rules YAML 项目增补，输出分类报告：每条 [类别N] P<段号4位> 摘录(前80字) 处置建议。
另扫 docProps/core.xml 与全部 XML part（第8类身份泄漏）。

用法: python3 bid_residue_scan.py <docx> [--mode main|pei] [--rules <yaml>]
  --mode pei（默认）全量 8 类；main 跳过第8类实名类（公司名/院自指/业绩归属/
         裸院/元数据署名），仍报工具痕迹（python-docx）。
exit 0 = PASS（0 findings）；exit 2 = FAIL n findings；exit 1 = 用法/IO 错误。
检测逻辑 SSOT 在同目录 bid_residue_lib.py（bid_finalize_sweep 复扫共用）。
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bid_residue_lib as lib


def main():
    ap = argparse.ArgumentParser(description="标书终稿 8 类残留只读扫描器")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--mode", choices=["main", "pei"], default="pei")
    ap.add_argument("--rules", type=Path, default=None)
    args = ap.parse_args()

    if not args.docx.is_file():
        print(f"错误: 文件不存在 {args.docx}", file=sys.stderr)
        return 1
    try:
        rules = lib.load_rules(args.rules)
        names, parts = lib.load_parts(args.docx)
        if "word/document.xml" not in parts:
            print(f"错误: 不是有效 docx（缺 word/document.xml）: {args.docx}", file=sys.stderr)
            return 1
        findings = lib.scan_parts(parts, mode=args.mode, rules=rules)
    except (OSError, ValueError, RuntimeError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    print(f"bid_residue_scan · {args.docx.name} · mode={args.mode}"
          f" · rules={args.rules.name if args.rules else '(内置通用)'}")
    if findings:
        by_cat = {}
        for f in findings:
            by_cat[f["cat"]] = by_cat.get(f["cat"], 0) + 1
        print("分类计数:", " ".join(f"类别{c}({lib.CAT_NAMES[c]})={n}" for c, n in sorted(by_cat.items())))
        for line in lib.format_findings(findings):
            print(line)
        print(f"FAIL {len(findings)} findings")
        return 2
    print("8 类残留全零（协作标记/拟hedge/评分脚手架/内部编号/二次残渣/断裂引用/口径meta/身份泄漏）")
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
