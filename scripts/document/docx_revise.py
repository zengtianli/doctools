#!/usr/bin/env python3
"""docx_revise.py — 修订注入 CLI：ops.yaml(意见=数据) → docx 修订标记+批注。

内容与引擎解耦：本文件只是薄壳，引擎在 lib/docx_revise.py；
专家意见/替换文本/批注全部写在 ops.yaml 里，项目目录不再现编注入脚本。
ops.yaml 写法见 config/spec-examples/revise-ops-example.yaml。

用法:
  python3 docx_revise.py <ops.yaml> [--dry-run] [--src X.docx] [--out Y.docx]

yaml 里的相对路径相对 ops.yaml 所在目录解析（数据文件跟着项目走，从哪跑都一样）。
"""
import argparse
import importlib.util
import sys
from pathlib import Path

import yaml

# 本文件与 lib/docx_revise.py 同名，裸 import 会命中 sys.path[0]（本目录）import 到自己——
# 按路径显式加载，绕开同名遮蔽
_LIB = Path(__file__).resolve().parent.parent.parent / 'lib'
sys.path.append(str(_LIB))          # docx_revise 内部还要 import docx_parts
_spec = importlib.util.spec_from_file_location('_docx_revise_lib', _LIB / 'docx_revise.py')
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
ReviseError, apply_ops = _mod.ReviseError, _mod.apply_ops


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument('ops_yaml', help='意见数据文件(作者/日期/src/out/ops 全在里面)')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--src', help='覆盖 yaml 里的 src')
    ap.add_argument('--out', help='覆盖 yaml 里的 out')
    a = ap.parse_args()

    yml = Path(a.ops_yaml)
    if not yml.exists():
        print(f'✗ 找不到 ops 文件：{yml}', file=sys.stderr)
        return 2
    cfg = yaml.safe_load(yml.read_text(encoding='utf-8'))
    missing = [k for k in ('author', 'date', 'ops') if not cfg.get(k)]
    if missing or not (a.src or cfg.get('src')) or not (a.out or cfg.get('out')):
        print(f'✗ ops.yaml 缺字段：{missing + [k for k in ("src", "out") if not cfg.get(k)]}',
              file=sys.stderr)
        return 2

    base = yml.resolve().parent
    src = Path(a.src) if a.src else base / cfg['src']
    out = Path(a.out) if a.out else base / cfg['out']

    try:
        stats = apply_ops(src, out, cfg, dry=a.dry_run)
    except ReviseError as e:
        print(e.code, file=sys.stderr)
        return 2
    for line in stats.pop('log'):
        print(f'  {line}')
    print(f"  → {stats}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
