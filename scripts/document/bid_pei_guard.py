#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_pei_guard.py — 陪标(pei)口径写入即检：谁刚写了标书正文，就立刻对着口径查谁。

Why（2026-08-31 用户 /govern「那么简单的陪标要求，你确保下次不会出错」）：
pei 口径只有 5 条，我全背得出，同一会话仍违反 2 次 —— 两次都是**在修别的东西时顺手
写入新正文，写完不查自己刚写的**：
  · 修一条陈旧交叉引用 →「本章 5.8」改成「见本节「X」」，去掉编号但留了指针，还是交叉引用
  · 写安全章 →「涉水岸线调查」这个招标里 0 命中的场景名；写溯源门 → 把自拟简称说成招标提出的
项目 CLAUDE.md 原来写的是「**交付前**必跑四门」—— 交付前是几十次写入之后，
中间任何一次引入的违规都要等到最后，而最后那一跑还未必有对应规则（无编号指针当天才补）。
结论：**检查点必须挪到写入的当下**，且由 harness 执行，不靠我记得。

用法: bid_pei_guard.py <刚改动的文件>     # 非 pei 项目/非正文文件 → 静默 exit 0
exit 0 = 无违规或不适用 · 2 = 有违规 · 1 = 用法/IO 错
"""
import re, subprocess, sys, zipfile
from pathlib import Path

DOC = Path(__file__).resolve().parent
GATE = DOC / 'bid_gate.py'


def project_root(p: Path):
    for d in [p if p.is_dir() else p.parent, *(p if p.is_dir() else p.parent).parents]:
        if (d / '_project.yaml').is_file():
            return d
    return None


def is_pei(root: Path):
    r"""→ True(pei) / False(main) / None(没声明口径)。
    ⚠ 值可能带引号（`bid_mode: "pei"`）—— 初版正则 (\w+) 匹配不上，守卫静默放行，
    这正是 fail-open 的典型形态（2026-08-31 自测当场撞上）。"""
    t = (root / '_project.yaml').read_text(encoding='utf-8', errors='ignore')
    m = re.search(r'^\s*bid_mode\s*:\s*["\']?(\w+)', t, re.M)
    if not m:
        return None
    return m.group(1) == 'pei'


def run(cmd):
    r = subprocess.run([sys.executable, *cmd], capture_output=True, text=True, timeout=120)
    return r.returncode, (r.stdout or '') + (r.stderr or '')


def main():
    if len(sys.argv) < 2:
        print(__doc__); return 1
    f = Path(sys.argv[1]).resolve()
    if not f.exists():
        return 0
    root = project_root(f)
    # 只管标书工作区里的 pei 项目正文
    if root is None or 'shared/bids' not in str(root):
        return 0
    mode = is_pei(root)
    if mode is None:
        print(f'[bid-pei-guard] {root.name}/_project.yaml 没声明 bid_mode —— 口径未定，'
              f'本轮不检查。开工第 0 步就该定 main|pei。', file=sys.stderr)
        return 0
    if not mode:
        return 0
    is_docx = f.suffix.lower() == '.docx'
    in_body = is_docx or (f.suffix.lower() == '.md' and '成果' in f.parts)
    if not in_body or f.name.startswith('~$') or '.bak-' in f.name:
        return 0

    findings = []
    tgt = str(f) if is_docx else str(f.parent)

    # 口径③ 禁交叉引用 —— 直接用 bid_residue_lib.XREF_PATS 这**一套**规则同时管 md 与 docx。
    # 初版走 `bid_gate deref --check` 判，漏了：deref 只做「号→标题名」的修复，不认无编号
    # 位置指针；而带 XREF_PATS 的 scan 只跑 docx，md 侧因此零检测 —— 守卫对着刚犯过的错
    # 报 exit 0（2026-08-31 自测撞上，同一守卫第三次 fail-open）。规则只有一处，别两套。
    sys.path.insert(0, str(DOC))
    import bid_residue_lib as _lib  # noqa: E402
    if is_docx:
        import zipfile as _z
        blob = _z.ZipFile(f).read('word/document.xml').decode('utf-8', 'ignore')
        lines = [''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', seg))
                 for seg in blob.split('</w:p>')]
    else:
        lines = f.read_text(encoding='utf-8', errors='ignore').split('\n')
    for ln in lines:
        s = ln.strip()
        if not s or (_lib.XREF_SKIP.match(s) and len(s) < 60):
            continue
        marks = [m for pat in _lib.XREF_PATS for m in pat.findall(s)]
        if marks:
            findings.append(f'[③禁交叉引用] {s[:70]}  ⟨命中: {marks[0]}⟩')
    # 口径② 素材只用公开资料·措辞回得了源
    aud = root / 'scripts' / 'provenance_audit.py'
    if aud.is_file():
        rc, out = run([str(aud), tgt])
        if rc == 2:
            findings += [f'[②溯源] {l.strip()}' for l in out.split('\n') if l.strip().startswith('✗')]
    # 口径①④ 身份泄漏 / 评分脚手架 / 加粗（docx 才有样式层）
    if is_docx:
        rc, out = run([str(GATE), 'scan', str(f), '--mode', 'pei'])
        if rc == 2:
            findings += [f'[①④残留] {l.strip()}' for l in out.split('\n') if l.strip().startswith('[类别')]

    # 显式豁免：项目 _project.yaml 里 pei_guard_waived 逐条声明「哪一类 + 为什么不适用」。
    # Why：交付形态若是合稿后 PDF 打印，docProps 元数据不随 PDF 带出 —— 这条真不适用。
    # 但守卫每次都为它报红 = 天天狼来了，最后整个守卫被无视（今天已在别的门上栽过）。
    # 所以给显式豁免口，**并把豁免打印出来**（可复查），不做静默放过。
    waived = []
    pt = (root / '_project.yaml').read_text(encoding='utf-8', errors='ignore')
    m = re.search(r'^pei_guard_waived:\s*$([\s\S]*?)(?=^\S|\Z)', pt, re.M)
    if m:
        for ln in m.group(1).split('\n'):
            ln = ln.strip()
            if ln.startswith('- '):
                waived.append(ln[2:].split(':')[0].strip())
    if waived:
        kept = []
        for x in findings:
            if any(w and w in x for w in waived):
                continue
            kept.append(x)
        if len(kept) != len(findings):
            print(f'[bid-pei-guard] 已按 _project.yaml pei_guard_waived 豁免 '
                  f'{len(findings) - len(kept)} 条（{"、".join(waived)}）', file=sys.stderr)
        findings = kept

    if not findings:
        return 0
    print('⛔ [bid-pei-guard] 刚写入的内容违反陪标口径 —— 现在改，别留到交付前：', file=sys.stderr)
    for x in findings[:12]:
        print('  ' + x, file=sys.stderr)
    if len(findings) > 12:
        print(f'  …另 {len(findings) - 12} 条', file=sys.stderr)
    print(f'  口径 5 条见 {root.name}/CLAUDE.md §口径纪律；逃生 BID_PEI_OK=1', file=sys.stderr)
    return 2


if __name__ == '__main__':
    sys.exit(main())
