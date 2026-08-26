#!/usr/bin/env python3
"""docx_quotes.py — 中文文档直角引号「」『』 检查 / 修复（用户 2026-08-22 /govern 钦定，全局有效）。

规则：外发/递交的中文文档一律用弯引号 “” ‘’，禁出现直角引号 「」『』。
      内部件 / 聊天 / 代码注释不受限——所以本工具只管被点名的文件，不自己决定谁是外发件。

用法：
  docx_quotes.py <file...>            检查；有命中 → 打印 文件:计数 并 exit 1；全干净 exit 0
  docx_quotes.py --apply <file...>    原地替换（docx 先备份 .bak-<ts>；md/txt 直接改），再复检
支持 .docx（word/*.xml 里的 <w:t> 文本）与 .md/.txt/.html 纯文本。
fail-closed：文件不存在 / 读不了 → exit 2。
"""
import io, os, re, shutil, sys, time, zipfile
from pathlib import Path

# 手工 zipfile 重打包 docx —— docx_safe_save 只 patch python-docx 的存盘路径，管不到这里，
# 所以按第二判据（2026-07-31 立）挂部件完整性断言：丢部件 / 改了没报备的部件即抛，不静默。
sys.path.append(str(Path(__file__).resolve().parents[2] / "lib"))
from docx_parts import assert_parts_intact  # noqa: E402

CORNER = str.maketrans({"「": "“", "」": "”", "『": "‘", "』": "’"})
PAT = re.compile(r"[「」『』]")
WT = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)


def count_text(s):
    return len(PAT.findall(s))


def docx_parts(z):
    return [n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")]


def check_docx(path):
    with zipfile.ZipFile(path) as z:
        return sum(count_text(z.read(n).decode("utf-8", "ignore")) for n in docx_parts(z))


def fix_docx(path):
    ts = time.strftime("%Y%m%d-%H%M%S")
    bak = f"{path}.bak-{ts}"
    shutil.copy2(path, bak)
    tmp = path + ".tmp"
    touched = set()          # 本次真改了引号的部件 —— 只有这些准变，其余一个字节都不许动
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in docx_parts(zin):
                s = data.decode("utf-8")
                fixed = WT.sub(lambda m: m.group(1) + m.group(2).translate(CORNER) + m.group(3), s)
                if fixed != s:
                    touched.add(item.filename)
                data = fixed.encode("utf-8")
            zout.writestr(item, data)
    # 拿备份当 src 比对：改的必须**正好**是 touched 这些，多一个少一个都抛
    assert_parts_intact(bak, tmp, allow_changed=touched, verbose=False)
    os.replace(tmp, path)


def check_text(path):
    return count_text(io.open(path, encoding="utf-8").read())


def fix_text(path):
    s = io.open(path, encoding="utf-8").read()
    io.open(path, "w", encoding="utf-8").write(s.translate(CORNER))


def main(argv):
    apply = "--apply" in argv
    files = [a for a in argv if not a.startswith("--")]
    if not files:
        print(__doc__); return 2
    bad = 0
    for p in files:
        if not os.path.isfile(p):
            print(f"⛔ 不存在：{p}"); return 2
        is_docx = p.lower().endswith(".docx")
        try:
            n = check_docx(p) if is_docx else check_text(p)
        except Exception as e:
            print(f"⛔ 读不了 {p}: {e}"); return 2
        if n and apply:
            (fix_docx if is_docx else fix_text)(p)
            n2 = check_docx(p) if is_docx else check_text(p)
            print(f"✔ 已替换 {p}：{n} → {n2}")
            n = n2
        if n:
            bad += 1
            print(f"✗ {p}：直角引号 {n} 处（修复：python3 {os.path.abspath(__file__)} --apply '{p}'）")
    if bad:
        return 1
    print(f"docx_quotes: PASS（{len(files)} 件，0 处直角引号）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
