"""`pdf_cli.py slim` 的回归门 —— 把 2026-08-14 那起真实事故固化成 fixture。

## 事故本体

一本 2228 页的汇编在 UPDF 里删到 1126 页，文件从 664,012,314 B 变成
664,023,365 B —— **删掉 1102 页，文件长大了 11,051 字节**。

根因是 **PDF 增量保存**：编辑器不重写原文件，只在末尾追加「改动对象 + 新 xref」。
删页 = 追加一张新页树说「只显示这些页」，正文一页没撕。实测上册前 664,012,314
字节与整本 md5 完全相同，即上册物理上就是「整本 + 11KB 新目录」；38,959 个对象里
只有 21,994 个从 /Root 走得到，其余 16,965 个（305 MB）是谁都指不到的死肉。

## 为什么单开一个文件

`pdf_cli` 不在本仓那套闸门（cli_surface / forward_probe / smoke）的覆盖面内 ——
实测 `cli_surface` 在加了 `slim` 前后 diff 为 0 行，它只指纹 docx 族。所以这条动词
**没有任何机器层兜底**，只能自带回归门。

## 判据（全部 fail-closed）

1. **孤儿要真被数出来**：增量删页过的文件必须报出 orphan 对象与孤儿字节。
   这条挂了 = 诊断退化成"看着像在诊断"。
2. **干净件不许吹牛**：单次全量写入的文件必须报 `%%EOF = 1`、孤儿 ≈ 0，
   且无损回收后不许声称省了一大截。假的节省比不诊断更坏。
3. **只读输入**：跑完源件 md5 必须一字节没变（除非显式 --in-place）。
4. **页数守卫**：产物页数必须等于源件。回收动了页数 = 出事了。
5. **-o 指向源件必须 rc=2**：别让「瘦身」变成原地覆盖。
"""
from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

import pytest

_CLI = Path(__file__).resolve().parents[1] / "pdf_cli.py"
_PY = sys.executable


# ──────────────────────────────────────────────────────────────────────
# fixture 造件：手写原始 PDF 字节
#
# 不用 pypdf 造 —— 要测的就是「增量层」这个字节级现象，用高层库造等于
# 让被测对象自己决定长什么样（铁律 #2：验证不能拿替身）。
# ──────────────────────────────────────────────────────────────────────

def _xref_block(entries: list[tuple[int, int]], size: int,
                root: int, prev: int | None) -> bytes:
    """entries = [(objnum, offset), ...]，按 objnum 排序后压成 xref 子段。"""
    out = [b"xref\n"]
    if prev is None:
        out.append(b"0 1\n0000000000 65535 f \n")
    for num, off in sorted(entries):
        out.append(f"{num} 1\n".encode())
        out.append(f"{off:010d} 00000 n \n".encode())
    tr = f"trailer\n<</Size {size}/Root {root} 0 R"
    if prev is not None:
        tr += f"/Prev {prev}"
    tr += ">>\n"
    out.append(tr.encode())
    return b"".join(out)


def _make_base(pad: int = 60_000) -> bytes:
    """2 页的干净 PDF，单次全量写入（%%EOF = 1）。

    第 2 页的内容流塞 pad 字节填充，好让它被孤立后孤儿字节明显可测。
    """
    buf = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs: dict[int, int] = {}

    def obj(num: int, body: bytes) -> None:
        offs[num] = len(buf)
        buf.extend(f"{num} 0 obj\n".encode() + body + b"\nendobj\n")

    def stream(num: int, data: bytes) -> None:
        obj(num, f"<</Length {len(data)}>>\nstream\n".encode()
            + data + b"\nendstream")

    obj(1, b"<</Type/Catalog/Pages 2 0 R>>")
    obj(2, b"<</Type/Pages/Kids[3 0 R 5 0 R]/Count 2>>")
    obj(3, b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Contents 4 0 R>>")
    stream(4, b"BT ET\n")
    obj(5, b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Contents 6 0 R>>")
    stream(6, b"BT ET\n" + b"%" + b"x" * pad + b"\n")

    xref_at = len(buf)
    buf.extend(_xref_block(list(offs.items()), size=7, root=1, prev=None))
    buf.extend(f"startxref\n{xref_at}\n%%EOF\n".encode())
    return bytes(buf)


def _append_incremental_delete_page2(base: bytes) -> bytes:
    """在 base 末尾追加一层增量更新：页树改成只剩第 1 页。

    这正是 UPDF「删页 + 保存」做的事 —— 前面的字节一个不动，
    旧页树（2）连同第 2 页（5）与它 60KB 的内容流（6）就此没人引用 = 孤儿。

    ⚠ 必须连第 1 页的 `/Parent` 一起改写指向新页树。只改 Catalog 的话，
      保留页会经 `/Parent` 反向指回旧页树，旧页树又 `/Kids` 指回被删的页 ——
      整条死链原地复活成"可达"，孤儿只剩 1 个。首版 fixture 就栽在这，
      引擎当时报的 1 个孤儿是对的。真实编辑器会一并改 /Parent。
    """
    prev_start = int(base.rsplit(b"startxref\n", 1)[1].split(b"\n")[0])
    buf = bytearray(base)
    offs: dict[int, int] = {}

    def obj(num: int, body: bytes) -> None:
        offs[num] = len(buf)
        buf.extend(f"{num} 0 obj\n".encode() + body + b"\nendobj\n")

    obj(7, b"<</Type/Pages/Kids[3 0 R]/Count 1>>")
    obj(1, b"<</Type/Catalog/Pages 7 0 R>>")
    obj(3, b"<</Type/Page/Parent 7 0 R/MediaBox[0 0 612 792]/Contents 4 0 R>>")

    xref_at = len(buf)
    buf.extend(_xref_block(list(offs.items()), size=8, root=1, prev=prev_start))
    buf.extend(f"startxref\n{xref_at}\n%%EOF\n".encode())
    return bytes(buf)


def _run(*argv: str) -> subprocess.CompletedProcess:
    return subprocess.run([_PY, str(_CLI), "slim", *argv],
                          capture_output=True, text=True)


@pytest.fixture()
def bloated(tmp_path: Path) -> Path:
    p = tmp_path / "bloated.pdf"
    p.write_bytes(_append_incremental_delete_page2(_make_base()))
    return p


@pytest.fixture()
def clean(tmp_path: Path) -> Path:
    p = tmp_path / "clean.pdf"
    p.write_bytes(_make_base())
    return p


# ──────────────────────────────────────────────────────────────────────
# 判据 1 · 孤儿要真被数出来
# ──────────────────────────────────────────────────────────────────────

def test_incremental_delete_is_detected(bloated: Path):
    # fixture 自检：删了一页，文件反而变大 —— 事故本体先复现出来
    assert len(bloated.read_bytes()) > len(_make_base()), \
        "fixture 没复现『删页反而变大』，后面的断言就没有意义"

    r = _run(str(bloated), "--diag-only")
    assert r.returncode == 0, r.stderr
    out = r.stdout
    assert "%%EOF = 2" in out, f"没数出增量层：\n{out}"
    assert "1 页" in out, f"页树没生效：\n{out}"
    # 旧页树 2 + 被删页 5 + 它的内容流 6，三个都该落成孤儿
    assert "孤儿 3 个对象" in out, f"孤儿对象没数出来：\n{out}"
    # 60KB 的死内容流必须体现在孤儿字节里
    dead_mb = [ln for ln in out.splitlines() if "孤儿流" in ln]
    assert dead_mb and "0.1 MB" in dead_mb[0], \
        f"孤儿字节没数出来：{dead_mb}"
    assert "死肉型" in out, f"病因分诊没判对：\n{out}"


# ──────────────────────────────────────────────────────────────────────
# 判据 2 · 干净件不许吹牛
# ──────────────────────────────────────────────────────────────────────

def test_clean_file_claims_nothing(clean: Path):
    r = _run(str(clean), "--diag-only")
    assert r.returncode == 0, r.stderr
    out = r.stdout
    assert "%%EOF = 1" in out, f"干净件被误判成增量：\n{out}"
    assert "死肉型" not in out, f"干净件被误诊成死肉型：\n{out}"


# ──────────────────────────────────────────────────────────────────────
# 判据 3+4 · 只读输入 · 页数守卫
# ──────────────────────────────────────────────────────────────────────

def test_shrink_is_readonly_and_preserves_pages(bloated: Path, tmp_path: Path):
    before = hashlib.md5(bloated.read_bytes()).hexdigest()
    out_pdf = tmp_path / "out.pdf"
    r = _run(str(bloated), "-o", str(out_pdf))
    assert r.returncode == 0, r.stderr + r.stdout
    assert out_pdf.is_file()

    assert hashlib.md5(bloated.read_bytes()).hexdigest() == before, \
        "源件被改了 —— slim 只读输入、只写 --output"

    assert out_pdf.stat().st_size < bloated.stat().st_size, \
        "回收后没变小，孤儿没被丢掉"

    import pypdf
    assert len(pypdf.PdfReader(str(out_pdf)).pages) == 1, "页数守卫失效"


# ──────────────────────────────────────────────────────────────────────
# 判据 5 · 出口码
# ──────────────────────────────────────────────────────────────────────

def test_missing_file_is_rc2(tmp_path: Path):
    r = _run(str(tmp_path / "nope.pdf"))
    assert r.returncode == 2, f"文件不存在应 rc=2，实得 {r.returncode}"


def test_output_equal_to_source_is_refused(bloated: Path):
    r = _run(str(bloated), "-o", str(bloated))
    assert r.returncode == 2, \
        f"-o 指向源件必须拒绝（否则瘦身=原地覆盖），实得 {r.returncode}"


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
