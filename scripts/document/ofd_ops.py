#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ofd_ops.py — OFD（GB/T 33190 版式文档）读取族。政务电子证照/公文的默认格式，Read 认不了。

子命令:
  read    <f.ofd> [--out FILE] [--pages 1-5]   抽正文文本（默认 stdout）
  info    <f.ofd>                              页数 / 签章 / 内嵌资源 / 元数据
  extract image <f.ofd> --out DIR              抽内嵌图（噪音过滤 + 页归属命名）
  to-pdf  <f.ofd|dir>... [--out-dir D]         版面转换 → 图片 PDF（多文件/目录；需 java + vendor/ofdrw）
  batch   <dir> [--out-suffix .正文摘录.txt]    递归批量 read 落盘

OFD = zip + XML。正文在 Doc_0/Pages/Page_N/Content.xml 的 <ofd:TextCode>，图在 Doc_0/Res/。
纯标准库，零依赖。**抽的是文本不是版面** —— 电子签章、附图、排版位置不在 read 结果里。

⚠ 别用 easyofd 转 PDF（2026-08-25 实测，磐安花溪取水许可证）：
   macOS 无 simsun.ttc/simhei.ttf，easyofd 的字体注册**逐个失败但不报错**，
   产出 PDF 只剩二维码与表格框线，**文字一个不剩**；3.4 MB 源膨胀成 20 MB。
   它有输出、有正确页数、体积还更大 —— 看着完全像成功，Read 一眼才发现是空壳。
   这就是「测量通道没验就信结果」的标准形态，所以本引擎的 to-pdf 宁可 fail-closed。

exit: 0 正常 · 1 失败/文件不存在/非 OFD · 2 空枚举（0 页 / 0 图 / 目录下 0 个 .ofd）
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

PAGE_RE = re.compile(r"Pages/Page_(\d+)/Content\.xml$")
RES_RE = re.compile(r"/Res/(.+)$")
IMG_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}
# 噪音阈值：政务 OFD 的 Res/ 里常混 soft-mask、纯色块、几十字节的占位图
MIN_IMG_BYTES = 3 * 1024


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _open(path: Path) -> zipfile.ZipFile:
    if not path.exists():
        sys.exit(f"[ofd_ops] 文件不存在: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        sys.exit(f"[ofd_ops] 不是合法 OFD（OFD 本质是 zip）: {path}")
    if not any(n.endswith("OFD.xml") for n in z.namelist()):
        sys.exit(f"[ofd_ops] zip 里没有 OFD.xml，不像 OFD 文档: {path}")
    return z


def _pages(z: zipfile.ZipFile) -> list[tuple[int, str]]:
    return sorted((int(m.group(1)), n) for n in z.namelist() if (m := PAGE_RE.search(n)))


def _page_text(z: zipfile.ZipFile, name: str) -> str:
    root = ET.fromstring(z.read(name))
    return "".join(e.text or "" for e in root.iter() if _local(e.tag) == "TextCode").strip()


def _parse_range(spec: str, total: int) -> set[int]:
    """'1-5,8' → {1,2,3,4,5,8}（1-based，与人读的页码一致）。"""
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    bad = {p for p in out if not 1 <= p <= total}
    if bad:
        sys.exit(f"[ofd_ops] --pages 超出范围（共 {total} 页）: {sorted(bad)}")
    return out


def cmd_read(args) -> int:
    src = Path(args.file)
    z = _open(src)
    pages = _pages(z)
    if not pages:
        print(f"[ofd_ops] {src.name}: 0 个页面 Content.xml —— 拒绝在空集上报绿", file=sys.stderr)
        return 2
    want = _parse_range(args.pages, len(pages)) if args.pages else None
    chunks = [f"[{src.name}] 共 {len(pages)} 页\n"]
    for idx, name in pages:
        if want and (idx + 1) not in want:
            continue
        chunks.append(f"===== 第 {idx + 1} 页 =====\n{_page_text(z, name)}")
    body = "\n\n".join(chunks)
    if args.out:
        Path(args.out).write_text(body, encoding="utf-8")
        print(f"[ofd_ops] → {args.out}（{len(body)} 字符 / {len(pages)} 页）")
    else:
        print(body)
    return 0


def cmd_info(args) -> int:
    src = Path(args.file)
    z = _open(src)
    pages = _pages(z)
    names = z.namelist()
    signs = sorted({n.split("/Sign_")[1].split("/")[0] for n in names if "/Sign_" in n})
    res = [(m.group(1), z.getinfo(n).file_size) for n in names if (m := RES_RE.search(n))]
    imgs = [(n, s) for n, s in res if Path(n).suffix.lower() in IMG_EXT]
    fonts = [(n, s) for n, s in res if Path(n).suffix.lower() in {".ttf", ".otf", ".ttc"}]

    print(f"文件      {src.name}")
    print(f"体积      {src.stat().st_size:,} 字节")
    print(f"页数      {len(pages)}")
    print(f"电子签章  {len(signs)} 个" + (f"（Sign_{', Sign_'.join(signs)}）" if signs else "  ← 无签章"))
    print(f"内嵌字体  {len(fonts)} 个" + ("  ← 矢量文本，read 抽得到字" if fonts else ""))
    print(f"内嵌图片  {len(imgs)} 个"
          + (f"（≥{MIN_IMG_BYTES // 1024}KB 的 {sum(1 for _, s in imgs if s >= MIN_IMG_BYTES)} 个）" if imgs else ""))
    if pages:
        head = _page_text(z, pages[0][1])
        print(f"首页文本  {len(head)} 字符" + ("" if head else "  ← 0 字符：可能是扫描图 OFD，read 会空手而归"))
        if head:
            print(f"          {head[:80]}…")
    return 0 if pages else 2


def _media_map(z: zipfile.ZipFile) -> dict[str, str]:
    """MultiMedia ID → MediaFile 文件名，出自 *Res.xml。

    ⚠ ResourceID **不等于**文件名里的数字（实测 ID 13 → image_12.jpg，差 1，且不是
    固定偏移）。2026-08-25 第一版按文件名剥数字去猜 ID，4 张图的页归属全部落空、
    静默降级成 item-NN —— 看着有输出，其实契约里那条「按页归属命名」压根没生效。
    映射必须从 Res.xml 读，不能推。
    """
    out: dict[str, str] = {}
    for n in z.namelist():
        if not n.endswith("Res.xml"):
            continue
        try:
            root = ET.fromstring(z.read(n))
        except ET.ParseError:
            continue
        for e in root.iter():
            if _local(e.tag) != "MultiMedia":
                continue
            mid = e.get("ID")
            f = next((c.text for c in e.iter() if _local(c.tag) == "MediaFile" and c.text), None)
            if mid and f:
                out[mid] = f.strip()
    return out


def cmd_extract_image(args) -> int:
    src = Path(args.file)
    z = _open(src)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    id2file = _media_map(z)
    # 文件名 → 首次出现的页码（1-based）。OFD 没有 caption，页归属就是能给的最强语义。
    file2page: dict[str, int] = {}
    for idx, name in _pages(z):
        root = ET.fromstring(z.read(name))
        for e in root.iter():
            if _local(e.tag) != "ImageObject":
                continue
            f = id2file.get(e.get("ResourceID") or "")
            if f and f not in file2page:
                file2page[f] = idx + 1

    entries = [(m.group(1), n) for n in z.namelist() if (m := RES_RE.search(n))]
    imgs = [(base, n) for base, n in entries if Path(base).suffix.lower() in IMG_EXT]
    if not imgs:
        print(f"[ofd_ops] {src.name}: Res/ 下 0 个图片资源 —— 拒绝在空集上报绿", file=sys.stderr)
        return 2

    kept = skipped = orphan = 0
    for seq, (base, name) in enumerate(sorted(imgs), 1):
        if z.getinfo(name).file_size < MIN_IMG_BYTES:
            skipped += 1
            continue
        page = file2page.get(base)
        if page:
            stem = f"page-{page:03d}-{base}"
        else:
            stem = f"page-000-item-{seq:02d}-{base}"   # 没被任何页引用：显式标 000，不静默混进正常命名
            orphan += 1
        (out / stem).write_bytes(z.read(name))
        kept += 1
    msg = f"[ofd_ops] → {out}  抽出 {kept} 张，按 <{MIN_IMG_BYTES // 1024}KB 过滤掉 {skipped} 张噪音"
    if orphan:
        msg += f"；{orphan} 张未被任何页引用（命名 page-000-item-NN）"
    print(msg)
    return 0 if kept else 2



def _page_has_content(z: zipfile.ZipFile, page_xml: str) -> bool:
    """这一页在原件里到底有没有东西（文本 / 图 / 路径）。

    用于把「原件就是空白页」与「渲染丢了内容」分开 —— 只有后者才该 fail-closed。
    """
    try:
        raw = z.read(page_xml).decode("utf-8", "ignore")
    except KeyError:
        return False
    if _page_text(z, page_xml).strip():
        return True
    return any(tag in raw for tag in ("ImageObject", "PathObject", "TextObject", "CompositeObject"))

def _one_to_pdf(src: Path, out: Path, ppm: int, quiet: bool = False) -> int:
    """OFD → PDF。走 **图片路线**（ofdrw ImageMaker 逐页自绘 → PIL 合成 PDF）。

    为什么不用 ofdrw 的 ConvertHelper.toPdf（PDF 直转）：
        2026-08-27 磐安花溪取水许可证实测 —— 直转出来的 PDF 有 11 页、抽得到 1938 字符、
        国徽红章边框二维码都在，**但字段值在视觉上是散的**：「单位名称 磐 …… 安 …… 自」，
        中间的字被 DeltaX 推到页外（页面右缘挂着一串孤字）。文本层完整所以任何自动检查都报绿，
        只有把页面渲染成图看一眼才发现证件是残的。根因是 macOS 无 simsun/simhei，
        ofdrw 退到 fallback 字体后字宽与原内嵌字体不符，偏移逐字累积。
        同一份走 ImageMaker 逐页渲染则**完全正确**（字段、印章、二维码齐全）。
    代价：产物是图片 PDF，**没有可选中的文字层**。需要检索时用 `ofd_ops.py read` 的
        `.正文摘录.txt`；需要验签时必须回 OFD 原件（转换后签章只剩图像）。

    依赖：java + `~/Dev/tools/doctools/vendor/ofdrw/`（classes + lib，见该目录 README）。
    """
    import shutil
    import subprocess
    import tempfile

    z = _open(src)
    pages_xml = [name for _, name in _pages(z)]
    n_expect = len(pages_xml)
    if n_expect == 0:
        print("[ofd_ops] 这份 OFD 解析出 0 页 —— 拒绝在空集上产出 PDF", file=sys.stderr)
        return 2

    vendor = Path.home() / "Dev/tools/doctools/vendor/ofdrw"
    classes, libs = vendor / "classes", vendor / "lib"
    java = shutil.which("java")
    if not java or not (classes / "OfdToImages.class").is_file() or not libs.is_dir():
        print("[ofd_ops] OFD → PDF 需要 java + ofdrw vendor 目录，本机不齐：", file=sys.stderr)
        print(f"          java={'有' if java else '无'}  {vendor}={'有' if classes.is_dir() else '无'}", file=sys.stderr)
        print(f"          建法见 {vendor}/README.md（brew install maven 后一条命令）", file=sys.stderr)
        print("  只要正文文字：python3 ofd_ops.py read '<f.ofd>'", file=sys.stderr)
        print("  ⚠ 别退而求其次用 easyofd —— 本机产出的是没有文字的空壳（见本文件抬头）。", file=sys.stderr)
        return 1

    try:
        from PIL import Image
    except ImportError:
        print("[ofd_ops] 需要 Pillow 合成 PDF：/opt/homebrew/bin/python3 -m pip install pillow", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory(prefix="ofd2pdf-") as td:
        cmd = [java, "-Xmx4g", "-cp", f"{classes}:{libs}/*", "OfdToImages", str(src), td, str(ppm)]
        r = subprocess.run(cmd, capture_output=True, text=True)
        pages = sorted(Path(td).glob("page-*.png"))
        if r.returncode != 0 or not pages:
            print("[ofd_ops] 渲染失败", file=sys.stderr)
            print((r.stderr or r.stdout).strip()[:1500], file=sys.stderr)
            return 1
        # fail-closed ①：页数必须与 OFD 声明一致（OOM 时会只出前几页且退出码非 0，但双保险）
        if len(pages) != n_expect:
            print(f"[ofd_ops] 页数不符：OFD {n_expect} 页，只渲出 {len(pages)} 页 —— 拒绝产出残册",
                  file=sys.stderr)
            return 1
        imgs = []
        blank = []
        for i, f in enumerate(pages, 1):
            im = Image.open(f).convert("RGB")
            # fail-closed ②：整页近乎纯白 = 这一页什么都没画出来
            small = im.resize((80, 80))
            if min(min(px) for px in small.getdata()) > 245:
                blank.append(i)
            imgs.append(im)
        # 判据不是「渲出来是白的」而是「**原件有内容**却渲成白的」——
        # 报批稿里常有真·空白页（扉页背面、章前空页），把它们判 FAIL 会拦下本来正常的转换。
        # 2026-08-27 实证：云山水厂论证报告第 3 页 Content.xml 只有 148 字节、0 文本 0 图元。
        lost = [i for i in blank if _page_has_content(z, pages_xml[i - 1])]
        if lost:
            print(f"[ofd_ops] 第 {lost} 页原件有内容却渲染成空白 —— 拒绝产出"
                  "（这正是 easyofd 那类空壳的样子）", file=sys.stderr)
            return 1
        if blank:
            print(f"    注：第 {blank} 页原件本身就是空白页，已原样保留。")
        dpi = int(round(ppm * 25.4))
        imgs[0].save(out, "PDF", save_all=True, append_images=imgs[1:], resolution=dpi)

    size = out.stat().st_size
    print(f"OK  {out}")
    print(f"    {len(pages)} 页 · {size / 1e6:.1f} MB · 约 {dpi} dpi · 图片 PDF（无文字层）")
    if not quiet:
        print("    ⚠ 验签能力只在 OFD 原件里，别把原件删了。")
    return 0


def cmd_to_pdf(args) -> int:
    """to-pdf 的 driver：吃 1..N 个文件或目录，目录递归找 *.ofd。

    批量时**一份失败不拖垮其余**（各自 fail-closed），最后汇总并按有无失败定退出码。
    """
    targets: list[Path] = []
    for raw in args.files:
        p = Path(raw)
        if p.is_dir():
            targets.extend(sorted(p.rglob("*.ofd")))
        elif p.is_file():
            targets.append(p)
        else:
            print(f"[ofd_ops] 不存在: {p}", file=sys.stderr)
            return 1
    # 去重（同一份可能既被目录扫到又被显式点名）
    seen, uniq = set(), []
    for p in targets:
        rp = p.resolve()
        if rp not in seen:
            seen.add(rp)
            uniq.append(p)
    targets = uniq

    if not targets:
        print("[ofd_ops] 没找到任何 .ofd —— 拒绝在空集上报成功", file=sys.stderr)
        return 2
    if args.out and len(targets) > 1:
        print(f"[ofd_ops] --out 只能配单个文件，这次匹配到 {len(targets)} 份；"
              "批量请用 --out-dir，或不给（默认落在各自源文件旁）", file=sys.stderr)
        return 1

    ppm = args.ppm or 8
    out_dir = Path(args.out_dir) if args.out_dir else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    ok, failed = [], []
    for k, src in enumerate(targets, 1):
        if len(targets) > 1:
            print(f"[{k}/{len(targets)}] {src.name}")
        if args.out:
            out = Path(args.out)
        elif out_dir:
            out = out_dir / (src.stem + ".pdf")
        else:
            out = src.with_suffix(".pdf")
        if out.exists() and not args.force:
            print(f"    跳过：{out.name} 已存在（--force 覆盖）")
            ok.append(src)
            continue
        rc = _one_to_pdf(src, out, ppm, quiet=(len(targets) > 1))
        (ok if rc == 0 else failed).append(src)

    if len(targets) > 1:
        print(f"\n合计 {len(targets)} 份：成功 {len(ok)} · 失败 {len(failed)}")
        for f in failed:
            print(f"  FAIL {f}")
        if ok:
            print("⚠ 验签能力只在 OFD 原件里，别把原件删了。")
    return 1 if failed else 0


def cmd_batch(args) -> int:
    root = Path(args.dir)
    if not root.exists():
        sys.exit(f"[ofd_ops] 目录不存在: {root}")
    files = sorted(root.rglob("*.ofd")) + sorted(root.rglob("*.OFD"))
    if not files:
        print(f"[ofd_ops] {root} 下 0 个 .ofd —— 拒绝在空集上报绿", file=sys.stderr)
        return 2
    ok = fail = 0
    for f in files:
        dst = f.with_suffix("").with_name(f.stem + args.out_suffix)
        try:
            z = _open(f)
            pages = _pages(z)
            body = "\n\n".join([f"[{f.name}] 共 {len(pages)} 页\n"]
                               + [f"===== 第 {i + 1} 页 =====\n{_page_text(z, n)}" for i, n in pages])
            dst.write_text(body, encoding="utf-8")
            print(f"  ✓ {dst.name}  {len(body)} 字符 / {len(pages)} 页")
            ok += 1
        except SystemExit as e:
            print(f"  ✗ {f.name}  {e}", file=sys.stderr)
            fail += 1
    print(f"[ofd_ops] {ok} 成功 / {fail} 失败 / 共 {len(files)}")
    return 1 if fail else 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="ofd_ops.py", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("read", help="抽正文文本")
    p.add_argument("file"); p.add_argument("--out"); p.add_argument("--pages")
    p.set_defaults(fn=cmd_read)

    p = sub.add_parser("info", help="页数/签章/资源/元数据")
    p.add_argument("file"); p.set_defaults(fn=cmd_info)

    p = sub.add_parser("extract", help="抽内嵌资源")
    s2 = p.add_subparsers(dest="what", required=True)
    p2 = s2.add_parser("image", help="抽内嵌图")
    p2.add_argument("file"); p2.add_argument("--out", required=True)
    p2.set_defaults(fn=cmd_extract_image)

    p = sub.add_parser("to-pdf", help="版面转换 → 图片 PDF（图片路线；不可用时 fail-closed）")
    p.add_argument("files", nargs="+", help="一个或多个 .ofd，或目录（递归找 *.ofd）")
    p.add_argument("--out", help="输出 PDF 路径（只能配单个文件）")
    p.add_argument("--out-dir", help="批量输出目录（默认落在各自源文件旁）")
    p.add_argument("--ppm", type=int, default=8, help="像素/毫米，8≈200dpi（不是 dpi）")
    p.add_argument("--force", action="store_true", help="覆盖已存在的同名 PDF（默认跳过）")
    p.set_defaults(fn=cmd_to_pdf)

    p = sub.add_parser("batch", help="递归批量抽正文落盘")
    p.add_argument("dir"); p.add_argument("--out-suffix", default=".正文摘录.txt")
    p.set_defaults(fn=cmd_batch)

    args = ap.parse_args()
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
