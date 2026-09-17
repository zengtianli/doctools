"""bid_body.py 回归门 —— 把 2026-09-17 海宁标实际踩过的坑固化成断言。

## 为什么有这份测试

海宁第 6 章被一次性脚本按「doc.paragraphs 列表」重排：表格不是段落，没跟着动，
结果 9 张表 + sectPr 全部落到文首，整份结构损坏；同一轮还把图插进了「引表短行」
和表题之间，并在同一秒内连续写回、后一份备份盖掉了前一份原始备份。

## 判据

  1. 挪图后：图在正文之后；每张表仍紧跟自己的表题；sectPr 仍在末位
  2. 正文末段若是引出表格的短行，图停在它之前
  3. 损坏形态（表在首个标题之前 / sectPr 不在末位）check 必须判红
  4. 同秒两次写回，两份备份都在
  6. 二级标题下平铺的 N）条目：check 判「标题跳级」；h3group 只插三级标题与引导段、
     组内编号从 1）重排、表仍跟表题；计划条数对不上或重复执行必须拒绝
  5. 来源渲染：简称→正式名、PDF 页→印刷页按已核偏移换算、同文件相邻页合并、
     无法确定的进 errs 而不是猜
"""
from __future__ import annotations

import base64
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
import bid_body as B  # noqa: E402

TOOL = HERE / "bid_body.py"
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")


def build(tmp_path, lead_line=False):
    png = tmp_path / "x.png"
    png.write_bytes(PNG)
    d = Document()
    d.add_heading("1.1 概况", level=3)
    d.add_picture(str(png))
    d.add_paragraph("图1-1 技术联系")
    d.add_heading("1）任务定位", level=4)
    d.add_paragraph("第一段正文，说明任务定位与总体认识。")
    d.add_paragraph("第二段正文，说明证据来源与核验方式。")
    if lead_line:
        d.add_paragraph("指标定义与取数判断")
    d.add_paragraph("表1-1 指标一览")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "指标"
    d.add_heading("2）资源约束", level=4)
    d.add_paragraph("第三段正文。")
    out = tmp_path / "t.docx"
    d.save(out)
    return out


def seq(path):
    doc = B.Doc(path)
    out = []
    for k in doc.body:
        if k.tag == B.w("tbl"):
            out.append("TBL")
        elif k.tag == B.w("sectPr"):
            out.append("SECT")
        elif doc.has_img(k):
            out.append("IMG")
        else:
            out.append(B.ptext(k).strip()[:6])
    return out


def run(*args):
    return subprocess.run([sys.executable, str(TOOL), *map(str, args)], capture_output=True, text=True)


def test_figs_moves_block_after_text_and_keeps_tables(tmp_path):
    f = build(tmp_path)
    assert run("check", f).returncode == 2
    r = run("figs", f, "--apply")
    assert r.returncode == 0, r.stdout + r.stderr
    s = seq(f)
    assert s.index("IMG") > s.index("第二段正文，")
    assert s[s.index("IMG") + 1] == "图1-1 技"
    assert s[s.index("TBL") - 1] == "表1-1 指"          # 表仍紧跟表题
    assert s[-1] == "SECT"
    assert run("check", f).returncode == 0


def test_figs_stops_before_table_lead_line(tmp_path):
    f = build(tmp_path, lead_line=True)
    assert run("figs", f, "--apply").returncode == 0
    s = seq(f)
    i = s.index("IMG")
    assert s[i + 2] == "指标定义与取" and s[i + 3] == "表1-1 指"   # 引表短行没有被图拆开


def test_check_flags_tables_dumped_to_front(tmp_path):
    f = build(tmp_path)
    doc = B.Doc(f)
    tbl = next(k for k in doc.body if k.tag == B.w("tbl"))
    sect = doc.body[-1]
    doc.body.remove(tbl)
    doc.body.remove(sect)
    doc.body.insert(0, sect)
    doc.body.insert(0, tbl)                              # 复现「按段落列表重排」的损坏形态
    doc.save()
    r = run("check", f)
    assert r.returncode == 2
    assert "sectPr 不在 body 末位" in r.stdout and "表格位于首个标题之前" in r.stdout


def test_same_second_backups_do_not_overwrite(tmp_path):
    f = build(tmp_path)
    for _ in range(3):
        B.Doc(f).save()
    assert len(list(tmp_path.glob("t.docx.bak-*"))) == 3


RULES = {
    "aliases": {"水网规划": "《海宁水网建设规划》", "水资源总体规划": "《水资源总体规划》"},
    "pdf_offset": {"《海宁水网建设规划》": 3},
    "overrides": {"S-05，资料包": ""},
    "drop_notes": ["拟议方法"],
}


@pytest.mark.parametrize("detail,want", [
    ("S-05，水网规划PDF第19页、印刷第16页", "（《海宁水网建设规划》第16页）"),
    ("S-05，水网规划PDF第13页", "（《海宁水网建设规划》第10页）"),                     # 仅 PDF 页 → 按已核偏移换算
    ("S-05，水网规划PDF第42页、印刷第39页；PDF第45页、印刷第42页", "（《海宁水网建设规划》第39页、第42页）"),
    ("S-01，招标需求印刷第6页；拟议方法", "（招标文件“招标需求”第6页）"),
    ("S-01，第二章“项目背景”；S-05，资料包", "（招标文件第二章“项目背景”）"),
])
def test_render(detail, want):
    errs = []
    assert B.render(detail, RULES, errs) == want and not errs


@pytest.mark.parametrize("detail", [
    "S-05",                                   # 裸号
    "S-05，某不认识的规划PDF第3页",              # 文件名无法识别
    "S-05，水资源总体规划PDF第8页",              # 只有 PDF 页且该文件偏移未核
    "S-03，案例参考，段617—624",                # 无页码无 override
])
def test_render_refuses_to_guess(detail):
    errs = []
    B.render(detail, RULES, errs)
    assert errs


def build_flat(tmp_path):
    d = Document()
    d.add_heading("7 技术路线", level=1)
    d.add_heading("7.1 现状调查", level=2)
    d.add_heading("7.1.1 思路", level=3)
    d.add_heading("1）样板条目", level=4)
    d.add_paragraph("样板正文" * 20)
    d.add_heading("7.2 实施方案", level=2)
    d.add_paragraph("节首说明" * 20)
    for i, name in enumerate(["方案框架", "规划约束", "供需平衡", "空间布局", "重点工程"], 1):
        d.add_heading(f"{i}）{name}", level=4)
        d.add_paragraph(f"{name}正文" * 15)
    d.add_paragraph("表7-1 指标")
    d.add_table(rows=2, cols=2)
    out = tmp_path / "flat.docx"
    d.save(out)
    return out


PLAN = """
sections:
  - h2: "7.2 实施方案"
    groups:
      - {title: "7.2.1 总体思路", n: 2, lead: "引导段一（招标文件第二章）。"}
      - title: "7.2.2 布局与工程"
        n: 3
        lead: "引导段二。"
        append:
          - {h4: "保障措施", paras: ["新增正文（招标文件第二章）。"]}
"""


def test_h3group_inserts_levels_and_renumbers(tmp_path):
    f = build_flat(tmp_path)
    plan = tmp_path / "plan.yaml"
    plan.write_text(PLAN, encoding="utf-8")
    r = run("check", f)
    assert r.returncode == 2 and "标题跳级: 1" in r.stdout
    r = run("h3group", f, "--plan", plan, "--apply")
    assert r.returncode == 0, r.stdout + r.stderr
    doc = B.Doc(f)
    heads = [(doc.level(k), B.ptext(k).strip()) for k in doc.body if doc.is_heading(k)]
    assert heads[4:] == [(2, "7.2 实施方案"), (3, "7.2.1 总体思路"), (4, "1）方案框架"), (4, "2）规划约束"),
                         (3, "7.2.2 布局与工程"), (4, "1）供需平衡"), (4, "2）空间布局"), (4, "3）重点工程"),
                         (4, "4）保障措施")]
    s = seq(f)
    assert s[s.index("TBL") - 1] == "表7-1 指" and s[-1] == "SECT"
    assert s[s.index("4）保障措施") + 1].startswith("新增正文")
    assert "标题跳级: 0" in run("check", f).stdout
    assert run("h3group", f, "--plan", plan, "--apply").returncode != 0      # 重复执行必须拒绝


def test_h3group_refuses_count_mismatch(tmp_path):
    f = build_flat(tmp_path)
    plan = tmp_path / "plan.yaml"
    plan.write_text(PLAN.replace("n: 3", "n: 2"), encoding="utf-8")
    before = f.read_bytes()
    assert run("h3group", f, "--plan", plan, "--apply").returncode != 0
    assert f.read_bytes() == before
