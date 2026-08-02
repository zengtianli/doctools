"""钉死 shape_contract 的**缺省容差 = 0**（2026-08-02 立）。

## 为什么需要这条

`fix_styleset.py` 的 8 个写盘 gate 原本各自显式传一份「全 0」的 `allowed_deltas`
（67 行），把「零容忍」写在调用点上。实测那是纯 no-op（`diff_structure` 对
「列出且为 0」与「不列出」处理完全相同），已删。

删掉之后，零容忍**全靠 `diff_structure` 的隐式缺省兜着**。哪天有人把缺省放松成
「未列入 = 不检查」，这 8 道 gate 会**静默从 fail-closed 变 fail-open** ——
调用点上没有任何本地证据能提示这件事，代码 review 也看不出来。

所以缺省值本身必须被一条断言钉住。**这条测试是那 67 行删除的前提条件。**

⚠ 202 个原有用例一条都没覆盖 `shape_contract`，所以「套件全绿」不能替代本文件。
"""
import importlib.util
import sys
from pathlib import Path

import pytest

SUB = Path(__file__).resolve().parents[1] / "sub"


def _load():
    sys.path.insert(0, str(SUB))
    spec = importlib.util.spec_from_file_location("shape_contract", SUB / "shape_contract.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


sc = _load()


def _snap(**over):
    d = {f: 10 for f in sc._SCALAR_FIELDS}
    d["heading_counts"] = {"Heading 1": 3}
    d["figure_number_set"] = ["图1-1"]
    d["table_number_set"] = ["表1-1"]
    d.update(over)
    return d


@pytest.mark.parametrize("field", list(sc._SCALAR_FIELDS))
def test_scalar_drift_is_violation_without_allowed_deltas(field):
    """任一标量字段漂 1，不传 allowed_deltas 就必须报违规。"""
    before = _snap()
    after = _snap(**{field: 11})
    violations = sc.diff_structure(before, after)
    assert violations, f"{field} 漂了 +1 却没报违规 —— 缺省容差不再是 0，8 道 gate 已 fail-open"
    assert any(field in v for v in violations)


def test_heading_counts_drift_is_violation():
    before = _snap()
    after = _snap(heading_counts={"Heading 1": 4})
    assert sc.diff_structure(before, after), "heading_counts 漂了却没报违规"


def test_caption_set_drift_is_violation():
    before = _snap()
    after = _snap(figure_number_set=["图1-1", "图1-2"])
    assert sc.diff_structure(before, after), "figure_number_set 变了却没报违规"


def test_explicit_all_zero_is_a_noop():
    """显式传「全 0」与不传必须完全等价 —— 这是删掉那 67 行的依据本身。

    形状不是手抄的：直接从 `fix_styleset.py` 里当年那 8 份的并集取字段名，
    也就是 `_SCALAR_FIELDS` + heading_counts + 两个集合字段。
    """
    all_zero = {f: 0 for f in sc._SCALAR_FIELDS}
    all_zero.update({"heading_counts": 0, "figure_number_set": 0, "table_number_set": 0})
    for field in list(sc._SCALAR_FIELDS) + ["heading_counts"]:
        before = _snap()
        after = _snap(**({field: 11} if field != "heading_counts"
                         else {"heading_counts": {"Heading 1": 4}}))
        assert sc.diff_structure(before, after) == sc.diff_structure(
            before, after, allowed_deltas=all_zero), f"{field}: 传全 0 与不传不等价"


def test_nonzero_allowed_delta_still_tolerates():
    """反向：真给非零容差时必须**不**报违规 —— 否则本测试只是在测「永远报违规」。"""
    before = _snap()
    after = _snap(paragraph_count=11)
    assert not sc.diff_structure(before, after, allowed_deltas={"paragraph_count": 1})
