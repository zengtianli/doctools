# `track compare` 落地全记录（从 CLAUDE.md 外迁 · 2026-08-04）

> 常驻结论（命令 + 退出码语义）在 `CLAUDE.md` §「track compare」；本页是空桩史与实现细节。

## 它此前是假装成功的空桩

`print("compare 功能将在 v2 实现。")`，parser 里却好好地声明着三个参数。前四道闸门
（surface / probe / function_axis / smoke）全绿 —— smoke 那条只跑 `track read`，因为
`track` 是 CMD_TABLE fast-path，**它的子命令根本不进指纹**（`cli_surface` 里 `track`
节点只有一个 `nargs="..."` 的 `rest`）。所以这类「二级子命令是空桩」在本仓的闸门体系里
是**结构性盲区**，只能靠真敲。

## 实现（2026-08-03 · 有回归门）

| | |
|---|---|
| 粒度 | **段落级**：改过的段 = 整段 `w:del` + 整段 `w:ins`，不做 run 级字符 diff |
| 范围 | body 顶层 `w:p`（difflib 对齐，键 = pStyle + 可见文本，`w:del` 内的字不参与） |
| 做法 | surgical：以**原稿**的 zip 为底只重写 `word/document.xml`，其余部件逐字节 verbatim + `assert_parts_intact` |
| 回归门 | `python3 -m pytest scripts/document/tests/test_track_compare.py`（13 条） |

**退出码不是只有 0**（照抄 `expect_rc=0` 会误判）：`0` 已产出修订件 · `1` 段落级无差异
（**不产出文件** —— 产一个没有修订标记的副本就是另一种假装成功）· `2` 输入不合法 /
一边 0 段落（空集不报绿）· `3` 已产出，但有本引擎标不了的**范围外差异**。

范围外 = 表格内 / 页眉页脚 / 脚注尾注 / 改稿新增段里引用改稿 rels 的图（那种图照搬进
原稿包 = Word 开门就报损坏，故摘掉）。这些一律打到 stderr 并把 rc 顶成 3，**不静默吞**。

删除侧复用总部引擎 `lib/docx_revise.tracked_delete_runs`，为它加了
`include_nontext=True`（默认值不变，历史调用方一个字节不受影响）：整段删除时**图片 run
也要包进 `w:del`**，否则接受修订后「段落没了图还在」，rc 照样 0。
