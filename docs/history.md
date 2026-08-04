# 结构史与实证注记（从 CLAUDE.md 外迁 · 2026-08-04）

> 本页收「已经发生完、平时不用背着走」的历史与实证细节。改到对应部位时回来查。
> 大盘改造史见 `handoffs/_archive/2026-08-04-docx-refactor-roadmap.md`（P1–P12 全销项）
> 与同目录 `…-before-after.md`。

## 折叠史（2026-07-31 三连）

- **家族折叠**：strip/audit/freeze/image/table/caption/chapter/renumber/blocks/split/
  typeset_ops/biddiff 12 族 42 旧件并成 `sub/` 里 12 个子命令族文件。
- **入口层折叠**：chapter_renumber/tabfig_align/docx_renumber_figures → `renum.py`；
  docx_apply_template/format_clone/font_normalize/text_formatter → `docx_fmt.py`；
  docx_apply_image_caption 平移进 `sub/`。
- **同日再折**：pptx 4→1（`pptx_cli`）· pdf 3→1（`pdf_cli`）· md_docx_template 并入
  `md_tools`（子命令 `md2docx`）。
- **`md_tools to-docx`（pandoc 版）2026-08-04 用户拍板退役**：绕开 docx-gen-guard 字面
  匹配（守卫旁路，路线图 P11 挂账项）+ 实测全生态零消费者；md→docx 统一走 `md2docx`。
  docx_cli 的 `md-to-docx` 动词本来就转发 `md2docx`，不受影响。
- `_groups.py` 之前是 20 个只做「argparse 声明 + `exec_script` 转发」的 group 模块
  （2265 行），2026-07-30 折成一张表。

## 入口与版本的实证

- 两条入口等价：实测同一条 `audit headings` 两边 stdout 逐字节相同。
- 版本号 SSOT 实证：把 `__version__` 改成 `9.9.9` 后，包元数据 / `doctools --version` /
  老路径 `--version` 三处一起变；改成非 PEP440 的 `9.9.9-probe` 会让 hatchling 构建
  直接失败（说明它真在读这个文件）。
- `--version` 若注册成 argparse 参数：实测 surface diff 立刻转红并点出 `_VersionAction`
  那一节 —— 要加就得先决定是否更新基线，别顺手塞。
- src-layout 是必需的：包放仓根时 editable 安装会把**整个仓根**写进共享 venv 的 `.pth`，
  凭空多出 lib/scripts/tools/config 等顶层名。

## 闸门抓过的真差异（留作「为什么要跑它」的证据）

- `cli_surface` 加上「子命令组的 dest/required/metavar」之后，当场抓到 `header-footer`
  的 metavar 被从 `<action>` 写成了 `<target>` —— 这种差异不报错、只在 `--help` 里
  显示不同。
- `script_graph` 动词轴 2026-08-03 前是 fail-open：`verb_map()` 载入失败只 print 一行 ⚠、
  rc 照样 0 —— 实测注入一句 ImportError 后仍输出「106 脚本 · 374 引用 · 0 孤儿」rc=0，
  而三视图里整整一视图已空。现已改 `exit 2`（反向验证：注入 rc=2 / 正常 rc=0）。

## `script_graph` 实现里的两个坑（第一版都踩过）

「实现脚本」那一列不是手抄的，三条来源都读 SSOT 本体：

- `Target.chain` 是 `(dest, 同组 target 名)`（**不是脚本名**，照字面收会把 `table borders`
  报成实现在 `center.py`）。
- 扁平组的 Target 挂在 `g.flat` 不在 `g.targets`（漏掉这支会把 `chrome` / `md-merge`
  报成实现在 `_groups.py`）。

交叉核验：与 `cli_forward_probe` 实录的转发脚本比对，65 条重合项 0 不一致。
清单里的**格式轴那一列是信号计数的启发式**（`FORMAT_SIG`），与边和动词映射不同级，
别拿它当事实用；页面顶部也这么标了。

## Raycast（2026-07-27 整体归档）

`raycast/` 整子树已进 `~/.Trash/dead-scripts-20260727/`：`commands/` 早就是空的，
9 个 `doc_*.sh` 在它自己的 `_archive/` 里，只剩一个没人 source 的 `lib/run_python.sh`。
现在的入口是直接命令行（`/docx` skill 已退役）。

## surgical 收口的由来（2026-07-30）

一份 301 部件 / 75 个嵌入公式的真报告，python-docx **什么都不改**地打开再存回，
**60 个部件字节变了、只有 1 个语义真变了**（59 个纯属被重新序列化）。收口把语义未变的
部件按原字节还原 → 炸开面 60→1，实测「无改动时输出与原件逐字节相同」。
