# DocKit

[English](README_EN.md)

[doctools](../README.md) 的 SwiftUI 桌面端，操作与选项从同一文档后端 `gui-ops` 加载。源码、文档引擎与构建由父仓统一维护；本客户端依赖本机 `~/Dev/.venv`，不是独立分发版。

```sh
./build.sh --check    # CodingKeys + 真实后端解码
./build.sh            # Release 构建、签名，不装机
./build.sh --install  # 明确要求时安装到 /Applications/DocKit.app
```

显示名取自 catalog，bundle ID 保持 `cyou.tianli.DocTools`。构建使用共享 Xcode 选择器。JSON 与客户端开发约定见 [CLAUDE.md](CLAUDE.md)。

固定验收在 `scripts/accept/`，通过 Chapter 运行并由其写入证据：

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py accept --app doc-tools-doctools --check functionality --check recovery --check privacy --check native_ui --json
```

功能与恢复测试只使用临时文档。原生界面验收会构建当前源码，并调用 App 的 `--ui-self-test` 离屏渲染真实视图；不装机、不抢焦点，也不操作剪贴板。它证明本地构建的行为，装机图标仍需本人在 Chapter 确认。

规范化、引号等操作由本机文档引擎执行。“敏感词扫描”使用外部 Claude 服务，会发送扫描目录中文档的文件名与正文；每次选择该操作时须明确勾选外发同意。验收只检查拒绝路径与本地处理，不发送文档，也不证明外部服务可用。

<!-- lightweight:start -->
## 资源占用

| 安装后占用 | 空闲内存 | 空闲 CPU | 读取操作列表（生产后端，含 Python 进程启动） |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

原生 SwiftUI 客户端，文档能力复用本机共享 Python 引擎；空闲时后端进程退出。

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · 已安装自用版首页，载入 14 项真实操作与动态选项；未打开或处理用户文档。 · 2026-09-26。数字来自所列设备实测，版本更新后重新测量。内存口径为 phys_footprint；CPU 为 60 秒采样窗内 CPU 时间 ÷ 墙钟；大小按十进制 MB。原始数据见 [perf/lightweight.json](perf/lightweight.json)。</sub>
<!-- lightweight:end -->
