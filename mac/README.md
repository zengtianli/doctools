# DocKit

[English](README_EN.md)

[doctools](../README.md) 的 SwiftUI 桌面端，操作与选项从同一文档后端 `gui-ops` 加载。源码、文档引擎与构建由父仓统一维护；本客户端依赖本机 `~/Dev/.venv`，不是独立分发版。

```sh
./build.sh --check    # CodingKeys + 真实后端解码
./build.sh            # Release 构建、签名，不装机
./build.sh --install  # 明确要求时安装到 /Applications/DocKit.app
```

显示名取自 catalog，bundle ID 保持 `cyou.tianli.DocTools`。构建使用共享 Xcode 选择器。JSON 与客户端开发约定见 [CLAUDE.md](CLAUDE.md)。

<!-- lightweight:start -->
## 资源占用

| 安装后占用 | 空闲内存 | 空闲 CPU | 读取操作列表（生产后端，含 Python 进程启动） |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

原生 SwiftUI 客户端，文档能力复用本机共享 Python 引擎；空闲时后端进程退出。

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · 已安装自用版首页，载入 14 项真实操作与动态选项；未打开或处理用户文档。 · 2026-09-26。数字来自所列设备实测，版本更新后重新测量。内存口径为 phys_footprint；CPU 为 60 秒采样窗内 CPU 时间 ÷ 墙钟；大小按十进制 MB。原始数据见 [perf/lightweight.json](perf/lightweight.json)。</sub>
<!-- lightweight:end -->
