# DocKit

[English](README_EN.md)

[doctools](../README.md) 的 SwiftUI 桌面端，操作与选项从同一文档后端 `gui-ops` 加载。源码、文档引擎与构建由父仓统一维护；本客户端依赖本机 `~/Dev/.venv`，不是独立分发版。

```sh
./build.sh --check    # CodingKeys + 真实后端解码
./build.sh            # Release 构建、签名，不装机
./build.sh --install  # 明确要求时安装到 /Applications/DocKit.app
```

显示名取自 catalog，bundle ID 保持 `cyou.tianli.DocTools`。构建使用共享 Xcode 选择器。JSON 与客户端开发约定见 [CLAUDE.md](CLAUDE.md)。
