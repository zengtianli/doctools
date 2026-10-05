# DocKit 私有 Mac 组件 · 2026-10-05

本组件已在原 Chapter 固定流程真实完成；DocKit 产品线还要合并公开版结果，不能据本组件宣称整行通过。

- 仅接总部 `LaneSignal.swift`、实际 `loadOps()` 与从不上屏的 NSHostingView 离屏绘制就绪信号。默认交互与 doctools 业务未改；静默实例不使用本人偏好，不处理输入文档。源码23ee77c，总部 lifecycle 同步491e0d1。
- 第一次原 build.sh 的 vendor 改变输入，来源闸拒写 receipt，旧回执保留。同步固定来源后原增量构建/装机成功：1.0.1(323)，可执行文件21e6207d…、构建输入68860382…、实际verify=true。receipt 输入仍为原清单，无新平行构建实现。
- 首次原 sim_lane 静默检查0.723s、os_log、UIElement、window=false、focus=false、front_unchanged=true；证明后才登记 in_use:true。
- 正式test9348dc7743ea49c0af00a34045479413与当前code_key一致。固定验收 functionality bda78c6ee41849869ca8ae82a7efbf01、recovery54768551889242f08d73a55ec1f1b76f、privacy9d8ed4d2561b422ba743d58e875d45a4、native_ui2e3a8962906f4e189af2ae2978403734、cli_entry2c75b26dc98f449fb3095f1f9ff4d1ce全部通过。既有图标有效证据复用。
- perf e7f3e4d435e64575845f57ae1bc9b639实际通过：五次os_log首屏就绪[579,509,491,548,523]ms、中位数523ms；静置45秒后采60秒，50MiB、CPU0%；安装2,818,048 bytes。采样全程前台未变、无窗口与抢焦点。测量输入d45ce87b…，原件`~/Library/Caches/app-lightweight/measurement-candidates/323f847dccee476c9b96b446947cf08e.raw.json`、SHA a19f21b5…。
- 最终 `sop run --check-only --retry`：current_passed、delivery_state=complete、coverage=[]、gaps=[]。未推送、发布、重录或重做推广。
