#!/usr/bin/env python3
"""
进度跟踪模块
"""

from display import show_info, show_processing


class ProgressTracker:
    """进度跟踪器"""

    def __init__(self, total: int = 0):
        self.success_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        self.total_count = 0
        self.current_count = 0
        self.total_expected = total

    def show(self, message: str = ""):
        """显示当前进度"""
        self.current_count += 1
        if self.total_expected > 0:
            percentage = (self.current_count * 100) // self.total_expected
            show_processing(f"进度: {percentage}% ({self.current_count}/{self.total_expected}) {message}")
        else:
            show_processing(f"处理中 ({self.current_count}): {message}")

    def add_success(self):
        """添加成功计数"""
        self.success_count += 1
        self.total_count += 1

    def add_failure(self):
        """添加失败计数"""
        self.failed_count += 1
        self.total_count += 1

    def add_skip(self):
        """添加跳过计数"""
        self.skipped_count += 1
        self.total_count += 1

    def show_summary(self, operation_name: str = "处理"):
        """显示统计摘要"""
        print()
        show_info(f"{operation_name}完成")
        print(f"✅ 成功: {self.success_count} 个")
        if self.failed_count > 0:
            print(f"❌ 失败: {self.failed_count} 个")
        if self.skipped_count > 0:
            print(f"⚠️ 跳过: {self.skipped_count} 个")
        print(f"📊 总计: {self.total_count} 个")

        if self.total_count > 0:
            success_rate = (self.success_count * 100) // self.total_count
            print(f"📊 成功率: {success_rate}%")

