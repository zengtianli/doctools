#!/usr/bin/env python3
"""
消息显示模块
提供统一的控制台输出格式
"""


def show_success(msg: str):
    """显示成功消息"""
    print(f"✅ {msg}")


def show_error(msg: str):
    """显示错误消息"""
    print(f"❌ {msg}")


def show_warning(msg: str):
    """显示警告消息"""
    print(f"⚠️ {msg}")


def show_info(msg: str):
    """显示信息消息"""
    print(f"ℹ️ {msg}")


def show_processing(msg: str):
    """显示处理中消息"""
    print(f"🔄 {msg}")


def show_progress(current: int, total: int, item: str = ""):
    """显示进度"""
    percentage = (current * 100) // total if total > 0 else 0
    print(f"🔄 进度: {percentage}% ({current}/{total}) {item}")

