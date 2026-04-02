#!/usr/bin/env python3
"""
使用统计日志模块
"""

import csv
import os
from datetime import datetime

from env import USAGE_LOG


def log_usage(script_name: str, category: str = ""):
    """记录脚本使用统计"""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not os.path.exists(USAGE_LOG):
            with open(USAGE_LOG, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['timestamp', 'script_name', 'category'])

        with open(USAGE_LOG, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([timestamp, script_name, category])
    except Exception:
        pass
