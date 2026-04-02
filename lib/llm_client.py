#!/usr/bin/env python3
"""统一 LLM 调用模块 — 通过 claude CLI 调用。

用法:
    from tools.llm_client import chat

    response = chat("你是助手", "帮我分析这个文件")

    # 指定模型 (sonnet/haiku/opus)
    response = chat("你是助手", "帮我分析", model="haiku")

也可直接运行测试:
    python3 llm_client.py
    python3 llm_client.py --model haiku
"""

import subprocess
import sys


def chat(
    system: str,
    message: str,
    provider: str | None = None,  # 保留参数兼容，忽略
    model: str | None = None,
    max_tokens: int | None = None,  # 保留参数兼容，忽略
) -> str:
    """Send a chat message via claude CLI and get a response.

    Args:
        system: System prompt
        message: User message
        provider: Ignored (kept for backward compatibility)
        model: Claude model alias (sonnet/haiku/opus, default: haiku)
        max_tokens: Ignored (kept for backward compatibility)

    Returns:
        Response text from Claude
    """
    cmd = ["claude", "-p", "--model", model or "haiku",
           "--system-prompt", system, message]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
        raise ConnectionError(
            f"claude CLI 失败 (code={result.returncode}): "
            f"{result.stderr.strip()[:200]}"
        )
    except FileNotFoundError:
        raise ConnectionError("claude CLI 未找到，请确认已安装 Claude Code")
    except subprocess.TimeoutExpired:
        raise ConnectionError("claude CLI 超时 (120s)")


# ---------------------------------------------------------------------------
# CLI test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="LLM Client 测试")
    parser.add_argument("--model", default="haiku", help="模型 (sonnet/haiku/opus)")
    parser.add_argument("--message", default="你好，请用一句话介绍自己。", help="测试消息")
    args = parser.parse_args()

    print(f"测试 LLM 调用 (model={args.model})...")
    try:
        result = chat(
            system="你是一个测试助手。简短回复。",
            message=args.message,
            model=args.model,
        )
        print(f"响应: {result}")
    except Exception as e:
        print(f"失败: {e}", file=sys.stderr)
        sys.exit(1)
