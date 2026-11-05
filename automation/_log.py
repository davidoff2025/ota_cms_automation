"""共享终端着色（绿=成功/一致，黄=警告，红=错误）。

仅当输出是 TTY 且未设 NO_COLOR 时上色；非交互/重定向/管道时自动退化为无色纯文本，
避免把 ANSI 码写进被下游捕获的输出。实现照搬 automation/prep.py 原有逻辑，供
prep / dass_* / active_ota 复用（纯标准库，不引第三方依赖）。
"""
from __future__ import annotations

import os
import sys

# import 时求值一次（与 prep 原行为一致，不改成调用时求值）。
_USE_COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None
_COLORS = {"green": "\033[32m", "yellow": "\033[33m", "red": "\033[31m",
           "bold": "\033[1m", "dim": "\033[2m"}
_RESET = "\033[0m"


def color(text: str, name: str) -> str:
    """把 text 包成指定颜色；非 TTY/未知颜色时原样返回（纯文本）。"""
    if not _USE_COLOR or name not in _COLORS:
        return text
    return f"{_COLORS[name]}{text}{_RESET}"


def ok(msg: str) -> str:      # 绿色 ✓
    return color(f"✓ {msg}", "green")


def warn(msg: str) -> str:    # 黄色 ⚠
    return color(f"⚠ {msg}", "yellow")


def err(msg: str) -> str:     # 红色 ✗
    return color(f"✗ {msg}", "red")
