"""开发时和打包成 exe 后，文件路径不一样，这里统一处理。"""

from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def bundle_dir() -> Path:
    if is_frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def app_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent
