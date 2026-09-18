"""Application resources shared by source, installed, and frozen builds."""
from __future__ import annotations
import base64
from functools import lru_cache
from pathlib import Path
import sys


def app_icon_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "icon.ico"
    source_icon = Path(__file__).resolve().parent.parent / "icon.ico"
    if source_icon.is_file():
        return source_icon
    return Path(sys.prefix) / "share" / "nijika-print" / "icon.ico"


@lru_cache(maxsize=1)
def app_icon_data_url() -> str:
    path = app_icon_path()
    if not path.is_file():
        return ""
    return "data:image/x-icon;base64," + base64.b64encode(path.read_bytes()).decode("ascii")
