# -*- mode: python ; coding: utf-8 -*-
"""Source-only build configuration for the standalone Nijika Print 1.0 release."""
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

ROOT = Path(SPECPATH).resolve()
sumatra = Path(os.environ.get("NIJIKA_SUMATRA_BUILD_PATH", ""))
if not sumatra.is_file():
    raise SystemExit("Set NIJIKA_SUMATRA_BUILD_PATH to an external SumatraPDF 3.6.1 executable before building.")

metadata = []
for distribution in ("pywebview", "pythonnet", "clr_loader", "PyMuPDF", "pywin32", "pypinyin"):
    metadata += copy_metadata(distribution)

a = Analysis(
    [str(ROOT / "nijika_print" / "__main__.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        (str(ROOT / "nijika_print" / "web"), "nijika_print/web"),
        (str(sumatra), "."),
        (str(ROOT / "THIRD_PARTY_NOTICES.md"), "."),
    ] + collect_data_files("webview") + metadata,
    hiddenimports=["webview.platforms.edgechromium", "webview.platforms.winforms"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter", "customtkinter", "tkinterdnd2", "PySide6", "PySide2", "PyQt6", "PyQt5",
        "matplotlib", "IPython", "pytest", "numpy", "pandas", "scipy", "pyarrow",
        "openpyxl", "sqlalchemy", "psycopg", "psycopg_binary", "cv2", "torch",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="Nijika_Print-1.0.0-windows-x64",
    debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False,
    disable_windowed_traceback=False,
    icon="NONE",
    version=str(ROOT / "packaging" / "version_info.txt"),
)
