"""Desktop window and native integration for Nijika Print."""
from __future__ import annotations
import ctypes
import importlib
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
import traceback

import webview
from webview.dom import DOMEventHandler
from .backend import DesktopAPI
from .branding import APP_NAME, APP_VERSION
from .resources import app_icon_path

WEB_ROOT = Path(__file__).resolve().parent / "web"


def run():
    if sys.platform == "win32":
        try:
            set_app_id = ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID
            set_app_id.argtypes = [ctypes.c_wchar_p]
            set_app_id.restype = ctypes.c_long
            set_app_id("Nijika.Print.Desktop")
        except (AttributeError, OSError):
            pass
    api = DesktopAPI()
    webview.settings["SHOW_DEFAULT_MENUS"] = False
    webview.settings["ALLOW_DOWNLOADS"] = False
    window = webview.create_window(
        APP_NAME, url=str(WEB_ROOT / "index.html"), js_api=api,
        width=1380, height=900, min_size=(1080, 740), frameless=True,
        easy_drag=False, shadow=True, background_color="#faf9f6", text_select=False,
    )
    api._attach(window)
    window.events.closing += api._can_close

    def on_loaded():
        def on_drop(target):
            def handler(event):
                try:
                    files = event.get("dataTransfer", {}).get("files", [])
                    paths = [file["pywebviewFullPath"] for file in files if file.get("pywebviewFullPath")]
                    if not paths:
                        result = {"error": "未能读取文件路径，请使用“选择文件夹”或“添加文件”。"}
                    elif target == "source":
                        folders = [path for path in paths if Path(path).is_dir()]
                        result = api.select_root(folders[0]) if folders else {"error": "左侧只接收文件夹；请将文件拖入右侧队列。"}
                    else:
                        result = api.queue_add(paths)
                    window.evaluate_js(f"window.desktop.onNativeDrop({json.dumps(target)}, {json.dumps(result, ensure_ascii=True)}).catch(window.desktop.reportError); true")
                except Exception as exc:
                    window.evaluate_js(f"window.desktop.reportError({json.dumps(str(exc))})")
            return handler
        for target, selector in (("source", "#source-drop"), ("queue", "#queue-drop")):
            window.dom.get_element(selector).on("drop", DOMEventHandler(on_drop(target), prevent_default=True, stop_propagation=True))

    window.events.loaded += on_loaded
    icon = app_icon_path()
    webview.start(gui="edgechromium", debug=False, private_mode=True,
                  icon=str(icon) if icon.is_file() else None)


def diagnose(output):
    """Check dependencies and real frontend startup without showing a window or printing."""
    from .engine import PrintEngine
    from . import engine as printing
    core = PrintEngine()
    result = {"name": APP_NAME, "version": APP_VERSION, "frozen": bool(getattr(sys, "frozen", False)),
              "win32": core.win32_ok, "pdf": core.have_fitz, "sumatra": bool(core.sumatra_path),
              "web_assets": all((WEB_ROOT / name).is_file() for name in ("index.html", "app.css", "app.js", "icons.js")),
              "tk_loaded": "tkinter" in sys.modules, "pdf_nup": False}
    if core.have_fitz:
        with tempfile.TemporaryDirectory(prefix="nijika-diagnose-") as directory:
            src, dst = Path(directory) / "input.pdf", Path(directory) / "nup.pdf"
            document = printing.fitz.open()
            for i in range(2):
                page = document.new_page()
                page.insert_text((72, 72), f"Diagnostic page {i + 1}")
            document.save(src)
            document.close()
            if core._create_pdf_nup_py(str(src), 2, str(dst), rows=1, cols=2):
                check = printing.fitz.open(dst)
                result["pdf_nup"] = check.page_count == 1
                check.close()
    icon = app_icon_path()
    result["app_icon"] = icon.is_file()
    result["app_icon_sha256"] = hashlib.sha256(icon.read_bytes()).hexdigest() if icon.is_file() else None
    result["webview_backend"] = False
    try:
        importlib.import_module("webview.platforms.edgechromium")
        importlib.import_module("webview.platforms.winforms")
        result["webview_backend"] = True
    except Exception as exc:
        result["webview_error"] = str(exc)
    result["ok"] = result["app_icon"] and result["webview_backend"] and result["win32"] and result["pdf"] and result["web_assets"] and result["pdf_nup"] and not result["tk_loaded"]
    if result["ok"]:
        from .diagnostics import frontend_startup_check
        result["frontend"] = frontend_startup_check(WEB_ROOT)
        result["ok"] = result["ok"] and result["frontend"]["ok"]
    Path(output).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if result["ok"] else 1


def main():
    if len(sys.argv) == 3 and sys.argv[1] == "--diagnose":
        raise SystemExit(diagnose(sys.argv[2]))
    try:
        run()
    except Exception:
        logs = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / APP_NAME / "logs"
        logs.mkdir(parents=True, exist_ok=True)
        log = logs / "startup-error.log"
        log.write_text(traceback.format_exc(), encoding="utf-8")
        if sys.platform == "win32":
            ctypes.windll.user32.MessageBoxW(None, f"无法启动 {APP_NAME} {APP_VERSION}。请检查 WebView2 运行时和应用依赖。\n\n错误详情：{log}", f"{APP_NAME} · 启动失败", 0x10)
        raise



