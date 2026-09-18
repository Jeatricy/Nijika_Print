"""Safe, hidden WebView2 startup diagnostics for the real frontend bridge."""
from __future__ import annotations
import threading
import time
from pathlib import Path
import webview
from .backend import DesktopAPI
from .resources import app_icon_path


class _ReadOnlyAPI(DesktopAPI):
    def __init__(self):
        super().__init__()
        self._bootstrap_calls = 0
        self._blocked_prints = 0

    def bootstrap(self):
        self._bootstrap_calls += 1
        return super().bootstrap()

    def start_print(self, settings):
        self._blocked_prints += 1
        return {"error": "诊断模式禁止提交打印任务。"}

    def choose_folder(self):
        return None

    def choose_files(self):
        return None


_CAPTURE_ERRORS = r"""
(() => {
  window.__nijikaStartupErrors = [];
  const record = value => window.__nijikaStartupErrors.push(String(value));
  window.addEventListener('error', e => record(e.message || 'JavaScript error'));
  window.addEventListener('unhandledrejection', e => record(e.reason));
  const original = console.error.bind(console);
  console.error = (...args) => { record(args.join(' ')); original(...args); };
  return true;
})()
"""

_SNAPSHOT = r"""
(() => {
  const el = id => document.getElementById(id);
  const image = el('app-icon');
  const printer = typeof selectModels === 'undefined' ? null : selectModels.get('printer-select');
  return {
    ready: typeof state !== 'undefined' && state.ready === true,
    bridge_ready: !!window.pywebview?.api,
    status: el('status-text')?.textContent || '',
    dialog_open: !!el('app-dialog')?.open,
    dialog_error: el('app-dialog')?.open ? el('dialog-body')?.textContent || '' : '',
    settings_initialized: !!printer && printer.placeholder !== '正在查找打印机…',
    icon_loaded: !!image && image.complete && image.naturalWidth > 0,
    controls_present: ['source-tree', 'queue-list', 'nup-input', 'duplex-select',
      'layout-select', 'office-select', 'print-button'].every(id => !!el(id)),
    errors: window.__nijikaStartupErrors || []
  };
})()
"""


def frontend_startup_check(web_root: Path, timeout: float = 45.0) -> dict:
    """Run actual HTML/JS against the Python API without showing any window."""
    api = _ReadOnlyAPI()
    result = {"ok": False, "error": "Frontend startup timed out", "snapshot": {}}
    finished = threading.Event()
    started = threading.Event()
    window = webview.create_window(
        "Nijika Print background diagnostics", url=str(web_root / "index.html"),
        js_api=api, width=1380, height=900, hidden=True, focus=False,
        frameless=True, easy_drag=False, shadow=False,
    )
    api._attach(window)

    def close():
        finished.set()
        try:
            window.destroy()
        except Exception:
            pass

    def inspect():
        if started.is_set():
            return
        started.set()
        stable_since = None
        deadline = time.monotonic() + timeout
        try:
            window.evaluate_js(_CAPTURE_ERRORS)
            while time.monotonic() < deadline and not finished.is_set():
                snapshot = window.evaluate_js(_SNAPSHOT)
                result["snapshot"] = snapshot
                if snapshot.get("dialog_open"):
                    result["error"] = snapshot.get("dialog_error") or "Unexpected startup dialog"
                    break
                if snapshot.get("errors"):
                    result["error"] = snapshot["errors"][0]
                    break
                ready = (snapshot.get("ready") and snapshot.get("bridge_ready")
                         and snapshot.get("status") == "准备就绪"
                         and snapshot.get("settings_initialized")
                         and snapshot.get("icon_loaded") and snapshot.get("controls_present")
                         and api._bootstrap_calls > 0 and api._blocked_prints == 0)
                if ready:
                    stable_since = stable_since or time.monotonic()
                    if time.monotonic() - stable_since >= 1.2:
                        result.update(ok=True, error="")
                        break
                else:
                    stable_since = None
                time.sleep(0.1)
        except Exception as exc:
            result["error"] = str(exc)
        finally:
            close()

    def watchdog():
        if not finished.wait(timeout + 5):
            result["error"] = "WebView2 did not finish frontend startup within the timeout"
            close()

    window.events.loaded += inspect
    watch = threading.Thread(target=watchdog, daemon=True)
    icon = app_icon_path()
    webview.start(watch.start, gui="edgechromium", debug=False, private_mode=True,
                  icon=str(icon) if icon.is_file() else None)
    result["bootstrap_calls"] = api._bootstrap_calls
    result["blocked_print_attempts"] = api._blocked_prints
    return result
