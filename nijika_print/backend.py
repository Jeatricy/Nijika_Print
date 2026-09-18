"""Thread-safe local file/queue/print bridge exposed to our bundled HTML only."""
from __future__ import annotations
import os
from pathlib import Path
import queue
import stat
import threading
import uuid
from .engine import PrintEngine
from .resources import app_icon_data_url
from . import engine as printing


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def is_linklike(path):
    try:
        return bool(getattr(os.lstat(path), "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)) or os.path.islink(path)
    except OSError:
        return True


def file_info(path):
    path = os.path.abspath(path)
    is_dir = os.path.isdir(path)
    extension = Path(path).suffix.lower()
    try:
        size = os.path.getsize(path) if not is_dir else 0
    except OSError:
        size = 0
    return {"path": path, "name": os.path.basename(path) or path, "folder": os.path.dirname(path),
            "isDir": is_dir, "extension": extension, "size": size,
            "supported": not is_dir and extension in printing.SUPPORTED_EXTS,
            "children": None if is_dir else []}


class DesktopAPI:
    def __init__(self, engine=None):
        self._engine = engine or PrintEngine()
        self._window = None
        self._lock = threading.RLock()
        self._files = []
        self._revision = 0
        self._root = ""
        self._thread = None
        self._done_received = False
        self._close_when_done = False
        self._job = {"id": "", "state": "idle", "status": "准备就绪", "processed": 0, "total": 0,
                     "failed": [], "success": 0, "cancelled": 0}

    def _attach(self, window):
        self._window = window

    def _queue_snapshot(self):
        return {"files": [file_info(path) for path in self._files], "revision": self._revision}

    def bootstrap(self):
        with self._lock:
            return {"app_icon": app_icon_data_url(), "queue": self._queue_snapshot(), "printers": self.get_printers(),
                    "capabilities": {"printing": self._engine.win32_ok, "pdf": self._engine.have_fitz,
                                     "sumatra": bool(self._engine.sumatra_path)}, "root": self._root, "job": dict(self._job)}

    def get_printers(self):
        if not self._engine.win32_ok:
            return {"names": [], "default": "", "error": "缺少 pywin32 打印组件"}
        try:
            flags = printing.win32print.PRINTER_ENUM_LOCAL | printing.win32print.PRINTER_ENUM_CONNECTIONS
            names = [p[2] for p in printing.win32print.EnumPrinters(flags)]
            try:
                default = printing.win32print.GetDefaultPrinter()
            except Exception:
                default = names[0] if names else ""
            return {"names": names, "default": default if default in names else (names[0] if names else ""),
                    "error": "" if names else "未发现打印机"}
        except Exception as exc:
            return {"names": [], "default": "", "error": str(exc)}

    def choose_folder(self):
        import webview
        result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
        return os.path.abspath(result[0]) if result else None

    def choose_files(self):
        import webview
        patterns = ";".join("*" + ext for ext in sorted(printing.SUPPORTED_EXTS))
        result = self._window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True,
                                                file_types=(f"可打印文件 ({patterns})", "所有文件 (*.*)"))
        return self.queue_add(list(result)) if result else None

    def browse(self, path, filter_text="", descending=False):
        if not isinstance(path, str) or not os.path.isdir(path):
            return {"error": "这个文件夹不存在，或已被移动。", "nodes": []}
        root = os.path.abspath(path)
        keyword = str(filter_text).strip().casefold()
        errors = []
        def read(folder):
            try:
                with os.scandir(folder) as entries:
                    paths = [entry.path for entry in entries]
            except OSError as exc:
                errors.append(str(exc))
                return []
            folders, files = [], []
            for full in paths:
                info = file_info(full)
                if info["isDir"]:
                    # Do not recursively follow directory symlinks/junctions.
                    if keyword:
                        if is_linklike(full):
                            continue
                        info["children"] = read(full)
                        if not info["children"]:
                            continue
                    folders.append(info)
                elif not keyword or keyword in info["name"].casefold():
                    files.append(info)
            key = lambda item: printing._natural_sort_key(item["name"])
            return sorted(folders, key=key, reverse=bool(descending)) + sorted(files, key=key, reverse=bool(descending))
        nodes = read(root)
        return {"root": root, "name": os.path.basename(root) or root, "nodes": nodes,
                "error": errors[0] if errors else ""}

    def select_root(self, path):
        if not isinstance(path, str) or not os.path.isdir(path):
            return {"error": "请拖入文件夹；单个文件请放入右侧队列。"}
        self._root = os.path.abspath(path)
        return {"root": self._root}

    def queue_add(self, paths):
        if not isinstance(paths, (list, tuple)):
            raise ValueError("文件路径应为列表")
        candidates = []
        skipped = 0
        for raw in paths:
            if not isinstance(raw, str):
                skipped += 1
                continue
            path = os.path.abspath(raw)
            if os.path.isdir(path):
                for folder, dirs, files in os.walk(path, followlinks=False):
                    dirs[:] = sorted((d for d in dirs if not is_linklike(os.path.join(folder, d))), key=printing._natural_sort_key)
                    for name in sorted(files, key=printing._natural_sort_key):
                        if Path(name).suffix.lower() in printing.SUPPORTED_EXTS:
                            candidates.append(os.path.join(folder, name))
            else:
                candidates.append(path)
        with self._lock:
            existing = {path_key(path) for path in self._files}
            added = 0
            for path in candidates:
                key = path_key(path)
                if key in existing or not os.path.isfile(path) or Path(path).suffix.lower() not in printing.SUPPORTED_EXTS:
                    skipped += 1
                    continue
                self._files.append(path)
                existing.add(key)
                added += 1
            self._revision += 1
            return {**self._queue_snapshot(), "added": added, "skipped": skipped}

    def queue_remove(self, paths):
        keys = {path_key(path) for path in paths}
        with self._lock:
            self._files = [path for path in self._files if path_key(path) not in keys]
            self._revision += 1
            return self._queue_snapshot()

    def queue_clear(self):
        with self._lock:
            self._files.clear()
            self._revision += 1
            return self._queue_snapshot()

    def queue_move(self, paths, direction):
        if direction not in (-1, 1):
            raise ValueError("无效的移动方向")
        selected = {path_key(path) for path in paths}
        with self._lock:
            indices = [i for i, path in enumerate(self._files) if path_key(path) in selected]
            moving, blocked = set(indices), set()
            for i in (indices if direction < 0 else reversed(indices)):
                neighbor = i + direction
                if neighbor < 0 or neighbor >= len(self._files) or (neighbor in moving and neighbor in blocked):
                    blocked.add(i)
                    continue
                self._files[i], self._files[neighbor] = self._files[neighbor], self._files[i]
            self._revision += 1
            return self._queue_snapshot()

    def start_print(self, settings):
        with self._lock:
            if self._thread and self._thread.is_alive():
                return {"error": "已有打印任务正在进行。"}
            if not self._files:
                return {"error": "请先将文件加入待打印队列。"}
            if not self._engine.win32_ok:
                return {"error": "缺少 Windows 打印组件，请安装 requirements.txt。"}
            printer = str(settings.get("printer", ""))
            if printer not in self.get_printers()["names"]:
                return {"error": "请选择当前可用的打印机。"}
            try:
                nup = int(str(settings.get("nup", "")))
            except (ValueError, TypeError):
                nup = 0
            if not 1 <= nup <= 16:
                return {"error": "PDF 每张页数应为 1～16 的整数。"}
            missing = [path for path in self._files if not os.path.isfile(path)]
            if missing:
                return {"error": f"有 {len(missing)} 个文件已移动或删除，请先移出队列。"}
            duplex = settings.get("duplex", "单面")
            office = settings.get("office", "自动")
            layout = str(settings.get("layout", "自动"))
            valid_layouts = ["自动"] if nup == 1 else (["自动", "Vertical（竖排）", "Horizontal（并排）"] if nup == 2 else (["自动（行优先）", "列优先"] if nup == 4 else ["自动（行优先）"]))
            if duplex not in printing.DUPLEX_OPTIONS or office not in printing.OFFICE_MODE_OPTIONS or layout not in valid_layouts:
                return {"error": "打印参数无效，请重新选择。"}
            job = {"files": list(self._files), "printer": printer, "color": bool(settings.get("color", True)),
                   "duplex": printing.DUPLEX_OPTIONS[duplex], "pdf_nup": nup,
                   "nup_layout": layout, "office_mode": office}
            self._engine.failed_files.clear()
            self._engine.stop_print.clear()
            self._engine.print_queue = queue.Queue()
            self._done_received = False
            self._job = {"id": uuid.uuid4().hex, "state": "running", "status": "正在准备打印…",
                         "processed": 0, "total": len(job["files"]), "failed": [], "success": 0, "cancelled": 0}
            self._thread = threading.Thread(target=self._run_job, args=(job,), daemon=False)
            self._thread.start()
            return {"job": dict(self._job)}

    def _run_job(self, job):
        try:
            self._engine._print_worker(job)
        except Exception as exc:
            self._engine.failed_files.append(("打印任务", str(exc)))
            self._engine.print_queue.put(f"任务异常：{exc}")
        finally:
            self._engine.print_queue.put("__DONE__")

    def poll_job(self):
        with self._lock:
            while True:
                try:
                    message = self._engine.print_queue.get_nowait()
                except queue.Empty:
                    break
                if message == "__PROGRESS__":
                    self._job["processed"] = min(self._job["processed"] + 1, self._job["total"])
                elif message == "__DONE__":
                    self._done_received = True
                else:
                    self._job["status"] = message
            if self._done_received and not (self._thread and self._thread.is_alive()):
                self._job["failed"] = [{"name": os.path.basename(path), "reason": reason}
                                       for path, reason in self._engine.failed_files]
                self._job["success"] = max(0, self._job["processed"] - len(self._job["failed"]))
                self._job["cancelled"] = max(0, self._job["total"] - self._job["processed"])
                self._job["state"] = "cancelled" if self._engine.stop_print.is_set() else "done"
            return dict(self._job)

    def cancel_print(self):
        with self._lock:
            if self._thread and self._thread.is_alive():
                self._job["cancelRequested"] = True
                self._engine.stop_print.set()
        return {"status": "正在停止后续提交；已提交的任务请在系统打印队列中取消。"}

    def window_action(self, action):
        if not self._window:
            return False
        if action == "minimize":
            self._window.minimize()
        elif action == "maximize":
            if getattr(self._window, "_desktop_maximized", False):
                self._window.restore()
                self._window._desktop_maximized = False
            else:
                self._window.maximize()
                self._window._desktop_maximized = True
        elif action == "close":
            if self._thread and self._thread.is_alive():
                return {"busy": True}
            self._window.destroy()
        return {"busy": False}

    def stop_and_close(self):
        self.cancel_print()
        self._close_when_done = True
        def wait():
            if self._thread:
                self._thread.join()
            if self._window:
                self._window.destroy()
        threading.Thread(target=wait, daemon=True).start()
        return True

    def _can_close(self):
        if self._thread and self._thread.is_alive():
            if not self._close_when_done:
                self._window.evaluate_js("window.desktop.requestClose()")
            return False
        return True


