"""Nijika Print's standalone, UI-independent Windows printing engine."""
from __future__ import annotations

import functools
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from typing import List, Optional, Tuple

WIN32_OK = False
win32print = win32api = win32con = pythoncom = None
Dispatch = DispatchEx = None
try:
    if sys.platform == "win32":
        import win32print
        import win32api
        import win32con
        import pythoncom
        from win32com.client import Dispatch, DispatchEx
        WIN32_OK = True
except ImportError:
    pass

HAVE_FITZ = False
try:
    import fitz
    HAVE_FITZ = True
except ImportError:
    fitz = None

HAVE_PYPINYIN = False
try:
    from pypinyin import lazy_pinyin
    HAVE_PYPINYIN = True
except ImportError:
    lazy_pinyin = None

SUPPORTED_EXTS = {
    ".doc", ".docx",
    ".xls", ".xlsx",
    ".ppt", ".pptx",
    ".pdf",
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp",
}

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}

DUPLEX_OPTIONS = {
    "单面": 1,
    "长边装订": 2,
    "短边装订": 3,
}

OFFICE_MODE_OPTIONS = ["自动", "优先 WPS", "优先 Excel"]

@functools.lru_cache(maxsize=65536)
def _natural_sort_key(s: str):
    # 排序时每个名字会被比较多次，拼音转换较慢，做缓存
    if HAVE_PYPINYIN:
        text_for_sort = ''.join(lazy_pinyin(s)).lower()
    else:
        text_for_sort = s.lower()
    return tuple((0, int(t)) if t.isdigit() else (1, t) for t in re.split(r'(\d+)', text_for_sort))

class PrintEngine:
    """One immutable job snapshot at a time; no UI or legacy module dependency."""

    def __init__(self):
        self.print_queue = queue.Queue()
        self.stop_print = threading.Event()
        self.failed_files = []
        self.win32_ok = WIN32_OK
        self.office_ok = WIN32_OK
        self.have_fitz = HAVE_FITZ
        self._job_printer = ""
        self._job_office_mode = "自动"
        self._shell_print_used = False
        self.sumatra_path = self._detect_sumatra()

    def _detect_sumatra(self) -> Optional[str]:
        candidates = [os.environ.get("NIJIKA_SUMATRA_PATH", "")]
        if getattr(sys, "frozen", False):
            candidates.extend([
                str(Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "SumatraPDF.exe"),
                str(Path(sys.executable).parent / "SumatraPDF.exe"),
            ])
        candidates.extend([
            shutil.which("SumatraPDF.exe") or shutil.which("sumatrapdf") or "",
            str(Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "SumatraPDF" / "SumatraPDF.exe"),
            str(Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "SumatraPDF" / "SumatraPDF.exe"),
            str(Path(os.environ.get("LOCALAPPDATA", "")) / "SumatraPDF" / "SumatraPDF.exe"),
        ])
        for candidate in candidates:
            if candidate:
                path = os.path.abspath(os.path.expandvars(os.path.expanduser(candidate)))
                if os.path.isfile(path):
                    return path
        return None

    def _spreadsheet_candidates(self) -> List[Tuple[str, str]]:
        mode = self._job_office_mode
        wps = [
            ("KET.Application", "WPS 表格"),
            ("ket.Application", "WPS 表格"),
            ("ET.Application", "WPS 表格"),
            ("et.Application", "WPS 表格"),
        ]
        excel = [("Excel.Application", "Microsoft Excel")]
        if mode == "优先 WPS":
            return wps + excel
        if mode == "优先 Excel":
            return excel + wps
        return wps + excel

    def _print_worker(self, job: dict):
        printer = job["printer"]
        color = job["color"]
        duplex = job["duplex"]
        pdf_nup = job["pdf_nup"]
        files = job["files"]
        # 子线程中使用 COM（Word/Excel/PPT）前必须初始化
        self._job_printer = printer
        self._job_office_mode = job["office_mode"]
        com_inited = False
        try:
            pythoncom.CoInitialize()
            com_inited = True
        except Exception as e:
            self.print_queue.put(f"COM 初始化失败（Office 文档将使用系统默认程序打印）：{e}")

        try:
            old_default = win32print.GetDefaultPrinter()
        except Exception:
            old_default = None

        old_devmode = None
        try:
            win32print.SetDefaultPrinter(printer)
        except Exception:
            pass
        old_devmode = self._apply_printer_devmode(printer, color, duplex)

        rows, cols, order = self._resolve_nup_layout(pdf_nup=pdf_nup, layout_choice=job["nup_layout"])
        used_shell_print = False

        try:
            total = len(files)
            for idx, path in enumerate(files, start=1):
                if self.stop_print.is_set():
                    self.print_queue.put("已取消")
                    break
                self.print_queue.put(f"正在打印 {idx}/{total}：{os.path.basename(path)}")
                self._shell_print_used = False
                try:
                    ext = os.path.splitext(path)[1].lower()
                    if ext in (".doc", ".docx") and self.office_ok:
                        ok = self._print_word(path)
                    elif ext in (".xls", ".xlsx") and self.office_ok:
                        ok = self._print_excel(path)
                    elif ext in (".ppt", ".pptx") and self.office_ok:
                        ok = self._print_powerpoint(path)
                    elif ext == ".pdf":
                        ok = self._print_pdf(path, nup=pdf_nup, rows=rows, cols=cols, order=order)
                    elif ext in IMAGE_EXTS:
                        ok = self._print_image(path)
                    else:
                        ok = self._print_via_shell(path)
                    if self._shell_print_used:
                        used_shell_print = True
                    if not ok:
                        self.print_queue.put(f"打印可能失败：{os.path.basename(path)}")
                        self.failed_files.append((path, "打印可能失败"))
                except Exception as e:
                    print(traceback.format_exc())
                    self.print_queue.put(f"打印异常：{os.path.basename(path)} → {e}")
                    self.failed_files.append((path, str(e)))
                self.print_queue.put("__PROGRESS__")
                time.sleep(0.6)
            else:
                self.print_queue.put("打印任务完成")
        finally:
            # 系统默认程序打印是异步的，等待一段时间再还原默认打印机，
            # 避免文档被送到原来的默认打印机。
            if used_shell_print:
                time.sleep(8.0)
            self._restore_printer_devmode(printer, old_devmode)
            try:
                if old_default:
                    win32print.SetDefaultPrinter(old_default)
            except Exception:
                pass
            if com_inited:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            self.print_queue.put("__DONE__")

    def _apply_printer_devmode(self, printer_name: str, color: bool, duplex_mode: int) -> Optional[Tuple[int, int, int]]:
        """修改打印机默认 DEVMODE，返回修改前的 (Color, Duplex, Copies) 供打印结束后还原。"""
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                pinfo = win32print.GetPrinter(hPrinter, 2)
                devmode = pinfo.get("pDevMode", None)
                if devmode is None:
                    return None
                old = (
                    getattr(devmode, "Color", win32con.DMCOLOR_COLOR),
                    getattr(devmode, "Duplex", win32con.DMDUP_SIMPLEX),
                    getattr(devmode, "Copies", 1),
                )
                try:
                    fields = getattr(devmode, "Fields", 0)
                    fields |= (win32con.DM_COLOR | win32con.DM_DUPLEX | win32con.DM_COPIES)
                    devmode.Fields = fields
                except Exception:
                    pass
                try:
                    # DMCOLOR_MONOCHROME = 1, DMCOLOR_COLOR = 2
                    devmode.Color = win32con.DMCOLOR_COLOR if color else win32con.DMCOLOR_MONOCHROME
                except Exception:
                    pass
                try:
                    devmode.Duplex = int(duplex_mode)
                except Exception:
                    pass
                try:
                    devmode.Copies = 1
                except Exception:
                    pass

                pinfo["pDevMode"] = devmode
                win32print.SetPrinter(hPrinter, 2, pinfo, 0)
                return old
            finally:
                win32print.ClosePrinter(hPrinter)
        except Exception as e:
            self.print_queue.put(f"提示：无法修改打印机默认设置（彩色/双面可能不生效）：{e}")
            return None

    def _restore_printer_devmode(self, printer_name: str, old: Optional[Tuple[int, int, int]]):
        if not old:
            return
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                pinfo = win32print.GetPrinter(hPrinter, 2)
                devmode = pinfo.get("pDevMode", None)
                if devmode is None:
                    return
                devmode.Color, devmode.Duplex, devmode.Copies = old
                pinfo["pDevMode"] = devmode
                win32print.SetPrinter(hPrinter, 2, pinfo, 0)
            finally:
                win32print.ClosePrinter(hPrinter)
        except Exception:
            pass

    def _print_via_shell(self, path: str) -> bool:
        try:
            os.startfile(path, "print")
            # 异步打印：标记后由 worker 在还原默认打印机前多等待一段时间
            self._shell_print_used = True
            return True
        except Exception as e:
            self.print_queue.put(f"系统打印调用失败：{e}")
            return False

    def _print_word(self, path: str) -> bool:
        try:
            # 使用 DispatchEx 新建独立实例，避免附着到用户已打开的 Word 并把它 Quit 掉
            word = self._dispatch_com_app("Word.Application")
            try:
                word.Visible = False
                word.DisplayAlerts = 0
                word.Options.PrintBackground = False
            except Exception:
                pass

            doc = None
            try:
                doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
                doc.PrintOut(Background=False, Copies=1, Collate=True)
                time.sleep(0.4)
            finally:
                if doc:
                    try: doc.Close(False)
                    except Exception: pass
                try: word.Quit()
                except Exception: pass
            return True
        except Exception as e:
            self.print_queue.put(f"Word COM 打印失败（回退）：{e}")
            return self._print_via_shell(path)

    def _dispatch_com_app(self, progid: str):
        errors = []
        for creator in (DispatchEx, Dispatch):
            if creator is None:
                continue
            try:
                return creator(progid)
            except Exception as e:
                errors.append(str(e))
        raise RuntimeError("；".join(errors) or f"无法创建 COM 对象：{progid}")

    def _open_spreadsheet_app(self):
        candidates = self._spreadsheet_candidates()
        errors = []
        for progid, name in candidates:
            try:
                app = self._dispatch_com_app(progid)
                self.print_queue.put(f"使用 {name} 导出 PDF")
                return app, name
            except Exception as e:
                errors.append(f"{name}({progid})：{e}")
        raise RuntimeError("；".join(errors))

    def _print_excel(self, path: str) -> bool:
        excel = None
        wb = None
        pdf_path = None
        app_name = "表格程序"

        try:
            if not self.sumatra_path:
                self.print_queue.put("Excel 打印失败：未找到 SumatraPDF.exe，无法按“先转 PDF 再打印”的流程打印")
                return False

            excel, app_name = self._open_spreadsheet_app()
            try:
                excel.Visible = False
            except Exception:
                pass
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass

            wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
            time.sleep(0.8)

            pdf_path = os.path.join(
                tempfile.gettempdir(),
                f"excel_print_{os.getpid()}_{int(time.time() * 1000)}.pdf"
            )

            wb.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            if not os.path.exists(pdf_path):
                raise Exception("PDF 导出失败")

            return self._print_pdf_by_sumatra(pdf_path)

        except Exception as e:
            self.print_queue.put(f"{app_name} 转 PDF 打印失败：{e}")
            return False

        finally:
            try:
                if wb:
                    wb.Close(False)
            except Exception:
                pass

            try:
                if excel:
                    excel.Quit()
            except Exception:
                pass
            # 释放 COM 引用（COM 的初始化/反初始化统一由 _print_worker 负责）
            wb = None
            excel = None

            if pdf_path:
                try:
                    time.sleep(1.0)
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)
                except Exception:
                    pass

    def _print_powerpoint(self, path: str) -> bool:
        try:
            ppt = self._dispatch_com_app("PowerPoint.Application")
            try:
                ppt.Visible = True
            except Exception:
                pass
            pres = None
            try:
                pres = ppt.Presentations.Open(os.path.abspath(path), ReadOnly=True, WithWindow=False)
                try:
                    pres.PrintOptions.NumberOfCopies = 1
                except Exception:
                    pass
                pres.PrintOut(ShowPrintDialog=False)
                time.sleep(0.6)
            finally:
                if pres:
                    try: pres.Close()
                    except Exception: pass
                try: ppt.Quit()
                except Exception: pass
            return True
        except Exception as e:
            self.print_queue.put(f"PPT COM 打印失败（回退）：{e}")
            return self._print_via_shell(path)

    def _print_image(self, path: str) -> bool:
        if not self.sumatra_path:
            self.print_queue.put("图片打印失败：未找到 SumatraPDF.exe")
            return False
        if not HAVE_FITZ:
            self.print_queue.put("图片打印失败：图片转 PDF 组件不可用")
            return False

        tmp_pdf = os.path.join(
            tempfile.gettempdir(),
            f"image_print_{os.getpid()}_{int(time.time() * 1000)}.pdf"
        )
        doc = None
        img_doc = None
        try:
            doc = fitz.open()
            img_doc = fitz.open(path)
            if img_doc.page_count:
                pix = img_doc[0].get_pixmap(alpha=False)
                img_w, img_h = pix.width, pix.height
            else:
                img_rect = fitz.Rect(0, 0, 595, 842)
                img_w, img_h = img_rect.width, img_rect.height

            if img_w > img_h:
                page_w, page_h = 842, 595
            else:
                page_w, page_h = 595, 842

            page = doc.new_page(width=page_w, height=page_h)
            margin = 24
            max_w = page_w - margin * 2
            max_h = page_h - margin * 2
            scale = min(max_w / img_w, max_h / img_h)
            draw_w = img_w * scale
            draw_h = img_h * scale
            rect = fitz.Rect(
                (page_w - draw_w) / 2,
                (page_h - draw_h) / 2,
                (page_w + draw_w) / 2,
                (page_h + draw_h) / 2,
            )
            page.insert_image(rect, filename=path, keep_proportion=True)
            doc.save(tmp_pdf)
            img_doc.close()
            img_doc = None
            doc.close()
            doc = None
            return self._print_pdf_by_sumatra(tmp_pdf)
        except Exception as e:
            self.print_queue.put(f"图片转 PDF 打印失败：{e}")
            return False
        finally:
            for d in (img_doc, doc):
                try:
                    if d is not None:
                        d.close()
                except Exception:
                    pass
            try:
                time.sleep(1.0)
                if os.path.exists(tmp_pdf):
                    os.remove(tmp_pdf)
            except Exception:
                pass

    def _print_pdf(self, path: str, nup: int, rows:int, cols:int, order:str) -> bool:
        tmp_pdf = None
        try:
            if nup <= 1:
                if self.sumatra_path:
                    return self._print_pdf_by_sumatra(path)
                return self._print_via_shell(path)

            target = path
            if not self.have_fitz:
                self.print_queue.put("提示：PDF 合并组件不可用，将按 1 合 1 打印")
            else:
                tmpdir = tempfile.mkdtemp(prefix="pdf_nup_")
                tmp_pdf = os.path.join(tmpdir, "nup_out.pdf")
                ok = self._create_pdf_nup_py(
                    path, n=nup, out_pdf=tmp_pdf, rows=rows, cols=cols, order=order
                )
                if ok:
                    target = tmp_pdf
                else:
                    self.print_queue.put(f"提示：{os.path.basename(path)} 生成 {nup} 合 1 失败，将按 1 合 1 打印")

            if self.sumatra_path:
                return self._print_pdf_by_sumatra(target)

            # 回退
            return self._print_via_shell(target)
        except Exception as e:
            self.print_queue.put(f"PDF 打印异常：{e}")
            return False
        finally:
            if tmp_pdf:
                time.sleep(2.0)
                try:
                    shutil.rmtree(os.path.dirname(tmp_pdf))
                except Exception:
                    pass

    def _print_pdf_by_sumatra(self, pdf_path: str, printer_name: Optional[str] = None) -> bool:
        if not self.sumatra_path:
            return False
        printer_name = printer_name or self._job_printer
        if not printer_name:
            self.print_queue.put("Sumatra 打印失败：未选择打印机")
            return False

        cmd = [
            self.sumatra_path,
            "-silent",
            "-exit-on-print",
            "-print-to", printer_name,
            "-print-settings", "copies=1",
            pdf_path,
        ]
        try:
            startupinfo = None
            if sys.platform == "win32":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=120,
                startupinfo=startupinfo,
            )
            if proc.returncode != 0:
                self.print_queue.put(f"Sumatra 打印失败: {proc.stderr.decode(errors='ignore')}")
                return False
            return True
        except Exception as e:
            self.print_queue.put(f"调用 Sumatra 打印异常: {e}")
            return False

    def _create_pdf_nup_py(self, src_pdf: str, n: int, out_pdf: str, rows: int = 0, cols: int = 0,
                           order: str = "row") -> bool:
        src = None
        dst = None
        try:
            if not HAVE_FITZ:
                return False

            src = fitz.open(src_pdf)
            total_pages = src.page_count
            if total_pages == 0:
                return False

            import math
            if rows <= 0 or cols <= 0:
                cols = int(math.ceil(math.sqrt(n)))
                rows = int(math.ceil(n / cols))

            first = src.load_page(0)
            rect = first.rect
            page_w, page_h = rect.width, rect.height

            # 补白页使用与首页相同的尺寸
            while total_pages % n != 0:
                src.new_page(-1, width=page_w, height=page_h)
                total_pages += 1

            # 根据网格形状选择输出纸张方向（与原页面方向无关，均能得到最大缩放比）：
            #   横向排列（列多于行）→ 横向纸；纵向堆叠（行多于列）→ 竖向纸；
            #   方形网格 → 保持原页面方向。
            long_side, short_side = max(page_w, page_h), min(page_w, page_h)
            if cols > rows:
                out_w, out_h = long_side, short_side
            elif rows > cols:
                out_w, out_h = short_side, long_side
            else:
                out_w, out_h = page_w, page_h

            cell_w = out_w / cols
            cell_h = out_h / rows

            if order == "column":
                positions = [(r, c) for c in range(cols) for r in range(rows)]
            else:
                positions = [(r, c) for r in range(rows) for c in range(cols)]

            dst = fitz.open()

            p_index = 0
            while p_index < total_pages:
                page = dst.new_page(width=out_w, height=out_h)

                for r, c in positions:
                    if p_index >= total_pages:
                        break

                    x0 = c * cell_w
                    y0 = r * cell_h
                    x1 = x0 + cell_w
                    y1 = y0 + cell_h

                    target_rect = fitz.Rect(x0, y0, x1, y1)

                    page.show_pdf_page(target_rect, src, p_index)

                    p_index += 1

            dst.save(out_pdf)
            return True

        except Exception as e:
            print("n-up 生成异常：", e)
            traceback.print_exc()
            return False
        finally:
            for d in (dst, src):
                try:
                    if d is not None:
                        d.close()
                except Exception:
                    pass

    def _resolve_nup_layout(self, pdf_nup:int, layout_choice:str):
        import math

        if pdf_nup == 2:
            if "Vertical" in layout_choice or "竖排" in layout_choice:
                return 2, 1, "row"
            if "Horizontal" in layout_choice or "并排" in layout_choice:
                return 1, 2, "row"

        if "列优先" in layout_choice:
            cols = int(math.ceil(math.sqrt(pdf_nup)))
            rows = int(math.ceil(pdf_nup / cols))
            return rows, cols, "column"

        cols = int(math.ceil(math.sqrt(pdf_nup)))
        rows = int(math.ceil(pdf_nup / cols))
        return rows, cols, "row"

