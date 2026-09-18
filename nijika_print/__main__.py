"""Run Nijika Print as a Python module or packaged desktop application."""
import json
from pathlib import Path
import sys


def entry():
    diagnostic = len(sys.argv) == 3 and sys.argv[1] == "--diagnose"
    try:
        from nijika_print.app import main
        main()
    except Exception as exc:
        if not diagnostic:
            raise
        Path(sys.argv[2]).write_text(json.dumps({"ok": False, "error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False, indent=2), encoding="utf-8")
        raise SystemExit(1)


if __name__ == "__main__":
    entry()
