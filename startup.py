from __future__ import annotations
import ctypes
import datetime as dt
import os
import platform
import sys
import traceback
from pathlib import Path

BASE = Path(__file__).resolve().parent
STARTUP_LOG = BASE / "startup.log"
ERROR_LOG = BASE / "error.log"

def append_log(path: Path, text: str) -> None:
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with path.open("a", encoding="utf-8") as f:
        f.write(f"\n[{stamp}]\n{text}\n")

def show_error(message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(0, message, "太炅 Lotto Lab 실행 오류", 0x10)
    except Exception:
        pass

def main() -> int:
    os.chdir(BASE)
    append_log(
        STARTUP_LOG,
        "START\n"
        f"python={sys.executable}\n"
        f"version={sys.version}\n"
        f"platform={platform.platform()}\n"
        f"cwd={Path.cwd()}"
    )
    try:
        import pandas
        import openpyxl
        import PySide6
        append_log(
            STARTUP_LOG,
            "MODULE CHECK OK\n"
            f"pandas={pandas.__version__}\n"
            f"openpyxl={openpyxl.__version__}\n"
            f"PySide6={PySide6.__version__}"
        )

        import app
        append_log(STARTUP_LOG, "APP IMPORT OK")
        result = int(app.main())
        append_log(STARTUP_LOG, f"APP EXIT code={result}")
        return result
    except BaseException:
        detail = traceback.format_exc()
        append_log(ERROR_LOG, detail)
        append_log(STARTUP_LOG, "APP FAILED - see error.log")
        show_error(
            "프로그램 시작 중 오류가 발생했습니다.\n\n"
            "같은 폴더의 error.log 파일에 원인이 기록되었습니다."
        )
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
