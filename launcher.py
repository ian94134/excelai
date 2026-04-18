"""
PyInstaller 打包用的啟動器。
直接執行：python launcher.py
打包後執行：ExcelAI.exe
"""

import sys
import os
import threading
import time
import webbrowser

# ── log 檔（放在 EXE 同目錄，方便排查錯誤）──────────────────────────────────
LOG_PATH = os.path.join(os.path.dirname(sys.executable)
                        if getattr(sys, "frozen", False)
                        else os.path.dirname(os.path.abspath(__file__)),
                        "ExcelAI_log.txt")

def log(msg: str):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _open_browser(port: int = 8501):
    time.sleep(4.0)
    url = f"http://localhost:{port}"
    log(f"開啟瀏覽器 {url}")
    webbrowser.open(url)


def _show_error(msg: str):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(msg), "Excel AI 助手 - 錯誤", 0x10)
    except Exception:
        pass


def main():
    # 清除舊 log
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"=== ExcelAI 啟動 {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
    except Exception:
        pass

    log(f"Python: {sys.version}")
    log(f"frozen: {getattr(sys, 'frozen', False)}")
    log(f"executable: {sys.executable}")

    try:
        if getattr(sys, "frozen", False):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        log(f"base_dir: {base_dir}")
        main_py = os.path.join(base_dir, "main.py")
        log(f"main_py exists: {os.path.exists(main_py)}")

        if not os.path.exists(main_py):
            msg = f"找不到 main.py：\n{main_py}"
            log(f"[錯誤] {msg}")
            _show_error(msg)
            sys.exit(1)

        # ── Port 自動遞增（8501 被佔用就試 8502、8503…）─────────────────────────
        import socket
        port = 8501
        for candidate in range(8501, 8510):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(("localhost", candidate))
                port = candidate
                break
            except OSError:
                log(f"Port {candidate} 已被佔用，嘗試下一個…")
        log(f"使用 Port：{port}")

        threading.Thread(target=_open_browser, args=(port,), daemon=True).start()

        log("啟動 Streamlit...")
        sys.argv = [
            "streamlit", "run", main_py,
            "--global.developmentMode",    "false",
            "--server.headless",           "false",
            "--server.port",               str(port),
            "--server.address",            "localhost",
            "--browser.gatherUsageStats",  "false",
        ]

        from streamlit.web import cli as stcli
        log("stcli.main() 呼叫中...")
        stcli.main()

    except Exception as e:
        import traceback
        err = traceback.format_exc()
        log(f"[例外] {err}")
        _show_error(
            f"啟動失敗：\n{e}\n\n"
            f"詳細錯誤請查看：\n{LOG_PATH}"
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
