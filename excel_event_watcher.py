"""
Excel 選取範圍即時監聽器。

設計原則：
- 背景執行緒每 POLL_INTERVAL 秒輪詢 Excel 目前選取範圍（比 win32com event sink 更穩定）
- 若選取沒有變化則不寫入 session_state（避免 Streamlit rerun）
- 只在 Streamlit 環境下工作；非 Streamlit 環境（pytest）自動跳過
- 執行緒為 daemon，App 結束時自動終止
"""

from __future__ import annotations
import threading
import time
import pythoncom
import win32com.client

POLL_INTERVAL = 1.5   # 秒（不宜太頻繁，避免 COM 開銷）
_watcher_started = False
_lock = threading.Lock()


def _poll_loop() -> None:
    """背景執行緒主體：持續輪詢 Excel 選取狀態。"""
    pythoncom.CoInitialize()
    last_signature: str = ""

    while True:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            wb    = excel.ActiveWorkbook
            ws    = excel.ActiveSheet

            if wb and ws:
                try:
                    sel_addr     = excel.Selection.Address
                    sheet_name   = ws.Name
                    wb_name      = wb.Name
                    signature    = f"{wb_name}|{sheet_name}|{sel_addr}"

                    if signature != last_signature:
                        last_signature = signature
                        _update_session_state(wb_name, sheet_name, sel_addr)
                except Exception:
                    pass  # 選取物件可能是圖表或其他非 Range 物件

        except Exception:
            last_signature = ""  # Excel 可能已關閉，重置

        time.sleep(POLL_INTERVAL)


def _update_session_state(wb_name: str, sheet_name: str, sel_addr: str) -> None:
    """將最新選取寫入 Streamlit session_state。"""
    try:
        import streamlit as st
        st.session_state["_excel_selection"] = {
            "workbook": wb_name,
            "sheet":    sheet_name,
            "address":  sel_addr,
        }
    except Exception:
        pass


def start_watcher() -> None:
    """
    啟動背景選取監聽執行緒（只啟動一次）。
    在 main.py 或 sidebar.py 呼叫一次即可。
    """
    global _watcher_started
    with _lock:
        if _watcher_started:
            return
        t = threading.Thread(target=_poll_loop, daemon=True, name="ExcelSelectionWatcher")
        t.start()
        _watcher_started = True


def get_current_selection() -> dict | None:
    """
    取得最新的選取資訊（從 session_state）。
    回傳 {"workbook": ..., "sheet": ..., "address": ...} 或 None。
    """
    try:
        import streamlit as st
        return st.session_state.get("_excel_selection")
    except Exception:
        return None
