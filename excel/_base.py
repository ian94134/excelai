"""excel/_base.py — auto-split from excel_tools.py."""
from __future__ import annotations
import json
import time
import threading
import pythoncom
import win32com.client
from constants import (
    XL_ALIGN, XL_CHART_TYPE, XL_BORDER_STYLE, XL_BORDER_SIDES,
    XL_ROW_FIELD, XL_COL_FIELD, XL_DATA_FIELD, XL_SUM, XL_DATABASE,
    XL_ASCENDING, XL_DESCENDING, XL_YES, XL_NO,
    XL_SPARKLINE_LINE, XL_SPARKLINE_COLUMN, XL_SPARKLINE_WINLOSS,
    XL_LEGEND_BOTTOM, XL_LEGEND_RIGHT, XL_LEGEND_TOP, XL_LEGEND_LEFT,
    XL_PORTRAIT, XL_LANDSCAPE, XL_PAPER_SIZE, TABLE_STYLE_MAP,
    XL_SERIES_LINEAR, XL_SERIES_DATE, XL_SERIES_COLUMNS, XL_SERIES_ROWS,
    XL_DATE_DAY, XL_DATE_WEEKDAY, XL_DATE_MONTH, XL_DATE_YEAR,
    XL_FILTER_IN_PLACE, XL_FILTER_COPY,
    XL_COUNT, XL_AVERAGE, XL_MAX, XL_MIN_FUNC,
    XL_DELIMITED,
)
from utils import col_letter as _col_letter
from exceptions import (
    ExcelNotFoundError,
    NoActiveWorkbookError,
    SheetNotFoundError,
    InvalidToolArgumentsError,
)

# Thread-local flag: CoInitialize is called at most once per thread.
_com_tls = threading.local()


def _get_excel():
    """
    取得已開啟的 Excel Application 實例。
    ★ B7：Streamlit 在子線程處理請求，COM 需要在每個線程中初始化，
    否則 GetActiveObject 會失敗。加入 CoInitialize + Dispatch 備援。
    """
    # Initialise COM apartment for this thread exactly once.
    if not getattr(_com_tls, 'initialised', False):
        pythoncom.CoInitialize()
        _com_tls.initialised = True

    # 方法 1：標準 ROT 查詢（最可靠的「找到已開啟的 Excel」方式）
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        pass

    # 方法 2：Dispatch 連接（某些情況 ROT 未註冊但 Dispatch 可連接既有實例）
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        if excel.Workbooks.Count > 0:
            return excel
    except Exception:
        pass

    raise ExcelNotFoundError("找不到開啟中的 Excel，請先開啟 Excel 再執行此指令")



def _get_sheet(excel, sheet_name: str | None = None):
    """
    取得工作表物件。
    ★ B5：ActiveWorkbook None guard
    ★ V4 P1：sheet_name 不存在時丟 SheetNotFoundError（原本 COM 透出的 HRESULT 訊息對 LLM 不友善）
    """
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    if sheet_name:
        try:
            return wb.Sheets(sheet_name)
        except Exception as e:
            existing = [ws.Name for ws in wb.Sheets]
            raise SheetNotFoundError(
                f"找不到工作表 '{sheet_name}'。目前可用的工作表：{existing}"
            ) from e
    return excel.ActiveSheet



def _hex_to_bgr(hex_color: str) -> int:
    """把 #RRGGBB 轉換為 Excel COM 的 BGR int"""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return b * 65536 + g * 256 + r



def _normalize_values(raw) -> list[list]:
    """
    ★ B6：統一 win32com Value 回傳格式
    - 單格：純量 → [[value]]
    - 單行：flat tuple → [[v1, v2, ...]]
    - 多行：tuple of tuples → [[...], [...]]
    """
    if raw is None:
        return []
    # 多行：第一個元素也是 tuple
    if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple):
        return [list(row) for row in raw]
    # 單行：flat tuple
    if isinstance(raw, tuple):
        return [list(raw)]
    # 單格：純量
    return [[raw]]



def _ensure_positive_int(name: str, value: int) -> None:
    if not isinstance(value, int) or value <= 0:
        raise InvalidToolArgumentsError(f"{name} 必須是大於 0 的整數")



def _ensure_positive_number(name: str, value: float | int) -> None:
    if not isinstance(value, (int, float)) or value <= 0:
        raise InvalidToolArgumentsError(f"{name} 必須是大於 0 的數值")


# ── 格式備份 helper（Phase 3）────────────────────────────────────────────────

# win32com 對齊常數（用於 HorizontalAlignment 讀取）
_XL_HALIGN_NONE     = -4142
_INTERIOR_NONE_IDX  = -4142   # xlColorIndexNone（無填色）

# 框線邊索引（與 XL_BORDER_SIDES 定義一致）
_ALL_BORDER_IDX = [7, 8, 9, 10, 11, 12]  # Left/Right/Top/Bottom/InsideH/InsideV




