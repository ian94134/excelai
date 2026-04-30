"""excel/sheet.py — auto-split from excel_tools.py."""
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
from excel._base import (
    _ensure_positive_int,
    _get_excel,
    _get_sheet,
    _normalize_values,
    _com_tls,
)

def insert_row(index: int, count: int = 1, sheet: str | None = None) -> dict:
    _ensure_positive_int("index", index)
    _ensure_positive_int("count", count)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(f"{index}:{index + count - 1}").Insert()
    return {"status": "ok", "inserted_at": index, "count": count}



def delete_row(index: int, count: int = 1, sheet: str | None = None) -> dict:
    _ensure_positive_int("index", index)
    _ensure_positive_int("count", count)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(f"{index}:{index + count - 1}").Delete()
    return {"status": "ok", "deleted_at": index, "count": count}


# ── 欄操作 ─────────────────────────────────────────────────────────────────────


def insert_column(index: int, count: int = 1, sheet: str | None = None) -> dict:
    _ensure_positive_int("index", index)
    _ensure_positive_int("count", count)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Range(ws.Columns(index), ws.Columns(index + count - 1)).Insert()
    return {"status": "ok", "inserted_at": index, "count": count}



def delete_column(index: int, count: int = 1, sheet: str | None = None) -> dict:
    _ensure_positive_int("index", index)
    _ensure_positive_int("count", count)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Range(ws.Columns(index), ws.Columns(index + count - 1)).Delete()
    return {"status": "ok", "deleted_at": index, "count": count}


# ── 工作表操作 ─────────────────────────────────────────────────────────────────


def add_sheet(name: str) -> dict:
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    try:
        wb.Sheets.Add().Name = name
    except Exception as e:
        raise InvalidToolArgumentsError(
            f"無法新增工作表 '{name}'。可能原因：名稱重複或包含非法字元。"
        ) from e
    return {"status": "ok", "added": name}



def rename_sheet(old_name: str, new_name: str) -> dict:
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    try:
        ws = wb.Sheets(old_name)
    except Exception as e:
        existing = [ws.Name for ws in wb.Sheets]
        raise SheetNotFoundError(
            f"找不到工作表 '{old_name}'。目前可用的工作表：{existing}"
        ) from e
    try:
        ws.Name = new_name
    except Exception as e:
        raise InvalidToolArgumentsError(
            f"無法將工作表改名為 '{new_name}'。可能原因：名稱重複、過長或含非法字元。"
        ) from e
    return {"status": "ok", "renamed": f"{old_name} → {new_name}"}



def delete_sheet(name: str) -> dict:
    """
    刪除指定工作表。⚠️ 危險操作：刪除後無法復原。
    活頁簿至少需保留一張工作表，否則會失敗。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿")
    if wb.Sheets.Count <= 1:
        raise InvalidToolArgumentsError("無法刪除：活頁簿至少需保留一張工作表")
    try:
        ws = wb.Sheets(name)
    except Exception as e:
        existing = [s.Name for s in wb.Sheets]
        raise SheetNotFoundError(
            f"找不到工作表 '{name}'。目前可用的工作表：{existing}"
        ) from e
    # 關閉刪除確認對話框
    excel.DisplayAlerts = False
    try:
        ws.Delete()
    finally:
        excel.DisplayAlerts = True
    return {"status": "ok", "deleted": name}



def move_sheet(
    name: str,
    before: str | None = None,
    after: str | None = None,
) -> dict:
    """
    移動工作表到指定位置。
    before：移動到此工作表之前；after：移動到此工作表之後。
    before 與 after 二選一；都省略則移到最後。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿")
    ws = _get_sheet(excel, name)
    if before:
        target = _get_sheet(excel, before)
        ws.Move(target, None)
    elif after:
        target = _get_sheet(excel, after)
        ws.Move(None, target)
    else:
        # 移到最後（After 最後一張）
        last = wb.Sheets(wb.Sheets.Count)
        if last.Name != name:
            ws.Move(None, last)
    return {"status": "ok", "moved": name, "before": before, "after": after}



def copy_sheet(
    name: str,
    new_name: str | None = None,
    before: str | None = None,
    after: str | None = None,
) -> dict:
    """
    複製工作表（含所有資料與格式）。
    new_name：複製後的新工作表名稱；省略則 Excel 自動命名（如 '工作表1 (2)'）。
    before / after：插入位置，都省略則複製到最後。
    常用場景：修改前先備份原始工作表。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿")
    ws = _get_sheet(excel, name)

    if before:
        target = _get_sheet(excel, before)
        ws.Copy(target, None)
    elif after:
        target = _get_sheet(excel, after)
        ws.Copy(None, target)
    else:
        last = wb.Sheets(wb.Sheets.Count)
        ws.Copy(None, last)

    # 複製後新工作表自動成為 ActiveSheet
    copied_ws = excel.ActiveSheet
    if new_name:
        try:
            copied_ws.Name = new_name
        except Exception as e:
            raise InvalidToolArgumentsError(
                f"工作表已複製，但改名為 '{new_name}' 失敗：{e}"
            ) from e

    return {"status": "ok", "copied_from": name, "new_sheet": copied_ws.Name}



def protect_sheet(
    password: str | None = None,
    allow_select_locked: bool = True,
    allow_select_unlocked: bool = True,
    allow_format_cells: bool = False,
    allow_insert_rows: bool = False,
    allow_delete_rows: bool = False,
    allow_sort: bool = False,
    allow_filter: bool = False,
    sheet: str | None = None,
) -> dict:
    """
    保護工作表，防止誤改公式或格式。
    password：保護密碼（省略則無密碼保護，任何人都可解除）。
    allow_* 參數控制保護後使用者仍可進行的操作。
    常見用法：保護公式欄，只開放輸入欄（先解鎖輸入欄的 locked 屬性再保護）。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Protect(
        Password=password or "",
        DrawingObjects=True,
        Contents=True,
        Scenarios=True,
        AllowFormattingCells=allow_format_cells,
        AllowInsertingRows=allow_insert_rows,
        AllowDeletingRows=allow_delete_rows,
        AllowSorting=allow_sort,
        AllowFiltering=allow_filter,
        AllowUsingPivotTables=False,
        UserInterfaceOnly=False,
    )
    return {
        "status":   "ok",
        "sheet":    ws.Name,
        "protected": True,
        "has_password": bool(password),
    }



def unprotect_sheet(
    password: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    解除工作表保護。
    password：若保護時有設定密碼，需提供正確密碼才能解除。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    try:
        ws.Unprotect(Password=password or "")
    except Exception as e:
        raise InvalidToolArgumentsError(
            f"無法解除保護：密碼錯誤或工作表未受保護。詳細：{e}"
        ) from e
    return {"status": "ok", "sheet": ws.Name, "protected": False}



def filter_range(
    range_addr: str,
    column_index: int,
    criteria: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    對範圍套用 AutoFilter 篩選。
    column_index：範圍內第幾欄（從 1 開始）。
    criteria：篩選值，如 "台北" 或 ">100"；省略則顯示全部（清除篩選）。
    """
    _ensure_positive_int("column_index", column_index)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    # 先清除既有篩選
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False

    if criteria is not None:
        rng.AutoFilter(Field=column_index, Criteria1=criteria)
    else:
        rng.AutoFilter(Field=column_index)  # 顯示全部

    return {"status": "ok", "filtered_column": column_index, "criteria": criteria}



def group_rows(
    start_row: int,
    end_row: int,
    action: str = "group",
    sheet: str | None = None,
) -> dict:
    """
    對指定列範圍進行分組（大綱），可折疊展開。
    action：group（建立分組）/ ungroup（移除分組）。
    常用場景：把明細列分組，折疊後只顯示小計列；多層級報表必備。
    """
    _ensure_positive_int("start_row", start_row)
    _ensure_positive_int("end_row",   end_row)
    if start_row > end_row:
        raise InvalidToolArgumentsError("start_row 不能大於 end_row")

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    row_range = ws.Rows(f"{start_row}:{end_row}")

    if action == "group":
        row_range.Group()
    elif action == "ungroup":
        row_range.Ungroup()
    else:
        raise InvalidToolArgumentsError("action 必須是 group 或 ungroup")

    return {"status": "ok", "action": action, "rows": f"{start_row}:{end_row}"}



def group_columns(
    start_col: int,
    end_col: int,
    action: str = "group",
    sheet: str | None = None,
) -> dict:
    """
    對指定欄範圍進行分組（大綱），可折疊展開。
    start_col / end_col：欄號（1-based，1=A、2=B…）。
    action：group（建立分組）/ ungroup（移除分組）。
    """
    _ensure_positive_int("start_col", start_col)
    _ensure_positive_int("end_col",   end_col)
    if start_col > end_col:
        raise InvalidToolArgumentsError("start_col 不能大於 end_col")

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    col_range = ws.Range(ws.Columns(start_col), ws.Columns(end_col))

    if action == "group":
        col_range.Group()
    elif action == "ungroup":
        col_range.Ungroup()
    else:
        raise InvalidToolArgumentsError("action 必須是 group 或 ungroup")

    return {"status": "ok", "action": action, "columns": f"{start_col}:{end_col}"}



def list_workbooks() -> dict:
    """列出目前所有已開啟的 Excel 活頁簿。"""
    excel = _get_excel()
    books = []
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        books.append({
            "name":   wb.Name,
            "path":   wb.FullName,
            "active": (wb.Name == excel.ActiveWorkbook.Name),
        })
    return {"status": "ok", "workbooks": books, "count": len(books)}



def switch_workbook(name: str) -> dict:
    """切換到指定活頁簿（讓它成為 ActiveWorkbook）。"""
    excel = _get_excel()
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        if wb.Name == name:
            wb.Activate()
            return {"status": "ok", "activated": wb.Name}
    available = [excel.Workbooks(i).Name for i in range(1, excel.Workbooks.Count + 1)]
    raise InvalidToolArgumentsError(
        f"找不到活頁簿「{name}」。目前開啟的活頁簿：{available}"
    )



def snapshot_sheet(sheet: str | None = None) -> dict:
    """
    將作用中工作表（或指定工作表）的所有儲存格值存入快照。
    回傳 {"status": "ok", "sheet": 名稱, "range": 範圍, "data": [[...]]}。
    快照儲存在呼叫端（sidebar）的 session_state，不寫入 Excel。
    """
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)
    used  = ws.UsedRange
    data  = _normalize_values(used.Value)
    return {
        "status": "ok",
        "sheet":  ws.Name,
        "range":  used.Address,
        "data":   data,
        "rows":   len(data),
        "cols":   len(data[0]) if data else 0,
    }



def restore_snapshot(data: list[list], range_addr: str, sheet: str | None = None) -> dict:
    """
    將快照資料回寫到指定工作表的起始範圍。
    清空 UsedRange 後，從 range_addr 開始逐格寫入。
    """
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)

    # 清空現有內容（保留格式）
    ws.UsedRange.ClearContents()

    if not data:
        return {"status": "ok", "message": "快照為空，已清空工作表內容"}

    rows = len(data)
    cols = max(len(row) for row in data)
    start = ws.Range(range_addr).Cells(1, 1)
    end   = start.Offset(rows - 1, cols - 1)
    rng   = ws.Range(start, end)

    # 補齊每列長度後寫入
    padded = [row + [None] * (cols - len(row)) for row in data]
    rng.Value = padded
    return {
        "status":  "ok",
        "sheet":   ws.Name,
        "range":   rng.Address,
        "rows":    rows,
        "cols":    cols,
    }


def copy_range_between_workbooks(
    source_range:  str,
    dest_range:    str,
    source_wb:     str | None = None,
    dest_wb:       str | None = None,
    source_sheet:  str | None = None,
    dest_sheet:    str | None = None,
    values_only:   bool = True,
) -> dict:
    """
    跨活頁簿複製範圍資料（v4.7.0）。
    values_only=True（預設）只複製值；False 同時複製格式。
    """
    excel = _get_excel()

    def _get_wb(name):
        if not name:
            return excel.ActiveWorkbook
        for i in range(1, excel.Workbooks.Count + 1):
            wb = excel.Workbooks(i)
            if wb.Name == name:
                return wb
        available = [excel.Workbooks(i).Name for i in range(1, excel.Workbooks.Count + 1)]
        raise InvalidToolArgumentsError(
            f"找不到活頁簿「{name}」。目前開啟：{available}"
        )

    def _get_ws(wb, sheet_name):
        if wb is None:
            raise NoActiveWorkbookError("沒有作用中的活頁簿")
        if sheet_name:
            try:
                return wb.Worksheets(sheet_name)
            except Exception as e:
                existing = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
                raise SheetNotFoundError(
                    f"找不到工作表 '{sheet_name}'。目前可用的工作表：{existing}"
                ) from e
        if excel.ActiveWorkbook is not None and wb.Name == excel.ActiveWorkbook.Name:
            return excel.ActiveSheet
        return wb.Worksheets(1)

    src_wb_obj = _get_wb(source_wb)
    dst_wb_obj = _get_wb(dest_wb)
    src_ws = _get_ws(src_wb_obj, source_sheet)
    dst_ws = _get_ws(dst_wb_obj, dest_sheet)

    src_rng = src_ws.Range(source_range)
    values  = _normalize_values(src_rng.Value)

    if not values:
        return {"status": "ok", "message": "來源範圍為空，無資料複製"}

    rows = len(values)
    cols = max(len(r) for r in values)
    start = dst_ws.Range(dest_range).Cells(1, 1)
    end   = dst_ws.Cells(start.Row + rows - 1, start.Column + cols - 1)
    dst_rng = dst_ws.Range(start, end)

    if values_only:
        padded = [row + [None] * (cols - len(row)) for row in values]
        dst_rng.Value = padded
    else:
        src_rng.Copy(Destination=start)

    return {
        "status":      "ok",
        "source":      f"{src_wb_obj.Name}/{src_ws.Name}!{src_rng.Address}",
        "destination": f"{dst_wb_obj.Name}/{dst_ws.Name}!{dst_rng.Address}",
        "rows":        rows,
        "cols":        cols,
        "values_only": values_only,
    }


# ── Phase 2 Undo ──────────────────────────────────────────────────────────────
