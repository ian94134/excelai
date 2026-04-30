"""excel/_undo.py — auto-split from excel_tools.py."""
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
from excel._base import _get_excel, _get_sheet, _com_tls
from excel.format import _restore_formats

def _undo_dispatch(entry) -> dict:
    """
    Core undo dispatcher.  Takes a BackupEntry and executes the reverse
    operation.  Called by undo_last() and by backup.restore().
    """
    return _undo_last_body(entry)



def undo_last() -> dict:
    """
    還原上一個可復原操作（Phase 2 實作）。

    策略分三類：
    Category A — 僅用 arguments 推算反向操作（insert_row / insert_column /
                 add_sheet / rename_sheet / merge_cells / unmerge_cells）
    Category B — 需要 values_before（executor 在工具執行前已讀取）
                 write_range / clear_range / trim_range
    Category C — 無法自動還原（回傳說明，不修改工作表）
    """
    import backup as bk
    stack = bk.get_session_stack()
    if stack is None or len(stack) == 0:
        return {"status": "no_op", "message": "備份堆疊為空，沒有可還原的操作。"}
    entry = stack.pop()
    try:
        return _undo_dispatch(entry)
    finally:
        bk.save_current_stack()



def _undo_last_body(entry):
    """Internal — real undo body (legacy path, kept for _undo_dispatch)."""
    name  = entry.tool_name
    args  = entry.arguments
    sheet = args.get("sheet")

    # ── Category B：回寫 formats_before（Phase 3：format_range / set_borders）──
    if name in ("format_range", "set_borders", "beautify_range"):
        if entry.formats_before is None or not entry.formats_before:
            return {
                "status":  "cannot_undo",
                "tool":    name,
                "message": f"「{name}」的格式備份未擷取（可能是升版前的舊備份項目），無法還原。",
            }
        rng_addr = args["range_addr"]
        _restore_formats(entry.formats_before, sheet)
        return {"status": "ok", "undone": name, "range": rng_addr}

    # ── Category B：回寫 values_before ────────────────────────────────────────
    if name in ("write_range", "clear_range", "trim_range"):
        if entry.values_before is None:
            return {
                "status":  "cannot_undo",
                "tool":    name,
                "message": f"「{name}」的前置資料未擷取（可能是升版前的舊備份項目），無法還原。",
            }
        rng_addr = args["range_addr"]
        excel = _get_excel()
        ws    = _get_sheet(excel, sheet)
        rng   = ws.Range(rng_addr)
        rows  = entry.values_before
        if rows == []:
            rng.ClearContents()
            return {"status": "ok", "undone": name, "range": rng_addr}
        # 單格範圍用 scalar 寫入；多格用批次寫入（一次 COM call，避免大範圍 O(n×m) 開銷）
        if len(rows) == 1 and len(rows[0]) == 1:
            rng.Value = rows[0][0] if rows[0][0] is not None else ""
        else:
            rng.Value = [[("" if v is None else v) for v in row] for row in rows]
        return {"status": "ok", "undone": name, "range": rng_addr}

    # ── Category B-widths: restore column widths (set_column_width / auto_fit) ──
    if name in ("set_column_width", "auto_fit") and entry.widths_before:
        excel = _get_excel()
        ws    = _get_sheet(excel, sheet)
        for col_idx, width in entry.widths_before.items():
            try:
                ws.Columns(col_idx).ColumnWidth = width
            except Exception:
                pass
        restored = list(entry.widths_before.keys())
        return {"status": "ok", "undone": name, "restored_cols": restored}

    # ── Category B-heights: restore row heights (set_row_height / auto_fit) ──────
    if name in ("set_row_height", "auto_fit") and entry.heights_before:
        excel = _get_excel()
        ws    = _get_sheet(excel, sheet)
        for row_idx, height in entry.heights_before.items():
            try:
                ws.Rows(row_idx).RowHeight = height
            except Exception:
                pass
        restored = list(entry.heights_before.keys())
        return {"status": "ok", "undone": name, "restored_rows": restored}

    # ── Category A：反向操作 ──────────────────────────────────────────────────
    excel = _get_excel()

    if name == "insert_row":
        idx   = args["index"]
        count = args.get("count", 1)
        ws    = _get_sheet(excel, sheet)
        # 插入了 count 列，從 idx 開始 → 刪除同樣範圍的列
        ws.Rows(f"{idx}:{idx + count - 1}").Delete()
        return {"status": "ok", "undone": name, "deleted_rows": f"{idx}:{idx + count - 1}"}

    if name == "insert_column":
        idx   = args["index"]
        count = args.get("count", 1)
        ws    = _get_sheet(excel, sheet)
        ws.Columns(f"{idx}:{idx + count - 1}").Delete()
        return {"status": "ok", "undone": name, "deleted_cols": f"{idx}:{idx + count - 1}"}

    if name == "add_sheet":
        new_name = args["name"]
        wb = excel.ActiveWorkbook
        if wb is None:
            raise NoActiveWorkbookError("沒有作用中的活頁簿")
        try:
            ws_to_del = wb.Sheets(new_name)
        except Exception:
            return {
                "status":  "cannot_undo",
                "tool":    name,
                "message": f"找不到工作表「{new_name}」，可能已被手動刪除。",
            }
        excel.DisplayAlerts = False
        try:
            ws_to_del.Delete()
        finally:
            excel.DisplayAlerts = True
        return {"status": "ok", "undone": name, "deleted_sheet": new_name}

    if name == "rename_sheet":
        original_name = args["old_name"]   # 執行前名稱（capture_before 存的是執行前引數）
        renamed_to    = args["new_name"]   # 執行後名稱（現在活頁簿裡的名稱）
        wb = excel.ActiveWorkbook
        if wb is None:
            raise NoActiveWorkbookError("沒有作用中的活頁簿")
        try:
            ws_renamed = wb.Sheets(renamed_to)
        except Exception:
            return {
                "status":  "cannot_undo",
                "tool":    name,
                "message": f"找不到工作表「{renamed_to}」（可能已再次被重新命名或刪除）。",
            }
        ws_renamed.Name = original_name
        return {"status": "ok", "undone": name, "restored_name": original_name}

    if name == "merge_cells":
        rng_addr = args["range_addr"]
        ws = _get_sheet(excel, sheet)
        ws.Range(rng_addr).UnMerge()
        return {"status": "ok", "undone": name, "unmerged": rng_addr}

    if name == "unmerge_cells":
        rng_addr = args["range_addr"]
        ws = _get_sheet(excel, sheet)
        ws.Range(rng_addr).Merge()
        return {"status": "ok", "undone": name, "merged": rng_addr}

    # ── Category C：嘗試 Excel 原生 Undo，失敗再回傳說明 ────────────────────
    # 先嘗試呼叫 Application.Undo()（相當於 Ctrl+Z）
    try:
        excel2 = _get_excel()
        excel2.CommandBars.ExecuteMso("Undo")
        return {
            "status":  "ok",
            "undone":  name,
            "method":  "native_undo",
            "message": f"「{name}」不支援自動還原，已透過 Excel 原生復原（Ctrl+Z）成功。",
        }
    except Exception:
        pass  # 原生 undo 失敗（例如 Excel undo stack 已空），繼續回傳說明

    CANNOT_UNDO_REASON: dict[str, str] = {
        "delete_row":             "列刪除後資料已遺失，無法還原。建議下次操作前先用 copy_sheet 備份。",
        "delete_column":          "欄刪除後資料已遺失，無法還原。",
        "delete_sheet":           "工作表刪除後無法自動還原，請改用 Excel 原生 Ctrl+Z。",
        "delete_chart":           "圖表刪除後無法還原。",
        "copy_sheet":             "複製工作表的還原需刪除複本，但無法確認複本名稱，請手動刪除。",
        "move_sheet":             "工作表移動無法自動反推原始位置，請手動調整。",
        # format_range 已升至 Category B（Phase 3），不應出現在此
        "sort_range":             "排序後原始順序已遺失，無法還原。",
        "find_replace":           "批次取代後舊值已遺失，無法還原。",
        "filter_range":           "篩選器狀態變更無法自動還原。",
        "freeze_panes":           "原始凍結狀態未記錄，無法還原。",
        "set_tab_color":          "原始索引標籤顏色未記錄，無法還原。",
        "split_text_to_columns":  "文字分欄後欄位結構已改變，無法自動還原。",
        "create_chart":           "圖表建立後若要撤銷請呼叫 delete_chart。",
        "create_pivot_table":     "樞紐分析表建立後若要撤銷請手動刪除。",
        # set_borders 已升至 Category B（Phase 3），不應出現在此
        "add_conditional_format": "條件格式備份尚未實作，無法還原。",
        "set_data_validation":    "資料驗證備份尚未實作，無法還原。",
        "fill_series":            "數列填滿後原始資料已遺失，無法還原。",
        "transpose_range":        "轉置後需刪除目標並還原來源，建議手動操作。",
        "group_rows":             "群組狀態變更無法自動還原。",
        "group_columns":          "群組狀態變更無法自動還原。",
        "add_subtotal":           "小計後工作表結構已改變，無法自動還原。",
        "advanced_filter":        "進階篩選結果無法自動還原。",
        "name_range":             "具名範圍可使用 Excel 名稱管理員手動刪除。",
        "find_duplicates":        "重複值標記或刪除後無法自動還原。",
        "beautify_range":         "一鍵美化格式備份未擷取，無法還原。",
        "apply_table_style":      "表格樣式格式備份尚未實作，無法還原。",
        "format_chart":           "圖表格式備份尚未實作，無法還原。",
        "create_combo_chart":     "組合圖建立後若要撤銷請呼叫 delete_chart。",
        "add_sparklines":         "走勢圖備份尚未實作，無法還原。",
        "page_setup":             "頁面設定備份尚未實作，無法還原。",
        "add_slicer":             "交叉分析篩選器備份尚未實作，無法還原。",
        "add_image":              "圖片插入後若要撤銷請手動刪除。",
        "protect_sheet":          "工作表保護狀態備份尚未實作，無法還原。",
        "unprotect_sheet":        "工作表取消保護備份尚未實作，無法還原。",
        "set_print_titles":       "列印標題備份尚未實作，無法還原。",
        "add_header_footer":      "頁首頁尾備份尚未實作，無法還原。",
        "move_chart":             "圖表原始位置未記錄，無法還原。",
        "format_pivot_table":     "樞紐格式備份尚未實作，無法還原。",
        "auto_fit":               "原始欄/列寬備份尚未實作，無法還原。",
        "set_column_width":       "原始欄寬備份尚未實作，無法還原。",
        "set_row_height":         "原始列高備份尚未實作，無法還原。",
        "copy_range":             "複製貼上後目標資料已覆蓋，無法自動還原。",
        "add_comment":            "批註新增後若要撤銷請手動刪除批註。",
        "trim_range":             "trim_range values_before 未擷取，無法還原。",
    }
    reason = CANNOT_UNDO_REASON.get(name, f"「{name}」屬於尚未支援還原的操作類型。")
    return {"status": "cannot_undo", "tool": name, "message": reason}
