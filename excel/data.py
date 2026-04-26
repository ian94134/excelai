"""excel/data.py — auto-split from excel_tools.py."""
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
from excel._base import _get_excel, _get_sheet, _hex_to_bgr, _normalize_values, _ensure_positive_int, _ensure_positive_number, _com_tls

def read_range(range_addr: str, sheet: str | None = None) -> list[list]:
    """讀取儲存格範圍，回傳二維 list（★ B3/B6 修正）"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    return _normalize_values(rng.Value)



def get_sheet_info() -> dict:
    """取得活頁簿基本資訊"""
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel 已開啟但沒有活頁簿")

    # 選取物件不一定是 Range（例如圖表被選取）；避免因此整個 tool 失敗。
    try:
        selection_addr = excel.Selection.Address
    except Exception:
        selection_addr = "(non-range selection)"

    return {
        "file_name":    wb.Name,
        "sheets":       [ws.Name for ws in wb.Sheets],
        "active_sheet": excel.ActiveSheet.Name,
        "selection":    selection_addr,
    }



def get_used_range(sheet: str | None = None) -> str:
    """回傳工作表中有資料的範圍位址"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    return ws.UsedRange.Address


def get_workbook_summary() -> dict:
    """
    One-shot summary of every sheet in the active workbook.
    Returns file name, active sheet, and for each sheet:
    used_range, row/column counts, and up to 10 header names from row 1.
    Use this before complex multi-sheet tasks so the LLM has a full picture
    without issuing separate get_sheet_info + get_used_range calls per sheet.
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("Excel is open but no workbook is active")

    sheets_info = []
    for ws in wb.Sheets:
        try:
            used = ws.UsedRange
            row_count = used.Rows.Count
            col_count = used.Columns.Count
            try:
                hdr_range = ws.Cells(1, 1).Resize(1, min(col_count, 20))
                raw = _normalize_values(hdr_range.Value)
                headers = [str(h) for h in (raw[0] if raw else [])
                           if h is not None and str(h).strip()]
            except Exception:
                headers = []
            sheets_info.append({
                "name":           ws.Name,
                "used_range":     used.Address,
                "rows":           row_count,
                "columns":        col_count,
                "sample_headers": headers[:10],
            })
        except Exception as e:
            sheets_info.append({"name": ws.Name, "error": str(e)})

    return {
        "file_name":    wb.Name,
        "active_sheet": excel.ActiveSheet.Name,
        "sheet_count":  wb.Sheets.Count,
        "sheets":       sheets_info,
    }


# ── 寫入 ──────────────────────────────────────────────────────────────────────


def write_range(range_addr: str, values: list[list], sheet: str | None = None) -> dict:
    """
    寫入二維陣列到指定範圍。
    ★ B4：偵測 = 開頭自動用 .Formula；否則用 .Value
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    start = ws.Range(range_addr).Cells(1, 1)

    rows = len(values)
    cols = max(len(row) for row in values) if values else 0

    for r_idx, row in enumerate(values):
        for c_idx, val in enumerate(row):
            cell = ws.Cells(start.Row + r_idx, start.Column + c_idx)
            if isinstance(val, str) and val.startswith("="):
                cell.Formula = val   # ★ 公式用 .Formula
            else:
                cell.Value = val     # 普通值用 .Value

    written_range = ws.Range(
        ws.Cells(start.Row, start.Column),
        ws.Cells(start.Row + rows - 1, start.Column + cols - 1)
    )

    # ── formula error validation: read back cells written with "=" formulas ──
    formula_errors = []
    _XL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?",
                  "#NUM!", "#N/A", "#GETTING_DATA", "#####"}
    for r_idx, row in enumerate(values):
        for c_idx, val in enumerate(row):
            if isinstance(val, str) and val.startswith("="):
                cell = ws.Cells(start.Row + r_idx, start.Column + c_idx)
                try:
                    cell_val = str(cell.Value) if cell.Value is not None else ""
                except Exception:
                    cell_val = ""
                if cell_val in _XL_ERRORS or cell_val.startswith("#"):
                    formula_errors.append({
                        "cell":    f"{_col_letter(start.Column + c_idx)}{start.Row + r_idx}",
                        "formula": val,
                        "error":   cell_val,
                    })

    result: dict = {"status": "ok", "written_to": written_range.Address}
    if formula_errors:
        result["formula_errors"] = formula_errors
        result["warning"] = (
            f"{len(formula_errors)} formula(s) produced errors after writing. "
            "Check cell references, sheet names, and argument types."
        )
    return result



def save_workbook() -> dict:
    """儲存目前作用中的活頁簿（★ 新增 save tool）"""
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise NoActiveWorkbookError("沒有開啟中的活頁簿")
    wb.Save()
    return {"status": "ok", "saved": wb.Name}


# ── 格式 ──────────────────────────────────────────────────────────────────────


def sort_range(
    range_addr: str,
    column_index: int,
    ascending: bool = True,
    has_header: bool = True,
    sheet: str | None = None,
) -> dict:
    _ensure_positive_int("column_index", column_index)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    rng.Sort(
        Key1=rng.Columns(column_index),
        Order1=XL_ASCENDING if ascending else XL_DESCENDING,
        Header=XL_YES if has_header else XL_NO,
    )
    return {"status": "ok", "sorted_by_column": column_index, "ascending": ascending}



def find_replace(find: str, replace: str, sheet: str | None = None) -> dict:
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Cells.Replace(What=find, Replacement=replace)
    return {"status": "ok", "find": find, "replace": replace}



def trim_range(
    range_addr: str,
    sheet: str | None = None,
) -> dict:
    """
    清除儲存格中多餘的空格（相當於 Excel TRIM 函數）：
    - 移除開頭與結尾的空格
    - 將文字中間的連續空格壓縮為單一空格
    只處理文字型儲存格，數值與公式不受影響。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    raw = rng.Value
    values = _normalize_values(raw)

    changed = 0
    start = rng.Cells(1, 1)

    for r_idx, row in enumerate(values):
        for c_idx, val in enumerate(row):
            if isinstance(val, str):
                cleaned = " ".join(val.split())
                if cleaned != val:
                    ws.Cells(start.Row + r_idx, start.Column + c_idx).Value = cleaned
                    changed += 1

    return {"status": "ok", "range": rng.Address, "cells_cleaned": changed}


# ── 圖表 ──────────────────────────────────────────────────────────────────────




def clear_range(
    range_addr: str | None = None,
    target: str = "values",
    sheet: str | None = None,
) -> dict:
    """
    清除儲存格內容或格式。
    target：'values'（清值保留格式）/ 'formats'（清格式保留值）/ 'all'（全清）
    range_addr：省略則清除已使用範圍。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr) if range_addr else ws.UsedRange

    if target == "values":
        rng.ClearContents()
    elif target == "formats":
        rng.ClearFormats()
    elif target == "all":
        rng.Clear()
    else:
        raise InvalidToolArgumentsError("clear_range.target 僅允許 values / formats / all")

    return {"status": "ok", "cleared": target, "range": range_addr or "UsedRange"}



def copy_range(
    source_range: str,
    dest_range: str,
    source_sheet: str | None = None,
    dest_sheet: str | None = None,
) -> dict:
    """
    複製範圍到另一位置（可跨工作表）。
    dest_sheet 不存在時自動建立。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook

    src_ws = _get_sheet(excel, source_sheet)

    if dest_sheet:
        try:
            dst_ws = wb.Sheets(dest_sheet)
        except Exception:
            dst_ws = wb.Sheets.Add()
            dst_ws.Name = dest_sheet
    else:
        dst_ws = excel.ActiveSheet

    src_ws.Range(source_range).Copy(dst_ws.Range(dest_range))

    src_label = f"{source_sheet}!{source_range}" if source_sheet else source_range
    dst_label = f"{dest_sheet}!{dest_range}"     if dest_sheet   else dest_range
    return {"status": "ok", "copied": f"{src_label} → {dst_label}"}



def add_comment(
    range_addr: str,
    comment: str,
    author: str | None = None,
    visible: bool = False,
    sheet: str | None = None,
) -> dict:
    """
    在儲存格加入批注（小紅三角提示）。
    comment：批注內文。
    author：作者名稱（省略則使用 Excel 登入使用者名稱）。
    visible：True = 批注常駐顯示；False = 滑鼠移上去才顯示（預設）。
    若儲存格已有批注，會先刪除再重新新增。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    cell = ws.Range(range_addr).Cells(1, 1)

    # 清除既有批注
    try:
        if cell.Comment is not None:
            cell.Comment.Delete()
    except Exception:
        pass

    c = cell.AddComment(comment)
    if author:
        try:
            c.Author = author
        except Exception:
            pass
    c.Visible = visible

    return {
        "status":  "ok",
        "cell":    cell.Address,
        "comment": comment,
        "visible": visible,
    }



def set_data_validation(
    range_addr: str,
    options: str | list,
    title: str | None = None,
    message: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    設定下拉選單資料驗證。
    options：選項清單（list）或直接引用範圍字串（如 "$E$1:$E$5"）
    title/message：選填，選取儲存格時顯示的提示標題與說明
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    if isinstance(options, list):
        formula = ",".join(str(o) for o in options)
    else:
        formula = str(options)  # 範圍參照

    rng.Validation.Delete()
    # xlValidateList=3, xlValidAlertStop=1, xlBetween=1
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula)
    rng.Validation.ShowError = True
    rng.Validation.ShowInput = bool(title or message)
    if title:
        rng.Validation.InputTitle = title
    if message:
        rng.Validation.InputMessage = message

    return {"status": "ok", "range": range_addr, "options": options}


# ── V4 新增：美化工具群 ──────────────────────────────────────────────────────



def summarize_range(
    range_addr: str,
    stats: list | None = None,
    sheet: str | None = None,
) -> dict:
    """
    對範圍內的數值儲存格計算統計摘要，直接回傳結果給 LLM（不寫入 Excel）。
    stats：指定要計算哪些統計項目；省略則全部計算。
    可選項：sum / average / max / min / count / stdev / median / count_all / count_blank
    """
    import statistics as _stats

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    raw = ws.Range(range_addr).Value
    values = _normalize_values(raw)

    all_cells  = [v for row in values for v in row]
    numbers    = [v for v in all_cells if isinstance(v, (int, float)) and v is not None]
    blank_count = sum(1 for v in all_cells if v is None or v == "")

    want = set(stats) if stats else {
        "sum", "average", "max", "min", "count", "stdev", "median",
        "count_all", "count_blank",
    }

    result: dict = {"status": "ok", "range": range_addr}

    if "count_all" in want:
        result["count_all"] = len(all_cells)
    if "count_blank" in want:
        result["count_blank"] = blank_count
    if "count" in want:
        result["count"] = len(numbers)

    if not numbers:
        result["note"] = "範圍內無數值，無法計算統計"
        return result

    if "sum" in want:
        result["sum"] = sum(numbers)
    if "average" in want:
        result["average"] = round(sum(numbers) / len(numbers), 6)
    if "max" in want:
        result["max"] = max(numbers)
    if "min" in want:
        result["min"] = min(numbers)
    if "median" in want:
        result["median"] = _stats.median(numbers)
    if "stdev" in want and len(numbers) > 1:
        result["stdev"] = round(_stats.stdev(numbers), 6)

    return result



def find_duplicates(
    range_addr: str,
    column_index: int = 1,
    action: str = "mark",
    mark_color: str = "#FFFF00",
    sheet: str | None = None,
) -> dict:
    """
    找出並處理重複值。
    column_index：以範圍內第幾欄判斷重複（1-based）。
    action：
      mark   = 標記重複列（背景色）
      delete = 刪除重複列（保留第一次出現）
      list   = 只回傳重複值清單，不修改工作表
    mark_color：標記顏色 #RRGGBB（action=mark 時使用）。
    """
    from collections import Counter

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    values = _normalize_values(rng.Value)

    start_row = rng.Cells(1, 1).Row
    start_col = rng.Cells(1, 1).Column
    col_idx   = column_index - 1  # 轉為 0-based

    col_vals = [
        row[col_idx] for row in values
        if len(row) > col_idx
    ]
    counts     = Counter(col_vals)
    dup_values = {v for v, c in counts.items() if c > 1 and v is not None and v != ""}

    if action == "list":
        return {"status": "ok", "duplicates": list(dup_values), "count": len(dup_values)}

    if action == "mark":
        bgr    = _hex_to_bgr(mark_color)
        marked = 0
        for r_idx, row in enumerate(values):
            if len(row) > col_idx and row[col_idx] in dup_values:
                cell_row = start_row + r_idx
                for c_idx in range(len(row)):
                    ws.Cells(cell_row, start_col + c_idx).Interior.Color = bgr
                marked += 1
        return {
            "status": "ok", "action": "mark",
            "marked_rows": marked, "duplicate_values": list(dup_values),
        }

    if action == "delete":
        to_delete = []
        seen: set = set()
        for r_idx, row in enumerate(values):
            if len(row) > col_idx:
                val = row[col_idx]
                if val is not None and val != "":
                    if val in seen:
                        to_delete.append(start_row + r_idx)
                    else:
                        seen.add(val)
        for row_num in reversed(to_delete):
            ws.Rows(row_num).Delete()
        return {"status": "ok", "action": "delete", "deleted_rows": len(to_delete)}

    raise InvalidToolArgumentsError("action 必須是 mark / delete / list")



def fill_series(
    start_cell: str,
    count: int,
    series_type: str = "number",
    step: float = 1,
    start_value: float | str | None = None,
    direction: str = "down",
    sheet: str | None = None,
) -> dict:
    """
    自動填充序列（數字 / 日期 / 平日 / 月份）。
    start_cell：起始儲存格，如 'A1'。
    count：填充幾個（含起始格）。
    series_type：number / date / weekday / month / year。
    step：每次遞增量；date 系列以天/平日/月/年為單位。
    start_value：起始值（省略則使用起始格現有值）。
    direction：down（向下，預設）/ right（向右）。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    _ensure_positive_int("count", count)

    start = ws.Range(start_cell).Cells(1, 1)

    if start_value is not None:
        start.Value = start_value

    if direction == "down":
        end_cell  = ws.Cells(start.Row + count - 1, start.Column)
        rowcol    = XL_SERIES_COLUMNS
    else:
        end_cell  = ws.Cells(start.Row, start.Column + count - 1)
        rowcol    = XL_SERIES_ROWS

    rng = ws.Range(start, end_cell)

    type_map = {
        "number":  (XL_SERIES_LINEAR, None),
        "date":    (XL_SERIES_DATE,   XL_DATE_DAY),
        "weekday": (XL_SERIES_DATE,   XL_DATE_WEEKDAY),
        "month":   (XL_SERIES_DATE,   XL_DATE_MONTH),
        "year":    (XL_SERIES_DATE,   XL_DATE_YEAR),
    }
    xl_type, date_unit = type_map.get(series_type.lower(), (XL_SERIES_LINEAR, None))

    kwargs: dict = {
        "Rowcol": rowcol,
        "Type":   xl_type,
        "Step":   step,
    }
    if date_unit is not None:
        kwargs["Date"] = date_unit

    rng.DataSeries(**kwargs)

    return {
        "status":      "ok",
        "range":       rng.Address,
        "series_type": series_type,
        "count":       count,
        "step":        step,
    }



def transpose_range(
    source_range: str,
    dest_cell: str,
    source_sheet: str | None = None,
    dest_sheet: str | None = None,
) -> dict:
    """
    將範圍行列轉置後寫入到新位置（列變欄、欄變列）。
    dest_cell：目標左上角起始儲存格。
    注意：只複製數值，不複製格式（格式請用 format_range 另行套用）。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook

    src_ws = _get_sheet(excel, source_sheet)

    if dest_sheet:
        try:
            dst_ws = wb.Sheets(dest_sheet)
        except Exception:
            dst_ws = wb.Sheets.Add()
            dst_ws.Name = dest_sheet
    else:
        dst_ws = excel.ActiveSheet

    raw    = src_ws.Range(source_range).Value
    values = _normalize_values(raw)

    if not values:
        return {"status": "ok", "note": "來源範圍為空"}

    rows_in = len(values)
    cols_in = max(len(r) for r in values)

    # 轉置
    transposed = [
        [values[r][c] if c < len(values[r]) else None for r in range(rows_in)]
        for c in range(cols_in)
    ]

    dest_start = dst_ws.Range(dest_cell).Cells(1, 1)
    rows_out   = len(transposed)
    cols_out   = rows_in  # 原本的行數變成欄數

    for r_idx, row in enumerate(transposed):
        for c_idx, val in enumerate(row):
            cell = dst_ws.Cells(dest_start.Row + r_idx, dest_start.Column + c_idx)
            if isinstance(val, str) and val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val

    dest_rng = dst_ws.Range(
        dst_ws.Cells(dest_start.Row, dest_start.Column),
        dst_ws.Cells(dest_start.Row + rows_out - 1, dest_start.Column + cols_out - 1),
    )
    return {
        "status":       "ok",
        "transposed_to": dest_rng.Address,
        "original":     f"{rows_in}列×{cols_in}欄",
        "result":       f"{rows_out}列×{cols_out}欄",
    }



def name_range(
    range_addr: str,
    name: str,
    sheet: str | None = None,
) -> dict:
    """
    為指定範圍建立具名範圍（Named Range）。
    命名後可在公式中直接使用名稱，如 =SUM(銷售額) 取代 =SUM(B2:B100)。
    name：只能包含字母、數字和底線，不可以數字開頭。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    wb = excel.ActiveWorkbook
    rng = ws.Range(range_addr)

    try:
        wb.Names.Add(Name=name, RefersTo=rng)
    except Exception as e:
        raise InvalidToolArgumentsError(
            f"無法建立具名範圍 '{name}'：{e}。"
            "名稱只能包含字母/數字/底線，且不可與現有名稱重複。"
        ) from e

    return {"status": "ok", "name": name, "refers_to": rng.Address}



def add_subtotal(
    range_addr: str,
    group_by_column: int,
    value_columns: list,
    function_type: str = "sum",
    sheet: str | None = None,
) -> dict:
    """
    依指定欄位自動插入分組小計列。
    ⚠️ 執行前資料必須已依 group_by_column 排序（sort_range），否則小計結果不正確。
    group_by_column：依第幾欄分組（範圍內 1-based），每次這欄值改變就插入小計列。
    value_columns：要加總的欄位索引清單（範圍內 1-based），如 [3, 4]。
    function_type：sum / count / average / max / min。
    """
    func_map = {
        "sum":     XL_SUM,
        "count":   XL_COUNT,
        "average": XL_AVERAGE,
        "max":     XL_MAX,
        "min":     XL_MIN_FUNC,
    }
    xl_func = func_map.get(function_type.lower())
    if xl_func is None:
        raise InvalidToolArgumentsError("function_type 必須是 sum / count / average / max / min")

    _ensure_positive_int("group_by_column", group_by_column)
    if not value_columns:
        raise InvalidToolArgumentsError("value_columns 不可為空")

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    rng.Subtotal(
        GroupBy=group_by_column,
        Function=xl_func,
        TotalList=value_columns,
        Replace=True,
        PageBreaks=False,
        SummaryBelowData=True,
    )

    return {
        "status":         "ok",
        "grouped_by_col": group_by_column,
        "value_cols":     value_columns,
        "function":       function_type,
    }



def advanced_filter(
    range_addr: str,
    criteria_range: str | None = None,
    dest_range: str | None = None,
    unique_only: bool = False,
    sheet: str | None = None,
) -> dict:
    """
    進階篩選：多條件篩選，可選擇就地篩選或複製結果到新位置。
    criteria_range：條件範圍（第一列為欄標題，需與資料標題完全一致；第二列起為條件值）。
    dest_range：複製目標起始格（填寫則結果複製到此；省略則就地篩選）。
    unique_only：True = 只保留不重複記錄。
    條件格式範例：同一列的條件為 AND，不同列為 OR。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    crit_rng = ws.Range(criteria_range) if criteria_range else None

    if dest_range:
        dest_rng = ws.Range(dest_range)
        rng.AdvancedFilter(
            Action=XL_FILTER_COPY,
            CriteriaRange=crit_rng,
            CopyToRange=dest_rng,
            Unique=unique_only,
        )
        result_location = dest_range
    else:
        rng.AdvancedFilter(
            Action=XL_FILTER_IN_PLACE,
            CriteriaRange=crit_rng,
            Unique=unique_only,
        )
        result_location = range_addr

    return {
        "status":        "ok",
        "criteria":      criteria_range,
        "result_at":     result_location,
        "unique_only":   unique_only,
    }



def split_text_to_columns(
    range_addr: str,
    delimiter: str = ",",
    sheet: str | None = None,
) -> dict:
    """
    依分隔符將文字欄位分割成多欄（等同 Excel「資料→資料剖析」）。
    ⚠️ 會直接覆蓋右側相鄰欄位，執行前請確認右側有足夠空白欄。
    delimiter：分隔符，支援 comma（逗號）/ tab（定位字元）/ semicolon（分號）/
               space（空格）/ 或任意單一字元（如 '|'、'-'）。
    常見用途：把 '台北,10000,業務' 這樣的單欄拆成三欄。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    preset = {"comma", "tab", "semicolon", "space"}
    is_preset = delimiter.lower() in preset

    rng.TextToColumns(
        DataType=XL_DELIMITED,
        TextQualifier=1,           # xlTextQualifierDoubleQuote
        ConsecutiveDelimiter=False,
        Tab=(delimiter.lower() == "tab"),
        Semicolon=(delimiter.lower() == "semicolon"),
        Comma=(delimiter.lower() == "comma"),
        Space=(delimiter.lower() == "space"),
        Other=not is_preset,
        OtherChar=delimiter if not is_preset else "",
    )
    return {"status": "ok", "range": range_addr, "delimiter": delimiter, "sheet": sheet}


# ── 工作表快照（突破 undo 20 步限制）────────────────────────────────────────

