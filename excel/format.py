"""excel/format.py — auto-split from excel_tools.py."""
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
    _get_excel, _get_sheet, _hex_to_bgr, _ensure_positive_int,
    _ensure_positive_number, _normalize_values, _com_tls,
)

_INTERIOR_NONE_IDX = -4142  # xlColorIndexNone
_ALL_BORDER_IDX = sorted({idx for indexes in XL_BORDER_SIDES.values() for idx in indexes})
_XL_CENTER_CONTINUOUS = 7
_XL_VERTICAL_CENTER = -4108

_BEAUTIFY_THEMES = {
    "blue": {
        "header_fill": "#4472C4",
        "header_font": "#FFFFFF",
        "stripe_fill": "#D9EAF7",
        "border_color": "#B4C6E7",
        "accent": "#1F4E79",
    },
    "green": {
        "header_fill": "#70AD47",
        "header_font": "#FFFFFF",
        "stripe_fill": "#E2F0D9",
        "border_color": "#A9D18E",
        "accent": "#548235",
    },
    "gray": {
        "header_fill": "#595959",
        "header_font": "#FFFFFF",
        "stripe_fill": "#F2F2F2",
        "border_color": "#BFBFBF",
        "accent": "#404040",
    },
    "orange": {
        "header_fill": "#ED7D31",
        "header_font": "#FFFFFF",
        "stripe_fill": "#FCE4D6",
        "border_color": "#F4B183",
        "accent": "#C65911",
    },
    "purple": {
        "header_fill": "#7030A0",
        "header_font": "#FFFFFF",
        "stripe_fill": "#EADCF8",
        "border_color": "#C9B2E2",
        "accent": "#5F249F",
    },
}


def capture_widths_before(col_start: int, count: int = 1, sheet: str | None = None) -> dict:
    """
    Read column widths (in points) before set_column_width / auto_fit so they
    can be restored by undo.  Returns {col_index: width}.
    """
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)
    result = {}
    for i in range(col_start, col_start + count):
        try:
            result[i] = ws.Columns(i).ColumnWidth
        except Exception:
            pass
    return result



def capture_heights_before(row_start: int, count: int = 1, sheet: str | None = None) -> dict:
    """
    Read row heights (in points) before set_row_height so they can be restored.
    Returns {row_index: height}.
    """
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)
    result = {}
    for i in range(row_start, row_start + count):
        try:
            result[i] = ws.Rows(i).RowHeight
        except Exception:
            pass
    return result



def capture_formats_before(range_addr: str, sheet: str | None, args: dict) -> dict:
    """
    Phase 3：在 format_range / set_borders 執行前讀取目前格式，供 undo_last 回寫。
    回傳結構：
      {"type": "format_range", "cells": [{address, bold, italic, color, fill, font_size, number_format, h_align}, ...]}
      {"type": "set_borders",  "cells": [{address, borders: {idx: {line_style, color}}}, ...]}

    效能保護：超過 200 格時改為範圍級別備份（適用均勻格式的常見場景）。
    """
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)
    rng   = ws.Range(range_addr)
    cell_count = rng.Count

    tool_type = args.get("_tool_type", "format_range")

    if tool_type == "beautify_range":
        props = {
            "bold", "italic", "color", "fill",
            "font_size", "number_format", "horizontal_alignment",
        }
        if cell_count <= 200:
            cells_data = []
            for cell in rng:
                info = {
                    "address": cell.Address,
                    "bold": cell.Font.Bold,
                    "italic": cell.Font.Italic,
                    "color": cell.Font.Color,
                    "font_size": cell.Font.Size,
                    "number_format": cell.NumberFormat,
                    "horizontal_alignment": cell.HorizontalAlignment,
                }
                ci = cell.Interior.ColorIndex
                info["fill"] = None if ci == _INTERIOR_NONE_IDX else cell.Interior.Color
                borders_info = {}
                for idx in _ALL_BORDER_IDX:
                    try:
                        b = cell.Borders(idx)
                        borders_info[idx] = {
                            "line_style": b.LineStyle,
                            "color": b.Color,
                        }
                    except Exception:
                        pass
                info["borders"] = borders_info
                cells_data.append(info)
        else:
            info = {
                "address": rng.Address,
                "_range_level": True,
                "bold": rng.Font.Bold,
                "italic": rng.Font.Italic,
                "color": rng.Font.Color,
                "font_size": rng.Font.Size,
                "number_format": rng.NumberFormat,
                "horizontal_alignment": rng.HorizontalAlignment,
            }
            ci = rng.Interior.ColorIndex
            info["fill"] = None if ci == _INTERIOR_NONE_IDX else rng.Interior.Color
            borders_info = {}
            for idx in _ALL_BORDER_IDX:
                try:
                    b = rng.Borders(idx)
                    borders_info[idx] = {"line_style": b.LineStyle, "color": b.Color}
                except Exception:
                    pass
            info["borders"] = borders_info
            cells_data = [info]
        return {"type": "beautify_range", "range": range_addr, "cells": cells_data}

    if tool_type == "set_borders":
        # 讀取即將被覆蓋的邊框設定
        if cell_count <= 200:
            cells_data = []
            for cell in rng:
                borders_info = {}
                for idx in _ALL_BORDER_IDX:
                    try:
                        b = cell.Borders(idx)
                        borders_info[idx] = {
                            "line_style": b.LineStyle,
                            "color":      b.Color,
                        }
                    except Exception:
                        pass
                cells_data.append({"address": cell.Address, "borders": borders_info})
        else:
            # 大範圍：只記範圍級別（均勻格式場景）
            borders_info = {}
            for idx in _ALL_BORDER_IDX:
                try:
                    b = rng.Borders(idx)
                    borders_info[idx] = {"line_style": b.LineStyle, "color": b.Color}
                except Exception:
                    pass
            cells_data = [{"address": rng.Address, "borders": borders_info, "_range_level": True}]
        return {"type": "set_borders", "range": range_addr, "cells": cells_data}

    else:  # format_range
        # 只讀 args 中實際傳入的屬性，避免無謂讀取
        props = {k for k in ("bold", "italic", "color", "fill",
                              "font_size", "number_format", "horizontal_alignment")
                 if args.get(k) is not None}
        if not props:
            return {}  # 沒有任何屬性需要備份

        if cell_count <= 200:
            cells_data = []
            for cell in rng:
                info: dict = {"address": cell.Address}
                if "bold" in props:
                    info["bold"] = cell.Font.Bold
                if "italic" in props:
                    info["italic"] = cell.Font.Italic
                if "color" in props:
                    info["color"] = cell.Font.Color
                if "fill" in props:
                    ci = cell.Interior.ColorIndex
                    info["fill"] = None if ci == _INTERIOR_NONE_IDX else cell.Interior.Color
                if "font_size" in props:
                    info["font_size"] = cell.Font.Size
                if "number_format" in props:
                    info["number_format"] = cell.NumberFormat
                if "horizontal_alignment" in props:
                    info["horizontal_alignment"] = cell.HorizontalAlignment
                cells_data.append(info)
        else:
            # 大範圍：範圍均勻值
            info: dict = {"address": rng.Address, "_range_level": True}
            if "bold" in props:
                info["bold"] = rng.Font.Bold
            if "italic" in props:
                info["italic"] = rng.Font.Italic
            if "color" in props:
                info["color"] = rng.Font.Color
            if "fill" in props:
                ci = rng.Interior.ColorIndex
                info["fill"] = None if ci == _INTERIOR_NONE_IDX else rng.Interior.Color
            if "font_size" in props:
                info["font_size"] = rng.Font.Size
            if "number_format" in props:
                info["number_format"] = rng.NumberFormat
            if "horizontal_alignment" in props:
                info["horizontal_alignment"] = rng.HorizontalAlignment
            cells_data = [info]

        return {"type": "format_range", "range": range_addr, "cells": cells_data}



def _restore_formats(formats_before: dict, sheet: str | None) -> None:
    """Phase 3：依 formats_before 回寫格式（供 undo_last 呼叫）。"""
    excel = _get_excel()
    ws    = _get_sheet(excel, sheet)
    fmt_type = formats_before.get("type")

    for cell_info in formats_before.get("cells", []):
        addr = cell_info["address"]
        rng  = ws.Range(addr)

        if fmt_type in ("set_borders", "beautify_range") and "borders" in cell_info:
            for idx, bdata in cell_info.get("borders", {}).items():
                try:
                    b = rng.Borders(int(idx))
                    b.LineStyle = bdata["line_style"]
                    b.Color     = bdata["color"]
                except Exception:
                    pass

        if fmt_type == "set_borders":
            continue

        # format_range / beautify_range
        if fmt_type in ("format_range", "beautify_range"):
            if "bold" in cell_info and cell_info["bold"] is not None:
                rng.Font.Bold = cell_info["bold"]
            if "italic" in cell_info and cell_info["italic"] is not None:
                rng.Font.Italic = cell_info["italic"]
            if "color" in cell_info and cell_info["color"] is not None:
                rng.Font.Color = cell_info["color"]
            if "fill" in cell_info:
                if cell_info["fill"] is None:
                    rng.Interior.ColorIndex = _INTERIOR_NONE_IDX
                else:
                    rng.Interior.Color = cell_info["fill"]
            if "font_size" in cell_info and cell_info["font_size"] is not None:
                rng.Font.Size = cell_info["font_size"]
            if "number_format" in cell_info and cell_info["number_format"] is not None:
                rng.NumberFormat = cell_info["number_format"]
            if "horizontal_alignment" in cell_info and cell_info["horizontal_alignment"] is not None:
                rng.HorizontalAlignment = cell_info["horizontal_alignment"]


# ── 讀取 ──────────────────────────────────────────────────────────────────────


def format_range(
    range_addr: str,
    sheet: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    fill: str | None = None,
    font_size: float | None = None,
    number_format: str | None = None,
    horizontal_alignment: str | None = None,
) -> dict:
    """設定儲存格格式（★ B3：參數改為 range_addr）"""
    # 防止 no-op：只有 range 沒有任何格式屬性時，回傳可辨識錯誤給 LLM。
    if all(
        v is None for v in (
            bold, italic, color, fill, font_size, number_format, horizontal_alignment
        )
    ):
        raise InvalidToolArgumentsError(
            "format_range 至少要提供一項格式參數："
            "bold/italic/color/fill/font_size/number_format/horizontal_alignment"
        )

    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    if bold is not None:
        rng.Font.Bold = bold
    if italic is not None:
        rng.Font.Italic = italic
    if color is not None:
        rng.Font.Color = _hex_to_bgr(color)
    if fill is not None:
        rng.Interior.Color = _hex_to_bgr(fill)
    if font_size is not None:
        rng.Font.Size = font_size
    if number_format is not None:
        rng.NumberFormat = number_format
    if horizontal_alignment is not None:
        rng.HorizontalAlignment = XL_ALIGN.get(horizontal_alignment, XL_ALIGN["center"])

    return {"status": "ok", "range": rng.Address}


def _beautify_theme(theme: str) -> dict:
    return _BEAUTIFY_THEMES.get((theme or "blue").lower(), _BEAUTIFY_THEMES["blue"])


def _is_numeric_value(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _guess_column_number_format(header: str, values: list) -> str | None:
    non_blank = [v for v in values if v not in (None, "")]
    if not non_blank or not all(_is_numeric_value(v) for v in non_blank):
        return None

    header_text = (header or "").lower()
    if any(token in header_text for token in ("%", "rate", "ratio", "percent", "比例", "比率", "率")):
        return "0.0%"
    if any(
        token in header_text
        for token in ("amount", "price", "sales", "revenue", "cost", "total", "金額", "收入", "成本", "總計", "合計")
    ):
        return "#,##0"
    if any(float(v) % 1 for v in non_blank):
        return "#,##0.00"
    return "#,##0"


def _looks_like_header_row(first_row: list, second_row: list | None = None) -> bool:
    non_blank = [str(value).strip() for value in first_row if value not in (None, "")]
    if not non_blank or len(non_blank) != len(set(non_blank)):
        return False
    if not all(isinstance(value, str) and str(value).strip() for value in first_row):
        return False
    if second_row and any(_is_numeric_value(value) for value in second_row):
        return True
    return len(non_blank) >= 2


def beautify_range(
    range_addr: str,
    sheet: str | None = None,
    theme: str = "blue",
    has_header: bool = True,
    banded_rows: bool = True,
    auto_fit_columns: bool = True,
    freeze_header: bool = False,
    apply_filter: bool = True,
    number_format: str | None = "auto",
    font_name: str = "Calibri",
) -> dict:
    """
    一鍵美化資料範圍：表頭、交錯列、框線、欄寬、數字格式、可選篩選與凍結。
    不建立 ListObject，避免和既有表格重疊時失敗。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    rows = int(rng.Rows.Count)
    cols = int(rng.Columns.Count)
    if rows < 1 or cols < 1:
        raise InvalidToolArgumentsError("beautify_range 的 range_addr 必須至少包含一個儲存格")

    theme_key = (theme or "blue").lower()
    palette = _beautify_theme(theme_key)
    first_row = int(rng.Row)
    first_col = int(rng.Column)
    last_row = first_row + rows - 1
    last_col = first_col + cols - 1

    rng.Font.Name = font_name
    rng.Font.Size = 11
    rng.VerticalAlignment = _XL_VERTICAL_CENTER

    border_color = _hex_to_bgr(palette["border_color"])
    for idx in XL_BORDER_SIDES["all"]:
        border = rng.Borders(idx)
        border.LineStyle = XL_BORDER_STYLE["thin"]
        border.Color = border_color

    applied: list[str] = ["font", "borders"]

    if has_header:
        header_rng = ws.Range(ws.Cells(first_row, first_col), ws.Cells(first_row, last_col))
        header_rng.Font.Bold = True
        header_rng.Font.Color = _hex_to_bgr(palette["header_font"])
        header_rng.Interior.Color = _hex_to_bgr(palette["header_fill"])
        header_rng.HorizontalAlignment = XL_ALIGN["center"]
        header_rng.VerticalAlignment = _XL_VERTICAL_CENTER
        header_rng.RowHeight = max(float(header_rng.RowHeight), 22)
        applied.append("header")

    data_start_row = first_row + 1 if has_header and rows > 1 else first_row
    if data_start_row <= last_row:
        data_rng = ws.Range(ws.Cells(data_start_row, first_col), ws.Cells(last_row, last_col))
        data_rng.Interior.ColorIndex = _INTERIOR_NONE_IDX
        data_rng.HorizontalAlignment = XL_ALIGN["left"]
        if banded_rows:
            stripe_color = _hex_to_bgr(palette["stripe_fill"])
            for row_idx in range(data_start_row, last_row + 1):
                if (row_idx - data_start_row) % 2 == 1:
                    ws.Range(ws.Cells(row_idx, first_col), ws.Cells(row_idx, last_col)).Interior.Color = stripe_color
            applied.append("banded_rows")

        if number_format and str(number_format).lower() not in ("none", "off", "false"):
            if str(number_format).lower() == "auto":
                header_values = _normalize_values(
                    ws.Range(ws.Cells(first_row, first_col), ws.Cells(first_row, last_col)).Value
                )[0]
                formatted_cols = []
                for offset in range(cols):
                    col_idx = first_col + offset
                    col_rng = ws.Range(ws.Cells(data_start_row, col_idx), ws.Cells(last_row, col_idx))
                    values = [row[0] for row in _normalize_values(col_rng.Value)]
                    fmt = _guess_column_number_format(str(header_values[offset] or ""), values)
                    if fmt:
                        col_rng.NumberFormat = fmt
                        formatted_cols.append(_col_letter(col_idx))
                if formatted_cols:
                    applied.append(f"number_format:{','.join(formatted_cols)}")
            else:
                data_rng.NumberFormat = str(number_format)
                applied.append("number_format")

    if apply_filter and has_header:
        try:
            ws.Parent.Activate()
            ws.Activate()
            if not ws.AutoFilterMode:
                rng.AutoFilter(Field=1)
            applied.append("filter")
        except Exception:
            pass

    if auto_fit_columns:
        rng.Columns.AutoFit()
        applied.append("auto_fit_columns")

    if freeze_header and has_header:
        try:
            freeze_panes(row=first_row, col=0, sheet=ws.Name)
            applied.append("freeze_header")
        except Exception:
            pass

    return {
        "status": "ok",
        "tool": "beautify_range",
        "sheet": ws.Name,
        "range": rng.Address,
        "theme": theme_key if theme_key in _BEAUTIFY_THEMES else "blue",
        "rows": rows,
        "columns": cols,
        "applied": applied,
    }


# ── 列操作 ─────────────────────────────────────────────────────────────────────


def set_print_titles(
    rows: str | None = None,
    columns: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    設定列印時每頁重複的標題列/欄。
    rows：重複的列範圍，如 '$1:$1'（第 1 列）或 '$1:$2'（第 1~2 列）。
    columns：重複的欄範圍，如 '$A:$A'（A 欄）。
    rows 與 columns 可同時設定。省略則清除對應的重複設定。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ps = ws.PageSetup
    if rows is not None:
        ps.PrintTitleRows = rows
    if columns is not None:
        ps.PrintTitleColumns = columns
    return {
        "status":  "ok",
        "sheet":   ws.Name,
        "title_rows":    ps.PrintTitleRows,
        "title_columns": ps.PrintTitleColumns,
    }



def add_header_footer(
    header: str | None = None,
    footer: str | None = None,
    left_header: str | None = None,
    center_header: str | None = None,
    right_header: str | None = None,
    left_footer: str | None = None,
    center_footer: str | None = None,
    right_footer: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    設定工作表的頁首與頁尾。
    header：頁首完整文字（置中），或分別用 left/center/right 設定三區。
    footer：頁尾完整文字（置中），或分別用 left/center/right 設定三區。
    Excel 特殊碼：&P=頁碼、&N=總頁數、&D=日期、&T=時間、&F=檔名、&A=工作表名。
    常用頁尾：'第 &P 頁，共 &N 頁' 或 '&D  &F'。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ps = ws.PageSetup

    # 頁首
    if header is not None:
        ps.CenterHeader = header
    if left_header is not None:
        ps.LeftHeader = left_header
    if center_header is not None:
        ps.CenterHeader = center_header
    if right_header is not None:
        ps.RightHeader = right_header

    # 頁尾
    if footer is not None:
        ps.CenterFooter = footer
    if left_footer is not None:
        ps.LeftFooter = left_footer
    if center_footer is not None:
        ps.CenterFooter = center_footer
    if right_footer is not None:
        ps.RightFooter = right_footer

    return {"status": "ok", "sheet": ws.Name}


# ── 資料操作 ──────────────────────────────────────────────────────────────────


def freeze_panes(
    row: int = 0,
    col: int = 0,
    sheet: str | None = None,
) -> dict:
    """
    凍結列或欄。row=1 表示凍結第 1 列（標題行）；col=1 表示凍結第 1 欄。
    row=0, col=0 表示解除所有凍結。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Parent.Activate()
    ws.Activate()

    win = excel.ActiveWindow

    try:
        if win.FreezePanes:
            win.FreezePanes = False
        win.SplitRow = 0
        win.SplitColumn = 0

        if row > 0 or col > 0:
            win.SplitRow = row
            win.SplitColumn = col
            win.FreezePanes = True
    except Exception:
        # Some Excel window states reject direct FreezePanes assignment even
        # after activating the workbook.  The Excel 4 macro path mirrors the
        # UI command and succeeds in those cases.
        ws.Cells(max(row + 1, 1), max(col + 1, 1)).Select()
        excel.ExecuteExcel4Macro("FREEZE.PANES(FALSE)")
        if row > 0 or col > 0:
            ws.Cells(row + 1, col + 1).Select()
            excel.ExecuteExcel4Macro("FREEZE.PANES(TRUE)")

    return {"status": "ok", "frozen_rows": row, "frozen_cols": col}



def auto_fit(
    target: str = "columns",
    range_addr: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    自動調整欄寬或列高。
    target：'columns'（自動欄寬）/ 'rows'（自動列高）/ 'both'（兩者皆是）
    range_addr：省略則整張工作表全部調整。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)

    if range_addr:
        rng = ws.Range(range_addr)
        if target in ("columns", "both"):
            rng.Columns.AutoFit()
        if target in ("rows", "both"):
            rng.Rows.AutoFit()
    else:
        if target in ("columns", "both"):
            ws.Columns.AutoFit()
        if target in ("rows", "both"):
            ws.Rows.AutoFit()

    return {"status": "ok", "auto_fit": target, "range": range_addr or "全表"}



def set_column_width(
    column_index: int,
    width: float,
    sheet: str | None = None,
) -> dict:
    """設定指定欄的寬度（Excel 單位，約 8.43 = 預設寬度）"""
    _ensure_positive_int("column_index", column_index)
    _ensure_positive_number("width", width)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Columns(column_index).ColumnWidth = width
    return {"status": "ok", "column": column_index, "width": width}


# ── V3 新增工具 ───────────────────────────────────────────────────────────────


def merge_cells(range_addr: str, sheet: str | None = None) -> dict:
    """合併儲存格，並自動置中"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    rng.Merge()
    rng.HorizontalAlignment = XL_ALIGN["center"]
    return {"status": "ok", "merged": range_addr}



def unmerge_cells(range_addr: str, sheet: str | None = None) -> dict:
    """取消合併儲存格"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Range(range_addr).UnMerge()
    return {"status": "ok", "unmerged": range_addr}





def set_borders(
    range_addr: str,
    style: str = "thin",
    color: str = "#000000",
    sides: str = "all",
    sheet: str | None = None,
) -> dict:
    """
    設定儲存格框線。
    style：thin / medium / thick / dashed
    sides：all（全部）/ outer（外框）/ inner（內線）/ left / top / bottom / right
    color：#RRGGBB
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    line_style = XL_BORDER_STYLE.get(style, XL_BORDER_STYLE["thin"])
    bgr = _hex_to_bgr(color)
    indices = XL_BORDER_SIDES.get(sides, XL_BORDER_SIDES["all"])

    for idx in indices:
        border = rng.Borders(idx)
        border.LineStyle = line_style
        border.Color = bgr

    return {"status": "ok", "range": range_addr, "style": style, "sides": sides}



def set_row_height(
    row_index: int,
    height: float,
    sheet: str | None = None,
) -> dict:
    """設定列高（Excel 點數單位，預設約 15）"""
    _ensure_positive_int("row_index", row_index)
    _ensure_positive_number("height", height)
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(row_index).RowHeight = height
    return {"status": "ok", "row": row_index, "height": height}



def add_conditional_format(
    range_addr: str,
    condition_type: str,
    value: str | int | float | list,
    fill_color: str,
    font_color: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    新增條件格式化規則。
    condition_type：greater / less / equal / between / contains
    value：數值或字串；between 時為 [min, max]
    fill_color：背景色 #RRGGBB
    font_color：字色 #RRGGBB（選填）
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)

    # 清除既有規則再新增
    rng.FormatConditions.Delete()

    # xlCellValue=1  Operator: xlGreater=5, xlLess=6, xlEqual=3, xlBetween=1
    # xlTextString=6（含文字）
    if condition_type == "between" and isinstance(value, list):
        fc = rng.FormatConditions.Add(
            Type=1, Operator=1,
            Formula1=str(value[0]), Formula2=str(value[1])
        )
    elif condition_type == "contains":
        fc = rng.FormatConditions.Add(Type=6, Formula1=str(value))
    else:
        op_map = {"greater": 5, "less": 6, "equal": 3}
        op = op_map.get(condition_type, 5)
        fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=str(value))

    fc.Interior.Color = _hex_to_bgr(fill_color)
    if font_color:
        fc.Font.Color = _hex_to_bgr(font_color)

    return {"status": "ok", "condition": condition_type, "value": value, "range": range_addr}



def apply_table_style(
    range_addr: str,
    style: str = "blue",
    table_name: str | None = None,
    has_header: bool = True,
    show_totals: bool = False,
    sheet: str | None = None,
) -> dict:
    """
    將儲存格範圍轉換為 Excel 正式表格（ListObject），套用內建樣式。
    style：blue / light_blue / green / light_green / orange / red / purple /
           gray / white / dark_blue / dark_green / dark_red，或直接填 Excel 樣式名稱。
    功能：帶狀底色、標題粗體、自動篩選按鈕、可選 Total 列。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    rng = ws.Range(range_addr)
    if not has_header and int(rng.Rows.Count) >= 2:
        values = _normalize_values(rng.Value)
        if values and _looks_like_header_row(values[0], values[1] if len(values) > 1 else None):
            has_header = True

    style_name = TABLE_STYLE_MAP.get(style.lower(), style)

    try:
        lo = ws.ListObjects.Add(
            SourceType=1,          # xlSrcRange
            Source=rng,
            XlListObjectHasHeaders=1 if has_header else 2,
        )
    except Exception as e:
        raise InvalidToolArgumentsError(
            f"無法建立表格（可能與現有表格重疊）：{e}"
        ) from e

    if table_name:
        try:
            lo.Name = table_name
        except Exception:
            pass  # 名稱衝突時沿用預設名稱

    lo.TableStyle = style_name

    if show_totals:
        lo.ShowTotals = True

    return {"status": "ok", "table": lo.Name, "style": style_name, "range": lo.Range.Address}



def add_sparklines(
    data_range: str,
    sparkline_range: str,
    sparkline_type: str = "line",
    color: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    在指定儲存格範圍插入 Sparkline 迷你圖。
    data_range：資料來源範圍（行數須與 sparkline_range 相符）。
    sparkline_range：迷你圖放置位置，如 'H2:H20'（每格一條迷你圖）。
    sparkline_type：line（折線）/ column（直條）/ winloss（盈虧）。
    color：迷你圖顏色 #RRGGBB（選填）。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)

    type_map = {
        "line":    XL_SPARKLINE_LINE,
        "column":  XL_SPARKLINE_COLUMN,
        "winloss": XL_SPARKLINE_WINLOSS,
    }
    xl_type = type_map.get(sparkline_type.lower(), XL_SPARKLINE_LINE)

    dest_rng = ws.Range(sparkline_range)
    grp = dest_rng.SparklineGroups.Add(Type=xl_type, SourceData=data_range)

    if color and grp is not None:
        try:
            grp.SeriesColor.Color = _hex_to_bgr(color)
        except Exception:
            pass

    return {
        "status":         "ok",
        "sparkline_type": sparkline_type,
        "data_range":     data_range,
        "sparkline_range": sparkline_range,
    }



def set_tab_color(
    color: str,
    sheet: str | None = None,
) -> dict:
    """
    設定工作表標籤顏色（Excel 底部分頁標籤）。
    color：#RRGGBB，例如 '#FF0000'（紅）、'#4472C4'（藍）。
    多個工作表不同顏色可一目了然區分用途。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Tab.Color = _hex_to_bgr(color)
    return {"status": "ok", "sheet": ws.Name, "color": color}



def page_setup(
    orientation: str = "portrait",
    paper_size: str = "a4",
    fit_to_wide: int | None = None,
    fit_to_tall: int | None = None,
    print_area: str | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
    sheet: str | None = None,
) -> dict:
    """
    設定工作表的頁面格式（列印方向、紙張、縮放、列印範圍等）。
    orientation：portrait（直印）/ landscape（橫印）。
    paper_size：a4 / letter / a3 / a5 / legal。
    fit_to_wide / fit_to_tall：縮放成幾頁寬 / 幾頁高（設定後 Zoom 自動關閉）。
    print_area：列印範圍，如 'A1:H50'；省略表示整張工作表。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ps = ws.PageSetup

    ps.Orientation = XL_PORTRAIT if orientation.lower() == "portrait" else XL_LANDSCAPE
    ps.PaperSize   = XL_PAPER_SIZE.get(paper_size.lower(), XL_PAPER_SIZE["a4"])

    if print_area is not None:
        ps.PrintArea = print_area

    if fit_to_wide is not None or fit_to_tall is not None:
        ps.Zoom = False
        if fit_to_wide is not None:
            ps.FitToPagesWide = False if fit_to_wide == 0 else fit_to_wide
        if fit_to_tall is not None:
            ps.FitToPagesTall = False if fit_to_tall == 0 else fit_to_tall

    if center_horizontally is not None:
        ps.CenterHorizontally = center_horizontally
    if center_vertically is not None:
        ps.CenterVertically = center_vertically

    return {
        "status":      "ok",
        "orientation": orientation,
        "paper_size":  paper_size,
        "print_area":  print_area,
    }



def add_image(
    image_path: str,
    range_addr: str,
    width: float | None = None,
    height: float | None = None,
    sheet: str | None = None,
) -> dict:
    """
    插入圖片到工作表中。圖片左上角對齊指定儲存格。
    image_path：圖片的絕對路徑或相對路徑（支援 jpg / png / bmp / gif）。
    range_addr：圖片放置位置的起始儲存格，如 'A1' 或 'D5'。
    width / height：圖片尺寸（點數）；省略則使用圖片原始尺寸。
    """
    import os
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)

    abs_path = os.path.abspath(image_path)
    if not os.path.exists(abs_path):
        raise InvalidToolArgumentsError(f"找不到圖片檔案：{abs_path}")

    cell = ws.Range(range_addr).Cells(1, 1)
    w = width  if width  is not None else -1
    h = height if height is not None else -1

    shape = ws.Shapes.AddPicture(
        Filename=abs_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=cell.Left,
        Top=cell.Top,
        Width=w,
        Height=h,
    )

    return {
        "status":   "ok",
        "image":    image_path,
        "position": range_addr,
        "shape":    shape.Name,
    }


# ── V4 分析工具群 ─────────────────────────────────────────────────────────────
