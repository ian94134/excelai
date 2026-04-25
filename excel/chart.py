"""excel/chart.py — auto-split from excel_tools.py."""
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
from excel._base import _get_excel, _get_sheet, _hex_to_bgr, _com_tls

def delete_chart(
    chart_index: int = 1,
    sheet: str | None = None,
) -> dict:
    """
    刪除工作表上指定的圖表。
    chart_index：第幾個圖表（1-based）；預設刪除第 1 個。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    total = ws.ChartObjects().Count
    if total == 0:
        raise InvalidToolArgumentsError("此工作表上沒有圖表可刪除")
    if chart_index < 1 or chart_index > total:
        raise InvalidToolArgumentsError(
            f"chart_index={chart_index} 超出範圍（此工作表共 {total} 個圖表）"
        )
    name = ws.ChartObjects(chart_index).Name
    ws.ChartObjects(chart_index).Delete()
    return {"status": "ok", "deleted_chart": name, "remaining": total - 1}



def move_chart(
    chart_index: int = 1,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    sheet: str | None = None,
) -> dict:
    """
    移動或調整工作表上圖表的位置與大小（點數單位）。
    chart_index：第幾個圖表（1-based）；預設第 1 個。
    left/top：圖表左上角距工作表左/上邊緣的距離。
    width/height：圖表的寬度與高度。
    省略任何參數則保持原值不變。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    total = ws.ChartObjects().Count
    if total == 0:
        raise InvalidToolArgumentsError("此工作表上沒有圖表")
    if chart_index < 1 or chart_index > total:
        raise InvalidToolArgumentsError(
            f"chart_index={chart_index} 超出範圍（共 {total} 個圖表）"
        )
    co = ws.ChartObjects(chart_index)
    if left   is not None: co.Left   = left
    if top    is not None: co.Top    = top
    if width  is not None: co.Width  = width
    if height is not None: co.Height = height
    return {
        "status": "ok",
        "chart":  co.Name,
        "left":   co.Left,
        "top":    co.Top,
        "width":  co.Width,
        "height": co.Height,
    }



def create_chart(
    range_addr: str,
    chart_type: str = "column",
    title: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    在工作表中建立圖表，圖表放在資料範圍右側。
    chart_type：column / bar / line / pie / area / scatter
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    data_range = ws.Range(range_addr)

    # 圖表放在資料範圍右側 10px 處
    left   = data_range.Left + data_range.Width + 10
    top    = data_range.Top
    width  = 400
    height = 280

    chart_obj = ws.ChartObjects().Add(Left=left, Top=top, Width=width, Height=height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=data_range)
    chart.ChartType = XL_CHART_TYPE.get(chart_type.lower(), XL_CHART_TYPE["column"])

    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    else:
        chart.HasTitle = False

    return {
        "status":     "ok",
        "chart_type": chart_type,
        "source":     range_addr,
        "title":      title,
    }


# ── 樞紐分析表 ────────────────────────────────────────────────────────────────


def create_pivot_table(
    source_range: str,
    dest_sheet: str,
    row_field: str,
    value_field: str,
    col_field: str | None = None,
    source_sheet: str | None = None,
) -> dict:
    """
    建立樞紐分析表。
    - source_range：來源資料範圍（含標題列），如 "A1:D100"
    - dest_sheet：放置樞紐的目標工作表名稱（不存在則自動建立）
    - row_field：列標籤欄位名稱（對應標題列的欄名）
    - value_field：值欄位名稱
    - col_field：欄標籤欄位名稱（可省略）
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook

    # 來源工作表與範圍
    src_ws = _get_sheet(excel, source_sheet)
    src_range = src_ws.Range(source_range)

    # 目標工作表：不存在則新建
    try:
        dest_ws = wb.Sheets(dest_sheet)
        # 清除舊內容
        dest_ws.Cells.Clear()
    except Exception:
        dest_ws = wb.Sheets.Add()
        dest_ws.Name = dest_sheet

    # 建立 PivotCache
    pivot_cache = wb.PivotCaches().Create(
        SourceType=XL_DATABASE,
        SourceData=src_range,
    )

    # 建立 PivotTable
    pt = pivot_cache.CreatePivotTable(
        TableDestination=dest_ws.Range("A3"),
        TableName=f"PivotTable_{int(time.time() * 1000)}",
    )

    # 列欄位
    pt.PivotFields(row_field).Orientation = XL_ROW_FIELD

    # 欄欄位（選填）
    if col_field:
        pt.PivotFields(col_field).Orientation = XL_COL_FIELD

    # 值欄位（加總）
    vf = pt.PivotFields(value_field)
    vf.Orientation = XL_DATA_FIELD
    vf.Function = XL_SUM

    dest_ws.Activate()
    return {
        "status":      "ok",
        "pivot_sheet": dest_sheet,
        "row_field":   row_field,
        "value_field": value_field,
        "col_field":   col_field,
    }



def refresh_pivot_table(
    pivot_sheet: str,
    sheet: str | None = None,
) -> dict:
    """
    重新整理樞紐分析表（更新來源資料後需執行）。
    pivot_sheet：樞紐分析表所在的工作表名稱。
    若工作表上有多個樞紐，全部一起重新整理。
    """
    excel = _get_excel()
    pv_ws = _get_sheet(excel, pivot_sheet)
    count = pv_ws.PivotTables().Count
    if count == 0:
        raise InvalidToolArgumentsError(f"工作表 '{pivot_sheet}' 上找不到樞紐分析表")
    for i in range(1, count + 1):
        pv_ws.PivotTables(i).RefreshTable()
    return {"status": "ok", "pivot_sheet": pivot_sheet, "refreshed": count}



def format_pivot_table(
    pivot_sheet: str,
    style: str = "PivotStyleMedium9",
    show_row_headers: bool = True,
    show_col_headers: bool = True,
    banded_rows: bool = True,
    banded_cols: bool = False,
) -> dict:
    """
    套用樞紐分析表的內建樣式與顯示選項。
    pivot_sheet：樞紐所在工作表；若有多個樞紐，套用到第一個。
    style：樞紐樣式名稱，如 'PivotStyleMedium9'（藍）、'PivotStyleMedium4'（橘）、
           'PivotStyleDark1'（深色）。
    banded_rows：帶狀列底色（交錯），建議 True。
    """
    excel = _get_excel()
    pv_ws = _get_sheet(excel, pivot_sheet)
    count = pv_ws.PivotTables().Count
    if count == 0:
        raise InvalidToolArgumentsError(f"工作表 '{pivot_sheet}' 上找不到樞紐分析表")
    pt = pv_ws.PivotTables(1)
    try:
        pt.TableStyle2 = style
    except Exception as e:
        raise InvalidToolArgumentsError(f"無效的樞紐樣式 '{style}'：{e}") from e
    pt.ShowTableStyleRowHeaders    = show_row_headers
    pt.ShowTableStyleColumnHeaders = show_col_headers
    pt.ShowTableStyleRowStripes    = banded_rows
    pt.ShowTableStyleColumnStripes = banded_cols
    return {
        "status":       "ok",
        "pivot_sheet":  pivot_sheet,
        "style":        style,
        "banded_rows":  banded_rows,
    }


# ── 視窗 / 欄列外觀 ───────────────────────────────────────────────────────────


def format_chart(
    chart_index: int = 1,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    has_legend: bool | None = None,
    legend_position: str | None = None,
    has_data_labels: bool | None = None,
    series_colors: list | None = None,
    plot_bg_color: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    格式化工作表上已存在的圖表（標題、座標軸、圖例、資料標籤、配色）。
    chart_index：工作表上第幾個圖表，從 1 開始。
    legend_position：bottom / right / top / left。
    series_colors：各數列顏色清單，如 ['#4472C4', '#ED7D31']。
    plot_bg_color：繪圖區背景色 #RRGGBB（留空 = 透明）。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)

    try:
        chart_obj = ws.ChartObjects(chart_index)
    except Exception as e:
        total = ws.ChartObjects().Count
        raise InvalidToolArgumentsError(
            f"找不到第 {chart_index} 個圖表（此工作表共有 {total} 個圖表）"
        ) from e

    chart = chart_obj.Chart

    if title is not None:
        chart.HasTitle = True
        chart.ChartTitle.Text = title

    # 座標軸標題（xlCategory=1, xlValue=2）
    if x_axis_title is not None:
        try:
            ax = chart.Axes(1)
            ax.HasTitle = True
            ax.AxisTitle.Text = x_axis_title
        except Exception:
            pass

    if y_axis_title is not None:
        try:
            ax = chart.Axes(2)
            ax.HasTitle = True
            ax.AxisTitle.Text = y_axis_title
        except Exception:
            pass

    if has_legend is not None:
        chart.HasLegend = has_legend

    if legend_position is not None and chart.HasLegend:
        pos_map = {
            "bottom": XL_LEGEND_BOTTOM,
            "right":  XL_LEGEND_RIGHT,
            "top":    XL_LEGEND_TOP,
            "left":   XL_LEGEND_LEFT,
        }
        chart.Legend.Position = pos_map.get(legend_position.lower(), XL_LEGEND_BOTTOM)

    if has_data_labels is not None:
        count = chart.SeriesCollection().Count
        for i in range(1, count + 1):
            try:
                chart.SeriesCollection(i).HasDataLabels = has_data_labels
            except Exception:
                pass

    if series_colors:
        for i, hex_color in enumerate(series_colors, 1):
            bgr = _hex_to_bgr(hex_color)
            try:
                s = chart.SeriesCollection(i)
                s.Interior.Color = bgr   # Column / Bar / Area
            except Exception:
                pass
            try:
                s = chart.SeriesCollection(i)
                s.Border.Color = bgr     # Line
            except Exception:
                pass

    if plot_bg_color is not None:
        try:
            chart.PlotArea.Interior.Color = _hex_to_bgr(plot_bg_color)
        except Exception:
            pass

    return {"status": "ok", "chart_index": chart_index, "sheet": ws.Name}



def create_combo_chart(
    range_addr: str,
    line_series_index: int = -1,
    secondary_axis: bool = True,
    title: str | None = None,
    sheet: str | None = None,
) -> dict:
    """
    建立組合圖（直條圖 + 折線圖），最後一個數列預設以折線顯示。
    line_series_index：哪個數列改為折線（1-based；-1 = 最後一個數列）。
    secondary_axis：折線數列是否使用次要 Y 軸（True 可避免尺度差異）。
    適合場景：銷售額（柱狀）+ 成長率（折線，次軸）。
    """
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    data_range = ws.Range(range_addr)

    left   = data_range.Left + data_range.Width + 10
    top    = data_range.Top
    width  = 420
    height = 300

    chart_obj = ws.ChartObjects().Add(Left=left, Top=top, Width=width, Height=height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=data_range)
    chart.ChartType = XL_CHART_TYPE["column"]  # 先全部設為直條

    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    else:
        chart.HasTitle = False

    series_count = chart.SeriesCollection().Count
    target_idx = line_series_index if line_series_index > 0 else series_count

    if target_idx > series_count:
        raise InvalidToolArgumentsError(
            f"line_series_index={target_idx} 超出範圍（共 {series_count} 個數列）"
        )

    s = chart.SeriesCollection(target_idx)
    s.ChartType = XL_CHART_TYPE["line"]
    if secondary_axis:
        s.AxisGroup = 2  # xlSecondary

    chart.HasLegend = True
    chart.Legend.Position = XL_LEGEND_BOTTOM

    return {
        "status":       "ok",
        "chart_type":   "combo",
        "series_count": series_count,
        "line_series":  target_idx,
        "secondary_axis": secondary_axis,
    }



def add_slicer(
    pivot_sheet: str,
    field_name: str,
    dest_sheet: str | None = None,
    left: float = 500,
    top: float = 50,
    width: float = 150,
    height: float = 200,
    sheet: str | None = None,
) -> dict:
    """
    為樞紐分析表新增切片器（Slicer）。需要 Excel 2010+。
    pivot_sheet：樞紐分析表所在的工作表名稱。
    field_name：要篩選的欄位名稱（必須與樞紐分析表欄名完全一致）。
    dest_sheet：切片器放置的工作表（省略則與樞紐同一張表）。
    left/top/width/height：切片器位置與尺寸（點數）。
    """
    excel = _get_excel()
    wb = excel.ActiveWorkbook

    pv_ws = _get_sheet(excel, pivot_sheet)

    try:
        pivot_count = pv_ws.PivotTables().Count
    except Exception:
        pivot_count = 0

    if pivot_count == 0:
        raise InvalidToolArgumentsError(f"工作表 '{pivot_sheet}' 上找不到樞紐分析表")

    pt = pv_ws.PivotTables(1)  # 使用第一個樞紐分析表

    try:
        sc = wb.SlicerCaches.Add2(pt, field_name)
    except Exception:
        try:
            sc = wb.SlicerCaches.Add(pt, field_name)
        except Exception as e:
            raise InvalidToolArgumentsError(
                f"無法建立切片器：欄位 '{field_name}' 可能不在樞紐分析表中。"
                f"詳細錯誤：{e}"
            ) from e

    target_sheet = dest_sheet or pivot_sheet
    target_ws = _get_sheet(excel, target_sheet)

    slicer = sc.Slicers.Add(
        SlicerDestination=target_ws,
        Top=top, Left=left, Width=width, Height=height,
    )

    return {
        "status":     "ok",
        "field":      field_name,
        "slicer":     slicer.Name,
        "sheet":      target_ws.Name,
    }



