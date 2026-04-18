"""
所有 win32com Excel 操作。
唯一與 Excel 程序互動的模組，不引入任何 LLM 相關套件。
"""

import pythoncom
import win32com.client
from constants import (
    XL_ALIGN, XL_CHART_TYPE, XL_BORDER_STYLE, XL_BORDER_SIDES,
    XL_ROW_FIELD, XL_COL_FIELD, XL_DATA_FIELD, XL_SUM, XL_DATABASE,
    XL_ASCENDING, XL_DESCENDING, XL_YES, XL_NO,
)


# ── 內部輔助函數 ──────────────────────────────────────────────────────────────

def _get_excel():
    """
    取得已開啟的 Excel Application 實例。
    ★ B7：Streamlit 在子線程處理請求，COM 需要在每個線程中初始化，
    否則 GetActiveObject 會失敗。加入 CoInitialize + Dispatch 備援。
    """
    pythoncom.CoInitialize()  # 確保當前線程已初始化 COM

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

    raise RuntimeError("找不到開啟中的 Excel，請先開啟 Excel 再執行此指令")


def _get_sheet(excel, sheet_name: str | None = None):
    """取得工作表物件；★ B5：加入 ActiveWorkbook None guard"""
    wb = excel.ActiveWorkbook
    if wb is None:
        raise RuntimeError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    if sheet_name:
        return wb.Sheets(sheet_name)
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


# ── 讀取 ──────────────────────────────────────────────────────────────────────

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
        raise RuntimeError("Excel 已開啟但沒有活頁簿")
    return {
        "file_name":    wb.Name,
        "sheets":       [ws.Name for ws in wb.Sheets],
        "active_sheet": excel.ActiveSheet.Name,
        "selection":    excel.Selection.Address,
    }


def get_used_range(sheet: str | None = None) -> str:
    """回傳工作表中有資料的範圍位址"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    return ws.UsedRange.Address


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
    return {"status": "ok", "written_to": written_range.Address}


def save_workbook() -> dict:
    """儲存目前作用中的活頁簿（★ 新增 save tool）"""
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise RuntimeError("沒有開啟中的活頁簿")
    wb.Save()
    return {"status": "ok", "saved": wb.Name}


# ── 格式 ──────────────────────────────────────────────────────────────────────

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


# ── 列操作 ─────────────────────────────────────────────────────────────────────

def insert_row(index: int, count: int = 1, sheet: str | None = None) -> dict:
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(f"{index}:{index + count - 1}").Insert()
    return {"status": "ok", "inserted_at": index, "count": count}


def delete_row(index: int, count: int = 1, sheet: str | None = None) -> dict:
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(f"{index}:{index + count - 1}").Delete()
    return {"status": "ok", "deleted_at": index, "count": count}


# ── 欄操作 ─────────────────────────────────────────────────────────────────────

def insert_column(index: int, count: int = 1, sheet: str | None = None) -> dict:
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Columns(f"{index}:{index + count - 1}").Insert()
    return {"status": "ok", "inserted_at": index, "count": count}


def delete_column(index: int, count: int = 1, sheet: str | None = None) -> dict:
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Columns(f"{index}:{index + count - 1}").Delete()
    return {"status": "ok", "deleted_at": index, "count": count}


# ── 工作表操作 ─────────────────────────────────────────────────────────────────

def add_sheet(name: str) -> dict:
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise RuntimeError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    wb.Sheets.Add().Name = name
    return {"status": "ok", "added": name}


def rename_sheet(old_name: str, new_name: str) -> dict:
    excel = _get_excel()
    wb = excel.ActiveWorkbook
    if wb is None:
        raise RuntimeError("Excel 已開啟但沒有活頁簿，請先開啟或新增一個 Excel 檔案")
    wb.Sheets(old_name).Name = new_name
    return {"status": "ok", "renamed": f"{old_name} → {new_name}"}


# ── 資料操作 ──────────────────────────────────────────────────────────────────

def sort_range(
    range_addr: str,
    column_index: int,
    ascending: bool = True,
    has_header: bool = True,
    sheet: str | None = None,
) -> dict:
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


# ── 圖表 ──────────────────────────────────────────────────────────────────────



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
        TableName="PivotTable1",
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


# ── 視窗 / 欄列外觀 ───────────────────────────────────────────────────────────

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
    ws.Activate()

    # 先解除現有凍結
    excel.ActiveWindow.FreezePanes = False

    if row > 0 or col > 0:
        # 選取凍結分割點右下角的儲存格
        ws.Cells(row + 1, col + 1).Select()
        excel.ActiveWindow.FreezePanes = True

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
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Columns(column_index).ColumnWidth = width
    return {"status": "ok", "column": column_index, "width": width}


# ── V3 新增工具 ───────────────────────────────────────────────────────────────

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
    else:
        rng.Clear()

    return {"status": "ok", "cleared": target, "range": range_addr or "UsedRange"}


def set_row_height(
    row_index: int,
    height: float,
    sheet: str | None = None,
) -> dict:
    """設定列高（Excel 點數單位，預設約 15）"""
    excel = _get_excel()
    ws = _get_sheet(excel, sheet)
    ws.Rows(row_index).RowHeight = height
    return {"status": "ok", "row": row_index, "height": height}


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
