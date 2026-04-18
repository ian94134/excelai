"""
Excel win32com COM 常數集中管理。
所有裸數字都在此定義，其他模組從這裡 import，不直接使用魔術數字。
"""

# ── 水平對齊（xlHAlign）────────────────────────────────────────────────────────
XL_ALIGN = {
    "left":   -4131,   # xlHAlignLeft
    "center": -4108,   # xlHAlignCenter
    "right":  -4152,   # xlHAlignRight
}

# ── 圖表類型（xlChartType）─────────────────────────────────────────────────────
XL_CHART_TYPE = {
    "column":  51,   # xlColumnClustered
    "bar":     57,   # xlBarClustered
    "line":    4,    # xlLine
    "pie":     5,    # xlPie
    "area":    1,    # xlArea
    "scatter": 74,   # xlXYScatter
}

# ── 框線樣式（xlLineStyle）─────────────────────────────────────────────────────
XL_BORDER_STYLE = {
    "thin":   1,
    "medium": -4138,
    "thick":  4,
    "dashed": -4115,
}

# ── 框線位置（xlBordersIndex）─────────────────────────────────────────────────
XL_BORDER_SIDES = {
    "all":   [7, 8, 9, 10, 11, 12],   # left / top / bottom / right / insideV / insideH
    "outer": [7, 8, 9, 10],
    "inner": [11, 12],
    "left":  [7],
    "top":   [8],
    "bottom":[9],
    "right": [10],
}

# ── PivotTable 欄位方向（xlPivotFieldOrientation）─────────────────────────────
XL_ROW_FIELD  = 1   # xlRowField
XL_COL_FIELD  = 2   # xlColumnField
XL_DATA_FIELD = 4   # xlDataField

# ── 函數類型（xlConsolidationFunction）───────────────────────────────────────
XL_SUM        = -4157   # xlSum

# ── PivotCache SourceType ─────────────────────────────────────────────────────
XL_DATABASE   = 1   # xlDatabase

# ── 排序方向（xlSortOrder）───────────────────────────────────────────────────
XL_ASCENDING  = 1   # xlAscending
XL_DESCENDING = 2   # xlDescending

# ── 標題列識別（xlYesNoGuess）────────────────────────────────────────────────
XL_YES        = 1   # xlYes（有標題列）
XL_NO         = 2   # xlNo（無標題列）
