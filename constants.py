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
    # 基本類型
    "column":         51,    # xlColumnClustered
    "column_stacked": 52,    # xlColumnStacked
    "bar":            57,    # xlBarClustered
    "bar_stacked":    58,    # xlBarStacked
    "line":           4,     # xlLine
    "line_markers":   65,    # xlLineMarkers
    "pie":            5,     # xlPie
    "doughnut":       -4120, # xlDoughnut
    "area":           1,     # xlArea
    "area_stacked":   76,    # xlAreaStacked
    "scatter":        74,    # xlXYScatter
    # Excel 2016+ 統計圖表
    "histogram":      118,   # xlHistogram（需 Excel 2016+）
    "waterfall":      119,   # xlWaterfall（需 Excel 2016+）
    "box_whisker":    121,   # xlBoxWhisker（需 Excel 2016+）
    "funnel":         123,   # xlFunnel（需 Excel 2016+）
}

# ── Sparkline 類型 ────────────────────────────────────────────────────────────
XL_SPARKLINE_LINE     = 1   # xlSparkLine
XL_SPARKLINE_COLUMN   = 2   # xlSparkColumn
XL_SPARKLINE_WINLOSS  = 3   # xlSparkColumnStacked100

# ── 圖例位置（xlLegendPosition）──────────────────────────────────────────────
XL_LEGEND_BOTTOM  = -4107   # xlLegendPositionBottom
XL_LEGEND_RIGHT   = -4152   # xlLegendPositionRight
XL_LEGEND_TOP     = -4160   # xlLegendPositionTop
XL_LEGEND_LEFT    = -4131   # xlLegendPositionLeft
XL_LEGEND_CORNER  = 2       # xlLegendPositionCorner

# ── 頁面方向（xlPageOrientation）─────────────────────────────────────────────
XL_PORTRAIT   = 1   # xlPortrait
XL_LANDSCAPE  = 2   # xlLandscape

# ── 紙張大小（xlPaperSize）───────────────────────────────────────────────────
XL_PAPER_SIZE = {
    "a4":     9,    # xlPaperA4
    "letter": 1,    # xlPaperLetter
    "a3":     8,    # xlPaperA3
    "a5":     11,   # xlPaperA5
    "legal":  5,    # xlPaperLegal
}

# ── Excel Table 樣式對照表 ────────────────────────────────────────────────────
TABLE_STYLE_MAP = {
    "blue":         "TableStyleMedium9",
    "light_blue":   "TableStyleLight9",
    "orange":       "TableStyleMedium3",
    "light_orange": "TableStyleLight3",
    "green":        "TableStyleMedium14",
    "light_green":  "TableStyleLight14",
    "red":          "TableStyleMedium4",
    "purple":       "TableStyleMedium11",
    "gray":         "TableStyleMedium15",
    "white":        "TableStyleLight1",
    "dark_blue":    "TableStyleDark1",
    "dark_green":   "TableStyleDark4",
    "dark_red":     "TableStyleDark3",
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

# ── DataSeries 類型（xlDataSeriesType）───────────────────────────────────────
XL_SERIES_LINEAR  = -4132   # xlLinear（數字線性遞增）
XL_SERIES_DATE    = 4       # xlDate（日期序列）
XL_SERIES_COLUMNS = 2       # xlColumns（向下填充）
XL_SERIES_ROWS    = 1       # xlRows（向右填充）

# DataSeries 日期單位（xlTimeUnit）
XL_DATE_DAY      = 1   # xlDay
XL_DATE_WEEKDAY  = 2   # xlWeekday
XL_DATE_MONTH    = 3   # xlMonth
XL_DATE_YEAR     = 4   # xlYear

# ── AdvancedFilter 動作 ───────────────────────────────────────────────────────
XL_FILTER_IN_PLACE = 1   # xlFilterInPlace
XL_FILTER_COPY     = 2   # xlFilterCopy

# ── Subtotal 函數類型（xlConsolidationFunction）──────────────────────────────
XL_COUNT   = -4112   # xlCount
XL_AVERAGE = -4106   # xlAverage
XL_MAX     = -4135   # xlMax
XL_MIN_FUNC = -4136  # xlMin（避免與內建 min 衝突）

# ── TextToColumns DataType ───────────────────────────────────────────────────
XL_DELIMITED   = 1   # xlDelimited
XL_FIXED_WIDTH = 2   # xlFixedWidth
