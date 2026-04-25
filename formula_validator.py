"""
公式智慧輔助（v4.7.0）

功能：
- validate_formula : 括號配對、函數名稱白名單、儲存格引用格式驗證
- explain_formula  : 生成公式的繁體中文說明（函數語意 + 引用範圍）

本模組純 Python，無 COM / Streamlit 依賴，可在 Linux CI 測試。
"""

from __future__ import annotations

import re
from typing import NamedTuple


# ── 已知 Excel 函數白名單 ─────────────────────────────────────────────────────

KNOWN_FUNCTIONS: frozenset[str] = frozenset({
    # 數學 / 三角
    "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN",
    "INT", "MOD", "SQRT", "POWER", "LOG", "LOG10", "LN", "EXP", "RAND", "RANDBETWEEN",
    "CEILING", "FLOOR", "TRUNC", "SIGN", "PI", "FACT", "COMBIN", "PERMUT",
    # 統計
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "COUNTBLANK", "MAX", "MIN", "MEDIAN", "MODE", "STDEV", "STDEVP", "VAR", "VARP",
    "LARGE", "SMALL", "RANK", "PERCENTILE", "QUARTILE", "CORREL", "COVAR",
    # 邏輯
    "IF", "IFS", "AND", "OR", "NOT", "XOR", "TRUE", "FALSE", "IFERROR", "IFNA",
    "SWITCH",
    # 文字
    "CONCATENATE", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "TRIM",
    "UPPER", "LOWER", "PROPER", "REPLACE", "SUBSTITUTE", "FIND", "SEARCH",
    "TEXT", "VALUE", "CHAR", "CODE", "REPT", "EXACT", "CLEAN", "NUMBERVALUE",
    # 查詢 / 引用
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "XMATCH", "OFFSET",
    "INDIRECT", "ADDRESS", "CHOOSE", "LOOKUP", "FORMULATEXT", "HYPERLINK",
    # 日期 / 時間
    "DATE", "DATEVALUE", "DAY", "MONTH", "YEAR", "NOW", "TODAY", "TIME",
    "HOUR", "MINUTE", "SECOND", "DATEDIF", "NETWORKDAYS", "WORKDAY",
    "EOMONTH", "EDATE", "WEEKDAY", "WEEKNUM", "DAYS", "DAYS360",
    # 財務
    "PMT", "PV", "FV", "RATE", "NPV", "IRR", "NPER", "IPMT", "PPMT", "XIRR", "XNPV",
    # 動態陣列（Excel 365）
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    # 資訊
    "ROW", "COLUMN", "ROWS", "COLUMNS", "TRANSPOSE",
    "ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISNA", "ISODD", "ISEVEN",
    "CELL", "TYPE", "N", "NA", "INFO", "SHEET", "SHEETS",
})

# 函數中文說明（用於 explain_formula）
_FUNC_DESC: dict[str, str] = {
    "SUM":         "加總數值",
    "AVERAGE":     "計算平均值",
    "COUNT":       "計算數字個數",
    "COUNTA":      "計算非空白格數",
    "COUNTBLANK":  "計算空白格數",
    "COUNTIF":     "依單一條件計算個數",
    "COUNTIFS":    "依多個條件計算個數",
    "SUMIF":       "依單一條件加總",
    "SUMIFS":      "依多個條件加總",
    "AVERAGEIF":   "依條件計算平均",
    "AVERAGEIFS":  "依多個條件計算平均",
    "SUMPRODUCT":  "對應元素相乘後加總（常用於多條件加總）",
    "MAX":         "取最大值",
    "MIN":         "取最小值",
    "LARGE":       "取第 N 大的值",
    "SMALL":       "取第 N 小的值",
    "MEDIAN":      "取中位數",
    "RANK":        "取數值在範圍中的排名",
    "IF":          "條件判斷（若…則…否則…）",
    "IFS":         "多重條件判斷",
    "IFERROR":     "發生錯誤時顯示備用值",
    "IFNA":        "#N/A 錯誤時顯示備用值",
    "AND":         "所有條件均為真時回傳 TRUE",
    "OR":          "任一條件為真時回傳 TRUE",
    "NOT":         "反轉條件（TRUE→FALSE）",
    "SWITCH":      "多值比對（類似 switch/case）",
    "VLOOKUP":     "垂直查詢：在第一欄搜尋，回傳同列指定欄的值",
    "HLOOKUP":     "水平查詢：在第一列搜尋，回傳同欄指定列的值",
    "INDEX":       "取得指定列欄交叉位置的儲存格值",
    "MATCH":       "查找目標在範圍中的位置（回傳數字）",
    "XLOOKUP":     "彈性查詢（可向左/向右搜尋，Excel 365+）",
    "OFFSET":      "從基準格偏移指定列欄後取得參照",
    "INDIRECT":    "將文字字串轉為儲存格參照",
    "CHOOSE":      "從清單中依索引取值",
    "LEFT":        "取文字左側 N 個字元",
    "RIGHT":       "取文字右側 N 個字元",
    "MID":         "取文字中間片段",
    "LEN":         "計算文字長度（字元數）",
    "TRIM":        "移除頭尾及多餘空白",
    "UPPER":       "轉為全大寫",
    "LOWER":       "轉為全小寫",
    "PROPER":      "每個單字首字母大寫",
    "CONCATENATE": "合併多個文字（舊版）",
    "CONCAT":      "合併多個文字或範圍（現代版）",
    "TEXTJOIN":    "用分隔符合併文字範圍",
    "TEXT":        "將數值格式化為指定樣式的文字",
    "VALUE":       "將文字轉為數值",
    "SUBSTITUTE":  "取代文字中的特定字串",
    "REPLACE":     "取代文字中指定位置的字元",
    "FIND":        "在文字中搜尋子字串（區分大小寫，回傳位置）",
    "SEARCH":      "在文字中搜尋子字串（不分大小寫，回傳位置）",
    "DATE":        "由年、月、日組成日期值",
    "TODAY":       "回傳今天日期",
    "NOW":         "回傳目前日期與時間",
    "DATEDIF":     "計算兩日期之差（年/月/日）",
    "NETWORKDAYS": "計算兩日期間的工作天數",
    "WORKDAY":     "從指定日期加/減工作天後的日期",
    "EOMONTH":     "指定日期所在月份（偏移 N 月）的最後一天",
    "EDATE":       "從指定日期偏移 N 個月後的日期",
    "YEAR":        "取日期的年份",
    "MONTH":       "取日期的月份",
    "DAY":         "取日期的日",
    "WEEKDAY":     "回傳星期幾（1-7）",
    "ROUND":       "四捨五入到指定小數位",
    "ROUNDUP":     "無條件進位",
    "ROUNDDOWN":   "無條件捨去",
    "ABS":         "取絕對值",
    "MOD":         "取餘數",
    "INT":         "取整數部分（向下取整）",
    "SQRT":        "取平方根",
    "POWER":       "計算次方（base^exp）",
    "PMT":         "計算貸款每期還款金額",
    "PV":          "計算現值",
    "FV":          "計算終值",
    "RATE":        "計算利率",
    "NPV":         "計算淨現值",
    "IRR":         "計算內部報酬率",
    "ROW":         "回傳儲存格的列號",
    "COLUMN":      "回傳儲存格的欄號",
    "ROWS":        "回傳範圍的列數",
    "COLUMNS":     "回傳範圍的欄數",
    "ISBLANK":     "判斷儲存格是否為空白",
    "ISNUMBER":    "判斷值是否為數字",
    "ISTEXT":      "判斷值是否為文字",
    "ISERROR":     "判斷值是否為任何錯誤",
    "ISNA":        "判斷值是否為 #N/A",
    "FILTER":      "依條件篩選範圍並回傳結果（Excel 365+）",
    "SORT":        "對範圍排序（Excel 365+）",
    "UNIQUE":      "取出唯一值清單（Excel 365+）",
    "SEQUENCE":    "產生等差數列（Excel 365+）",
    "TRANSPOSE":   "將橫向範圍轉為直向（或反之）",
    "STDEV":       "計算樣本標準差",
    "VAR":         "計算樣本變異數",
    "PERCENTILE":  "取第 N 百分位數",
    "QUARTILE":    "取四分位數",
    "CORREL":      "計算兩組資料的相關係數",
}

# ── 正規表示式 ────────────────────────────────────────────────────────────────

# 匹配函數調用：FUNCNAME(
_FUNC_RE = re.compile(r"\b([A-Z][A-Z0-9_\.]*)\s*\(", re.IGNORECASE)

# 匹配儲存格 / 範圍引用（含跨工作表 Sheet!A1）
_CELL_REF_RE = re.compile(
    r"""
    (?:'[^']*'!|[A-Za-z_][\w]*!)?  # 可選的工作表前綴（'Sheet Name'! 或 Sheet!）
    \$?[A-Z]{1,3}\$?\d{1,7}         # 欄字母 + 列號，可含 $
    (?::\$?[A-Z]{1,3}\$?\d{1,7})?  # 可選的範圍終點
    """,
    re.VERBOSE | re.IGNORECASE,
)


# ── 資料結構 ──────────────────────────────────────────────────────────────────

class ValidationResult(NamedTuple):
    valid:          bool
    errors:         list[str]
    warnings:       list[str]
    functions_used: list[str]


# ── 公開 API ──────────────────────────────────────────────────────────────────

def validate_formula(formula: str) -> ValidationResult:
    """
    驗證 Excel 公式的語法正確性。

    檢查項目
    --------
    1. 必須以 = 開頭（含 =+ / =- 等合法變體）
    2. 括號配對（不含字串內括號）
    3. 函數名稱白名單（非白名單項目給 warning，不視為錯誤）

    Returns
    -------
    ValidationResult(valid, errors, warnings, functions_used)
    """
    errors:   list[str] = []
    warnings: list[str] = []

    formula = (formula or "").strip()

    if not formula.startswith("="):
        errors.append("公式必須以 = 開頭")
        return ValidationResult(False, errors, warnings, [])

    # ── 括號配對（跳過字串內容）──────────────────────────────────────────────
    depth      = 0
    in_string  = False
    for i, ch in enumerate(formula):
        if ch == '"':
            # 簡易字串偵測（不處理跳脫，Excel 公式用 "" 表示引號）
            in_string = not in_string
        if not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    errors.append(f"第 {i + 1} 個字元位置：多餘的右括號 ')'")
                    depth = 0  # 繼續掃描其餘括號問題

    if depth > 0:
        errors.append(f"括號未完整閉合：尚缺 {depth} 個右括號 ')'")

    # ── 函數名稱檢查 ──────────────────────────────────────────────────────────
    funcs_seen: list[str] = []
    for m in _FUNC_RE.finditer(formula):
        fname = m.group(1).upper()
        if fname not in funcs_seen:
            funcs_seen.append(fname)
        if fname not in KNOWN_FUNCTIONS:
            warnings.append(f"未知函數 {fname}（可能是自訂名稱、拼寫錯誤，或需 Excel 365 才支援）")

    return ValidationResult(
        valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        functions_used=funcs_seen,
    )


def explain_formula(formula: str) -> str:
    """
    生成 Excel 公式的繁體中文說明。

    Parameters
    ----------
    formula : Excel 公式字串（含或不含開頭的 =）

    Returns
    -------
    人類可讀的說明文字（多行字串）
    """
    formula = (formula or "").strip()
    if not formula:
        return "（空公式）"

    if not formula.startswith("="):
        return f"「{formula}」不是公式（需以 = 開頭）"

    body = formula[1:]

    # 找出使用的函數（去重，保留順序）
    funcs_seen: list[str] = []
    for m in _FUNC_RE.finditer(body):
        fname = m.group(1).upper()
        if fname not in funcs_seen:
            funcs_seen.append(fname)

    # 找出引用的儲存格 / 範圍（去重）
    refs_seen: list[str] = []
    for m in _CELL_REF_RE.finditer(body):
        ref = m.group(0)
        if ref not in refs_seen:
            refs_seen.append(ref)

    parts: list[str] = [f"📐 公式：`{formula}`"]

    if funcs_seen:
        parts.append(f"🔧 使用函數：{', '.join(funcs_seen)}")

    if refs_seen:
        parts.append(f"📍 引用範圍：{', '.join(refs_seen)}")

    # 函數逐一說明
    desc_lines = [
        f"  · {f}：{_FUNC_DESC[f]}"
        for f in funcs_seen
        if f in _FUNC_DESC
    ]
    unknown_funcs = [f for f in funcs_seen if f not in _FUNC_DESC]
    if unknown_funcs:
        desc_lines.append(
            f"  · {', '.join(unknown_funcs)}：自訂或較罕見的函數，請查閱 Excel 說明"
        )

    if desc_lines:
        parts.append("📖 函數說明：\n" + "\n".join(desc_lines))

    return "\n".join(parts)


def validate_formula_tool(formula: str) -> dict:
    """
    explain_formula 的工具包裝版本：回傳 dict，供 executor 呼叫。
    """
    result = validate_formula(formula)
    return {
        "valid":          result.valid,
        "errors":         result.errors,
        "warnings":       result.warnings,
        "functions_used": result.functions_used,
        "status":         "ok" if result.valid else "invalid",
        "summary": (
            "公式語法正確" + (f"，但有 {len(result.warnings)} 個警告" if result.warnings else "")
            if result.valid
            else f"公式有 {len(result.errors)} 個錯誤"
        ),
    }


def explain_formula_tool(formula: str) -> dict:
    """
    explain_formula 的工具包裝版本：回傳 dict，供 executor 呼叫。
    """
    text = explain_formula(formula)
    result = validate_formula(formula)
    return {
        "formula":        formula,
        "explanation":    text,
        "functions_used": result.functions_used,
        "valid":          result.valid,
        "warnings":       result.warnings,
    }
