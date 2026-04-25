"""
自然語言資料查詢模組（v4.7.0）

設計原則：
- query_range 讀取 Excel 範圍，在記憶體內執行過濾 / 排序 / 聚合
- 完全非破壞性：不修改任何 Excel 儲存格
- _query_data 是純函式（接受 2D list），可在 Linux CI 完整測試
- query_range 呼叫 excel_tools.read_range 取得資料後交給 _query_data 處理

過濾條件 JSON 格式
------------------
filters：條件清單（AND 邏輯）
  每個條件：{"column": int(1-based), "operator": str, "value": any}
  operator 可用：>, <, >=, <=, =, !=, contains, startswith, endswith, isblank, notblank

聚合格式
--------
aggregation：{"function": "sum"|"avg"|"count"|"max"|"min", "column": int(1-based)}

排序格式
--------
sort_by：{"column": int(1-based), "descending": bool}

top_n：int（取前 N 筆，排序後截取）
"""

from __future__ import annotations

import json
import operator as _op
from typing import Any


# ── 支援的運算子 ──────────────────────────────────────────────────────────────

_NUMERIC_OPS: dict[str, Any] = {
    ">":  _op.gt,
    "<":  _op.lt,
    ">=": _op.ge,
    "<=": _op.le,
    "=":  _op.eq,
    "!=": _op.ne,
    "==": _op.eq,
}

_STRING_OPS = {"contains", "startswith", "endswith", "isblank", "notblank"}


def _coerce(value: Any) -> Any:
    """嘗試把字串轉為數字；失敗回傳原值。"""
    if isinstance(value, str):
        try:
            return float(value) if "." in value else int(value)
        except (ValueError, TypeError):
            return value
    return value


def _apply_filter(row: list[Any], condition: dict) -> bool:
    """
    對單列套用一個過濾條件。
    column 為 1-based 欄號。不合法條件預設通過（寬鬆策略）。
    """
    col_idx = condition.get("column", 1)
    op_str  = str(condition.get("operator", "=")).strip().lower()
    cond_val = condition.get("value", "")

    if col_idx < 1 or col_idx > len(row):
        return True  # 欄號超出範圍→不過濾

    cell = row[col_idx - 1]

    # 字串操作
    if op_str == "isblank":
        return cell is None or str(cell).strip() == ""
    if op_str == "notblank":
        return cell is not None and str(cell).strip() != ""

    cell_str = str(cell) if cell is not None else ""
    if op_str == "contains":
        return str(cond_val).lower() in cell_str.lower()
    if op_str == "startswith":
        return cell_str.lower().startswith(str(cond_val).lower())
    if op_str == "endswith":
        return cell_str.lower().endswith(str(cond_val).lower())

    # 數值比較
    if op_str in _NUMERIC_OPS:
        try:
            cell_num = _coerce(cell)
            cond_num = _coerce(cond_val)
            return _NUMERIC_OPS[op_str](cell_num, cond_num)
        except (TypeError, ValueError):
            # fallback: 字串比較
            return _NUMERIC_OPS[op_str](str(cell), str(cond_val))

    return True  # 未知 operator → 不過濾


def _parse_condition_json(condition_json: str) -> dict:
    """解析 condition_json 字串，容錯處理。"""
    if not condition_json or not condition_json.strip():
        return {}
    try:
        parsed = json.loads(condition_json)
        return parsed if isinstance(parsed, dict) else {}
    except (json.JSONDecodeError, ValueError):
        return {}


def _query_data(
    data: list[list[Any]],
    condition_json: str = "",
    aggregation_json: str = "",
    has_header: bool = True,
) -> dict:
    """
    純函式：對二維陣列執行過濾 / 排序 / 聚合。

    Parameters
    ----------
    data            : read_range 回傳的二維陣列（含標題列）
    condition_json  : 過濾 / 排序 / top_n 條件（JSON 字串）
    aggregation_json: 聚合設定（JSON 字串）
    has_header      : 第一列是否為標題（預設 True）

    Returns
    -------
    dict with keys: headers, filtered_rows, total_rows, filtered_count,
                    aggregation, top_n_applied, sort_applied
    """
    if not data:
        return {
            "headers":        [],
            "filtered_rows":  [],
            "total_rows":     0,
            "filtered_count": 0,
            "aggregation":    None,
        }

    # 分離標題與資料
    if has_header and len(data) >= 1:
        headers = [str(h) if h is not None else "" for h in data[0]]
        rows    = list(data[1:])
    else:
        headers = []
        rows    = list(data)

    total_rows = len(rows)

    # ── 解析條件 ──────────────────────────────────────────────────────────────
    cond = _parse_condition_json(condition_json)
    agg  = _parse_condition_json(aggregation_json)

    filters  = cond.get("filters", [])
    sort_by  = cond.get("sort_by")
    top_n    = cond.get("top_n")

    # ── 過濾 ──────────────────────────────────────────────────────────────────
    filtered = rows
    if filters:
        filtered = [
            row for row in rows
            if all(_apply_filter(row, f) for f in filters)
        ]

    # ── 排序 ──────────────────────────────────────────────────────────────────
    sort_applied = False
    if sort_by and isinstance(sort_by, dict):
        col_idx = sort_by.get("column", 1)
        desc    = bool(sort_by.get("descending", False))
        if 1 <= col_idx <= (len(headers) or (len(filtered[0]) if filtered else 0)):
            def _sort_key(row):
                val = row[col_idx - 1] if col_idx <= len(row) else None
                # 數字排在文字前；None 排最後
                if val is None:
                    return (2, "")
                coerced = _coerce(val)
                if isinstance(coerced, (int, float)):
                    return (0, coerced)
                return (1, str(val).lower())

            filtered = sorted(filtered, key=_sort_key, reverse=desc)
            sort_applied = True

    # ── Top-N ──────────────────────────────────────────────────────────────────
    top_n_applied = False
    if isinstance(top_n, int) and top_n > 0:
        filtered = filtered[:top_n]
        top_n_applied = True

    # ── 聚合 ──────────────────────────────────────────────────────────────────
    agg_result = None
    if agg:
        func     = str(agg.get("function", "count")).lower()
        agg_col  = agg.get("column", 1)
        label    = agg.get("label", "")

        if func == "count":
            agg_result = {"function": "count", "value": len(filtered), "label": label or "筆數"}
        elif agg_col and 1 <= agg_col <= (len(headers) or 999):
            nums = []
            for row in filtered:
                if agg_col <= len(row):
                    v = _coerce(row[agg_col - 1])
                    if isinstance(v, (int, float)):
                        nums.append(v)

            if nums:
                col_name = headers[agg_col - 1] if headers and agg_col <= len(headers) else f"欄{agg_col}"
                if func == "sum":
                    agg_result = {"function": "sum",   "value": round(sum(nums), 4),   "column": col_name, "label": label}
                elif func == "avg":
                    agg_result = {"function": "avg",   "value": round(sum(nums)/len(nums), 4), "column": col_name, "label": label}
                elif func == "max":
                    agg_result = {"function": "max",   "value": max(nums),             "column": col_name, "label": label}
                elif func == "min":
                    agg_result = {"function": "min",   "value": min(nums),             "column": col_name, "label": label}
                else:
                    agg_result = {"function": func, "value": None, "error": f"不支援的聚合函數：{func}"}
            else:
                agg_result = {"function": func, "value": None, "note": "過濾後無可計算的數值"}

    # 回傳（限制回傳列數以控制 context window 大小）
    MAX_RETURN_ROWS = 200
    return_rows = filtered[:MAX_RETURN_ROWS]
    truncated   = len(filtered) > MAX_RETURN_ROWS

    return {
        "headers":        headers,
        "filtered_rows":  return_rows,
        "total_rows":     total_rows,
        "filtered_count": len(filtered),
        "returned_count": len(return_rows),
        "truncated":      truncated,
        "aggregation":    agg_result,
        "sort_applied":   sort_applied,
        "top_n_applied":  top_n_applied,
    }


def query_range(
    range_addr: str,
    condition_json: str = "",
    aggregation_json: str = "",
    has_header: bool = True,
    sheet: str | None = None,
) -> dict:
    """
    從 Excel 讀取範圍後，在記憶體內執行查詢（非破壞性）。

    Parameters
    ----------
    range_addr      : Excel 範圍位址（如 "A1:F100"）
    condition_json  : JSON 字串，描述 filters / sort_by / top_n
    aggregation_json: JSON 字串，描述聚合函數與目標欄
    has_header      : 第一列是否為標題（預設 True）
    sheet           : 工作表名稱；省略時用作用中工作表

    Returns
    -------
    dict（含 headers / filtered_rows / aggregation 等欄位）
    """
    import excel_tools as et
    data = et.read_range(range_addr, sheet)
    if not isinstance(data, list):
        return {"error": f"read_range 回傳非預期格式：{type(data).__name__}"}

    result = _query_data(data, condition_json, aggregation_json, has_header)
    result["range_addr"] = range_addr
    result["sheet"]      = sheet
    return result
