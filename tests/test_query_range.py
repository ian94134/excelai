"""
excel_query 模組單元測試（v4.7.0）

測試 _query_data（純函式，不需 COM）：
- 過濾：>、<、>=、<=、=、!=、contains、startswith、endswith、isblank、notblank
- 排序：升序、降序
- top_n
- 聚合：sum、avg、count、max、min
- 組合：過濾 + 聚合、過濾 + 排序 + top_n
- 邊界：空資料、無標題列、欄號超出範圍、未知聚合函數
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_query import _query_data, _apply_filter, _coerce


# ── 測試資料 ──────────────────────────────────────────────────────────────────

SAMPLE_DATA = [
    ["姓名",  "部門",  "業績",   "月份"],
    ["Alice", "業務",  15000,  1],
    ["Bob",   "技術",  8000,   1],
    ["Carol", "業務",  22000,  2],
    ["Dave",  "技術",  5000,   2],
    ["Eve",   "業務",  18000,  3],
    ["Frank", "管理",  None,   3],   # 空值
]


def cj(filters=None, sort_by=None, top_n=None) -> str:
    """Helper：建立 condition_json 字串。"""
    d = {}
    if filters  is not None: d["filters"]  = filters
    if sort_by  is not None: d["sort_by"]  = sort_by
    if top_n    is not None: d["top_n"]    = top_n
    return json.dumps(d)


def aj(function, column=None, label="") -> str:
    """Helper：建立 aggregation_json 字串。"""
    d = {"function": function}
    if column is not None: d["column"] = column
    if label: d["label"] = label
    return json.dumps(d)


# ── 基本結構 ──────────────────────────────────────────────────────────────────

def test_returns_headers_and_rows():
    r = _query_data(SAMPLE_DATA)
    assert r["headers"] == ["姓名", "部門", "業績", "月份"]
    assert r["total_rows"] == 6
    assert r["filtered_count"] == 6


def test_empty_data():
    r = _query_data([])
    assert r["filtered_rows"] == []
    assert r["total_rows"] == 0


def test_no_header():
    data = [["A", 1], ["B", 2], ["C", 3]]
    r = _query_data(data, has_header=False)
    assert r["headers"] == []
    assert r["total_rows"] == 3


# ── 過濾 ─────────────────────────────────────────────────────────────────────

def test_filter_greater_than():
    r = _query_data(SAMPLE_DATA, cj([{"column": 3, "operator": ">", "value": 10000}]))
    names = [row[0] for row in r["filtered_rows"]]
    assert "Alice" in names and "Carol" in names and "Eve" in names
    assert "Bob" not in names and "Dave" not in names


def test_filter_less_than_or_equal():
    r = _query_data(SAMPLE_DATA, cj([{"column": 3, "operator": "<=", "value": 8000}]))
    names = [row[0] for row in r["filtered_rows"]]
    assert "Bob" in names and "Dave" in names
    assert "Alice" not in names


def test_filter_equals():
    r = _query_data(SAMPLE_DATA, cj([{"column": 2, "operator": "=", "value": "業務"}]))
    assert r["filtered_count"] == 3


def test_filter_accepts_header_name_and_op_alias():
    r = _query_data(SAMPLE_DATA, cj([{"column": "部門", "op": "equals", "value": "業務"}]))
    assert r["filtered_count"] == 3


def test_unknown_string_column_is_ignored():
    r = _query_data(SAMPLE_DATA, cj([{"column": "不存在欄位", "op": "equals", "value": "業務"}]))
    assert r["filtered_count"] == 6


def test_filter_not_equals():
    r = _query_data(SAMPLE_DATA, cj([{"column": 2, "operator": "!=", "value": "技術"}]))
    departments = [row[1] for row in r["filtered_rows"]]
    assert "技術" not in departments


def test_filter_contains():
    r = _query_data(SAMPLE_DATA, cj([{"column": 1, "operator": "contains", "value": "a"}]))
    # Alice, Carol, Dave (case-insensitive)
    names = [row[0] for row in r["filtered_rows"]]
    assert "Alice" in names or "Carol" in names or "Dave" in names


def test_filter_startswith():
    r = _query_data(SAMPLE_DATA, cj([{"column": 1, "operator": "startswith", "value": "A"}]))
    names = [row[0] for row in r["filtered_rows"]]
    assert "Alice" in names


def test_filter_endswith():
    r = _query_data(SAMPLE_DATA, cj([{"column": 1, "operator": "endswith", "value": "e"}]))
    names = [row[0] for row in r["filtered_rows"]]
    assert "Alice" in names or "Dave" in names or "Eve" in names


def test_filter_isblank():
    r = _query_data(SAMPLE_DATA, cj([{"column": 3, "operator": "isblank"}]))
    assert r["filtered_count"] == 1
    assert r["filtered_rows"][0][0] == "Frank"


def test_filter_notblank():
    r = _query_data(SAMPLE_DATA, cj([{"column": 3, "operator": "notblank"}]))
    assert r["filtered_count"] == 5


def test_multiple_filters_and_logic():
    r = _query_data(SAMPLE_DATA, cj([
        {"column": 2, "operator": "=",  "value": "業務"},
        {"column": 3, "operator": ">",  "value": 16000},
    ]))
    names = [row[0] for row in r["filtered_rows"]]
    assert "Carol" in names and "Eve" in names
    assert "Alice" not in names  # 15000 < 16000


# ── 排序 ─────────────────────────────────────────────────────────────────────

def test_sort_ascending():
    r = _query_data(SAMPLE_DATA, cj(sort_by={"column": 3, "descending": False}))
    vals = [row[2] for row in r["filtered_rows"] if row[2] is not None]
    assert vals == sorted(vals)


def test_sort_descending():
    r = _query_data(SAMPLE_DATA, cj(sort_by={"column": 3, "descending": True}))
    vals = [row[2] for row in r["filtered_rows"] if row[2] is not None]
    # None 會排到後面，只驗證有值的部分有序
    non_none = [v for v in [row[2] for row in r["filtered_rows"]] if v is not None]
    assert non_none == sorted(non_none, reverse=True)


# ── top_n ─────────────────────────────────────────────────────────────────────

def test_top_n():
    r = _query_data(SAMPLE_DATA, cj(
        sort_by={"column": 3, "descending": True},
        top_n=3,
    ))
    assert r["filtered_count"] == 3
    assert r["top_n_applied"] is True


# ── 聚合 ─────────────────────────────────────────────────────────────────────

def test_aggregation_sum():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("sum", 3))
    expected = 15000 + 8000 + 22000 + 5000 + 18000  # Frank is None, skipped
    assert r["aggregation"]["value"] == expected


def test_aggregation_avg():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("avg", 3))
    expected = (15000 + 8000 + 22000 + 5000 + 18000) / 5
    assert abs(r["aggregation"]["value"] - expected) < 0.01


def test_aggregation_count():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("count"))
    assert r["aggregation"]["value"] == 6  # count 計所有過濾後列數


def test_aggregation_max():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("max", 3))
    assert r["aggregation"]["value"] == 22000


def test_aggregation_min():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("min", 3))
    assert r["aggregation"]["value"] == 5000


def test_aggregation_after_filter():
    r = _query_data(
        SAMPLE_DATA,
        cj([{"column": 2, "operator": "=", "value": "業務"}]),
        aj("sum", 3),
    )
    expected = 15000 + 22000 + 18000
    assert r["aggregation"]["value"] == expected


def test_aggregation_no_numeric_values():
    r = _query_data(SAMPLE_DATA, aggregation_json=aj("sum", 1))  # 欄1=姓名（文字）
    assert r["aggregation"]["value"] is None or r["aggregation"].get("note")


# ── 邊界情況 ──────────────────────────────────────────────────────────────────

def test_column_out_of_range_filter_passes():
    """欄號超出範圍的過濾條件應寬鬆通過（不報錯）。"""
    r = _query_data(SAMPLE_DATA, cj([{"column": 99, "operator": ">", "value": 0}]))
    assert r["filtered_count"] == 6  # 全部通過


def test_invalid_condition_json_treated_as_no_filter():
    r = _query_data(SAMPLE_DATA, condition_json="not json {{{")
    assert r["filtered_count"] == 6


def test_empty_condition_json():
    r = _query_data(SAMPLE_DATA, condition_json="")
    assert r["filtered_count"] == 6


def test_coerce_string_to_int():
    assert _coerce("42") == 42


def test_coerce_string_to_float():
    assert abs(_coerce("3.14") - 3.14) < 0.001


def test_coerce_non_numeric_unchanged():
    assert _coerce("hello") == "hello"
