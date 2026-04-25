"""
formula_validator 模組單元測試（v4.7.0）

測試範圍：
- validate_formula：括號配對、非 = 開頭拒絕、已知/未知函數、複合公式
- explain_formula：回傳說明文字包含函數名稱與引用範圍
- validate_formula_tool / explain_formula_tool：工具包裝回傳 dict
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from formula_validator import (
    validate_formula, explain_formula,
    validate_formula_tool, explain_formula_tool,
    KNOWN_FUNCTIONS,
)


# ── validate_formula ──────────────────────────────────────────────────────────

def test_valid_simple_sum():
    result = validate_formula("=SUM(A1:A10)")
    assert result.valid is True
    assert result.errors == []
    assert "SUM" in result.functions_used


def test_valid_nested():
    result = validate_formula("=IF(SUM(A1:A5)>100,\"高\",\"低\")")
    assert result.valid is True
    assert "IF" in result.functions_used
    assert "SUM" in result.functions_used


def test_missing_equals_is_invalid():
    result = validate_formula("SUM(A1:A10)")
    assert result.valid is False
    assert any("=" in e for e in result.errors)


def test_unmatched_open_paren():
    result = validate_formula("=SUM(A1:A10")
    assert result.valid is False
    assert any("閉合" in e or "括號" in e for e in result.errors)


def test_unmatched_close_paren():
    result = validate_formula("=SUM(A1:A10))")
    assert result.valid is False
    assert any("右括號" in e for e in result.errors)


def test_balanced_nested_parens():
    result = validate_formula("=IF(AND(A1>0,B1>0),SUM(A1:B1),0)")
    assert result.valid is True


def test_known_function_no_warning():
    result = validate_formula("=VLOOKUP(A1,B1:C10,2,FALSE)")
    assert result.valid is True
    assert result.warnings == []


def test_unknown_function_gives_warning():
    result = validate_formula("=MYFUNC(A1)")
    assert result.valid is True  # 未知函數不是錯誤
    assert any("MYFUNC" in w for w in result.warnings)


def test_functions_used_deduplication():
    result = validate_formula("=SUM(A1)+SUM(B1)+SUM(C1)")
    assert result.functions_used.count("SUM") == 1


def test_empty_formula_invalid():
    result = validate_formula("")
    assert result.valid is False


def test_equals_only_valid():
    """=1+1 是合法公式（無函數調用）"""
    result = validate_formula("=1+1")
    assert result.valid is True
    assert result.functions_used == []


def test_string_with_parens_not_counted():
    """字串內的括號不應影響括號計數"""
    result = validate_formula('=IF(A1="(ok)","yes","no")')
    assert result.valid is True


def test_known_functions_set_includes_common():
    for func in ("SUM", "AVERAGE", "VLOOKUP", "IF", "IFERROR", "COUNTIFS", "XLOOKUP"):
        assert func in KNOWN_FUNCTIONS


def test_validate_formula_tool_returns_dict():
    result = validate_formula_tool("=SUM(A1:A10)")
    assert isinstance(result, dict)
    assert result["valid"] is True
    assert "status" in result
    assert "summary" in result
    assert "functions_used" in result


def test_validate_formula_tool_invalid():
    result = validate_formula_tool("=SUM(A1:A10")
    assert result["valid"] is False
    assert result["status"] == "invalid"


# ── explain_formula ───────────────────────────────────────────────────────────

def test_explain_includes_formula():
    text = explain_formula("=SUM(A1:A10)")
    assert "SUM" in text
    assert "A1:A10" in text


def test_explain_includes_function_description():
    text = explain_formula("=VLOOKUP(A1,B1:C10,2,0)")
    assert "查詢" in text or "VLOOKUP" in text


def test_explain_non_formula():
    text = explain_formula("hello")
    assert "不是公式" in text or "=" in text


def test_explain_empty():
    text = explain_formula("")
    assert text  # 不應是空字串


def test_explain_multiple_functions():
    text = explain_formula("=IF(ISBLANK(A1),\"\",SUM(A1:A10))")
    assert "IF" in text
    assert "ISBLANK" in text
    assert "SUM" in text


def test_explain_formula_tool_returns_dict():
    result = explain_formula_tool("=AVERAGE(B2:B20)")
    assert isinstance(result, dict)
    assert "explanation" in result
    assert "functions_used" in result
    assert "AVERAGE" in result["functions_used"]
    assert "formula" in result


def test_explain_formula_tool_valid_flag():
    ok_result  = explain_formula_tool("=SUM(A1:A5)")
    bad_result = explain_formula_tool("=SUM(A1:A5")
    assert ok_result["valid"] is True
    assert bad_result["valid"] is False


def test_explain_cross_sheet_ref():
    text = explain_formula("=SUM(Sheet2!A1:A10)")
    assert "Sheet2" in text or "A1:A10" in text
