"""
tests/test_session_compress.py

Unit tests for compress_tool_result() — the v4.6.0 tool result compression.
Pure function tests: no Streamlit, no COM, no network.
"""
from __future__ import annotations
import sys, json
from pathlib import Path
from unittest.mock import MagicMock

# COM stubs
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())

# Streamlit stub (session.py imports it at module level)
sys.modules.setdefault("streamlit", MagicMock())

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
from compress import compress_tool_result, TOOL_CONTENT_LIMIT as _TOOL_CONTENT_LIMIT


# ── Helpers ───────────────────────────────────────────────────────────────────

def _json(obj) -> str:
    return json.dumps(obj, ensure_ascii=False)

LIMIT = _TOOL_CONTENT_LIMIT


# ── Short content: returned unchanged ─────────────────────────────────────────

def test_short_content_unchanged():
    content = _json({"status": "ok", "sheet": "Sheet1"})
    assert compress_tool_result(content, limit=LIMIT) == content


def test_exactly_at_limit_unchanged():
    content = "x" * LIMIT
    assert compress_tool_result(content, limit=LIMIT) == content


# ── Error results: kept as-is ──────────────────────────────────────────────────

def test_error_result_not_compressed():
    """Error results must be returned unchanged so Qwen has full context."""
    payload = {"error": "SheetNotFoundError", "error_type": "SheetNotFoundError",
               "hint": "A" * 5000}   # even very large
    content = _json(payload)
    assert compress_tool_result(content, limit=100) == content


def test_error_result_with_large_body_not_compressed():
    payload = {"error": "something", "data": list(range(1000))}
    content = _json(payload)
    result = compress_tool_result(content, limit=50)
    assert json.loads(result)["error"] == "something"


# ── Success dict: large arrays pruned ─────────────────────────────────────────

def test_large_array_field_summarised():
    big_list = list(range(500))
    payload = {"status": "ok", "values": big_list, "sheet": "Sheet1"}
    content = _json(payload)
    # content is ~2437 chars; use limit=2000 so the overall limit is exceeded
    # and field-level compression kicks in for the large array
    result = compress_tool_result(content, limit=2000)
    parsed = json.loads(result)
    assert isinstance(parsed["values"], str)
    assert "500" in parsed["values"]          # row count mentioned
    assert parsed["sheet"] == "Sheet1"        # other fields preserved
    assert parsed.get("_compressed") is True


def test_large_string_field_truncated():
    payload = {"status": "ok", "description": "A" * 1000}
    content = _json(payload)
    # content is ~1035 chars; use limit=600 so:
    # - overall limit is exceeded → compression logic runs
    # - compressed dict (description truncated to 500 + suffix ~557 chars) fits within 600
    # - json.loads() succeeds on the returned valid JSON
    result = compress_tool_result(content, limit=600)
    parsed = json.loads(result)
    assert len(parsed["description"]) < 1000
    assert "已截斷" in parsed["description"]


def test_small_array_field_preserved():
    payload = {"status": "ok", "headers": ["A", "B", "C"]}
    content = _json(payload)
    result = compress_tool_result(content, limit=LIMIT)
    parsed = json.loads(result)
    assert parsed["headers"] == ["A", "B", "C"]


def test_compressed_flag_present_after_compression():
    big_list = list(range(1000))
    payload = {"data": big_list}
    content = _json(payload)
    result = compress_tool_result(content, limit=LIMIT)
    parsed = json.loads(result)
    assert parsed.get("_compressed") is True


def test_no_compression_flag_when_unchanged():
    payload = {"status": "ok"}
    content = _json(payload)
    result = compress_tool_result(content, limit=LIMIT)
    parsed = json.loads(result)
    assert "_compressed" not in parsed


# ── List payload (e.g. read_range 2D array) ───────────────────────────────────

def test_large_list_payload_summarised():
    data = [[i, i+1] for i in range(500)]   # 500 rows
    content = _json(data)
    result = compress_tool_result(content, limit=LIMIT)
    parsed = json.loads(result)
    assert parsed["rows"] == 500
    assert parsed["columns"] == 2
    assert len(parsed["sample"]) <= 3
    assert parsed.get("_compressed") is True


def test_small_list_payload_unchanged():
    data = [[1, 2], [3, 4]]
    content = _json(data)
    assert compress_tool_result(content, limit=LIMIT) == content


# ── Non-JSON content ──────────────────────────────────────────────────────────

def test_non_json_truncated_with_note():
    content = "X" * (LIMIT + 100)
    result = compress_tool_result(content, limit=LIMIT)
    assert len(result) > LIMIT           # truncation marker adds some chars
    assert "已截斷" in result


def test_non_json_short_unchanged():
    content = "plain text ok"
    assert compress_tool_result(content, limit=LIMIT) == content


# ── Custom limit parameter ────────────────────────────────────────────────────

def test_custom_limit_respected():
    payload = {"status": "ok", "big": "Z" * 2000}
    content = _json(payload)
    result = compress_tool_result(content, limit=200)
    assert len(result) <= 250    # slight overrun OK due to truncation suffix
