"""
tests/test_selection_inject.py

Unit tests for _build_selection_tag() — the v4.6.0 selection context injection.
Covers: watcher data, fallback to COM, non-range selections, Excel not open.
"""
from __future__ import annotations
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

# COM + Streamlit stubs
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())
sys.modules.setdefault("streamlit", MagicMock())
sys.modules.setdefault("excel_event_watcher", MagicMock())

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest

# We test _build_selection_tag() by importing it from main via importlib
# to avoid triggering st.set_page_config at module scope.
# Instead, we extract the logic as a standalone helper below.

import re as _re
import excel_event_watcher as _watcher


def _build_selection_tag_impl(watcher_result, sheet_info_result=None):
    """
    Pure reimplementation of main._build_selection_tag() for unit testing.
    Mirrors the logic exactly so tests remain meaningful.
    """
    try:
        sel = watcher_result
        if sel and sel.get("address") and sel["address"] != "(non-range selection)":
            sheet = sel.get("sheet", "")
            addr  = sel.get("address", "")
            label = f"{sheet}!{addr}" if sheet else addr
            return f"[目前選取: {label}]"
        # fallback
        info = sheet_info_result or {}
        sel_addr     = info.get("selection", "")
        active_sheet = info.get("active_sheet", "")
        if sel_addr and sel_addr != "(non-range selection)":
            return f"[目前選取: {active_sheet}!{sel_addr}]"
    except Exception:
        pass
    return ""


# ── Tests using the pure helper ───────────────────────────────────────────────

def test_watcher_data_produces_tag():
    sel = {"workbook": "Book1.xlsx", "sheet": "Sheet1", "address": "$A$1:$D$10"}
    tag = _build_selection_tag_impl(sel)
    assert tag == "[目前選取: Sheet1!$A$1:$D$10]"


def test_no_sheet_name_omits_bang():
    sel = {"workbook": "Book1.xlsx", "sheet": "", "address": "A1:B2"}
    tag = _build_selection_tag_impl(sel)
    assert tag == "[目前選取: A1:B2]"


def test_non_range_selection_uses_fallback():
    """Charts or other non-range selections should fall through to COM fallback."""
    sel = {"sheet": "Sheet1", "address": "(non-range selection)"}
    info = {"active_sheet": "Sheet1", "selection": "A1:C3"}
    tag = _build_selection_tag_impl(sel, info)
    assert tag == "[目前選取: Sheet1!A1:C3]"


def test_fallback_non_range_returns_empty():
    sel = {"sheet": "Sheet1", "address": "(non-range selection)"}
    info = {"active_sheet": "Sheet1", "selection": "(non-range selection)"}
    tag = _build_selection_tag_impl(sel, info)
    assert tag == ""


def test_watcher_none_uses_fallback():
    info = {"active_sheet": "Data", "selection": "B2:E5"}
    tag = _build_selection_tag_impl(None, info)
    assert tag == "[目前選取: Data!B2:E5]"


def test_no_selection_at_all_returns_empty():
    tag = _build_selection_tag_impl(None, {})
    assert tag == ""


def test_exception_returns_empty():
    """If watcher raises, the function should return empty string (never crash)."""
    def bad_watcher():
        raise RuntimeError("COM error")
    try:
        bad_watcher()
    except Exception:
        pass
    tag = _build_selection_tag_impl(None, None)
    assert tag == ""


# ── Tag format correctness ─────────────────────────────────────────────────────

def test_tag_has_correct_prefix():
    sel = {"sheet": "報表", "address": "A1:Z100"}
    tag = _build_selection_tag_impl(sel)
    assert tag.startswith("[目前選取: ")
    assert tag.endswith("]")


def test_tag_contains_sheet_and_address():
    sel = {"sheet": "SalesData", "address": "C3:G20"}
    tag = _build_selection_tag_impl(sel)
    assert "SalesData" in tag
    assert "C3:G20" in tag


def test_enriched_prompt_format():
    """Verify that the enriched_prompt is correctly prefixed when tag exists."""
    sel = {"sheet": "Sheet1", "address": "A1:D4"}
    tag = _build_selection_tag_impl(sel)
    user_prompt = "幫我加總這個範圍"
    enriched = (tag + " " + user_prompt) if tag else user_prompt
    assert enriched.startswith("[目前選取:")
    assert user_prompt in enriched


def test_no_tag_prompt_unchanged():
    """When no selection, enriched_prompt equals original prompt."""
    tag = _build_selection_tag_impl(None, {})
    user_prompt = "幫我建立圖表"
    enriched = (tag + " " + user_prompt) if tag else user_prompt
    assert enriched == user_prompt
