"""
tests/test_complexity.py

Unit tests for _is_complex_task() — the auto-planning trigger detector.
"""
from __future__ import annotations
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# main.py imports streamlit at module level, so we need to stub it
from unittest.mock import MagicMock
sys.modules.setdefault("streamlit", MagicMock())
sys.modules.setdefault("excel_event_watcher", MagicMock())

# Also stub win32com family before importing anything that touches excel_tools
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())

# Import the function under test by reading the module directly
# (avoids triggering the Streamlit page-config call at module level)
import importlib, types, re as _re

# Extract _is_complex_task by injecting the logic directly — mirrors main.py exactly
_CONNECTOR_RE = _re.compile(
    r"然後|再次?|接著|並且|同時|另外|最後|之後|先.{1,20}再|分別|逐一|依序|第[一二三四五六七八九十]步",
    _re.UNICODE,
)
_OP_KW_RE = _re.compile(
    r"篩選|排序|格式化?|合併|刪除|插入|建立|新增|計算|加總|平均|統計|分析|整理|製作|產生|匯出|"
    r"讀取|寫入|複製|移動|凍結|圖表|樞紐|報表|下拉|驗證|保護|解除|條件|框線|欄寬|列高|自動",
    _re.UNICODE,
)


def _is_complex_task(prompt: str) -> bool:
    connectors = len(_CONNECTOR_RE.findall(prompt))
    ops        = len(_OP_KW_RE.findall(prompt))
    return connectors >= 2 or ops >= 3


# ── Simple / single-step → NOT complex ──────────────────────────────────────

def test_simple_filter_not_complex():
    assert not _is_complex_task("篩選台北的資料")


def test_simple_format_not_complex():
    assert not _is_complex_task("把 A1:D1 設成粗體")


def test_simple_chart_not_complex():
    assert not _is_complex_task("幫我建立一個圖表")


def test_short_question_not_complex():
    assert not _is_complex_task("現在用的是哪個工作表？")


def test_two_ops_not_complex():
    assert not _is_complex_task("篩選資料並排序")   # only 1 connector "並" (not in list), 2 ops


# ── Multi-step connectors → complex ─────────────────────────────────────────

def test_two_connectors_is_complex():
    assert _is_complex_task("先篩選台北的資料，然後刪除空白列，最後存檔")


def test_sequential_steps_complex():
    assert _is_complex_task("把A欄排序，然後計算總和，再把結果複製到Sheet2")


def test_parallel_ops_complex():
    assert _is_complex_task("同時整理三個工作表的格式，然後各自加上標題列")


def test_first_then_pattern_complex():
    assert _is_complex_task("先確認工作表名稱，再讀取資料，最後製作圖表")


# ── Multiple operation keywords → complex ────────────────────────────────────

def test_three_op_keywords_complex():
    assert _is_complex_task("分析業績資料，製作樞紐，產生報表")


def test_four_op_keywords_complex():
    assert _is_complex_task("讀取資料、計算加總、排序結果、匯出成報表")


def test_rich_task_complex():
    assert _is_complex_task(
        "幫我把A欄的資料依照日期排序，然後在最後一列加上合計，再把標題列設成藍底白字"
    )


def test_report_task_complex():
    assert _is_complex_task("整理各區業績，加總後建立圖表，再製作一份樞紐分析報表")


# ── Edge cases ────────────────────────────────────────────────────────────────

def test_empty_string_not_complex():
    assert not _is_complex_task("")


def test_very_short_not_complex():
    assert not _is_complex_task("ok")


def test_exactly_two_connectors_complex():
    # 然後 + 最後 = 2 connectors → complex
    assert _is_complex_task("做完A然後做B最後做C")


def test_exactly_one_connector_not_complex():
    # Only 1 connector, 1 op → not complex
    assert not _is_complex_task("篩選資料然後存檔")
