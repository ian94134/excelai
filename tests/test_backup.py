"""
backup 模組單元測試。

目的：
- BackupStack 達上限後會丟棄最舊項目（20 步 FIFO）
- push / pop / peek / clear / __len__ / snapshot 行為正確
- capture_before 對唯讀工具回傳 None、對可還原工具回傳 BackupEntry
- BackupEntry 預設 timestamp 為 UTC 且會被填入
- restore() Phase 2 已實作；對有 values_before 的 entry 呼叫 et.write_range
"""

from __future__ import annotations
import sys
from datetime import datetime, timezone
from unittest.mock import MagicMock, patch

# Stub Windows-only COM modules BEFORE any project import
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _mod in _WIN_STUBS:
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()

from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
import backup
from backup import (
    BACKUP_NEEDED,
    BACKUP_STACK_MAX,
    BackupEntry,
    BackupStack,
    capture_before,
    restore,
)


# ── BackupEntry dataclass ────────────────────────────────────────────────────

def test_backup_entry_defaults_timestamp_utc():
    entry = BackupEntry(tool_name="write_range", arguments={"a": 1})
    assert entry.tool_name == "write_range"
    assert entry.arguments == {"a": 1}
    assert isinstance(entry.timestamp, datetime)
    assert entry.timestamp.tzinfo is not None
    assert entry.values_before is None
    assert entry.formats_before is None


def test_backup_entry_describe_contains_tool_name():
    entry = BackupEntry(tool_name="delete_row", arguments={})
    assert "delete_row" in entry.describe()


# ── capture_before ───────────────────────────────────────────────────────────

def test_capture_before_returns_none_for_readonly_tools():
    assert capture_before("read_range", {"range_addr": "A1:B2"}) is None
    assert capture_before("get_sheet_info", {}) is None
    assert capture_before("get_used_range", {}) is None
    assert capture_before("save_workbook", {}) is None


def test_capture_before_returns_entry_for_mutating_tools():
    entry = capture_before("write_range", {"range_addr": "A1", "values": [[1]]})
    assert isinstance(entry, BackupEntry)
    assert entry.tool_name == "write_range"
    assert entry.arguments == {"range_addr": "A1", "values": [[1]]}


def test_capture_before_copies_arguments():
    args = {"sheet_name": "Sheet1"}
    entry = capture_before("add_sheet", args)
    assert entry is not None
    args["sheet_name"] = "CHANGED"
    assert entry.arguments["sheet_name"] == "Sheet1"


def test_capture_before_unknown_tool_returns_none():
    assert capture_before("not_a_real_tool", {}) is None


def test_backup_needed_readonly_flags():
    readonly = ["read_range", "get_sheet_info", "get_used_range", "save_workbook"]
    for name in readonly:
        assert BACKUP_NEEDED.get(name) is False, f"{name} should be False (read-only)"


# ── BackupStack ──────────────────────────────────────────────────────────────

def _entry(name: str) -> BackupEntry:
    return BackupEntry(tool_name=name, arguments={})


def test_stack_push_pop_peek():
    stack = BackupStack()
    assert len(stack) == 0
    assert stack.peek() is None
    assert stack.pop() is None

    e1, e2 = _entry("t1"), _entry("t2")
    stack.push(e1)
    stack.push(e2)
    assert len(stack) == 2
    assert stack.peek() is e2
    assert len(stack) == 2
    assert stack.pop() is e2
    assert stack.pop() is e1
    assert stack.pop() is None


def test_stack_enforces_max_size_fifo_drop():
    stack = BackupStack(max_size=3)
    for i in range(5):
        stack.push(_entry(f"t{i}"))
    assert len(stack) == 3
    remaining = [e.tool_name for e in stack]
    assert remaining == ["t2", "t3", "t4"]


def test_stack_default_max_is_20():
    assert BACKUP_STACK_MAX == 20
    stack = BackupStack()
    for i in range(25):
        stack.push(_entry(f"t{i}"))
    assert len(stack) == 20
    assert stack.peek().tool_name == "t24"


def test_stack_clear_and_snapshot():
    stack = BackupStack()
    stack.push(_entry("a"))
    stack.push(_entry("b"))
    snap = stack.snapshot()
    assert len(snap) == 2
    snap.clear()          # modifying snapshot does NOT affect stack
    assert len(stack) == 2
    stack.clear()
    assert len(stack) == 0
    assert stack.peek() is None


# ── restore (Phase 2 implemented) ────────────────────────────────────────────

def test_restore_values_before_calls_write_range():
    """restore() with values_before should delegate to et.write_range.

    excel_tools is a late import inside restore(), so we patch sys.modules.
    """
    entry = BackupEntry(
        tool_name="write_range",
        arguments={"range_addr": "A1:B2", "sheet": "Sheet1"},
        values_before=[[1, 2], [3, 4]],
    )
    mock_et = MagicMock()
    mock_et.write_range.return_value = {"status": "ok"}
    with patch.dict(sys.modules, {"excel_tools": mock_et}):
        result = restore(entry)
    mock_et.write_range.assert_called_once_with("A1:B2", [[1, 2], [3, 4]], "Sheet1")
    assert result["status"] == "ok"


def test_restore_no_values_delegates_to_undo_dispatch():
    """restore() without values_before falls through to _undo_dispatch.

    excel_tools is a late import inside restore(), so we patch sys.modules.
    """
    entry = BackupEntry(
        tool_name="insert_row",
        arguments={"row_index": 3, "sheet": "Sheet1"},
    )
    mock_et = MagicMock()
    mock_et._undo_dispatch.return_value = {"status": "ok", "undone": "insert_row"}
    with patch.dict(sys.modules, {"excel_tools": mock_et}):
        result = restore(entry)
    mock_et._undo_dispatch.assert_called_once_with(entry)
    assert result["undone"] == "insert_row"
