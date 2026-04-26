"""
macro 模組單元測試（v4.7.0）

測試範圍：
- record_macro：從 steps 參數錄製、空名稱拒絕、步驟為空拒絕、覆蓋同名巨集
- list_macros：空時回傳空清單、多個巨集均列出
- run_macro：委派給 execute_batch、失敗時正確回傳 error
- delete_macro：成功刪除、找不到時回傳 error
- 持久化：_save_macros / _load_macros 版本驗證
- get_macro_steps：找到 / 找不到
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

# ── Win32 stub（Linux CI）────────────────────────────────────────────────────
for _mod in ("pythoncom", "pywintypes", "win32com", "win32com.client", "win32con", "win32api"):
    sys.modules.setdefault(_mod, MagicMock())

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# pre-import so that patch("backup.*") and patch("tools.executor.*") can
# resolve correctly via sys.modules inside the tests package context
import backup  # noqa: F401
import tools.executor  # noqa: F401

import macro
from macro import (
    record_macro, list_macros, run_macro, delete_macro,
    get_macro_steps, _load_macros, _save_macros, _MACROS_PATH, _MACROS_VERSION,
)


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def isolated_macros_file(tmp_path, monkeypatch):
    """每個測試使用獨立的臨時 macros.json，不干擾真實磁碟。"""
    fake_path = tmp_path / "macros.json"
    monkeypatch.setattr(macro, "_MACROS_PATH", fake_path)
    yield fake_path


SAMPLE_STEPS = [
    {"tool": "format_range", "args": {"range_addr": "A1:D1", "bold": True}},
    {"tool": "auto_fit",     "args": {"target": "columns"}},
]

DANGEROUS_STEPS = [
    {"tool": "clear_range", "args": {"range_addr": "A1:D10", "target": "all"}},
]


# ── record_macro ──────────────────────────────────────────────────────────────

def test_record_macro_with_steps():
    result = record_macro("test_macro", "測試巨集", steps=SAMPLE_STEPS)
    assert result["status"] == "ok"
    assert result["name"] == "test_macro"
    assert result["step_count"] == 2


def test_record_macro_empty_name_rejected():
    result = record_macro("", steps=SAMPLE_STEPS)
    assert result["status"] == "error"
    assert "名稱" in result["message"]


def test_record_macro_whitespace_name_rejected():
    result = record_macro("   ", steps=SAMPLE_STEPS)
    assert result["status"] == "error"


def test_record_macro_empty_steps_rejected():
    result = record_macro("empty", steps=[])
    assert result["status"] == "error"
    assert "空" in result["message"]


def test_record_macro_overwrites_existing():
    record_macro("dup", steps=SAMPLE_STEPS)
    new_steps = [{"tool": "save_workbook", "args": {}}]
    result = record_macro("dup", steps=new_steps)
    assert result["status"] == "ok"
    assert result["step_count"] == 1
    # 驗證磁碟已更新
    loaded = _load_macros()
    assert len(loaded["dup"]["steps"]) == 1


def test_record_macro_persists_to_disk():
    record_macro("persist_test", description="說明", steps=SAMPLE_STEPS)
    loaded = _load_macros()
    assert "persist_test" in loaded
    assert loaded["persist_test"]["description"] == "說明"


def test_record_macro_from_backup_stack():
    """從 BackupStack 自動取得步驟的路徑。"""
    fake_entry = MagicMock()
    fake_entry.tool_name = "write_range"
    fake_entry.arguments = {"range_addr": "A1", "values": [["x"]]}

    fake_stack = MagicMock()
    fake_stack.__len__ = lambda self: 1
    fake_stack.snapshot.return_value = [fake_entry]

    # get_session_stack 在 macro.py 函式內懶載入，patch 來源模組
    with patch("backup.get_session_stack", return_value=fake_stack):
        result = record_macro("from_stack")

    assert result["status"] == "ok"
    assert result["step_count"] == 1


def test_record_macro_from_empty_stack_rejected():
    fake_stack = MagicMock()
    fake_stack.__len__ = lambda self: 0

    with patch("backup.get_session_stack", return_value=fake_stack):
        result = record_macro("fail_stack")

    assert result["status"] == "error"
    assert "歷史" in result["message"]


# ── list_macros ───────────────────────────────────────────────────────────────

def test_list_macros_empty():
    result = list_macros()
    assert result["status"] == "ok"
    assert result["macros"] == []
    assert result["count"] == 0


def test_list_macros_multiple():
    record_macro("m1", steps=SAMPLE_STEPS)
    record_macro("m2", steps=SAMPLE_STEPS[:1])
    result = list_macros()
    assert result["count"] == 2
    names = {m["name"] for m in result["macros"]}
    assert "m1" in names and "m2" in names


def test_list_macros_includes_step_count():
    record_macro("counted", steps=SAMPLE_STEPS)
    result = list_macros()
    m = next(m for m in result["macros"] if m["name"] == "counted")
    assert m["step_count"] == 2


# ── run_macro ─────────────────────────────────────────────────────────────────

def test_run_macro_delegates_to_execute_batch():
    record_macro("runner", steps=SAMPLE_STEPS)
    # execute_batch 在 macro.py 函式內懶載入，patch 來源模組
    with patch("tools.executor.execute_batch") as mock_batch:
        mock_batch.return_value = [
            {"tool": "format_range", "result": {"status": "ok"}, "rolled_back": False},
            {"tool": "auto_fit",     "result": {"status": "ok"}, "rolled_back": False},
        ]
        result = run_macro("runner")

    assert result["status"] == "ok"
    assert result["total_steps"] == 2
    mock_batch.assert_called_once()


def test_run_macro_not_found():
    result = run_macro("nonexistent")
    assert result["status"] == "error"
    assert "找不到" in result["message"]


def test_run_macro_empty_name():
    result = run_macro("")
    assert result["status"] == "error"


def test_run_macro_reports_failure():
    record_macro("fail_runner", steps=SAMPLE_STEPS)
    with patch("tools.executor.execute_batch") as mock_batch:
        mock_batch.return_value = [
            {"tool": "format_range", "result": {"error": "sheet not found"}, "rolled_back": True},
        ]
        result = run_macro("fail_runner")

    assert result["status"] == "error"
    assert "失敗" in result["message"]


def test_run_macro_reports_status_error_failure():
    record_macro("fail_runner", steps=SAMPLE_STEPS)
    with patch("tools.executor.execute_batch") as mock_batch:
        mock_batch.return_value = [
            {
                "tool": "format_range",
                "result": {"status": "error", "message": "blocked"},
                "rolled_back": True,
            },
        ]
        result = run_macro("fail_runner")

    assert result["status"] == "error"
    assert "失敗" in result["message"]


def test_run_macro_blocks_dangerous_steps_by_default():
    record_macro("danger_runner", steps=DANGEROUS_STEPS)
    with patch("tools.executor.execute_batch") as mock_batch:
        result = run_macro("danger_runner")

    assert result["status"] == "error"
    assert result["requires_confirmation"] is True
    assert result["error_type"] == "DangerousMacroRequiresConfirmation"
    assert result["dangerous_steps"][0]["tool"] == "clear_range"
    mock_batch.assert_not_called()


def test_run_macro_confirmed_dangerous_steps_execute():
    record_macro("danger_runner", steps=DANGEROUS_STEPS)
    with patch("tools.executor.execute_batch") as mock_batch:
        mock_batch.return_value = [
            {"tool": "clear_range", "result": {"status": "ok"}, "rolled_back": False},
        ]
        result = run_macro("danger_runner", confirm_dangerous=True)

    assert result["status"] == "ok"
    mock_batch.assert_called_once_with(DANGEROUS_STEPS, confirm_dangerous=True)


def test_executor_run_macro_ignores_untrusted_confirmation_arg():
    record_macro("danger_runner", steps=DANGEROUS_STEPS)
    from tools.executor import execute

    with patch("tools.executor.execute_batch") as mock_batch:
        raw = execute("run_macro", {"name": "danger_runner", "confirm_dangerous": True})

    result = json.loads(raw)
    assert result["requires_confirmation"] is True
    mock_batch.assert_not_called()


# ── delete_macro ──────────────────────────────────────────────────────────────

def test_delete_macro_success():
    record_macro("to_delete", steps=SAMPLE_STEPS)
    result = delete_macro("to_delete")
    assert result["status"] == "ok"
    assert _load_macros() == {}


def test_delete_macro_not_found():
    result = delete_macro("ghost")
    assert result["status"] == "error"
    assert "找不到" in result["message"]


def test_delete_macro_empty_name():
    result = delete_macro("")
    assert result["status"] == "error"


# ── get_macro_steps ───────────────────────────────────────────────────────────

def test_get_macro_steps_found():
    record_macro("steps_test", steps=SAMPLE_STEPS)
    steps = get_macro_steps("steps_test")
    assert len(steps) == 2
    assert steps[0]["tool"] == "format_range"


def test_get_macro_steps_not_found():
    steps = get_macro_steps("no_such_macro")
    assert steps == []


# ── 持久化版本驗證 ────────────────────────────────────────────────────────────

def test_load_macros_version_mismatch_returns_empty(isolated_macros_file):
    isolated_macros_file.write_text(
        json.dumps({"version": 999, "macros": {"x": {"steps": [], "step_count": 0}}}),
        encoding="utf-8",
    )
    loaded = _load_macros()
    assert loaded == {}


def test_load_macros_missing_file_returns_empty():
    # isolated_macros_file fixture 確保不存在
    loaded = _load_macros()
    assert loaded == {}


def test_save_then_load_round_trip():
    record_macro("round_trip", description="RT", steps=SAMPLE_STEPS)
    loaded = _load_macros()
    assert "round_trip" in loaded
    assert loaded["round_trip"]["description"] == "RT"
    assert len(loaded["round_trip"]["steps"]) == 2
