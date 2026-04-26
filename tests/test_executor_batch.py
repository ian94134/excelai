"""
tests/test_executor_batch.py

Unit tests for execute_batch() rollback behaviour.
All win32com / Excel calls are mocked out so these run on Linux CI.

execute_batch() API
-------------------
Input:  list of {"tool": str, "args": dict}
Output: list of {"tool": str, "result": dict, "rolled_back": bool}

On error in step N:
  - Steps 0..N-1 are retroactively marked rolled_back=True
  - Step N is appended with rolled_back=True
  - Steps N+1.. are never executed
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

# Stub Windows-only COM modules BEFORE any project import
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _mod in _WIN_STUBS:
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.executor import execute_batch  # noqa: E402


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ok(tool: str) -> str:
    return json.dumps({"status": "ok", "tool": tool})


def _err(tool: str, msg: str = "boom") -> str:
    return json.dumps({"error": msg, "tool": tool, "error_type": "TestError"})


def _status_err(tool: str, msg: str = "blocked") -> str:
    return json.dumps({
        "status": "error",
        "message": msg,
        "tool": tool,
        "error_type": "StatusError",
    })


def _steps(*names: str) -> list[dict]:
    return [{"tool": n, "args": {}} for n in names]


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def mock_exec():
    with patch("tools.executor.execute") as m:
        yield m


@pytest.fixture()
def no_stack():
    with patch("tools.executor.backup.get_session_stack", return_value=None):
        yield


@pytest.fixture()
def fake_stack():
    stk = MagicMock()
    stk.__len__ = MagicMock(return_value=0)
    stk.pop.return_value = MagicMock()
    with patch("tools.executor.backup.get_session_stack", return_value=stk), \
         patch("tools.executor.backup.restore", return_value={"status": "ok"}):
        yield stk


# ---------------------------------------------------------------------------
# Tests — success path
# ---------------------------------------------------------------------------

def test_all_success_length(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("a"), _ok("b"), _ok("c")]
    results = execute_batch(_steps("a", "b", "c"))
    assert len(results) == 3


def test_all_success_not_rolled_back(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("x"), _ok("y")]
    results = execute_batch(_steps("x", "y"))
    assert all(r["rolled_back"] is False for r in results)


def test_all_success_result_payload(mock_exec, no_stack):
    mock_exec.return_value = _ok("t")
    results = execute_batch(_steps("t"))
    assert results[0]["result"]["status"] == "ok"


def test_empty_steps_returns_empty(mock_exec, no_stack):
    results = execute_batch([])
    assert results == []
    mock_exec.assert_not_called()


def test_confirmed_batch_marks_dangerous_steps(mock_exec, no_stack):
    mock_exec.return_value = _ok("delete_row")
    results = execute_batch(
        [{"tool": "delete_row", "args": {"index": 3}}],
        confirm_dangerous=True,
    )
    assert results[0]["rolled_back"] is False
    mock_exec.assert_called_once_with(
        "delete_row",
        {"index": 3, "confirm_dangerous": True},
    )


# ---------------------------------------------------------------------------
# Tests — failure path
# ---------------------------------------------------------------------------

def test_stops_after_error(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("a"), _err("b"), _ok("c")]
    results = execute_batch(_steps("a", "b", "c"))
    assert len(results) == 2          # "c" never ran
    assert mock_exec.call_count == 2


def test_error_step_rolled_back(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("a"), _err("b")]
    results = execute_batch(_steps("a", "b"))
    assert results[1]["rolled_back"] is True


def test_prior_steps_also_rolled_back(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("a"), _ok("b"), _err("c")]
    results = execute_batch(_steps("a", "b", "c"))
    assert all(r["rolled_back"] is True for r in results)


def test_first_step_failure(mock_exec, no_stack):
    mock_exec.return_value = _err("only")
    results = execute_batch(_steps("only"))
    assert results[0]["rolled_back"] is True


def test_error_payload_preserved(mock_exec, no_stack):
    mock_exec.return_value = _err("t", "something went wrong")
    results = execute_batch(_steps("t"))
    assert results[0]["result"]["error"] == "something went wrong"


def test_status_error_payload_stops_batch(mock_exec, no_stack):
    mock_exec.side_effect = [_ok("a"), _status_err("b"), _ok("c")]
    results = execute_batch(_steps("a", "b", "c"))
    assert len(results) == 2
    assert all(r["rolled_back"] is True for r in results)
    assert results[1]["result"]["error"] == "blocked"
    assert mock_exec.call_count == 2


# ---------------------------------------------------------------------------
# Tests — stack rollback
# ---------------------------------------------------------------------------

def test_stack_pop_called_on_failure(mock_exec, fake_stack):
    # Simulate stack growing by 1 after the successful first step
    depth_vals = [0, 1, 1, 0]
    call_i = [0]

    def _len():
        v = depth_vals[min(call_i[0], len(depth_vals) - 1)]
        call_i[0] += 1
        return v

    fake_stack.__len__ = MagicMock(side_effect=_len)
    mock_exec.side_effect = [_ok("a"), _err("b")]
    execute_batch(_steps("a", "b"))
    fake_stack.pop.assert_called()


def test_no_stack_does_not_crash(mock_exec, no_stack):
    """None stack → rollback loop is skipped, function completes normally."""
    mock_exec.side_effect = [_ok("a"), _err("b")]
    results = execute_batch(_steps("a", "b"))
    assert len(results) == 2
