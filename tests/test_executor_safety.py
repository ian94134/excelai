from __future__ import annotations

import json
from unittest.mock import patch

from tools import executor


def test_dangerous_tool_requires_executor_confirmation(monkeypatch):
    called = False

    def _tool(_args):
        nonlocal called
        called = True
        return {"status": "ok"}

    monkeypatch.setitem(executor.TOOL_MAP, "clear_range", _tool)

    with patch("tools.executor.backup.capture_before") as capture:
        raw = executor.execute("clear_range", {"range_addr": "A1"})

    payload = json.loads(raw)
    assert payload["requires_confirmation"] is True
    assert payload["error_type"] == "DangerousToolRequiresConfirmation"
    assert called is False
    capture.assert_not_called()


def test_confirmed_dangerous_tool_executes_without_confirmation_arg(monkeypatch):
    captured = {}

    def _tool(args):
        captured.update(args)
        return {"status": "ok"}

    monkeypatch.setitem(executor.TOOL_MAP, "clear_range", _tool)

    with patch("tools.executor.backup.capture_before", return_value=None):
        raw = executor.execute(
            "clear_range",
            {"range_addr": "A1", "confirm_dangerous": True},
        )

    payload = json.loads(raw)
    assert payload["status"] == "ok"
    assert captured == {"range_addr": "A1"}


def test_tool_status_error_is_normalized_and_does_not_push_backup(monkeypatch):
    def _tool(_args):
        return {
            "status": "error",
            "message": "needs attention",
            "error_type": "ToolSpecificError",
        }

    monkeypatch.setitem(executor.TOOL_MAP, "format_range", _tool)

    with patch("tools.executor.backup.capture_before", return_value=object()), \
         patch("tools.executor.backup.get_session_stack") as get_stack:
        raw = executor.execute("format_range", {"range_addr": "A1"})

    payload = json.loads(raw)
    assert payload["status"] == "error"
    assert payload["error"] == "needs attention"
    assert payload["error_type"] == "ToolSpecificError"
    get_stack.assert_not_called()
