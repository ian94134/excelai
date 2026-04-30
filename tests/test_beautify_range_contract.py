from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import backup
from excel.format import _looks_like_header_row
import tools.executor as executor
from tools.definition import OPENAI_TOOLS


def _tool_schema(name: str) -> dict:
    for tool in OPENAI_TOOLS:
        fn = tool.get("function", {})
        if fn.get("name") == name:
            return fn
    raise AssertionError(f"missing tool schema: {name}")


def test_beautify_range_schema_is_available():
    schema = _tool_schema("beautify_range")
    params = schema["parameters"]
    props = params["properties"]

    assert len(OPENAI_TOOLS) == 72
    assert params["required"] == ["range_addr"]
    assert props["theme"]["enum"] == ["blue", "green", "gray", "orange", "purple"]
    assert "一鍵美化" in schema["description"]
    assert "不要再自動呼叫 apply_table_style" in schema["description"]


def test_apply_table_style_schema_is_not_default_beautify_path():
    schema = _tool_schema("apply_table_style")

    assert "正式表格（ListObject）" in schema["description"]
    assert "不要把一般的「美化、變漂亮、整理成報表」自動導向此工具" in schema["description"]


def test_header_row_detection_prefers_existing_unique_text_headers():
    assert _looks_like_header_row(
        ["Date", "Region", "Amount"],
        ["2026-01-01", "North", 120],
    ) is True
    assert _looks_like_header_row(["Date", "Date"], ["2026-01-01", 120]) is False
    assert _looks_like_header_row(["North", 120, "Alpha"], ["South", 80, "Beta"]) is False


def test_beautify_range_is_backup_tracked():
    assert backup.BACKUP_NEEDED["beautify_range"] is True


def test_executor_dispatches_beautify_range(monkeypatch):
    calls = []

    def fake_beautify_range(**kwargs):
        calls.append(kwargs)
        return {
            "status": "ok",
            "tool": "beautify_range",
            "range": kwargs["range_addr"],
            "theme": kwargs.get("theme", "blue"),
        }

    monkeypatch.setattr(executor.backup, "capture_before", lambda *args, **kwargs: None)
    monkeypatch.setattr(executor.telemetry, "record", lambda *args, **kwargs: None)
    monkeypatch.setattr(executor.et, "beautify_range", fake_beautify_range)

    raw = executor.execute(
        "beautify_range",
        {"sheet": "SalesData", "range_addr": "A1:G10", "theme": "green"},
    )
    payload = json.loads(raw)

    assert payload["status"] == "ok"
    assert payload["tool"] == "beautify_range"
    assert payload["theme"] == "green"
    assert calls == [{"sheet": "SalesData", "range_addr": "A1:G10", "theme": "green"}]
