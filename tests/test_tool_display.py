from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from ui.tool_display import friendly_tool_label, sanitize_assistant_text


def test_friendly_tool_label_hides_internal_tool_name():
    assert friendly_tool_label("query_range") == "已查詢資料"
    assert friendly_tool_label("beautify_range") == "已美化表格"
    assert "_" not in friendly_tool_label("write_range")


def test_sanitize_assistant_text_replaces_tool_checklist():
    text = """smart_toy

已完成。

check

✅ get_used_range

✅ query_range

已呼叫 `write_range` 完成。"""

    cleaned = sanitize_assistant_text(text)

    assert "smart_toy" not in cleaned
    assert "check" not in cleaned
    assert "get_used_range" not in cleaned
    assert "query_range" not in cleaned
    assert "write_range" not in cleaned
    assert "✅ 已確認資料範圍" in cleaned
    assert "✅ 已查詢資料" in cleaned
    assert "已呼叫 已寫入 Excel 完成。" in cleaned
