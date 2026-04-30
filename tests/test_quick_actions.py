from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from ui.quick_actions import QUICK_ACTIONS, queue_quick_action, quick_action_prompt


def test_quick_actions_are_fixed_and_user_facing():
    assert [action.key for action in QUICK_ACTIONS] == [
        "beautify_report",
        "summarize_data",
        "sum_by_group",
        "create_report_chart",
        "undo_last",
    ]
    assert [action.label for action in QUICK_ACTIONS] == [
        "美化目前表格",
        "產生資料摘要",
        "查詢加總資料",
        "建立報告圖表",
        "復原上一步",
    ]


def test_quick_action_prompts_do_not_expose_internal_tools():
    tool_names = {
        "beautify_range",
        "summarize_range",
        "query_range",
        "create_chart",
        "undo_last",
        "write_range",
    }

    for action in QUICK_ACTIONS:
        assert action.prompt
        assert "呼叫" not in action.prompt
        assert "工具" not in action.prompt
        assert not any(tool_name in action.prompt for tool_name in tool_names)


def test_queue_quick_action_uses_existing_chat_flow():
    state = {}

    prompt = queue_quick_action(state, "sum_by_group")

    assert prompt == quick_action_prompt("sum_by_group")
    assert state == {"_queued_prompt": prompt}
