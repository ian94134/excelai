from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from ui.quick_actions import (
    DEFAULT_QUICK_ACTION_OPTIONS,
    QUICK_ACTIONS,
    QUICK_ACTION_FORM_CHOICES,
    build_quick_action_prompt,
    clear_quick_action_form,
    open_quick_action_form,
    queue_quick_action,
    quick_action_options,
    quick_action_prompt,
)


INTERNAL_TERMS = {
    "beautify_range",
    "summarize_range",
    "query_range",
    "create_chart",
    "undo_last",
    "write_range",
}


def assert_user_facing_prompt(prompt: str):
    assert prompt
    assert "呼叫" not in prompt
    assert "工具" not in prompt
    assert not any(term in prompt for term in INTERNAL_TERMS)


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
    assert set(DEFAULT_QUICK_ACTION_OPTIONS) == {action.key for action in QUICK_ACTIONS}


def test_quick_action_prompts_do_not_expose_internal_tools():
    for action in QUICK_ACTIONS:
        prompt = quick_action_prompt(action.key)
        assert_user_facing_prompt(prompt)
        assert "完成後" in prompt


def test_quick_action_form_choices_cover_configurable_actions():
    assert QUICK_ACTION_FORM_CHOICES["beautify_report"]["theme"] == ("藍色", "綠色", "灰色")
    assert QUICK_ACTION_FORM_CHOICES["summarize_data"]["depth"] == ("標準", "快速", "詳細")
    assert "地區" in QUICK_ACTION_FORM_CHOICES["sum_by_group"]["group_by"]
    assert "長條圖" in QUICK_ACTION_FORM_CHOICES["create_report_chart"]["chart_type"]


def test_configured_quick_action_prompts_include_user_options():
    prompt = build_quick_action_prompt(
        "sum_by_group",
        {"group_by": "地區", "value_col": "金額", "include_total": False},
    )
    assert_user_facing_prompt(prompt)
    assert "地區" in prompt
    assert "金額" in prompt
    assert "不用顯示總計" in prompt

    chart_prompt = build_quick_action_prompt(
        "create_report_chart",
        {"chart_type": "長條圖", "placement": "新工作表", "include_title": True},
    )
    assert_user_facing_prompt(chart_prompt)
    assert "長條圖" in chart_prompt
    assert "新工作表" in chart_prompt

    beautify_prompt = build_quick_action_prompt(
        "beautify_report",
        {"theme": "綠色", "freeze_header": False, "save_after": True},
    )
    assert_user_facing_prompt(beautify_prompt)
    assert "綠色" in beautify_prompt
    assert "不要凍結表頭" in beautify_prompt
    assert "儲存檔案" in beautify_prompt


def test_quick_action_options_merge_defaults():
    options = quick_action_options("summarize_data", {"depth": "詳細"})

    assert options == {
        "depth": "詳細",
        "include_recommendations": True,
    }


def test_queue_quick_action_uses_existing_chat_flow():
    state = {}

    prompt = queue_quick_action(state, "sum_by_group", {"group_by": "類別"})

    assert prompt == build_quick_action_prompt("sum_by_group", {"group_by": "類別"})
    assert state["_queued_prompt"] == prompt
    assert state["_quick_action_last"]["key"] == "sum_by_group"
    assert state["_quick_action_last"]["label"] == "查詢加總資料"
    assert state["_quick_action_last"]["options"]["group_by"] == "類別"


def test_quick_action_form_state_helpers():
    state = {}

    open_quick_action_form(state, "create_report_chart")
    assert state["_quick_action_form"] == "create_report_chart"

    clear_quick_action_form(state)
    assert "_quick_action_form" not in state
