"""
tests/test_agent_clarify.py

Unit tests for EVT_CLARIFY — the v4.6.0 clarification-turn mechanism.
Covers: _is_clarification() logic, EVT_CLARIFY yielded from run_turn(),
        EVT_DONE still yielded for non-question responses,
        EVT_CLARIFY not triggered when tool calls precede the question,
        planning mode suppresses clarify detection.
"""
from __future__ import annotations
import sys, json
from pathlib import Path
from unittest.mock import MagicMock

# COM stubs
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
from agent import (
    run_turn, _is_clarification,
    EVT_CLARIFY, EVT_DONE, EVT_PLAN_READY,
    EVT_TOOL_START, EVT_TOOL_DONE, EVT_ASST_MSG,
)
from providers.base import LLMProvider, LLMResponse, ToolCall


# ── _is_clarification() unit tests ───────────────────────────────────────────

def test_clarification_with_question_mark():
    assert _is_clarification("請問您要篩選哪個範圍？")


def test_clarification_which_sheet():
    assert _is_clarification("您希望操作哪張工作表？")


def test_clarification_confirm():
    assert _is_clarification("是否要刪除整個欄位，還是只清除內容？")


def test_clarification_please_provide():
    assert _is_clarification("請提供您要合併的範圍，例如 A1:D10？")


def test_plain_statement_not_clarification():
    assert not _is_clarification("已將 A1:D10 的資料排序完成。")


def test_statement_with_question_word_no_mark_not_clarification():
    assert not _is_clarification("請問這個功能的用途")  # no ？ or ?


def test_empty_string_not_clarification():
    assert not _is_clarification("")


def test_task_complete_not_clarification():
    assert not _is_clarification("任務已完成，共處理了 50 列資料。")


def test_error_explanation_not_clarification():
    assert not _is_clarification("發生了 SheetNotFoundError，工作表不存在。")


def test_ascii_question_mark_accepted():
    assert _is_clarification("您要使用哪個工作表? 請告知名稱。")


# ── run_turn() yields EVT_CLARIFY ─────────────────────────────────────────────

def _fake_provider(text: str):
    """Provider that yields a single 'done' event with the given text."""
    provider = MagicMock(spec=LLMProvider)
    provider.chat_stream.return_value = iter([("done", text)])
    return provider


def _run(provider, messages=None, plan_inject="", tools=None):
    msgs = messages or [{"role": "user", "content": "hello"}]
    return list(run_turn(
        get_messages=lambda: msgs,
        tools=tools or [],
        provider=provider,
        plan_inject=plan_inject,
    ))


def test_clarify_event_yielded_for_question():
    events = _run(_fake_provider("請問您要操作哪個工作表？"))
    kinds = [k for k, _ in events]
    assert EVT_CLARIFY in kinds
    assert EVT_DONE not in kinds


def test_clarify_event_data_is_question_text():
    q = "您希望篩選哪個欄位？"
    events = _run(_fake_provider(q))
    for kind, data in events:
        if kind == EVT_CLARIFY:
            assert data == q
            break


def test_done_event_for_plain_response():
    events = _run(_fake_provider("已完成排序操作。"))
    kinds = [k for k, _ in events]
    assert EVT_DONE in kinds
    assert EVT_CLARIFY not in kinds


def test_planning_mode_suppresses_clarify():
    """In planning mode, EVT_PLAN_READY should be yielded, not EVT_CLARIFY."""
    events = _run(
        _fake_provider("請問您要哪種格式？"),
        plan_inject="[規劃模式]",
    )
    kinds = [k for k, _ in events]
    assert EVT_PLAN_READY in kinds
    assert EVT_CLARIFY not in kinds


def test_clarify_triggered_in_follow_up_after_tool_calls():
    """
    When tools are called in iteration 1, and the LLM then asks a clarification
    question in iteration 2 (with no tool calls that iteration), EVT_CLARIFY IS
    yielded — the agent pauses so the user can answer the question.

    This is correct UX: if the agent completed some steps and then needs guidance,
    showing the question and waiting is the right behaviour.
    """
    tool_call = ToolCall(
        id="tc1", name="get_sheet_info", arguments={},
    )
    llm_resp = LLMResponse(
        text="",
        tool_calls=[tool_call],
        raw_assistant_message={"role": "assistant", "content": None,
                                "tool_calls": [{"id": "tc1", "type": "function",
                                                "function": {"name": "get_sheet_info",
                                                             "arguments": "{}"}}]},
    )

    call_count = [0]
    def streaming(msgs, tools):
        call_count[0] += 1
        if call_count[0] == 1:
            yield ("tool_calls", llm_resp)
        else:
            # Second iteration: LLM asks a clarification question
            yield ("done", "請問還有其他需要處理的工作表？")

    import tools.executor as executor_mod
    from unittest.mock import patch

    provider = MagicMock(spec=LLMProvider)
    provider.chat_stream.side_effect = streaming

    with patch.object(executor_mod, "execute", return_value='{"status": "ok"}'), \
         patch("backup.get_session_stack", return_value=None):
        events = list(run_turn(
            get_messages=lambda: [{"role": "user", "content": "hello"}],
            tools=[],
            provider=provider,
        ))

    kinds = [k for k, _ in events]
    # In iteration 2, has_tool_call=False and the response is a question → EVT_CLARIFY
    assert EVT_CLARIFY in kinds
    assert EVT_DONE not in kinds


def test_empty_response_not_clarification():
    events = _run(_fake_provider(""))
    kinds = [k for k, _ in events]
    assert EVT_CLARIFY not in kinds
