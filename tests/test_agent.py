"""
tests/test_agent.py

Unit tests for agent.run_turn() — the LLM tool-calling loop.

All win32com / Excel / Streamlit calls are mocked so these run on Linux CI.

Key behaviours under test
──────────────────────────
- Pure text response → yields EVT_DONE
- Single tool call → yields EVT_ASST_MSG + EVT_TOOL_START + EVT_TOOL_DONE
- Multi-tool round → all tools executed, session messages correct
- Repeat-loop detection → halts after same call × 3, yields EVT_REPEAT_HALT
- Dangerous tool → halts immediately, yields EVT_DANGEROUS
- Error in tool → yields EVT_TOOL_DONE(has_error=True) + EVT_ROLLBACK
- Planning mode → yields EVT_PLAN_READY, stops before executing tools
- max_iterations guard → stops after limit
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, call

# ── COM stubs BEFORE any project imports ───────────────────────────────────
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _mod in _WIN_STUBS:
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import agent
from agent import (
    EVT_TEXT_CHUNK, EVT_RETRY, EVT_ASST_MSG, EVT_TOOL_START,
    EVT_TOOL_DONE, EVT_ROLLBACK, EVT_DONE, EVT_PLAN_READY,
    EVT_DANGEROUS, EVT_REPEAT_HALT, EVT_ERROR,
    ToolExecution, run_turn,
)
from providers.base import LLMResponse, ToolCall


# ── Helpers ──────────────────────────────────────────────────────────────────

def _msgs():
    return [{"role": "system", "content": "sys"}, {"role": "user", "content": "hi"}]


def _tc(name: str = "write_range", args: dict | None = None) -> ToolCall:
    return ToolCall(id=f"id_{name}", name=name, arguments=args or {"range_addr": "A1"})


def _ok_result(name: str = "write_range") -> str:
    return json.dumps({"status": "ok", "tool": name})


def _err_result(name: str = "write_range") -> str:
    return json.dumps({"error": "boom", "tool": name, "error_type": "TestError"})


def _text_stream(text: str):
    """Fake provider.chat_stream that yields a single text done event."""
    def _inner(msgs, tools):
        yield ("text", text)
        yield ("done", text)
    return _inner


def _tool_stream(tool_calls: list[ToolCall], then_done: str = "完成"):
    """Fake provider.chat_stream that yields tool_calls then done on next call."""
    _call_count = [0]

    raw_msg = {
        "role": "assistant",
        "content": None,
        "tool_calls": [{"id": tc.id, "type": "function",
                        "function": {"name": tc.name, "arguments": json.dumps(tc.arguments)}}
                       for tc in tool_calls],
    }
    resp = LLMResponse(text=None, tool_calls=tool_calls, raw_assistant_message=raw_msg)

    def _inner(msgs, tools):
        _call_count[0] += 1
        if _call_count[0] == 1:
            yield ("tool_calls", resp)
        else:
            yield ("text", then_done)
            yield ("done", then_done)

    return _inner


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture()
def no_stack():
    with patch("agent.backup.get_session_stack", return_value=None):
        yield


@pytest.fixture()
def mock_execute():
    with patch("agent.execute") as m:
        yield m


@pytest.fixture()
def mock_undo():
    with patch("agent.et.undo_last", return_value={"status": "ok"}) as m:
        yield m


def _make_provider(stream_fn):
    p = MagicMock()
    p.chat_stream.side_effect = stream_fn
    return p


# ── Text-only responses ───────────────────────────────────────────────────────

def test_text_response_yields_done(no_stack):
    provider = _make_provider(_text_stream("Hello!"))
    events = list(run_turn(_msgs, [], provider))
    kinds = [k for k, _ in events]
    assert EVT_DONE in kinds
    done_data = next(d for k, d in events if k == EVT_DONE)
    assert done_data == "Hello!"


def test_text_response_yields_chunks(no_stack):
    provider = _make_provider(_text_stream("Hello!"))
    events = list(run_turn(_msgs, [], provider))
    assert any(k == EVT_TEXT_CHUNK for k, _ in events)


def test_text_response_no_tool_events(no_stack):
    provider = _make_provider(_text_stream("Hi"))
    events = list(run_turn(_msgs, [], provider))
    for k, _ in events:
        assert k not in (EVT_TOOL_DONE, EVT_ASST_MSG, EVT_DANGEROUS)


# ── Single tool call ──────────────────────────────────────────────────────────

def test_single_tool_yields_asst_msg(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc()
    provider = _make_provider(_tool_stream([tc]))
    events = list(run_turn(_msgs, [], provider))
    assert any(k == EVT_ASST_MSG for k, _ in events)


def test_single_tool_yields_tool_start(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc()
    provider = _make_provider(_tool_stream([tc]))
    events = list(run_turn(_msgs, [], provider))
    assert any(k == EVT_TOOL_START for k, _ in events)


def test_single_tool_yields_tool_done(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc()
    provider = _make_provider(_tool_stream([tc]))
    events = list(run_turn(_msgs, [], provider))
    tool_dones = [(k, d) for k, d in events if k == EVT_TOOL_DONE]
    assert len(tool_dones) == 1
    tex: ToolExecution = tool_dones[0][1]
    assert tex.tc.name == "write_range"
    assert tex.has_error is False


def test_single_tool_result_json(no_stack, mock_execute):
    mock_execute.return_value = _ok_result("write_range")
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    assert json.loads(tex.result_json)["status"] == "ok"


def test_tool_execute_called_with_correct_args(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc("format_range", {"range_addr": "A1:B2", "bold": True})
    provider = _make_provider(_tool_stream([tc]))
    list(run_turn(_msgs, [], provider))
    mock_execute.assert_called_once_with("format_range", {"range_addr": "A1:B2", "bold": True})


# ── Multi-tool round ──────────────────────────────────────────────────────────

def test_multi_tool_all_executed(no_stack, mock_execute):
    mock_execute.side_effect = [_ok_result("a"), _ok_result("b"), _ok_result("c")]
    tcs = [_tc("a"), _tc("b"), _tc("c")]
    provider = _make_provider(_tool_stream(tcs))
    events = list(run_turn(_msgs, [], provider))
    dones = [d for k, d in events if k == EVT_TOOL_DONE]
    assert len(dones) == 3
    assert [d.tc.name for d in dones] == ["a", "b", "c"]


# ── Dangerous tool ────────────────────────────────────────────────────────────

def test_dangerous_tool_yields_dangerous_halt(no_stack, mock_execute):
    tc = _tc("delete_row", {"index": 5})

    def _stream(msgs, tools):
        raw = {"role": "assistant", "content": None,
               "tool_calls": [{"id": tc.id, "type": "function",
                                "function": {"name": tc.name,
                                             "arguments": json.dumps(tc.arguments)}}]}
        yield ("tool_calls", LLMResponse(text=None, tool_calls=[tc], raw_assistant_message=raw))

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider, dangerous_tools={"delete_row"}))
    assert any(k == EVT_DANGEROUS for k, _ in events)


def test_dangerous_tool_stops_execution(no_stack, mock_execute):
    tc = _tc("delete_row")

    def _stream(msgs, tools):
        raw = {"role": "assistant", "content": None, "tool_calls": []}
        yield ("tool_calls", LLMResponse(text=None, tool_calls=[tc], raw_assistant_message=raw))

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider, dangerous_tools={"delete_row"}))
    # execute() must never be called for a dangerous tool
    mock_execute.assert_not_called()


# ── Repeat-loop detection ─────────────────────────────────────────────────────

def test_repeat_halt_after_3_identical_calls(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc("format_range", {"range_addr": "A1"})
    call_count = [0]

    def _stream(msgs, tools):
        raw = {"role": "assistant", "content": None,
               "tool_calls": [{"id": tc.id, "type": "function",
                                "function": {"name": tc.name,
                                             "arguments": json.dumps(tc.arguments)}}]}
        call_count[0] += 1
        yield ("tool_calls", LLMResponse(text=None, tool_calls=[tc], raw_assistant_message=raw))

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider, dangerous_tools=set()))
    assert any(k == EVT_REPEAT_HALT for k, _ in events)


def test_repeat_halt_message_content(no_stack, mock_execute):
    mock_execute.return_value = _ok_result()
    tc = _tc("format_range", {"range_addr": "A1"})

    def _stream(msgs, tools):
        raw = {"role": "assistant", "content": None,
               "tool_calls": [{"id": tc.id, "type": "function",
                                "function": {"name": tc.name,
                                             "arguments": json.dumps(tc.arguments)}}]}
        yield ("tool_calls", LLMResponse(text=None, tool_calls=[tc], raw_assistant_message=raw))

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider, dangerous_tools=set()))
    halt_msgs = [d for k, d in events if k == EVT_REPEAT_HALT]
    assert halt_msgs and "重複" in halt_msgs[0]


# ── Error and rollback ────────────────────────────────────────────────────────

def test_error_tool_done_has_error_true(no_stack, mock_execute):
    mock_execute.return_value = _err_result()
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    assert tex.has_error is True


def test_error_no_rollback_when_nothing_pushed(no_stack, mock_execute, mock_undo):
    """First tool fails → nothing to roll back → no EVT_ROLLBACK."""
    mock_execute.return_value = _err_result()
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    assert not any(k == EVT_ROLLBACK for k, _ in events)
    mock_undo.assert_not_called()


def test_error_rollback_prior_pushed_steps():
    """Second tool fails → first (backed-up) step should be rolled back."""
    # Set up a fake stack that reports one entry after first execute
    fake_stack = MagicMock()
    depth_vals = [0, 1, 1, 0]  # before_a, after_a, before_b, after_b(fail)
    call_i = [0]

    def _len():
        v = depth_vals[min(call_i[0], len(depth_vals) - 1)]
        call_i[0] += 1
        return v

    fake_stack.__len__ = MagicMock(side_effect=_len)
    fake_stack.peek.return_value = MagicMock()

    exec_side_effects = [_ok_result("a"), _err_result("b")]

    with patch("agent.backup.get_session_stack", return_value=fake_stack), \
         patch("agent.execute", side_effect=exec_side_effects), \
         patch("agent.et.undo_last", return_value={"status": "ok"}) as mock_undo:

        tcs = [_tc("a"), _tc("b")]
        provider = _make_provider(_tool_stream(tcs))
        events = list(run_turn(_msgs, [], provider))

    rollback_events = [d for k, d in events if k == EVT_ROLLBACK]
    assert rollback_events, "Expected at least one EVT_ROLLBACK"
    mock_undo.assert_called()


# ── Planning mode ─────────────────────────────────────────────────────────────

def test_plan_mode_yields_plan_ready(no_stack):
    provider = _make_provider(_text_stream("1. 步驟一\n2. 步驟二"))
    events = list(run_turn(_msgs, [], provider, plan_inject="[規劃模式]"))
    assert any(k == EVT_PLAN_READY for k, _ in events)


def test_plan_mode_no_tool_calls(no_stack, mock_execute):
    provider = _make_provider(_text_stream("計劃：…"))
    list(run_turn(_msgs, [], provider, plan_inject="[規劃模式]"))
    mock_execute.assert_not_called()


def test_plan_mode_stops_after_plan(no_stack):
    """Generator returns after EVT_PLAN_READY — no EVT_DONE."""
    provider = _make_provider(_text_stream("計劃"))
    events = list(run_turn(_msgs, [], provider, plan_inject="[規劃模式]"))
    assert not any(k == EVT_DONE for k, _ in events)


# ── max_iterations guard ──────────────────────────────────────────────────────

def test_max_iterations_respected(no_stack, mock_execute):
    """With max_iterations=2 and infinite tool loop, must stop after 2 LLM calls."""
    mock_execute.return_value = _ok_result()
    tc = _tc("format_range", {"range_addr": "B2"})  # different each time

    call_n = [0]

    def _stream(msgs, tools):
        call_n[0] += 1
        # Always return a different tool so repeat-halt never triggers
        tc_n = _tc(f"tool_{call_n[0]}", {"n": call_n[0]})
        raw = {"role": "assistant", "content": None,
               "tool_calls": [{"id": tc_n.id, "type": "function",
                                "function": {"name": tc_n.name,
                                             "arguments": json.dumps(tc_n.arguments)}}]}
        yield ("tool_calls", LLMResponse(text=None, tool_calls=[tc_n], raw_assistant_message=raw))

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider, max_iterations=2, dangerous_tools=set()))
    assert call_n[0] == 2


# ── Retry info propagation ────────────────────────────────────────────────────

def test_retry_info_propagated(no_stack):
    def _stream(msgs, tools):
        yield ("retry_info", 2)
        yield ("text", "ok")
        yield ("done", "ok")

    provider = _make_provider(_stream)
    events = list(run_turn(_msgs, [], provider))
    retry_events = [d for k, d in events if k == EVT_RETRY]
    assert retry_events == [2]


# ── Workbook context injection ────────────────────────────────────────────────

def test_wb_context_injected_into_system_message(no_stack):
    """wb_context_fn result should appear in messages passed to provider."""
    received_msgs = []

    def _stream(msgs, tools):
        received_msgs.extend(msgs)
        yield ("text", "hi")
        yield ("done", "hi")

    provider = _make_provider(_stream)
    list(run_turn(
        _msgs, [], provider,
        wb_context_fn=lambda: "WORKBOOK_CTX",
    ))
    assert received_msgs
    system_msg = next((m for m in received_msgs if m.get("role") == "system"), None)
    assert system_msg is not None
    assert "WORKBOOK_CTX" in system_msg["content"]


# ── Error hint enrichment ─────────────────────────────────────────────────────

def test_error_result_gets_hint(no_stack, mock_execute):
    """When a tool returns an error with known error_type, result_json must contain 'hint'."""
    err = json.dumps({"error": "not found", "error_type": "SheetNotFoundError"})
    mock_execute.return_value = err
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    payload = json.loads(tex.result_json)
    assert "hint" in payload


def test_error_hint_contains_suggested_next(no_stack, mock_execute):
    """SheetNotFoundError hint should include suggested_next = get_sheet_info."""
    err = json.dumps({"error": "no sheet", "error_type": "SheetNotFoundError"})
    mock_execute.return_value = err
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    payload = json.loads(tex.result_json)
    assert payload.get("suggested_next") == "get_sheet_info"


def test_unknown_error_type_gets_default_hint(no_stack, mock_execute):
    """Unknown error_type should still receive the default hint string."""
    err = json.dumps({"error": "something", "error_type": "WeirdCustomError"})
    mock_execute.return_value = err
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    payload = json.loads(tex.result_json)
    assert "hint" in payload


def test_ok_result_not_enriched(no_stack, mock_execute):
    """Successful tool results must NOT have a 'hint' key added."""
    mock_execute.return_value = _ok_result()
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    payload = json.loads(tex.result_json)
    assert "hint" not in payload


def test_enrich_preserves_original_error_field(no_stack, mock_execute):
    """The original 'error' field must survive enrichment unchanged."""
    err = json.dumps({"error": "original message", "error_type": "RangeError"})
    mock_execute.return_value = err
    provider = _make_provider(_tool_stream([_tc()]))
    events = list(run_turn(_msgs, [], provider))
    tex = next(d for k, d in events if k == EVT_TOOL_DONE)
    payload = json.loads(tex.result_json)
    assert payload["error"] == "original message"
