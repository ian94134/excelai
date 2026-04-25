"""
tests/test_providers.py
LocalQwenProvider unit tests.
Uses unittest.mock.patch to stub the openai.OpenAI client.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# Helpers: build fake OpenAI response objects
# ---------------------------------------------------------------------------

def _make_message(content=None, tool_calls=None):
    msg = MagicMock()
    msg.content = content
    msg.tool_calls = tool_calls or []
    msg.model_dump.return_value = {
        "role": "assistant", "content": content, "tool_calls": []
    }
    return msg


def _make_completion(content=None, tool_calls=None):
    msg = _make_message(content, tool_calls)
    choice = MagicMock()
    choice.message = msg
    resp = MagicMock()
    resp.choices = [choice]
    return resp


def _make_tc_obj(tc_id, name, arguments_str):
    tc = MagicMock()
    tc.id = tc_id
    tc.function = MagicMock()
    tc.function.name = name
    tc.function.arguments = arguments_str
    return tc


def _make_stream_chunks(text_parts=None, tool_calls_parts=None):
    """Return an iterator of fake stream chunks."""
    chunks = []
    if text_parts:
        for i, part in enumerate(text_parts):
            delta = MagicMock()
            delta.content = part
            delta.tool_calls = None
            choice = MagicMock()
            choice.delta = delta
            choice.finish_reason = "stop" if i == len(text_parts) - 1 else None
            chunk = MagicMock()
            chunk.choices = [choice]
            chunks.append(chunk)
    if tool_calls_parts:
        for i, (idx, tc_id, name, args_frag) in enumerate(tool_calls_parts):
            tc_delta = MagicMock()
            tc_delta.index = idx
            tc_delta.id = tc_id
            tc_delta.function = MagicMock()
            tc_delta.function.name = name
            tc_delta.function.arguments = args_frag
            delta = MagicMock()
            delta.content = None
            delta.tool_calls = [tc_delta]
            choice = MagicMock()
            choice.delta = delta
            choice.finish_reason = (
                "tool_calls" if i == len(tool_calls_parts) - 1 else None
            )
            chunk = MagicMock()
            chunk.choices = [choice]
            chunks.append(chunk)
    return iter(chunks)


# ---------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------

@pytest.fixture
def provider():
    """
    Return (LocalQwenProvider_instance, mock_create).
    The underlying openai.OpenAI client is replaced with a MagicMock so no
    real HTTP requests are made.
    """
    mock_client = MagicMock()
    mock_create = MagicMock()
    mock_client.chat.completions.create = mock_create

    with patch("providers.local_qwen.OpenAI", return_value=mock_client):
        from providers.local_qwen import LocalQwenProvider
        p = LocalQwenProvider(base_url="http://fake", model="qwen-test")

    return p, mock_create


# ---------------------------------------------------------------------------
# chat() tests
# ---------------------------------------------------------------------------

class TestChat:
    def test_plain_text_response(self, provider):
        """chat() with a plain-text reply sets LLMResponse.text correctly."""
        p, mock_create = provider
        mock_create.return_value = _make_completion(content="Hello!")

        result = p.chat([{"role": "user", "content": "hi"}], [])

        assert result.text == "Hello!"
        assert result.tool_calls == []

    def test_think_tag_stripped(self, provider):
        """<think>...</think> blocks must be removed from text."""
        p, mock_create = provider
        mock_create.return_value = _make_completion(
            content="<think>internal reasoning</think>Final answer"
        )

        result = p.chat([], [])

        assert result.text == "Final answer"
        assert "<think>" not in (result.text or "")

    def test_tool_call_parsed(self, provider):
        """ToolCall fields (id, name, arguments) are correctly parsed."""
        p, mock_create = provider
        tc_obj = _make_tc_obj(
            "call_1", "write_range",
            '{"range_addr": "A1", "data": [[1]]}'
        )
        mock_create.return_value = _make_completion(
            content=None, tool_calls=[tc_obj]
        )

        result = p.chat([], [])

        assert len(result.tool_calls) == 1
        tc = result.tool_calls[0]
        assert tc.id == "call_1"
        assert tc.name == "write_range"
        assert tc.arguments == {"range_addr": "A1", "data": [[1]]}

    def test_tool_call_json_decode_fallback(self, provider):
        """Malformed tool_call JSON falls back to {} without raising."""
        p, mock_create = provider
        tc_obj = _make_tc_obj("call_bad", "write_range", "{INVALID JSON}")
        mock_create.return_value = _make_completion(
            content=None, tool_calls=[tc_obj]
        )

        result = p.chat([], [])

        assert len(result.tool_calls) == 1
        assert result.tool_calls[0].arguments == {}

    def test_empty_content_becomes_none(self, provider):
        """Empty or whitespace-only content is normalised to None."""
        p, mock_create = provider
        mock_create.return_value = _make_completion(content="")

        result = p.chat([], [])

        assert result.text is None


# ---------------------------------------------------------------------------
# chat_stream() tests
# ---------------------------------------------------------------------------

class TestChatStream:
    def test_text_stream_yields_done(self, provider):
        """Plain-text stream ends with ("done", full_text), think tags removed."""
        p, mock_create = provider
        mock_create.return_value = _make_stream_chunks(
            text_parts=["<think>skip</think>Hello", " world"]
        )

        events = list(p.chat_stream([], []))
        done_events = [d for e, d in events if e == "done"]

        assert len(done_events) == 1
        assert done_events[0] == "Hello world"

    def test_text_stream_yields_text_events(self, provider):
        """Streaming text yields at least one ("text", ...) event for the UI."""
        p, mock_create = provider
        mock_create.return_value = _make_stream_chunks(
            text_parts=["Part1", " Part2", " Part3"]
        )

        events = list(p.chat_stream([], []))
        text_events = [e for e, _ in events if e == "text"]

        assert len(text_events) >= 1

    def test_tool_call_stream_yields_tool_calls(self, provider):
        """Tool-call stream ends with ("tool_calls", LLMResponse)."""
        p, mock_create = provider
        mock_create.return_value = _make_stream_chunks(
            tool_calls_parts=[(0, "call_x", "get_sheet_info", "")]
        )

        events = list(p.chat_stream([], []))
        tc_events = [d for e, d in events if e == "tool_calls"]

        assert len(tc_events) == 1
        assert tc_events[0].tool_calls[0].name == "get_sheet_info"

    def test_tool_call_args_json_decode_fallback_stream(self, provider):
        """Malformed tool_call JSON in stream falls back to {}."""
        p, mock_create = provider
        mock_create.return_value = _make_stream_chunks(
            tool_calls_parts=[(0, "call_bad", "write_range", "{BAD")]
        )

        events = list(p.chat_stream([], []))
        tc_events = [d for e, d in events if e == "tool_calls"]

        assert len(tc_events) == 1
        assert tc_events[0].tool_calls[0].arguments == {}

    def test_connect_error_propagates(self, provider):
        """
        ConnectError propagates to the caller after retries are exhausted.
        (In real usage tenacity retries 3x; here the decorator is real,
        so we let it retry and eventually re-raise.)
        """
        from httpx import ConnectError
        p, mock_create = provider
        mock_create.side_effect = ConnectError("refused", request=MagicMock())

        with pytest.raises(ConnectError):
            list(p.chat_stream([], []))
