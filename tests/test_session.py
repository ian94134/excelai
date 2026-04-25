"""
session 模組單元測試。

目的：
- reset_messages 只留 system prompt
- append_message / non_system_count 的計數正確
- messages_to_json 不含 system prompt
- load_messages_from_json 會先 reset 再載入
- maybe_summarize 未達閾值時不動作；達閾值時壓縮為 system + summary + 最近 KEEP_RECENT 輪

因 session 模組依賴 streamlit 的 session_state，本檔所有測試都透過
conftest.py 的 `fake_session_state` fixture 取代為 dict 替身。
"""

from __future__ import annotations
import json
from unittest.mock import MagicMock

import pytest

# 若環境沒有 streamlit，整個檔就跳過（跟 fake_session_state fixture 行為一致）
streamlit = pytest.importorskip("streamlit")

import session
from config import SYSTEM_PROMPT


# ── reset_messages ───────────────────────────────────────────────────────────

def test_reset_messages_keeps_only_system(fake_session_state):
    session.reset_messages()
    msgs = session.get_messages()
    assert len(msgs) == 1
    assert msgs[0]["role"] == "system"
    assert msgs[0]["content"] == SYSTEM_PROMPT


def test_reset_messages_is_idempotent(fake_session_state):
    """重複呼叫 reset 結果一致，不會疊加多個 system。"""
    session.reset_messages()
    session.append_message({"role": "user", "content": "hi"})
    session.reset_messages()
    msgs = session.get_messages()
    assert len(msgs) == 1
    assert msgs[0]["role"] == "system"


# ── append_message / non_system_count ────────────────────────────────────────

def test_append_and_non_system_count(fake_session_state):
    session.reset_messages()
    assert session.non_system_count() == 0

    session.append_message({"role": "user", "content": "A"})
    session.append_message({"role": "assistant", "content": "B"})
    session.append_message({"role": "tool", "content": "{}"})

    assert session.non_system_count() == 3
    # system 訊息不計入
    assert len(session.get_messages()) == 4  # system + 3


# ── messages_to_json ─────────────────────────────────────────────────────────

def test_messages_to_json_excludes_system(fake_session_state):
    session.reset_messages()
    session.append_message({"role": "user", "content": "你好"})
    session.append_message({"role": "assistant", "content": "哈囉"})

    raw = session.messages_to_json()
    data = json.loads(raw)

    assert isinstance(data, list)
    assert len(data) == 2
    assert all(m["role"] != "system" for m in data)
    assert data[0]["content"] == "你好"


def test_messages_to_json_empty_when_only_system(fake_session_state):
    session.reset_messages()
    data = json.loads(session.messages_to_json())
    assert data == []


# ── load_messages_from_json ──────────────────────────────────────────────────

def test_load_messages_from_json_resets_then_loads(fake_session_state):
    session.reset_messages()
    session.append_message({"role": "user", "content": "舊訊息"})

    payload = json.dumps([
        {"role": "user", "content": "新1"},
        {"role": "assistant", "content": "新2"},
    ], ensure_ascii=False)

    count = session.load_messages_from_json(payload)

    assert count == 2
    msgs = session.get_messages()
    # reset 後僅剩 system + 載入的 2 筆
    assert len(msgs) == 3
    assert msgs[0]["role"] == "system"
    assert msgs[1]["content"] == "新1"
    # 舊訊息應被清掉
    assert not any(m.get("content") == "舊訊息" for m in msgs)


def test_load_messages_accepts_bytes(fake_session_state):
    session.reset_messages()
    payload = json.dumps([{"role": "user", "content": "x"}]).encode("utf-8")
    count = session.load_messages_from_json(payload)
    assert count == 1


def test_load_messages_from_json_rejects_non_list(fake_session_state):
    with pytest.raises(ValueError, match="最外層必須是陣列"):
        session.load_messages_from_json(json.dumps({"role": "user", "content": "x"}))


def test_load_messages_from_json_rejects_invalid_item(fake_session_state):
    with pytest.raises(ValueError, match="不是物件"):
        session.load_messages_from_json(json.dumps(["bad-item"]))


def test_load_messages_from_json_ignores_system_from_file(fake_session_state):
    payload = json.dumps([
        {"role": "system", "content": "外部 system"},
        {"role": "user", "content": "u1"},
    ], ensure_ascii=False)
    count = session.load_messages_from_json(payload)
    assert count == 1
    msgs = session.get_messages()
    assert msgs[0]["role"] == "system"
    assert "外部 system" not in msgs[0]["content"]
    assert msgs[1]["role"] == "user"
    assert msgs[1]["content"] == "u1"


# ── maybe_summarize ──────────────────────────────────────────────────────────

def test_maybe_summarize_noop_below_threshold(fake_session_state, fake_provider):
    session.reset_messages()
    # 只加 5 輪，遠低於 CONTEXT_SUMMARIZE_THRESHOLD
    for i in range(5):
        session.append_message({"role": "user", "content": f"u{i}"})
        session.append_message({"role": "assistant", "content": f"a{i}"})

    before = list(session.get_messages())
    session.maybe_summarize(fake_provider)
    after = session.get_messages()

    # 未觸發摘要：訊息應完全不變，且不應呼叫 provider.chat
    assert before == after
    fake_provider.chat.assert_not_called()


def test_maybe_summarize_compresses_over_threshold(fake_session_state):
    """超過閾值時：結果 = system + 摘要 assistant + 最近 KEEP_RECENT 輪。"""
    session.reset_messages()
    # 塞入足以觸發的量（THRESHOLD + 幾輪餘裕）
    total = session.CONTEXT_SUMMARIZE_THRESHOLD + 5
    for i in range(total):
        session.append_message({"role": "user", "content": f"msg-{i}"})

    provider = MagicMock()
    provider.chat.return_value = MagicMock(text="摘要內容 OK")

    session.maybe_summarize(provider)
    msgs = session.get_messages()

    # 結構：system + 摘要（assistant）+ KEEP_RECENT 筆最近訊息
    assert msgs[0]["role"] == "system"
    assert msgs[1]["role"] == "assistant"
    assert "摘要內容 OK" in msgs[1]["content"]
    assert len(msgs) == 2 + session.KEEP_RECENT

    # provider.chat 應被呼叫一次
    provider.chat.assert_called_once()


def test_maybe_summarize_handles_provider_failure(fake_session_state):
    """provider 拋例外時仍能壓縮，摘要文字 fallback。"""
    session.reset_messages()
    total = session.CONTEXT_SUMMARIZE_THRESHOLD + 5
    for i in range(total):
        session.append_message({"role": "user", "content": f"msg-{i}"})

    provider = MagicMock()
    provider.chat.side_effect = RuntimeError("network down")

    session.maybe_summarize(provider)
    msgs = session.get_messages()

    assert msgs[0]["role"] == "system"
    assert msgs[1]["role"] == "assistant"
    assert "摘要失敗" in msgs[1]["content"]
    assert len(msgs) == 2 + session.KEEP_RECENT
