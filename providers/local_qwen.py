import re
import time
import json
from typing import Generator
from openai import OpenAI
from httpx import ConnectError, ReadTimeout
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    RetryCallState,
)
from .base import LLMProvider, LLMResponse, ToolCall
from logger import get_logger

_THINK_RE = re.compile(r"<think>.*?</think>", re.DOTALL)
_log = get_logger("provider.qwen")

# 重試設定：ConnectError / ReadTimeout，指數退避，最多 3 次
_RETRY_EXCEPTIONS = (ConnectError, ReadTimeout)
_RETRY_KWARGS = dict(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=1, max=10),
    retry=retry_if_exception_type(_RETRY_EXCEPTIONS),
    reraise=True,
)


def _make_retry_callback(label: str):
    """回傳 before_sleep callback，用來 log 重試資訊。"""
    def _cb(state: RetryCallState):
        _log.warning(f"{label}_retry", extra={
            "attempt": state.attempt_number,
            "error_type": type(state.outcome.exception()).__name__,
        })
    return _cb


class LocalQwenProvider(LLMProvider):
    def __init__(self, base_url: str, model: str):
        self.client = OpenAI(base_url=base_url, api_key="not-required")
        self.model = model
        _log.info("provider_init", extra={"base_url": base_url, "model": model})

    def chat(self, messages: list[dict], tools: list[dict]) -> LLMResponse:
        """非串流模式（tool call 情境使用）。連線失敗時自動重試最多 3 次。"""
        @retry(**_RETRY_KWARGS, before_sleep=_make_retry_callback("chat"))
        def _call():
            return self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=tools,
                tool_choice="auto",
                extra_body={"chat_template_kwargs": {"enable_thinking": False}},
            )

        resp = _call()
        msg = resp.choices[0].message
        content = _THINK_RE.sub("", msg.content or "").strip() or None

        tool_calls = []
        if msg.tool_calls:
            for tc in msg.tool_calls:
                try:
                    parsed_args = json.loads(tc.function.arguments) if tc.function.arguments else {}
                except json.JSONDecodeError:
                    parsed_args = {}
                    _log.warning("chat_tool_args_json_decode_failed", extra={
                        "tool_name": tc.function.name,
                    })
                tool_calls.append(ToolCall(
                    id=tc.id,
                    name=tc.function.name,
                    arguments=parsed_args,
                ))

        raw = msg.model_dump()
        raw["content"] = content
        return LLMResponse(text=content, tool_calls=tool_calls, raw_assistant_message=raw)

    def chat_stream(
        self, messages: list[dict], tools: list[dict]
    ) -> Generator[tuple[str, object], None, None]:
        """
        串流模式。
        純文字回覆 → yield ("text", chunk_str) ... yield ("done", full_text)
        tool call  → yield ("tool_calls", LLMResponse)
        重試提示   → yield ("retry_info", attempt_number)

        用法：
            for event, data in provider.chat_stream(messages, tools):
                if event == "text":      placeholder.markdown(data + "▌")
                elif event == "done":    placeholder.markdown(data)
                elif event == "tool_calls": handle(data)  # LLMResponse
                elif event == "retry_info": show_retry_warning(data)
        """
        start = time.perf_counter()

        # 記錄最後一次重試 attempt，供連線成功後向 UI 回報
        last_retry_attempt = [0]

        def _retry_cb(state: RetryCallState):
            last_retry_attempt[0] = state.attempt_number
            _log.warning("chat_stream_retry", extra={
                "attempt": state.attempt_number,
                "error_type": type(state.outcome.exception()).__name__,
            })

        @retry(**_RETRY_KWARGS, before_sleep=_retry_cb)
        def _connect():
            return self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=tools,
                tool_choice="auto",
                stream=True,
                extra_body={"chat_template_kwargs": {"enable_thinking": False}},
            )

        try:
            stream = _connect()
        except Exception as e:
            _log.exception("chat_stream_connect_failed", extra={
                "msg_count": len(messages), "error_type": type(e).__name__,
            })
            raise

        # 若曾重試過，先向 UI 發出提示事件（重試成功後才走到這裡）
        if last_retry_attempt[0] > 0:
            yield ("retry_info", last_retry_attempt[0])

        collected_text = ""
        collected_tc: list[dict] = []   # [{id, name, arguments_str}]
        has_tool_calls = False

        for chunk in stream:
            choice = chunk.choices[0]
            delta = choice.delta

            # ── 累積 tool call 分片 ────────────────────────────────────────────
            if delta.tool_calls:
                has_tool_calls = True
                for tc_delta in delta.tool_calls:
                    idx = tc_delta.index
                    while len(collected_tc) <= idx:
                        collected_tc.append({"id": "", "name": "", "args": ""})
                    if tc_delta.id:
                        collected_tc[idx]["id"] = tc_delta.id
                    if tc_delta.function and tc_delta.function.name:
                        collected_tc[idx]["name"] += tc_delta.function.name
                    if tc_delta.function and tc_delta.function.arguments:
                        collected_tc[idx]["args"] += tc_delta.function.arguments

            # ── 純文字分片 ────────────────────────────────────────────────────
            if delta.content:
                collected_text += delta.content
                if not has_tool_calls:
                    # 給 UI 的顯示版：清除已完整的 think 標籤
                    display = _THINK_RE.sub("", collected_text)
                    yield ("text", display)

            if choice.finish_reason:
                break

        duration_ms = int((time.perf_counter() - start) * 1000)

        # ── 回傳最終結果 ──────────────────────────────────────────────────────
        if has_tool_calls:
            tool_calls = []
            raw_tc_list = []
            for tc in collected_tc:
                try:
                    args = json.loads(tc["args"]) if tc["args"] else {}
                except json.JSONDecodeError:
                    args = {}
                tool_calls.append(ToolCall(id=tc["id"], name=tc["name"], arguments=args))
                raw_tc_list.append({
                    "id": tc["id"], "type": "function",
                    "function": {"name": tc["name"], "arguments": tc["args"]},
                })

            raw_msg = {
                "role": "assistant",
                "content": collected_text or None,
                "tool_calls": raw_tc_list,
            }
            _log.info("chat_stream_tool_calls", extra={
                "msg_count": len(messages), "duration_ms": duration_ms,
                "tool_call_count": len(tool_calls),
                "tool_names": [tc.name for tc in tool_calls],
            })
            yield ("tool_calls", LLMResponse(
                text=collected_text or None,
                tool_calls=tool_calls,
                raw_assistant_message=raw_msg,
            ))
        else:
            final = _THINK_RE.sub("", collected_text).strip()
            _log.info("chat_stream_done", extra={
                "msg_count": len(messages), "duration_ms": duration_ms,
                "text_len": len(final),
            })
            yield ("done", final)
