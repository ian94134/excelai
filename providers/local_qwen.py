import re
import json
from typing import Generator
from openai import OpenAI
from .base import LLMProvider, LLMResponse, ToolCall

_THINK_RE = re.compile(r"<think>.*?</think>", re.DOTALL)


class LocalQwenProvider(LLMProvider):
    def __init__(self, base_url: str, model: str):
        self.client = OpenAI(base_url=base_url, api_key="not-required")
        self.model = model

    def chat(self, messages: list[dict], tools: list[dict]) -> LLMResponse:
        """非串流模式（tool call 情境使用）"""
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            tools=tools,
            tool_choice="auto",
            extra_body={"chat_template_kwargs": {"enable_thinking": False}},
        )
        msg = resp.choices[0].message
        content = _THINK_RE.sub("", msg.content or "").strip() or None

        tool_calls = []
        if msg.tool_calls:
            for tc in msg.tool_calls:
                tool_calls.append(ToolCall(
                    id=tc.id,
                    name=tc.function.name,
                    arguments=json.loads(tc.function.arguments),
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

        用法：
            for event, data in provider.chat_stream(messages, tools):
                if event == "text":   placeholder.markdown(data + "▌")
                elif event == "done": placeholder.markdown(data)
                elif event == "tool_calls": handle(data)  # LLMResponse
        """
        stream = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            tools=tools,
            tool_choice="auto",
            stream=True,
            extra_body={"chat_template_kwargs": {"enable_thinking": False}},
        )

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
                collected_text += delta.content   # 原始累積（含可能跨片的 <think> 標籤）
                if not has_tool_calls:
                    # 給 UI 的顯示版：清除已完整的 think 標籤（跨片殘留無害，最終全文會再清一次）
                    display = _THINK_RE.sub("", collected_text)
                    yield ("text", display)

            if choice.finish_reason:
                break

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
            yield ("tool_calls", LLMResponse(
                text=collected_text or None,
                tool_calls=tool_calls,
                raw_assistant_message=raw_msg,
            ))
        else:
            final = _THINK_RE.sub("", collected_text).strip()
            yield ("done", final)
