from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Generator


@dataclass
class ToolCall:
    id: str
    name: str
    arguments: dict


@dataclass
class LLMResponse:
    text: str | None
    tool_calls: list[ToolCall] = field(default_factory=list)
    # 原始 assistant 訊息物件（供 messages list 追加用）
    raw_assistant_message: dict | None = None


class LLMProvider(ABC):
    @abstractmethod
    def chat(self, messages: list[dict], tools: list[dict]) -> LLMResponse:
        """送出對話，回傳 LLMResponse。messages 遵循 OpenAI message list 格式。"""
        ...

    @abstractmethod
    def chat_stream(
        self, messages: list[dict], tools: list[dict]
    ) -> Generator[tuple[str, object], None, None]:
        """
        串流對話。
        純文字回覆 → yield ("text", chunk_str) ... yield ("done", full_text)
        tool call  → yield ("tool_calls", LLMResponse)
        """
        ...
