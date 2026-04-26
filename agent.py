"""
agent.py — LLM tool-calling loop (v4.5.0).

Extracted from main.py so the loop logic is:
  - Testable without Streamlit
  - Reusable across different UIs
  - A clean boundary between logic and presentation

run_turn() is a generator that drives one "user message → response" cycle,
including multi-step tool calling, planning mode, repeat-loop detection,
auto-rollback on error, and structured error-hint enrichment.
It has zero Streamlit imports.

Event protocol
──────────────
Yielded tuples: (event_kind: str, data: Any)

  EVT_TEXT_CHUNK   │ str            │ Streaming text fragment from LLM
  EVT_RETRY        │ int            │ Qwen Tenacity retry attempt number
  EVT_ASST_MSG     │ dict           │ Raw assistant message — caller must append to session
  EVT_TOOL_START   │ ToolCall       │ Tool about to execute (show spinner / status label)
  EVT_TOOL_DONE    │ ToolExecution  │ Tool executed — caller must append tool result to session
  EVT_ROLLBACK     │ int            │ N prior successful steps rolled back after an error
  EVT_DONE         │ str            │ Final assistant text; loop ends after this event
  EVT_PLAN_READY   │ str            │ Plan text (planning mode); loop ends, caller shows confirm
  EVT_DANGEROUS    │ ToolCall       │ Dangerous tool detected; loop ends, caller shows confirm
  EVT_REPEAT_HALT  │ str            │ Same tool called ≥ 3x; loop ends with warning message
  EVT_ERROR        │ Exception      │ Unhandled exception; loop ends
  EVT_CLARIFY      │ str            │ LLM is asking a clarification question; loop ends, caller shows question

Caller responsibilities per event
──────────────────────────────────
  EVT_ASST_MSG    → session.append_message(data)
  EVT_TOOL_DONE   → session.append_message({
                        "role": "tool", "tool_call_id": data.tc.id,
                        "name": data.tc.name, "content": data.result_json
                    })
  EVT_DONE        → session.append_message({"role": "assistant", "content": data})
  EVT_PLAN_READY  → save plan; show confirm UI; do NOT append to session yet
  EVT_DANGEROUS   → save tool call; show confirm dialog; do NOT continue generator
  EVT_REPEAT_HALT → session.append_message({"role": "assistant", "content": data})
  EVT_CLARIFY     → session.append_message({"role": "assistant", "content": data}) (same as EVT_DONE)
  EVT_ROLLBACK    → show rollback warning to user (Excel state already restored)
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any, Callable, Iterator, Set

import backup
import excel_tools as et
from tools.executor import execute, DANGEROUS_TOOLS
from providers.base import LLMProvider, LLMResponse, ToolCall


# ── Structured error hints ──────────────────────────────────────────────────
# Maps error_type → {"hint": str, "suggested_next": str | None}
# Appended to error result JSON so the LLM receives actionable guidance.

_ERROR_HINTS: dict[str, dict] = {
    "SheetNotFoundError": {
        "hint": (
            "指定的工作表不存在。請先呼叫 get_sheet_info 取得正確工作表名稱清單，"
            "名稱須完全符合（含大小寫與空白）。"
        ),
        "suggested_next": "get_sheet_info",
    },
    "WorkbookNotFoundError": {
        "hint": "找不到指定的活頁簿，請確認 Excel 已開啟且 file_name 參數正確。",
        "suggested_next": "get_sheet_info",
    },
    "RangeError": {
        "hint": (
            "範圍位址格式有誤或超出工作表已使用範圍。"
            "請先呼叫 get_used_range 確認有效範圍後重新指定，格式如 'A1:D10'。"
        ),
        "suggested_next": "get_used_range",
    },
    "ProtectedSheetError": {
        "hint": (
            "工作表受到保護，無法修改。"
            "請先呼叫 unprotect_sheet 解除保護後再執行操作。"
        ),
        "suggested_next": "unprotect_sheet",
    },
    "FormulaError": {
        "hint": (
            "公式語法錯誤。請確認：(1) 使用英文函數名稱（即使中文版 Excel 亦同）；"
            "(2) 公式以 = 開頭；(3) 引用的儲存格/範圍確實存在；"
            "(4) 括號與逗號格式正確。"
        ),
        "suggested_next": None,
    },
    "ValueError": {
        "hint": "參數值不合法，請確認資料型別（數字/字串/布林）與允許的範圍是否正確。",
        "suggested_next": None,
    },
    "IndexError": {
        "hint": (
            "欄/列索引超出範圍，索引從 1 開始。"
            "請先呼叫 get_used_range 確認工作表的實際大小。"
        ),
        "suggested_next": "get_used_range",
    },
    "ChartError": {
        "hint": (
            "圖表操作失敗。請確認：(1) data_range 有數值資料；"
            "(2) chart_type 為支援的類型（'bar'/'line'/'pie'/'scatter' 等）；"
            "(3) 指定的工作表存在。"
        ),
        "suggested_next": "get_used_range",
    },
    "PivotError": {
        "hint": (
            "樞紐分析表建立失敗。請確認來源範圍包含標題列且資料不為空，"
            "dest_sheet 若不存在請先呼叫 add_sheet 建立。"
        ),
        "suggested_next": "get_sheet_info",
    },
    "UnknownTool": {
        "hint": "呼叫了不存在的工具名稱，請從可用工具清單中選擇正確名稱。",
        "suggested_next": None,
    },
    "UnexpectedError": {
        "hint": (
            "發生未預期的錯誤。請先呼叫 get_sheet_info 確認工作表狀態，"
            "再以更簡單的參數重試，或拆分成更小的步驟執行。"
        ),
        "suggested_next": "get_sheet_info",
    },
}

_DEFAULT_HINT = (
    "請確認參數是否正確，或先呼叫 get_sheet_info / get_used_range "
    "取得當前工作表狀態後再重試。"
)


def _enrich_error_result(tool_name: str, result_json: str) -> str:
    """
    Append structured hint and suggested_next to an error result JSON.

    The enriched JSON is returned to the LLM as the tool result, giving it
    actionable guidance on how to recover without the user needing to intervene.
    Original fields are preserved; non-error results are returned unchanged.
    """
    try:
        payload = json.loads(result_json)
    except Exception:
        return result_json

    if "error" not in payload and payload.get("status") == "error":
        payload["error"] = payload.get("message", "工具回傳錯誤狀態")

    if "error" not in payload:
        return result_json

    err_type  = payload.get("error_type", "")
    hint_info = _ERROR_HINTS.get(err_type, {})
    hint      = hint_info.get("hint") or _DEFAULT_HINT
    suggested = hint_info.get("suggested_next")

    payload["hint"] = hint
    if suggested:
        payload["suggested_next"] = suggested

    try:
        return json.dumps(payload, ensure_ascii=False)
    except Exception:
        return result_json


# ── Event kind constants ────────────────────────────────────────────────────

EVT_TEXT_CHUNK  = "text_chunk"
EVT_RETRY       = "retry_info"
EVT_ASST_MSG    = "asst_msg"
EVT_TOOL_START  = "tool_start"
EVT_TOOL_DONE   = "tool_done"
EVT_ROLLBACK    = "rollback"
EVT_DONE        = "done"
EVT_PLAN_READY  = "plan_ready"
EVT_DANGEROUS   = "dangerous_halt"
EVT_REPEAT_HALT = "repeat_halt"
EVT_ERROR       = "error"
EVT_CLARIFY     = "clarify"


# ── Clarification detection (v4.6.0) ────────────────────────────────────────
# Triggered when the LLM asks a question instead of immediately acting,
# signalling it needs more information before proceeding.
# Pattern: contains a question-seeking keyword + ends with ? or ？
_CLARIFY_RE = re.compile(
    r"(請問|您要|是否要|請確認|還是|您希望|需要我|您想要|"
    r"哪[個份張種]|哪一[個份張]|指的是|請說明|請提供|請告知|"
    r"可以告訴我|能告訴我|需要更多|能確認)[^？?]*[？?]",
    re.UNICODE,
)


def _is_clarification(text: str) -> bool:
    """
    Return True if the LLM response is a clarification question rather than
    a direct action or final answer.

    Conditions (both must hold):
    - Contains a clarification-seeking pattern (see _CLARIFY_RE)
    - Contains at least one question mark (？ or ?)

    This prevents false-positives on statements that merely contain keywords
    but are not genuine questions.
    """
    has_question_mark = "？" in text or "?" in text
    return has_question_mark and bool(_CLARIFY_RE.search(text))


# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class ToolExecution:
    """Result of a single tool call execution."""
    tc: ToolCall
    result_json: str
    backup_entry: Any   # BackupEntry | None — None for read-only tools
    has_error: bool


# ── Main entry point ─────────────────────────────────────────────────────────

def run_turn(
    get_messages: Callable[[], list[dict]],
    tools: list[dict],
    provider: LLMProvider,
    *,
    # Behaviour settings
    dangerous_tools: Set[str] | None = None,
    max_iterations: int = 100,
    plan_inject: str = "",
    # Context injection (optional)
    wb_context_fn: Callable[[], str] | None = None,
    inject_ctx_fn: Callable[[list[dict], str], list[dict]] | None = None,
) -> Iterator[tuple[str, Any]]:
    """
    Agent tool-calling loop.

    Parameters
    ----------
    get_messages:
        Called at the start of each outer iteration to read the current
        session messages (including tool results from previous rounds).
    tools:
        OpenAI-format tool schemas passed to the LLM.
    provider:
        A LLMProvider instance (e.g. LocalQwenProvider).
    dangerous_tools:
        Tool names that require user confirmation before executing.
        Defaults to executor.DANGEROUS_TOOLS.
    max_iterations:
        Maximum number of outer loop iterations before halting.
    plan_inject:
        Extra text appended to the last user message to trigger planning mode.
        When non-empty, tools are disabled so the LLM generates a plan first.
    wb_context_fn:
        Optional callable returning a workbook summary string for the system message.
    inject_ctx_fn:
        Optional callable (msgs, ctx) -> msgs for custom context injection.
    """
    if dangerous_tools is None:
        dangerous_tools = DANGEROUS_TOOLS

    last_sig: tuple[str, str] | None = None
    same_count = 0

    for _iter in range(max_iterations):
        has_tool_call = False
        _round_pushed = 0   # backup entries successfully pushed this iteration

        try:
            # ── Build message list for this iteration ──────────────────────
            msgs = list(get_messages())

            # Inject live workbook context into system message
            if wb_context_fn:
                wb_ctx = wb_context_fn()
                if wb_ctx:
                    if inject_ctx_fn:
                        msgs = inject_ctx_fn(msgs, wb_ctx)
                    elif msgs and msgs[0].get("role") == "system":
                        msgs[0] = dict(msgs[0])
                        msgs[0]["content"] = msgs[0]["content"] + "\n\n" + wb_ctx
                    else:
                        msgs.insert(0, {"role": "system", "content": wb_ctx})

            # Plan inject: append to last user message, disable tool use
            use_tools = tools
            if plan_inject:
                use_tools = []
                for i in range(len(msgs) - 1, -1, -1):
                    if msgs[i].get("role") == "user":
                        msgs[i] = dict(msgs[i])
                        msgs[i]["content"] = msgs[i]["content"] + plan_inject
                        break

            # ── LLM streaming call ─────────────────────────────────────────
            for event, data in provider.chat_stream(msgs, use_tools):

                if event == "text":
                    yield (EVT_TEXT_CHUNK, data)

                elif event == "retry_info":
                    yield (EVT_RETRY, int(data))

                elif event == "done":
                    final = str(data) if data else ""
                    if plan_inject:
                        yield (EVT_PLAN_READY, final)
                    elif not has_tool_call and _is_clarification(final):
                        # LLM is asking for more info before acting — pause the loop
                        yield (EVT_CLARIFY, final)
                    else:
                        yield (EVT_DONE, final)
                    return

                elif event == "tool_calls":
                    has_tool_call = True
                    resp: LLMResponse = data

                    # Signal caller to append the assistant message immediately
                    yield (EVT_ASST_MSG, resp.raw_assistant_message or {
                        "role": "assistant", "content": None,
                    })

                    for tc in resp.tool_calls:

                        # ── Dangerous tool → halt for user confirmation ────
                        if tc.name in dangerous_tools:
                            yield (EVT_DANGEROUS, tc)
                            return

                        # ── Repeat-loop detection ──────────────────────────
                        sig = (
                            tc.name,
                            json.dumps(tc.arguments, ensure_ascii=False, sort_keys=True),
                        )
                        if sig == last_sig:
                            same_count += 1
                        else:
                            last_sig = sig
                            same_count = 1
                        if same_count >= 3:
                            halt_msg = (
                                "⚠️ 偵測到 AI 連續重複呼叫相同工具參數，已停止本輪避免卡住。"
                                "請補充更明確的格式需求後再試"
                                "（例如：標題列藍底白字、資料列加細框線、欄寬自動調整）。"
                            )
                            yield (EVT_REPEAT_HALT, halt_msg)
                            return

                        # ── Execute tool ───────────────────────────────────
                        stk = backup.get_session_stack()
                        stk_before = len(stk) if stk else 0

                        yield (EVT_TOOL_START, tc)  # UI: show "starting…" label

                        result_json = execute(tc.name, tc.arguments)

                        # Identify newly pushed backup entry (if any)
                        stk_after = backup.get_session_stack()
                        entry = None
                        if stk_after and len(stk_after) > stk_before:
                            entry = stk_after.peek()
                            _round_pushed += 1

                        # Parse result to detect errors
                        try:
                            _parsed = json.loads(result_json)
                        except Exception:
                            _parsed = {}
                        has_err = "error" in _parsed or _parsed.get("status") == "error"

                        # Enrich error result with actionable hint for the LLM
                        if has_err:
                            result_json = _enrich_error_result(tc.name, result_json)

                        yield (EVT_TOOL_DONE, ToolExecution(
                            tc=tc,
                            result_json=result_json,
                            backup_entry=entry,
                            has_error=has_err,
                        ))

                        # ── Auto-rollback on error ─────────────────────────
                        # Roll back all successfully backed-up steps from this
                        # iteration so Excel stays consistent after a mid-batch failure.
                        if has_err and _round_pushed > 0:
                            rolled = 0
                            for _ in range(_round_pushed):
                                try:
                                    et.undo_last()
                                    rolled += 1
                                except Exception:
                                    break
                            if rolled:
                                yield (EVT_ROLLBACK, rolled)
                            _round_pushed = 0

        except Exception as exc:
            yield (EVT_ERROR, exc)
            return

        # No tool calls this iteration → LLM is done
        if not has_tool_call:
            break
