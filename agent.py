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

_REPEAT_SUMMARY_TOOLS = {
    "read_range",
    "get_used_range",
    "get_sheet_info",
    "get_workbook_summary",
    "query_range",
    "summarize_range",
    "find_duplicates",
    "list_workbooks",
}


def _markdown_table(headers: list[Any], rows: list[list[Any]], max_rows: int = 8) -> str:
    safe_headers = [str(h) if h is not None else "" for h in headers]
    safe_rows = rows[:max_rows]
    if not safe_headers and safe_rows:
        safe_headers = [f"欄{i + 1}" for i in range(max(len(r) for r in safe_rows))]
    if not safe_headers:
        return ""

    def _cell(value: Any) -> str:
        text = "" if value is None else str(value)
        return text.replace("|", "\\|").replace("\n", " ")

    lines = [
        "| " + " | ".join(_cell(h) for h in safe_headers) + " |",
        "| " + " | ".join("---" for _ in safe_headers) + " |",
    ]
    for row in safe_rows:
        padded = list(row) + [""] * max(0, len(safe_headers) - len(row))
        lines.append("| " + " | ".join(_cell(v) for v in padded[:len(safe_headers)]) + " |")
    return "\n".join(lines)


def _summarize_repeated_tool_result(tool_name: str, result_json: str) -> str:
    """
    Build a user-facing answer when the LLM repeats an already-successful
    read/query tool call instead of using the returned data.
    """
    try:
        payload = json.loads(result_json)
    except Exception:
        payload = result_json

    prefix = (
        f"工具 `{tool_name}` 已成功取得結果；AI 又重複呼叫同一工具，"
        "已停止重複執行並直接呈現上次結果。"
    )

    if tool_name == "get_used_range" and isinstance(payload, str):
        return f"{prefix}\n\n已使用範圍：`{payload}`"

    if tool_name == "read_range" and isinstance(payload, list):
        if payload and all(isinstance(row, list) for row in payload):
            headers = payload[0] if payload and all(isinstance(v, str) for v in payload[0]) else []
            rows = payload[1:] if headers else payload
            table = _markdown_table(headers, rows)
            suffix = f"\n\n共 {len(rows)} 列，顯示前 {min(len(rows), 8)} 列。"
            return f"{prefix}{suffix}\n\n{table}" if table else f"{prefix}{suffix}"
        return f"{prefix}\n\n結果：`{payload}`"

    if tool_name == "query_range" and isinstance(payload, dict):
        headers = payload.get("headers", [])
        rows = payload.get("filtered_rows", [])
        filtered_count = payload.get("filtered_count", len(rows) if isinstance(rows, list) else 0)
        if isinstance(headers, list) and isinstance(rows, list):
            table = _markdown_table(headers, [r for r in rows if isinstance(r, list)])
            msg = f"{prefix}\n\n符合條件：{filtered_count} 筆，顯示前 {min(filtered_count, 8)} 筆。"
            return f"{msg}\n\n{table}" if table else msg

    if tool_name == "summarize_range" and isinstance(payload, dict):
        pairs = [(k, v) for k, v in payload.items() if v is not None]
        if pairs:
            table = _markdown_table(["Metric", "Value"], [[k, v] for k, v in pairs])
            return f"{prefix}\n\n{table}"

    return f"{prefix}\n\n請展開上方工具結果查看完整內容。"


def _summarize_completed_tool_result(tool_name: str, result_json: str) -> str:
    """Build a plain final message when a successful tool call had no LLM final."""
    try:
        payload = json.loads(result_json)
    except Exception:
        payload = {}

    if tool_name == "beautify_range" and isinstance(payload, dict):
        sheet = payload.get("sheet") or "目前工作表"
        range_addr = payload.get("range") or payload.get("range_addr") or "指定範圍"
        theme = payload.get("theme") or "預設"
        applied = payload.get("applied") or []
        details = []
        if any(str(item).startswith("header") for item in applied):
            details.append("表頭")
        if any(str(item).startswith("banded_rows") for item in applied):
            details.append("交錯列底色")
        if any(str(item).startswith("number_format") for item in applied):
            details.append("數字格式")
        if "filter" in applied:
            details.append("篩選按鈕")
        if "auto_fit_columns" in applied:
            details.append("自動欄寬")
        detail_text = "、".join(details) if details else "基礎美化格式"
        return f"已完成表格美化：`{sheet}!{range_addr}`，套用 `{theme}` 主題，包含{detail_text}。"

    if tool_name == "write_range" and isinstance(payload, dict):
        target = payload.get("range") or payload.get("range_addr") or "指定範圍"
        return f"已完成寫入：`{target}`。"

    if tool_name == "fill_series" and isinstance(payload, dict):
        target = payload.get("range") or payload.get("filled_range") or payload.get("start_cell") or "指定範圍"
        return f"已完成數列填入：`{target}`。"

    if tool_name == "query_range" and isinstance(payload, dict):
        count = payload.get("filtered_count")
        aggregation = payload.get("aggregation_result") or payload.get("aggregation")
        if count is not None and aggregation is not None:
            return f"查詢完成：符合條件 {count} 筆，彙總結果為 `{aggregation}`。"
        if count is not None:
            return f"查詢完成：符合條件 {count} 筆。"

    return "已完成操作。"


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

    if not isinstance(payload, dict):
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


def _payload_has_error(payload: object) -> bool:
    """Return True only for dict payloads that carry tool error fields."""
    return isinstance(payload, dict) and (
        "error" in payload or payload.get("status") == "error"
    )


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


_FORMAT_STYLE_KEYS = {
    "bold",
    "italic",
    "color",
    "fill",
    "font_size",
    "number_format",
    "horizontal_alignment",
}


def _last_user_text(messages: list[dict]) -> str:
    for msg in reversed(messages):
        if msg.get("role") == "user":
            return str(msg.get("content") or "")
    return ""


def _repair_format_range_args(arguments: dict, user_text: str) -> dict:
    """
    Some OpenAI-compatible local models occasionally call format_range with
    only range_addr/sheet even when the prompt explicitly names simple styles.
    Keep executor-level no-op rejection, but conservatively recover common
    style words from the user's latest prompt.
    """
    if any(arguments.get(k) is not None for k in _FORMAT_STYLE_KEYS):
        return arguments

    text = user_text.lower()
    repaired = dict(arguments)

    if "粗體" in user_text or "bold" in text:
        repaired["bold"] = True
    if "斜體" in user_text or "italic" in text:
        repaired["italic"] = True

    fill_map = {
        "藍底": "#4472C4",
        "藍色底": "#4472C4",
        "黃底": "#FFFF00",
        "黃色底": "#FFFF00",
        "綠底": "#00B050",
        "綠色底": "#00B050",
        "紅底": "#FF0000",
        "紅色底": "#FF0000",
        "黑底": "#000000",
        "黑色底": "#000000",
        "白底": "#FFFFFF",
        "白色底": "#FFFFFF",
    }
    for token, color in fill_map.items():
        if token in user_text:
            repaired["fill"] = color
            break

    font_color_map = {
        "白字": "#FFFFFF",
        "白色字": "#FFFFFF",
        "黑字": "#000000",
        "黑色字": "#000000",
        "紅字": "#FF0000",
        "紅色字": "#FF0000",
        "藍字": "#0070C0",
        "藍色字": "#0070C0",
        "綠字": "#00B050",
        "綠色字": "#00B050",
    }
    for token, color in font_color_map.items():
        if token in user_text:
            repaired["color"] = color
            break

    if "置中" in user_text or "居中" in user_text or "center" in text:
        repaired["horizontal_alignment"] = "center"
    elif "靠左" in user_text or "向左" in user_text or "left" in text:
        repaired["horizontal_alignment"] = "left"
    elif "靠右" in user_text or "向右" in user_text or "right" in text:
        repaired["horizontal_alignment"] = "right"

    return repaired


def _header_column_index(arguments: dict, user_text: str) -> int | None:
    range_addr = arguments.get("range_addr")
    if not range_addr:
        return None

    try:
        rows = et.read_range(str(range_addr), arguments.get("sheet"))
    except Exception:
        return None
    if not rows:
        return None

    headers = rows[0]
    lowered = user_text.lower()
    for idx, header in enumerate(headers, start=1):
        if header is None:
            continue
        name = str(header).strip()
        if name and name.lower() in lowered:
            return idx
    return None


def _repair_range_column_args(tool_name: str, arguments: dict, user_text: str) -> dict:
    if tool_name not in {"sort_range", "filter_range"}:
        return arguments

    repaired = dict(arguments)
    idx = _header_column_index(repaired, user_text)
    if idx is not None:
        repaired["column_index"] = idx

    if tool_name == "sort_range":
        if any(token in user_text for token in ("表頭", "標題", "第一列")):
            repaired["has_header"] = True
        lowered = user_text.lower()
        if any(token in user_text for token in ("由大到小", "大到小", "降冪")) or "descending" in lowered:
            repaired["ascending"] = False
        elif any(token in user_text for token in ("由小到大", "小到大", "升冪")) or "ascending" in lowered:
            repaired["ascending"] = True

    elif tool_name == "filter_range" and not repaired.get("criteria"):
        # Common phrasing: "Region 是 North" / "Region=North".
        try:
            rows = et.read_range(str(repaired["range_addr"]), repaired.get("sheet"))
            headers = [str(h).strip() for h in (rows[0] if rows else [])]
            if idx is not None and 1 <= idx <= len(headers):
                header = re.escape(headers[idx - 1])
                match = re.search(rf"{header}\s*(?:=|是|為)\s*([A-Za-z0-9_\-]+)", user_text)
                if match:
                    repaired["criteria"] = match.group(1)
        except Exception:
            pass

    return repaired


def _repair_data_validation_args(arguments: dict, user_text: str) -> dict:
    repaired = dict(arguments)

    if not repaired.get("options"):
        for alias in ("formula1", "formula", "values", "items"):
            value = repaired.get(alias)
            if value:
                repaired["options"] = value
                break

    if not repaired.get("options"):
        match = re.search(
            r"(?:選項|清單|下拉選單|options?)\s*(?:是|為|=|:|：)?\s*([^\n。]+)",
            user_text,
            flags=re.IGNORECASE,
        )
        if match:
            options = match.group(1).strip().strip("'\"`")
            options = options.replace("，", ",").replace("、", ",").replace("；", ";")
            repaired["options"] = options

    for alias in ("formula1", "formula", "values", "items", "validation_type"):
        repaired.pop(alias, None)

    return repaired


def _repair_fill_series_args(arguments: dict, user_text: str) -> dict:
    repaired = dict(arguments)
    text = user_text.lower()

    if not repaired.get("start_value"):
        match = re.search(r"start_value\s*=\s*([\-]?\d+(?:\.\d+)?)", text)
        if not match:
            match = re.search(r"從\s*([\-]?\d+(?:\.\d+)?)\s*(?:到|至)\s*([\-]?\d+(?:\.\d+)?)", user_text)
        if not match:
            match = re.search(r"填\s*([\-]?\d+(?:\.\d+)?)\s*(?:到|至)\s*([\-]?\d+(?:\.\d+)?)", user_text)
        if match:
            repaired["start_value"] = match.group(1)

    if not repaired.get("count"):
        match = re.search(r"count\s*=\s*(\d+)", text)
        if match:
            repaired["count"] = int(match.group(1))
        else:
            span = re.search(r"從\s*([\-]?\d+(?:\.\d+)?)\s*(?:到|至)\s*([\-]?\d+(?:\.\d+)?)", user_text)
            if not span:
                span = re.search(r"填\s*([\-]?\d+(?:\.\d+)?)\s*(?:到|至)\s*([\-]?\d+(?:\.\d+)?)", user_text)
            if span:
                try:
                    start = float(span.group(1))
                    end = float(span.group(2))
                    step = float(repaired.get("step") or 1)
                    if step:
                        repaired["count"] = int(abs((end - start) / step)) + 1
                except Exception:
                    pass

    return repaired


def _repair_page_setup_args(arguments: dict, user_text: str) -> dict:
    repaired = dict(arguments)
    text = user_text.lower()

    if "orientation" not in repaired:
        if "landscape" in text or "橫向" in user_text or "橫印" in user_text:
            repaired["orientation"] = "landscape"
        elif "portrait" in text or "直向" in user_text or "直印" in user_text:
            repaired["orientation"] = "portrait"

    if "paper_size" not in repaired:
        match = re.search(r"paper_size\s*=\s*([A-Za-z0-9]+)", text)
        if match:
            repaired["paper_size"] = match.group(1)
        elif "a4" in text:
            repaired["paper_size"] = "a4"

    for key in ("fit_to_wide", "fit_to_tall"):
        if key not in repaired:
            match = re.search(rf"{key}\s*=\s*(\d+)", text)
            if match:
                repaired[key] = int(match.group(1))

    if "print_area" not in repaired:
        match = re.search(r"print_area\s*=\s*([A-Za-z]+\d+\s*:\s*[A-Za-z]+\d+)", user_text, flags=re.IGNORECASE)
        if match:
            repaired["print_area"] = match.group(1).replace(" ", "")

    if "center_horizontally" not in repaired:
        match = re.search(r"center_horizontally\s*=\s*(true|false)", text)
        if match:
            repaired["center_horizontally"] = match.group(1) == "true"
        elif "水平置中" in user_text:
            repaired["center_horizontally"] = True

    if "center_vertically" not in repaired:
        match = re.search(r"center_vertically\s*=\s*(true|false)", text)
        if match:
            repaired["center_vertically"] = match.group(1) == "true"
        elif "垂直置中" in user_text:
            repaired["center_vertically"] = True

    return repaired


def _repair_query_range_args(arguments: dict, user_text: str) -> dict:
    repaired = dict(arguments)
    query_text = f"{repaired.pop('query', '')} {user_text}".strip()

    headers: list[str] = []
    range_addr = repaired.get("range_addr")
    if range_addr:
        try:
            rows = et.read_range(str(range_addr), repaired.get("sheet"))
            headers = [str(h).strip() for h in (rows[0] if rows else []) if h is not None]
        except Exception:
            headers = []

    if not repaired.get("filters") and not repaired.get("condition_json") and headers:
        filters = []
        for header in headers:
            pattern = rf"{re.escape(header)}\s*(?:=|是|為)\s*([A-Za-z0-9_\-]+)"
            match = re.search(pattern, query_text, flags=re.IGNORECASE)
            if match:
                filters.append({"column": header, "operator": "=", "value": match.group(1)})
        if filters:
            repaired["filters"] = filters

    if not repaired.get("aggregation") and not repaired.get("aggregation_json") and headers:
        lowered = query_text.lower()
        function = None
        if "加總" in query_text or "總和" in query_text or "sum" in lowered:
            function = "sum"
        elif "平均" in query_text or "average" in lowered or "avg" in lowered:
            function = "avg"
        elif "最大" in query_text or "max" in lowered:
            function = "max"
        elif "最小" in query_text or "min" in lowered:
            function = "min"
        elif "計數" in query_text or "筆數" in query_text or "count" in lowered:
            function = "count"

        if function:
            agg_header = None
            stat_words = "加總|總和|平均|最大|最小|計數|筆數|sum|average|avg|max|min|count"
            for header in headers:
                if header and re.search(rf"{re.escape(header)}\s*(?:{stat_words})", query_text, flags=re.IGNORECASE):
                    agg_header = header
                    break

            if agg_header is None:
                filter_columns = {
                    str(f.get("column", "")).strip().lower()
                    for f in repaired.get("filters", [])
                    if isinstance(f, dict)
                }
                for header in headers:
                    if header and header.lower() in lowered and header.lower() not in filter_columns:
                        agg_header = header
                        break

            if agg_header is not None:
                repaired["aggregation"] = {"function": function, "column": agg_header}
            elif function == "count":
                repaired["aggregation"] = {"function": "count", "column": headers[0]}

    return repaired


def _repair_tool_calls(resp: LLMResponse, messages: list[dict]) -> None:
    user_text = _last_user_text(messages)
    changed = False
    for tc in resp.tool_calls:
        if tc.name == "format_range":
            repaired = _repair_format_range_args(tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True
        elif tc.name in {"sort_range", "filter_range"}:
            repaired = _repair_range_column_args(tc.name, tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True
        elif tc.name == "set_data_validation":
            repaired = _repair_data_validation_args(tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True
        elif tc.name == "fill_series":
            repaired = _repair_fill_series_args(tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True
        elif tc.name == "page_setup":
            repaired = _repair_page_setup_args(tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True
        elif tc.name == "query_range":
            repaired = _repair_query_range_args(tc.arguments, user_text)
            if repaired != tc.arguments:
                tc.arguments = repaired
                changed = True

    raw = resp.raw_assistant_message
    if changed and isinstance(raw, dict):
        raw_calls = raw.get("tool_calls")
        if isinstance(raw_calls, list):
            by_id = {tc.id: tc for tc in resp.tool_calls}
            for raw_tc in raw_calls:
                try:
                    tc_id = raw_tc.get("id")
                    tc = by_id.get(tc_id)
                    if tc is not None:
                        raw_tc["function"]["arguments"] = json.dumps(
                            tc.arguments, ensure_ascii=False
                        )
                except Exception:
                    pass


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
    successful_results_by_sig: dict[tuple[str, str], str] = {}
    last_successful_tool: tuple[str, str] | None = None

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
                    _repair_tool_calls(resp, msgs)

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
                        if (
                            same_count >= 2
                            and tc.name in _REPEAT_SUMMARY_TOOLS
                            and sig in successful_results_by_sig
                        ):
                            yield (
                                EVT_DONE,
                                _summarize_repeated_tool_result(
                                    tc.name,
                                    successful_results_by_sig[sig],
                                ),
                            )
                            return
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
                        has_err = _payload_has_error(_parsed)

                        # Enrich error result with actionable hint for the LLM
                        if has_err:
                            result_json = _enrich_error_result(tc.name, result_json)

                        yield (EVT_TOOL_DONE, ToolExecution(
                            tc=tc,
                            result_json=result_json,
                            backup_entry=entry,
                            has_error=has_err,
                        ))
                        if not has_err:
                            successful_results_by_sig[sig] = result_json
                            last_successful_tool = (tc.name, result_json)

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

    if last_successful_tool is not None:
        yield (
            EVT_DONE,
            _summarize_completed_tool_result(
                last_successful_tool[0],
                last_successful_tool[1],
            ),
        )
    else:
        yield (EVT_DONE, "完成 ✓")
