"""
compress.py — Tool result JSON compression utilities (v4.6.0).

Pure functions with no Streamlit dependency; safe to import in test files
without triggering session_state side-effects.

Used by session.append_message() to keep tool results within the LLM's
context window budget.
"""
from __future__ import annotations

import json

# Single tool result exceeding this character count (~750 tokens) is compressed.
TOOL_CONTENT_LIMIT = 3000


def compress_tool_result(content: str, limit: int = TOOL_CONTENT_LIMIT) -> str:
    """
    Compress a large tool result JSON to stay under `limit` characters.

    Design rules:
    - Error results (containing "error" key) are returned unchanged —
      Qwen needs the full error context to decide on recovery steps.
    - Success dict: large arrays (> 500 chars) replaced with descriptive
      summaries; long strings truncated to 500 chars; other fields kept.
    - List payloads (e.g. read_range 2D arrays): replaced with row/col count
      + 3-row sample so Qwen retains data awareness.
    - Non-JSON content: naively truncated with a note.
    - Compressed payloads get "_compressed": true so logs can detect them.

    This is a pure function — no side effects, no Streamlit dependency.
    """
    if len(content) <= limit:
        return content

    try:
        payload = json.loads(content)
    except Exception:
        return content[:limit] + f"…[已截斷，原始長度 {len(content)} 字元]"

    # Error results: never compress (LLM must see full context)
    if isinstance(payload, dict) and "error" in payload:
        return content

    # Success dict: prune large fields
    if isinstance(payload, dict):
        pruned: dict = {}
        for k, v in payload.items():
            serialized = json.dumps(v, ensure_ascii=False)
            if isinstance(v, list) and len(serialized) > 500:
                pruned[k] = f"[陣列，{len(v)} 列，已略去詳細資料]"
            elif isinstance(v, str) and len(v) > 500:
                pruned[k] = v[:500] + "…[已截斷]"
            else:
                pruned[k] = v
        pruned["_compressed"] = True
        result = json.dumps(pruned, ensure_ascii=False)
        if len(result) <= limit:
            return result
        return result[:limit] + f"…[已截斷，原始長度 {len(content)} 字元]"

    # List payload (e.g. read_range returns 2D list)
    if isinstance(payload, list):
        row_count = len(payload)
        col_count = len(payload[0]) if payload else 0
        sample    = payload[:3]
        summary   = {
            "rows":        row_count,
            "columns":     col_count,
            "sample":      sample,
            "_compressed": True,
            "_note":       f"共 {row_count} 列資料，僅顯示前 3 列",
        }
        return json.dumps(summary, ensure_ascii=False, default=str)

    # Other types: truncate
    return content[:limit] + f"…[已截斷，原始長度 {len(content)} 字元]"
