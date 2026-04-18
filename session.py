"""
對話訊息管理：初始化、上下文摘要壓縮、儲存與載入。
所有 Streamlit session_state["messages"] 的讀寫都經由此模組。
"""

import json
import streamlit as st
from config import SYSTEM_PROMPT

# 超過此輪數時觸發上下文摘要壓縮
CONTEXT_SUMMARIZE_THRESHOLD = 30
# 摘要後保留最近幾輪
KEEP_RECENT = 10


def reset_messages() -> None:
    """重置對話，只保留 system prompt。"""
    st.session_state["messages"] = [{"role": "system", "content": SYSTEM_PROMPT}]


def get_messages() -> list[dict]:
    """取得完整訊息串列（含 system prompt）。"""
    return st.session_state.get("messages", [])


def append_message(msg: dict) -> None:
    """追加一則訊息。"""
    st.session_state["messages"].append(msg)


def non_system_count() -> int:
    """回傳非 system 訊息的數量（等同對話輪數指標）。"""
    return sum(1 for m in get_messages() if m["role"] != "system")


def maybe_summarize(provider) -> None:
    """
    上下文自動摘要：非 system 訊息超過閾值時，
    保留最近 KEEP_RECENT 輪，中間歷史壓縮成摘要。
    """
    msgs = get_messages()
    non_sys = [m for m in msgs if m["role"] != "system"]
    if len(non_sys) <= CONTEXT_SUMMARIZE_THRESHOLD:
        return

    to_summarize = non_sys[:-KEEP_RECENT]
    recent = non_sys[-KEEP_RECENT:]

    summary_prompt = [
        {"role": "system", "content": "請用繁體中文，用 3~5 句話摘要以下對話的重點操作與結果："},
        *to_summarize,
        {"role": "user", "content": "請摘要以上對話。"},
    ]
    try:
        resp = provider.chat(summary_prompt, [])
        summary_text = resp.text or "（先前操作摘要）"
    except Exception:
        summary_text = "（先前有多輪操作，摘要失敗）"

    st.session_state["messages"] = [
        {"role": "system",    "content": SYSTEM_PROMPT},
        {"role": "assistant", "content": f"📋 **先前操作摘要**\n\n{summary_text}"},
        *recent,
    ]


def messages_to_json() -> str:
    """匯出非 system 訊息為 JSON 字串（供下載）。"""
    exportable = [m for m in get_messages() if m["role"] != "system"]
    return json.dumps(exportable, ensure_ascii=False, indent=2)


def load_messages_from_json(raw: bytes | str) -> int:
    """
    從 JSON bytes/字串載入訊息，追加到現有對話。
    回傳載入的訊息筆數。
    """
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8")
    loaded = json.loads(raw)
    reset_messages()
    st.session_state["messages"].extend(loaded)
    return len(loaded)
