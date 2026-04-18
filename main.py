"""
Excel AI 助手 V3（重構版）
執行方式：streamlit run main.py
"""

import json
import streamlit as st
from tools import OPENAI_TOOLS, execute, DANGEROUS_TOOLS
import session
import ui.sidebar as sidebar

VERSION = "v3.1.0"

st.set_page_config(page_title="Excel AI 助手", page_icon="📊", layout="wide")

# ── 側邊欄（完全委派給 ui/sidebar.py）────────────────────────────────────────
sidebar.render(VERSION)

# ── 初始化訊息 ────────────────────────────────────────────────────────────────
if "messages" not in st.session_state:
    session.reset_messages()

# ── 顯示歷史訊息 ──────────────────────────────────────────────────────────────
for m in session.get_messages():
    if m["role"] == "user":
        with st.chat_message("user"):
            st.markdown(m["content"])
    elif m["role"] == "assistant" and m.get("content"):
        with st.chat_message("assistant"):
            st.markdown(m["content"])


# ── 危險操作確認區 ────────────────────────────────────────────────────────────
if "_pending_confirm" in st.session_state:
    pending = st.session_state["_pending_confirm"]
    tc = pending["tool_call"]
    st.warning(
        f"⚠️ **確認操作**\n\n"
        f"AI 想要執行：`{tc['name']}`\n\n"
        f"參數：`{tc['arguments']}`"
    )
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ 確認執行", use_container_width=True, type="primary"):
            result = execute(tc["name"], tc["arguments"])
            session.append_message({
                "role": "tool", "tool_call_id": tc["id"],
                "name": tc["name"], "content": result,
            })
            st.session_state.pop("_pending_confirm")
            st.rerun()
    with col2:
        if st.button("❌ 取消", use_container_width=True):
            session.append_message({
                "role": "tool", "tool_call_id": tc["id"],
                "name": tc["name"],
                "content": json.dumps({"error": "使用者取消執行"}, ensure_ascii=False),
            })
            st.session_state.pop("_pending_confirm")
            st.rerun()
    st.stop()


# ── 使用者輸入 ────────────────────────────────────────────────────────────────
if prompt := st.chat_input("例：篩選台北的資料 / 合併 A1:D1 / 設定外框線 / 建立下拉選單"):
    session.append_message({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    provider = st.session_state.get("_provider")
    if provider is None:
        st.error("Qwen 模型未連接，請在側邊欄設定後重試")
        st.stop()

    # 上下文自動摘要（超過閾值時）
    session.maybe_summarize(provider)

    with st.chat_message("assistant"):
        placeholder = st.empty()
        placeholder.markdown("⏳ 思考中…")

        for _ in range(10):
            has_tool_call_this_round = False

            for event, data in provider.chat_stream(session.get_messages(), OPENAI_TOOLS):
                if event == "text":
                    placeholder.markdown(data + "▌")

                elif event == "done":
                    final = data or "完成 ✓"
                    placeholder.markdown(final)
                    session.append_message({"role": "assistant", "content": final})

                elif event == "tool_calls":
                    has_tool_call_this_round = True
                    resp = data  # LLMResponse

                    session.append_message(
                        resp.raw_assistant_message or {"role": "assistant", "content": None}
                    )

                    for tc in resp.tool_calls:
                        # 危險操作 → 暫停確認
                        if tc.name in DANGEROUS_TOOLS:
                            st.session_state["_pending_confirm"] = {
                                "tool_call": {
                                    "id": tc.id, "name": tc.name,
                                    "arguments": tc.arguments,
                                }
                            }
                            placeholder.empty()
                            st.rerun()

                        # 一般操作 → 直接執行並顯示
                        label = f"🔧 {tc.name}　`{tc.arguments}`"
                        with st.status(label, expanded=False) as status:
                            result = execute(tc.name, tc.arguments)
                            st.code(result, language="json")
                            status.update(label=label, state="complete")

                        session.append_message({
                            "role": "tool", "tool_call_id": tc.id,
                            "name": tc.name, "content": result,
                        })

            if not has_tool_call_this_round:
                break  # 純文字回覆，對話結束
        else:
            placeholder.warning("⚠️ 執行輪數超過上限，請重新嘗試")
