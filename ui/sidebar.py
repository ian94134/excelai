"""
Streamlit 側邊欄 UI 元件。
負責：Qwen 伺服器設定、Excel 狀態顯示、對話記錄儲存/載入/清除。
"""

import streamlit as st
import excel_tools as et
from tools.executor import execute
from config import QWEN_BASE_URL, QWEN_MODEL
import session


def _init_provider(base_url: str, model: str):
    from config import get_provider
    return get_provider(base_url=base_url, model=model)


def render(version: str = "") -> None:
    """渲染完整側邊欄。Provider 已快取於 st.session_state['_provider']。"""
    with st.sidebar:
        st.title("📊 Excel AI 助手")
        if version:
            st.caption(version)

        _render_qwen_settings()
        st.divider()
        _render_excel_status()
        st.divider()
        _render_chat_history()


# ── 內部元件 ──────────────────────────────────────────────────────────────────

def _render_qwen_settings() -> None:
    with st.expander("⚙️ Qwen 伺服器設定", expanded=False):
        qwen_url = st.text_input(
            "Server URL",
            value=st.session_state.get("qwen_url", QWEN_BASE_URL),
        )
        qwen_model = st.text_input(
            "Model ID",
            value=st.session_state.get("qwen_model", QWEN_MODEL),
        )
        if st.button("🔌 套用並測試連線", use_container_width=True):
            st.session_state["qwen_url"]   = qwen_url
            st.session_state["qwen_model"] = qwen_model
            try:
                st.session_state["_provider"] = _init_provider(qwen_url, qwen_model)
                st.session_state.pop("_provider_err", None)
            except Exception as e:
                st.session_state["_provider_err"] = str(e)
            st.rerun()

    # 首次載入時初始化 provider
    if "_provider" not in st.session_state:
        try:
            st.session_state["_provider"] = _init_provider(
                st.session_state.get("qwen_url",   QWEN_BASE_URL),
                st.session_state.get("qwen_model", QWEN_MODEL),
            )
        except Exception as e:
            st.session_state["_provider_err"] = str(e)

    if st.session_state.get("_provider_err"):
        st.error(f"連線失敗：{st.session_state['_provider_err']}")
    else:
        st.success("✅ Qwen 本地模型已連接")


def _render_excel_status() -> None:
    st.subheader("📁 Excel 狀態")
    try:
        info = et.get_sheet_info()
        st.success(f"✅ {info['file_name']}")
        st.caption(f"作用中：**{info['active_sheet']}**　共 {len(info['sheets'])} 張工作表")
        with st.expander("所有工作表"):
            for s in info["sheets"]:
                marker = "▶" if s == info["active_sheet"] else "　"
                st.text(f"{marker} {s}")
    except Exception:
        st.warning("⚠️ 未偵測到 Excel：請先開啟 Excel")

    col_r, col_s = st.columns(2)
    with col_r:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()
    with col_s:
        if st.button("💾 儲存檔案", use_container_width=True):
            try:
                execute("save_workbook", {})
                st.toast("✅ 已儲存", icon="💾")
            except Exception as e:
                st.error(str(e))


def _render_chat_history() -> None:
    st.subheader("💬 對話記錄")

    if session.non_system_count() > 0:
        st.download_button(
            label="💾 儲存對話",
            data=session.messages_to_json(),
            file_name="excel_ai_chat.json",
            mime="application/json",
            use_container_width=True,
        )

    uploaded = st.file_uploader("📂 載入對話", type="json", label_visibility="collapsed")
    if uploaded:
        try:
            count = session.load_messages_from_json(uploaded.read())
            st.success(f"已載入 {count} 則對話")
            st.rerun()
        except Exception as e:
            st.error(f"載入失敗：{e}")

    if st.button("🗑️ 清除對話", use_container_width=True):
        session.reset_messages()
        st.session_state.pop("_pending_confirm", None)
        st.rerun()

    count = session.non_system_count()
    st.caption(f"對話輪數：{count} / {session.CONTEXT_SUMMARIZE_THRESHOLD}（超過自動摘要）")
