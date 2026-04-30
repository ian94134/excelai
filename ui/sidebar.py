"""
Streamlit 側邊欄 UI 元件。
負責：Qwen 伺服器設定、Excel 狀態顯示、對話記錄儲存/載入/清除。
"""

import json
import streamlit as st
import excel_tools as et
from tools.executor import execute
from config import QWEN_BASE_URL, QWEN_MODEL
import session
import backup
import telemetry
import excel_event_watcher
from ui.tool_display import friendly_tool_label


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
        _render_selection_status()
        st.divider()
        _render_undo_panel()
        st.divider()
        _render_snapshot_panel()
        st.divider()
        _render_csv_import()
        st.divider()
        _render_plan_mode_toggle()
        st.divider()
        _render_macro_panel()
        st.divider()
        _render_op_log()
        st.divider()
        _render_telemetry()
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

        # 多工作簿切換
        try:
            wb_result = et.list_workbooks()
            books = wb_result.get("workbooks", [])
            if len(books) > 1:
                with st.expander(f"📂 切換活頁簿（共 {len(books)} 個）"):
                    for wb in books:
                        label = f"{'▶ ' if wb['active'] else '　'}{wb['name']}"
                        if not wb["active"]:
                            if st.button(label, use_container_width=True, key=f"wb_{wb['name']}"):
                                et.switch_workbook(wb["name"])
                                st.toast(f"✅ 已切換至「{wb['name']}」", icon="📂")
                                st.rerun()
                        else:
                            st.caption(f"▶ **{wb['name']}**（目前使用中）")
        except Exception:
            pass  # 找不到多個活頁簿時靜默略過

    except Exception:
        st.warning("⚠️ 未偵測到 Excel：請先開啟 Excel")

    col_r, col_s = st.columns(2)
    with col_r:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()
    with col_s:
        if st.button("💾 儲存檔案", use_container_width=True):
            try:
                raw = execute("save_workbook", {})
                payload = json.loads(raw)
                if isinstance(payload, dict) and payload.get("error"):
                    st.error(f"儲存失敗：{payload['error']}")
                else:
                    st.toast("✅ 已儲存", icon="💾")
            except Exception as e:
                st.error(str(e))


def _render_undo_panel() -> None:
    """顯示備份堆疊狀態與「↶ 復原上一步」按鈕。"""
    st.subheader("↶ 復原操作")

    stack = backup.get_session_stack()
    stack_size = len(stack) if stack else 0
    last_entry = stack.peek() if stack else None

    if stack_size == 0:
        st.caption("目前沒有可復原的操作")
        st.button("↶ 復原上一步", disabled=True, use_container_width=True)
        return

    # 顯示堆疊狀態
    st.caption(f"可復原步數：**{stack_size}** / 20")
    if last_entry:
        st.caption(
            f"最後操作：**{friendly_tool_label(last_entry.tool_name)}**"
            f"（{last_entry.timestamp:%H:%M:%S}）"
        )

    # 展開顯示完整堆疊（最新在前）
    if stack_size > 1:
        with st.expander(f"查看全部 {stack_size} 步"):
            for i, entry in enumerate(reversed(stack.snapshot())):
                st.text(
                    f"{i + 1}. {friendly_tool_label(entry.tool_name)}"
                    f"（{entry.timestamp:%H:%M:%S}）"
                )

    if st.button("↶ 復原上一步", use_container_width=True, type="primary"):
        try:
            raw = execute("undo_last", {})
            payload = json.loads(raw)
            status = payload.get("status")
            if status == "ok":
                st.toast(f"✅ 已復原：{friendly_tool_label(payload.get('undone', ''))}", icon="↶")
            elif status == "no_op":
                st.info("備份堆疊為空，沒有可復原的操作。")
            elif status == "cannot_undo":
                st.warning(f"⚠️ 無法自動還原\n\n{payload.get('message', '')}")
            else:
                st.error(f"復原失敗：{payload.get('error', raw)}")
            st.rerun()
        except Exception as e:
            st.error(f"復原失敗：{e}")


def _render_snapshot_panel() -> None:
    """工作表快照：建立 / 顯示 / 還原。"""
    st.subheader("📸 工作表快照")

    snap = st.session_state.get("_sheet_snapshot")

    col_snap, col_restore = st.columns(2)
    with col_snap:
        if st.button("📸 建立快照", use_container_width=True):
            try:
                result = et.snapshot_sheet()
                st.session_state["_sheet_snapshot"] = result
                st.toast(
                    f"✅ 已快照「{result['sheet']}」"
                    f"（{result['rows']} 列 × {result['cols']} 欄）",
                    icon="📸",
                )
                st.rerun()
            except Exception as e:
                st.error(f"快照失敗：{e}")

    with col_restore:
        if st.button(
            "↩ 還原快照",
            use_container_width=True,
            disabled=(snap is None),
            type="primary" if snap else "secondary",
        ):
            st.session_state["_confirm_restore"] = True
            st.rerun()

    if snap:
        st.caption(
            f"快照：「{snap['sheet']}」　"
            f"{snap['rows']} 列 × {snap['cols']} 欄　"
            f"範圍 {snap['range']}"
        )
    else:
        st.caption("尚無快照（建立快照後可一鍵還原整張工作表）")

    # 確認還原對話
    if st.session_state.get("_confirm_restore") and snap:
        st.warning(
            f"⚠️ 確認要把「{snap['sheet']}」還原成快照狀態？\n\n"
            "目前工作表內容將被覆蓋（此操作不可 undo）。"
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ 確認還原", use_container_width=True, type="primary"):
                try:
                    et.restore_snapshot(
                        snap["data"], "A1",
                        snap["sheet"],
                    )
                    st.session_state.pop("_confirm_restore", None)
                    st.toast("✅ 已還原至快照狀態", icon="↩")
                    st.rerun()
                except Exception as e:
                    st.error(f"還原失敗：{e}")
        with c2:
            if st.button("❌ 取消", use_container_width=True):
                st.session_state.pop("_confirm_restore", None)
                st.rerun()


def _render_csv_import() -> None:
    """CSV / 剪貼簿快速導入：貼入文字後直接寫入 Excel 目前選取位置。"""
    with st.expander("📋 CSV 快速導入", expanded=False):
        st.caption("把 CSV 或試算表複製的資料貼入下方，按確認寫入 Excel 目前選取的起始儲存格。")

        csv_text = st.text_area(
            "貼上資料（CSV 格式，以逗號或 Tab 分隔）",
            height=120,
            placeholder="月份,銷售額,目標\n1月,125000,100000\n2月,98000,100000",
            label_visibility="collapsed",
        )

        col_dest, col_delim = st.columns([2, 1])
        with col_dest:
            dest_cell = st.text_input("起始儲存格", value="A1", placeholder="A1")
        with col_delim:
            delimiter = st.selectbox("分隔符", ["自動偵測", "逗號", "Tab", "分號"], index=0)

        if st.button("✅ 寫入 Excel", use_container_width=True, disabled=not csv_text.strip()):
            try:
                import csv, io

                # 決定分隔符
                delim_map = {"逗號": ",", "Tab": "\t", "分號": ";"}
                if delimiter == "自動偵測":
                    sample = csv_text[:500]
                    sniffer = csv.Sniffer()
                    try:
                        sep = sniffer.sniff(sample, delimiters=",\t;").delimiter
                    except Exception:
                        sep = ","
                else:
                    sep = delim_map[delimiter]

                reader = csv.reader(io.StringIO(csv_text.strip()), delimiter=sep)
                rows   = [list(row) for row in reader]

                if not rows:
                    st.warning("資料為空，請重新貼上。")
                else:
                    # 取得目前作用中工作表
                    info   = et.get_sheet_info()
                    sheet  = info.get("active_sheet")
                    result = execute("write_range", {
                        "range_addr": dest_cell,
                        "values":     rows,
                        "sheet":      sheet,
                    })
                    import json as _json
                    payload = _json.loads(result)
                    if payload.get("error"):
                        st.error(f"寫入失敗：{payload['error']}")
                    else:
                        st.toast(
                            f"✅ 已寫入 {len(rows)} 列 × {len(rows[0])} 欄到「{sheet}」{dest_cell}",
                            icon="📋",
                        )
                        st.rerun()
            except Exception as e:
                st.error(f"導入失敗：{e}")


def _render_plan_mode_toggle() -> None:
    """任務規劃模式開關：啟用後 AI 先列計劃，等確認再執行。"""
    plan_mode = st.toggle(
        "📋 任務規劃模式",
        value=st.session_state.get("_plan_mode", False),
        help=(
            "啟用後，AI 遇到複雜任務時會先用文字列出步驟計劃，"
            "等你按「▶ 開始執行」確認後才真正操作 Excel。\n\n"
            "適合：多步驟任務、第一次嘗試新操作、不確定 AI 會做什麼時。"
        ),
    )
    st.session_state["_plan_mode"] = plan_mode
    if plan_mode:
        st.caption("✅ 啟用中：AI 會先規劃再執行")


def _render_op_log() -> None:
    """本次 session 的操作日誌（來自 BackupStack + 唯讀工具記錄）。"""
    stack      = backup.get_session_stack()
    all_ops    = list(stack.snapshot()) if stack else []
    read_ops   = st.session_state.get("_read_op_log", [])  # 唯讀工具另外記
    total      = len(all_ops) + len(read_ops)

    with st.expander(f"📜 操作記錄（{total} 筆）", expanded=False):
        if total == 0:
            st.caption("本次尚無操作記錄")
            return

        # 可還原操作（從 BackupStack，最新在前）
        if all_ops:
            st.caption("**可還原操作**")
            for i, entry in enumerate(reversed(all_ops)):
                rng = entry.arguments.get("range_addr", "")
                detail = f"  `{rng}`" if rng else ""
                st.text(
                    f"{i+1:02d}. {entry.timestamp:%H:%M:%S}  "
                    f"{friendly_tool_label(entry.tool_name)}{detail}"
                )

        # 唯讀操作
        if read_ops:
            st.caption("**讀取 / 唯讀操作**")
            for item in read_ops[-10:]:  # 最多顯示最後 10 筆
                label = item.get("label") or friendly_tool_label(item.get("tool"))
                st.text(f"     {item['time']}  {label}")

        if st.button("🗑️ 清除記錄", use_container_width=True):
            if stack:
                stack.clear()
            st.session_state.pop("_read_op_log", None)
            st.rerun()


def _render_telemetry() -> None:
    st.subheader("📈 使用統計")
    stats = telemetry.get_summary()

    if stats["total"] == 0:
        st.caption("尚無使用記錄，執行工具後會自動累積。")
        return

    # ── 摘要指標 ──────────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    c1.metric("總操作數", stats["total"])
    c2.metric("成功率", f"{stats['success_rate'] * 100:.1f}%")

    # ── 最常用操作 ────────────────────────────────────────────────────────────
    if stats["top_tools"]:
        with st.expander("🏆 最常用操作", expanded=True):
            for rank, (name, cnt) in enumerate(stats["top_tools"], 1):
                st.text(f"  {rank}. {friendly_tool_label(name):<28} {cnt} 次")

    # ── 平均最慢工具 ───────────────────────────────────────────────
    if stats["slowest_tools"]:
        with st.expander("⏱️ 平均耗時最久（≥3 次）", expanded=False):
            for name, avg_ms in stats["slowest_tools"]:
                st.text(f"  {friendly_tool_label(name):<28} {avg_ms} ms")

    # ── 最近錯誤 ───────────────────────────────────────────────────────────
    if stats["recent_errors"]:
        with st.expander(
            f"❌ 最近錯誤（{len(stats['recent_errors'])} 筆）",
            expanded=False,
        ):
            for tool, err_type, ts in stats["recent_errors"]:
                st.text(f"  {ts[11:19]}  {friendly_tool_label(tool)} → {err_type or 'Unknown'}")

    # ── 清除按鈕 ───────────────────────────────────────────────────────
    if st.button("🗑️ 清除使用統計", use_container_width=True):
        telemetry.clear()
        st.rerun()


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
        except Exception as e:
            st.error(f"載入失敗：{e}")

    if st.button("🗑️ 清除對話", use_container_width=True):
        session.reset_messages()
        st.session_state.pop("_pending_confirm", None)
        st.rerun()

    count = session.non_system_count()
    st.caption(f"對話輪數：{count} / {session.CONTEXT_SUMMARIZE_THRESHOLD}（超過自動摘要）")

def _render_selection_status() -> None:
    """顯示 Excel 目前選取範圍（由背景 watcher 即時更新）。"""
    sel = excel_event_watcher.get_current_selection()
    if sel and sel.get("address") and sel["address"] != "(non-range selection)":
        addr  = sel["address"]
        sheet = sel.get("sheet", "")
        wb    = sel.get("workbook", "")
        label = f"{sheet}!{addr}" if sheet else addr
        st.caption(f"🎯 **目前選取**　`{label}`")
        if wb:
            st.caption(f"　活頁簿：{wb}")
    else:
        st.caption("🎯 **目前選取**　（無選取）")


def _render_macro_panel() -> None:
    """巨集管理面板：列出 / 執行 / 錄製 / 刪除巨集（v4.7.0）。"""
    import macro as _macro

    with st.expander("🔴 巨集管理", expanded=False):
        result = _macro.list_macros()
        macros = result.get("macros", [])
        pending = st.session_state.get("_pending_macro_confirm")

        if pending:
            st.warning(f"巨集「{pending['name']}」包含危險工具，請確認是否執行。")
            for step in pending.get("dangerous_steps", []):
                st.caption(f"第 {step['index']} 步：{friendly_tool_label(step['tool'])}")
            confirm_col, cancel_col = st.columns(2)
            with confirm_col:
                if st.button("確認執行", key="confirm_macro_run", type="primary", use_container_width=True):
                    run_result = _macro.run_macro(pending["name"], confirm_dangerous=True)
                    st.session_state.pop("_pending_macro_confirm", None)
                    if run_result.get("status") == "ok":
                        st.toast(f"✅ {run_result['message']}", icon="🔴")
                    else:
                        st.error(run_result.get("message", "執行失敗"))
                    st.rerun()
            with cancel_col:
                if st.button("取消", key="cancel_macro_run", use_container_width=True):
                    st.session_state.pop("_pending_macro_confirm", None)
                    st.rerun()
            st.divider()

        if not macros:
            st.caption("尚無已儲存的巨集。執行一些操作後，可在下方錄製。")
        else:
            st.caption(f"已儲存 {len(macros)} 個巨集：")
            for m in macros:
                col_info, col_run, col_del = st.columns([3, 1, 1])
                with col_info:
                    desc = m.get("description") or ""
                    st.markdown(
                        f"**{m['name']}**　{m['step_count']} 步"
                        + (f"\n\n_{desc}_" if desc else "")
                    )
                with col_run:
                    if st.button("▶", key=f"run_macro_{m['name']}", help=f"執行「{m['name']}」"):
                        run_result = _macro.run_macro(m["name"])
                        if run_result.get("requires_confirmation"):
                            st.session_state["_pending_macro_confirm"] = {
                                "name": run_result["name"],
                                "dangerous_steps": run_result.get("dangerous_steps", []),
                            }
                            st.warning(run_result.get("message", "此巨集包含危險工具，請確認後再執行"))
                        elif run_result.get("status") == "ok":
                            st.toast(f"✅ {run_result['message']}", icon="🔴")
                        else:
                            st.error(run_result.get("message", "執行失敗"))
                        st.rerun()
                with col_del:
                    if st.button("🗑", key=f"del_macro_{m['name']}", help=f"刪除「{m['name']}」"):
                        _macro.delete_macro(m["name"])
                        st.toast(f"已刪除巨集「{m['name']}」", icon="🗑️")
                        st.rerun()

        st.divider()
        st.caption("**錄製新巨集**（從最近操作歷史）")
        macro_name = st.text_input(
            "巨集名稱",
            key="new_macro_name",
            placeholder="如：月報格式化",
            label_visibility="collapsed",
        )
        macro_desc = st.text_input(
            "說明（可省略）",
            key="new_macro_desc",
            placeholder="說明這個巨集的用途",
            label_visibility="collapsed",
        )
        if st.button("🔴 錄製", use_container_width=True, disabled=not macro_name.strip()):
            rec_result = _macro.record_macro(macro_name.strip(), macro_desc.strip())
            if rec_result.get("status") == "ok":
                st.toast(f"✅ {rec_result['message']}", icon="🔴")
                st.rerun()
            else:
                st.error(rec_result.get("message", "錄製失敗"))
