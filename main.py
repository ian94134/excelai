"""
Excel AI 助手 v4.4.0
執行方式：streamlit run main.py
"""

import importlib
import json
import time
import streamlit as st
from tools import OPENAI_TOOLS, execute, execute_batch, DANGEROUS_TOOLS
import session
import ui.sidebar as sidebar
from logger import get_logger, redact_prompt
import excel_tools as et
import excel_event_watcher
import backup
import agent
from agent import (
    EVT_TEXT_CHUNK, EVT_RETRY, EVT_ASST_MSG, EVT_TOOL_START,
    EVT_TOOL_DONE, EVT_ROLLBACK, EVT_DONE, EVT_PLAN_READY,
    EVT_DANGEROUS, EVT_REPEAT_HALT, EVT_ERROR, EVT_CLARIFY,
)
from utils import col_letter as _col_letter
from ui.tool_display import friendly_tool_label, friendly_tool_status, sanitize_assistant_text
from ui import quick_actions as _quick_actions

_quick_actions = importlib.reload(_quick_actions)
QUICK_ACTIONS = _quick_actions.QUICK_ACTIONS
QUICK_ACTION_FORM_CHOICES = _quick_actions.QUICK_ACTION_FORM_CHOICES
build_quick_action_prompt = _quick_actions.build_quick_action_prompt
clear_quick_action_form = _quick_actions.clear_quick_action_form
get_quick_action = _quick_actions.get_quick_action
open_quick_action_form = _quick_actions.open_quick_action_form
queue_quick_action = _quick_actions.queue_quick_action

VERSION = "v4.8.0"

_log = get_logger("main")


def _build_workbook_context() -> str:
    """
    Try to fetch a one-line workbook summary to prepend to the LLM system
    context.  Fails silently if Excel is not open or the call errors out.
    Returns an empty string when unavailable.
    """
    try:
        summary = et.get_workbook_summary()
        sheet_lines = []
        for s in summary.get("sheets", []):
            hdrs = ", ".join(s.get("sample_headers", [])) or "(no headers)"
            sheet_lines.append(
                f"  - {s['name']}: {s.get('rows', '?')}R x {s.get('columns', '?')}C "
                f"used={s.get('used_range', '?')} headers=[{hdrs}]"
            )
        sheets_str = "\n".join(sheet_lines)
        active = summary.get("active_sheet", "?")
        fname  = summary.get("file_name", "?")
        return (
            f"[Workbook context — {fname}, active: {active}]\n{sheets_str}"
        )
    except Exception:
        return ""


def _inject_ephemeral_system_ctx(msgs: list[dict], ctx: str) -> list[dict]:
    """
    Merge ctx into the existing system message at index 0 (append to content).
    If no system message exists, prepend one.
    Always returns a new list — the original session messages are never mutated.
    Rule: Qwen (and most OpenAI-compatible models) require system role to appear
    exactly once and only at position 0.  Never insert a second system message.
    """
    if not ctx:
        return list(msgs)
    result = list(msgs)
    if result and result[0].get("role") == "system":
        result[0] = dict(result[0])
        result[0]["content"] = result[0]["content"] + "\n\n" + ctx
    else:
        result.insert(0, {"role": "system", "content": ctx})
    return result


def _build_selection_tag() -> str:
    """
    回傳簡潔的選取範圍標記字串（如 [目前選取: Sheet1!A1:D10]），
    供注入 user message 讓 Qwen 直接參照。
    無選取或 Excel 未開啟時回傳空字串。
    """
    try:
        sel = excel_event_watcher.get_current_selection()
        if sel and sel.get("address") and sel["address"] != "(non-range selection)":
            sheet = sel.get("sheet", "")
            addr  = sel.get("address", "")
            label = f"{sheet}!{addr}" if sheet else addr
            return f"[目前選取: {label}]"
        # fallback：同步 COM 讀取（watcher 尚未取得資料時）
        info = et.get_sheet_info()
        sel_addr     = info.get("selection", "")
        active_sheet = info.get("active_sheet", "")
        if sel_addr and sel_addr != "(non-range selection)":
            return f"[目前選取: {active_sheet}!{sel_addr}]"
    except Exception:
        pass
    return ""

# 唯讀工具集合（執行後記錄到 _read_op_log，不進 BackupStack）
_READ_ONLY_TOOLS = {
    "read_range", "get_sheet_info", "get_used_range", "get_workbook_summary",
    "summarize_range", "find_duplicates", "query_range", "list_workbooks",
}

# ── 複雜任務自動偵測 ──────────────────────────────────────────────────────────
import re as _re_complexity

# 多步驟連接詞（出現 ≥ 2 個 → 視為複雜任務）
_CONNECTOR_RE = _re_complexity.compile(
    r"然後|再次?|接著|並且|同時|另外|最後|之後|先.{1,20}再|分別|逐一|依序|第[一二三四五六七八九十]步",
    _re_complexity.UNICODE,
)
# Excel 操作關鍵字（出現 ≥ 3 個 → 視為複雜任務）
_OP_KW_RE = _re_complexity.compile(
    r"篩選|排序|格式化?|合併|刪除|插入|建立|新增|計算|加總|平均|統計|分析|整理|製作|產生|匯出|"
    r"讀取|寫入|複製|移動|凍結|圖表|樞紐|報表|下拉|驗證|保護|解除|條件|框線|欄寬|列高|自動",
    _re_complexity.UNICODE,
)


def _is_complex_task(prompt: str) -> bool:
    """
    回傳 True 表示輸入屬於多步驟複雜任務，應自動啟用規劃模式。
    判斷依據：
      - 多步驟連接詞 ≥ 2 個（然後/再/接著/並且…）
      - Excel 操作關鍵字 ≥ 3 個（篩選/排序/圖表/報表…）
    """
    connectors = len(_CONNECTOR_RE.findall(prompt))
    ops        = len(_OP_KW_RE.findall(prompt))
    return connectors >= 2 or ops >= 3




st.set_page_config(page_title="Excel AI 助手", page_icon="📊", layout="wide")


def _preview_dangerous_op(tool_name: str, arguments: dict) -> None:
    """在危險操作確認視窗中顯示即將受影響的資料預覽。"""
    try:
        sheet = arguments.get("sheet")

        if tool_name in ("delete_row",):
            idx   = arguments.get("index", 1)
            count = arguments.get("count", 1)
            end   = idx + count - 1
            rng   = f"{idx}:{end}" if count > 1 else str(idx)
            st.info(f"🗑️ 即將刪除第 **{rng}** 列，以下是受影響的資料：")
            try:
                # 讀取要刪除的那幾列（最多顯示 20 列）
                read_end = min(end, idx + 19)
                col_end  = "Z"  # 讀到 Z 欄已足夠
                data = et.read_range(f"A{idx}:{col_end}{read_end}", sheet)
                if data and any(any(c for c in row) for row in data):
                    st.dataframe(data, use_container_width=True)
                else:
                    st.caption("（該範圍為空白）")
            except Exception:
                st.caption("（無法讀取預覽）")

        elif tool_name in ("delete_column",):
            idx   = arguments.get("index", 1)
            count = arguments.get("count", 1)
            col_letter = _col_letter(idx)
            end_col    = _col_letter(idx + count - 1)
            st.info(f"🗑️ 即將刪除第 **{idx}** 欄（{col_letter} 欄），以下是受影響的資料（前 10 列）：")
            try:
                data = et.read_range(f"{col_letter}1:{end_col}10", sheet)
                if data and any(any(c for c in row) for row in data):
                    st.dataframe(data, use_container_width=True)
                else:
                    st.caption("（該範圍為空白）")
            except Exception:
                st.caption("（無法讀取預覽）")

        elif tool_name == "delete_sheet":
            name = arguments.get("name", "")
            st.error(f"🗑️ 即將永久刪除工作表「**{name}**」，此操作無法由 undo_last 復原！")
            st.caption("建議先執行 copy_sheet 備份後再刪除。")

        elif tool_name == "clear_range":
            rng = arguments.get("range_addr", "")
            st.info(f"🗑️ 即將清空範圍 **{rng}** 的所有內容，以下是即將被清除的資料：")
            try:
                data = et.read_range(rng, sheet)
                if data and any(any(c for c in row) for row in data):
                    st.dataframe(data, use_container_width=True)
                else:
                    st.caption("（該範圍為空白）")
            except Exception:
                st.caption("（無法讀取預覽）")

        elif tool_name == "find_replace":
            find_text    = arguments.get("find_text", "")
            replace_text = arguments.get("replace_text", "")
            st.info(
                f"🔄 即將把工作表中所有「**{find_text}**」取代為「**{replace_text}**」。\n\n"
                "此操作無法由 undo_last 自動還原，確認前請注意影響範圍。"
            )

        elif tool_name == "split_text_to_columns":
            rng = arguments.get("range_addr", "")
            st.info(
                f"✂️ 即將對 **{rng}** 進行文字分欄，右側欄位的現有資料**可能被覆蓋**。\n\n"
                "此操作無法由 undo_last 自動還原。"
            )

    except Exception:
        pass  # 預覽失敗不影響確認流程


def _render_tool_result(tool_name: str, result_json: str) -> None:
    """
    美化工具執行結果的呈現方式：
    - read_range / summarize_range → st.dataframe
    - get_sheet_info / get_used_range → 精簡文字
    - undo_last → 彩色提示
    - 其他 → JSON（保持原樣）
    """
    try:
        payload = json.loads(result_json)
    except Exception:
        st.code(result_json, language="json")
        return

    # ── read_range：二維陣列 → 表格 ───────────────────────────────────────────
    if tool_name == "read_range":
        if isinstance(payload, list) and payload:
            import pandas as pd
            try:
                # 第一列視為標題（若全為字串），否則產生自動欄名
                if all(isinstance(v, str) for v in payload[0]):
                    df = pd.DataFrame(payload[1:], columns=payload[0])
                else:
                    df = pd.DataFrame(payload)
                st.dataframe(df, use_container_width=True, hide_index=True)
                st.caption(f"共 {len(df)} 列 × {len(df.columns)} 欄")
                return
            except Exception:
                pass
        st.code(result_json, language="json")
        return

    # ── summarize_range：統計結果 → 指標卡片 ─────────────────────────────────
    if tool_name == "summarize_range":
        if isinstance(payload, dict) and "sum" in payload:
            cols = st.columns(5)
            labels = [("合計", "sum"), ("平均", "avg"), ("最大", "max"),
                      ("最小", "min"), ("筆數", "count")]
            for col, (label, key) in zip(cols, labels):
                val = payload.get(key)
                if val is not None:
                    col.metric(label, f"{val:,.2f}" if isinstance(val, float) else val)
            return

    # ── get_sheet_info：精簡顯示 ─────────────────────────────────────────────
    if tool_name == "get_sheet_info" and isinstance(payload, dict):
        st.caption(
            f"📄 **{payload.get('file_name', '')}**　"
            f"作用中：`{payload.get('active_sheet', '')}`　"
            f"共 {len(payload.get('sheets', []))} 張工作表"
        )
        return

    # ── get_used_range：精簡顯示 ─────────────────────────────────────────────
    if tool_name == "get_used_range" and isinstance(payload, str):
        st.caption(f"📐 已使用範圍：`{payload}`")
        return

    # ── undo_last：狀態提示 ───────────────────────────────────────────────────
    if tool_name == "undo_last" and isinstance(payload, dict):
        status = payload.get("status")
        if status == "ok":
            st.success(f"↶ 已還原：{payload.get('undone', '')}")
        elif status == "no_op":
            st.info(payload.get("message", ""))
        elif status == "cannot_undo":
            st.warning(payload.get("message", ""))
        return

    if tool_name == "beautify_range" and isinstance(payload, dict) and payload.get("status") == "ok":
        st.caption(
            f"已美化 `{payload.get('sheet', '')}!{payload.get('range', '')}`，"
            f"主題：`{payload.get('theme', 'blue')}`。"
        )
        return

    if tool_name == "write_range" and isinstance(payload, dict) and payload.get("status") == "ok":
        target = payload.get("range") or payload.get("range_addr") or ""
        st.caption(f"已寫入 `{target}`。")
        return

    if tool_name == "query_range" and isinstance(payload, dict):
        count = payload.get("filtered_count") or payload.get("count")
        aggregation = payload.get("aggregation_result") or payload.get("aggregation")
        parts = []
        if count is not None:
            parts.append(f"符合條件：`{count}` 筆")
        if aggregation is not None:
            parts.append(f"彙總結果：`{aggregation}`")
        if parts:
            st.caption("　".join(parts))
            return

    # ── 其他工具：檢查是否有 error，否則精簡 ok 狀態 ─────────────────────────
    if isinstance(payload, dict):
        if "error" in payload:
            st.error(f"❌ {payload['error']}")
            return
        if payload.get("status") == "ok":
            # 只保留非 status 欄位做簡短顯示
            extra = {
                k: v
                for k, v in payload.items()
                if k not in ("status", "tool") and v is not None
            }
            if extra:
                st.caption("　".join(f"`{k}` = `{v}`" for k, v in extra.items()))
            else:
                st.caption("✓ 執行成功")
            return

    # 其他情況 fallback
    st.code(result_json, language="json")

def _render_diff(tool_name: str, entry) -> None:
    """
    在 st.status 區塊內顯示 Before/After diff。
    - values_before（write_range / clear_range / trim_range）→ 儲存格前後對照表
    - formats_before（format_range / set_borders）→ 格式色彩摘要
    """
    import pandas as pd

    has_values  = entry is not None and entry.values_before is not None
    has_formats = entry is not None and entry.formats_before is not None

    if not has_values and not has_formats:
        return

    with st.expander("📋 變更前後", expanded=False):

        # ── 資料 diff（write_range / clear_range / trim_range）──────────────
        if has_values:
            before_rows = entry.values_before  # list[list]
            args = entry.arguments or {}

            if tool_name == "write_range":
                after_rows = args.get("data", [])
            elif tool_name == "clear_range":
                # 清空後每格為空
                after_rows = [
                    ["" for _ in row] for row in before_rows
                ]
            else:
                # trim_range：無法事先知道 after，僅顯示 before
                after_rows = None

            rng_addr = args.get("range_addr", "")
            st.caption(f"範圍：`{rng_addr}`")

            if after_rows is not None:
                # 把 before / after 並排顯示
                rows_out = []
                max_rows = max(len(before_rows), len(after_rows))
                for r in range(max_rows):
                    b_row = before_rows[r] if r < len(before_rows) else []
                    a_row = after_rows[r]   if r < len(after_rows)  else []
                    max_cols = max(len(b_row), len(a_row))
                    for c in range(max_cols):
                        b_val = b_row[c] if c < len(b_row) else ""
                        a_val = a_row[c] if c < len(a_row) else ""
                        # 只列出有變化的儲存格
                        if str(b_val) != str(a_val):
                            col_letter = _col_letter(c + 1)
                            # 嘗試從 range_addr 解析起始列（簡單解析 A1 格式）
                            import re as _re
                            start_row = 1
                            m = _re.match(r"[A-Z]+(\d+)", rng_addr.upper())
                            if m:
                                start_row = int(m.group(1))
                            cell = f"{col_letter}{start_row + r}"
                            rows_out.append({
                                "儲存格": cell,
                                "變更前": b_val,
                                "變更後": a_val,
                            })
                if rows_out:
                    df = pd.DataFrame(rows_out)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.caption("值未發生變化")
            else:
                # trim_range：僅顯示 before
                st.caption("ℹ️ trim_range 僅顯示操作前的值，空白已被修剪")
                try:
                    df = pd.DataFrame(before_rows)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                except Exception:
                    st.code(str(before_rows))

        # ── 格式 diff（format_range / set_borders）──────────────────────────
        if has_formats:
            fmt = entry.formats_before  # dict[cell_addr, dict]
            args = entry.arguments or {}
            rng_addr = args.get("range_addr", "")
            st.caption(f"格式變更範圍：`{rng_addr}`（顯示變更前的格式）")

            rows_out = []
            for cell_addr, props in list(fmt.items())[:30]:  # 最多顯示 30 格
                fill = props.get("fill_color") or props.get("bg_color")
                font_color = props.get("font_color") or props.get("color")
                bold = props.get("bold")
                row = {"儲存格": cell_addr}
                if fill:
                    # 把 BGR int 轉為 #RRGGBB（Excel COM 用 BGR 整數）
                    try:
                        bgr = int(fill)
                        r_c = (bgr & 0xFF0000) >> 16
                        g_c = (bgr & 0x00FF00) >> 8
                        b_c =  bgr & 0x0000FF
                        hex_color = f"#{b_c:02X}{g_c:02X}{r_c:02X}"
                        row["背景色（前）"] = hex_color
                    except Exception:
                        row["背景色（前）"] = str(fill)
                if font_color:
                    try:
                        bgr = int(font_color)
                        r_c = (bgr & 0xFF0000) >> 16
                        g_c = (bgr & 0x00FF00) >> 8
                        b_c =  bgr & 0x0000FF
                        hex_fc = f"#{b_c:02X}{g_c:02X}{r_c:02X}"
                        row["字色（前）"] = hex_fc
                    except Exception:
                        row["字色（前）"] = str(font_color)
                if bold is not None:
                    row["粗體（前）"] = "✓" if bold else "✗"
                rows_out.append(row)

            if rows_out:
                df = pd.DataFrame(rows_out)
                st.dataframe(df, use_container_width=True, hide_index=True)
                if len(fmt) > 30:
                    st.caption(f"（僅顯示前 30 格，共 {len(fmt)} 格有格式記錄）")
            else:
                st.caption("無格式備份資料")


def _choice_index(choices, value) -> int:
    try:
        return list(choices).index(value)
    except ValueError:
        return 0


def _render_quick_action_panel(action_key: str) -> None:
    try:
        action = get_quick_action(action_key)
    except KeyError:
        clear_quick_action_form(st.session_state)
        return

    choices = QUICK_ACTION_FORM_CHOICES.get(action.key, {})
    options = {}

    with st.form(f"quick_action_form_{action.key}"):
        st.markdown(f"**{action.label}**")

        if action.key == "beautify_report":
            theme_choices = choices["theme"]
            options["theme"] = st.selectbox("主題", theme_choices, index=_choice_index(theme_choices, "藍色"))
            options["freeze_header"] = st.checkbox("凍結表頭", value=True)
            options["save_after"] = st.checkbox("完成後儲存", value=False)

        elif action.key == "summarize_data":
            depth_choices = choices["depth"]
            options["depth"] = st.radio("摘要深度", depth_choices, index=_choice_index(depth_choices, "標準"), horizontal=True)
            options["include_recommendations"] = st.checkbox("包含下一步建議", value=True)
            options["write_to_report"] = st.checkbox("把摘要寫入 Report 工作表", value=False)
            if options["write_to_report"]:
                options["save_after"] = st.checkbox("寫入後儲存檔案", value=False)
            else:
                options["save_after"] = False

        elif action.key == "sum_by_group":
            group_choices = choices["group_by"]
            value_choices = choices["value_col"]
            options["group_by"] = st.selectbox("分組欄位", group_choices, index=0)
            options["value_col"] = st.selectbox("加總欄位", value_choices, index=0)
            options["include_total"] = st.checkbox("顯示總計", value=True)

        elif action.key == "create_report_chart":
            chart_choices = choices["chart_type"]
            placement_choices = choices["placement"]
            options["chart_type"] = st.selectbox("圖表類型", chart_choices, index=0)
            options["placement"] = st.selectbox("放置位置", placement_choices, index=0)
            options["include_title"] = st.checkbox("加上清楚標題", value=True)

        elif action.key == "undo_last":
            st.warning("按下執行會復原上一個可復原操作。")

        preview_prompt = build_quick_action_prompt(action.key, options)
        st.text_area("將送出的任務", value=preview_prompt, height=110, disabled=True)
        st.caption("確認後會以白話任務送出，並沿用原本安全、備份與復原流程。")

        run_col, cancel_col = st.columns(2)
        with run_col:
            submitted = st.form_submit_button("執行", type="primary", use_container_width=True)
        with cancel_col:
            cancelled = st.form_submit_button("取消", use_container_width=True)

    if submitted:
        queue_quick_action(st.session_state, action.key, options)
        clear_quick_action_form(st.session_state)
        st.rerun()
    if cancelled:
        clear_quick_action_form(st.session_state)
        st.rerun()


# 每次 Streamlit rerun 都會重入此檔，session_state 首次才 log 啟動
if "_session_logged" not in st.session_state:
    _log.info("session_started", extra={"version": VERSION})
    st.session_state["_session_logged"] = True

# 啟動 Excel 選取範圍背景監聽（只啟動一次）
try:
    excel_event_watcher.start_watcher()
except Exception:
    pass  # win32com 不可用時（如 Linux CI）靜默略過

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
    pending_label = friendly_tool_label(tc["name"])

    st.warning(f"⚠️ **危險操作確認**\n\nAI 想要執行：**{pending_label}**")

    # 依工具類型顯示「即將受影響的資料」預覽
    _preview_dangerous_op(tc["name"], tc["arguments"])

    with st.expander("📋 技術細節", expanded=False):
        st.caption(f"工具：`{tc['name']}`")
        st.code(json.dumps(tc["arguments"], ensure_ascii=False, indent=2), language="json")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ 確認執行", use_container_width=True, type="primary"):
            confirmed_args = dict(tc["arguments"])
            confirmed_args["confirm_dangerous"] = True
            result = execute(tc["name"], confirmed_args)
            session.append_message({
                "role": "tool", "tool_call_id": tc["id"],
                "name": tc["name"], "content": result,
            })
            try:
                payload = json.loads(result)
            except Exception:
                payload = {}
            if isinstance(payload, dict) and (
                payload.get("status") == "error" or "error" in payload
            ):
                summary = f"❌ 已確認但「{pending_label}」執行失敗：{payload.get('message') or payload.get('error') or result}"
            else:
                summary = f"✅ 已確認並執行「{pending_label}」。"
            session.append_message({"role": "assistant", "content": summary})
            st.session_state.pop("_pending_confirm")
            st.rerun()
    with col2:
        if st.button("❌ 取消", use_container_width=True):
            session.append_message({
                "role": "tool", "tool_call_id": tc["id"],
                "name": tc["name"],
                "content": json.dumps({"error": "使用者取消執行"}, ensure_ascii=False),
            })
            session.append_message({"role": "assistant", "content": f"已取消「{pending_label}」。"})
            st.session_state.pop("_pending_confirm")
            st.rerun()
    st.stop()


# ── 使用者輸入 ────────────────────────────────────────────────────────────────
# ── 任務規劃模式：等待使用者確認計劃後開始執行 ──────────────────────────────
if "_pending_plan" in st.session_state:
    plan_data = st.session_state["_pending_plan"]
    st.info(
        "📋 **任務計劃已生成**\n\n"
        "請確認上方的執行步驟，然後點擊「▶ 開始執行」讓 AI 逐步完成任務。\n"
        "若有需要修改，請直接在對話框輸入補充說明。"
    )
    col_go, col_cancel = st.columns(2)
    with col_go:
        if st.button("▶ 開始執行", use_container_width=True, type="primary"):
            # 注入執行指令，讓下一輪進入與 chat_input 相同的執行流程
            exec_prompt = (
                "好的，計劃確認。請現在依照上面的步驟逐一執行，"
                "每一步完成後告知結果，遇到問題立即停下來說明。"
            )
            st.session_state["_queued_prompt"] = exec_prompt
            st.session_state.pop("_pending_plan")
            st.rerun()
    with col_cancel:
        if st.button("✏️ 修改計劃", use_container_width=True):
            st.session_state.pop("_pending_plan")
            st.rerun()
    st.stop()

queued_prompt = st.session_state.pop("_queued_prompt", None)
if queued_prompt is None:
    st.caption("常用任務")
    quick_cols = st.columns(len(QUICK_ACTIONS))
    for quick_col, action in zip(quick_cols, QUICK_ACTIONS):
        with quick_col:
            if st.button(action.label, key=f"quick_action_{action.key}", use_container_width=True):
                open_quick_action_form(st.session_state, action.key)
                st.rerun()
    quick_action_key = st.session_state.get("_quick_action_form")
    if quick_action_key:
        _render_quick_action_panel(quick_action_key)

prompt = queued_prompt or st.chat_input("例：篩選台北的資料 / 合併 A1:D1 / 設定外框線 / 建立下拉選單")

if prompt:
    _log.info("user_prompt", extra={
        "prompt_preview": redact_prompt(prompt),
        "prompt_len": len(prompt),
    })

    # ── 選取範圍感知（v4.6.0）─────────────────────────────────────────────────
    # 注入格式：[目前選取: Sheet1!A1:D10]，讓 Qwen 可直接參照範圍
    # 優先使用 watcher 即時資料（背景執行緒），fallback 到同步 COM 讀取
    selection_tag = _build_selection_tag()
    enriched_prompt = (selection_tag + " " + prompt) if selection_tag else prompt
    session.append_message({"role": "user", "content": enriched_prompt})
    with st.chat_message("user"):
        if selection_tag:
            st.caption(f"🎯 {selection_tag}")
        st.markdown(prompt)

    provider = st.session_state.get("_provider")
    if provider is None:
        st.error("Qwen 模型未連接，請在側邊欄設定後重試")
        st.stop()

    # ── 任務規劃模式：手動開啟 或 自動偵測複雜任務 ───────────────────────────
    plan_mode_active = st.session_state.get("_plan_mode", False)
    _auto_plan_triggered = False
    plan_inject = ""

    _PLAN_INJECT_TEXT = (
        "\n\n[任務規劃模式已啟用] 請先用繁體中文條列出完成此任務的所有步驟（1. 2. 3. ...），"
        "說明每一步要用哪個工具、操作哪個範圍、預期結果。"
        "暫時不要呼叫任何工具，等待使用者確認計劃後再執行。"
    )

    if plan_mode_active:
        plan_inject = _PLAN_INJECT_TEXT
    elif _is_complex_task(prompt):
        # 自動偵測到多步驟任務 → 靜默啟用規劃模式
        plan_inject = _PLAN_INJECT_TEXT
        _auto_plan_triggered = True

    # 上下文自動摘要（超過閾值時）
    session.maybe_summarize(provider)

    with st.chat_message("assistant"):
        placeholder = st.empty()
        tool_output_area = st.container()
        placeholder.markdown("⏳ 思考中…")
        last_stream_render_at = 0.0
        last_stream_text = ""
        backup_stack_changed = False

        # 自動規劃模式提示（使用者未手動開啟但被偵測觸發）
        if _auto_plan_triggered:
            st.caption("🧭 偵測到多步驟任務，已自動啟用規劃模式。")

        # ── 呼叫 agent.run_turn() ─────────────────────────────────────────
        for kind, data in agent.run_turn(
            get_messages=session.get_messages,
            tools=OPENAI_TOOLS,
            provider=provider,
            dangerous_tools=DANGEROUS_TOOLS,
            max_iterations=100,
            plan_inject=plan_inject,
            wb_context_fn=_build_workbook_context,
            inject_ctx_fn=_inject_ephemeral_system_ctx,
        ):
            if kind == EVT_TEXT_CHUNK:
                # 串流中以純文字+節流更新，done 事件再用 markdown 呈現
                stream_text = str(data)
                now = time.perf_counter()
                should_render = (
                    (now - last_stream_render_at) >= 0.05
                    or (len(stream_text) - len(last_stream_text)) >= 80
                )
                if should_render:
                    placeholder.text(stream_text + "▌")
                    last_stream_render_at = now
                    last_stream_text = stream_text

            elif kind == EVT_RETRY:
                attempt = int(data)
                placeholder.warning(f"🔄 Qwen 連線失敗，第 {attempt}/3 次重試中…")

            elif kind == EVT_DONE:
                final = sanitize_assistant_text(str(data)) if data else "完成 ✓"
                placeholder.markdown(final)
                session.append_message({"role": "assistant", "content": final})

            elif kind == EVT_PLAN_READY:
                # 規劃模式：顯示計劃，等待使用者確認後執行
                final = sanitize_assistant_text(str(data)) if data else ""
                placeholder.markdown(final)
                session.append_message({"role": "assistant", "content": final})
                st.session_state["_pending_plan"] = {"plan": final}
                st.rerun()

            elif kind == EVT_ASST_MSG:
                # 含 tool_calls 的 assistant message 立即寫入 session
                session.append_message(data)

            elif kind == EVT_TOOL_START:
                # 保留給未來擴充（進度條等）
                pass

            elif kind == EVT_TOOL_DONE:
                tex = data   # agent.ToolExecution
                tc  = tex.tc
                tool_label = friendly_tool_label(tc.name)
                status_label = friendly_tool_status(tc.name, has_error=tex.has_error)

                # 可收折的 status 元件顯示工具執行結果
                with tool_output_area:
                    with st.status(f"🔧 {tool_label}", expanded=False) as status:
                        with st.expander("技術細節", expanded=False):
                            st.caption(f"工具：`{tc.name}`")
                            st.caption("參數")
                            st.code(
                                json.dumps(tc.arguments, ensure_ascii=False, indent=2),
                                language="json",
                            )
                            st.caption("結果")
                            st.code(tex.result_json, language="json")
                        _render_tool_result(tc.name, tex.result_json)
                        if tex.backup_entry is not None:
                            _render_diff(tc.name, tex.backup_entry)
                        status.update(
                            label=f"{'❌' if tex.has_error else '✅'} {status_label}",
                            state="complete",
                            expanded=False,
                        )

                # 追加工具結果訊息至 session
                session.append_message({
                    "role": "tool", "tool_call_id": tc.id,
                    "name": tc.name, "content": tex.result_json,
                })
                if tex.backup_entry is not None:
                    backup_stack_changed = True

                # 唯讀工具另外記錄到 _read_op_log
                if tc.name in _READ_ONLY_TOOLS:
                    log = st.session_state.setdefault("_read_op_log", [])
                    log.append({
                        "time": time.strftime("%H:%M:%S"),
                        "tool": tc.name,
                        "label": friendly_tool_label(tc.name),
                    })
                    if len(log) > 50:
                        st.session_state["_read_op_log"] = log[-50:]

            elif kind == EVT_ROLLBACK:
                rolled = int(data)
                st.warning(f"⚠️ 發生錯誤，已自動回滾本輪前 {rolled} 個成功步驟。")

            elif kind == EVT_DANGEROUS:
                tc = data
                st.session_state["_pending_confirm"] = {"tool_call": {
                    "id": tc.id, "name": tc.name, "arguments": tc.arguments,
                }}
                placeholder.empty()
                st.rerun()

            elif kind == EVT_REPEAT_HALT:
                msg = sanitize_assistant_text(str(data))
                placeholder.warning(msg)
                session.append_message({"role": "assistant", "content": msg})

            elif kind == EVT_CLARIFY:
                # LLM is asking a clarification question — show it and wait for user reply
                question = sanitize_assistant_text(str(data)) if data else ""
                placeholder.markdown(question)
                session.append_message({"role": "assistant", "content": question})
                # 顯示提示標籤，讓使用者知道 AI 在等待補充說明
                st.caption("💬 AI 需要更多資訊才能繼續，請在下方對話框補充說明。")

            elif kind == EVT_ERROR:
                exc = data
                placeholder.error(f"執行錯誤：{exc}")
                _log.exception("main_loop_error", extra={"error_type": type(exc).__name__})

        if backup_stack_changed:
            st.rerun()
