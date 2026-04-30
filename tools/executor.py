import json
import time
import excel_tools as et
from tools.registry import get_registered_tool_map
from exceptions import ExcelAIError
from logger import get_logger, hash_args
import backup
import telemetry

_log = get_logger("executor")


TOOL_MAP = {
    # ── V1 工具 ────────────────────────────────────────────────────────────────
    "read_range":             lambda a: et.read_range(**a),
    "write_range":            lambda a: et.write_range(**a),
    "get_sheet_info":         lambda _: et.get_sheet_info(),   # ★ V4 P1-E：無參數工具 lambda 一致化
    "get_used_range":         lambda a: et.get_used_range(**a),
    "format_range":           lambda a: et.format_range(**a),
    "insert_row":             lambda a: et.insert_row(**a),
    "delete_row":             lambda a: et.delete_row(**a),
    "insert_column":          lambda a: et.insert_column(**a),
    "delete_column":          lambda a: et.delete_column(**a),
    "add_sheet":              lambda a: et.add_sheet(**a),
    "rename_sheet":           lambda a: et.rename_sheet(**a),
    "delete_sheet":           lambda a: et.delete_sheet(**a),
    "move_sheet":             lambda a: et.move_sheet(**a),
    "copy_sheet":             lambda a: et.copy_sheet(**a),
    "set_print_titles":       lambda a: et.set_print_titles(**a),
    "add_header_footer":      lambda a: et.add_header_footer(**a),
    "protect_sheet":          lambda a: et.protect_sheet(**a),
    "unprotect_sheet":        lambda a: et.unprotect_sheet(**a),
    "sort_range":             lambda a: et.sort_range(**a),
    "find_replace":           lambda a: et.find_replace(**a),
    "trim_range":             lambda a: et.trim_range(**a),
    "save_workbook":          lambda _: et.save_workbook(),    # ★ V4 P1-E：同上
    # ── V2 新增工具 ────────────────────────────────────────────────────────────
    "delete_chart":           lambda a: et.delete_chart(**a),
    "move_chart":             lambda a: et.move_chart(**a),
    "create_chart":           lambda a: et.create_chart(**a),
    "create_pivot_table":     lambda a: et.create_pivot_table(**a),
    "refresh_pivot_table":    lambda a: et.refresh_pivot_table(**a),
    "format_pivot_table":     lambda a: et.format_pivot_table(**a),
    "freeze_panes":           lambda a: et.freeze_panes(**a),
    "auto_fit":               lambda a: et.auto_fit(**a),
    "set_column_width":       lambda a: et.set_column_width(**a),
    # ── V3 新增工具 ────────────────────────────────────────────────────────────
    "filter_range":           lambda a: et.filter_range(**a),
    "merge_cells":            lambda a: et.merge_cells(**a),
    "unmerge_cells":          lambda a: et.unmerge_cells(**a),
    "set_borders":            lambda a: et.set_borders(**a),
    "clear_range":            lambda a: et.clear_range(**a),
    "set_row_height":         lambda a: et.set_row_height(**a),
    "copy_range":             lambda a: et.copy_range(**a),
    "add_conditional_format": lambda a: et.add_conditional_format(**a),
    "add_comment":            lambda a: et.add_comment(**a),
    "set_data_validation":    lambda a: et.set_data_validation(**a),
    # ── V4 美化工具群 ──────────────────────────────────────────────────────────
    "beautify_range":         lambda a: et.beautify_range(**a),
    "apply_table_style":      lambda a: et.apply_table_style(**a),
    "format_chart":           lambda a: et.format_chart(**a),
    "create_combo_chart":     lambda a: et.create_combo_chart(**a),
    "add_sparklines":         lambda a: et.add_sparklines(**a),
    "set_tab_color":          lambda a: et.set_tab_color(**a),
    "page_setup":             lambda a: et.page_setup(**a),
    "add_slicer":             lambda a: et.add_slicer(**a),
    "add_image":              lambda a: et.add_image(**a),
    # ── V4 分析工具群 ──────────────────────────────────────────────────────────
    "summarize_range":        lambda a: et.summarize_range(**a),
    "find_duplicates":        lambda a: et.find_duplicates(**a),
    "fill_series":            lambda a: et.fill_series(**a),
    "group_rows":             lambda a: et.group_rows(**a),
    "group_columns":          lambda a: et.group_columns(**a),
    "transpose_range":        lambda a: et.transpose_range(**a),
    "name_range":             lambda a: et.name_range(**a),
    "add_subtotal":           lambda a: et.add_subtotal(**a),
    "advanced_filter":        lambda a: et.advanced_filter(**a),
    "split_text_to_columns":  lambda a: et.split_text_to_columns(**a),
    # ── Phase 2：Undo ──────────────────────────────────────────────────────────
    "undo_last":              lambda _: et.undo_last(),
    # ── V4.2 新工具 ────────────────────────────────────────────────────────────
    "get_workbook_summary":   lambda _: et.get_workbook_summary(),
    # ── V4.7.0 A：巨集工具 ────────────────────────────────────────────────────
    "record_macro":           lambda a: __import__("macro").record_macro(**a),
    "list_macros":            lambda _: __import__("macro").list_macros(),
    "run_macro":              lambda a: __import__("macro").run_macro(a.get("name", "")),
    "delete_macro":           lambda a: __import__("macro").delete_macro(**a),
    # ── V4.7.0 B：公式智慧輔助 ────────────────────────────────────────────────
    "validate_formula":       lambda a: __import__("formula_validator").validate_formula_tool(**a),
    "explain_formula":        lambda a: __import__("formula_validator").explain_formula_tool(**a),
    # ── V4.7.0 C：自然語言查詢 ────────────────────────────────────────────────
    "query_range":            lambda a: __import__("excel_query").query_range(**a),
    # ── V4.7.0 D：多工作簿協作 ────────────────────────────────────────────────
    "list_workbooks":                 lambda _: et.list_workbooks(),
    "switch_workbook":                lambda a: et.switch_workbook(**a),
    "copy_range_between_workbooks":   lambda a: et.copy_range_between_workbooks(**a),
}
# Merge registry-registered tools (new tools decorated with @register_tool).
# Registry entries take precedence so they can override legacy lambdas.
TOOL_MAP.update(get_registered_tool_map())

# 執行前需要使用者確認的危險工具
DANGEROUS_TOOLS = {"delete_row", "delete_column", "find_replace", "clear_range", "split_text_to_columns", "delete_sheet"}


def _is_error_payload(payload: object) -> bool:
    return isinstance(payload, dict) and (
        "error" in payload or payload.get("status") == "error"
    )


def _ensure_error_field(payload: dict, tool_name: str) -> dict:
    if "error" in payload:
        return payload
    normalized = dict(payload)
    normalized["error"] = str(
        normalized.get("message") or f"工具「{tool_name}」回傳錯誤狀態"
    )
    return normalized


def execute(tool_name: str, arguments: dict) -> str:
    """
    執行單一工具，回傳 JSON 字串供 LLM 讀取。

    V4 Phase 1 變更：
    - 錯誤 JSON 多附 `error_type` 欄位（例如 "SheetNotFoundError"）供 LLM 依類型
      採取補救動作；既有呼叫端若只讀 `error` 欄位完全不受影響。
    - 每次呼叫會寫入結構化 JSON 日誌（~/.excel-ai/logs/）：tool_name / args_hash /
      duration_ms / status / error_type。參數以 hash 記錄不寫原值。
    - 執行前透過 `backup.capture_before()` 取得 BackupEntry；執行成功後推入
      `session_state["_backup_stack"]`（上限 20 步）。Phase 1 僅累積不啟用 Undo，
      Phase 2 才會新增 undo_last 工具與側邊欄按鈕消費這個 stack。
    """
    arguments = dict(arguments or {})
    confirmed_dangerous = bool(arguments.pop("confirm_dangerous", False))
    args_h = hash_args(arguments)
    start = time.perf_counter()

    if tool_name in DANGEROUS_TOOLS and not confirmed_dangerous:
        duration_ms = int((time.perf_counter() - start) * 1000)
        _log.info("dangerous_tool_confirmation_required", extra={
            "tool": tool_name, "args_hash": args_h,
            "duration_ms": duration_ms, "status": "blocked",
            "error_type": "DangerousToolRequiresConfirmation",
        })
        telemetry.record(tool_name, duration_ms, "error", "DangerousToolRequiresConfirmation")
        return json.dumps(
            {
                "status": "error",
                "error": f"工具「{tool_name}」屬於危險操作，請確認後再執行",
                "error_type": "DangerousToolRequiresConfirmation",
                "requires_confirmation": True,
                "tool": tool_name,
                "arguments": arguments,
            },
            ensure_ascii=False,
        )

    # Phase 2：執行前抓 backup entry（失敗不影響主流程）
    backup_entry = None
    try:
        backup_entry = backup.capture_before(tool_name, arguments)
        # Phase 2 Category B：對資料修改工具預先讀取 values_before
        if backup_entry is not None and tool_name in ("write_range", "clear_range", "trim_range"):
            try:
                rng_addr = arguments.get("range_addr", "")
                sheet = arguments.get("sheet")
                if rng_addr:
                    values_before = et.read_range(rng_addr, sheet)
                    if values_before == []:
                        excel = et._get_excel()
                        ws = et._get_sheet(excel, sheet)
                        rng = ws.Range(rng_addr)
                        if rng.Cells.Count == 1:
                            values_before = [[None]]
                    backup_entry.values_before = values_before
            except Exception as ve:
                _log.warning("backup_values_before_failed", extra={
                    "tool": tool_name, "args_hash": args_h,
                    "error_type": type(ve).__name__,
                })
        # Phase 3 Category B：對格式修改工具預先讀取 formats_before
        if backup_entry is not None and tool_name in ("format_range", "set_borders", "beautify_range"):
            try:
                rng_addr = arguments.get("range_addr", "")
                sheet = arguments.get("sheet")
                if rng_addr:
                    capture_args = dict(arguments)
                    capture_args["_tool_type"] = tool_name
                    backup_entry.formats_before = et.capture_formats_before(
                        rng_addr, sheet, capture_args
                    )
            except Exception as fe:
                _log.warning("backup_formats_before_failed", extra={
                    "tool": tool_name, "args_hash": args_h,
                    "error_type": type(fe).__name__,
                })
        # Phase 4 (TD-03)：欄寬 / 列高 捕捉（set_column_width / set_row_height / auto_fit）
        if backup_entry is not None and tool_name == "set_column_width":
            try:
                col_idx  = arguments.get("column_index", arguments.get("col_index", 1))
                count    = arguments.get("count", 1)
                sheet    = arguments.get("sheet")
                backup_entry.widths_before = et.capture_widths_before(col_idx, count, sheet)
            except Exception:
                pass
        if backup_entry is not None and tool_name == "set_row_height":
            try:
                row_idx  = arguments.get("row_index", 1)
                count    = arguments.get("count", 1)
                sheet    = arguments.get("sheet")
                backup_entry.heights_before = et.capture_heights_before(row_idx, count, sheet)
            except Exception:
                pass
        if backup_entry is not None and tool_name == "auto_fit":
            try:
                sheet = arguments.get("sheet")
                used = et.get_used_range(sheet)
                import re as _re
                m = _re.match(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", used or "")
                if m:
                    from utils import col_index
                    c1, r1 = col_index(m.group(1)), int(m.group(2))
                    c2, r2 = col_index(m.group(3)), int(m.group(4))
                    backup_entry.widths_before  = et.capture_widths_before(c1, c2-c1+1, sheet)
                    backup_entry.heights_before = et.capture_heights_before(r1, r2-r1+1, sheet)
            except Exception:
                pass
    except Exception as e:  # 防禦：capture_before 失誤不該阻斷工具執行
        _log.warning("backup_capture_failed", extra={
            "tool": tool_name, "args_hash": args_h,
            "error_type": type(e).__name__,
        })

    try:
        fn = TOOL_MAP.get(tool_name)
        if fn is None:
            duration_ms = int((time.perf_counter() - start) * 1000)
            _log.warning("tool_unknown", extra={
                "tool": tool_name, "args_hash": args_h,
                "duration_ms": duration_ms, "status": "error",
                "error_type": "UnknownTool",
            })
            return json.dumps(
                {"error": f"未知的 tool：{tool_name}", "error_type": "UnknownTool"},
                ensure_ascii=False,
            )
        result = fn(arguments)
        duration_ms = int((time.perf_counter() - start) * 1000)

        if _is_error_payload(result):
            result = _ensure_error_field(result, tool_name)
            err_type = str(result.get("error_type") or "ToolReturnedError")
            _log.info("tool_failed", extra={
                "tool": tool_name, "args_hash": args_h,
                "duration_ms": duration_ms, "status": "error",
                "error_type": err_type,
            })
            telemetry.record(tool_name, duration_ms, "error", err_type)
            return json.dumps(result, ensure_ascii=False, default=str)

        # Phase 2：執行成功後把 entry 推入 stack
        if backup_entry is not None:
            stack = backup.get_session_stack()
            if stack is not None:
                stack.push(backup_entry)
                backup.save_current_stack()   # v4.6.0: persist after push
                _log.debug("backup_pushed", extra={
                    "tool": tool_name, "args_hash": args_h,
                    "stack_size": len(stack),
                })

        _log.info("tool_executed", extra={
            "tool": tool_name, "args_hash": args_h,
            "duration_ms": duration_ms, "status": "ok",
        })
        telemetry.record(tool_name, duration_ms, "ok")
        return json.dumps(result, ensure_ascii=False, default=str)
    except ExcelAIError as e:
        duration_ms = int((time.perf_counter() - start) * 1000)
        err_type = type(e).__name__
        _log.info("tool_failed", extra={
            "tool": tool_name, "args_hash": args_h,
            "duration_ms": duration_ms, "status": "error",
            "error_type": err_type,
        })
        telemetry.record(tool_name, duration_ms, "error", err_type)
        return json.dumps(
            {"error": str(e), "error_type": err_type},
            ensure_ascii=False,
        )
    except Exception as e:
        duration_ms = int((time.perf_counter() - start) * 1000)
        _log.exception("tool_crashed", extra={
            "tool": tool_name, "args_hash": args_h,
            "duration_ms": duration_ms, "status": "error",
            "error_type": "UnexpectedError",
        })
        telemetry.record(tool_name, duration_ms, "error", "UnexpectedError")
        return json.dumps(
            {"error": str(e), "error_type": "UnexpectedError"},
            ensure_ascii=False,
        )


def execute_batch(steps: list[dict], *, confirm_dangerous: bool = False) -> list[dict]:
    """
    Execute a batch of tool calls; auto-rollback on first error.

    Each step: {"tool": str, "args": dict}
    confirm_dangerous=True is only for already-confirmed batch runners such as
    macro replay; normal callers should let execute() return a confirmation
    request for dangerous tools.
    Returns list of {"tool": str, "result": dict, "rolled_back": bool}
    """
    import json as _json

    results: list[dict] = []
    stack = backup.get_session_stack()
    depth_before = len(stack) if stack is not None else 0

    for step in steps:
        tool_name = step.get("tool", "")
        arguments = dict(step.get("args", {}) or {})
        if confirm_dangerous and tool_name in DANGEROUS_TOOLS:
            arguments["confirm_dangerous"] = True
        raw = execute(tool_name, arguments)
        try:
            parsed = _json.loads(raw)
        except Exception:
            parsed = {"raw": raw}

        if _is_error_payload(parsed):
            parsed = _ensure_error_field(parsed, tool_name)
            if stack is not None:
                while len(stack) > depth_before:
                    entry = stack.pop()
                    if entry is not None:
                        backup.restore(entry)
                backup.save_current_stack()   # v4.6.0: persist after rollback
            results.append({"tool": tool_name, "result": parsed, "rolled_back": True})
            for r in results[:-1]:
                r["rolled_back"] = True
            break
        else:
            results.append({"tool": tool_name, "result": parsed, "rolled_back": False})

    return results
