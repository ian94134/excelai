import json
import excel_tools as et

TOOL_MAP = {
    # ── V1 工具 ────────────────────────────────────────────────────────────────
    "read_range":             lambda a: et.read_range(**a),
    "write_range":            lambda a: et.write_range(**a),
    "get_sheet_info":         lambda a: et.get_sheet_info(),
    "get_used_range":         lambda a: et.get_used_range(**a),
    "format_range":           lambda a: et.format_range(**a),
    "insert_row":             lambda a: et.insert_row(**a),
    "delete_row":             lambda a: et.delete_row(**a),
    "insert_column":          lambda a: et.insert_column(**a),
    "delete_column":          lambda a: et.delete_column(**a),
    "add_sheet":              lambda a: et.add_sheet(**a),
    "rename_sheet":           lambda a: et.rename_sheet(**a),
    "sort_range":             lambda a: et.sort_range(**a),
    "find_replace":           lambda a: et.find_replace(**a),
    "save_workbook":          lambda a: et.save_workbook(),
    # ── V2 新增工具 ────────────────────────────────────────────────────────────
    "create_chart":           lambda a: et.create_chart(**a),
    "create_pivot_table":     lambda a: et.create_pivot_table(**a),
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
    "set_data_validation":    lambda a: et.set_data_validation(**a),
}

# 執行前需要使用者確認的危險工具
DANGEROUS_TOOLS = {"delete_row", "delete_column", "find_replace", "clear_range"}


def execute(tool_name: str, arguments: dict) -> str:
    try:
        fn = TOOL_MAP.get(tool_name)
        if fn is None:
            return json.dumps({"error": f"未知的 tool：{tool_name}"}, ensure_ascii=False)
        result = fn(arguments)
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)
