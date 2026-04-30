"""User-facing labels for internal Excel tool names."""
from __future__ import annotations

import re


FRIENDLY_TOOL_LABELS: dict[str, str] = {
    "read_range": "已讀取資料",
    "get_sheet_info": "已確認 Excel 狀態",
    "get_used_range": "已確認資料範圍",
    "get_workbook_summary": "已整理活頁簿摘要",
    "write_range": "已寫入 Excel",
    "save_workbook": "已儲存檔案",
    "format_range": "已套用格式",
    "set_borders": "已設定框線",
    "add_conditional_format": "已設定條件格式",
    "merge_cells": "已合併儲存格",
    "unmerge_cells": "已取消合併儲存格",
    "clear_range": "已清除範圍",
    "insert_row": "已插入列",
    "delete_row": "已刪除列",
    "insert_column": "已插入欄",
    "delete_column": "已刪除欄",
    "set_row_height": "已調整列高",
    "add_sheet": "已新增工作表",
    "delete_sheet": "已刪除工作表",
    "move_sheet": "已移動工作表",
    "copy_sheet": "已複製工作表",
    "set_print_titles": "已設定列印標題",
    "add_header_footer": "已設定頁首頁尾",
    "protect_sheet": "已保護工作表",
    "unprotect_sheet": "已解除工作表保護",
    "rename_sheet": "已重新命名工作表",
    "sort_range": "已排序資料",
    "find_replace": "已尋找並取代",
    "trim_range": "已清理空白",
    "filter_range": "已篩選資料",
    "copy_range": "已複製範圍",
    "add_comment": "已新增註解",
    "set_data_validation": "已設定下拉選單",
    "delete_chart": "已刪除圖表",
    "move_chart": "已移動圖表",
    "create_chart": "已建立圖表",
    "create_pivot_table": "已建立樞紐分析表",
    "refresh_pivot_table": "已刷新樞紐分析表",
    "format_pivot_table": "已美化樞紐分析表",
    "freeze_panes": "已凍結窗格",
    "auto_fit": "已自動調整欄寬列高",
    "set_column_width": "已調整欄寬",
    "beautify_range": "已美化表格",
    "apply_table_style": "已轉成正式表格",
    "format_chart": "已美化圖表",
    "create_combo_chart": "已建立組合圖",
    "add_sparklines": "已新增走勢圖",
    "set_tab_color": "已設定工作表標籤色",
    "page_setup": "已設定頁面格式",
    "add_slicer": "已新增交叉分析篩選器",
    "summarize_range": "已產生統計摘要",
    "find_duplicates": "已檢查重複資料",
    "fill_series": "已填入數列",
    "group_rows": "已群組列",
    "group_columns": "已群組欄",
    "transpose_range": "已轉置資料",
    "name_range": "已建立範圍名稱",
    "add_subtotal": "已新增小計",
    "advanced_filter": "已完成進階篩選",
    "split_text_to_columns": "已完成文字分欄",
    "add_image": "已插入圖片",
    "undo_last": "已復原上一步",
    "record_macro": "已錄製巨集",
    "list_macros": "已列出巨集",
    "run_macro": "已執行巨集",
    "delete_macro": "已刪除巨集",
    "validate_formula": "已檢查公式",
    "explain_formula": "已說明公式",
    "query_range": "已查詢資料",
    "list_workbooks": "已列出活頁簿",
    "switch_workbook": "已切換活頁簿",
    "copy_range_between_workbooks": "已跨活頁簿複製資料",
}


def friendly_tool_label(tool_name: str | None) -> str:
    if not tool_name:
        return "已完成操作"
    return FRIENDLY_TOOL_LABELS.get(tool_name, f"已執行操作：{tool_name.replace('_', ' ')}")


def friendly_tool_status(tool_name: str | None, *, has_error: bool = False) -> str:
    prefix = "執行失敗：" if has_error else ""
    return prefix + friendly_tool_label(tool_name)


_STANDALONE_NOISE = {"check", "smart_toy"}
_TOOL_STATUS_RE = re.compile(r"^(?P<mark>[✅✔☑✓\-*]\s*)?`?(?P<name>[a-z][a-z0-9_]+)`?[。.!！]?$")


def sanitize_assistant_text(text: str) -> str:
    """
    Convert model-generated technical tool checklists into plain user-facing text.

    The raw tool messages still stay in session context; this is only for what
    users read in the chat transcript.
    """
    if not text:
        return text

    cleaned_lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.lower() in _STANDALONE_NOISE:
            continue

        match = _TOOL_STATUS_RE.match(stripped)
        if match and match.group("name") in FRIENDLY_TOOL_LABELS:
            mark = match.group("mark") or "✅ "
            cleaned_lines.append(f"{mark}{FRIENDLY_TOOL_LABELS[match.group('name')]}")
            continue

        cleaned_lines.append(_replace_inline_tool_names(line))

    return "\n".join(cleaned_lines).strip()


def _replace_inline_tool_names(text: str) -> str:
    for name, label in sorted(FRIENDLY_TOOL_LABELS.items(), key=lambda item: len(item[0]), reverse=True):
        text = re.sub(rf"`{re.escape(name)}`", label, text)
        text = re.sub(rf"\b{re.escape(name)}\b", label, text)
    return text
