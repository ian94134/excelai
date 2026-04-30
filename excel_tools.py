"""
excel_tools.py — backward-compatibility shim.

All implementation has moved to the excel/ sub-package:
  excel/_base.py   — shared COM helpers (_get_excel, _get_sheet, etc.)
  excel/data.py    — read/write, sort, trim, search, analysis tools
  excel/format.py  — format_range, borders, merge, sizing, print setup
  excel/sheet.py   — row/col/sheet management, workbook helpers
  excel/chart.py   — charts, pivot tables, slicers
  excel/_undo.py   — undo_last, _undo_dispatch, _undo_last_body

Importing this module provides the same public API as before.
"""
from excel._base import (
    _get_excel, _get_sheet, _hex_to_bgr, _normalize_values,
    _ensure_positive_int, _ensure_positive_number,
    _com_tls,
)
from excel.data import (
    read_range, get_sheet_info, get_used_range, get_workbook_summary,
    write_range, save_workbook,
    sort_range, find_replace, trim_range, clear_range, copy_range,
    add_comment, set_data_validation, name_range,
    summarize_range, find_duplicates, fill_series, advanced_filter,
    split_text_to_columns, add_subtotal, transpose_range,
)
from excel.format import (
    capture_widths_before, capture_heights_before,
    capture_formats_before, _restore_formats,
    format_range, set_borders, add_conditional_format,
    merge_cells, unmerge_cells,
    freeze_panes, auto_fit, set_column_width, set_row_height,
    beautify_range, apply_table_style, set_tab_color, add_image, add_sparklines,
    set_print_titles, add_header_footer, page_setup,
)
from excel.sheet import (
    insert_row, delete_row, insert_column, delete_column,
    add_sheet, rename_sheet, delete_sheet, move_sheet, copy_sheet,
    protect_sheet, unprotect_sheet,
    filter_range, group_rows, group_columns,
    list_workbooks, switch_workbook, snapshot_sheet, restore_snapshot,
    copy_range_between_workbooks,
)
from excel.chart import (
    create_chart, delete_chart, move_chart, format_chart, create_combo_chart,
    create_pivot_table, refresh_pivot_table, format_pivot_table,
    add_slicer,
)
from excel._undo import (
    _undo_dispatch, undo_last, _undo_last_body,
)

__all__ = [
    # _base
    "_get_excel", "_get_sheet", "_com_tls",
    # data
    "read_range", "get_sheet_info", "get_used_range", "get_workbook_summary",
    "write_range", "save_workbook",
    "sort_range", "find_replace", "trim_range", "clear_range", "copy_range",
    "add_comment", "set_data_validation", "name_range",
    "summarize_range", "find_duplicates", "fill_series", "advanced_filter",
    "split_text_to_columns", "add_subtotal", "transpose_range",
    # format
    "capture_widths_before", "capture_heights_before",
    "capture_formats_before", "_restore_formats",
    "format_range", "set_borders", "add_conditional_format",
    "merge_cells", "unmerge_cells",
    "freeze_panes", "auto_fit", "set_column_width", "set_row_height",
    "beautify_range", "apply_table_style", "set_tab_color", "add_image", "add_sparklines",
    "set_print_titles", "add_header_footer", "page_setup",
    # sheet
    "insert_row", "delete_row", "insert_column", "delete_column",
    "add_sheet", "rename_sheet", "delete_sheet", "move_sheet", "copy_sheet",
    "protect_sheet", "unprotect_sheet",
    "filter_range", "group_rows", "group_columns",
    "list_workbooks", "switch_workbook", "snapshot_sheet", "restore_snapshot",
    "copy_range_between_workbooks",
    # chart
    "create_chart", "delete_chart", "move_chart", "format_chart", "create_combo_chart",
    "create_pivot_table", "refresh_pivot_table", "format_pivot_table",
    "add_slicer",
    # undo
    "_undo_dispatch", "undo_last", "_undo_last_body",
]
