"""excel/__init__.py — re-exports all public symbols."""
from __future__ import annotations

from excel._base import (
    _ensure_positive_int,
    _ensure_positive_number,
    _get_excel,
    _get_sheet,
    _hex_to_bgr,
    _normalize_values,
)

from excel._undo import (
    _undo_dispatch,
    _undo_last_body,
    undo_last,
)

from excel.chart import (
    add_slicer,
    create_chart,
    create_combo_chart,
    create_pivot_table,
    delete_chart,
    format_chart,
    format_pivot_table,
    move_chart,
    refresh_pivot_table,
)

from excel.data import (
    add_comment,
    add_subtotal,
    advanced_filter,
    clear_range,
    copy_range,
    fill_series,
    find_duplicates,
    find_replace,
    get_sheet_info,
    get_used_range,
    get_workbook_summary,
    name_range,
    read_range,
    save_workbook,
    set_data_validation,
    sort_range,
    split_text_to_columns,
    summarize_range,
    transpose_range,
    trim_range,
    write_range,
)

from excel.format import (
    _restore_formats,
    add_conditional_format,
    add_header_footer,
    add_image,
    add_sparklines,
    apply_table_style,
    auto_fit,
    capture_formats_before,
    capture_heights_before,
    capture_widths_before,
    format_range,
    freeze_panes,
    merge_cells,
    page_setup,
    set_borders,
    set_column_width,
    set_print_titles,
    set_row_height,
    set_tab_color,
    unmerge_cells,
)

from excel.sheet import (
    add_sheet,
    copy_range_between_workbooks,
    copy_sheet,
    delete_column,
    delete_row,
    delete_sheet,
    filter_range,
    group_columns,
    group_rows,
    insert_column,
    insert_row,
    list_workbooks,
    move_sheet,
    protect_sheet,
    rename_sheet,
    restore_snapshot,
    snapshot_sheet,
    switch_workbook,
    unprotect_sheet,
)
