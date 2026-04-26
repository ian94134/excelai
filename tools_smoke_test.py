from __future__ import annotations

import argparse
import base64
import json
import os
import tempfile
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

import excel.chart as chart_mod
import excel.data as data_mod
import excel.format as format_mod
import excel.sheet as sheet_mod
import excel._undo as undo_mod
import macro
from tools.executor import DANGEROUS_TOOLS, execute


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR_ENV = "EXCEL_AI_SMOKE_OUTPUT_DIR"

ONE_PIXEL_PNG = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
    "/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
)


def _json(raw: str) -> Any:
    try:
        return json.loads(raw)
    except Exception:
        return {"raw": raw}


def _is_error_response(payload: Any) -> bool:
    return isinstance(payload, dict) and (
        "error" in payload or payload.get("status") == "error"
    )


def _call(name: str, args: dict[str, Any] | None = None) -> Any:
    call_args = dict(args or {})
    if name in DANGEROUS_TOOLS:
        call_args.setdefault("confirm_dangerous", True)
    parsed = _json(execute(name, call_args))
    return parsed


def _patch_excel(app) -> None:
    for mod in (data_mod, format_mod, sheet_mod, chart_mod, undo_mod):
        mod._get_excel = lambda app=app: app


def _setup_workbook(app, name: str = "tool_smoke"):
    wb = app.Workbooks.Add()
    while wb.Worksheets.Count < 3:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    wb.Worksheets(1).Name = "Data"
    wb.Worksheets(2).Name = "Aux"
    wb.Worksheets(3).Name = "Criteria"
    ws = wb.Worksheets("Data")
    ws.Activate()
    ws.Range("A1:H1").Value = [["Name", "Region", "Amount", "Qty", "Date", "Category", "Text", "Parts"]]
    rows = [
        ["Alpha", "North", 120, 2, "2026-01-01", "A", "  hello   world  ", "A,B,C"],
        ["Beta", "South", 80, 1, "2026-01-02", "B", "text", "D,E,F"],
        ["Gamma", "North", 150, 3, "2026-01-03", "A", "text", "G,H,I"],
        ["Alpha", "East", 90, 4, "2026-01-04", "C", "text", "J,K,L"],
    ]
    ws.Range("A2:H5").Value = rows
    ws.Range("I1:K3").Value = [["Month", "Sales", "Rate"], ["Jan", 100, 0.1], ["Feb", 120, 0.2]]
    crit = wb.Worksheets("Criteria")
    crit.Range("A1:A2").Value = [["Region"], ["North"]]
    wb.Worksheets("Aux").Range("A1:B3").Value = [["Key", "Value"], ["x", 1], ["y", 2]]
    return wb


def _close_workbook(wb) -> None:
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def _prepare_chart() -> None:
    _call("create_chart", {"range_addr": "I1:K3", "chart_type": "line", "title": "Smoke Chart", "sheet": "Data"})


def _prepare_pivot() -> None:
    _call("create_pivot_table", {
        "source_range": "A1:H5",
        "dest_sheet": "Pivot",
        "row_field": "Region",
        "value_field": "Amount",
        "source_sheet": "Data",
    })


def _resolve_output_dir(output_dir: str | Path | None = None) -> Path:
    requested = output_dir or os.environ.get(OUTPUT_DIR_ENV)
    if requested:
        out_dir = Path(requested).expanduser().resolve()
    else:
        out_dir = Path(tempfile.mkdtemp(prefix="excel_ai_tool_smoke_")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def run_tool_smoke(output_dir: str | Path | None = None) -> dict[str, Any]:
    out_dir = _resolve_output_dir(output_dir)
    original_macro_path = macro._MACROS_PATH
    macro._MACROS_PATH = out_dir / "smoke_macros.json"

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        _patch_excel(app)

        image_path = out_dir / "smoke_image.png"
        image_path.write_bytes(base64.b64decode(ONE_PIXEL_PNG))
        save_path = out_dir / "smoke_saved.xlsx"

        tool_cases: list[tuple[str, dict[str, Any], str | None]] = [
        ("read_range", {"range_addr": "A1:C3", "sheet": "Data"}, None),
        ("get_sheet_info", {}, None),
        ("get_used_range", {"sheet": "Data"}, None),
        ("get_workbook_summary", {}, None),
        ("write_range", {"range_addr": "J5", "values": [["ok"]], "sheet": "Data"}, None),
        ("save_workbook", {}, "save"),
        ("format_range", {"range_addr": "A1:H1", "sheet": "Data", "bold": True, "fill": "#D9EAF7"}, None),
        ("set_borders", {"range_addr": "A1:H5", "sheet": "Data", "style": "continuous", "sides": "all"}, None),
        ("add_conditional_format", {"range_addr": "C2:C5", "sheet": "Data", "condition_type": "greater", "value": 100, "fill_color": "#FCE4D6"}, None),
        ("merge_cells", {"range_addr": "J1:K1", "sheet": "Data"}, None),
        ("unmerge_cells", {"range_addr": "J1:K1", "sheet": "Data"}, "merge"),
        ("clear_range", {"range_addr": "G2:G3", "target": "values", "sheet": "Data"}, None),
        ("insert_row", {"index": 3, "count": 1, "sheet": "Data"}, None),
        ("delete_row", {"index": 5, "count": 1, "sheet": "Data"}, None),
        ("insert_column", {"index": 3, "count": 1, "sheet": "Data"}, None),
        ("delete_column", {"index": 8, "count": 1, "sheet": "Data"}, None),
        ("set_row_height", {"row_index": 2, "height": 24, "sheet": "Data"}, None),
        ("add_sheet", {"name": "Added"}, None),
        ("delete_sheet", {"name": "Aux"}, None),
        ("move_sheet", {"name": "Aux", "after": "Data"}, None),
        ("copy_sheet", {"name": "Aux", "new_name": "Aux Copy"}, None),
        ("set_print_titles", {"rows": "$1:$1", "sheet": "Data"}, None),
        ("add_header_footer", {"header": "Smoke", "footer": "Page &P", "sheet": "Data"}, None),
        ("protect_sheet", {"sheet": "Data", "password": "pw"}, None),
        ("unprotect_sheet", {"sheet": "Data", "password": "pw"}, "protect"),
        ("rename_sheet", {"old_name": "Aux", "new_name": "AuxRenamed"}, None),
        ("sort_range", {"range_addr": "A1:H5", "column_index": 3, "ascending": False, "has_header": True, "sheet": "Data"}, None),
        ("find_replace", {"find": "Beta", "replace": "Beta2", "sheet": "Data"}, None),
        ("trim_range", {"range_addr": "G2:G2", "sheet": "Data"}, None),
        ("filter_range", {"range_addr": "A1:H5", "column_index": 2, "criteria": "North", "sheet": "Data"}, None),
        ("copy_range", {"source_range": "A1:B2", "dest_range": "A10", "source_sheet": "Data", "dest_sheet": "Aux"}, None),
        ("add_comment", {"range_addr": "A2", "comment": "smoke", "author": "test", "sheet": "Data"}, None),
        ("set_data_validation", {"range_addr": "F2:F5", "options": "A,B,C", "sheet": "Data"}, None),
        ("delete_chart", {"chart_index": 1, "sheet": "Data"}, "chart"),
        ("move_chart", {"chart_index": 1, "left": 300, "top": 30, "width": 300, "height": 200, "sheet": "Data"}, "chart"),
        ("create_chart", {"range_addr": "I1:K3", "chart_type": "line", "title": "Smoke Chart", "sheet": "Data"}, None),
        ("create_pivot_table", {"source_range": "A1:H5", "dest_sheet": "Pivot", "row_field": "Region", "value_field": "Amount", "source_sheet": "Data"}, None),
        ("refresh_pivot_table", {"pivot_sheet": "Pivot"}, "pivot"),
        ("format_pivot_table", {"pivot_sheet": "Pivot", "style": "PivotStyleMedium9"}, "pivot"),
        ("freeze_panes", {"row": 1, "col": 0, "sheet": "Data"}, None),
        ("auto_fit", {"target": "columns", "range_addr": "A:H", "sheet": "Data"}, None),
        ("set_column_width", {"column_index": 2, "width": 18, "sheet": "Data"}, None),
        ("apply_table_style", {"range_addr": "A1:H5", "style": "TableStyleMedium2", "table_name": "SmokeTable", "sheet": "Data"}, None),
        ("format_chart", {"chart_index": 1, "title": "Updated Chart", "has_legend": True, "sheet": "Data"}, "chart"),
        ("create_combo_chart", {"range_addr": "I1:K3", "line_series_index": 2, "secondary_axis": True, "title": "Combo", "sheet": "Data"}, None),
        ("add_sparklines", {"data_range": "C2:D2", "sparkline_range": "J2", "sparkline_type": "line", "sheet": "Data"}, None),
        ("set_tab_color", {"color": "#70AD47", "sheet": "Data"}, None),
        ("page_setup", {"orientation": "landscape", "paper_size": "A4", "print_area": "A1:H5", "sheet": "Data"}, None),
        ("add_slicer", {"pivot_sheet": "Pivot", "field_name": "Region", "dest_sheet": "Pivot"}, "pivot"),
        ("summarize_range", {"range_addr": "C2:C5", "sheet": "Data"}, None),
        ("find_duplicates", {"range_addr": "A1:H5", "column_index": 1, "action": "mark", "sheet": "Data"}, None),
        ("fill_series", {"start_cell": "L2", "count": 5, "series_type": "number", "start_value": 1, "step": 1, "direction": "down", "sheet": "Data"}, None),
        ("group_rows", {"start_row": 2, "end_row": 4, "action": "group", "sheet": "Data"}, None),
        ("group_columns", {"start_col": 3, "end_col": 4, "action": "group", "sheet": "Data"}, None),
        ("transpose_range", {"source_range": "A1:B3", "dest_cell": "L1", "source_sheet": "Data", "dest_sheet": "Data"}, None),
        ("name_range", {"range_addr": "C2:C5", "name": "SmokeAmount", "sheet": "Data"}, None),
        ("add_subtotal", {"range_addr": "A1:H5", "group_by_column": 2, "value_columns": [3], "function_type": "sum", "sheet": "Data"}, None),
        ("advanced_filter", {"range_addr": "A1:H5", "criteria_range": "M1:M2", "dest_range": "M5", "unique_only": False, "sheet": "Data"}, "criteria"),
        ("split_text_to_columns", {"range_addr": "H2:H5", "delimiter": "comma", "sheet": "Data"}, None),
        ("add_image", {"image_path": str(image_path), "range_addr": "J10", "width": 20, "height": 20, "sheet": "Data"}, None),
        ("undo_last", {}, None),
        ("record_macro", {"name": "smoke_macro", "description": "smoke", "steps": [{"tool": "write_range", "args": {"range_addr": "A1", "values": [["x"]], "sheet": "Data"}}]}, None),
        ("list_macros", {}, None),
        ("run_macro", {"name": "smoke_macro"}, "macro"),
        ("delete_macro", {"name": "smoke_macro"}, "macro"),
        ("validate_formula", {"formula": "=SUM(A1:A3)"}, None),
        ("explain_formula", {"formula": "=SUM(A1:A3)"}, None),
        ("query_range", {"range_addr": "A1:H5", "condition_json": '{"filters":[{"column":"Region","op":"equals","value":"North"}]}', "sheet": "Data"}, None),
        ("list_workbooks", {}, None),
        ("switch_workbook", {"name": None}, "workbook_name"),
        ("copy_range_between_workbooks", {}, "cross_workbook"),
    ]

        results: list[dict[str, Any]] = []
        for tool_name, args, prep in tool_cases:
            wb = None
            extra_wb = None
            started = time.perf_counter()
            try:
                wb = _setup_workbook(app)
                if prep == "save":
                    wb.SaveAs(str(save_path))
                elif prep == "merge":
                    wb.Worksheets("Data").Range("J1:K1").Merge()
                elif prep == "protect":
                    wb.Worksheets("Data").Protect(Password="pw")
                elif prep == "chart":
                    _prepare_chart()
                elif prep == "pivot":
                    _prepare_pivot()
                elif prep == "criteria":
                    wb.Worksheets("Criteria").Range("A1:A2").Copy(wb.Worksheets("Data").Range("M1"))
                elif prep == "macro":
                    macro.record_macro("smoke_macro", steps=[{"tool": "write_range", "args": {"range_addr": "A1", "values": [["x"]], "sheet": "Data"}}])
                elif prep == "workbook_name":
                    args = {"name": wb.Name}
                elif prep == "cross_workbook":
                    extra_wb = app.Workbooks.Add()
                    args = {
                        "source_range": "A1:B2",
                        "dest_range": "A1",
                        "source_wb": wb.Name,
                        "dest_wb": extra_wb.Name,
                        "source_sheet": "Data",
                        "dest_sheet": extra_wb.Worksheets(1).Name,
                        "values_only": True,
                    }

                parsed = _call(tool_name, args)
                ok = not _is_error_response(parsed)
                results.append({
                    "tool": tool_name,
                    "ok": ok,
                    "dangerous": tool_name in DANGEROUS_TOOLS,
                    "duration_ms": round((time.perf_counter() - started) * 1000, 1),
                    "result": parsed,
                })
            except Exception as exc:
                results.append({
                    "tool": tool_name,
                    "ok": False,
                    "dangerous": tool_name in DANGEROUS_TOOLS,
                    "duration_ms": round((time.perf_counter() - started) * 1000, 1),
                    "exception": type(exc).__name__,
                    "message": str(exc),
                })
            finally:
                if extra_wb is not None:
                    _close_workbook(extra_wb)
                if wb is not None:
                    _close_workbook(wb)
    finally:
        macro._MACROS_PATH = original_macro_path
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    report = {
        "total": len(results),
        "passed": sum(1 for r in results if r["ok"]),
        "failed": [r for r in results if not r["ok"]],
        "results": results,
    }
    report_path = out_dir / "tool_smoke_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    report["report_path"] = str(report_path)
    return report


def main(output_dir: str | Path | None = None) -> int:
    report = run_tool_smoke(output_dir)
    print(json.dumps({
        "total": report["total"],
        "passed": report["passed"],
        "failed": len(report["failed"]),
        "report": report["report_path"],
        "failed_tools": [r["tool"] for r in report["failed"]],
    }, ensure_ascii=False, indent=2))
    return 0 if not report["failed"] else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run isolated Excel COM smoke tests for all tools.")
    parser.add_argument(
        "--output-dir",
        help=f"Directory for smoke artifacts. Defaults to ${OUTPUT_DIR_ENV} or a system temp dir.",
    )
    args = parser.parse_args()
    raise SystemExit(main(args.output_dir))
