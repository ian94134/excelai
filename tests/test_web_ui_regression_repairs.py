from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import agent
import excel._undo as undo_mod


class _SilentAfterToolProvider(agent.LLMProvider):
    def __init__(self):
        self.calls = 0

    def chat(self, messages, tools):
        raise NotImplementedError

    def chat_stream(self, messages, tools):
        self.calls += 1
        if self.calls == 1:
            yield (
                "tool_calls",
                agent.LLMResponse(
                    text=None,
                    tool_calls=[
                        agent.ToolCall(
                            id="call_1",
                            name="beautify_range",
                            arguments={
                                "sheet": "SalesData",
                                "range_addr": "A1:G10",
                                "theme": "blue",
                            },
                        )
                    ],
                    raw_assistant_message={
                        "role": "assistant",
                        "content": None,
                        "tool_calls": [
                            {
                                "id": "call_1",
                                "type": "function",
                                "function": {
                                    "name": "beautify_range",
                                    "arguments": "{}",
                                },
                            }
                        ],
                    },
                ),
            )
        else:
            if False:
                yield None


def test_format_range_repair_recovers_visible_style_args():
    args = {"sheet": "Report", "range_addr": "A1:C1"}

    repaired = agent._repair_format_range_args(
        args,
        "把 Report!A1:C1 設成粗體、藍底白字、置中",
    )

    assert repaired["bold"] is True
    assert repaired["fill"] == "#4472C4"
    assert repaired["color"] == "#FFFFFF"
    assert repaired["horizontal_alignment"] == "center"


def test_sort_and_filter_repair_map_header_names(monkeypatch):
    monkeypatch.setattr(
        agent.et,
        "read_range",
        lambda range_addr, sheet=None: [
            ["Date", "Region", "Product", "Quantity", "Amount", "Category", "Notes"],
            ["2026-01-01", "North", "Alpha", 12, 120, "A", "north apple"],
        ],
    )

    sort_args = agent._repair_range_column_args(
        "sort_range",
        {"sheet": "SalesData", "range_addr": "A1:G10", "column_index": 1},
        "把 SalesData!A1:G10 依 Amount 欄由大到小排序，第一列是表頭",
    )
    assert sort_args["column_index"] == 5
    assert sort_args["ascending"] is False
    assert sort_args["has_header"] is True

    filter_args = agent._repair_range_column_args(
        "filter_range",
        {"sheet": "SalesData", "range_addr": "A1:G10"},
        "對 SalesData!A1:G10 篩選出 Region 是 North 的列",
    )
    assert filter_args["column_index"] == 2
    assert filter_args["criteria"] == "North"


def test_data_validation_repair_accepts_formula1_alias():
    repaired = agent._repair_data_validation_args(
        {
            "sheet": "ValidationData",
            "range_addr": "B2:B4",
            "validation_type": "list",
            "formula1": "A,B,C",
        },
        "對 ValidationData!B2:B4 設定下拉選單，選項是 A,B,C",
    )

    assert repaired["options"] == "A,B,C"
    assert "formula1" not in repaired
    assert "validation_type" not in repaired


def test_fill_series_repair_recovers_start_value_and_count():
    repaired = agent._repair_fill_series_args(
        {"sheet": "Report", "start_cell": "K1", "direction": "down"},
        "從 Report!K1 開始往下填 1 到 5",
    )

    assert repaired["start_value"] == "1"
    assert repaired["count"] == 5


def test_page_setup_repair_recovers_print_settings():
    repaired = agent._repair_page_setup_args(
        {"sheet": "SalesData"},
        (
            "請呼叫 page_setup。sheet=SalesData，orientation=landscape，paper_size=a4，"
            "fit_to_wide=1，fit_to_tall=0，print_area=A1:O20，center_horizontally=true。"
        ),
    )

    assert repaired["orientation"] == "landscape"
    assert repaired["paper_size"] == "a4"
    assert repaired["fit_to_wide"] == 1
    assert repaired["fit_to_tall"] == 0
    assert repaired["print_area"] == "A1:O20"
    assert repaired["center_horizontally"] is True


def test_query_range_repair_builds_filter_and_amount_sum(monkeypatch):
    monkeypatch.setattr(
        agent.et,
        "read_range",
        lambda range_addr, sheet=None: [
            ["Date", "Region", "Product", "Quantity", "Amount", "Category", "Notes"],
            ["2026-01-01", "North", "Alpha", 12, 120, "A", "north apple"],
        ],
    )

    repaired = agent._repair_query_range_args(
        {
            "sheet": "SalesData",
            "range_addr": "A1:G10",
            "query": "Region=North 的所有列，並把 Amount 加總",
        },
        "請直接呼叫 query_range。",
    )

    assert repaired["filters"] == [
        {"column": "Region", "operator": "=", "value": "North"}
    ]
    assert repaired["aggregation"] == {"function": "sum", "column": "Amount"}


def test_undo_blank_single_cell_clears_range(monkeypatch):
    class FakeRange:
        def __init__(self):
            self.cleared = False
            self.Value = "UNDO_MARK"

        def ClearContents(self):
            self.cleared = True
            self.Value = None

    class FakeSheet:
        def __init__(self):
            self.range = FakeRange()

        def Range(self, addr):
            assert addr == "Z1"
            return self.range

    fake_sheet = FakeSheet()
    monkeypatch.setattr(undo_mod, "_get_excel", lambda: object())
    monkeypatch.setattr(undo_mod, "_get_sheet", lambda excel, sheet=None: fake_sheet)

    entry = SimpleNamespace(
        tool_name="write_range",
        arguments={"sheet": "Target", "range_addr": "Z1"},
        values_before=[],
        formats_before=None,
        widths_before=None,
        heights_before=None,
    )

    result = undo_mod._undo_last_body(entry)

    assert result["status"] == "ok"
    assert result["undone"] == "write_range"
    assert fake_sheet.range.cleared is True
    assert fake_sheet.range.Value is None


def test_run_turn_falls_back_to_plain_completion_after_successful_tool(monkeypatch):
    messages = [{"role": "user", "content": "幫我把 SalesData 這張表整理漂亮一點"}]
    provider = _SilentAfterToolProvider()

    monkeypatch.setattr(agent.backup, "get_session_stack", lambda: None)
    monkeypatch.setattr(
        agent,
        "execute",
        lambda name, args: (
            '{"status":"ok","tool":"beautify_range","sheet":"SalesData",'
            '"range":"$A$1:$G$10","theme":"blue",'
            '"applied":["header","banded_rows","number_format:E","filter","auto_fit_columns"]}'
        ),
    )

    events = []
    for kind, data in agent.run_turn(
        get_messages=lambda: list(messages),
        tools=[],
        provider=provider,
        dangerous_tools=set(),
        max_iterations=3,
    ):
        events.append((kind, data))
        if kind == agent.EVT_ASST_MSG:
            messages.append(data)
        elif kind == agent.EVT_TOOL_DONE:
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": data.tc.id,
                    "name": data.tc.name,
                    "content": data.result_json,
                }
            )

    assert events[-1][0] == agent.EVT_DONE
    assert "已完成表格美化" in events[-1][1]
    assert "SalesData" in events[-1][1]
