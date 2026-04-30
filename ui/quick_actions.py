"""Quick action prompts for common Excel user workflows."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping, MutableMapping


@dataclass(frozen=True)
class QuickAction:
    key: str
    label: str
    prompt: str


QUICK_ACTIONS: tuple[QuickAction, ...] = (
    QuickAction(
        key="beautify_report",
        label="美化目前表格",
        prompt="幫我把目前工作簿中的主要資料表整理得漂亮一點，做成可以給主管看的報表格式。",
    ),
    QuickAction(
        key="summarize_data",
        label="產生資料摘要",
        prompt="幫我針對目前工作簿中的主要資料表產生資料摘要，包含資料範圍、欄位、筆數、數值欄合計與平均。",
    ),
    QuickAction(
        key="sum_by_group",
        label="查詢加總資料",
        prompt="幫我找出目前資料中各類別或地區的總金額，並指出最高的一組。",
    ),
    QuickAction(
        key="create_report_chart",
        label="建立報告圖表",
        prompt="幫我用目前資料建立一張適合報告使用的圖表，放在資料右側空白區。",
    ),
    QuickAction(
        key="undo_last",
        label="復原上一步",
        prompt="復原上一步操作。",
    ),
)


QUICK_ACTION_FORM_CHOICES: dict[str, dict[str, tuple[Any, ...]]] = {
    "beautify_report": {
        "theme": ("藍色", "綠色", "灰色"),
    },
    "summarize_data": {
        "depth": ("標準", "快速", "詳細"),
    },
    "sum_by_group": {
        "group_by": ("自動判斷", "地區", "類別", "月份", "負責人"),
        "value_col": ("自動判斷", "金額", "銷售額", "數量"),
    },
    "create_report_chart": {
        "chart_type": ("自動判斷", "長條圖", "折線圖", "圓餅圖"),
        "placement": ("資料右側空白區", "新工作表", "目前工作表下方"),
    },
}

DEFAULT_QUICK_ACTION_OPTIONS: dict[str, dict[str, Any]] = {
    "beautify_report": {
        "theme": "藍色",
        "freeze_header": True,
        "save_after": False,
    },
    "summarize_data": {
        "depth": "標準",
        "include_recommendations": True,
    },
    "sum_by_group": {
        "group_by": "自動判斷",
        "value_col": "自動判斷",
        "include_total": True,
    },
    "create_report_chart": {
        "chart_type": "自動判斷",
        "placement": "資料右側空白區",
        "include_title": True,
    },
    "undo_last": {},
}


def get_quick_action(action_key: str) -> QuickAction:
    for action in QUICK_ACTIONS:
        if action.key == action_key:
            return action
    raise KeyError(f"Unknown quick action: {action_key}")


def quick_action_options(action_key: str, options: Mapping[str, Any] | None = None) -> dict[str, Any]:
    merged = dict(DEFAULT_QUICK_ACTION_OPTIONS.get(action_key, {}))
    if options:
        merged.update(options)
    get_quick_action(action_key)
    return merged


def build_quick_action_prompt(action_key: str, options: Mapping[str, Any] | None = None) -> str:
    opts = quick_action_options(action_key, options)

    if action_key == "beautify_report":
        freeze = "請凍結表頭，方便往下瀏覽" if opts.get("freeze_header") else "不要凍結表頭"
        save = "完成後請儲存檔案" if opts.get("save_after") else "完成後先不要自動儲存，讓我確認結果"
        return (
            f"幫我把目前工作簿中的主要資料表整理成可以給主管看的報表格式。"
            f"視覺主題用{opts.get('theme', '藍色')}，{freeze}，{save}。"
            "完成後請用白話摘要列出處理範圍、做了哪些整理、以及是否需要我再確認下一步。"
        )

    if action_key == "summarize_data":
        recommendation = "並補充你建議我下一步可以檢查什麼" if opts.get("include_recommendations") else "不需要額外建議"
        return (
            f"幫我針對目前工作簿中的主要資料表產生{opts.get('depth', '標準')}資料摘要，"
            f"包含資料範圍、欄位、筆數、數值欄合計與平均，{recommendation}。"
            "完成後請用白話整理成容易給主管看的重點。"
        )

    if action_key == "sum_by_group":
        group_by = opts.get("group_by", "自動判斷")
        value_col = opts.get("value_col", "自動判斷")
        group_text = "請自動判斷最適合的分類欄位" if group_by == "自動判斷" else f"優先用「{group_by}」分組"
        value_text = "請自動判斷最適合加總的數值欄位" if value_col == "自動判斷" else f"優先加總「{value_col}」"
        total_text = "請顯示總計" if opts.get("include_total") else "可以不用顯示總計"
        return (
            f"幫我找出目前資料中各類別或地區的總金額。{group_text}，{value_text}，{total_text}，"
            "並指出最高的一組。完成後請用表格或條列摘要說明結果。"
        )

    if action_key == "create_report_chart":
        chart_type = opts.get("chart_type", "自動判斷")
        chart_text = "請自動判斷最適合的圖表類型" if chart_type == "自動判斷" else f"請建立{chart_type}"
        title_text = "請加上清楚標題" if opts.get("include_title") else "不需要額外標題"
        return (
            f"幫我用目前資料建立一張適合報告使用的圖表。{chart_text}，"
            f"放在{opts.get('placement', '資料右側空白區')}，{title_text}。"
            "完成後請用白話說明圖表位置、使用的資料與主要觀察。"
        )

    if action_key == "undo_last":
        return "請復原上一步操作。完成後請用白話說明復原了什麼、目前資料是否已回到前一個狀態。"

    get_quick_action(action_key)
    raise KeyError(f"Unknown quick action: {action_key}")


def quick_action_prompt(action_key: str) -> str:
    return build_quick_action_prompt(action_key)


def open_quick_action_form(state: MutableMapping[str, Any], action_key: str) -> None:
    get_quick_action(action_key)
    state["_quick_action_form"] = action_key


def clear_quick_action_form(state: MutableMapping[str, Any]) -> None:
    state.pop("_quick_action_form", None)


def queue_quick_action(
    state: MutableMapping[str, Any],
    action_key: str,
    options: Mapping[str, Any] | None = None,
) -> str:
    action = get_quick_action(action_key)
    prompt = build_quick_action_prompt(action_key, options)
    state["_queued_prompt"] = prompt
    state["_quick_action_last"] = {
        "key": action.key,
        "label": action.label,
        "options": quick_action_options(action_key, options),
    }
    return prompt
