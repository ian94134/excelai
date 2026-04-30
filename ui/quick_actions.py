"""Quick action prompts for common Excel user workflows."""
from __future__ import annotations

from dataclasses import dataclass
from typing import MutableMapping


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


def quick_action_prompt(action_key: str) -> str:
    for action in QUICK_ACTIONS:
        if action.key == action_key:
            return action.prompt
    raise KeyError(f"Unknown quick action: {action_key}")


def queue_quick_action(state: MutableMapping[str, str], action_key: str) -> str:
    prompt = quick_action_prompt(action_key)
    state["_queued_prompt"] = prompt
    return prompt
