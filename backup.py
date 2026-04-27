"""
操作備份系統（v4.6.0 Phase 2/3 + 持久化）。

設計依據：PLAN_v4.md §A-1 / §A-2。

## 本檔職責

- BackupEntry dataclass — 單次操作的備份記錄（含 values_before / formats_before）
- BACKUP_NEEDED 對照表 — 標示哪些工具需要備份、哪些屬於唯讀可略過
- capture_before() — 記錄 tool_name / arguments / timestamp；回傳 BackupEntry 骨架
  （values_before / formats_before 由 executor.py 在執行前填入）
- BackupStack — 20 步上限 + push / pop / peek / clear / snapshot 完整功能
- get_session_stack() — Streamlit session_state 的 BackupStack 存取點（供 executor / undo_last 共用）
- restore() — 明確拋出 NotImplementedError；實際還原邏輯在 excel_tools.undo_last()
- save_current_stack() / _load_stack() — v4.6.0 新增：跨 rerun 持久化至 ~/.excel-ai/backup_stack.json

## Undo 策略（實作於 excel_tools.undo_last）

- Category A：僅用 arguments 推算反向操作（insert_row / insert_column / add/rename/merge/unmerge sheet）
- Category B：values_before 回寫（write_range / clear_range / trim_range）
- Category B-ext：formats_before 回寫（format_range / set_borders）；executor 執行前呼叫 capture_formats_before()
- Category C：嘗試 Excel 原生 Undo（CommandBars.ExecuteMso("Undo")），失敗才回傳 cannot_undo

## 設計原則

- 本檔不依賴 Streamlit，可純 Python 測試
- capture_before 只記 metadata；values_before / formats_before 由 executor.py 在 execute() 內填入
- BackupStack 達上限後丟棄最舊項目，符合「只能還原最近 20 步」語意
- 持久化寫入使用 atomic write（tmp → replace），不會產生半寫狀態
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from tools.registry import get_registered_backup_needed


# ── 常數 ──────────────────────────────────────────────────────────────────────

BACKUP_STACK_MAX = 20
BACKUP_CELL_LIMIT = 10000  # Phase 2 會用到：超過此 cell 數的範圍改為地址-only 備份

# 持久化路徑與版本（版本不符時丟棄舊檔，安全地重建）
_PERSIST_PATH    = Path.home() / ".excel-ai" / "backup_stack.json"
_PERSIST_VERSION = 1


# 哪些工具的操作需要備份（Phase 2 的 restore 會針對這些做反向操作）。
# 唯讀與純檢視類工具不備份（沒有狀態變更可還原）。
BACKUP_NEEDED: dict[str, bool] = {
    # ── V1 工具 ────────────────────────────────────────────────────────────────
    "read_range":             False,  # 唯讀
    "write_range":            True,
    "get_sheet_info":         False,  # 唯讀
    "get_used_range":         False,  # 唯讀
    "format_range":           True,
    "insert_row":             True,
    "delete_row":             True,
    "insert_column":          True,
    "delete_column":          True,
    "add_sheet":              True,
    "rename_sheet":           True,
    "delete_sheet":           True,
    "move_sheet":             True,
    "copy_sheet":             True,
    "set_print_titles":       True,
    "add_header_footer":      True,
    "protect_sheet":          True,
    "unprotect_sheet":        True,
    "sort_range":             True,
    "find_replace":           True,
    "save_workbook":          False,  # 不需復原（使用者只會想「重存一次」）
    # ── V2 工具 ────────────────────────────────────────────────────────────────
    "delete_chart":           True,
    "move_chart":             True,
    "create_chart":           True,
    "create_pivot_table":     True,
    "refresh_pivot_table":    False,  # 唯讀操作（只重新計算）
    "format_pivot_table":     True,
    "freeze_panes":           True,
    "auto_fit":               True,
    "set_column_width":       True,
    # ── V3 工具 ────────────────────────────────────────────────────────────────
    "filter_range":           True,
    "merge_cells":            True,
    "unmerge_cells":          True,
    "set_borders":            True,
    "clear_range":            True,
    "set_row_height":         True,
    "copy_range":             True,
    "add_conditional_format": True,
    "add_comment":            True,
    "set_data_validation":    True,
    # ── V4 新增工具 ────────────────────────────────────────────────────────────
    "trim_range":             True,
    # ── V4 美化工具群 ──────────────────────────────────────────────────────────
    "apply_table_style":      True,
    "format_chart":           True,
    "create_combo_chart":     True,
    "add_sparklines":         True,
    "set_tab_color":          True,
    "page_setup":             True,
    "add_slicer":             True,
    "add_image":              True,
    # ── V4 分析工具群 ──────────────────────────────────────────────────────────
    "summarize_range":        False,  # 唯讀計算，不修改工作表
    "get_workbook_summary":    False,  # 唯讀
    "find_duplicates":        True,   # mark/delete 會修改工作表
    "fill_series":            True,
    "group_rows":             True,
    "group_columns":          True,
    "transpose_range":        True,
    "name_range":             True,
    "add_subtotal":           True,
    "advanced_filter":        True,
    "split_text_to_columns":  True,
    # ── Phase 2 Undo ───────────────────────────────────────────────────────────
    "undo_last":              False,  # Meta 操作，不備份自身（避免 undo undo 的混亂）
    # ── V4.7.0 A：巨集工具 ────────────────────────────────────────────────────
    "record_macro":           False,  # 寫 JSON 至磁碟，不修改工作表
    "list_macros":            False,  # 唯讀
    "run_macro":              True,   # 執行一系列工具，可能修改工作表
    "delete_macro":           False,  # 只刪磁碟上的 JSON
    # ── V4.7.0 B：公式智慧輔助 ────────────────────────────────────────────────
    "validate_formula":       False,  # 純分析，不修改工作表
    "explain_formula":        False,  # 純分析，不修改工作表
    # ── V4.7.0 C：自然語言查詢 ────────────────────────────────────────────────
    "query_range":            False,  # 唯讀查詢
    # ── V4.7.0 D：多工作簿協作 ────────────────────────────────────────────────
    "list_workbooks":                False,  # 唯讀
    "switch_workbook":               False,  # 切換視窗，不改資料
    "copy_range_between_workbooks":  True,   # 寫入目標工作簿
}
# Merge registry-registered tools (decorated with @register_tool).
# Registry flag takes precedence for new tools.
BACKUP_NEEDED.update(get_registered_backup_needed())


# ── BackupEntry dataclass ────────────────────────────────────────────────────

@dataclass
class BackupEntry:
    """單次可還原操作的備份記錄。

    Phase 1 僅填 tool_name / arguments / timestamp；Phase 2 擴充時會填後面的
    values_before / formats_before / sheet_structure_before。
    """
    tool_name: str
    arguments: dict
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))

    # Phase 2 擴充：
    affected_range: str | None = None           # 例如 "Sheet1!A1:D10"
    values_before: list[list[Any]] | None = None
    formats_before: dict | None = None          # {bold, italic, color, fill, ...}
    sheet_structure_before: dict | None = None  # 僅 add/rename/delete sheet 用
    # Phase 4 (TD-03)：欄寬 / 列高 備份
    widths_before:  dict | None = None   # {col_index: width_pts}
    heights_before: dict | None = None   # {row_index: height_pts}

    def describe(self) -> str:
        """供 UI 顯示用的人類可讀描述。Phase 2 的側邊欄按鈕會用到。"""
        return f"{self.tool_name}（{self.timestamp:%H:%M:%S}）"


class BackupStack:
    """
    LIFO 操作堆疊，用於 undo_last。
    預設上限 20 步（FIFO 溢出：超過上限時丟棄最舊項目）。
    """

    def __init__(self, max_size: int = 20) -> None:
        self._entries: list[BackupEntry] = []
        self.max_size = max_size

    def push(self, entry: BackupEntry) -> None:
        """推入新 entry；超過 max_size 時丟棄最舊（index 0）。"""
        if len(self._entries) >= self.max_size:
            self._entries.pop(0)
        self._entries.append(entry)

    def pop(self) -> BackupEntry | None:
        """取出並移除最新 entry；空時回傳 None。"""
        return self._entries.pop() if self._entries else None

    def peek(self) -> BackupEntry | None:
        """查看最新 entry 但不移除；空時回傳 None。"""
        return self._entries[-1] if self._entries else None

    def clear(self) -> None:
        self._entries.clear()

    def snapshot(self) -> list[BackupEntry]:
        """回傳所有 entry 的淺複製清單（舊→新順序）。"""
        return list(self._entries)

    def __len__(self) -> int:
        return len(self._entries)

    def __iter__(self):
        return iter(self._entries)


# ── 核心函數 ──────────────────────────────────────────────────────────────────

def get_session_stack() -> "BackupStack | None":
    """
    取得目前 Streamlit session 的 BackupStack（executor 與 excel_tools 共用）。

    v4.6.0 變更：首次建立時自動從 ~/.excel-ai/backup_stack.json 載入持久化資料，
    讓 undo 歷史在頁面重整後依然存在。

    非 Streamlit 環境（如 pytest）回傳 None，讓呼叫端跳過備份邏輯。
    """
    try:
        import streamlit as st
    except ImportError:
        return None
    try:
        session_state = st.session_state
    except Exception:
        return None
    if "_backup_stack" not in session_state:
        loaded = _load_stack()
        session_state["_backup_stack"] = loaded if loaded is not None else BackupStack()
    return session_state["_backup_stack"]


def capture_before(tool_name: str, arguments: dict) -> BackupEntry | None:
    """
    在工具執行前呼叫。回傳應該推入 BackupStack 的 entry，或 None 表示無需備份。

    Phase 1：回傳的 entry 僅含 tool_name / arguments / timestamp。
    Phase 2 會依 BACKUP_STRATEGY_MAP 分派，針對不同 tool 類型填入 values_before
    / formats_before / sheet_structure_before。
    """
    if not BACKUP_NEEDED.get(tool_name, False):
        return None
    return BackupEntry(tool_name=tool_name, arguments=dict(arguments))


def restore(entry: BackupEntry) -> dict:
    """
    Reverse the operation recorded in entry, restoring Excel to its prior state.

    Dispatch strategy (mirrors excel_tools.undo_last categories):
      - values_before present  → write_range to restore cell values
      - formats_before present → format_range / set_borders to restore formats
      - sheet_structure_before → rename / delete the sheet as needed
      - fallback               → delegate to excel_tools.undo_last() which
                                 handles per-tool Category C (native Undo) and
                                 Category D (cannot-undo explanations)

    This is the authoritative restore implementation.  execute_batch and any
    future undo paths should call this function rather than calling
    excel_tools.undo_last() directly.
    """
    import excel_tools as et   # late import to avoid circular dependency

    name = entry.tool_name
    args = entry.arguments or {}
    sheet = args.get("sheet")

    # ── Category B-values: restore cell data ──────────────────────────────
    if entry.values_before is not None:
        rng_addr = args.get("range_addr", "")
        if rng_addr and entry.values_before:
            try:
                et.write_range(rng_addr, entry.values_before, sheet)
                return {"status": "ok", "undone": name, "method": "values_restore"}
            except Exception as e:
                return {"status": "error", "undone": name, "error": str(e)}

    # ── Category B-formats: restore formatting ────────────────────────────
    if entry.formats_before is not None:
        rng_addr = args.get("range_addr", "")
        if rng_addr:
            try:
                et._restore_formats(entry.formats_before, sheet)
                return {"status": "ok", "undone": name, "method": "formats_restore"}
            except Exception as e:
                return {"status": "error", "undone": name, "error": str(e)}

    # ── Category A/C/D: delegate to per-tool logic in excel_tools ─────────
    # excel_tools.undo_last() pops from the stack internally, but here we
    # already have the entry, so we call the inner dispatcher directly.
    try:
        return et._undo_dispatch(entry)
    except AttributeError:
        # Fallback for versions that do not expose _undo_dispatch yet
        return et.undo_last()


# ── 持久化（v4.6.0）──────────────────────────────────────────────────────────

def _entry_to_dict(entry: BackupEntry) -> dict:
    """BackupEntry → JSON 相容的 dict。"""
    return {
        "tool_name":              entry.tool_name,
        "arguments":              entry.arguments,
        "timestamp":              entry.timestamp.isoformat(),
        "affected_range":         entry.affected_range,
        "values_before":          entry.values_before,
        "formats_before":         entry.formats_before,
        "sheet_structure_before": entry.sheet_structure_before,
        "widths_before":          entry.widths_before,
        "heights_before":         entry.heights_before,
    }


def _dict_to_entry(d: dict) -> BackupEntry:
    """JSON dict → BackupEntry。型態轉換失敗時回傳帶預設值的 entry。"""
    ts_str = d.get("timestamp")
    try:
        ts = datetime.fromisoformat(ts_str) if ts_str else datetime.now(timezone.utc)
    except (ValueError, TypeError):
        ts = datetime.now(timezone.utc)
    return BackupEntry(
        tool_name=d.get("tool_name", "unknown"),
        arguments=d.get("arguments") or {},
        timestamp=ts,
        affected_range=d.get("affected_range"),
        values_before=d.get("values_before"),
        formats_before=d.get("formats_before"),
        sheet_structure_before=d.get("sheet_structure_before"),
        widths_before=d.get("widths_before"),
        heights_before=d.get("heights_before"),
    )


def save_current_stack() -> None:
    """
    將目前 session 的 BackupStack 持久化到 ~/.excel-ai/backup_stack.json。

    由 executor.py 在 push / pop 後呼叫，確保 undo 歷史在 Streamlit 頁面重整後
    依然存在。寫入採 atomic write（.tmp 再 replace），不產生半寫狀態。

    任何例外均靜默略過——持久化失敗絕對不能阻斷工具執行主流程。
    """
    stack = get_session_stack()
    if stack is None:
        return
    try:
        _PERSIST_PATH.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": _PERSIST_VERSION,
            "entries": [_entry_to_dict(e) for e in stack.snapshot()],
        }
        tmp = _PERSIST_PATH.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, default=str)
        tmp.replace(_PERSIST_PATH)
    except Exception:
        pass


def _load_stack() -> "BackupStack | None":
    """
    從 ~/.excel-ai/backup_stack.json 載入 BackupStack。

    版本不符、檔案不存在、JSON 損毀均回傳 None（由呼叫端建立空 stack）。
    單筆 entry 反序列化失敗時跳過該筆，不影響其他 entry。
    """
    if not _PERSIST_PATH.exists():
        return None
    try:
        with open(_PERSIST_PATH, "r", encoding="utf-8") as f:
            payload = json.load(f)
        if payload.get("version") != _PERSIST_VERSION:
            return None
        entries_raw = payload.get("entries", [])
        stack = BackupStack()
        for d in entries_raw:
            try:
                stack.push(_dict_to_entry(d))
            except Exception:
                pass  # 跳過損毀的單筆 entry
        return stack if len(stack) > 0 else None
    except Exception:
        return None
