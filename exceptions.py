"""
Excel AI 助手的 Custom Exception 家族。

設計原則：
- 所有 Excel / LLM / 工具執行錯誤都繼承自 ExcelAIError，方便 executor.execute() 統一捕捉。
- 每個子類別對應一個具體失敗情境，LLM 收到 JSON 結果內 `error_type` 欄位後可依類型採取對應補救動作
  （例如 SheetNotFoundError 就去呼叫 get_sheet_info 重新確認工作表清單）。
- 訊息以使用者語言（繁體中文）呈現；`error_type` 則保留英文類別名供程式與 LLM 判斷。

V4 Phase 1 引入。後續 Phase 2 的 Undo、Phase 3 的 telemetry 會依 error_type 分類統計。
相關文件：PLAN_v4.md §C-3、TROUBLESHOOTING.md、CONVENTIONS.md
"""


class ExcelAIError(Exception):
    """所有 Excel AI 錯誤的根類別。executor 統一捕捉這個類型以附加 error_type。"""


# ── Excel / COM 相關 ──────────────────────────────────────────────────────────

class ExcelNotFoundError(ExcelAIError):
    """Excel 應用程式未啟動或 COM 無法連線。對應歷史情境：B7 子線程 COM 未初始化。"""


class NoActiveWorkbookError(ExcelAIError):
    """Excel 已啟動但沒有開啟的活頁簿。對應歷史情境：B5 ActiveWorkbook None。"""


class SheetNotFoundError(ExcelAIError):
    """指定的工作表名稱不存在於當前活頁簿。"""


class InvalidRangeError(ExcelAIError):
    """A1 記號錯誤或範圍格式不合法（例如 'X999999999'、含中文字符的範圍字串）。"""


# ── 工具參數相關 ──────────────────────────────────────────────────────────────

class InvalidToolArgumentsError(ExcelAIError):
    """
    工具參數缺少必要語意，導致操作無實際效果。
    例如 format_range 只提供 range_addr、但沒有任何格式屬性（bold/fill/...）。
    """


# ── 工具執行相關 ──────────────────────────────────────────────────────────────

class ToolTimeoutError(ExcelAIError):
    """單一工具執行超過 TOOL_TIMEOUT_SEC（預設 30s）未完成。Phase 2 才會啟用 timeout 包裝。"""


# ── LLM Provider 相關 ────────────────────────────────────────────────────────

class ProviderConnectionError(ExcelAIError):
    """LLM Provider 連線失敗（例如 Qwen 伺服器 down、網路斷線）。Phase 2 會搭配 tenacity 重試。"""
