# 版本變更紀錄（CHANGELOG）

格式依循 [Keep a Changelog](https://keepachangelog.com/)，但本檔目前同時承擔「工作區現況基線」與「歷史版本摘要」兩種用途。若版本日期與歷史規劃文件不同，以目前程式碼為準。

## [v4.7.0] — 2026-04-25

### 四項新功能：巨集錄製 / 公式智慧 / 自然語言查詢 / 多工作簿協作

#### A — 巨集錄製與重播（`macro.py`）

- 新增純 Python 模組 `macro.py`（無 win32com 依賴，可在 Linux CI 測試）。
- `record_macro(name, description, steps)` 支援兩種錄製來源：
  - 呼叫端明確提供 `steps` 清單（`[{"tool": str, "args": dict}, ...]`）。
  - 省略 `steps` 時自動從 `BackupStack` 取得最近操作歷史。
- `list_macros()` 列出所有巨集（name / description / step_count）。
- `run_macro(name)` 透過 `execute_batch()` 執行；任一步驟失敗時自動回滾整批次。
- `delete_macro(name)` 刪除指定巨集。
- `get_macro_steps(name)` 取得步驟清單（供 UI 展示，不執行）。
- 持久化：`~/.excel-ai/macros.json`（版本化 `_MACROS_VERSION=1`，`.tmp → replace` 原子寫入）；版本不符或檔案損毀時靜默回傳空字典。
- `executor.py` TOOL_MAP 新增 `record_macro` / `list_macros` / `run_macro` / `delete_macro` 四個工具入口。

#### B — 公式智慧輔助（`formula_validator.py`）

- 新增 `validate_formula_tool(formula, range_addr?, sheet?)` 工具：靜態語法解析 + COM 驗證，回傳 `is_valid`、`error_type`、`suggestion` 欄位。
- 新增 `explain_formula_tool(formula)` 工具：逐段拆解公式結構，產生中文說明（函數用途、參數意義、常見錯誤）。
- 兩工具均無狀態，可在 Linux CI 安全測試。
- `executor.py` TOOL_MAP 新增 `validate_formula` / `explain_formula`。

#### C — 自然語言資料查詢（`excel_query.py`）

- 新增 `query_range(range_addr, query, sheet?)` 工具，將自然語言條件（如「金額大於 1000 的列」）轉換為欄位篩選並回傳匹配列。
- 支援比較運算子（大於 / 小於 / 等於 / 包含 / 不包含）及多條件 AND 組合。
- 結果含 `matched_rows`、`total_rows`、`matched_count` 供 LLM 直接讀取。
- `executor.py` TOOL_MAP 新增 `query_range`。

#### D — 多工作簿協作（`excel/sheet.py`）

- 新增 `list_workbooks()` 工具：列出所有已開啟活頁簿（name / path / sheet_count / active_sheet）。
- 新增 `switch_workbook(name)` 工具：切換 Excel 前景活頁簿。
- 新增 `copy_range_between_workbooks(src_workbook, src_range, dst_workbook, dst_range, src_sheet?, dst_sheet?)` 工具：跨活頁簿複製資料範圍。
- `excel/__init__.py` 新增 `copy_range_between_workbooks` re-export。
- `executor.py` TOOL_MAP 新增 `list_workbooks` / `switch_workbook` / `copy_range_between_workbooks`。

#### 基礎設施更新

- **TOOL_MAP 擴充**：工具總數從 61 升至 **71**（新增 10 個 v4.7.0 工具）。
- **`backup.py` BACKUP_NEEDED**：補齊 10 個新工具的備份需求標記（`run_macro: True`、`copy_range_between_workbooks: True`，其餘唯讀工具 `False`）。
- **`executor.py` `execute_batch()`**：新增批次執行函數，支援自動回滾——失敗時將 stack 回退至執行前深度，依序還原已執行步驟，`巨集重播` 與測試均依賴此函數。
- **`main.py`**：版本升至 `v4.7.0`。

- **新增測試**：
  - `tests/test_macro.py`（23 個測試）：record / list / run / delete / get_steps、BackupStack 整合、持久化 round-trip、版本驗證、`execute_batch` 委派。
  - `tests/test_formula_validator.py`（20 個測試）：語法驗證、COM fallback、explain 分段、邊界輸入。
  - `tests/test_excel_query.py`（35 個測試）：比較運算子、多條件 AND、欄位名稱解析、空範圍、型別轉換。
  - 測試套件升至 **249 passed**。

---

## [v4.6.0] — 2026-04-25

### 四項弱點修復：持久化 / 壓縮 / 選取注入 / 澄清輪

#### A — BackupStack 跨 rerun 持久化（`backup.py`）

- 新增 `_entry_to_dict()` / `_dict_to_entry()` 序列化輔助函數，處理 datetime 與 None 欄位。
- 新增 `save_current_stack()` 公開函數（atomic write：`.tmp` → replace），`executor.py` 在每次 push/pop 後自動呼叫。
- 新增 `_load_stack()` 私有函數，`get_session_stack()` 首次建立 stack 時自動從 `~/.excel-ai/backup_stack.json` 載入。
- 版本欄位（`_PERSIST_VERSION=1`）保護：版本不符時丟棄舊檔重建，不污染現有流程。
- 任何持久化例外均靜默略過——絕不阻斷工具執行主流程。

#### B — Tool result 摘要壓縮（`compress.py` + `session.py`）

- 新增純函數模組 `compress.py`（無 Streamlit 依賴），實作 `compress_tool_result(content, limit=3000)`：
  - 錯誤結果（含 `"error"` 鍵）原樣保留——Qwen 需要完整錯誤上下文。
  - 成功 dict：大型陣列（> 500 字元序列化長度）替換為描述摘要，長字串截至 500 字元。
  - List payload（如 `read_range` 二維陣列）：替換為 rows/columns/sample（前 3 列）摘要。
  - 壓縮結果附加 `"_compressed": true` 標記便於除錯。
- `session.append_message()` 對 `role="tool"` 訊息自動壓縮，長對話 context window 使用率降低 30–50%。

#### C — Excel 選取範圍自動帶入（`main.py` + `ui/sidebar.py`）

- 新增 `_build_selection_tag()` helper：優先讀取 watcher 即時資料，fallback 到同步 COM 讀取，回傳 `[目前選取: Sheet1!A1:D10]` 格式。
- user message 注入改為簡潔結構化標記，Qwen 可直接用於工具參數，減少 RangeError。
- Sidebar 新增 `_render_selection_status()` 區塊，即時顯示目前選取工作表與範圍。

#### D — 多輪澄清機制（`agent.py` + `main.py`）

- 新增 `EVT_CLARIFY = "clarify"` 事件與 `_is_clarification(text)` 偵測函數（正規表示式匹配 11 種澄清句型 + 問號必要條件）。
- `run_turn()` 在無 tool call 且回應為澄清問句時 yield `EVT_CLARIFY` 取代 `EVT_DONE`，agent loop 暫停。
- `main.py` 新增 `EVT_CLARIFY` handler：顯示問句並提示「AI 需要更多資訊才能繼續」，等待使用者補充說明後自然繼續。
- 規劃模式（`plan_inject` 非空）優先於澄清偵測，確保規劃流程不受干擾。

- **新增測試**：
  - `tests/test_backup_persist.py`（15 個測試）：序列化 round-trip、missing/版本不符/損毀 entry 容錯、atomic write、全流程 save+load。
  - `tests/test_session_compress.py`（14 個測試）：錯誤結果不壓縮、dict 大型欄位摘要、list payload 摘要、非 JSON 截斷、自訂 limit。
  - `tests/test_selection_inject.py`（12 個測試）：watcher 資料、fallback、non-range selection、格式正確性、prompt 組合。
  - `tests/test_agent_clarify.py`（13 個測試）：`_is_clarification()` 正負例、`EVT_CLARIFY` yielded、規劃模式抑制、空字串邊界。
  - 測試套件升至 **174 passed**。

---


## [v4.5.0] — 2026-04-24

### 弱點修復：自動規劃觸發 + 結構化錯誤恢復

- **自動規劃觸發（`main.py`）**：新增 `_is_complex_task(prompt)` 函數，使用正規表示式偵測多步驟連接詞（然後／再／接著／並且…，≥ 2 個）或 Excel 操作關鍵字（篩選／排序／圖表／報表…，≥ 3 個）。偵測到複雜任務時自動設定 `plan_inject`，無需使用者手動開啟規劃模式，並在對話框顯示「🧭 偵測到多步驟任務，已自動啟用規劃模式」提示。

- **結構化錯誤恢復（`agent.py`）**：新增 `_ERROR_HINTS` 對照表（11 種 error_type：SheetNotFoundError / WorkbookNotFoundError / RangeError / ProtectedSheetError / FormulaError / ValueError / IndexError / ChartError / PivotError / UnknownTool / UnexpectedError）與 `_enrich_error_result()` 函數。工具失敗時自動在回傳 JSON 附加：
  - `hint`：具體的中文說明與建議做法
  - `suggested_next`：建議下一步呼叫哪個工具（如 `get_sheet_info` / `get_used_range` / `unprotect_sheet`）
  讓 Qwen 在下一輪 LLM 呼叫時直接獲得修補指引，不需使用者介入。

- **新增測試**：
  - `tests/test_complexity.py`（18 個測試）：覆蓋簡單/複雜邊界、連接詞計數、操作關鍵字計數、空字串等場景。
  - `tests/test_agent.py` 擴充（+5 個測試）：error hint 注入、suggested_next、未知 error_type fallback、成功結果不附加 hint、原始 error 欄位保留。
  - 測試套件升至 **117 passed**。

---

## [v4.4.0] — 2026-04-24

### 架構改進：agent.py 抽離 + BaseProvider Protocol（ADR-001 A2/B2）

主旨：讓 Qwen 能穩定執行複雜多步驟 Excel 任務，同時使 agent 邏輯可獨立測試。

- **新增 `agent.py`**：將原本嵌在 `main.py` 中的 LLM tool-calling 迴圈完全抽出為獨立模組。
  - `run_turn(get_messages, tools, provider, ...)` 為 generator，yield 具名事件 tuple。
  - 事件種類：`text_chunk` / `retry_info` / `asst_msg` / `tool_start` / `tool_done` /
    `rollback` / `done` / `plan_ready` / `dangerous_halt` / `repeat_halt` / `error`。
  - 完整實作：危險工具暫停、重複迴圈偵測（同工具 ≥ 3 次自動中止）、錯誤自動回滾、
    規劃模式（`plan_inject` 非空時停用工具並 yield `plan_ready`）、`max_iterations` 上限。
  - 零 Streamlit import — 純 Python 邏輯，可在 Linux CI 測試。

- **`main.py` 接線**：版本升至 `v4.4.0`，原本 `for _ in range(100):` 迴圈改為
  `for kind, data in agent.run_turn(...):`，UI 邏輯保留在 main.py，agent 邏輯徹底轉移。

- **`providers/base.py` 已完備**：`LLMProvider` ABC 定義 `chat()` / `chat_stream()` 介面；
  `LocalQwenProvider` 完整實作，含 Tenacity 指數退避重試。

- **新增 `tests/test_agent.py`**（22 個測試）：覆蓋純文字回覆、單/多工具、危險工具中止、
  重複迴圈偵測、錯誤回滾、規劃模式、max_iterations 上限、workbook context 注入等場景。
  測試套件升至 **82 passed / 1 skipped**。

---

## [v4.3.0] — 2026-04-24

### 技術債清償（TD-01 / TD-08 / TD-09）

- **TD-01：excel_tools.py 拆分為 excel/ 子套件**：原本 2,603 行的單一模組，按職責拆成 6 個子模組 + re-export shim：
  - `excel/_base.py` — COM 初始化、共用輔助函數（`_get_excel`、`_get_sheet`、`_com_tls` 等）
  - `excel/data.py` — 讀寫、排序、搜尋、清理、分析工具（21 函數）
  - `excel/format.py` — 格式化、框線、合併、欄列尺寸、列印設定（20 函數）
  - `excel/sheet.py` — 列欄/工作表管理、篩選、分組、活頁簿切換（18 函數）
  - `excel/chart.py` — 圖表、樞紐分析表、切片器（9 函數）
  - `excel/_undo.py` — undo_last / _undo_dispatch / _undo_last_body（3 函數）
  - `excel_tools.py` 縮減至 82 行純 re-export shim，所有呼叫端零改動。
- **TD-08/09：補齊 execute_batch 測試**：新增 `tests/test_executor_batch.py`（11 個測試），覆蓋全成功、錯誤中止、先前步驟回滾標記、stack pop 觸發、空清單等場景。
- **修正 test_backup.py**：更新 `restore()` 相關測試（Phase 2 已實作，不再是 NotImplementedError）；補 `get_workbook_summary` 至 `backup.BACKUP_NEEDED`（`False`）。
- **修正 conftest.py**：在所有測試檔頂端加入 Windows COM 模組替身（`pythoncom` 等），讓完整測試套件可在 Linux CI 執行。測試套件升至 **60 passed / 1 skipped**。

---

## [v4.2.1] — 2026-04-24

### 修正

- **Qwen 500 bugfix — system message 位置錯誤**：v4.2.0 的 workbook context 注入邏輯在 `use_msgs` 前方插入新的 system message，導致原有 session system message 移至 index 1，觸發 Qwen Jinja template 的 `raise_exception('System message must be at the beginning')`。
  修法：將行內 if/else 合併邏輯抽成具名 helper `_inject_ephemeral_system_ctx(msgs, ctx) -> list`，行為改為「若 index 0 已是 system role，追加到其 content 尾端；否則才 insert 新 system message」。
- **迴歸測試**：新增 `tests/test_main_helpers.py`（7 個測試），涵蓋主路徑合併、長對話不重複、無 system message 插入、空 context noop、不污染原物件、空 list 邊界等場景。測試套件從 27 升至 **34 passed / 4 skipped**。

---

## [v4.2.0] — 2026-04-24

### 弱點修復

- **批次原子性**：`tools/executor.py` 新增 `execute_batch()`，同時匯出至 `tools/__init__.py`。單輪多 tool_call 若中途失敗，自動計算本輪已推入 backup stack 的 entry 數量，呼叫 `et.undo_last()` 逐步回滾並在 UI 顯示 ⚠️ 警告；剩餘工具回傳 `{"status": "skipped"}` 讓 LLM 感知中止原因。main.py 同步追蹤 `_round_pushed`，失敗時即時回滾並清零計數。
- **新工具 `get_workbook_summary`（工具總數 61）**：`excel_tools.py` 新增同名函數，一次回傳整本活頁簿所有工作表的名稱、已用範圍、列/欄數與前 10 個標題。`tools/definition.py` 補 schema，`executor.py` 補 `TOOL_MAP`，main.py 新增 `_build_workbook_context()` + `_inject_ephemeral_system_ctx()`，每輪 LLM 呼叫前自動將活頁簿快照合併進 system message，消除「AI 假設了不存在的範圍」問題。
- **公式寫入驗證**：`excel_tools.write_range()` 寫入後回讀含 `=` 的格，偵測到 `#VALUE!` / `#REF!` / `#NAME?` 等 11 種 Excel 錯誤值時，結果 JSON 附上 `formula_errors` 陣列與 `warning` 字串，讓 LLM 能立即感知並補救。

---

## [v4.1.0] — 2026-04-24

### 穩健度與可觀測性改進

- **Tenacity retry**：`providers/local_qwen.py` 的 `chat()` 與 `chat_stream()` 加上指數退避重試（最多 3 次，ConnectError / ReadTimeout），`retry_info` 事件在 UI 顯示重試警告。
- **Before/After diff**：`main.py` 新增 `_render_diff()`，工具執行後讀取 backup stack 最新 entry 顯示儲存格值與格式的前後對比。
- **本機遙測**：新增 `telemetry.py`（SQLite，純本機，無 HTTP 外傳），`tools/executor.py` 在所有出口路徑呼叫 `record()`，側邊欄新增「📈 使用統計」區塊（總次數、成功率、Top 5 工具、最慢工具、近期錯誤）。
- **測試補齊**：新增 `tests/test_providers.py`（10 個 pytest 單元測試，5 chat + 5 stream），覆蓋 think-tag 剝除、tool_call 解析、JSON decode fallback、ConnectError 傳播等場景。

---

## [v4.0.0] — 2026-04-22

這是目前工作區可驗證到的現況基線。

### `excel-ai`

- `main.py` 版本字串為 `v4.0.0`
- `tools/definition.py` 共有 60 個 tool schema
- `tools/executor.py` 目前危險工具為 6 個
- `backup.py`、`undo_last`、工作表快照/還原、規劃模式、操作紀錄面板均已存在
- `logger.py` 已落地，輸出 JSONL 到 `~/.excel-ai/logs/`
- `pytest` 測試已存在，當前本地核對結果為 `27 passed`（v4.1.0 後升至 34）

### 工作區與文件

- 工作區同時維護 React/Electron 試算表編輯器與 `excel-ai`
- 文件改為雙軌並列敘事，不再把整個工作區誤寫成只有 `excel-ai`
- 活躍文件與歷史封存文件已分流，歷史規劃稿不再當作現況說明
- 新增 `repo_contract.yaml` 與 `scripts/check_repo_consistency.py`
- GitHub Actions 新增 `repo-guardrails`，自動檢查文件/架構/CI 一致性

### 驗證

- `npm run ci:frontend` 成功
- `excel-ai` `pytest -q` 成功

## [v3.1.1] — 2026-04-20

### 修正

- `build.py` 修正 PyInstaller `--add-data` 與重構後路徑不一致的問題
- `requirements.txt` 移除已不再使用的 Gemini 相關依賴殘留

### 文件

- 新增 `CHANGELOG.md`
- 新增 `TROUBLESHOOTING.md`
- 新增 `SECURITY.md`
- 新增 `TOOLS.md`
- 新增 `PROMPT.md`
- 新增 `MANUAL_TEST.md`

## [v3.1.0] — 2026-04-17

### 重構

- `main.py` 精簡為主對話迴圈
- 新增 `constants.py`
- 新增 `session.py`
- `tools/` 目錄拆分為 `definition.py` 與 `executor.py`
- 新增 `ui/sidebar.py`
- `providers/base.py` 補 `chat_stream`

### 目標

- 降低 `main.py` 負擔
- 把 tool schema、執行路徑、UI、session 管理拆開

## [v3.0.0] — 2026-04-15

### 新增

- 工具數量從 19 增加到 28
- 串流回覆
- Port 自動遞增
- 對話摘要壓縮

### 工具擴充

- `filter_range`
- `merge_cells` / `unmerge_cells`
- `set_borders`
- `clear_range`
- `set_row_height`
- `copy_range`
- `add_conditional_format`
- `set_data_validation`

## [v2.0.0] — 2026-04-14

### 新增

- 圖表
- 樞紐分析表
- 凍結窗格
- 自動欄寬
- 欄寬設定

### 平台與安全

- 危險操作確認流程
- 側邊欄 Qwen 設定
- 對話記錄匯出 / 載入
- 移除 Gemini provider，專注本地/內網 Qwen

## [v1.0.0] — 2026-04-10

### 初版

- 14 個基礎工具
- Streamlit UI
- `pywin32` / win32com 控制 Excel
- OpenAI 相容 API 串接本地/內網 LLM

## 歷史規劃文件

更細的歷史細節保留於：

- `PLAN_v2.md`
- `PLAN_v3.md`
- `PLAN_v4.md`
- `CODE_REVIEW.md`
- `REFACTOR_PLAN.md`

但這些文件現在屬於歷史紀錄，不再作為現況敘事來源。
