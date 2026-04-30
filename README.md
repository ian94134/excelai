# Excel AI Streamlit Runtime

這個資料夾是精簡後的 Excel AI Streamlit 版本，只保留執行 `streamlit run main.py` 需要的檔案。React/Electron 前端、測試、建置輸出、文件草稿、快取與打包檔案已移除。

## 版本

- 目前版本：`v4.8.0`
- 入口檔案：`main.py`
- 啟動方式：`streamlit run main.py`
- 工具數量：72 個 OpenAI-compatible tool schema
- 危險工具數量：6 個，需要明確確認後才會執行
- 預設模型服務：OpenAI-compatible Qwen endpoint
- 執行平台：Windows + Microsoft Excel + pywin32 COM

## 目錄結構

```text
excel-ai/
  main.py                  Streamlit UI 入口
  agent.py                 LLM 工具呼叫迴圈與事件流
  config.py                環境變數、模型設定、系統提示詞
  session.py               Streamlit 對話狀態與壓縮
  backup.py                操作前備份與 undo stack
  excel_tools.py           Excel tool re-export shim
  excel_event_watcher.py   Excel 選取範圍監聽
  excel_query.py           範圍查詢、排序、聚合
  formula_validator.py     公式驗證與公式說明
  macro.py                 巨集錄製、執行、管理
  logger.py                本機 JSONL log
  telemetry.py             本機 SQLite 工具統計
  constants.py             共用常數
  exceptions.py            結構化錯誤類型
  utils.py                 欄號、位址等小工具
  excel/                   pywin32 Excel 實作
  providers/               LLM provider
  tools/                   tool schema、executor、registry
  ui/                      Streamlit sidebar
  requirements.txt         runtime Python 依賴
  run.bat                  Windows 快速啟動腳本
```

## 架構

```text
User
  -> Streamlit UI (main.py, ui/sidebar.py)
  -> Agent loop (agent.py)
  -> LLM provider (providers/local_qwen.py)
  -> Tool schema + executor (tools/definition.py, tools/executor.py)
  -> Excel operations (excel_tools.py, excel/*.py)
  -> Microsoft Excel COM (pywin32 / win32com)
```

核心流程：

1. `main.py` 建立 Streamlit 介面，讀取目前 Excel 活頁簿、選取範圍與對話狀態。
2. 使用者輸入需求後，`agent.py` 呼叫 LLM，並依模型回傳的 tool calls 分段執行。
3. `tools/executor.py` 驗證工具名稱與參數，處理危險工具確認、log、telemetry、backup。
4. 實際 Excel 操作由 `excel/` 模組透過 `win32com` 操作目前開啟的 Excel。
5. 寫入或高風險動作會進入 `backup.py`，讓 `undo_last` 有機會回復。

## 主要模組

| 檔案 | 作用 |
|---|---|
| `main.py` | Streamlit app 入口、畫面渲染、使用者輸入、工具結果顯示 |
| `agent.py` | LLM streaming、工具呼叫迴圈、重複呼叫保護、錯誤事件 |
| `config.py` | `QWEN_BASE_URL`、`QWEN_MODEL`、`SYSTEM_PROMPT` |
| `tools/definition.py` | 72 個工具的 OpenAI-compatible schema |
| `tools/executor.py` | 工具派發、安全檢查、backup、telemetry |
| `excel/data.py` | 讀寫範圍、查找取代、清除、資料整理 |
| `excel/format.py` | 格式、表格樣式、條件格式、列印設定 |
| `excel/sheet.py` | 工作表、列欄、工作簿、保護、篩選 |
| `excel/chart.py` | 圖表、樞紐分析表、走勢圖、交叉分析篩選器 |
| `excel/_undo.py` | `undo_last` 的反向操作 |
| `backup.py` | 操作前快照與本機備份堆疊 |
| `macro.py` | 巨集錄製與重播 |
| `excel_query.py` | 自然語言資料查詢背後的結構化範圍查詢 |
| `formula_validator.py` | 公式檢查與公式解釋 |

## 工具能力

目前 `v4.8.0` 共 72 個工具，涵蓋：

- 活頁簿與工作表資訊
- 範圍讀寫、清除、複製、儲存
- 一鍵表格美化、格式化、框線、欄寬列高、表格樣式
- 列欄插入刪除、合併取消合併、凍結窗格
- 排序、篩選、進階篩選、移除重複、小計
- 圖表、組合圖、樞紐分析表、走勢圖、圖片
- 資料驗證、批註、具名範圍
- 巨集錄製與執行
- 公式驗證與公式解釋
- 自然語言範圍查詢
- 多工作簿切換與跨工作簿複製
- `undo_last`

危險工具清單：

```text
delete_row
delete_column
find_replace
clear_range
split_text_to_columns
delete_sheet
```

這些工具在 `tools/executor.py` 內強制要求 `confirm_dangerous=True`，不是只靠 UI 提醒。

## 環境設定

`.env` 支援以下設定：

```env
QWEN_BASE_URL=http://host:port/v1
QWEN_MODEL=Qwen-3.5-122B-A10B
```

不要把真實 API key 或內部憑證寫進文件。需要範例時請放在 `.env.example`。

## 安裝與啟動

在 Windows PowerShell：

```powershell
cd C:\Users\User\Desktop\excel-main\excel-main\excel-ai
python -m pip install -r requirements.txt
streamlit run main.py
```

如果 8501 已被占用，Streamlit 會提示或可指定其他 port：

```powershell
streamlit run main.py --server.port 8502
```

也可以使用：

```powershell
.\run.bat
```

## 使用前檢查

1. 先開啟 Microsoft Excel。
2. 開啟一個要操作的活頁簿。
3. 確認 `.env` 的模型服務可連線。
4. 啟動 Streamlit。
5. 在側邊欄確認 Excel 連線狀態與目前選取範圍。

## 本機資料

程式會在使用者家目錄建立本機資料：

```text
~/.excel-ai/logs/              JSONL 操作紀錄
~/.excel-ai/backup_stack.json  undo/backup 堆疊
~/.excel-ai/macros.json        巨集資料
~/.excel-ai/telemetry.db       本機工具使用統計
```

這些資料不屬於專案 runtime source，可以視需求清除。

## 維護原則

- `main.py` 是唯一 Streamlit 入口。
- 新增工具時要同步 `tools/definition.py`、`tools/executor.py` 與必要的 `excel/` 實作。
- 高風險寫入工具必須在 executor 層做確認，不只在 UI 層做提醒。
- 修改 Excel 寫入邏輯時，要確認 backup/undo 行為是否仍正確。
- 不要把 React/Electron 前端檔案放回這個精簡 runtime 目錄，除非目標重新改為完整雙產品 repo。

## 驗證

目前清理後已做過：

```text
AST OK: 32 files
streamlit run main.py -> HTTP 200
```

若後續修改程式碼，至少先做：

```powershell
python -B -c "import ast, pathlib; files=[p for p in pathlib.Path('.').rglob('*.py') if '.git' not in p.parts]; [ast.parse(p.read_text(encoding='utf-8'), filename=str(p)) for p in files]; print(f'AST OK: {len(files)} files')"
streamlit run main.py
```

### Web UI smoke 測試

固定測試活頁簿由下列指令產生：

```powershell
python -B scripts\build_web_ui_smoke_fixture.py
```

短版 Web UI smoke 會複製 `tests\fixtures\web_ui_smoke_base.xlsx` 成 `web_ui_smoke_work.xlsx`，開啟 Excel，確認 Streamlit health，然後用瀏覽器跑 `fill_series`、`query_range`、`beautify_range`、`write_range`、`undo_last`：

```powershell
python -B tools_web_ui_smoke.py --prepare-only
python -B tools_web_ui_smoke.py --url http://localhost:8501/
```

白話使用者 smoke 已固定在同一個 runner，可用 `--case-set plain` 或 `--plain-language` 重跑，不需要在 prompt 裡說工具名：

```powershell
python -B tools_web_ui_smoke.py --case-set plain --workbook web_ui_plain_smoke_work.xlsx --url http://localhost:8501/
```

`tools_web_ui_smoke.py` 需要 Playwright 才能自動操作瀏覽器；一般 `pytest` 不依賴 Playwright，只驗證 fixture、工具名 smoke case、白話 smoke case 與 UI 友善顯示規格。

## 常見問題

### Streamlit 開了但抓不到 Excel

確認 Excel 已經開啟，且至少有一個活頁簿。此工具主要透過 `GetActiveObject` 連到既有 Excel instance。

### 工具顯示需要確認

代表模型準備執行危險工具。確認前請檢查範圍、工作表名稱與影響內容。

### 模型沒有呼叫工具

先確認 `.env` 的 `QWEN_BASE_URL` 與 `QWEN_MODEL`。此程式使用 OpenAI-compatible `/v1/chat/completions` 介面。

### 操作後想復原

在對話中輸入「復原」「undo」「還原上一步」會觸發 `undo_last`。並非所有 Excel 操作都可完整還原，回傳 `cannot_undo` 時會說明原因。
