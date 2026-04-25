# Excel AI 助手（`excel-ai` 子系統）

這份 README 只描述 `excel-main/excel-ai/` 這條 Python/Streamlit + `pywin32` 產品線。若你要看整個工作區、React/Electron 編輯器或歷史文件導覽，請先回到 [../README.md](../README.md) 與 [../DOCS.md](../DOCS.md)。

## 目前狀態

- 版本字串：`v4.0.0`
- LLM 端：OpenAI 相容介面，預設為本地/內網 Qwen
- 預設設定：
  - `QWEN_BASE_URL=http://140.96.96.16:8079/v1`
  - `QWEN_MODEL=Qwen-3.5-122B-A10B`
- Tool schema：**60 個**
- 危險工具：**6 個**
- 自動化驗證：
  - `pytest -q`：`33 passed`
  - GitHub Actions：前端 CI + Windows `excel-ai` 測試

## 這個子系統做什麼

`excel-ai` 會在 Windows 上透過 `pywin32` 直接控制「使用者已經開啟中的」Microsoft Excel。使用者在 Streamlit 聊天介面輸入自然語言後，LLM 會選擇對應工具，再由 `excel_tools.py` 執行 Excel 操作。

適合的場景：

- 真正操作桌面版 Excel 2021 活頁簿
- 排序、篩選、圖表、樞紐、框線、資料驗證等 Excel 實際功能
- 危險操作前要有確認流程
- 需要 `undo_last`、快照、操作紀錄與側邊欄控制

## 目前能力摘要

### Tool 層

- 60 個 OpenAI 相容 tools
- 6 個危險工具會在 UI 中要求使用者確認
- `undo_last` 已存在，支援可復原操作與不可復原狀況說明

### 側邊欄與 UI

- Qwen URL / Model 設定
- Excel 狀態顯示
- 備份堆疊與「↶ 復原上一步」
- 工作表快照 / 還原
- CSV 快速導入
- 任務規劃模式
- 唯讀操作紀錄
- 對話歷史儲存 / 載入 / 清除

### 穩健性與觀測

- `error_type` 結構化錯誤回傳
- JSON Lines 結構化日誌：`~/.excel-ai/logs/YYYY-MM-DD.jsonl`
- Launcher 日誌：`ExcelAI_log.txt`
- 重複 tool call 保護
- `pytest` 測試與 GitHub Actions

## 系統需求

| 項目 | 需求 |
|------|------|
| 作業系統 | Windows 10 / 11 |
| Excel | Microsoft Excel 2016 以上；需已開啟活頁簿 |
| Python | 3.10 以上（開發/測試環境） |
| LLM 端點 | 可存取的 OpenAI 相容 API；預設為 Qwen |

## 快速開始

### 1. 安裝依賴

```bash
pip install -r requirements.txt
```

若要跑測試：

```bash
pip install -r requirements-dev.txt
```

### 2. 設定 `.env`

```bash
copy .env.example .env
```

常用欄位：

```env
QWEN_BASE_URL=http://你的伺服器IP:埠號/v1
QWEN_MODEL=你的模型名稱
```

### 3. 開啟 Excel，再啟動 Streamlit

```bash
streamlit run main.py
```

瀏覽器會開到 `http://localhost:8501`；若埠號被占用，launcher/打包版本會自動嘗試 `8502~8509`。

## 測試與驗證

### Pytest

```bash
pytest -q
```

### Qwen 連線測試

```bash
python ..\\test_qwen\\test_01_basic_chat.py
python ..\\test_qwen\\test_02_tool_calling.py
```

### 手動整合測試

- [../MANUAL_TEST.md](../MANUAL_TEST.md)
- [../undo_test_guide.md](../undo_test_guide.md)

## 打包為 EXE

```bash
python build.py
```

打包後：

- EXE 主程式位於 `dist/ExcelAI/`
- 啟動器會寫 `ExcelAI_log.txt`
- 執行中的主流程會另外寫 `~/.excel-ai/logs/*.jsonl`

## 目錄結構

```text
excel-ai/
├── main.py
├── config.py
├── constants.py
├── excel_tools.py
├── backup.py
├── logger.py
├── session.py
├── launcher.py
├── build.py
├── exceptions.py
├── requirements.txt
├── requirements-dev.txt
├── providers/
│   ├── base.py
│   └── local_qwen.py
├── tools/
│   ├── definition.py
│   └── executor.py
├── ui/
│   └── sidebar.py
└── tests/
    ├── conftest.py
    ├── test_backup.py
    ├── test_session.py
    └── test_tools_schema.py
```

## 相關文件

- [../TOOLS.md](../TOOLS.md)：60 個 tool 速查
- [../ARCHITECTURE.md](../ARCHITECTURE.md)：工作區與 `excel-ai` 架構
- [../PROMPT.md](../PROMPT.md)：`SYSTEM_PROMPT` 說明
- [../TROUBLESHOOTING.md](../TROUBLESHOOTING.md)：排錯與日誌位置
- [../SECURITY.md](../SECURITY.md)：資料邊界與安全說明
- [../CHANGELOG.md](../CHANGELOG.md)：版本演進
