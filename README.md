# Excel AI 助手 `v4.7.0`

Streamlit + `pywin32` 自然語言控制桌面版 Excel 的 AI 助手。

## 目前狀態

| 項目 | 內容 |
|------|------|
| 版本 | `v4.7.0` |
| LLM 端 | OpenAI 相容介面，預設本地 Qwen |
| 預設端點 | `http://140.96.96.16:8079/v1` |
| 工具數量 | **71 個**（含 6 個危險工具需確認） |
| 測試套件 | `pytest -q` → **249 passed** |

## 系統需求

| 項目 | 需求 |
|------|------|
| 作業系統 | Windows 10 / 11 |
| Excel | Microsoft Excel 2016 以上；需已開啟活頁簿 |
| Python | 3.10 以上 |
| LLM 端點 | OpenAI 相容 API（預設 Qwen） |

## 快速開始

```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 設定 .env
copy .env.example .env
# 編輯 QWEN_BASE_URL 與 QWEN_MODEL

# 3. 開啟 Excel，再啟動
streamlit run main.py
```

測試：

```bash
pip install -r requirements-dev.txt
pytest -q
```

## 功能摘要

### 工具層（71 個）

- **讀取**：`read_range`、`get_sheet_info`、`get_used_range`、`get_workbook_summary`
- **寫入**：`write_range`、`save_workbook`
- **格式**：`format_range`、`set_borders`、`merge_cells`、`unmerge_cells`、`clear_range`、`add_conditional_format`
- **列/欄**：`insert_row`、`delete_row`、`insert_column`、`delete_column`、`set_row_height`
- **工作表**：`add_sheet`、`rename_sheet`、`delete_sheet`、`move_sheet`、`copy_sheet`、`protect_sheet`、`unprotect_sheet`、`set_print_titles`、`add_header_footer`
- **資料操作**：`sort_range`、`find_replace`、`trim_range`、`filter_range`、`copy_range`、`set_data_validation`、`add_comment`
- **圖表**：`create_chart`、`create_combo_chart`、`format_chart`、`delete_chart`、`move_chart`
- **樞紐**：`create_pivot_table`、`refresh_pivot_table`、`format_pivot_table`
- **視窗/外觀**：`freeze_panes`、`auto_fit`、`set_column_width`、`set_tab_color`
- **美化**：`apply_table_style`、`add_sparklines`、`page_setup`、`add_image`、`add_slicer`
- **分析**：`summarize_range`、`find_duplicates`、`fill_series`、`group_rows`、`group_columns`、`transpose_range`、`name_range`、`add_subtotal`、`advanced_filter`、`split_text_to_columns`
- **Undo**：`undo_last`
- **巨集**：`record_macro`、`list_macros`、`run_macro`、`delete_macro`
- **公式智慧**：`validate_formula`、`explain_formula`
- **自然語言查詢**：`query_range`
- **多工作簿**：`list_workbooks`、`switch_workbook`、`copy_range_between_workbooks`

### 側邊欄功能

- Qwen URL / Model 設定
- 即時 Excel 選取範圍顯示與自動注入
- BackupStack 備份與「↶ 復原上一步」（跨 rerun 持久化）
- 工作表快照 / 還原
- CSV 快速導入
- 任務規劃模式（自動偵測複雜多步驟任務）
- 操作紀錄側邊欄
- 對話歷史儲存 / 載入 / 清除

### 穩健性

- 結構化錯誤回傳（`error_type` + `hint` + `suggested_next`）
- 多輪澄清機制（`EVT_CLARIFY`）
- 重複 tool call 保護（同簽章 ≥ 3 次自動中止）
- Tool result 自動壓縮（長對話 context 節省 30–50%）
- JSON Lines 結構化日誌：`~/.excel-ai/logs/YYYY-MM-DD.jsonl`

## 打包為 EXE

```bash
python build.py
```

打包輸出位於 `dist/ExcelAI/`。

## 目錄結構

```text
excel-ai/
├── main.py                  # Streamlit 入口
├── agent.py                 # LLM tool-calling generator loop
├── config.py
├── constants.py
├── session.py               # 對話歷史管理
├── compress.py              # Tool result 壓縮
├── backup.py                # BackupStack + 持久化
├── excel_tools.py           # excel/ 子套件的 shim
├── excel_event_watcher.py   # 背景選取範圍監聽
├── excel_query.py           # 自然語言查詢
├── formula_validator.py     # 公式驗證與解釋
├── macro.py                 # 巨集錄製與重播
├── logger.py
├── telemetry.py
├── utils.py
├── exceptions.py
├── launcher.py
├── build.py
├── requirements.txt
├── requirements-dev.txt
├── excel/
│   ├── __init__.py
│   ├── _base.py             # COM helpers
│   ├── _undo.py             # undo_last
│   ├── data.py              # 讀寫、排序、分析
│   ├── format.py            # 格式、邊框、列高欄寬
│   ├── sheet.py             # 工作表、多工作簿
│   └── chart.py             # 圖表、樞紐
├── providers/
│   ├── base.py              # LLMProvider ABC
│   └── local_qwen.py
├── tools/
│   ├── definition.py        # 71 個 OpenAI tool schema
│   ├── executor.py          # TOOL_MAP + execute / execute_batch
│   └── registry.py          # @register_tool decorator
├── ui/
│   └── sidebar.py
└── tests/
    ├── conftest.py
    ├── test_agent.py
    ├── test_agent_clarify.py
    ├── test_backup.py
    ├── test_backup_persist.py
    ├── test_complexity.py
    ├── test_executor_batch.py
    ├── test_formula_validator.py
    ├── test_macro.py
    ├── test_main_helpers.py
    ├── test_providers.py
    ├── test_query_range.py
    ├── test_selection_inject.py
    ├── test_session.py
    ├── test_session_compress.py
    └── test_tools_schema.py
```

## 版本紀錄

詳見 [CHANGELOG.md](CHANGELOG.md)。
