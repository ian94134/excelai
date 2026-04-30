# Excel AI Web UI 工具測試報告

日期：2026-04-30  
測試介面：Streamlit Web UI `http://localhost:8501/`  
UI 版本：v4.8.0  
測試方式：Codex in-app browser 模擬使用者輸入 + Excel COM 驗證實際活頁簿狀態  
測試活頁簿：`web_ui_v47_full.xlsx`、`web_ui_v47_second.xlsx`

## 總結

- 既有 71 個工具主流程已全部跑完。
- 新增 `beautify_range` 一鍵表格美化工具，目前工具總數為 72。
- 最終結果：72 PASS / 0 FAIL（既有 71 工具 + 新增 `beautify_range`）。
- 針對本輪修復與新增工具的 `pytest` 回歸測試：18 PASS / 0 FAIL。
- 重新用 Web UI 模擬使用者做短煙霧測試：`fill_series`、`query_range`、`beautify_range`、`write_range` + `undo_last` 全部 PASS。
- 追加完全白話輸入測試已自動化為 `tools_web_ui_smoke.py --case-set plain`：美化報表、查 North 總金額、寫入 `Report!AA1`。
- Web UI 已將一般使用者可見的工具名稱改為白話操作名稱；原始工具名保留在「技術細節」展開區。
- 新增可重跑的 Web UI smoke 流程與固定測試活頁簿 fixture。
- 測試中發現並修復多個「UI 顯示成功但 Excel 實際沒有變」的問題。
- 最後語法檢查通過：`agent.py`、`config.py`、`main.py`、`ui/sidebar.py`、`ui/tool_display.py`、`excel/data.py`、`excel/format.py`、`excel/_undo.py`、`excel_tools.py`、`tools/definition.py`、`tools/executor.py`、`backup.py`、`tools_web_ui_smoke.py`、`scripts/build_web_ui_smoke_fixture.py`、`tests/test_beautify_range_contract.py`、`tests/test_web_ui_regression_repairs.py`、`tests/test_web_ui_smoke_assets.py`、`tests/test_tool_display.py`。

## 修復項目

| 問題 | 修復 |
| --- | --- |
| `format_range` 模型只傳 `range_addr/sheet`，造成格式無動作 | `agent.py` 會從提示補回粗體、底色、字色、對齊 |
| `sort_range/filter_range` 欄位索引錯誤，曾破壞表頭或篩錯欄 | `agent.py` 依標題列把 `Amount/Region` 映射成正確 `column_index` |
| `freeze_panes` COM 失敗 | `excel/format.py` 增加 ActiveWindow 與 Excel4Macro fallback |
| `find_replace` 在篩選狀態下可能假成功 | `excel/data.py` 先清除篩選，改用 `UsedRange.Replace` 並回傳 `replaced` |
| `set_data_validation` 模型常送 `formula1/validation_type` | `excel/data.py` 接受別名，`agent.py` 會轉成 `options` |
| `fill_series` 起始格空白時假成功 | `excel/data.py` 要求 `start_value` 或報錯，`agent.py` 會解析 `1 到 5`、`填 1 到 5`、`start_value=1` |
| `advanced_filter` 不支援 `Criteria!A1:A2` | `excel/data.py` 支援 sheet-qualified range |
| `page_setup` 模型只傳 `sheet` 導致假成功 | `agent.py` 從提示補回方向、紙張、縮放、列印範圍、置中 |
| `query_range` 沒把自然語句轉成 filters/aggregation | `agent.py` 補 `Region=North` filter 與 `Amount` sum aggregation |
| `undo_last` 空白單格備份成 `[]`，還原無效 | `tools/executor.py`、`excel/_undo.py`、`backup.py` 修復空白單格/空白範圍還原 |
| 使用者想快速美化 Excel 成品時需要多次格式工具組合 | 新增 `beautify_range`，一次完成表頭、交錯列、框線、數字格式、欄寬與篩選 |
| 白話「整理漂亮」後模型又轉正式表格，可能造成 Excel 產生 `欄1/欄2` 並推移原表頭 | 收斂 `config.py` 與工具描述；一般美化只用 `beautify_range`，`apply_table_style` 僅限明確要求正式 Excel Table；並加上表頭偵測保護 |
| 工具已成功執行但模型最後沒有輸出文字時，Web UI 可能停在思考狀態 | `agent.py` 會用已成功的工具結果產生簡短完成訊息 |
| 一般使用者看不懂 `query_range` / `write_range` 等內部工具名 | 新增 `ui/tool_display.py`，狀態列、側邊欄、確認訊息與模型回覆會轉成白話操作名稱；技術細節仍可展開查看 |

## 覆蓋結果

| 區塊 | 工具 | 結果 |
| --- | --- | --- |
| 讀取 | `read_range`, `get_sheet_info`, `get_used_range`, `get_workbook_summary`, `list_workbooks` | PASS |
| 寫入與格式 | `write_range`, `format_range`, `beautify_range`, `set_borders`, `clear_range`, `merge_cells`, `unmerge_cells`, `add_conditional_format`, `apply_table_style`, `set_tab_color` | PASS |
| 列欄結構 | `insert_row`, `delete_row`, `insert_column`, `delete_column`, `set_row_height`, `set_column_width`, `auto_fit`, `freeze_panes`, `group_rows`, `group_columns` | PASS |
| 工作表 | `add_sheet`, `rename_sheet`, `move_sheet`, `copy_sheet`, `delete_sheet`, `protect_sheet`, `unprotect_sheet`, `set_print_titles`, `add_header_footer` | PASS |
| 資料處理 | `sort_range`, `filter_range`, `find_replace`, `trim_range`, `copy_range`, `add_comment`, `set_data_validation`, `name_range`, `transpose_range`, `fill_series`, `split_text_to_columns`, `add_subtotal`, `advanced_filter`, `summarize_range`, `find_duplicates` | PASS |
| 圖表與樞紐 | `create_chart`, `format_chart`, `move_chart`, `create_combo_chart`, `delete_chart`, `add_sparklines`, `create_pivot_table`, `refresh_pivot_table`, `format_pivot_table`, `add_slicer` | PASS |
| 頁面/圖片/巨集 | `page_setup`, `add_image`, `record_macro`, `list_macros`, `run_macro`, `delete_macro` | PASS |
| 公式與查詢 | `validate_formula`, `explain_formula`, `query_range` | PASS |
| 多工作簿/儲存/復原 | `switch_workbook`, `copy_range_between_workbooks`, `save_workbook`, `undo_last` | PASS |

## 代表性驗證

- `Report!A1:C1`：藍底白字、粗體、置中。
- `SalesData`：Amount 降冪排序、Region=North 篩選、`Beta` 替換成 `Beta-X`。
- `ValidationData!B2:B4`：資料驗證清單公式為 `A,B,C`。
- `Report!I1:I3`：`Date / Region / Product` 轉置成功。
- `Report!K1:K5`：填入 `1,2,3,4,5`。
- `Report!M1:O3`：逗號文字成功拆欄。
- `PivotOut`：樞紐分析表建立、刷新、樣式套用，並新增 Region slicer。
- `SalesData`：圖表、sparkline、頁面設定、圖片插入均由 COM 驗證。
- `MacroArea!A2:B2`：執行 `TestMacro` 後寫入 `MACRO / run`。
- `query_range`：`Region=North` 回 4 筆，`Amount` 加總 `500.0`。
- `beautify_range`：`SalesData!A1:G10` 套用藍色主題；COM 驗證表頭粗體/底色、交錯列底色、細框線、`Amount` 千分位格式與 AutoFilter。
- 白話美化：輸入「幫我把 SalesData 這張表整理得漂亮一點，做成可以給主管看的樣子。」後，Web UI 自行完成美化、標籤配色與儲存；COM 驗證 `SalesData` 仍為 `$A$1:$G$10`，表頭未被推移。
- 白話查詢：輸入「這張表裡 North 的總金額是多少？」後，Web UI 回覆 4 筆、合計 `500.0`。
- 白話寫入：輸入「幫我在 Report 工作表的 AA1 寫上『白話輸入OK』。」後，COM 驗證 `Report!AA1 = 白話輸入OK`。
- UI 友善顯示：輸入「這張表裡 North 的總金額是多少？」後，畫面狀態顯示「已讀取資料 / 已查詢資料」等白話操作名稱；可見 checklist 未再出現 `query_range` 或 `get_used_range`。
- `web_ui_v47_second.xlsx!Target!D1:F3`：成功複製主活頁簿 `SalesData!A1:C3`。
- `undo_last`：`Target!Z1` 從 `UNDO_MARK` 還原為空白。

## 本輪新增回歸測試

檔案：`tests/test_web_ui_regression_repairs.py`  
新增 Web UI smoke 規格測試：`tests/test_web_ui_smoke_assets.py`  
新增美化工具契約測試：`tests/test_beautify_range_contract.py`  
新增 UI 工具名友善顯示測試：`tests/test_tool_display.py`  
指令：`python -B -m pytest tests\test_tool_display.py tests\test_web_ui_smoke_assets.py tests\test_beautify_range_contract.py tests\test_web_ui_regression_repairs.py -q`  
結果：18 PASS / 0 FAIL

覆蓋範圍：

- `format_range` 從自然語句補回粗體、底色、字色、置中。
- `sort_range` / `filter_range` 依表頭補正欄位索引與篩選條件。
- `set_data_validation` 接受 `formula1` / `validation_type` 別名並轉成工具需要的參數。
- `fill_series` 解析「從 Report!K1 開始往下填 1 到 5」。
- `page_setup` 補回列印方向、紙張、縮放、列印範圍與水平置中。
- `query_range` 從自然語句建立 `Region=North` filter 與 `Amount` sum aggregation。
- `undo_last` 對空白單格備份會清空儲存格，而不是假還原。
- `beautify_range` schema、executor 派發與 backup 追蹤。
- `apply_table_style` schema 不再被描述成一般美化預設路徑。
- `_looks_like_header_row` 偵測唯一文字表頭，降低 `has_header=false` 時把原表頭推成資料列的風險。
- 工具成功但模型無後續文字時，`run_turn` 會回傳可讀完成訊息。
- 固定 fixture 的 `SalesData` 內 `Region=North` 為 4 筆，`Amount` 合計固定為 `500`。
- Web UI smoke case 順序固定為 `fill_series` → `query_range` → `beautify_range` → `write_range` → `undo_last`。
- 白話 smoke case 固定為「美化主管報表」→「查 North 總金額」→「寫入 Report!AA1」，且 prompt 不含工具名。
- `sanitize_assistant_text` 會把模型輸出的工具 checklist 轉成白話操作名稱。

## 本輪新增自動化檔案

| 檔案 | 作用 |
| --- | --- |
| `scripts/build_web_ui_smoke_fixture.py` | 產生固定測試活頁簿 `tests/fixtures/web_ui_smoke_base.xlsx` |
| `tests/fixtures/web_ui_smoke_base.xlsx` | Web UI smoke 的 golden workbook；包含 `Report` 與 `SalesData` |
| `tools_web_ui_smoke.py` | 可重跑的短版 Web UI smoke runner；支援 `--case-set tool/plain/all` |
| `tests/test_beautify_range_contract.py` | 驗證 `beautify_range` schema、executor 與 backup 設定 |
| `tests/test_web_ui_smoke_assets.py` | 不依賴瀏覽器的 smoke case 與 fixture 規格測試 |
| `tests/test_tool_display.py` | 驗證 UI 會把工具名轉成白話操作名稱 |

基本指令：

```powershell
python -B scripts\build_web_ui_smoke_fixture.py
python -B tools_web_ui_smoke.py --prepare-only
python -B tools_web_ui_smoke.py --url http://localhost:8501/
python -B tools_web_ui_smoke.py --case-set plain --workbook web_ui_plain_smoke_work.xlsx --url http://localhost:8501/
```

## 本輪 Web UI 短煙霧重跑

| 使用者操作 | UI 結果 | COM 驗證 |
| --- | --- | --- |
| `fill_series`：從 `Report!K1` 往下填 `1` 到 `5` | 回覆已填入 `K1:K5` | `Report!K1:K5 = 1,2,3,4,5` |
| `query_range`：查 `SalesData!A1:G10` 的 `Region=North` 並加總 `Amount` | 回覆 4 筆 North，總和 `500.0` | 與資料表內容一致 |
| `beautify_range`：美化 `SalesData!A1:G10`，主題 `blue` | 回覆已套用藍色主題美化 | 表頭粗體/底色、交錯列、框線、`#,##0`、AutoFilter 全部存在 |
| `write_range` + `undo_last`：先寫 `Report!Z1=SMOKE_UNDO`，再撤銷 | 回覆已還原上一個 `write_range` | `Report!Z1` 從 `SMOKE_UNDO` 回到空白 |

## 本輪白話輸入追加測試

測試活頁簿：`web_ui_natural_smoke_20260429_174527.xlsx`
UI 顯示驗證活頁簿：`web_ui_plain_smoke_ui.xlsx`

| 使用者白話輸入 | UI 結果 | COM 驗證 |
| --- | --- | --- |
| `幫我把 SalesData 這張表整理得漂亮一點，做成可以給主管看的樣子。` | 自行完成美化、標籤配色與儲存；未要求使用者提供工具名 | `SalesData` 維持 `$A$1:$G$10`，`A1=Date`、`A2=2026-01-01`，AutoFilter 與 `#,##0` 格式存在 |
| `這張表裡 North 的總金額是多少？` | 回覆 North 有 4 筆，合計 `500.0` | 原資料列未被推移，4 筆 North 合計正確 |
| `幫我在 Report 工作表的 AA1 寫上「白話輸入OK」。` | 回覆已寫入 `Report!AA1` | `Report!AA1 = 白話輸入OK` |
| `這張表裡 North 的總金額是多少？`（UI 顯示檢查） | 回覆合計 `500.0`，可見狀態為「已讀取資料 / 已查詢資料」 | 未在可見 checklist 顯示 `query_range` 或 `get_used_range` |

## 注意

- 外部 Python Playwright 未安裝；本輪 Web UI 短煙霧測試使用 Codex in-app browser 執行，沒有新增 Playwright 依賴。`tools_web_ui_smoke.py` 已做好 Playwright runner，但實際自動瀏覽器執行需要先安裝 Playwright。
- 測試過程曾為排查 Excel COM data validation 狀態，產生 `web_ui_v47_full_resaved.xlsx` 與 `tmp_openpyxl_dv_test.xlsx` 等臨時檔；未自動刪除。
- 目前工作區原本已有多個未追蹤/已刪除檔案，本次沒有還原或清理那些既有變更。
