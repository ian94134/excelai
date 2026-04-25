"""所有 tool 以 OpenAI 格式定義（單一事實來源）。"""

OPENAI_TOOLS = [

    # ══════════════════════════════════════════════════════════════════════════
    # 讀取工具
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "read_range",
            "description": (
                "讀取 Excel 指定範圍的儲存格值，回傳二維陣列。\n"
                "使用時機：需要確認資料內容、欄名、現有值時使用。\n"
                "注意：大量資料（超過 200 列）會讓 AI context 變大，請先用 get_used_range 確認範圍再決定要讀多少。\n"
                "確認欄名時只讀第一列（如 'A1:Z1'），不要整張表一次讀完。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "範圍位址，如 'A1:D10' 或 'A1:Z1'（只讀標題列）"},
                    "sheet": {"type": "string", "description": "工作表名稱；省略則用作用中工作表"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "get_sheet_info",
            "description": (
                "取得活頁簿的基本狀態：檔案名稱、所有工作表清單、作用中工作表名稱、目前選取範圍。\n"
                "使用時機：每次新對話的第一個動作；需要確認工作表是否存在；需要知道有幾張工作表時。\n"
                "這是最輕量的狀態查詢，應優先使用，不要用其他工具替代。"
            ),
            "parameters": {"type": "object", "properties": {}},
        },
    },

    {
        "type": "function",
        "function": {
            "name": "get_used_range",
            "description": (
                "回傳工作表中有資料的範圍位址（如 '$A$1:$H$50'）。\n"
                "使用時機：\n"
                "  - 在 read_range 大量讀取前，確認資料邊界\n"
                "  - 在 format_range / set_borders / clear_range 前，確認要操作的完整範圍\n"
                "  - 在 create_chart / create_pivot_table 前，確認來源資料範圍\n"
                "注意：空白工作表會回傳 '$A$1'，不代表有資料，需搭配實際資料確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {"sheet": {"type": "string", "description": "工作表名稱；省略則用作用中工作表"}},
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "get_workbook_summary",
            "description": (
                "一次取得整本活頁簿所有工作表的摘要（名稱、已用範圍、列欄數、前 10 個標題名稱）。\n"
                "使用時機：\n"
                "  - 任何需要跨多張工作表操作的任務開始前\n"
                "  - 使用者說「幫我整理這份檔案」但未指定工作表時\n"
                "  - 需要一次了解整本活頁簿結構，避免逐張呼叫 get_sheet_info + get_used_range\n"
                "優點：單次呼叫取得全局資訊，token 消耗遠低於逐張查詢。\n"
                "注意：只回傳結構摘要，不回傳實際資料值，需要資料請再用 read_range。"
            ),
            "parameters": {"type": "object", "properties": {}},
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 寫入工具
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "write_range",
            "description": (
                "寫入值或公式到 Excel 指定範圍。\n"
                "values 為二維陣列，如 [[1, 2], [3, 4]] 或 [['姓名', '金額'], ['王小明', 5000]]。\n"
                "公式請以 = 開頭（如 =SUM(A1:A10)、=IF(B2>60,'通過','未通過')），函數名稱用英文。\n"
                "起始位址自動根據 values 陣列大小計算寫入範圍，range_addr 只需填左上角起始格（如 'A1'）。\n"
                "建議每次不超過 500 格，大量資料請分批寫入。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "起始儲存格或範圍，如 'A1' 或 'B2:D5'"},
                    "values": {"type": "array", "description": "二維陣列，如 [[1,2],[3,4]] 或 [['=SUM(A1:A5)']]"},
                    "sheet": {"type": "string", "description": "工作表名稱；省略則用作用中工作表"},
                },
                "required": ["range_addr", "values"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "save_workbook",
            "description": (
                "儲存目前作用中的活頁簿（等同於 Ctrl+S）。\n"
                "使用時機：完成重要的寫入、格式修改、新增工作表後主動呼叫。\n"
                "建議在每次完整任務結束時呼叫，不要等用戶提醒。"
            ),
            "parameters": {"type": "object", "properties": {}},
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 格式工具
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "format_range",
            "description": (
                "設定儲存格的字型、顏色、背景色、數字格式、對齊方式。\n"
                "至少要提供一項格式屬性（bold/italic/color/fill/font_size/number_format/horizontal_alignment）。\n"
                "常用數字格式：'#,##0'（整數千分位）、'#,##0.00'（兩位小數）、'0%'（百分比）、'YYYY/MM/DD'（日期）。\n"
                "顏色用 #RRGGBB：#FFFF00=黃、#FF0000=紅、#00B050=綠、#0070C0=藍、#FFFFFF=白、#000000=黑。\n"
                "標題列常用組合：bold=true, fill='#4472C4', color='#FFFFFF'（白字藍底）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":          {"type": "string", "description": "目標範圍，如 'A1:H1'"},
                    "bold":                {"type": "boolean", "description": "true=粗體、false=取消粗體"},
                    "italic":              {"type": "boolean", "description": "true=斜體"},
                    "color":               {"type": "string", "description": "字色 #RRGGBB，如 '#FFFFFF'"},
                    "fill":                {"type": "string", "description": "背景色 #RRGGBB，如 '#4472C4'"},
                    "font_size":           {"type": "number", "description": "字體大小，如 12、14"},
                    "number_format":       {"type": "string", "description": "數字格式字串，如 '#,##0.00' 或 'YYYY/MM/DD'"},
                    "horizontal_alignment":{"type": "string", "enum": ["left", "center", "right"], "description": "水平對齊"},
                    "sheet":               {"type": "string"},
                },
                "required": ["range_addr"],
                "anyOf": [
                    {"required": ["bold"]},
                    {"required": ["italic"]},
                    {"required": ["color"]},
                    {"required": ["fill"]},
                    {"required": ["font_size"]},
                    {"required": ["number_format"]},
                    {"required": ["horizontal_alignment"]},
                ],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_borders",
            "description": (
                "設定儲存格框線的樣式與顏色。\n"
                "sides 說明：\n"
                "  all=全部框線（最常用）、outer=只畫外框、inner=只畫內部格線\n"
                "  left/top/bottom/right=單邊\n"
                "style 說明：thin=細線（最常用）、medium=中線、thick=粗線、dashed=虛線\n"
                "常見用法：整張資料表加細外框 → sides='outer', style='medium'；內部加細格線 → sides='all', style='thin'。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "目標範圍"},
                    "style": {
                        "type": "string",
                        "enum": ["thin", "medium", "thick", "dashed"],
                        "description": "框線粗細；預設 thin",
                    },
                    "color": {"type": "string", "description": "框線顏色 #RRGGBB；預設黑色 #000000"},
                    "sides": {
                        "type": "string",
                        "enum": ["all", "outer", "inner", "left", "top", "bottom", "right"],
                        "description": "套用位置；預設 all",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_conditional_format",
            "description": (
                "新增條件格式化規則：符合條件的儲存格自動上色。\n"
                "condition_type 說明：\n"
                "  greater = 大於 value / less = 小於 value / equal = 等於 value\n"
                "  between = 介於兩值之間（value 填 '[最小, 最大]'，如 '[60, 80]'）\n"
                "  contains = 包含文字（value 填要搜尋的字串）\n"
                "注意：此工具會先清除該範圍既有的條件格式再新增。若要保留舊規則，請告知用戶。\n"
                "常用顏色：綠色 #92D050、黃色 #FFFF00、紅色 #FF0000、橘色 #FFC000。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要套用的儲存格範圍"},
                    "condition_type": {
                        "type": "string",
                        "enum": ["greater", "less", "equal", "between", "contains"],
                    },
                    "value": {
                        "type": "string",
                        "description": "條件值。between 時填 '[最小, 最大]'，如 '[60, 80]'；其他類型填數字或字串",
                    },
                    "fill_color": {"type": "string", "description": "符合條件時的背景色 #RRGGBB"},
                    "font_color": {"type": "string", "description": "符合條件時的字色 #RRGGBB（選填）"},
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr", "condition_type", "value", "fill_color"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "merge_cells",
            "description": (
                "合併儲存格並自動水平置中。\n"
                "常見用途：合併標題行（如 A1:H1 合併作為整張表的大標題）。\n"
                "注意：合併後只保留左上角儲存格的值，其餘值會遺失。如果合併範圍內有多個值，請先確認用戶是否知道。\n"
                "不可對已合併的儲存格重複合併，需先 unmerge_cells。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要合併的範圍，如 'A1:H1'"},
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "unmerge_cells",
            "description": (
                "取消合併儲存格，還原為獨立的儲存格。\n"
                "注意：取消合併後，原本的值只會出現在左上角儲存格，其他格為空白。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要取消合併的範圍"},
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "clear_range",
            "description": (
                "清除儲存格的內容或格式。\n"
                "target 說明：\n"
                "  values = 只清內容，保留格式（如背景色、框線）\n"
                "  formats = 只清格式，保留數值\n"
                "  all = 全部清除（內容+格式+條件格式）\n"
                "range_addr 省略時清除整張工作表的已使用範圍（危險操作，需用戶確認）。\n"
                "⚠️ 此為危險工具，執行前系統會請用戶確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要清除的範圍；省略則清除整張表的 UsedRange"},
                    "target": {
                        "type": "string",
                        "enum": ["values", "formats", "all"],
                        "description": "清除目標；預設 values",
                    },
                    "sheet": {"type": "string"},
                },
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 列 / 欄操作
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "insert_row",
            "description": (
                "在指定行號之前插入空白列。\n"
                "index=1 表示在第 1 列之前插入（工作表最頂端）。\n"
                "index=5, count=2 表示在第 5 列之前插入 2 列（原本 5 列以後的資料往下移）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "插入位置的行號（從 1 開始），在此行之前插入"},
                    "count": {"type": "integer", "description": "插入幾列；預設 1"},
                    "sheet": {"type": "string"},
                },
                "required": ["index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "delete_row",
            "description": (
                "刪除指定行號的列。\n"
                "index=3, count=2 表示刪除第 3 列和第 4 列（共 2 列）。\n"
                "⚠️ 此為危險工具，執行前系統會請用戶確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "要刪除的起始行號（從 1 開始）"},
                    "count": {"type": "integer", "description": "要刪除幾列；預設 1"},
                    "sheet": {"type": "string"},
                },
                "required": ["index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "insert_column",
            "description": (
                "在指定欄號之前插入空白欄。\n"
                "index=1 是在 A 欄之前插入（最左端）。\n"
                "index=3 是在 C 欄之前插入（原 C 欄變 D 欄）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "插入位置的欄號（從 1 開始，1=A、2=B…），在此欄之前插入"},
                    "count": {"type": "integer", "description": "插入幾欄；預設 1"},
                    "sheet": {"type": "string"},
                },
                "required": ["index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "delete_column",
            "description": (
                "刪除指定欄號的欄。\n"
                "index=2 表示刪除 B 欄（原 C 欄變 B 欄）。\n"
                "⚠️ 此為危險工具，執行前系統會請用戶確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "要刪除的起始欄號（從 1 開始，1=A、2=B…）"},
                    "count": {"type": "integer", "description": "要刪除幾欄；預設 1"},
                    "sheet": {"type": "string"},
                },
                "required": ["index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_row_height",
            "description": (
                "設定指定列的高度（Excel 點數單位）。\n"
                "預設列高約 15 點。常用值：20~25（較寬鬆）、30~40（大字體或含圖片）。\n"
                "通常搭配 format_range 設定大字體後使用，確保文字不被截斷。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "row_index": {"type": "integer", "description": "列號（從 1 開始）"},
                    "height":    {"type": "number",  "description": "高度值（點數），建議範圍 15~60"},
                    "sheet":     {"type": "string"},
                },
                "required": ["row_index", "height"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 工作表操作
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "add_sheet",
            "description": (
                "新增一張工作表到活頁簿。\n"
                "注意：如果工作表名稱已存在會報錯。執行前先用 get_sheet_info 確認名稱是否重複。\n"
                "新工作表會插入在所有工作表的最後方。"
            ),
            "parameters": {
                "type": "object",
                "properties": {"name": {"type": "string", "description": "新工作表名稱，不可與現有工作表重複"}},
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "delete_sheet",
            "description": (
                "刪除指定工作表。⚠️ 危險操作：刪除後無法復原，執行前系統會請使用者確認。\n"
                "活頁簿至少需保留一張工作表，否則操作失敗。\n"
                "建議先用 copy_sheet 備份後再刪除。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要刪除的工作表名稱"},
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "move_sheet",
            "description": (
                "移動工作表到指定位置（調整工作表索引標籤的順序）。\n"
                "before：移動到此工作表之前；after：移動到此工作表之後。\n"
                "before 與 after 二選一；都省略則移到最後。\n"
                "常用場景：把摘要表移到最前面，原始資料表移到最後。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name":   {"type": "string", "description": "要移動的工作表名稱"},
                    "before": {"type": "string", "description": "移動到此工作表之前（二選一）"},
                    "after":  {"type": "string", "description": "移動到此工作表之後（二選一）"},
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "copy_sheet",
            "description": (
                "複製工作表（含所有資料與格式）到新工作表。\n"
                "new_name：複製後的名稱；省略則 Excel 自動命名（如 '工作表1 (2)'）。\n"
                "before / after：插入位置；都省略則複製到最後。\n"
                "常用場景：\n"
                "  - 修改前先備份原始工作表（copy 後再編輯副本）\n"
                "  - 從模板工作表複製出新的月份報表"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name":     {"type": "string", "description": "要複製的來源工作表名稱"},
                    "new_name": {"type": "string", "description": "複製後的新名稱（選填；省略則自動命名）"},
                    "before":   {"type": "string", "description": "插入到此工作表之前（選填）"},
                    "after":    {"type": "string", "description": "插入到此工作表之後（選填）"},
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_print_titles",
            "description": (
                "設定列印時每頁自動重複的標題列或標題欄。\n"
                "長表格列印時必備：不設定的話第 2 頁以後看不到欄名。\n"
                "rows：列範圍，格式必須含 $，如 '$1:$1'（只重複第 1 列）或 '$1:$2'。\n"
                "columns：欄範圍，格式如 '$A:$A'（只重複 A 欄）。\n"
                "省略某項 = 不修改那項的設定。\n"
                "搭配 page_setup 一起用，可完整設定列印版面。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "rows":    {"type": "string", "description": "重複的標題列，如 '$1:$1' 或 '$1:$2'"},
                    "columns": {"type": "string", "description": "重複的標題欄，如 '$A:$A' 或 '$A:$B'"},
                    "sheet":   {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_header_footer",
            "description": (
                "設定工作表的列印頁首與頁尾文字。\n"
                "header / footer：頁首/頁尾置中文字（最常用）。\n"
                "或分別用 left/center/right 精確控制三個區塊。\n"
                "Excel 特殊碼（直接寫在文字中）：\n"
                "  &P = 目前頁碼　&N = 總頁數　&D = 日期　&T = 時間\n"
                "  &F = 檔案名稱　&A = 工作表名稱　&B = 粗體開關\n"
                "常用範例：\n"
                "  footer='第 &P 頁，共 &N 頁'（頁碼置中）\n"
                "  left_footer='&D'（左下角顯示日期）right_footer='&P/&N'（右下角頁碼）"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "header":         {"type": "string", "description": "頁首置中文字（最常用）"},
                    "footer":         {"type": "string", "description": "頁尾置中文字（最常用）"},
                    "left_header":    {"type": "string", "description": "頁首左區文字"},
                    "center_header":  {"type": "string", "description": "頁首中區文字"},
                    "right_header":   {"type": "string", "description": "頁首右區文字"},
                    "left_footer":    {"type": "string", "description": "頁尾左區文字"},
                    "center_footer":  {"type": "string", "description": "頁尾中區文字"},
                    "right_footer":   {"type": "string", "description": "頁尾右區文字"},
                    "sheet":          {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "protect_sheet",
            "description": (
                "保護工作表，防止使用者誤改公式、格式或結構。\n"
                "password：保護密碼（省略=無密碼，任何人可解除）。\n"
                "預設：只允許選取儲存格，其餘操作全部鎖定。\n"
                "常見用法：\n"
                "  1. 先用 format_range 解鎖輸入欄（Excel 預設所有格都是 locked=True）\n"
                "  2. 再呼叫 protect_sheet → 只有解鎖的欄可以輸入\n"
                "allow_sort=true + allow_filter=true 適合保護公式但允許使用者排序篩選。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "password":              {"type": "string",  "description": "保護密碼（選填；省略=無密碼）"},
                    "allow_select_locked":   {"type": "boolean", "description": "允許選取被鎖定的格；預設 true"},
                    "allow_select_unlocked": {"type": "boolean", "description": "允許選取未鎖定的格；預設 true"},
                    "allow_format_cells":    {"type": "boolean", "description": "允許修改格式；預設 false"},
                    "allow_insert_rows":     {"type": "boolean", "description": "允許插入列；預設 false"},
                    "allow_delete_rows":     {"type": "boolean", "description": "允許刪除列；預設 false"},
                    "allow_sort":            {"type": "boolean", "description": "允許排序；預設 false"},
                    "allow_filter":          {"type": "boolean", "description": "允許篩選；預設 false"},
                    "sheet":                 {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "unprotect_sheet",
            "description": (
                "解除工作表保護，讓所有儲存格可以再次編輯。\n"
                "password：如果保護時有設定密碼，需提供正確密碼才能解除。\n"
                "密碼錯誤會回傳錯誤，不會強制解除。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "password": {"type": "string", "description": "保護密碼（若有設定才需填）"},
                    "sheet":    {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "rename_sheet",
            "description": (
                "重新命名工作表。\n"
                "old_name 必須與現有工作表名稱完全一致（大小寫敏感）。\n"
                "建議先用 get_sheet_info 確認確切的工作表名稱後再呼叫。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "old_name": {"type": "string", "description": "現有工作表名稱（需完全一致）"},
                    "new_name": {"type": "string", "description": "新名稱"},
                },
                "required": ["old_name", "new_name"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 資料操作
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "sort_range",
            "description": (
                "對指定範圍依某欄排序。\n"
                "⚠️ column_index 是「範圍內的第幾欄」，不是工作表欄號。\n"
                "   例如範圍 C1:F100，要依 D 欄排序則 column_index=2（D 是範圍內第 2 欄）。\n"
                "has_header=true（預設）：第一列視為標題，不參與排序。\n"
                "ascending=true：由小到大（A→Z）；ascending=false：由大到小（Z→A）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":   {"type": "string", "description": "要排序的資料範圍（含標題列），如 'A1:D100'"},
                    "column_index": {"type": "integer", "description": "依此欄排序（範圍內第幾欄，從 1 開始）"},
                    "ascending":    {"type": "boolean", "description": "true=升冪（A→Z、小→大）；false=降冪；預設 true"},
                    "has_header":   {"type": "boolean", "description": "第一列是否為標題列；預設 true"},
                    "sheet":        {"type": "string"},
                },
                "required": ["range_addr", "column_index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "find_replace",
            "description": (
                "在工作表中尋找文字並全部取代。\n"
                "⚠️ 此工具會取代工作表內所有符合的文字，無法復原，為危險操作，執行前系統會請用戶確認。\n"
                "取代範圍為整張工作表（不限定某個範圍）。\n"
                "常用場景：批次修正錯字、統一格式（如「台北市」→「台北」）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "find":    {"type": "string", "description": "要尋找的文字"},
                    "replace": {"type": "string", "description": "取代後的文字"},
                    "sheet":   {"type": "string"},
                },
                "required": ["find", "replace"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "trim_range",
            "description": (
                "清除儲存格中多餘的空格（等同 Excel TRIM 函數）。\n"
                "處理範圍：移除開頭與結尾的空格、將文字中間的連續空格壓縮為單一空格。\n"
                "只處理文字型儲存格，數值與公式不受影響，不會破壞資料。\n"
                "常見用途：清理從外部系統匯入的資料（職稱、姓名、地址欄位常含多餘空格）。\n"
                "建議：整欄或整個資料範圍直接套用，不需逐格處理；"
                "勿用 find_replace 多輪替換來模擬 TRIM，效率差且無法處理尾隨空格。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要清理的儲存格範圍，如 'G2:G100' 或 'A1:Z50'"},
                    "sheet":      {"type": "string", "description": "工作表名稱；省略則用作用中工作表"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "filter_range",
            "description": (
                "對範圍套用 AutoFilter 篩選，只顯示符合條件的列。\n"
                "⚠️ column_index 是「範圍內的第幾欄」，不是工作表欄號。\n"
                "   例如範圍 A1:E100，要篩選 C 欄（地區）則 column_index=3。\n"
                "criteria 省略時：套用篩選但顯示全部（相當於清除篩選）。\n"
                "criteria 支援：完全符合（'台北'）、比較運算（'>100'、'<=50'、'<>0'）。\n"
                "套用前建議先清除既有篩選（criteria 省略呼叫一次），避免多重篩選互相干擾。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":   {"type": "string",  "description": "資料範圍（含標題列），如 'A1:E100'"},
                    "column_index": {"type": "integer", "description": "要篩選的欄（範圍內第幾欄，從 1 開始）"},
                    "criteria":     {"type": "string",  "description": "篩選條件，如 '台北' 或 '>100'；省略=清除篩選顯示全部"},
                    "sheet":        {"type": "string"},
                },
                "required": ["range_addr", "column_index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "copy_range",
            "description": (
                "複製儲存格範圍到另一位置（可跨工作表）。\n"
                "dest_sheet 不存在時自動建立。\n"
                "dest_range 只需填目標左上角的起始格（如 'A1'），不用填完整範圍。\n"
                "複製包含值、公式與格式；若只想複製值請先用 read_range 讀出再 write_range 寫入。\n"
                "常用場景：備份原始資料到另一張表、跨表彙整資料。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_range": {"type": "string", "description": "來源範圍，如 'A1:D50'"},
                    "dest_range":   {"type": "string", "description": "目標起始格，如 'A1'"},
                    "source_sheet": {"type": "string", "description": "來源工作表名稱；省略=目前工作表"},
                    "dest_sheet":   {"type": "string", "description": "目標工作表名稱；省略=目前工作表；不存在時自動建立"},
                },
                "required": ["source_range", "dest_range"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_comment",
            "description": (
                "在儲存格加入批注（右上角小紅三角，滑鼠移上去顯示說明文字）。\n"
                "批注不佔版面，適合為公式、假設值或注意事項留下說明。\n"
                "若儲存格已有批注，會自動覆蓋。\n"
                "visible=true：批注框常駐顯示（適合重要警示）。\n"
                "常見用途：\n"
                "  - 解釋公式邏輯（如「此處使用 VLOOKUP 匹配產品代碼」）\n"
                "  - 標記假設值（如「假設年增長率 5%，請依實際調整」）\n"
                "  - 給填表人說明（如「請填入含稅金額」）"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string",  "description": "要加批注的儲存格，如 'B1' 或 'C5'"},
                    "comment":    {"type": "string",  "description": "批注內容文字"},
                    "author":     {"type": "string",  "description": "作者名稱（選填；省略=Excel 登入使用者名稱）"},
                    "visible":    {"type": "boolean", "description": "是否常駐顯示批注框；預設 false（滑鼠移上才顯示）"},
                    "sheet":      {"type": "string"},
                },
                "required": ["range_addr", "comment"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_data_validation",
            "description": (
                "設定儲存格的下拉選單資料驗證，限制用戶只能輸入清單中的值。\n"
                "options 兩種格式：\n"
                "  1. 逗號分隔字串：'是,否,待定'（最多建議 20 個選項）\n"
                "  2. Excel 範圍參照：'$E$1:$E$10'（引用工作表現有清單，項目可動態增加）\n"
                "title/message：選填，用戶點選儲存格時顯示的提示說明。\n"
                "設定後，輸入不在清單中的值會彈出錯誤訊息。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要套用驗證的儲存格範圍"},
                    "options":    {"type": "string", "description": "選項清單（逗號分隔，如 '是,否'）或範圍參照（如 '$E$1:$E$5'）"},
                    "title":      {"type": "string", "description": "提示框標題（選填）"},
                    "message":    {"type": "string", "description": "提示框說明（選填）"},
                    "sheet":      {"type": "string"},
                },
                "required": ["range_addr", "options"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 圖表
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "delete_chart",
            "description": (
                "刪除工作表上指定的圖表。\n"
                "chart_index：第幾個圖表（1-based）；預設刪除第 1 個。\n"
                "刪除前可先呼叫 get_sheet_info 確認工作表名稱，不過目前無法列出圖表清單，"
                "只能依索引操作。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "chart_index": {"type": "integer", "description": "要刪除的圖表索引（1-based）；預設 1"},
                    "sheet":       {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "move_chart",
            "description": (
                "移動圖表位置或調整圖表大小（點數單位）。\n"
                "chart_index：第幾個圖表（1-based）；預設第 1 個。\n"
                "left/top：距工作表左上角的距離（點數）。\n"
                "width/height：圖表寬高（點數）；常用值：width=400~600, height=250~350。\n"
                "省略任何參數則保持原值不變，可只調整位置不改大小。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "chart_index": {"type": "integer", "description": "圖表索引（1-based）；預設 1"},
                    "left":        {"type": "number",  "description": "左邊距（點數）"},
                    "top":         {"type": "number",  "description": "上邊距（點數）"},
                    "width":       {"type": "number",  "description": "圖表寬度（點數）"},
                    "height":      {"type": "number",  "description": "圖表高度（點數）"},
                    "sheet":       {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "create_chart",
            "description": (
                "在工作表中依指定資料範圍建立圖表，自動放在資料右側。\n"
                "執行前必須確認：\n"
                "  1. range_addr 含標題列（第一列為欄名，作為圖例）\n"
                "  2. 資料為數值型（非文字）\n"
                "chart_type 選擇指南：\n"
                "  column / bar = 比較各類別的數量（直條 vs 橫條）\n"
                "  line = 顯示趨勢、時間序列\n"
                "  pie = 顯示佔比（建議資料不超過 7 個類別）\n"
                "  area = 顯示累積趨勢\n"
                "  scatter = 顯示兩變數的相關性"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "資料範圍（含標題列），如 'A1:B20'"},
                    "chart_type": {
                        "type": "string",
                        "enum": [
                            "column", "column_stacked",
                            "bar", "bar_stacked",
                            "line", "line_markers",
                            "pie", "doughnut",
                            "area", "area_stacked",
                            "scatter",
                            "histogram", "waterfall", "box_whisker", "funnel",
                        ],
                        "description": (
                            "圖表類型；預設 column。\n"
                            "選擇指南：\n"
                            "  column/bar = 類別比較（直條/橫條）\n"
                            "  column_stacked/bar_stacked = 堆疊比例\n"
                            "  line/line_markers = 趨勢時間序列\n"
                            "  pie/doughnut = 佔比（建議 ≤7 類）\n"
                            "  area/area_stacked = 累積趨勢\n"
                            "  scatter = 兩變數相關性\n"
                            "  histogram = 分佈（需 Excel 2016+）\n"
                            "  waterfall = 瀑布圖財務分析（需 Excel 2016+）\n"
                            "  box_whisker = 箱型圖統計分析（需 Excel 2016+）\n"
                            "  funnel = 漏斗圖（需 Excel 2016+）"
                        ),
                    },
                    "title":  {"type": "string", "description": "圖表標題（選填）"},
                    "sheet":  {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 樞紐分析表
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "create_pivot_table",
            "description": (
                "建立樞紐分析表。值欄位自動以加總（SUM）計算。\n"
                "⚠️ 最重要注意事項：row_field / col_field / value_field 必須與來源資料標題列欄名「完全一致」（大小寫、空白都算）。\n"
                "   執行前必須先用 read_range 讀取標題列確認欄名，不可靠記憶或猜測。\n"
                "來源資料必須含標題列（第一列為欄名）。\n"
                "dest_sheet 不存在時自動建立；若已存在則先清除內容再重建。\n"
                "col_field 選填，加入後可做交叉分析（列×欄的矩陣）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_range":  {"type": "string", "description": "來源資料範圍（含標題列），如 'A1:D100'"},
                    "dest_sheet":    {"type": "string", "description": "放置樞紐分析表的工作表名稱"},
                    "row_field":     {"type": "string", "description": "列標籤欄位名稱（必須與標題列欄名完全一致）"},
                    "value_field":   {"type": "string", "description": "值欄位名稱（數值欄，將加總；必須與標題列欄名完全一致）"},
                    "col_field":     {"type": "string", "description": "欄標籤欄位名稱（選填；必須與標題列欄名完全一致）"},
                    "source_sheet":  {"type": "string", "description": "來源工作表名稱；省略=作用中工作表"},
                },
                "required": ["source_range", "dest_sheet", "row_field", "value_field"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "refresh_pivot_table",
            "description": (
                "重新整理樞紐分析表，讓它反映最新的來源資料。\n"
                "每次修改來源資料（write_range / find_replace 等）後，"
                "都應呼叫此工具更新樞紐，否則數字會是舊的。\n"
                "若工作表上有多個樞紐，會全部一起重新整理。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "pivot_sheet": {"type": "string", "description": "樞紐分析表所在的工作表名稱"},
                },
                "required": ["pivot_sheet"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "format_pivot_table",
            "description": (
                "套用樞紐分析表的內建樣式，讓樞紐外觀更專業。\n"
                "style 常用值（PivotStyleMedium 系列最推薦）：\n"
                "  PivotStyleMedium9  = 藍色（預設推薦）\n"
                "  PivotStyleMedium4  = 橘色\n"
                "  PivotStyleMedium7  = 綠色\n"
                "  PivotStyleDark1    = 深藍（強調）\n"
                "  PivotStyleLight16  = 淡藍（低調）\n"
                "banded_rows=true 交錯帶狀列色，資料多時更易閱讀。\n"
                "若工作表有多個樞紐，只套用到第一個。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "pivot_sheet":     {"type": "string",  "description": "樞紐所在工作表名稱"},
                    "style":           {"type": "string",  "description": "樞紐樣式名稱；預設 PivotStyleMedium9"},
                    "show_row_headers":{"type": "boolean", "description": "顯示列標題樣式；預設 true"},
                    "show_col_headers":{"type": "boolean", "description": "顯示欄標題樣式；預設 true"},
                    "banded_rows":     {"type": "boolean", "description": "帶狀列底色（交錯）；預設 true"},
                    "banded_cols":     {"type": "boolean", "description": "帶狀欄底色；預設 false"},
                },
                "required": ["pivot_sheet"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # 視窗 / 欄列外觀
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "freeze_panes",
            "description": (
                "凍結工作表的列或欄，捲動時標題保持可見。\n"
                "最常見用法：freeze_panes row=1, col=0 → 凍結第一列（標題行）。\n"
                "freeze_panes row=1, col=1 → 同時凍結第一列和第一欄。\n"
                "freeze_panes row=0, col=0 → 解除所有凍結。\n"
                "row/col 都可省略（省略視為 0）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "row":   {"type": "integer", "description": "凍結幾列（0=不凍結）；最常用 1（凍結標題行）"},
                    "col":   {"type": "integer", "description": "凍結幾欄（0=不凍結）"},
                    "sheet": {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "auto_fit",
            "description": (
                "自動調整欄寬或列高，讓儲存格內容完整顯示不被截斷。\n"
                "target='columns'（最常用）：自動調整欄寬。\n"
                "target='rows'：自動調整列高（適合有換行內容的格子）。\n"
                "target='both'：同時調整欄寬和列高。\n"
                "range_addr 省略時調整整張工作表所有欄/列。\n"
                "格式化完成後建議都呼叫一次 auto_fit，確保外觀整齊。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "target": {
                        "type": "string",
                        "enum": ["columns", "rows", "both"],
                        "description": "調整目標；預設 columns",
                    },
                    "range_addr": {"type": "string", "description": "指定範圍（省略=整張工作表）"},
                    "sheet":      {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_column_width",
            "description": (
                "精確設定指定欄的寬度（Excel 字元寬度單位）。\n"
                "預設欄寬約 8.43。常用值：5~8（窄欄，如序號）、12~18（一般文字）、25~35（長文字或說明）。\n"
                "適合需要統一欄寬（如報表對齊）的場景；一般情況優先用 auto_fit。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "column_index": {"type": "integer", "description": "欄號（從 1 開始，1=A、2=B…）"},
                    "width":        {"type": "number",  "description": "寬度值（字元寬度單位，建議範圍 5~50）"},
                    "sheet":        {"type": "string"},
                },
                "required": ["column_index", "width"],
            },
        },
    },


    # ══════════════════════════════════════════════════════════════════════════
    # V4 美化工具群
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "apply_table_style",
            "description": (
                "將儲存格範圍轉換為 Excel 正式表格（ListObject），一次套用帶狀底色、標題粗體、自動篩選按鈕。\n"
                "這是讓資料表立刻變漂亮最有效的工具，建議每次格式化後優先使用。\n"
                "style 常用推薦：\n"
                "  blue（藍，最專業）、green（綠）、orange（橘）、gray（灰，低調）\n"
                "  dark_blue（深藍，強調）、light_blue（淡藍，輕盈）\n"
                "show_totals=true 時，表格底部自動加入合計列。\n"
                "注意：範圍不可與現有表格重疊；操作後欄寬可能需要重新 auto_fit。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":  {"type": "string", "description": "要轉換為表格的範圍（含標題列），如 'A1:F50'"},
                    "style": {
                        "type": "string",
                        "enum": [
                            "blue", "light_blue", "green", "light_green",
                            "orange", "light_orange", "red", "purple",
                            "gray", "white", "dark_blue", "dark_green", "dark_red",
                        ],
                        "description": "表格樣式；預設 blue",
                    },
                    "table_name":  {"type": "string",  "description": "表格名稱（選填，供公式引用；不填則自動命名）"},
                    "has_header":  {"type": "boolean", "description": "第一列是否為標題列；預設 true"},
                    "show_totals": {"type": "boolean", "description": "是否顯示合計列；預設 false"},
                    "sheet":       {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "format_chart",
            "description": (
                "格式化工作表上已存在的圖表（需先用 create_chart / create_combo_chart 建立）。\n"
                "可設定：標題、X/Y 軸標題、圖例、資料標籤、各數列顏色、繪圖區背景色。\n"
                "chart_index：工作表上第幾個圖表（從 1 開始）。\n"
                "series_colors：依序指定每個數列的顏色，如 ['#4472C4', '#ED7D31', '#A9D18E']。\n"
                "使用流程：create_chart → format_chart → auto_fit（調整欄寬）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "chart_index":    {"type": "integer", "description": "圖表索引（第幾個圖，從 1 開始）；預設 1"},
                    "title":          {"type": "string",  "description": "圖表標題文字"},
                    "x_axis_title":   {"type": "string",  "description": "X 軸（分類軸）標題"},
                    "y_axis_title":   {"type": "string",  "description": "Y 軸（數值軸）標題"},
                    "has_legend":     {"type": "boolean", "description": "是否顯示圖例"},
                    "legend_position":{
                        "type": "string",
                        "enum": ["bottom", "right", "top", "left"],
                        "description": "圖例位置；預設 bottom",
                    },
                    "has_data_labels":{"type": "boolean", "description": "是否在每個資料點顯示數值標籤"},
                    "series_colors":  {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "各數列顏色 #RRGGBB 清單，如 ['#4472C4','#ED7D31']",
                    },
                    "plot_bg_color":  {"type": "string", "description": "繪圖區背景色 #RRGGBB（如 '#F2F2F2'）"},
                    "sheet":          {"type": "string"},
                },
                "required": ["chart_index"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "create_combo_chart",
            "description": (
                "建立組合圖（直條圖 + 折線圖）。最後一個數列預設改為折線並使用次要 Y 軸。\n"
                "最適合場景：主數列用柱狀顯示絕對值（如銷售額），次數列用折線顯示相對指標（如成長率 %）。\n"
                "line_series_index：哪個數列改為折線（1-based；-1 = 最後一個，通常是比率或成長率欄）。\n"
                "secondary_axis=true：折線使用右側次軸（解決主次軸尺度差異的必要設定）。\n"
                "建立後可呼叫 format_chart 進一步美化顏色與標題。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":         {"type": "string",  "description": "資料範圍（含標題列），如 'A1:C20'"},
                    "line_series_index":  {"type": "integer", "description": "改為折線的數列索引（1-based；-1=最後一個）；預設 -1"},
                    "secondary_axis":     {"type": "boolean", "description": "折線是否使用次要 Y 軸（右軸）；預設 true"},
                    "title":              {"type": "string",  "description": "圖表標題（選填）"},
                    "sheet":              {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_sparklines",
            "description": (
                "在儲存格內插入 Sparkline 迷你圖（折線 / 直條 / 盈虧）。\n"
                "迷你圖嵌在格子裡，不佔額外空間，適合摘要列或 KPI Dashboard。\n"
                "常見用法：data_range='B2:M2'（12 個月資料）、sparkline_range='N2'（放在 N 欄）。\n"
                "多列時：data_range='B2:M20'（19 列資料）、sparkline_range='N2:N20'（N 欄逐列配對）。\n"
                "sparkline_type：line（趨勢折線，最常用）/ column（逐月直條）/ winloss（正負盈虧）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "data_range":      {"type": "string", "description": "資料來源範圍（列數須與 sparkline_range 相符）"},
                    "sparkline_range": {"type": "string", "description": "迷你圖放置位置，如 'N2:N20'"},
                    "sparkline_type": {
                        "type": "string",
                        "enum": ["line", "column", "winloss"],
                        "description": "迷你圖類型；預設 line",
                    },
                    "color": {"type": "string", "description": "迷你圖顏色 #RRGGBB（選填）"},
                    "sheet": {"type": "string"},
                },
                "required": ["data_range", "sparkline_range"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "set_tab_color",
            "description": (
                "設定工作表底部標籤的顏色，讓多個工作表一目了然。\n"
                "常見配色策略：原始資料=灰色、報表=藍色、圖表=綠色、注意事項=紅色。\n"
                "顏色用 #RRGGBB，如 '#4472C4'（藍）、'#70AD47'（綠）、'#FF0000'（紅）、'#808080'（灰）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "color": {"type": "string", "description": "標籤顏色 #RRGGBB"},
                    "sheet": {"type": "string", "description": "工作表名稱；省略則用作用中工作表"},
                },
                "required": ["color"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "page_setup",
            "description": (
                "設定工作表列印版面：方向、紙張、列印範圍、縮放。讓報表可直接列印送出。\n"
                "orientation：portrait（直印，A4 直放）/ landscape（橫印，寬表格用）。\n"
                "fit_to_wide=1, fit_to_tall=0 → 縮放成一頁寬（最常用，高度不限）。\n"
                "print_area：指定列印範圍，如 'A1:H50'（省略 = 整張表）。\n"
                "center_horizontally=true → 列印內容水平置中。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "orientation": {
                        "type": "string",
                        "enum": ["portrait", "landscape"],
                        "description": "列印方向；預設 portrait",
                    },
                    "paper_size": {
                        "type": "string",
                        "enum": ["a4", "letter", "a3", "a5", "legal"],
                        "description": "紙張大小；預設 a4",
                    },
                    "fit_to_wide": {
                        "type": "integer",
                        "description": "縮放成幾頁寬（如 1=一頁寬）；設定後會關閉 Zoom 百分比",
                    },
                    "fit_to_tall": {
                        "type": "integer",
                        "description": "縮放成幾頁高（如 1=一頁高；0=不限頁數）",
                    },
                    "print_area": {
                        "type": "string",
                        "description": "列印範圍，如 'A1:H50'；省略 = 整張工作表",
                    },
                    "center_horizontally": {"type": "boolean", "description": "是否水平置中列印"},
                    "center_vertically":   {"type": "boolean", "description": "是否垂直置中列印"},
                    "sheet": {"type": "string"},
                },
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_slicer",
            "description": (
                "為樞紐分析表新增切片器（Slicer），讓使用者可點按篩選資料。需要 Excel 2010+。\n"
                "前提：工作表上必須已有 create_pivot_table 建立的樞紐分析表。\n"
                "field_name：要篩選的欄位名稱，必須與樞紐分析表欄名完全一致。\n"
                "切片器預設放在樞紐所在的工作表右側（left=500）；可調整 left/top 避開圖表。\n"
                "常見用法：建完樞紐後呼叫此工具加 '地區' 或 '產品類別' 切片器。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "pivot_sheet": {"type": "string",  "description": "樞紐分析表所在工作表名稱"},
                    "field_name":  {"type": "string",  "description": "切片器對應的欄位名稱（須與樞紐欄名完全一致）"},
                    "dest_sheet":  {"type": "string",  "description": "切片器放置的工作表（省略 = 與樞紐同一張）"},
                    "left":        {"type": "number",  "description": "切片器左邊距（點數）；預設 500"},
                    "top":         {"type": "number",  "description": "切片器上邊距（點數）；預設 50"},
                    "width":       {"type": "number",  "description": "切片器寬度（點數）；預設 150"},
                    "height":      {"type": "number",  "description": "切片器高度（點數）；預設 200"},
                },
                "required": ["pivot_sheet", "field_name"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # V4 分析工具群
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "summarize_range",
            "description": (
                "對範圍內的數值儲存格計算統計摘要，直接回傳結果給 AI，不需先 read_range 再手動計算。\n"
                "預設計算全部統計；stats 可指定只算需要的項目，加快速度。\n"
                "stats 可選值：sum / average / max / min / count / stdev / median / count_all / count_blank。\n"
                "常用場景：\n"
                "  - 回答「這欄平均是多少」→ stats=['average']\n"
                "  - 確認資料範圍 → stats=['max','min','count']\n"
                "  - 品質檢查 → stats=['count_all','count_blank','count']（確認空白格比例）"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "目標範圍，如 'B2:B100' 或 'B2:F50'"},
                    "stats": {
                        "type": "array",
                        "items": {
                            "type": "string",
                            "enum": ["sum","average","max","min","count","stdev","median","count_all","count_blank"],
                        },
                        "description": "要計算的統計項目；省略則全部計算",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "find_duplicates",
            "description": (
                "找出範圍內的重複值，可標記、刪除或只列出。\n"
                "column_index：依範圍內第幾欄判斷重複（1-based）。\n"
                "action 三種模式：\n"
                "  mark   = 把重複列整列標記指定背景色（不刪除，可事後人工確認）\n"
                "  delete = 刪除重複列，只保留第一次出現的列（⚠️ 不可復原）\n"
                "  list   = 只回傳重複值清單，完全不修改工作表（最安全）\n"
                "建議先用 list 確認後，再決定是否 mark 或 delete。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":    {"type": "string",  "description": "資料範圍（含標題列），如 'A1:E100'"},
                    "column_index":  {"type": "integer", "description": "依第幾欄判斷重複（範圍內 1-based）；預設 1"},
                    "action": {
                        "type": "string",
                        "enum": ["mark", "delete", "list"],
                        "description": "處理方式；預設 mark",
                    },
                    "mark_color":    {"type": "string",  "description": "標記顏色 #RRGGBB（action=mark 時）；預設黃色 #FFFF00"},
                    "sheet":         {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "fill_series",
            "description": (
                "自動填充序列（數字 / 日期 / 平日 / 月份 / 年份）。\n"
                "start_cell 的現有值作為起始值（或用 start_value 覆蓋）。\n"
                "series_type 說明：\n"
                "  number  = 數字線性遞增（1,2,3… 或 step=5 → 5,10,15…）\n"
                "  date    = 依天遞增（2024/1/1, 2024/1/2…）\n"
                "  weekday = 跳過週末（只填平日）\n"
                "  month   = 依月遞增（2024/1/1, 2024/2/1…）\n"
                "  year    = 依年遞增\n"
                "direction：down=向下填充（最常用）/ right=向右填充。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "start_cell":  {"type": "string",  "description": "起始儲存格，如 'A1'"},
                    "count":       {"type": "integer", "description": "填充幾個（含起始格）"},
                    "series_type": {
                        "type": "string",
                        "enum": ["number", "date", "weekday", "month", "year"],
                        "description": "序列類型；預設 number",
                    },
                    "step":        {"type": "number",  "description": "每次遞增量；預設 1"},
                    "start_value": {
                        "type": "string",
                        "description": "覆蓋起始格的值（如 '1' 或 '2024/01/01'）；省略則使用儲存格現有值",
                    },
                    "direction": {
                        "type": "string",
                        "enum": ["down", "right"],
                        "description": "填充方向；預設 down（向下）",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["start_cell", "count"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "group_rows",
            "description": (
                "對指定列範圍建立大綱分組，讓使用者可折疊/展開明細列。\n"
                "多層級報表必備：把明細列分組後折疊，只顯示小計列，外觀更整潔。\n"
                "action：group（建立分組）/ ungroup（移除分組）。\n"
                "常見用法：先用 add_subtotal 加小計，再 group_rows 把明細分組折疊。\n"
                "注意：Excel 大綱最多 8 層，start_row 不可等於 end_row（至少 2 列）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "start_row": {"type": "integer", "description": "分組起始列號（1-based）"},
                    "end_row":   {"type": "integer", "description": "分組結束列號（須 ≥ start_row）"},
                    "action": {
                        "type": "string",
                        "enum": ["group", "ungroup"],
                        "description": "操作類型；預設 group",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["start_row", "end_row"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "group_columns",
            "description": (
                "對指定欄範圍建立大綱分組，讓使用者可折疊/展開輔助欄位。\n"
                "常見用法：把計算過程欄（如中間計算值）分組折疊，報表更簡潔。\n"
                "start_col / end_col：欄號（1-based，1=A、2=B…）。\n"
                "action：group（建立分組）/ ungroup（移除分組）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "start_col": {"type": "integer", "description": "分組起始欄號（1-based，1=A）"},
                    "end_col":   {"type": "integer", "description": "分組結束欄號（須 ≥ start_col）"},
                    "action": {
                        "type": "string",
                        "enum": ["group", "ungroup"],
                        "description": "操作類型；預設 group",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["start_col", "end_col"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "transpose_range",
            "description": (
                "將範圍行列轉置（列↔欄互換）後寫入到新位置。\n"
                "例：原本 1列×5欄 → 轉置後變 5列×1欄。\n"
                "只複製數值，不複製格式（格式請之後用 format_range 套用）。\n"
                "dest_cell：目標左上角起始格，與來源範圍不重疊即可。\n"
                "dest_sheet 不存在時自動建立。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_range": {"type": "string", "description": "來源範圍，如 'A1:E1' 或 'A1:C10'"},
                    "dest_cell":    {"type": "string", "description": "目標起始格，如 'A5' 或 'G1'"},
                    "source_sheet": {"type": "string", "description": "來源工作表；省略=目前工作表"},
                    "dest_sheet":   {"type": "string", "description": "目標工作表；省略=目前工作表；不存在自動建立"},
                },
                "required": ["source_range", "dest_cell"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "name_range",
            "description": (
                "為儲存格範圍建立具名範圍（Named Range）。\n"
                "命名後可在公式中直接使用名稱，如 =SUM(銷售額) 取代 =SUM(B2:B100)。\n"
                "名稱規則：只能包含字母、數字、底線，不可以數字開頭，不可與現有名稱重複。\n"
                "常見用途：\n"
                "  - 固定範圍命名後，公式更易讀、更易維護\n"
                "  - 資料驗證下拉選單可直接引用具名範圍\n"
                "  - 跨工作表公式使用名稱取代複雜的 Sheet!A1:A100 引用"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要命名的範圍，如 'B2:B100'"},
                    "name":       {"type": "string", "description": "具名範圍名稱（字母/數字/底線，如 '銷售額' 或 'SalesData'）"},
                    "sheet":      {"type": "string"},
                },
                "required": ["range_addr", "name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_subtotal",
            "description": (
                "依分組欄位自動插入小計列（等同 Excel「資料→小計」功能）。\n"
                "⚠️ 執行前資料必須已依 group_by_column 排序（先呼叫 sort_range），否則小計結果錯誤。\n"
                "group_by_column：當這欄的值改變時，自動插入一列小計（範圍內 1-based）。\n"
                "value_columns：要加總/計算的欄位索引清單（範圍內 1-based），如 [3,4,5]。\n"
                "function_type：sum（加總）/ count（計數）/ average（平均）/ max（最大）/ min（最小）。\n"
                "建立後可搭配 group_rows 把明細列折疊，只顯示小計。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":       {"type": "string",  "description": "資料範圍（含標題列），如 'A1:E100'"},
                    "group_by_column":  {"type": "integer", "description": "分組依據欄（範圍內 1-based）"},
                    "value_columns": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "要計算小計的欄索引清單（範圍內 1-based），如 [3,4]",
                    },
                    "function_type": {
                        "type": "string",
                        "enum": ["sum", "count", "average", "max", "min"],
                        "description": "統計函數；預設 sum",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr", "group_by_column", "value_columns"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "advanced_filter",
            "description": (
                "進階篩選：多條件組合篩選，可就地篩選或將結果複製到新位置。\n"
                "criteria_range：條件範圍，必須：\n"
                "  - 第一列為欄標題（須與資料標題完全一致）\n"
                "  - 第二列起為條件值；同一列的條件為 AND，不同列的條件為 OR\n"
                "  - 例：A列='台北', B列='>10000' → 篩選台北且金額>10000\n"
                "dest_range：填寫則複製篩選結果到此位置；省略則就地隱藏不符合的列。\n"
                "unique_only=true：只保留不重複記錄（不需 criteria_range）。\n"
                "注意：就地篩選時需呼叫 filter_range 清除篩選以顯示全部資料。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":      {"type": "string",  "description": "資料範圍（含標題列），如 'A1:F100'"},
                    "criteria_range":  {"type": "string",  "description": "條件範圍（含標題列），如 'H1:J3'；省略=只做唯一記錄篩選"},
                    "dest_range":      {"type": "string",  "description": "結果複製目標起始格，如 'A105'；省略=就地篩選"},
                    "unique_only":     {"type": "boolean", "description": "只保留不重複記錄；預設 false"},
                    "sheet":           {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "split_text_to_columns",
            "description": (
                "依分隔符將文字欄位拆分成多欄（等同 Excel「資料→資料剖析」）。\n"
                "⚠️ 會直接覆蓋右側相鄰欄位，執行前請確認右側有足夠空白欄。\n"
                "delimiter 支援：\n"
                "  comma（,）/ tab（定位字元）/ semicolon（;）/ space（空格）\n"
                "  或直接填任意單一字元，如 '|'、'-'、':'。\n"
                "常見用途：\n"
                "  - 從 CSV 貼入的資料（逗號分隔）拆欄\n"
                "  - '姓名-部門-職稱' 拆成三欄\n"
                "  - 日期 '2024/01/15' 拆成年/月/日三欄"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr": {"type": "string", "description": "要拆分的欄位範圍，如 'A2:A100'（通常只選一欄）"},
                    "delimiter":  {
                        "type": "string",
                        "description": "分隔符：comma / tab / semicolon / space 或任意單一字元（如 '|'）；預設 comma",
                    },
                    "sheet": {"type": "string"},
                },
                "required": ["range_addr"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "add_image",
            "description": (
                "插入圖片到工作表，圖片左上角對齊指定儲存格。\n"
                "支援格式：jpg / png / bmp / gif。\n"
                "image_path：圖片路徑（絕對路徑或相對於執行目錄的相對路徑）。\n"
                "width / height：省略時使用圖片原始尺寸；只填一邊可按比例縮放（另一邊自動調整）。\n"
                "常見用途：在報表左上角插入公司 LOGO、在標題旁放置產品圖片。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "image_path": {"type": "string", "description": "圖片檔案路徑，如 'C:/logo.png' 或 'images/logo.png'"},
                    "range_addr": {"type": "string", "description": "圖片左上角對齊的儲存格，如 'A1' 或 'B3'"},
                    "width":      {"type": "number", "description": "圖片寬度（點數）；省略 = 原始尺寸"},
                    "height":     {"type": "number", "description": "圖片高度（點數）；省略 = 原始尺寸"},
                    "sheet":      {"type": "string"},
                },
                "required": ["image_path", "range_addr"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # Phase 2：Undo
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "undo_last",
            "description": (
                "還原上一個可復原操作（Phase 2）。\n"
                "支援三類策略：\n"
                "  Category A（反向操作）：insert_row / insert_column / add_sheet / "
                "rename_sheet / merge_cells / unmerge_cells\n"
                "  Category B（回寫備份資料）：write_range / clear_range / trim_range\n"
                "  Category C（無法還原）：其餘操作 — 回傳說明訊息，不修改工作表\n"
                "備份堆疊上限 20 步；若堆疊為空回傳 no_op。\n"
                "注意：undo_last 本身不可再被 undo（Meta 操作，不放入備份堆疊）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },


    # ══════════════════════════════════════════════════════════════════════════
    # V4.7.0 A：巨集錄製與重播
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "record_macro",
            "description": (
                "錄製巨集：將最近的操作歷史（BackupStack）存成可重複執行的巨集。\n"
                "使用時機：使用者說「把剛才的步驟存成巨集」、「錄製這些操作」時。\n"
                "若提供 steps 參數，可指定任意步驟清單，不限於最近操作。\n"
                "相同名稱的巨集會覆蓋舊版本。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name":        {"type": "string", "description": "巨集名稱（唯一鍵）"},
                    "description": {"type": "string", "description": "巨集說明文字（可省略）"},
                    "steps": {
                        "type": "array",
                        "description": "步驟清單；省略時從操作歷史自動取得",
                        "items": {"type": "object"},
                    },
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "list_macros",
            "description": "列出所有已儲存的巨集（名稱、說明、步驟數）。",
            "parameters": {"type": "object", "properties": {}},
        },
    },

    {
        "type": "function",
        "function": {
            "name": "run_macro",
            "description": (
                "執行已儲存的巨集。失敗時自動回滾已執行步驟。\n"
                "不確定名稱時先呼叫 list_macros 確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要執行的巨集名稱"},
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "delete_macro",
            "description": "刪除指定巨集。",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要刪除的巨集名稱"},
                },
                "required": ["name"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # V4.7.0 B：公式智慧輔助
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "validate_formula",
            "description": (
                "驗證 Excel 公式的語法正確性（括號配對、函數名稱是否已知）。\n"
                "使用時機：準備寫入公式前確認語法無誤時。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "description": "要驗證的 Excel 公式（含開頭的 =）"},
                },
                "required": ["formula"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "explain_formula",
            "description": "生成 Excel 公式的繁體中文說明（函數語意 + 引用範圍）。",
            "parameters": {
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "description": "要說明的 Excel 公式（含開頭的 =）"},
                },
                "required": ["formula"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # V4.7.0 C：自然語言資料查詢
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "query_range",
            "description": (
                "在記憶體內查詢 Excel 範圍資料（完全非破壞性，不修改 Excel）。\n"
                "支援過濾（>, <, >=, <=, =, !=, contains, startswith, endswith, isblank, notblank）、\n"
                "排序（sort_by）、取前 N 筆（top_n）、聚合（sum/avg/count/max/min）。\n"
                "condition_json 格式：\n"
                "  {\"filters\": [{\"column\": 2, \"operator\": \">\", \"value\": 10000}],\n"
                "   \"sort_by\": {\"column\": 2, \"descending\": true}, \"top_n\": 5}\n"
                "aggregation_json 格式：{\"function\": \"sum\", \"column\": 2}"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "range_addr":       {"type": "string",  "description": "要查詢的範圍（如 'A1:F100'）"},
                    "condition_json":   {"type": "string",  "description": "過濾/排序/top_n 條件，JSON 字串"},
                    "aggregation_json": {"type": "string",  "description": "聚合設定，JSON 字串"},
                    "has_header":       {"type": "boolean", "description": "第一列是否為標題（預設 true）"},
                    "sheet":            {"type": "string",  "description": "工作表名稱；省略時用作用中工作表"},
                },
                "required": ["range_addr"],
            },
        },
    },

    # ══════════════════════════════════════════════════════════════════════════
    # V4.7.0 D：多工作簿協作
    # ══════════════════════════════════════════════════════════════════════════

    {
        "type": "function",
        "function": {
            "name": "list_workbooks",
            "description": (
                "列出目前 Excel 中所有已開啟的活頁簿（名稱、路徑、是否為作用中）。\n"
                "跨活頁簿操作前先呼叫此工具確認名稱。"
            ),
            "parameters": {"type": "object", "properties": {}},
        },
    },

    {
        "type": "function",
        "function": {
            "name": "switch_workbook",
            "description": (
                "切換作用中活頁簿。\n"
                "不確定名稱時先呼叫 list_workbooks 確認。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要切換的活頁簿名稱（如 'Report.xlsx'）"},
                },
                "required": ["name"],
            },
        },
    },

    {
        "type": "function",
        "function": {
            "name": "copy_range_between_workbooks",
            "description": (
                "跨活頁簿複製範圍資料（值複製，來源不受影響）。\n"
                "操作前請先用 list_workbooks 確認來源和目標活頁簿名稱。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_range": {"type": "string",  "description": "來源範圍（如 'A1:D10'）"},
                    "dest_range":   {"type": "string",  "description": "目標起始儲存格（如 'A1'）"},
                    "source_wb":    {"type": "string",  "description": "來源活頁簿名稱；省略時用目前作用中活頁簿"},
                    "dest_wb":      {"type": "string",  "description": "目標活頁簿名稱；省略時用目前作用中活頁簿"},
                    "source_sheet": {"type": "string",  "description": "來源工作表名稱；省略時用作用中工作表"},
                    "dest_sheet":   {"type": "string",  "description": "目標工作表名稱；省略時用作用中工作表"},
                    "values_only":  {"type": "boolean", "description": "僅複製值（預設 true）"},
                },
                "required": ["source_range", "dest_range"],
            },
        },
    },

]
