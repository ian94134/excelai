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
                "所有參數皆為選填，只需傳入要修改的屬性。\n"
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
                        "enum": ["column", "bar", "line", "pie", "area", "scatter"],
                        "description": "圖表類型；預設 column",
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

]
