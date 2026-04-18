# 📊 Excel AI 助手

透過自然語言直接操控 Microsoft Excel，無需手動操作，說出你想做的事，AI 自動完成。

基於本地 Qwen 大型語言模型，28 個 Excel 工具，Streamlit 網頁介面，可打包為 Windows EXE 供同事直接使用。

---

## ✨ 功能一覽

| 類別 | 工具 |
|------|------|
| 讀寫 | 讀取範圍、寫入值／公式、儲存檔案 |
| 格式 | 字型、顏色、背景色、數字格式、對齊、框線、合併儲存格 |
| 列欄 | 插入／刪除列、插入／刪除欄、設定列高、設定欄寬、自動欄寬 |
| 工作表 | 新增、重新命名、凍結窗格 |
| 資料 | 排序、篩選、尋找取代、跨表複製、清除 |
| 視覺化 | 建立圖表（直條、橫條、折線、圓餅、區域、散佈）|
| 分析 | 建立樞紐分析表 |
| 自動化 | 條件格式化、下拉選單驗證 |

**指令範例：**
- `把標題列加粗、藍底白字，然後自動調整欄寬`
- `篩選出地區為台北的資料`
- `用銷售額欄位建立長條圖`
- `做一張樞紐分析表，列用地區，值用金額加總`
- `把大於 80 分的格子變綠色`
- `在 B 欄設定下拉選單：是、否、待定`

---

## 🖥️ 系統需求

- Windows 10 / 11
- Microsoft Excel 2016 以上（需已開啟 Excel 檔案）
- Python 3.10 以上（開發環境用）
- 可存取 Qwen 本地模型伺服器

---

## 🚀 快速開始（開發環境）

### 1. 安裝依賴套件

```bash
pip install -r requirements.txt
```

### 2. 設定環境變數

複製 `.env.example` 並建立 `.env`：

```bash
copy .env.example .env
```

編輯 `.env`，填入你的 Qwen 伺服器資訊：

```env
QWEN_BASE_URL=http://你的伺服器IP:埠號/v1
QWEN_MODEL=你的模型名稱
```

### 3. 開啟 Excel，然後啟動

```bash
streamlit run main.py
```

瀏覽器會自動開啟 `http://localhost:8501`。

---

## 📦 打包為 EXE（供同事使用）

```bash
python build.py
```

打包完成後，`dist/ExcelAI/` 資料夾內會有 `ExcelAI.exe`，將整個資料夾複製給同事即可。

### 同事使用方式

1. 安裝 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)（若尚未安裝）
2. 先開啟 Excel，並打開要操作的檔案
3. 雙擊 `ExcelAI.exe`，瀏覽器會自動開啟操作介面
4. 輸入自然語言指令即可

> ⚠️ EXE 與 Excel 必須同時開啟才能正常運作

---

## 🗂️ 專案結構

```
excel-ai/
├── main.py              # Streamlit 主程式（入口）
├── config.py            # 環境變數、System Prompt、Provider 工廠
├── constants.py         # Excel COM 常數集中管理
├── excel_tools.py       # 28 個工具的 win32com 實作
├── session.py           # 對話訊息管理（摘要、儲存、載入）
├── launcher.py          # PyInstaller EXE 啟動器
├── build.py             # PyInstaller 打包腳本
│
├── providers/
│   ├── base.py          # LLMProvider 抽象介面
│   └── local_qwen.py    # Qwen 串流／非串流實作
│
├── tools/
│   ├── definition.py    # 28 個工具的 OpenAI 格式 schema
│   └── executor.py      # 工具分派與執行
│
└── ui/
    └── sidebar.py       # Streamlit 側邊欄元件
```

---

## ⚙️ 主要技術

| 技術 | 用途 |
|------|------|
| [Streamlit](https://streamlit.io) | 網頁介面 |
| [pywin32](https://github.com/mhammond/pywin32) | win32com 控制 Excel |
| [OpenAI Python SDK](https://github.com/openai/openai-python) | 對接 Qwen OpenAI 相容 API |
| [PyInstaller](https://pyinstaller.org) | 打包為 Windows EXE |
| Qwen 本地模型 | 自然語言理解與工具呼叫 |

---

## 🔒 安全說明

- `.env` 含伺服器位址，已加入 `.gitignore`，不會上傳至 GitHub
- 危險操作（刪除列、清除範圍、全文取代）執行前會彈出確認視窗
- 所有 Excel 操作只影響已開啟的活頁簿，不會讀取其他檔案

---

## 📋 版本紀錄

| 版本 | 主要功能 |
|------|---------|
| V1 | 14 個基礎工具（讀寫、格式、列欄、工作表操作） |
| V2 | +5 工具（圖表、樞紐、凍結、自動欄寬）、危險操作確認、對話儲存/載入 |
| V3 | +9 工具（篩選、合併、框線、條件格式、資料驗證）、串流回覆、Port 自動遞增、上下文摘要 |
| V3.1 | 模組化重構、constants.py、強化 AI 精準度（SOP + few-shot）|
