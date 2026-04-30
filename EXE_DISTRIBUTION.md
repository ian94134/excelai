# Excel AI Assistant EXE 發佈說明

這個專案可以用 PyInstaller 打成 Windows 發佈包，讓同事不需要安裝 Python。

## 打包

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows_exe.ps1
```

輸出位置：

- `dist\Excel AI Assistant\Excel AI Assistant.exe`
- `dist\Excel-AI-Assistant-windows.zip`

## 交付同事

請交付整個 zip 或整個 `dist\Excel AI Assistant` 資料夾，不要只傳單一 exe，因為 `_internal` 目錄包含內嵌 Python runtime 與套件。

同事端需求：

- Windows
- Microsoft Excel
- 可連到 Qwen / OpenAI-compatible server

同事操作：

1. 解壓 `Excel-AI-Assistant-windows.zip`
2. 執行 `Excel AI Assistant.exe`
3. 瀏覽器會開啟 `http://localhost:8501/`
4. 在側邊欄確認 Qwen 伺服器設定

注意：發佈包不包含 `.env`，也不包含測試用 workbook。
