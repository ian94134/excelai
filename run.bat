@echo off
chcp 65001 >nul
echo 啟動 Excel AI 助手...
echo 請先確認 Excel 已開啟
echo.
streamlit run main.py --server.headless false --server.port 8501
pause
