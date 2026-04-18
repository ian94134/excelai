"""
打包腳本 - 執行方式：python build.py
取代 build.bat，避免 Windows 換行符號問題
"""
import subprocess
import sys
import os
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def run(cmd):
    print(f"\n>>> {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        print(f"\n[錯誤] 指令失敗，請查看上方訊息")
        sys.exit(1)

def main():
    print("=" * 50)
    print(" Excel AI 助手 - 打包成 EXE")
    print("=" * 50)

    # 清除舊的建置（先強制關閉殘留的 ExcelAI.exe）
    subprocess.run(["taskkill", "/f", "/im", "ExcelAI.exe"],
                   capture_output=True)  # 忽略錯誤（沒跑就沒關係）

    for d in ["dist", "build"]:
        p = os.path.join(BASE_DIR, d)
        if os.path.exists(p):
            try:
                shutil.rmtree(p)
                print(f"已清除 {d}/")
            except PermissionError:
                print(f"[警告] 無法清除 {d}/，請手動關閉 ExcelAI.exe 後重試")
                sys.exit(1)
    spec = os.path.join(BASE_DIR, "ExcelAI.spec")
    if os.path.exists(spec):
        os.remove(spec)

    # 組合 PyInstaller 指令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--console",       # 保留 console 視窗，錯誤訊息看得到
        "--name", "ExcelAI",
        "--collect-all", "streamlit",
        "--collect-all", "altair",
        "--collect-all", "vega_datasets",
        "--collect-all", "toolz",
        "--collect-all", "pyarrow",
        "--collect-all", "pandas",
        "--collect-all", "pydeck",
        "--collect-all", "validators",
        "--collect-all", "click",
        "--collect-all", "openai",
        "--collect-all", "dotenv",
        "--exclude-module", "google.genai",
        "--exclude-module", "google.auth",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "win32api",
        "--hidden-import", "win32con",
        "--hidden-import", "streamlit.runtime.scriptrunner.magic_funcs",
        "--add-data", "main.py;.",
        "--add-data", "config.py;.",
        "--add-data", "tools_definition.py;.",
        "--add-data", "tool_executor.py;.",
        "--add-data", "excel_tools.py;.",
        "--add-data", "providers;providers",
        "--add-data", ".env.example;.",
        "launcher.py",
    ]

    print("\n[開始打包，約需 3~8 分鐘...]\n")
    run(cmd)

    # 複製 pywin32 DLL
    print("\n[複製 pywin32 DLL...]")
    try:
        import sysconfig
        win32_dir = os.path.join(sysconfig.get_paths()["platlib"], "win32")
        dist_dir = os.path.join(BASE_DIR, "dist", "ExcelAI")
        copied = 0
        for fname in os.listdir(win32_dir):
            if fname.startswith(("pywintypes", "pythoncom")) and fname.endswith(".dll"):
                shutil.copy2(os.path.join(win32_dir, fname), dist_dir)
                print(f"  ✓ {fname}")
                copied += 1
        if copied == 0:
            print("  (找不到 DLL，若啟動失敗請手動複製)")
    except Exception as e:
        print(f"  (DLL 複製跳過：{e})")

    # 建立使用者啟動腳本
    dist_dir = os.path.join(BASE_DIR, "dist", "ExcelAI")
    bat_path = os.path.join(dist_dir, "啟動Excel助手.bat")
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write("@echo off\r\n")
        f.write("echo 請先確認 Excel 已開啟...\r\n")
        f.write("start ExcelAI.exe\r\n")

    print("\n" + "=" * 50)
    print(" 打包完成！")
    print(f" 輸出位置：{dist_dir}")
    print("\n 將整個 dist\\ExcelAI\\ 資料夾壓縮成 ZIP 給同仁")
    print(" 同仁執行 ExcelAI.exe 即可（不需要安裝 Python）")
    print("=" * 50)

    # 開啟輸出資料夾
    os.startfile(dist_dir)

if __name__ == "__main__":
    main()
