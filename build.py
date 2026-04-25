"""
Excel AI 助手 v4.0.0 — 打包腳本
執行方式：python build.py
"""
import subprocess
import sys
import os
import shutil

VERSION  = "v4.0.0"
APP_NAME = "ExcelAI"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ── 輔助函數 ──────────────────────────────────────────────────────────────────

def run(cmd: list[str]) -> None:
    print(f"\n>>> {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        print("\n[錯誤] 指令失敗，請查看上方訊息")
        sys.exit(1)


def check_prerequisites() -> None:
    """打包前確認必要工具與檔案存在。"""
    # PyInstaller
    r = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--version"],
        capture_output=True, text=True,
    )
    if r.returncode != 0:
        print("[錯誤] 未安裝 PyInstaller，請先執行：pip install pyinstaller")
        sys.exit(1)
    print(f"  ✓ PyInstaller {r.stdout.strip()}")

    # launcher.py
    launcher = os.path.join(BASE_DIR, "launcher.py")
    if not os.path.exists(launcher):
        print(f"[錯誤] 找不到 {launcher}")
        sys.exit(1)
    print(f"  ✓ launcher.py")

    # .env.example（.env 若不存在則自動從 example 複製）
    env_file     = os.path.join(BASE_DIR, ".env")
    env_example  = os.path.join(BASE_DIR, ".env.example")
    if not os.path.exists(env_file):
        if os.path.exists(env_example):
            shutil.copy2(env_example, env_file)
            print("  ✓ .env（由 .env.example 複製，請確認填入正確的 QWEN_BASE_URL）")
        else:
            print("  ⚠ 找不到 .env 和 .env.example，打包後使用者需手動建立 .env")
    else:
        print("  ✓ .env")


def clean_old_build() -> None:
    """清除上一次的建置輸出。"""
    # 先嘗試關閉殘留的 EXE
    subprocess.run(
        ["taskkill", "/f", "/im", f"{APP_NAME}.exe"],
        capture_output=True,
    )
    for d in ["dist", "build"]:
        p = os.path.join(BASE_DIR, d)
        if os.path.exists(p):
            try:
                shutil.rmtree(p)
                print(f"  ✓ 已清除 {d}/")
            except PermissionError:
                print(f"[警告] 無法清除 {d}/，請手動關閉 {APP_NAME}.exe 後重試")
                sys.exit(1)
    spec = os.path.join(BASE_DIR, f"{APP_NAME}.spec")
    if os.path.exists(spec):
        os.remove(spec)
        print(f"  ✓ 已清除 {APP_NAME}.spec")


def build_exe() -> None:
    """組合並執行 PyInstaller 指令。"""
    env_file    = os.path.join(BASE_DIR, ".env")
    has_env     = os.path.exists(env_file)
    env_arg     = f".env;." if has_env else None

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--console",        # 保留 console 視窗，方便查看錯誤
        "--name", APP_NAME,

        # ── Streamlit 相關套件（完整收集）────────────────────────────────────
        "--collect-all", "streamlit",
        "--collect-all", "altair",
        "--collect-all", "vega_datasets",
        "--collect-all", "toolz",
        "--collect-all", "pyarrow",
        "--collect-all", "pandas",
        "--collect-all", "pydeck",
        "--collect-all", "validators",
        "--collect-all", "click",

        # ── 網路與 LLM ────────────────────────────────────────────────────────
        "--collect-all", "openai",
        "--collect-all", "httpx",       # openai 內部依賴
        "--collect-all", "dotenv",

        # ── 排除不需要的大型套件（縮小體積）──────────────────────────────────
        "--exclude-module", "google.genai",
        "--exclude-module", "google.auth",
        "--exclude-module", "tensorflow",
        "--exclude-module", "torch",

        # ── pywin32 / COM（hidden import 確保不被排除）────────────────────────
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "win32api",
        "--hidden-import", "win32con",

        # ── Streamlit 內部 magic ──────────────────────────────────────────────
        "--hidden-import", "streamlit.runtime.scriptrunner.magic_funcs",

        # ── 專案原始碼（全部必要的 .py 檔）──────────────────────────────────
        "--add-data", "main.py;.",
        "--add-data", "launcher.py;.",
        "--add-data", "config.py;.",
        "--add-data", "constants.py;.",
        "--add-data", "session.py;.",
        "--add-data", "excel_tools.py;.",
        "--add-data", "backup.py;.",
        "--add-data", "exceptions.py;.",
        "--add-data", "logger.py;.",
        "--add-data", "excel_event_watcher.py;.",

        # ── 子目錄（整包帶入）────────────────────────────────────────────────
        "--add-data", "tools;tools",
        "--add-data", "ui;ui",
        "--add-data", "providers;providers",

        # ── 設定檔 ───────────────────────────────────────────────────────────
        "--add-data", ".env.example;.",
    ]

    # .env 若存在一起打包（含 Qwen 伺服器位址）
    if env_arg:
        cmd += ["--add-data", env_arg]

    cmd.append("launcher.py")

    print("\n[開始打包，約需 3~8 分鐘...]\n")
    run(cmd)


def copy_pywin32_dlls() -> None:
    """複製 pywin32 DLL 到輸出目錄（EXE 執行時需要）。"""
    dist_dir = os.path.join(BASE_DIR, "dist", APP_NAME)
    copied   = 0

    # 搜尋順序：site-packages/win32、build_venv/...
    import sysconfig
    search_dirs = [
        os.path.join(sysconfig.get_paths()["platlib"], "win32"),
        os.path.join(sysconfig.get_paths()["platlib"], "pywin32_system32"),
    ]
    # build_venv（若專案有虛擬環境）
    venv_dir = os.path.join(BASE_DIR, "build_venv")
    if os.path.exists(venv_dir):
        for root, dirs, files in os.walk(venv_dir):
            if "win32" in dirs:
                search_dirs.append(os.path.join(root, "win32"))
            if "pywin32_system32" in dirs:
                search_dirs.append(os.path.join(root, "pywin32_system32"))

    for search_dir in search_dirs:
        if not os.path.isdir(search_dir):
            continue
        for fname in os.listdir(search_dir):
            if fname.startswith(("pywintypes", "pythoncom")) and fname.endswith(".dll"):
                src = os.path.join(search_dir, fname)
                dst = os.path.join(dist_dir, fname)
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
                    print(f"  ✓ {fname}  ←  {search_dir}")
                    copied += 1

    if copied == 0:
        print("  ⚠ 找不到 pywintypes / pythoncom DLL")
        print("    若啟動後出現 ImportError，請手動複製這兩個 DLL 到 dist\\ExcelAI\\")


def create_launcher_bat() -> None:
    """建立給使用者點擊的啟動 BAT。"""
    dist_dir = os.path.join(BASE_DIR, "dist", APP_NAME)
    bat_path = os.path.join(dist_dir, "啟動Excel助手.bat")
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write("@echo off\r\n")
        f.write("chcp 65001 > nul\r\n")
        f.write("echo ==============================\r\n")
        f.write(f"echo  Excel AI 助手 {VERSION}\r\n")
        f.write("echo ==============================\r\n")
        f.write("echo 請先確認 Microsoft Excel 已開啟，再繼續...\r\n")
        f.write("pause\r\n")
        f.write(f"start {APP_NAME}.exe\r\n")
    print(f"  ✓ 啟動Excel助手.bat")


def create_readme() -> None:
    """在 dist 目錄建立給使用者的說明文字。"""
    dist_dir  = os.path.join(BASE_DIR, "dist", APP_NAME)
    readme    = os.path.join(dist_dir, "使用說明.txt")
    env_file  = os.path.join(dist_dir, ".env")
    with open(readme, "w", encoding="utf-8") as f:
        f.write(f"Excel AI 助手 {VERSION}\n")
        f.write("=" * 40 + "\n\n")
        f.write("【啟動方式】\n")
        f.write("  雙擊「啟動Excel助手.bat」\n")
        f.write("  （請先開啟 Microsoft Excel 再啟動）\n\n")
        f.write("【首次使用設定】\n")
        f.write("  1. 用記事本開啟同目錄的 .env 檔案\n")
        f.write("  2. 填入 QWEN_BASE_URL=http://你的伺服器IP:埠號/v1\n")
        f.write("  3. 填入 QWEN_MODEL=模型名稱\n")
        f.write("  4. 儲存後重新啟動\n\n")
        f.write("【錯誤排查】\n")
        f.write("  啟動失敗請查看同目錄的 ExcelAI_log.txt\n")
    print(f"  ✓ 使用說明.txt")


# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 50)
    print(f" Excel AI 助手 {VERSION} — 打包成 EXE")
    print("=" * 50)

    print("\n[1/5] 前置檢查...")
    check_prerequisites()

    print("\n[2/5] 清除舊建置...")
    clean_old_build()

    print("\n[3/5] PyInstaller 打包...")
    build_exe()

    print("\n[4/5] 複製 pywin32 DLL...")
    copy_pywin32_dlls()

    print("\n[5/5] 建立啟動腳本與說明文件...")
    create_launcher_bat()
    create_readme()

    dist_dir = os.path.join(BASE_DIR, "dist", APP_NAME)
    print("\n" + "=" * 50)
    print(f" ✅ 打包完成！版本：{VERSION}")
    print(f" 輸出位置：{dist_dir}")
    print()
    print(" 發佈步驟：")
    print(f"   1. 確認 dist\\{APP_NAME}\\.env 的伺服器設定正確")
    print(f"   2. 將整個 dist\\{APP_NAME}\\ 資料夾壓縮成 ZIP")
    print("   3. 同仁解壓縮後執行「啟動Excel助手.bat」即可")
    print("=" * 50)

    try:
        os.startfile(dist_dir)
    except Exception:
        pass


if __name__ == "__main__":
    main()
