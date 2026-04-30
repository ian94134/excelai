"""Windows executable launcher for the Excel AI Streamlit app."""
from __future__ import annotations

import os
import sys
from pathlib import Path


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parents[1]


def main() -> int:
    app_root = _runtime_root()
    app_path = app_root / "main.py"
    if not app_path.exists():
        raise FileNotFoundError(f"Cannot find Streamlit entrypoint: {app_path}")

    os.chdir(app_root)
    sys.path.insert(0, str(app_root))

    port = os.environ.get("EXCEL_AI_PORT", "8501")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    from streamlit.web import cli as streamlit_cli

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address=localhost",
        f"--server.port={port}",
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    return int(streamlit_cli.main() or 0)


if __name__ == "__main__":
    raise SystemExit(main())
