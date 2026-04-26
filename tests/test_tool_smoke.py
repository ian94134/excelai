from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parent.parent


def _probe_excel_com() -> subprocess.CompletedProcess[str]:
    code = """
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = None
try:
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
finally:
    if app is not None:
        app.Quit()
    pythoncom.CoUninitialize()
"""
    return subprocess.run(
        [sys.executable, "-c", code],
        cwd=ROOT,
        text=True,
        capture_output=True,
        timeout=60,
    )


def test_all_tools_excel_com_smoke(tmp_path: Path):
    if os.environ.get("EXCEL_AI_RUN_COM_SMOKE") != "1":
        pytest.skip("Set EXCEL_AI_RUN_COM_SMOKE=1 to run the Windows Excel COM smoke suite.")

    probe = _probe_excel_com()
    if probe.returncode != 0:
        detail = (probe.stderr or probe.stdout or "").strip()
        pytest.skip(f"Excel COM is not available: {detail}")

    result = subprocess.run(
        [
            sys.executable,
            str(ROOT / "tools_smoke_test.py"),
            "--output-dir",
            str(tmp_path),
        ],
        cwd=ROOT,
        text=True,
        capture_output=True,
        timeout=360,
    )

    assert result.returncode == 0, result.stdout + result.stderr
