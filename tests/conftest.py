"""
pytest 共用 fixtures。

設計原則：
- 不載入 win32com（Linux CI 無此套件），由 fixture 提供 MagicMock
- 不載入 streamlit session_state 需要的完整 runtime，由 fixture 提供 dict 替身
- 測試重點在 schema / session / backup / provider 的純邏輯層，win32com 真實互動保留在 MANUAL_TEST.md
"""

from __future__ import annotations
import sys
from pathlib import Path
from unittest.mock import MagicMock
import pytest

# 讓 tests/ 內部的 import 能找到專案根（等同把 excel-ai 加入 PYTHONPATH）
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# ── Windows-only COM 模組替身（Linux CI 無此套件）────────────────────────────
# 在任何 excel_tools / executor 被匯入之前，把缺少的模組塞進 sys.modules，
# 讓 import 時不會拋出 ModuleNotFoundError。
_WIN_STUBS = [
    "pythoncom", "pywintypes",
    "win32com", "win32com.client",
    "win32con", "win32api",
]
for _mod in _WIN_STUBS:
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()


# ── Streamlit session_state 替身 ─────────────────────────────────────────────

class _FakeSessionState(dict):
    """模擬 streamlit.session_state 的 dict-like 物件，加上 __getattr__/__setattr__ 存取。"""

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError as e:
            raise AttributeError(name) from e

    def __setattr__(self, name, value):
        self[name] = value


@pytest.fixture
def fake_session_state(monkeypatch):
    """把 streamlit.session_state 換成 dict 替身；適用於依賴 session 的測試。"""
    try:
        import streamlit as st
    except ImportError:
        pytest.skip("streamlit 未安裝，跳過需要 session_state 的測試")

    fake = _FakeSessionState()
    monkeypatch.setattr(st, "session_state", fake, raising=False)
    return fake


# ── 假 LLM Provider（避免測試打到真實 Qwen）──────────────────────────────────

@pytest.fixture
def fake_provider():
    """最小化的 Provider mock：chat() 回傳可控的 LLMResponse，chat_stream() 可 yield 固定序列。"""
    provider = MagicMock()
    provider.chat.return_value = MagicMock(text="mocked-summary", tool_calls=[])
    provider.chat_stream.return_value = iter([("done", "mocked-done")])
    return provider


# ── 假 Excel Application（阻止測試真的連到 COM）─────────────────────────────

@pytest.fixture
def fake_excel(monkeypatch):
    """
    把 excel_tools._get_excel / _get_sheet 換成 MagicMock。
    測試只想驗證邏輯路徑，不實際操作 COM。
    """
    try:
        import excel_tools  # noqa: F401
    except ImportError:
        pytest.skip("excel_tools 需 pywin32，跳過於 Linux CI")

    excel = MagicMock()
    sheet = MagicMock()
    monkeypatch.setattr("excel_tools._get_excel", lambda: excel)
    monkeypatch.setattr("excel_tools._get_sheet", lambda e=None, n=None: sheet)
    return excel, sheet
