"""
tests/test_backup_persist.py

Unit tests for BackupStack persistence (v4.6.0).
Covers: _entry_to_dict / _dict_to_entry round-trip, save_current_stack,
_load_stack, get_session_stack auto-load, atomic write (tmp file replaced),
version mismatch discard, malformed-entry tolerance.
"""
from __future__ import annotations
import sys, json, tempfile
from pathlib import Path
from datetime import datetime, timezone
from unittest.mock import MagicMock, patch

# ── COM stubs (must come before any project import) ──────────────────────────
_WIN_STUBS = ["pythoncom", "pywintypes", "win32com", "win32com.client",
              "win32con", "win32api"]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
import backup
from backup import (
    BackupEntry, BackupStack,
    _entry_to_dict, _dict_to_entry,
    save_current_stack, _load_stack,
    _PERSIST_VERSION,
)


# ── helpers ──────────────────────────────────────────────────────────────────

def _make_entry(tool="write_range", args=None, ts=None) -> BackupEntry:
    return BackupEntry(
        tool_name=tool,
        arguments=args or {"range_addr": "A1:B2", "sheet": "Sheet1"},
        timestamp=ts or datetime(2026, 4, 25, 12, 0, 0, tzinfo=timezone.utc),
        values_before=[[1, 2], [3, 4]],
    )


def _make_stack(*tools) -> BackupStack:
    stk = BackupStack()
    for t in tools:
        stk.push(_make_entry(t))
    return stk


# ── _entry_to_dict / _dict_to_entry round-trip ───────────────────────────────

def test_entry_to_dict_has_required_keys():
    e = _make_entry()
    d = _entry_to_dict(e)
    assert d["tool_name"] == "write_range"
    assert d["arguments"] == {"range_addr": "A1:B2", "sheet": "Sheet1"}
    assert "timestamp" in d
    assert d["values_before"] == [[1, 2], [3, 4]]


def test_entry_to_dict_timestamp_is_iso():
    e = _make_entry()
    d = _entry_to_dict(e)
    # Should be parseable as ISO datetime
    parsed = datetime.fromisoformat(d["timestamp"])
    assert parsed.year == 2026


def test_dict_to_entry_round_trip():
    original = _make_entry()
    d = _entry_to_dict(original)
    restored = _dict_to_entry(d)
    assert restored.tool_name == original.tool_name
    assert restored.arguments == original.arguments
    assert restored.values_before == original.values_before
    assert restored.timestamp.year == original.timestamp.year


def test_dict_to_entry_bad_timestamp_defaults_to_now():
    d = _entry_to_dict(_make_entry())
    d["timestamp"] = "not-a-date"
    entry = _dict_to_entry(d)
    assert entry.timestamp.tzinfo is not None   # defaults to aware UTC


def test_dict_to_entry_missing_tool_name_defaults():
    entry = _dict_to_entry({"arguments": {}})
    assert entry.tool_name == "unknown"


def test_entry_none_fields_preserved():
    e = BackupEntry(tool_name="add_sheet", arguments={"name": "New"})
    d = _entry_to_dict(e)
    restored = _dict_to_entry(d)
    assert restored.values_before is None
    assert restored.formats_before is None


# ── _load_stack / save_current_stack ─────────────────────────────────────────

def test_load_stack_missing_file_returns_none(tmp_path):
    with patch.object(backup, "_PERSIST_PATH", tmp_path / "nonexistent.json"):
        result = _load_stack()
    assert result is None


def test_load_stack_version_mismatch_returns_none(tmp_path):
    p = tmp_path / "backup_stack.json"
    p.write_text(json.dumps({"version": 999, "entries": []}), encoding="utf-8")
    with patch.object(backup, "_PERSIST_PATH", p):
        result = _load_stack()
    assert result is None


def test_load_stack_empty_entries_returns_none(tmp_path):
    p = tmp_path / "backup_stack.json"
    p.write_text(json.dumps({"version": _PERSIST_VERSION, "entries": []}), encoding="utf-8")
    with patch.object(backup, "_PERSIST_PATH", p):
        result = _load_stack()
    assert result is None


def test_load_stack_restores_entries(tmp_path):
    entries = [_entry_to_dict(_make_entry("write_range")),
               _entry_to_dict(_make_entry("format_range"))]
    p = tmp_path / "backup_stack.json"
    p.write_text(json.dumps({"version": _PERSIST_VERSION, "entries": entries}), encoding="utf-8")
    with patch.object(backup, "_PERSIST_PATH", p):
        stk = _load_stack()
    assert stk is not None
    assert len(stk) == 2
    assert stk.peek().tool_name == "format_range"


def test_load_stack_skips_malformed_entry(tmp_path):
    good = _entry_to_dict(_make_entry("write_range"))
    bad  = {"__bad__": True}   # will fail _dict_to_entry but not crash
    p = tmp_path / "backup_stack.json"
    p.write_text(json.dumps({"version": _PERSIST_VERSION, "entries": [good, bad]}), encoding="utf-8")
    with patch.object(backup, "_PERSIST_PATH", p):
        stk = _load_stack()
    # good entry should still load
    assert stk is not None
    assert len(stk) >= 1


def test_save_current_stack_writes_file(tmp_path, monkeypatch):
    """save_current_stack() should write a valid JSON file."""
    p = tmp_path / "backup_stack.json"
    stk = _make_stack("write_range", "format_range")

    # Patch get_session_stack to return our stack
    monkeypatch.setattr(backup, "get_session_stack", lambda: stk)
    with patch.object(backup, "_PERSIST_PATH", p):
        save_current_stack()

    assert p.exists()
    payload = json.loads(p.read_text(encoding="utf-8"))
    assert payload["version"] == _PERSIST_VERSION
    assert len(payload["entries"]) == 2


def test_save_current_stack_atomic_no_tmp_left(tmp_path, monkeypatch):
    """Temp file should be replaced (not left on disk) after successful write."""
    p = tmp_path / "backup_stack.json"
    stk = _make_stack("write_range")
    monkeypatch.setattr(backup, "get_session_stack", lambda: stk)
    with patch.object(backup, "_PERSIST_PATH", p):
        save_current_stack()
    tmp_file = p.with_suffix(".tmp")
    assert not tmp_file.exists()


def test_save_current_stack_silent_on_no_stack(monkeypatch):
    """save_current_stack() should not raise if get_session_stack returns None."""
    monkeypatch.setattr(backup, "get_session_stack", lambda: None)
    save_current_stack()  # must not raise


def test_round_trip_save_and_load(tmp_path, monkeypatch):
    """Full round-trip: save stack, reload it, verify entries match."""
    original_stk = _make_stack("write_range", "set_borders", "insert_row")
    monkeypatch.setattr(backup, "get_session_stack", lambda: original_stk)
    p = tmp_path / "backup_stack.json"
    with patch.object(backup, "_PERSIST_PATH", p):
        save_current_stack()
        loaded = _load_stack()

    assert loaded is not None
    assert len(loaded) == 3
    names = [e.tool_name for e in loaded.snapshot()]
    assert "write_range" in names
    assert "insert_row" in names


def test_get_session_stack_auto_loads_persisted(tmp_path, monkeypatch):
    """get_session_stack() should restore stack from file on first access."""
    entries = [_entry_to_dict(_make_entry("write_range"))]
    p = tmp_path / "backup_stack.json"
    p.write_text(json.dumps({"version": _PERSIST_VERSION, "entries": entries}), encoding="utf-8")

    # Save original streamlit module so we can restore it afterwards
    _orig_streamlit = sys.modules.get("streamlit")

    # Fake streamlit session_state
    fake_state = {}
    fake_st = MagicMock()
    fake_st.session_state = fake_state
    sys.modules["streamlit"] = fake_st

    try:
        with patch.object(backup, "_PERSIST_PATH", p):
            stk = backup.get_session_stack()

        assert stk is not None
        assert len(stk) == 1
        assert stk.peek().tool_name == "write_range"
    finally:
        # IMPORTANT: restore the ORIGINAL streamlit module object (not a new MagicMock).
        # Creating a new MagicMock here would break subsequent tests because session.py's
        # module-level `st` variable is still bound to the original object.
        if _orig_streamlit is not None:
            sys.modules["streamlit"] = _orig_streamlit
        else:
            sys.modules.pop("streamlit", None)
