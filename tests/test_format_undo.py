from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock


_WIN_STUBS = [
    "pythoncom", "pywintypes", "win32com", "win32com.client", "win32con", "win32api",
]
for _m in _WIN_STUBS:
    sys.modules.setdefault(_m, MagicMock())


ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import excel.format as fmt
import backup
import excel._undo as undo_mod
import excel_tools


class _Font:
    def __init__(self) -> None:
        self.Bold = False
        self.Italic = False
        self.Color = 0
        self.Size = 11


class _Interior:
    def __init__(self) -> None:
        self.ColorIndex = fmt._INTERIOR_NONE_IDX
        self.Color = 16777215


class _Cell:
    def __init__(self, address: str = "$H$14") -> None:
        self.Address = address
        self.Font = _Font()
        self.Interior = _Interior()
        self.NumberFormat = "General"
        self.HorizontalAlignment = -4131


class _Range:
    def __init__(self, cell: _Cell) -> None:
        self._cell = cell
        self.Count = 1
        self.Font = cell.Font
        self.Interior = cell.Interior

    def __iter__(self):
        return iter([self._cell])

    @property
    def NumberFormat(self):
        return self._cell.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self._cell.NumberFormat = value

    @property
    def HorizontalAlignment(self):
        return self._cell.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self._cell.HorizontalAlignment = value


class _Sheet:
    def __init__(self, cell: _Cell) -> None:
        self._cell = cell

    def Range(self, _addr: str) -> _Range:
        return _Range(self._cell)


def test_capture_formats_before_records_no_fill_without_nameerror(monkeypatch):
    cell = _Cell()
    sheet = _Sheet(cell)
    monkeypatch.setattr(fmt, "_get_excel", lambda: object())
    monkeypatch.setattr(fmt, "_get_sheet", lambda _excel, _sheet_name=None: sheet)

    result = fmt.capture_formats_before(
        "H14",
        "Dashboard",
        {
            "_tool_type": "format_range",
            "bold": True,
            "color": "#FF0000",
            "fill": "#FFFF00",
            "horizontal_alignment": "center",
        },
    )

    assert result["type"] == "format_range"
    assert result["cells"] == [{
        "address": "$H$14",
        "bold": False,
        "color": 0,
        "fill": None,
        "horizontal_alignment": -4131,
    }]


def test_restore_formats_can_clear_fill_and_restore_style(monkeypatch):
    cell = _Cell()
    cell.Font.Bold = True
    cell.Font.Color = 255
    cell.Interior.ColorIndex = 1
    cell.Interior.Color = 65535
    cell.HorizontalAlignment = -4108
    sheet = _Sheet(cell)
    monkeypatch.setattr(fmt, "_get_excel", lambda: object())
    monkeypatch.setattr(fmt, "_get_sheet", lambda _excel, _sheet_name=None: sheet)

    fmt._restore_formats(
        {
            "type": "format_range",
            "range": "H14",
            "cells": [{
                "address": "$H$14",
                "bold": False,
                "color": 0,
                "fill": None,
                "horizontal_alignment": -4131,
            }],
        },
        "Dashboard",
    )

    assert cell.Font.Bold is False
    assert cell.Font.Color == 0
    assert cell.Interior.ColorIndex == fmt._INTERIOR_NONE_IDX
    assert cell.HorizontalAlignment == -4131


def test_capture_borders_uses_defined_border_indexes(monkeypatch):
    class _Border:
        LineStyle = 1
        Color = 0

    class _BorderCell(_Cell):
        def Borders(self, _idx: int) -> _Border:
            return _Border()

    cell = _BorderCell()
    sheet = _Sheet(cell)
    monkeypatch.setattr(fmt, "_get_excel", lambda: object())
    monkeypatch.setattr(fmt, "_get_sheet", lambda _excel, _sheet_name=None: sheet)

    result = fmt.capture_formats_before(
        "H14",
        "Dashboard",
        {"_tool_type": "set_borders"},
    )

    assert result["type"] == "set_borders"
    assert set(result["cells"][0]["borders"]) == set(fmt._ALL_BORDER_IDX)


def test_undo_last_persists_after_popping_entry(monkeypatch):
    stack = backup.BackupStack()
    stack.push(backup.BackupEntry(
        tool_name="write_range",
        arguments={"range_addr": "H14", "sheet": "Dashboard"},
        values_before=[["before"]],
    ))
    saved = []
    monkeypatch.setattr(backup, "get_session_stack", lambda: stack)
    monkeypatch.setattr(backup, "save_current_stack", lambda: saved.append(True))
    monkeypatch.setattr(undo_mod, "_undo_dispatch", lambda _entry: {"status": "ok"})

    result = undo_mod.undo_last()

    assert result == {"status": "ok"}
    assert len(stack) == 0
    assert saved == [True]


def test_backup_restore_formats_uses_restore_formats(monkeypatch):
    formats_before = {
        "type": "format_range",
        "range": "H14",
        "cells": [{"address": "$H$14", "bold": False}],
    }
    entry = backup.BackupEntry(
        tool_name="format_range",
        arguments={"range_addr": "H14", "sheet": "Dashboard"},
        formats_before=formats_before,
    )
    calls = []
    monkeypatch.setattr(
        excel_tools,
        "_restore_formats",
        lambda formats, sheet: calls.append((formats, sheet)),
    )

    result = backup.restore(entry)

    assert result == {"status": "ok", "undone": "format_range", "method": "formats_restore"}
    assert calls == [(formats_before, "Dashboard")]
