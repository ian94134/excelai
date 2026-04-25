"""
tests/test_main_helpers.py

Regression tests for pure helper functions in main.py that have no
Streamlit or win32com dependency.

Currently covers:
  _inject_ephemeral_system_ctx  — prevents the Qwen 500 error caused by
      inserting a second system message into the conversation history.
      Qwen's Jinja chat template requires the system role to appear at most
      once and only at index 0.
"""
from __future__ import annotations
import pathlib
import re

import pytest

# ---------------------------------------------------------------------------
# Extract _inject_ephemeral_system_ctx from main.py without importing it
# (importing main.py would execute Streamlit page-config calls and fail).
# ---------------------------------------------------------------------------

_MAIN_PY = pathlib.Path(__file__).resolve().parent.parent / "main.py"


def _load_fn(fn_name: str):
    """
    Parse main.py source, extract the named top-level function via a simple
    indentation-based scan, compile, and return the function object.
    """
    src = _MAIN_PY.read_text(encoding="utf-8")
    lines = src.splitlines()

    start = None
    for i, line in enumerate(lines):
        if re.match(rf"^def {fn_name}\(", line):
            start = i
            break
    assert start is not None, f"{fn_name} not found in main.py"

    fn_lines = [lines[start]]
    for line in lines[start + 1:]:
        if line and not line[0].isspace():
            break
        fn_lines.append(line)

    globs: dict = {}
    exec("\n".join(fn_lines), globs)
    return globs[fn_name]


@pytest.fixture(scope="module")
def inject():
    return _load_fn("_inject_ephemeral_system_ctx")


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestInjectEphemeralSystemCtx:
    """
    Guard against Qwen 500: 'System message must be at the beginning.'

    Root cause: prepending a new system message pushes the existing one to
    index 1, which Qwen's Jinja template rejects.
    Fix: merge ctx into the existing system message content instead.
    """

    def test_merges_into_existing_system_message(self, inject):
        msgs = [
            {"role": "system",    "content": "You are an Excel assistant."},
            {"role": "user",      "content": "hi"},
            {"role": "assistant", "content": "hello"},
        ]
        result = inject(msgs, "[wb: Book1.xlsx, active: Sheet1]")

        system_msgs = [m for m in result if m["role"] == "system"]
        assert len(system_msgs) == 1,         "must have exactly one system message"
        assert result[0]["role"] == "system", "system message must be at index 0"
        assert "[wb: Book1.xlsx" in result[0]["content"]
        assert "You are an Excel assistant." in result[0]["content"]

    def test_no_duplicate_system_on_long_history(self, inject):
        """Simulate a 10-round conversation — still only one system message."""
        msgs = [{"role": "system", "content": "base prompt"}]
        for i in range(10):
            msgs.append({"role": "user",      "content": f"q{i}"})
            msgs.append({"role": "assistant", "content": f"a{i}"})
        result = inject(msgs, "wb-context")
        system_count = sum(1 for m in result if m["role"] == "system")
        assert system_count == 1
        assert result[0]["role"] == "system"

    def test_prepends_when_no_system_message(self, inject):
        """Edge case: conversation started without a system message."""
        msgs = [{"role": "user", "content": "hi"}]
        result = inject(msgs, "[wb: Book1.xlsx]")
        assert result[0]["role"] == "system"
        assert result[0]["content"] == "[wb: Book1.xlsx]"
        assert result[1]["role"] == "user"

    def test_noop_on_empty_context(self, inject):
        """Empty context string must return an equivalent list unchanged."""
        msgs = [
            {"role": "system", "content": "base"},
            {"role": "user",   "content": "q"},
        ]
        result = inject(msgs, "")
        assert result == msgs

    def test_does_not_mutate_original_list(self, inject):
        """inject must be non-destructive — never modify caller's list."""
        original = "original system content"
        msgs = [{"role": "system", "content": original}]
        inject(msgs, "extra context")
        assert msgs[0]["content"] == original

    def test_does_not_mutate_original_dict(self, inject):
        """inject must not mutate the dict objects inside the list either."""
        msg = {"role": "system", "content": "original"}
        msgs = [msg]
        inject(msgs, "extra")
        assert msg["content"] == "original"

    def test_empty_msgs_list(self, inject):
        """Empty message list gets a new system message prepended."""
        result = inject([], "wb-ctx")
        assert result == [{"role": "system", "content": "wb-ctx"}]


# ---------------------------------------------------------------------------
# utils.col_letter / col_index
# ---------------------------------------------------------------------------

from utils import col_letter, col_index


class TestColLetter:
    def test_single_letters(self):
        assert col_letter(1)  == "A"
        assert col_letter(26) == "Z"

    def test_double_letters(self):
        assert col_letter(27) == "AA"
        assert col_letter(52) == "AZ"
        assert col_letter(53) == "BA"
        assert col_letter(702) == "ZZ"

    def test_triple(self):
        assert col_letter(703) == "AAA"

    def test_invalid(self):
        import pytest
        with pytest.raises(ValueError):
            col_letter(0)


class TestColIndex:
    def test_single(self):
        assert col_index("A")  == 1
        assert col_index("Z")  == 26

    def test_double(self):
        assert col_index("AA") == 27
        assert col_index("ZZ") == 702

    def test_case_insensitive(self):
        assert col_index("a") == col_index("A")

    def test_roundtrip(self):
        for i in [1, 26, 27, 100, 702, 703]:
            assert col_index(col_letter(i)) == i

    def test_invalid(self):
        import pytest
        with pytest.raises(ValueError):
            col_index("A1")
