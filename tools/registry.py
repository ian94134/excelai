"""
tools/registry.py — single source of truth for tool function bindings.

PROBLEM BEFORE: adding a tool required touching 5 files:
  excel_tools.py  (implementation)
  tools/definition.py  (schema)
  tools/executor.py    (TOOL_MAP)
  backup.py            (BACKUP_NEEDED)
  tests/test_tools_schema.py  (hardcoded count)

SOLUTION: @register_tool decorator in excel_tools.py provides the function
binding and backup flag.  executor.py merges the registry into TOOL_MAP.
backup.py merges it into BACKUP_NEEDED.  Schemas stay in definition.py.

Adding a new tool now only requires:
  1. Add @register_tool(needs_backup=...) to the function in excel_tools.py
  2. Add the schema dict to tools/definition.py

The test_tools_schema consistency check validates that every registered
function has a matching schema and vice-versa.
"""
from __future__ import annotations
from typing import Callable

_REGISTRY: dict[str, dict] = {}


def register_tool(needs_backup: bool = True) -> Callable:
    """
    Decorator that registers an excel_tools function into the global registry.

    Usage:
        @register_tool(needs_backup=True)
        def write_range(range_addr, values, sheet=None):
            ...

    The function name becomes the tool name.  executor.py picks up the
    callable; backup.py picks up the needs_backup flag.
    """
    def decorator(fn: Callable) -> Callable:
        name = fn.__name__
        _REGISTRY[name] = {
            "fn":           fn,
            "needs_backup": needs_backup,
        }
        return fn
    return decorator


def get_registered_tool_map() -> dict[str, Callable]:
    """
    Return {tool_name: lambda args: fn(**args)} for all registered tools.
    executor.py merges this with its legacy TOOL_MAP.
    """
    return {
        name: (lambda a, f=entry["fn"]: f(**a))
        for name, entry in _REGISTRY.items()
    }


def get_registered_backup_needed() -> dict[str, bool]:
    """
    Return {tool_name: needs_backup} for all registered tools.
    backup.py merges this with its legacy BACKUP_NEEDED dict.
    """
    return {name: entry["needs_backup"] for name, entry in _REGISTRY.items()}


def registered_names() -> set[str]:
    """Return the set of all registered tool names."""
    return set(_REGISTRY.keys())
