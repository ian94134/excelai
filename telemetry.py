"""
telemetry.py — local-only usage statistics.
- Pure local SQLite, no HTTP calls ever.
- Records every tool call: name, duration, status, error type.
- Exposes get_summary() for the sidebar dashboard.

TD-06 refactor: replaced per-call connect/close with a module-level
persistent connection (_conn).  atexit ensures the connection is closed
cleanly on interpreter shutdown.  Thread safety is handled by
check_same_thread=False plus SQLite's built-in serialisation for writes.
"""

import atexit
import sqlite3
import time
from pathlib import Path
from typing import Optional

# Database path: ~/.excel-ai/telemetry.db
_DB_DIR  = Path.home() / ".excel-ai"
_DB_PATH = _DB_DIR / "telemetry.db"

_CREATE_SQL = """
CREATE TABLE IF NOT EXISTS tool_calls (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts          TEXT    NOT NULL,
    tool_name   TEXT    NOT NULL,
    duration_ms INTEGER NOT NULL,
    status      TEXT    NOT NULL,
    error_type  TEXT
);
CREATE INDEX IF NOT EXISTS idx_tool_name ON tool_calls (tool_name);
CREATE INDEX IF NOT EXISTS idx_ts        ON tool_calls (ts);
"""

# ---------------------------------------------------------------------------
# Module-level persistent connection
# ---------------------------------------------------------------------------

_conn: Optional[sqlite3.Connection] = None


def _get_conn() -> Optional[sqlite3.Connection]:
    """
    Return the module-level SQLite connection, creating it on first call.
    Returns None if the DB cannot be opened (e.g. no write permission).
    Subsequent calls always return the same connection object.
    """
    global _conn
    if _conn is not None:
        return _conn
    try:
        _DB_DIR.mkdir(parents=True, exist_ok=True)
        _conn = sqlite3.connect(
            str(_DB_PATH),
            check_same_thread=False,  # safe: SQLite serialises writes
        )
        _conn.executescript(_CREATE_SQL)
        _conn.commit()
        atexit.register(_close_conn)
        return _conn
    except OSError:
        return None


def _close_conn() -> None:
    """atexit handler — flush and close the persistent connection."""
    global _conn
    if _conn is not None:
        try:
            _conn.commit()
            _conn.close()
        except Exception:
            pass
        _conn = None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def record(tool_name: str, duration_ms: int, status: str,
           error_type: Optional[str] = None) -> None:
    """Record one tool call.  Fails silently — never disrupts the main flow."""
    try:
        conn = _get_conn()
        if conn is None:
            return
        ts = time.strftime("%Y-%m-%dT%H:%M:%S")
        conn.execute(
            "INSERT INTO tool_calls (ts, tool_name, duration_ms, status, error_type) "
            "VALUES (?, ?, ?, ?, ?)",
            (ts, tool_name, duration_ms, status, error_type),
        )
        conn.commit()
    except Exception:
        pass


def get_summary() -> dict:
    """
    Return usage statistics:
    {
        "total": int,
        "success_rate": float,              # 0.0–1.0
        "top_tools": [(name, count)],       # top 5 by call count
        "slowest_tools": [(name, avg_ms)],  # top 5 by avg duration (min 3 calls)
        "recent_errors": [(tool, error_type, ts)],  # 5 most recent errors
    }
    Returns an empty structure when no data or DB unavailable.
    """
    empty: dict = {
        "total": 0, "success_rate": 0.0,
        "top_tools": [], "slowest_tools": [], "recent_errors": [],
    }
    try:
        conn = _get_conn()
        if conn is None:
            return empty

        total = conn.execute("SELECT COUNT(*) FROM tool_calls").fetchone()[0]
        if total == 0:
            return empty

        ok_count = conn.execute(
            "SELECT COUNT(*) FROM tool_calls WHERE status = 'ok'"
        ).fetchone()[0]

        top_tools = conn.execute(
            "SELECT tool_name, COUNT(*) AS cnt FROM tool_calls "
            "GROUP BY tool_name ORDER BY cnt DESC LIMIT 5"
        ).fetchall()

        slowest = conn.execute(
            "SELECT tool_name, AVG(duration_ms) AS avg_ms "
            "FROM tool_calls WHERE status = 'ok' "
            "GROUP BY tool_name HAVING COUNT(*) >= 3 "
            "ORDER BY avg_ms DESC LIMIT 5"
        ).fetchall()

        recent_errors = conn.execute(
            "SELECT tool_name, error_type, ts FROM tool_calls "
            "WHERE status != 'ok' AND error_type IS NOT NULL "
            "ORDER BY id DESC LIMIT 5"
        ).fetchall()

        return {
            "total": total,
            "success_rate": ok_count / total if total else 0.0,
            "top_tools": list(top_tools),
            "slowest_tools": [(n, round(ms)) for n, ms in slowest],
            "recent_errors": list(recent_errors),
        }
    except Exception:
        return empty


def clear() -> None:
    """Delete all recorded data."""
    try:
        conn = _get_conn()
        if conn is None:
            return
        conn.execute("DELETE FROM tool_calls")
        conn.commit()
    except Exception:
        pass
