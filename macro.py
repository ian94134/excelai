"""
巨集錄製與重播系統（v4.7.0）

設計原則：
- 巨集 = 一組有序的 {tool, args} 步驟
- 錄製來源：BackupStack（最近操作歷史）或呼叫端明確提供 steps
- 儲存格式：~/.excel-ai/macros.json（版本化，原子寫入）
- 執行機制：透過 execute_batch 執行，自動回滾失敗步驟
- 本模組無 Streamlit / win32com 依賴，可在 Linux CI 測試
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

# backup をモジュールレベルでインポートしておくことで patch("backup.get_session_stack")
# が sys.modules 経由で正しく解決できる（テスト互換性のため）
try:
    import backup as _backup_module  # noqa: F401
except ImportError:
    _backup_module = None  # type: ignore

_MACROS_PATH    = Path.home() / ".excel-ai" / "macros.json"
_MACROS_VERSION = 1


# ── 持久化 ────────────────────────────────────────────────────────────────────

def _load_macros() -> dict[str, Any]:
    """從磁碟載入巨集字典。格式不符或檔案不存在時回傳空字典。"""
    try:
        if not _MACROS_PATH.exists():
            return {}
        data = json.loads(_MACROS_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or data.get("version") != _MACROS_VERSION:
            return {}
        macros = data.get("macros", {})
        return macros if isinstance(macros, dict) else {}
    except Exception:
        return {}


def _save_macros(macros: dict[str, Any]) -> None:
    """原子寫入巨集字典（.tmp → replace）；失敗靜默吸收。"""
    try:
        _MACROS_PATH.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(
            {"version": _MACROS_VERSION, "macros": macros},
            ensure_ascii=False,
            indent=2,
        )
        tmp = _MACROS_PATH.with_suffix(".tmp")
        tmp.write_text(payload, encoding="utf-8")
        tmp.replace(_MACROS_PATH)
    except Exception:
        pass


# ── 公開 API ──────────────────────────────────────────────────────────────────

def record_macro(
    name: str,
    description: str = "",
    steps: list[dict] | None = None,
) -> dict:
    """
    錄製巨集並寫入磁碟。

    Parameters
    ----------
    name        : 巨集名稱（唯一鍵，相同名稱會覆蓋）
    description : 巨集說明文字（可省略）
    steps       : 步驟清單 [{"tool": str, "args": dict}, ...]；
                  若省略則自動從 BackupStack 取得最近操作

    Returns
    -------
    dict  {"status": "ok", "name", "step_count", "message"}
          or {"status": "error", "message"}
    """
    name = (name or "").strip()
    if not name:
        return {"status": "error", "message": "巨集名稱不能為空"}

    if steps is None:
        # 從 BackupStack 取得最近操作歷史
        try:
            from backup import get_session_stack
            stack = get_session_stack()
        except Exception:
            stack = None

        if stack is None or len(stack) == 0:
            return {"status": "error", "message": "沒有可錄製的操作歷史，請先執行一些 Excel 操作"}

        steps = [
            {"tool": entry.tool_name, "args": dict(entry.arguments)}
            for entry in stack.snapshot()
        ]

    if not steps:
        return {"status": "error", "message": "步驟清單為空，無法錄製巨集"}

    macros = _load_macros()
    macros[name] = {
        "description": description,
        "steps": steps,
        "step_count": len(steps),
    }
    _save_macros(macros)

    return {
        "status":     "ok",
        "name":       name,
        "step_count": len(steps),
        "message":    f"已錄製巨集「{name}」，共 {len(steps)} 個步驟",
    }


def list_macros() -> dict:
    """
    列出所有已儲存的巨集。

    Returns
    -------
    dict  {"status": "ok", "macros": [...], "count": int}
    """
    macros = _load_macros()
    macro_list = [
        {
            "name":        name,
            "description": info.get("description", ""),
            "step_count":  info.get("step_count", len(info.get("steps", []))),
        }
        for name, info in macros.items()
    ]
    return {"status": "ok", "macros": macro_list, "count": len(macro_list)}


def _dangerous_steps(steps: list[dict]) -> list[dict]:
    """Return dangerous macro steps with 1-based indices for UI confirmation."""
    from tools.executor import DANGEROUS_TOOLS

    found: list[dict] = []
    for index, step in enumerate(steps, start=1):
        tool_name = step.get("tool", "")
        if tool_name in DANGEROUS_TOOLS:
            found.append({
                "index": index,
                "tool": tool_name,
                "args": step.get("args", {}),
            })
    return found


def _is_error_result(result: object) -> bool:
    return isinstance(result, dict) and (
        "error" in result or result.get("status") == "error"
    )


def run_macro(name: str, confirm_dangerous: bool = False) -> dict:
    """
    執行已儲存的巨集。失敗時 execute_batch 自動回滾已執行步驟。

    Returns
    -------
    dict  {"status": "ok"/"error", "name", "total_steps", "executed_steps",
           "results": [...], "message"}
    """
    name = (name or "").strip()
    if not name:
        return {"status": "error", "message": "請提供巨集名稱"}

    macros = _load_macros()
    macro  = macros.get(name)
    if macro is None:
        return {
            "status":          "error",
            "message":         f"找不到巨集「{name}」",
            "available_macros": list(macros.keys()),
        }

    steps = macro.get("steps", [])
    if not steps:
        return {"status": "error", "message": f"巨集「{name}」沒有步驟"}

    dangerous = _dangerous_steps(steps)
    if dangerous and not confirm_dangerous:
        tools = "、".join(f"第 {s['index']} 步 {s['tool']}" for s in dangerous)
        return {
            "status": "error",
            "error_type": "DangerousMacroRequiresConfirmation",
            "requires_confirmation": True,
            "name": name,
            "dangerous_steps": dangerous,
            "message": f"巨集「{name}」包含危險工具（{tools}），請確認後再執行",
        }

    from tools.executor import execute_batch
    results = execute_batch(steps, confirm_dangerous=confirm_dangerous)

    failed = [r for r in results if _is_error_result(r.get("result", {}))]
    return {
        "status":         "error" if failed else "ok",
        "name":           name,
        "total_steps":    len(steps),
        "executed_steps": len(results),
        "results":        results,
        "message": (
            f"巨集「{name}」執行完成，共 {len(steps)} 步驟"
            if not failed
            else f"巨集「{name}」在第 {len(results)} 步失敗，已自動回滾"
        ),
    }


def delete_macro(name: str) -> dict:
    """
    刪除指定巨集。

    Returns
    -------
    dict  {"status": "ok"/"error", "message"}
    """
    name = (name or "").strip()
    if not name:
        return {"status": "error", "message": "請提供巨集名稱"}

    macros = _load_macros()
    if name not in macros:
        return {
            "status":          "error",
            "message":         f"找不到巨集「{name}」",
            "available_macros": list(macros.keys()),
        }

    del macros[name]
    _save_macros(macros)
    return {"status": "ok", "message": f"已刪除巨集「{name}」"}


def get_macro_steps(name: str) -> list[dict]:
    """
    取得巨集的步驟清單（供 UI 展示，不執行）。
    找不到時回傳空清單。
    """
    macros = _load_macros()
    macro  = macros.get((name or "").strip())
    return macro.get("steps", []) if macro else []
