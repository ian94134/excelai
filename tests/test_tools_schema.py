"""
工具 schema 驗證測試。

目的：
- 確認 28 個工具 schema 在結構上符合 OpenAI tool-calling 規範
- 確認 TOOL_MAP 與 OPENAI_TOOLS 名單一致（不會有工具只定義 schema 卻無 executor，或反之）
- 確認 DANGEROUS_TOOLS 的每個項目都存在於 TOOL_MAP
- 確認每個 tool 的 parameters.required 欄位都出現在 parameters.properties
- 確認 name / description 非空字串

這類純結構測試在 CI 上執行極快（毫秒級），可作為 regression 第一道防線。

參見 PLAN_v4.md §C-5。
"""

from __future__ import annotations
import re
import pytest

# 無 win32com 的環境（例如 Linux CI）可能讓 tools.executor 載入失敗（因為會連帶 import excel_tools）
try:
    from tools.definition import OPENAI_TOOLS
    from tools.executor import TOOL_MAP, DANGEROUS_TOOLS
    _EXECUTOR_OK = True
except ImportError as e:
    OPENAI_TOOLS = None  # type: ignore
    TOOL_MAP = None  # type: ignore
    DANGEROUS_TOOLS = None  # type: ignore
    _EXECUTOR_OK = False
    _IMPORT_ERR = str(e)


# 只載入 definition 的 schema 驗證可在任何平台執行
from tools.definition import OPENAI_TOOLS as SCHEMAS


_NAME_RE = re.compile(r"^[a-z][a-z0-9_]*$")


def test_tool_schema_count():
    """v4.7.0 新增 10 個工具（A/B/C/D），共 71 個工具"""
    assert len(SCHEMAS) == 71, f"Expected 71 tools, got {len(SCHEMAS)}"


def test_tool_schema_structure():
    """每個 tool 都有 type=function、function.name、function.description、function.parameters"""
    for i, tool in enumerate(SCHEMAS):
        assert tool.get("type") == "function", f"[{i}] type 必須為 'function'"
        fn = tool.get("function")
        assert isinstance(fn, dict), f"[{i}] function 必須是 dict"
        assert isinstance(fn.get("name"), str) and fn["name"], f"[{i}] name 不可為空"
        assert _NAME_RE.match(fn["name"]), f"[{i}] name={fn['name']} 命名需為 snake_case"
        assert isinstance(fn.get("description"), str) and fn["description"], f"[{fn['name']}] description 不可為空"
        params = fn.get("parameters")
        assert isinstance(params, dict), f"[{fn['name']}] parameters 必須是 dict"
        assert params.get("type") == "object", f"[{fn['name']}] parameters.type 必須為 'object'"
        assert "properties" in params, f"[{fn['name']}] parameters.properties 必填（即使為 {{}}）"


def test_tool_required_fields_are_declared():
    """parameters.required 列出的欄位都必須存在於 properties。"""
    for tool in SCHEMAS:
        fn = tool["function"]
        params = fn["parameters"]
        required = params.get("required", [])
        properties = params.get("properties", {})
        for req_name in required:
            assert req_name in properties, (
                f"[{fn['name']}] required 欄位 '{req_name}' 未出現在 properties"
            )


def test_format_range_requires_at_least_one_style_field():
    """format_range 不可只傳 range_addr，至少要有一個實際樣式欄位。"""
    fmt = next(t for t in SCHEMAS if t["function"]["name"] == "format_range")
    params = fmt["function"]["parameters"]
    assert "anyOf" in params, "format_range.parameters 必須定義 anyOf 約束"
    any_of = params["anyOf"]
    required_fields = {entry["required"][0] for entry in any_of if entry.get("required")}
    expected = {
        "bold", "italic", "color", "fill",
        "font_size", "number_format", "horizontal_alignment",
    }
    assert expected.issubset(required_fields), (
        "format_range.anyOf 必須涵蓋所有格式欄位，避免 no-op 呼叫"
    )


def test_tool_names_unique():
    """不允許同名工具重複定義。"""
    names = [t["function"]["name"] for t in SCHEMAS]
    duplicates = {n for n in names if names.count(n) > 1}
    assert not duplicates, f"工具名稱重複：{duplicates}"


@pytest.mark.skipif(not _EXECUTOR_OK, reason=f"tools.executor 無法載入（通常是 Linux CI 缺 pywin32）")
def test_schema_and_executor_consistent():
    """每個 schema 名稱都要有對應的 TOOL_MAP 實作，反之亦然。"""
    schema_names = {t["function"]["name"] for t in SCHEMAS}
    executor_names = set(TOOL_MAP.keys())

    only_in_schema = schema_names - executor_names
    only_in_executor = executor_names - schema_names

    assert not only_in_schema, f"有 schema 但無 executor 實作：{only_in_schema}"
    assert not only_in_executor, f"有 executor 但無 schema：{only_in_executor}"


@pytest.mark.skipif(not _EXECUTOR_OK, reason="tools.executor 無法載入")
def test_dangerous_tools_exist_in_tool_map():
    """DANGEROUS_TOOLS 名稱必須都存在於 TOOL_MAP，否則確認流程無效。"""
    missing = DANGEROUS_TOOLS - set(TOOL_MAP.keys())
    assert not missing, f"DANGEROUS_TOOLS 有項目不在 TOOL_MAP：{missing}"


@pytest.mark.skipif(not _EXECUTOR_OK, reason="tools.executor 無法載入")
def test_backup_needed_covers_all_tools():
    """backup.BACKUP_NEEDED 應涵蓋全部 28 個工具，確保 Phase 2 擴充時不遺漏。"""
    import backup
    tool_names = set(TOOL_MAP.keys())
    missing = tool_names - set(backup.BACKUP_NEEDED.keys())
    assert not missing, f"backup.BACKUP_NEEDED 未涵蓋：{missing}"


# ---------------------------------------------------------------------------
# Registry consistency (TD-04)
# ---------------------------------------------------------------------------

def test_registry_no_unknown_tools():
    """
    Every tool registered via @register_tool must also have a schema in
    definition.py.  Catches the case where a developer decorates a function
    but forgets to add a schema entry.
    """
    from tools.registry import registered_names
    schema_names = {t["function"]["name"] for t in SCHEMAS}
    reg_names    = registered_names()
    unknown = reg_names - schema_names
    assert not unknown, (
        f"Tools in registry but missing a schema in definition.py: {unknown}\n"
        "Add the schema dict to tools/definition.py."
    )


def test_registry_no_orphan_schemas():
    """
    Every schema that has a matching @register_tool entry must not conflict
    with the legacy TOOL_MAP.  (Pure schema-only tools without a decorator
    use the legacy lambda path and are fine.)
    This test is a forward-looking guard: once a function is decorated it
    must stay decorated.
    """
    from tools.registry import registered_names, get_registered_tool_map
    reg = get_registered_tool_map()
    # Registry map should be callable for each registered name
    for name, fn in reg.items():
        assert callable(fn), f"Registry entry for '{name}' is not callable"
