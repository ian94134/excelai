"""Small repeatable Web UI smoke runner for the Streamlit Excel AI app.

The runner intentionally keeps the default smoke short. It exercises the
browser surface, the model-to-tool routing layer, and Excel COM verification
without re-running the full 71-tool matrix.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
import urllib.request
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


PROJECT_ROOT = Path(__file__).resolve().parent
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "web_ui_smoke_base.xlsx"
WORKBOOK_PATH = PROJECT_ROOT / "web_ui_smoke_work.xlsx"
LOG_PATH = PROJECT_ROOT / "web_ui_smoke_log.json"
DEFAULT_URL = "http://localhost:8501/"
WAIT_TIMEOUT_SEC = 90


@dataclass(frozen=True)
class SmokeCase:
    name: str
    tool: str
    prompt: str


@dataclass
class SmokeResult:
    name: str
    tool: str
    status: str
    duration_ms: int
    evidence: str
    assistant_preview: str = ""
    error: str | None = None


SMOKE_CASES = [
    SmokeCase(
        name="fill_series_natural_phrase",
        tool="fill_series",
        prompt="請直接呼叫 fill_series。從 Report!K1 開始往下填 1 到 5。",
    ),
    SmokeCase(
        name="query_range_filter_sum",
        tool="query_range",
        prompt=(
            "請直接呼叫 query_range。sheet=SalesData，range_addr=A1:G10，"
            "query=Region=North 的所有列，並把 Amount 加總。"
        ),
    ),
    SmokeCase(
        name="beautify_range_business_blue",
        tool="beautify_range",
        prompt=(
            "請直接呼叫 beautify_range。sheet=SalesData，range_addr=A1:G10，"
            "theme=blue，has_header=true，freeze_header=false。"
        ),
    ),
    SmokeCase(
        name="write_range_undo_marker",
        tool="write_range",
        prompt=(
            "請直接呼叫 write_range 工具。sheet=Report，range_addr=Z1，"
            'values=[["SMOKE_UNDO"]]。'
        ),
    ),
    SmokeCase(
        name="undo_last_blank_restore",
        tool="undo_last",
        prompt="請直接呼叫 undo_last 工具，撤銷剛剛對 Report!Z1 寫入 SMOKE_UNDO 的動作。",
    ),
]

PLAIN_LANGUAGE_SMOKE_CASES = [
    SmokeCase(
        name="plain_beautify_salesdata_report",
        tool="beautify_range",
        prompt="幫我把 SalesData 這張表整理得漂亮一點，做成可以給主管看的樣子。",
    ),
    SmokeCase(
        name="plain_query_north_amount_sum",
        tool="query_range",
        prompt="這張表裡 North 的總金額是多少？",
    ),
    SmokeCase(
        name="plain_write_report_aa1",
        tool="write_range",
        prompt="幫我在 Report 工作表的 AA1 寫上「白話輸入OK」。",
    ),
]

CASE_SETS = {
    "tool": SMOKE_CASES,
    "plain": PLAIN_LANGUAGE_SMOKE_CASES,
    "all": SMOKE_CASES + PLAIN_LANGUAGE_SMOKE_CASES,
}


class SmokeError(RuntimeError):
    pass


def _ensure_fixture(path: Path) -> None:
    if path.exists():
        return
    scripts_dir = PROJECT_ROOT / "scripts"
    if str(scripts_dir) not in sys.path:
        sys.path.insert(0, str(scripts_dir))
    from build_web_ui_smoke_fixture import build_workbook

    build_workbook(path)


def _reset_workbook(fixture: Path, workbook: Path) -> None:
    _close_default_workbook_if_open(workbook)
    if workbook.exists():
        try:
            workbook.unlink()
        except PermissionError as exc:
            raise SmokeError(
                f"{workbook.name} is open or locked. Close it, or pass --workbook "
                "with a different path."
            ) from exc
    shutil.copy2(fixture, workbook)


def _close_default_workbook_if_open(workbook: Path) -> None:
    default_target = WORKBOOK_PATH.resolve()
    try:
        target = workbook.resolve()
    except FileNotFoundError:
        target = workbook.absolute()

    if target != default_target:
        return

    try:
        import win32com.client as win32

        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        return

    target_text = str(target).lower()
    for idx in range(excel.Workbooks.Count, 0, -1):
        wb = excel.Workbooks(idx)
        if str(wb.FullName).lower() == target_text:
            wb.Close(SaveChanges=False)
            return


def _healthcheck(url: str) -> None:
    health_url = f"{url.rstrip('/')}/_stcore/health"
    try:
        with urllib.request.urlopen(health_url, timeout=5) as response:
            body = response.read().decode("utf-8", errors="replace").strip()
    except Exception as exc:
        raise SmokeError(f"Streamlit is not reachable at {health_url}: {exc!r}") from exc
    if body != "ok":
        raise SmokeError(f"Streamlit healthcheck returned {body!r}, expected 'ok'.")


def _get_excel() -> Any:
    import win32com.client as win32

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    return excel


def _open_workbook(path: Path) -> Any:
    excel = _get_excel()
    target = str(path.resolve()).lower()
    for idx in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(idx)
        if str(wb.FullName).lower() == target:
            wb.Activate()
            wb.Worksheets("Report").Activate()
            return wb

    wb = excel.Workbooks.Open(str(path.resolve()))
    wb.Activate()
    wb.Worksheets("Report").Activate()
    return wb


def _range_values(wb: Any, sheet: str, addr: str) -> list[Any]:
    value = wb.Worksheets(sheet).Range(addr).Value
    if isinstance(value, tuple):
        flattened: list[Any] = []
        for row in value:
            if isinstance(row, tuple):
                flattened.extend(row)
            else:
                flattened.append(row)
        return flattened
    return [value]


def _cell_value(wb: Any, sheet: str, addr: str) -> Any:
    return wb.Worksheets(sheet).Range(addr).Value


def _verify(case: SmokeCase, wb: Any, assistant_text: str) -> tuple[bool, str]:
    if case.name == "fill_series_natural_phrase":
        values = _range_values(wb, "Report", "K1:K5")
        ok = values == [1, 2, 3, 4, 5]
        return ok, f"Report!K1:K5={values!r}"

    if case.name == "query_range_filter_sum":
        ok = "500" in assistant_text and ("North" in assistant_text or "north" in assistant_text)
        preview = assistant_text.replace("\n", " ")[:180]
        return ok, f"assistant contains North and 500; preview={preview!r}"

    if case.name == "beautify_range_business_blue":
        ws = wb.Worksheets("SalesData")
        header = ws.Range("A1")
        striped = ws.Range("A3")
        ok = (
            bool(header.Font.Bold)
            and header.Interior.ColorIndex != -4142
            and striped.Interior.ColorIndex != -4142
            and bool(ws.AutoFilterMode)
        )
        return (
            ok,
            "SalesData!A1 bold/fill applied, A3 stripe applied, AutoFilterMode="
            f"{bool(ws.AutoFilterMode)}",
        )

    if case.name == "plain_beautify_salesdata_report":
        ws = wb.Worksheets("SalesData")
        header = ws.Range("A1")
        first_data = ws.Range("A2")
        amount_cell = ws.Range("E2")
        ok = (
            ws.UsedRange.Address == "$A$1:$G$10"
            and header.Text == "Date"
            and str(first_data.Text).startswith("2026-01-01")
            and bool(header.Font.Bold)
            and header.Interior.ColorIndex != -4142
            and bool(ws.AutoFilterMode)
            and amount_cell.NumberFormat == "#,##0"
        )
        return (
            ok,
            "SalesData stayed A1:G10 with Date header, first data row intact, "
            f"AutoFilterMode={bool(ws.AutoFilterMode)}, Amount format={amount_cell.NumberFormat!r}",
        )

    if case.name == "plain_query_north_amount_sum":
        ok = "500" in assistant_text and ("North" in assistant_text or "north" in assistant_text)
        preview = assistant_text.replace("\n", " ")[:180]
        return ok, f"assistant contains North and 500; preview={preview!r}"

    if case.name == "plain_write_report_aa1":
        value = _cell_value(wb, "Report", "AA1")
        return value == "白話輸入OK", f"Report!AA1={value!r}"

    if case.name == "write_range_undo_marker":
        value = _cell_value(wb, "Report", "Z1")
        return value == "SMOKE_UNDO", f"Report!Z1={value!r}"

    if case.name == "undo_last_blank_restore":
        value = _cell_value(wb, "Report", "Z1")
        return value in (None, ""), f"Report!Z1={value!r}"

    raise SmokeError(f"No verifier for {case.name}")


def _import_playwright() -> tuple[Any, type[BaseException]]:
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeout
        from playwright.sync_api import sync_playwright
    except ModuleNotFoundError as exc:
        raise SmokeError(
            "Playwright is not installed. Install it only when you want to run "
            "browser automation: python -m pip install playwright; python -m "
            "playwright install chromium"
        ) from exc
    return sync_playwright, PlaywrightTimeout


def _send_prompt(page: Any, prompt: str, timeout_sec: int) -> str:
    messages = page.locator('[data-testid="stChatMessage"]')
    before = messages.count()

    input_box = page.locator('[data-testid="stChatInputTextArea"]')
    input_box.wait_for(state="visible", timeout=15_000)
    input_box.fill(prompt)
    input_box.press("Enter")

    deadline = time.monotonic() + timeout_sec
    last_text = ""
    while time.monotonic() < deadline:
        count = messages.count()
        if count:
            last_text = messages.nth(count - 1).inner_text(timeout=8_000)
        page_text = page.locator("body").inner_text(timeout=8_000)
        still_running = "Running..." in page_text or "⏳ 思考中" in page_text
        if count >= before + 2 and not still_running:
            return last_text
        time.sleep(1.0)

    raise SmokeError(f"Timed out waiting for assistant response. Last text: {last_text[:300]}")


def _run_browser_smoke(
    *,
    url: str,
    workbook: Path,
    cases: list[SmokeCase],
    timeout_sec: int,
    headless: bool,
) -> list[SmokeResult]:
    sync_playwright, playwright_timeout = _import_playwright()
    wb = _open_workbook(workbook)
    results: list[SmokeResult] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context(viewport={"width": 1440, "height": 950})
        page = context.new_page()
        page.goto(url, wait_until="networkidle")
        page.locator('[data-testid="stChatInputTextArea"]').wait_for(
            state="visible",
            timeout=20_000,
        )

        for case in cases:
            started = time.monotonic()
            try:
                assistant_text = _send_prompt(page, case.prompt, timeout_sec)
                ok, evidence = _verify(case, wb, assistant_text)
                status = "PASS" if ok else "FAIL"
                error = None
            except playwright_timeout as exc:
                assistant_text = ""
                status = "TIMEOUT"
                evidence = ""
                error = str(exc)
            except Exception as exc:
                assistant_text = ""
                status = "ERROR"
                evidence = ""
                error = repr(exc)

            results.append(
                SmokeResult(
                    name=case.name,
                    tool=case.tool,
                    status=status,
                    duration_ms=int((time.monotonic() - started) * 1000),
                    evidence=evidence,
                    assistant_preview=assistant_text[:500],
                    error=error,
                )
            )
            if status != "PASS":
                break

        browser.close()

    return results


def _write_log(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument("--fixture", type=Path, default=FIXTURE_PATH)
    parser.add_argument("--workbook", type=Path, default=WORKBOOK_PATH)
    parser.add_argument("--log", type=Path, default=LOG_PATH)
    parser.add_argument("--timeout-sec", type=int, default=WAIT_TIMEOUT_SEC)
    parser.add_argument("--headless", action="store_true")
    parser.add_argument(
        "--case-set",
        choices=sorted(CASE_SETS),
        default="tool",
        help="tool=explicit tool-name smoke, plain=natural user wording, all=both",
    )
    parser.add_argument(
        "--plain-language",
        action="store_true",
        help="Shortcut for --case-set plain.",
    )
    parser.add_argument(
        "--reuse-workbook",
        action="store_true",
        help="Do not copy the fixture over the workbook before running.",
    )
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help="Create/copy the workbook and run preflight checks, but do not open the browser.",
    )
    args = parser.parse_args()
    case_set_name = "plain" if args.plain_language else args.case_set
    cases = CASE_SETS[case_set_name]

    _ensure_fixture(args.fixture)
    if not args.reuse_workbook:
        _reset_workbook(args.fixture, args.workbook)

    _healthcheck(args.url)
    _open_workbook(args.workbook)

    if args.prepare_only:
        print(f"prepared {args.workbook}")
        return 0

    results = _run_browser_smoke(
        url=args.url,
        workbook=args.workbook,
        cases=cases,
        timeout_sec=args.timeout_sec,
        headless=args.headless,
    )
    summary: dict[str, int] = {}
    for result in results:
        summary[result.status] = summary.get(result.status, 0) + 1

    payload = {
        "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "url": args.url,
        "case_set": case_set_name,
        "fixture": str(args.fixture),
        "workbook": str(args.workbook),
        "summary": summary,
        "results": [asdict(result) for result in results],
    }
    _write_log(args.log, payload)

    print(json.dumps(payload["summary"], ensure_ascii=False))
    print(f"log written: {args.log}")
    return 0 if summary.get("FAIL", 0) == summary.get("ERROR", 0) == summary.get("TIMEOUT", 0) == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
