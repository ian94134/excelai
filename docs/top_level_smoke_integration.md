# Top-Level Smoke Test Integration

This `excel-ai` repository is nested under a top-level application folder:

`C:\Users\User\Desktop\excel-main\excel-main`

At the time this note was written, the top-level folder was not a Git
repository, while `excel-ai` was. The top-level smoke-test integration therefore
cannot be captured by the `excel-ai` commit directly.

## Top-Level Files To Preserve

- `package.json`
  - Adds `npm run test:tools-smoke`
  - Runs `cd excel-ai && cross-env EXCEL_AI_RUN_COM_SMOKE=1 pytest -q tests/test_tool_smoke.py`
- `.github/workflows/excel-com-smoke.yml`
  - Adds a manual `workflow_dispatch` job for Excel COM smoke testing
  - Keeps the COM smoke outside normal push / pull_request CI

## Restore / Apply

From the top-level folder, apply the companion patch:

```powershell
cd C:\Users\User\Desktop\excel-main\excel-main
git apply excel-ai\docs\top_level_smoke_integration.patch
```

If the top-level folder is still not a Git repository, use the patch as a
plain-text change record and manually confirm the two files above.

## Validation

```powershell
npm run test:tools-smoke
```

Expected local result on a Windows machine with desktop Excel available:

```text
1 passed
```
