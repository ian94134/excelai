# Top-Level Git Ownership Decision

Status: conservative checkout decision recorded on 2026-04-27.

## Current Layout

- Top-level application folder: `C:\Users\User\Desktop\excel-main\excel-main`
  - Not a Git repository at inspection time.
  - Contains the React/Electron frontend, top-level Node scripts, and top-level GitHub workflow files.
- Nested package folder: `C:\Users\User\Desktop\excel-main\excel-main\excel-ai`
  - Git repository on `main`, tracking `origin/main`.
  - Remote: `https://github.com/ian94134/excelai.git`
  - Source of truth for the Python, Streamlit, Excel COM, tool executor, macro, and smoke-test code.
- Top-level smoke integration:
  - `package.json` adds `npm run test:tools-smoke`.
  - `.github/workflows/excel-com-smoke.yml` adds manual Excel COM smoke execution.
  - The restore record is preserved in `docs/top_level_smoke_integration.patch`.

## Decision

Keep `excel-ai` as an independent Git repository. Treat the top-level application folder as a separate host/consumer until an outer repository and remote ownership model are explicitly established.

Do not convert the checkout to a monorepo or submodule in this pass.

## Rationale

- Monorepo conversion would require absorbing or rewriting the nested `excel-ai` repository boundary, which risks losing the already-verified remote/history relationship.
- Submodule conversion requires an outer Git repository, a remote for that outer repository, and a deliberate submodule policy. None of those are established in this checkout.
- Keeping the top-level app separate while preserving `excel-ai` as its own repository is the least disruptive path and preserves the validated commits already pushed to GitHub.

## Guardrails

- Run `git status`, commits, and pushes inside `excel-ai` for Python/Streamlit/Excel automation work.
- If the top-level folder later becomes its own frontend Git repository, keep `excel-ai/` ignored unless a deliberate submodule conversion is chosen.
- Do not run `git add .` from the top-level folder until `git status` confirms `excel-ai/` is ignored or intentionally registered as a submodule.
- Keep the top-level smoke integration patch in `docs/top_level_smoke_integration.patch` as the recovery record for files outside the `excel-ai` repository.
- Do not delete historical smoke artifacts such as `C:\Users\User\Desktop\excel-main\outputs\tool_smoke`.

## Applied Local Guardrail

This checkout's top-level `.gitignore` includes:

```gitignore
excel-ai/
```

The companion replay patch is `docs/top_level_git_ownership.patch`.
