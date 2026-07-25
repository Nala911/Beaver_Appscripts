# Workspace Sync Appscripts - Agent Rules

## Purpose
This project implements local-first Google Apps Script tools under the "Beaver" system. It connects Google Sheets with Workspace services (Calendar, Contacts, Drive, Tasks, Gmail, Forms, etc.) via custom HTML sidebars and backend scripts.

## Project Structure
- `core/`: Core engine files. Filenames are numbered sequentially (e.g., `00_Config_Constants.js`, `01_SheetManager.js`) to enforce load/evaluation order in Apps Script.
- `tools/`: Modular tool implementations, each containing paired code and view files:
  - `*_Code.js`: Apps Script logic.
  - `*_Sidebar.html`: Embedded HTML/CSS/JS sidebar UI.
- `tests/`: Jest tests, mock environments, and custom test reporters.
- `scripts/agent-cli.js`: Unified CLI harness for AI agents (`npm run agent:*`).
- `validate-and-push.js`: Pipeline entry script that runs Jest tests, creates failure diagnostics, and executes `clasp push` on success.
- `ENTRY_POINTS.md`: Comprehensive reference sitemap for AI agents.

## Key Rules & Guidelines

1. **Evaluation Order**: 
   - Files in `core/` must retain their sequential numbered prefixes. Adding files to `core/` requires matching the prefix scheme to control evaluation order in the Google Apps Script global scope.
2. **Global Namespace Safety**:
   - Google Apps Script executes all script files in a shared global namespace. Use IIFEs or namespace object wrappers (e.g., `SyncEngine`) to prevent variable collisions.
3. **Deployment Safety & Agent Entry Points**:
   - Never run `npx clasp push` directly. Always run validation using `npm run agent:deploy` or `npm run deploy` (which runs `node scripts/agent-cli.js deploy` / `node validate-and-push.js`) to ensure syntax checks and tests pass.
4. **Low CPU Governor**:
   - Always run tests using the agent harness (`npm run agent:test`) which enforces `--maxWorkers=50%` to keep CPU load minimal.
5. **Diagnostic Failure Reports**:
   - When tests fail, read [`tests/reporters/agent_failure_report.json`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_failure_report.json) or [`agent_failure_report.md`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_failure_report.md) for code line links, excerpts, missing mock code suggestions, and API call traces.
6. **Tool Scaffolding**:
   - Create new tools using `npm run agent:scaffold -- --name=ToolName`.
7. **HTML Sidebars**:
   - Sidebars are embedded inside the sheet UI and should share UI styling guidelines from the core configuration where applicable.
