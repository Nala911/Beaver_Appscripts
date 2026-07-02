# Workspace Sync Appscripts - Agent Rules

## Purpose
This project implements local-first Google Apps Script tools under the "Beaver" system. It connects Google Sheets with Workspace services (Calendar, Contacts, Drive, Tasks, Gmail, Forms, etc.) via custom HTML sidebars and backend scripts.

## Project Structure
- `core/`: Core engine files. Filenames are numbered sequentially (e.g., `00_Config_Constants.js`, `01_SheetManager.js`) to enforce load/evaluation order in Apps Script.
- `tools/`: Modular tool implementations, each containing paired code and view files:
  - `*_Code.js`: Apps Script logic.
  - `*_Sidebar.html`: Embedded HTML/CSS/JS sidebar UI.
- `tests/`: Jest tests, mock environments, and custom test reporters.
- `validate-and-push.js`: Pipeline entry script that runs Jest tests, creates failure diagnostics, and executes `clasp push` on success.

## Key Rules & Guidelines

1. **Evaluation Order**: 
   - Files in `core/` must retain their sequential numbered prefixes. Adding files to `core/` requires matching the prefix scheme to control evaluation order in the Google Apps Script global scope.
2. **Global Namespace Safety**:
   - Google Apps Script executes all script files in a shared global namespace. Use IIFEs or namespace object wrappers (e.g., `SyncEngine`) to prevent variable collisions.
3. **Deployment Safety**:
   - Never run `npx clasp push` directly. Always run validation using `npm run deploy` (which runs `node validate-and-push.js`) to ensure all tests pass and a failure report is generated if they fail.
4. **HTML Sidebars**:
   - Sidebars are embedded inside the sheet UI and should share UI styling guidelines from the core configuration where applicable.
