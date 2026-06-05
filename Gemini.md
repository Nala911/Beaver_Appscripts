# 🦫 WorkspaceSync Appscripts - Blueprint & Architecture Guide

Read this before changing any `.js` or `.html` file in this repository. **This project is maintained 100% by AI agents.** Human coders do not actively edit this codebase. Therefore, all architectural rules, patterns, and conventions must be strictly preserved to maintain systematic scalability.

Consolidating both the architectural blueprint and developer rules here ensures that future agents automatically receive this full context within their system prompt rule block, avoiding unnecessary file reads and maintaining design consistency.

---

## 🚀 Key Commands & Workflow
This project uses **Clasp** (Command Line Apps Script Projects) for local development.

- **Pull Code:** `clasp pull`
- **Deploy/Push:** `clasp push`
- **Open Script Editor:** `clasp open`
- **Testing:** Since this is a Google Workspace add-on/script, testing is performed by running functions directly from the Apps Script editor or interacting with the "Workspace Sync Tools" menu in the associated Google Sheet.

---

## 🏛️ Ground Rules & Source Of Truth
- Treat `Gemini.md` as the ultimate architectural map and source of guidelines.
- Treat `00_Config_Constants.js` as the source of truth for global state keys and sheet names.
- **Do not invent new patterns.** If a tool needs local storage, declare it in `APP_PROPS` and access it through `_App_getProperty` / `_App_setProperty`. If a tool needs to modify the spreadsheet UI, it must use the `_App_` ecosystem.
- Before adding a new tool, use `CalendarSync_Code.js` and `CalendarSync_Sidebar.html` as your primary architectural benchmark models.

---

## ⛔ "Do Not Touch" Core Modules
The system is split into two halves: the Core Engine and the Tool Modules. Agents maintaining specific features should treat the **Core Engine** files as immutable unless explicitly tasked to refactor the framework itself.

**Core Engine Files (Do Not Modify for Feature Work):**
- `00_Config_Constants.js` through `09_Engine_UI.js`
- `UI.js`
- `SidebarShared.html`
- `Logger.js`

If you are just editing or adding a feature (like Mail Merge, etc.), stick entirely to your tool's `_Code.js` and `_Sidebar.html` files.

---

## 📂 File System Structure & Tool Mapping
The project consists of `.js` (Google Apps Script server-side code) and `.html` (Sidebar interfaces) pairs for each tool.

### Core System Files
The system logic is split into sequential modules evaluated in order:
- `00_Config_Constants.js`: Global registries, `SHEET_NAMES`, `APP_PROPS`, and enum structures.
- `01_Config_Theme.js`: Default theme definitions, colors, and `SHEET_THEME` proxy.
- `01_SheetManager.js`: Centralized data access object (DAO). Uses `SyncEngine` configurations to map sheet data to JavaScript objects and vice-versa.
- `02_Config_Storage.js`: Unified properties service wrappers (`_App_getProperty`, `_App_setProperty`, `_App_deleteProperty`, `_App_getRawProperty`).
- `03_Core_Utils.js`: Core utilities (`_App_throttle`, `_App_callWithBackoff`, `_App_setProgress`, `_App_getColumnLetter`, `_App_escapeRegExp`, `_App_getDriveAttachment`, `_App_mergeEmails`, etc.).
- `04_Core_Validators.js`: Validation helpers for types and constraints.
- `05_Core_State.js`: Global application state management.
- `06_Sheets_Helpers.js`: Low-level spreadsheet helpers (`_App_canScaffoldSheet`, `_App_assertActiveSheet`, `_App_validateActiveSheet`).
- `07_Sheets_Formatting.js`: UI/styling application to sheets (`_App_applyBodyFormatting`).
- `08_Engine_Core.js`: The `SyncEngine` plugin registration and retrieval system.
- `09_Engine_UI.js`: UI abstractions for opening sidebars/dialogs and scaffolding sheets (`_App_openSidebar`, `_App_launchTool`, `_App_ensureSheetExists`).
- `UI.js`: The central UI orchestrator. Responsible for creating the custom "Workspace Sync Tools" menu (`onOpen`), providing the global wrapper for the Settings sidebar, and connecting user actions to the tools.
- `SidebarShared.html`: Shared HTML, CSS, and JS runtime for sidebars. Owns the common loading/toast/tooltip shell, `SyncSidebar` action helpers, global button locking, and reusable layout/action primitives consumed by tool sidebars.
- `PipelineControl_Sidebar.html`: Standalone pipeline dashboard sidebar. The legacy `PipelineControl_CSS.html` / `PipelineControl_JS.html` partials have been retired and their logic now lives here.
- `Logger.js`: Silent execution boundary. Provides `Logger.run` and no-op logging methods; expected failures should return `_App_fail(...)`, while unexpected exceptions are rethrown to the caller. All console-based logging is strictly forbidden.
- `appsscript.json` / `.clasp.json`: Google Apps Script configuration and Clasp deployment environment details.

### Tool Modules & Connections
Each tool has a Backend file, a Frontend sidebar file, and a global Entry Function triggered by the custom menu in `UI.js`.

| Tool Name | Tool Key | Backend (`.js`) | Frontend (`.html`) | UI Menu Entry Function |
|---|---|---|---|---|
| **Google Calendar** | `CALENDAR_SYNC` | `CalendarSync_Code.js` | `CalendarSync_Sidebar.html` | `CalendarSync_openSidebar` |
| **Google Contacts** | `CONTACTS_SYNC` | `ContactsSync_Code.js` | `ContactsSync_Sidebar.html` | `ContactsSync_openSidebar` |
| **Mail Merge** | `MAIL_MERGE` | `MailMerge_Code.js` | `MailMerge_Sidebar.html` | `MailMerge_openSidebar` |
| **Mail Sender** | `MAIL_SENDER` | `MailSender_Code.js` | `MailSender_Sidebar.html` | `MailSender_openSidebar` |
| **Docs Merge** | `DOCS_MERGE` | `DocsMerge_Code.js` | `DocsMerge_Sidebar.html` | `DocsMerge_openSidebar` |
| **Google Forms** | `FORMS_SYNC` | `FormsSync_Code.js` | `FormsSync_Sidebar.html` | `FormsSync_openSidebar` |
| **Bulk Folder Creation** | `BULK_FOLDER` | `BulkFolderCreation_Code.js` | `BulkFolderCreation_Sidebar.html` | `BulkFolderCreation_openSidebar` |
| **Google Drive** | `DRIVE_SYNC` | `DriveFileDetails_Code.js` | `DriveFileDetails_Sidebar.html` | `DriveFileDetails_openSidebar` |
| **Pipeline** | `PIPELINE` | `PipelineControl_Code.js` | `PipelineControl_Sidebar.html` | `PipelineControl_openSidebar` |
| **Google Chat Spaces** | `CHAT_SYNC` | `ChatSpaceSync_Code.js` | `ChatSpaceSync_Sidebar.html` | `ChatSpaceSync_openSidebar` |
| **Gmail Filters** | `GMAIL_FILTERS` | `GmailFilters_Code.js` | `GmailFilters_Sidebar.html` | `GmailFilters_openSidebar` |
| **Google Tasks** | `TASKS_SYNC` | `TasksSync_Code.js` | `TasksSync_Sidebar.html` | `TasksSync_openSidebar` |

> [!NOTE]
> Every tool backend self-registers with `SyncEngine.registerTool('<KEY>', ...)` at the top of its file.

> [!CAUTION]
> **Large File Warning:** The following files are large (25KB+). Use surgical/partial reads.
> - `SidebarShared.html` (~45KB): Central CSS stylesheet, UI layout logic, and the global `SyncSidebar` wrapper.
> - `DriveFileDetails_Code.js` (~32KB): Complex Drive synchronization logic.
> - `PipelineControl_Sidebar.html` (~27KB): The standalone pipeline dashboard interface and controls.
> - `ContactsSync_Code.js` (~22KB): People API integration logic.

---

## ⚙️ Mandatory Code Contracts

### 1. Naming Standards
- **Tool Backend Logic:** `<ToolName>_Code.js` (e.g., `MailMerge_Code.js`). Always PascalCase for the tool name.
- **Tool Sidebar Interface:** `<ToolName>_Sidebar.html` (e.g., `MailMerge_Sidebar.html`).
- **Public Backend Functions:** `ToolName_FunctionName` (e.g., `MailMerge_openSidebar`). These are called from Sidebars or the UI menu.
- **Internal Helper Functions:** `_ToolName_InternalFunction` (e.g., `_MailMerge_validateData`).
- **Core System Utilities:** `_App_UtilityName` (e.g., `_App_launchTool`).
- **Sidebar Composition:** Prefer a single `ToolName_Sidebar.html` that includes `<?!= _App_include('SidebarShared'); ?>`. Legacy tool-specific HTML partials such as `ToolName_CSS.html` / `ToolName_JS.html` are no longer the default pattern.

### 1a. UI Naming Standards (Uniformity)
To maintain a professional and consistent user experience, the following strings MUST match exactly:
- **`TITLE`** (in `SyncEngine.registerTool`): Must match the tool's `SHEET_NAME` value exactly (including emoji).
- **`MENU_LABEL`** (in `SyncEngine.registerTool`): Must match the tool's `SHEET_NAME` value exactly (including emoji).
- **`Gemini.md` (Bold Tool Name)**: Must match the tool's `SHEET_NAME` value but without the emoji (in the Tool Modules & Connections table).
- **Sidebar Header**: Must use the `<div class="header">` structure with the `<i data-lucide="...">` explicitly placed *inside* the `<div class="header-title">` container to ensure uniform alignment. The text must match the base tool name (without emoji or suffixes).
- **Status Column**: Every tool sheet MUST include a `Status` column immediately following the `Action` column. It must be defined in `COL_SCHEMA` as `{ header: 'Status', type: 'STATUS' }`.
- **Column Categories**: Formatting is strictly schema-driven and positional. The engine maps types to three visual categories:
    - **First Columns (Action/Status)**: Includes `type: 'ACTION'` and `type: 'STATUS'`. Colors use `SHEET_THEME.FIRST_COLS_COLOR`.
    - **Last Columns (Read-Only/IDs)**: Includes `type: 'READ_ONLY'` and `type: 'ID'`. **Crucial**: All System IDs (e.g., `type: 'ID'`) MUST be placed at the very end of the `COL_SCHEMA` array to hide non-actionable technical data from the user's immediate view. Colors use `SHEET_THEME.LAST_COLS_COLOR`.
    - **Middle Columns (Editable)**: Includes all other data input types (`TEXT`, `URL`, `DROPDOWN`, `CHECKBOX`, `EMAIL`, etc.). Colors use `SHEET_THEME.MIDDLE_COLS_COLOR`.
- **Frozen Columns**: To ensure these system columns remain visible at all times, the engine enforces a default of 2 frozen columns. Tools can omit `FROZEN_COLS` in registration metadata to let the engine apply this default, but if specified, it must be set to `2`.
- **Sidebar Documentation**: Only tool sidebars that benefit from guided onboarding need a "Help & Guide" section at the bottom, using the standardized `.sync-sidebar-help-guide-card` architecture.

### 1b. Sidebar Help & Documentation (Uniformity)
To ensure user clarity, sidebars should include a "Help & Guide" section using one of these two standardized patterns:
1. **Dynamic Config-Driven Help (Preferred)**: Define a `HELP_ITEMS` object inside the tool's backend `registerTool` config containing `gettingStarted` HTML content and list of items/tooltips. In the sidebar HTML file, place `<div id="sync-sidebar-help-container"></div>` at the bottom, and initialize via `SyncSidebar.initSidebar({ toolKey: 'YOUR_TOOL_KEY' })` to automatically fetch and render the help section.
2. **Manual HTML Guide (Fallback)**: Directly write the guide layout at the bottom of the HTML sidebar file using the `.sync-sidebar-help-guide-card` container and `.sync-sidebar-help-guide-item` rows from `SidebarShared.html`, then initialize via `SyncSidebar.initSidebar({ createIcons: true })`.
- **Content Guidelines**: Keep it focused; document only the core columns, a 3-step quick start, and critical performance or behavioral "gotchas".
- **Standard Tooltip System**:
  - `help-trigger`: A CSS class applied to icons (usually `help-circle`).
  - `data-help-target`: An attribute on the trigger pointing to the `ID` of a hidden content element.
  - Hidden Content Container: A `div` at the bottom of the HTML file (set to `display: none`) containing multiple divs with specific IDs (e.g., `help-getting-started`).
  - Global event handlers in `SidebarShared.html` handle tooltip positioning and boundary overflows.

### 2. The `SyncEngine` Contract & Plugin Architecture
- **Registration**: Every tool backend file must register itself with the engine at the very top of the script using `SyncEngine.registerTool(key, config)`. Do not hardcode columns inside backend logic; rely on the registry's `FORMAT_CONFIG.COL_SCHEMA`.
- **Default Config Inference**: To simplify registrations, `FROZEN_ROWS` (defaults to 1), `FROZEN_COLS` (defaults to 2), `COL_WIDTHS`, and `conditionalRules` are optional. If `COL_WIDTHS` is omitted, the engine automatically infers standard widths. If `conditionalRules` is omitted, the engine automatically injects standard highlighting rules for pending actions (amber), successes (green), warnings (orange), and errors (red) by scanning the `COL_SCHEMA` for ACTION and STATUS column types.

### 3. Frontend Unified Wrapper (`SyncSidebar` / Frontend-Backend Connection)
All client-to-server communication MUST use the `SyncSidebar` layer from `SidebarShared.html`. `SyncSidebar.run()` unwraps the standard `{ success, message, data, meta }` payloads and provides consistent toast notifications.
- **Connection Flow**:
  1. **Trigger**: User clicks menu or sidebar button.
  2. **Launch**: `_App_openSidebar('TOOL_KEY')` handles sheet preparation and sidebar rendering.
  3. **Execution**: Sidebar calls `SyncSidebar.run('ToolName_publicFunc')` -> Backend function -> `Logger.run()` for a consistent execution boundary.
  4. **Response**: Backend returns standardized `_App_ok(...)` / `_App_fail(...)` payloads.
- **Automatic Locking**: This wrapper automatically locks all sidebar buttons and applies a "grayed-out" style during the call.
- **Overlapping Calls**: The engine uses a counter; buttons stay locked until *all* concurrent `SyncSidebar.run` calls complete.
- **Opting Out**: For silent background tasks (like progress polling), use `SyncSidebar.run(method, args, { lockButtons: false })`.
- **Redundancy**: DO NOT manually disable buttons in sidebar code (e.g., `btn.disabled = true`); rely entirely on the core wrapper to maintain UI state.
- **Preferred Helpers**: Default to `SyncSidebar.initSidebar()`, `SyncSidebar.runPullAction()`, `SyncSidebar.runPushAction()`, and `SyncSidebar.runAction()` instead of rebuilding the same orchestration in each sidebar.
- **Safety Prompts for Pull Actions**: For tools that import external data and overwrite spreadsheet rows, always provide the centralized `'UI_checkForUnsavedChanges'` helper as `unsavedCheckMethod` inside `SyncSidebar.runPullAction` and pass the tool key (e.g. `['YOUR_TOOL_KEY']`) inside `unsavedCheckArgs`. The centralized check returns a boolean `{hasChanges: true|false}` indicating if there are unsaved edits in the `Action` column, presenting a standard warning dialog before data is overwritten.
- **Styling Boundary**: Standard sync sidebars should reuse shared shell tokens and layout primitives (e.g., `btn-pull`, `btn-push`, `sync-sidebar-action-grid`, `sync-sidebar-action-stack`, `sync-sidebar-inline-options`, etc.) from `SidebarShared.html`. Only specialized dashboards with materially different layouts, such as `PipelineControl_Sidebar.html`, should keep larger local style blocks.
- **Icons**: All sidebars must use the Lucide icon framework exclusively (`<i data-lucide="..."></i>`).
- **Dynamic Icons & Hydration**: If the sidebar dynamically updates the DOM or injects dynamic HTML content containing `<i data-lucide="...">` tags, you MUST call `SyncSidebar.refreshIcons()` immediately after the DOM update to ensure the Lucide engine parses and renders the new icons.

### 4. Logger.run & Unified Reporting Contract
- **Silent Backend**: Every public function called from a sidebar or the Sheets menu must use the `Logger.run` execution wrapper to preserve a consistent execution boundary. **Direct use of `console.log`, `console.warn`, or `console.error` is strictly prohibited.** Expected validation failures should return `_App_fail(...)`; unexpected system-level errors should be thrown so `SyncSidebar` / Apps Script failure handlers can surface them.
  ```javascript
  function MyTool_publicFunction() {
      return Logger.run('MY_TOOL', 'Action Context', function() {
          // ... logic
          return _App_ok('Done');
      });
  }
  ```
- **Unified Reporting Architecture**:
  1. **Row-Level Errors**: Processing errors must be caught and returned as `{ isError: true, error: msg }` to `_App_BatchProcessor`. The `onBatchComplete` hook MUST then write these to the `Status` column prefixed with `SHEET_THEME.STATUS_PREFIXES.ERROR` (❌).
  2. **Row-Level Success**: All row-level success messages MUST be prefixed with `SHEET_THEME.STATUS_PREFIXES.SUCCESS` (✅).
  3. **Row-Level Warnings**: Non-blocking issues should use `SHEET_THEME.STATUS_PREFIXES.WARNING` (⚠️).
  4. **General/System Errors**: Any system-wide errors in sidebars must use `SyncSidebar.handleError(err, { severity: 'mild|medium|critical' })` to surface a standardized modal box:
     - `mild`: Blue "Notice" modal for informational non-errors.
     - `medium`: Amber "Warning" modal for recoverable issues.
     - `critical`: Red "System Error" modal for blocking/critical failures (Default).
- **API Error Translation**: The engine provides `_App_translateApiError(err)` inside `03_Core_Utils.js` which automatically intercepts Google JSON exception strings (like 403 authorization failures or 429 rate limit errors) and translates them into actionable, human-friendly troubleshooting guides displayed in sidebars.

### 5. The PropertiesService Contract
Never use `PropertiesService.getDocumentProperties()` directly in a tool. 
- Define your new key strictly in `APP_PROPS` inside `00_Config_Constants.js`.
- Use `_App_getProperty` and `_App_setProperty` from `02_Config_Storage.js`.

### 6. Batch Processing, Execution Time, & Trigger Management
- **Centralized Batch Processor**: Row-by-row data processing must use `_App_BatchProcessor` from `03_Core_Utils.js`. This utility handles progress tracking, backoff retries, and protects against the 6-minute script timeout.
  - **Error Propagation**: The `processFn` should throw errors directly.
  - **Status Reporting**: Use `SheetManager.batchPatchRows` within the `onBatchComplete` hook to write results into the `Status` column.
  - **Automatic Retries**: Wraps each item in `_App_callWithBackoff` to handle transient API errors.
  - **Progress Tracking**: Automatically updates CacheService with progress data for sidebar polling.
- **Execution Time Limits**: Global timing is managed via `_App_resetExecutionTimer()` and `_App_isExecutionLimitApproaching()`. The processor pauses execution at 5.5 minutes, allowing for safe partial completions and saving the progress.
- **Trigger Management**: Background sync tools should manage their own `ScriptApp` triggers. Use an internal `_ToolName_manageTrigger` function called from the setting update handler to ensure triggers are created/removed in sync with user preferences.
- **Intelligent Halting**: If the `_App_BatchProcessor` intercepts a fatal system or authorization error (e.g. rate limit exceeded or access token revoked), it flushes all currently successful segment row status modifications and halts execution immediately. This prevents cascading raw API errors and preserves the Action column state for the remaining unprocessed rows.

### 7. Centralized Validation Contract
- **Deprecating Local Helpers**: Do not write custom local validators inside tool modules for common validation needs (like email syntax verification).
- **Core Validators**: Always call `_App_validateEmail` or `_App_validateEmailList` from `04_Core_Validators.js`. Any future general validators should be added to that central module rather than being duplicated in tool backends.
- **Schema-Driven Pre-Validation**: The engine provides a unified data pre-validation engine. The `_App_BatchProcessor` dynamically invokes `_App_validateRowAgainstSchema(item, toolKey)` based on the tool's registered `COL_SCHEMA` rules before executing the tool's backend logic. Schema types like `EMAIL`, `DATE`, `DATETIME`, `BOOLEAN`, `URL`, `DOCS_URL`, `DRIVE_URL`, and `DROPDOWN` are validated automatically using `_App_validateValueByType()` in `04_Core_Validators.js`.

### 8. Batch Result Patching & Action Preservation
- **Reporting Success/Failure**: When batch execution completes, do not write raw status strings to the sheet manually. Call `_App_batchPatchResults` (defined in `03_Core_Utils.js`) within the `onBatchComplete` hook of your `_App_BatchProcessor`.
- **Action Column Rule**: If a row-level execution fails (e.g., invalid email address), the patch MUST preserve the user's `Action` column value. This ensures users can fix typos in the input columns and re-run without having to re-type the sync actions/verbs. `_App_batchPatchResults` handles this automatically by only resetting the `Action` column on success.

### 9. Concurrency Protection & Document Locks
- **Wrapper Enforce**: All high-impact backend entry points (such as push/pull workflows) MUST be wrapped in the unified concurrency shield `_App_withDocumentLock('<TOOL_KEY>_<ACTION>', function() { ... })`.
- **Block & Notify**: The lock service retrieves a document-level lock. If it fails, it will raise a structured error that will notify the user with a standardized notice modal rather than silent failure or race condition state contamination.

### 10. Timezone Standardization
- **Use Unified Formatter**: Never hardcode formatting offsets or use `Session.getScriptTimeZone()` directly for writing date strings to sheet displays.
- **Utility**: Always use `_App_formatDateTime(date, format)` from `03_Core_Utils.js`. It automatically resolves the active spreadsheet timezone and handles offset conversions correctly.

### 11. State Isolation
- **Spreadsheet-Scoped Cache**: Storing global progress state under raw keys leads to collisions when multiple spreadsheets run syncs concurrently.
- **Helper**: Use `_App_setProgress`, `_App_getProgress`, and `_App_clearProgress` in `05_Core_State.js`, which automatically utilize the internal `_App_getProgressKey_` helper to namespace cache keys with the active Spreadsheet ID.

### 12. Formatting Optimization
- **Rule Bypassing**: Re-evaluating and applying conditional formatting rules on every batch write is highly expensive.
- **Bypass Flag**: Ensure conditional formatting rule application is decoupled from regular body formatting writes. The engine should only overwrite conditional formatting rules when explicitly scaffolding or structurally sync-modifying the sheets.

---

## 🔑 Google API Scopes & Services Used
Each tool relies on specific Google APIs. Do NOT use an API in a tool that doesn't need it.

| Tool | Google APIs / Services | Advanced Service? |
|---|---|---|
| **Google Calendar** | `CalendarApp`, `Calendar` (Advanced) | Yes — `Calendar API v3` |
| **Google Contacts** | `People` (Advanced) | Yes — `People API v1` |
| **Mail Merge** | `GmailApp`, `DocumentApp` | No |
| **Mail Sender** | `GmailApp`, `MailApp` | No |
| **Docs Merge** | `DocumentApp`, `DriveApp` | No |
| **Google Forms** | `FormApp`, `DriveApp` | No |
| **Bulk Folder Creation** | `DriveApp` | No |
| **Google Drive** | `DriveApp`, `Drive` (Advanced) | Yes — `Drive API v3` |
| **Google Chat Spaces** | `Chat` (Advanced) | Yes — `Chat API v1` |
| **Gmail Filters** | `Gmail` (Advanced) | Yes — `Gmail API v1` |
| **Google Tasks** | `Tasks` (Advanced) | Yes — `Tasks API v1` |
| **Pipeline** | `PropertiesService`, `SpreadsheetApp`, `ScriptApp` | No |

---

## 🗝️ PropertiesService Key Registry
All keys used across the codebase. **Do NOT invent new key names** — check here first and follow the naming convention in `APP_PROPS`.

| Key | File | Store Type | Purpose |
|---|---|---|---|
| `SYSTEM_ENABLED` | `PipelineControl_Code.js` | `ScriptProperties` | Master on/off toggle for pipeline |
| `DOCS_MERGE_TEMPLATE_URL` | `DocsMerge_Code.js` | `DocumentProperties` | Saved template Doc URL |
| `DOCS_MERGE_FOLDER_URL` | `DocsMerge_Code.js` | `DocumentProperties` | Saved output folder URL |
| `DOCS_MERGE_TEMPLATE_NAME` | `DocsMerge_Code.js` | `DocumentProperties` | Cached template file name |
| `DOCS_MERGE_FOLDER_NAME` | `DocsMerge_Code.js` | `DocumentProperties` | Cached folder name |
| `DOCS_MERGE_MASTER_DOC_ID` | `DocsMerge_Code.js` | `DocumentProperties` | Cached master document ID used to resume a multi-batch merge |
| `selectedCalIds` | `CalendarSync_Code.js` | `UserProperties` | JSON array of selected calendar IDs |
| `startDate` | `CalendarSync_Code.js` | `UserProperties` | Saved start date filter |
| `endDate` | `CalendarSync_Code.js` | `UserProperties` | Saved end date filter |
| `selectedContactGroups` | `ContactsSync_Code.js` | `UserProperties` | JSON array of selected contact group IDs |
| `selectedChatSpaces` | `ChatSpaceSync_Code.js` | `UserProperties` | JSON array of selected chat space IDs |
| `FORMSSYNC_CURRENT_FORM` | `FormsSync_Code.js` | `DocumentProperties` | Stores currently synced form ID |
| `FORMSSYNC_SELECTED_FORM` | `FormsSync_Code.js` | `UserProperties` | Stores user's selected form ID for sidebar auto-selection |
| `selectedTasksList` | `TasksSync_Code.js` | `UserProperties` | (Registered but currently unused/reserved; TasksSync pulls all lists) |

---

## 🤖 Gemini Workflow Rules
1. **Minimize file reads**: ONLY read the specific tool files needed.
2. **Consult Core Modules first**: Global configuration and logic are defined in `00_Config_Constants.js` through `09_Engine_UI.js`.
3. **Use `Logger.run()`**: Wrap primary tool operations in `Logger.run('TOOL_KEY', 'Context', () => { ... })` for consistent error boundary management. **NEVER use `console.log` for debugging or reporting.**
4. **Follow `SyncEngine`**: When modifying sheet structure, update the registration metadata in the tool's backend file, which registers with `SyncEngine`.

---

## 📋 Gemini Pre-Flight Checklist
Before completing any task, mentally run this checklist. Do not proceed until you have verified all points:
- [ ] Does my backend file register with `SyncEngine` at the very top?
- [ ] Are all public functions prefixed with `ToolName_` (e.g., `MailMerge_doWork`)?
- [ ] Are all internal helper functions prefixed with `_ToolName_` (e.g., `_MailMerge_validate`)?
- [ ] Did I wrap my core action inside `Logger.run('KEY', 'Context', function() {...})`?
- [ ] Does my backend file include the mandatory `Status` column in `COL_SCHEMA`?
- [ ] Does my `onBatchComplete` logic handle `res.isError` to report failures in the `Status` column?
- [ ] Does my public function return a standard object via `_App_ok` or `_App_fail?`
- [ ] Did I use `_App_callWithBackoff` around any external Google API calls?
- [ ] If I added a new setting, is it declared in `APP_PROPS` in `00_Config_Constants.js`?
- [ ] If my tool processes rows, did I use `_App_BatchProcessor` and `_App_batchPatchResults` to apply patches and preserve Action columns on failure?
- [ ] Did I use centralized validation functions (e.g., `_App_validateEmailList` in `04_Core_Validators.js`) instead of writing local duplicate validation helpers?
- [ ] Did I register my tool's schema in `FORMAT_CONFIG.COL_SCHEMA` correctly so the core validation engine (`_App_validateRowAgainstSchema`) validates row cell types automatically before sync begins?
- [ ] Did I ensure that any custom backend exceptions thrown return descriptive messages, allowing the translation engine `_App_translateApiError` to map errors correctly?
- [ ] If my sidebar dynamically modifies the DOM to add elements with icons, did I call `SyncSidebar.refreshIcons()` afterward?
- [ ] Is my sidebar strictly including `<?!= _App_include('SidebarShared'); ?>` to inherit standard WorkspaceSync UI libraries?
- [ ] Are all backend calls in the sidebar routed through `SyncSidebar.run()` instead of raw `google.script.run`?
- [ ] Did I use `lockButtons: false` for background polling tasks to avoid UI jitter?
- [ ] Did I remove all manual `btn.disabled = true` logic, relying on the `SyncSidebar` core instead?
- [ ] If my tool overwrites spreadsheet rows during a Pull action, did I bind the centralized `'UI_checkForUnsavedChanges'` helper as `unsavedCheckMethod` in the sidebar `runPullAction` call and pass the tool key in `unsavedCheckArgs`?
- [ ] Did I wrap heavy backend push and pull entry points with `_App_withDocumentLock` to ensure concurrency protection?
- [ ] Did I use the central `_App_formatDateTime` utility for timezone-safe date-time formatting to the spreadsheet?

---

## 🚀 Adding a New Tool
Since humans do not code here, follow the **CalendarSync Benchmark**:
1. Duplicate `CalendarSync_Code.js` and rename it to `<NewName>_Code.js`.
2. Duplicate `CalendarSync_Sidebar.html` and rename it to `<NewName>_Sidebar.html`.
3. Add the `SHEET_NAMES` entry to `00_Config_Constants.js`.
4. Update the plugin registration inside `<NewName>_Code.js` (Key, Sheet Name, Title, etc.).
5. Ensure `COL_SCHEMA` includes the mandatory `Status` column (type: `STATUS`) as the second entry. Do NOT use legacy positional properties like `numReadOnlyColsAtEnd`.
6. Leverage Config Inference: You can omit `FROZEN_ROWS`, `FROZEN_COLS`, `COL_WIDTHS`, and `conditionalRules` from your config; the engine will automatically infer standard defaults and auto-inject standard conditional formatting highlights based on the `COL_SCHEMA` column types.
7. Implement backend logic with `Logger.run` and `_App_ok`.
8. Implement frontend logic using `SidebarShared.html` plus the `SyncSidebar` helpers (`runPullAction`, `runPushAction`, `runAction`) and native CSS variables (NO TailwindCSS).
