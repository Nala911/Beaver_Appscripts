# 🦫 WorkspaceSync Appscripts - BLUEPRINT & Architecture Context

## Overview
This file serves as the absolute Architectural Map for AI agents working on the "WorkspaceSync Appscripts" workspace. It contains details on file connections, global state management, and module structures.

> [!NOTE]
> Workflow rules and procedural instructions for agents are strictly located in `Gemini.md`. Refer to that file before making any changes.

## 📂 File System Structure

The project consists of `.js` (Google Apps Script server-side code) and `.html` (Sidebar interfaces) pairs for each tool.

### Core System Files
The system logic is split into sequential modules evaluated in order:
- `00_Config_Constants.js`: Global registries, `SHEET_NAMES`, `APP_PROPS`, and enum structures.
- `01_Config_Theme.js`: Default theme definitions, colors, and `SHEET_THEME` proxy.
- `01_SheetManager.js`: Centralized data access object (DAO). Uses `SyncEngine` configurations to map sheet data to JavaScript objects and vice-versa.
- `02_Config_Storage.js`: Unified properties service wrappers (`_App_getProperty`, `_App_setProperty`, `_App_deleteProperty`, `_App_getRawProperty`).
- `03_Core_Utils.js`: Core utilities (`_App_throttle`, `_App_callWithBackoff`, `_App_setProgress`, etc.).
- `04_Core_Validators.js`: Validation helpers for types and constraints.
- `05_Core_State.js`: Global application state management.
- `06_Sheets_Helpers.js`: Low-level spreadsheet helpers (`_App_canScaffoldSheet`, `_App_assertActiveSheet`, `_App_validateActiveSheet`).
- `07_Sheets_Formatting.js`: UI/styling application to sheets (`_App_applyBodyFormatting`).
- `08_Engine_Core.js`: The `SyncEngine` plugin registration and retrieval system.
- `09_Engine_UI.js`: UI abstractions for opening sidebars/dialogs and scaffolding sheets (`_App_openSidebar`, `_App_launchTool`, `_App_ensureSheetExists`).
- `UI.js`: The central UI orchestrator. Responsible for creating the custom "Workspace Sync Tools" menu (`onOpen`), providing the global wrapper for the Settings sidebar, and connecting user actions to the tools.
- `SidebarShared.html`: Shared HTML, CSS, and JS runtime for sidebars. Owns the common loading/toast/tooltip shell, `SyncSidebar` action helpers, global button locking, and reusable layout/action primitives consumed by tool sidebars.
- `Settings_Sidebar.html`: Standalone settings dashboard sidebar. The legacy `Settings_CSS.html` / `Settings_JS.html` partials have been retired and their logic now lives here.
- `PipelineControl_Sidebar.html`: Standalone pipeline dashboard sidebar. The legacy `PipelineControl_CSS.html` / `PipelineControl_JS.html` partials have been retired and their logic now lives here.
- `Logger.js`: Silent execution boundary. Provides `Logger.run` and no-op logging methods; expected failures should return `_App_fail(...)`, while unexpected exceptions are rethrown to the caller. All console-based logging is strictly forbidden.
- `SystemAudit.js`: Runs comprehensive diagnostic audits across all registered tools, verifying sheet integrity, API access, and schema setup. Outputs results directly to the Settings UI.
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
| **Settings** | N/A | (Inside `UI.js`) | `Settings_Sidebar.html` | `UI_openSettingsDialog` |

> [!NOTE]
> Every tool backend self-registers with `SyncEngine.registerTool('<KEY>', ...)` at the top of its file. `Settings` is owned directly by `UI.js` and is not a registered tool.

> [!CAUTION]
> **Large File Warning:** The following files are large (25KB+). Use surgical reads.
> - `DriveFileDetails_Code.js` (~32KB): Complex Drive synchronization logic.
> - `ContactsSync_Code.js` (~27KB): People API integration logic.

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
| **Pipeline** | `PropertiesService`, `SpreadsheetApp`, `ScriptApp` | No |

| **Settings** | `PropertiesService` only | No |

## 🗝️ PropertiesService Key Registry

All keys used across the codebase. **Do NOT invent new key names** — check here first and follow the naming convention in `APP_PROPS`.

| Key | File | Store Type | Purpose |
|---|---|---|---|
| `WorkspaceSync_SHEET_THEME` | `01_Config_Theme.js` | `DocumentProperties` | Custom theme JSON overrides |
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
| `FORMSSYNC_SELECTED_FORM` | `FormsSync_Code.js` | `UserProperties` | Stores user's selected form ID for sidebar auto-selection; accessed through `APP_PROPS.FORMS_SELECTED_FORM` |


## 🏗️ Architectural Patterns

The codebase follows a strict and predictable design pattern across all tools. **`CalendarSync_Code.js` and `CalendarSync_Sidebar.html` serve as the benchmark models for all new implementations.**

### 1. Decentralized Plugin Architecture (`SyncEngine`)
The project uses a decentralized registration pattern to manage tools.
- **`SyncEngine`**: A singleton in `08_Engine_Core.js` that handles tool registration (`registerTool`) and retrieval (`getTool`).
- **Self-Registration**: Each tool module registers its own configuration block at the top of its file.
- **Registry Metadata**: Configuration includes `REQUIRED_SERVICES`, `SHEET_NAME`, `TITLE`, `MENU_LABEL`, `MENU_ENTRYPOINT`, `MENU_ORDER`, `SIDEBAR_HTML`, `SIDEBAR_WIDTH`, `FROZEN_ROWS`, `FROZEN_COLS`, `COL_WIDTHS`, and a `FORMAT_CONFIG` object (containing `conditionalRules` and `COL_SCHEMA` for declarative column validations, types, and schema-driven background categorization).

### 2. Unified Utilities (`_App_`)
Core logic is abstracted into `_App_` prefixed functions spread across `03_Core_Utils.js` to `09_Engine_UI.js`.

### 3. Flat Function Prefix Architecture
To avoid naming collisions and facilitate AI interaction, functions use `ToolName_` prefixes.

### 4. Global Entry Points
Entry functions for the UI menu generally look like:
```javascript
function ToolName_openSidebar() {
  _App_launchTool('TOOL_KEY');
}
```

### 5. Background Automation & Trigger Management
Tools that require background execution should manage their own triggers programmatically.

### 6. Frontend Unified Wrapper (`SyncSidebar`)
All client-to-server communication must use the `SyncSidebar` wrapper located in `SidebarShared.html`. `SyncSidebar.run()` abstracts `google.script.run`, standardizes loading states, unwraps the `{ success, message, data, meta }` payloads, and provides consistent toast notifications.

**Global Button Locking**: To prevent double-clicks and provide visual feedback, `SyncSidebar` automatically disables all buttons (`.btn` and `button` tags) and applies a grayed-out style during every server call. It uses an internal lock counter to ensure buttons are only re-enabled when the *last* active call finishes. Direct use of `google.script.run` is prohibited in feature sidebars.

**Action Helpers**: Standard sidebar flows should use the higher-level helpers from `SidebarShared.html`:
- `SyncSidebar.initSidebar()` for common startup and icon hydration.
- `SyncSidebar.runPullAction()` for pull/import actions, especially where unsaved sheet changes may be overwritten.
- `SyncSidebar.runPushAction()` for push/apply actions, including optional preflight confirmation steps.
- `SyncSidebar.runAction()` for single-action or bespoke operations that still need standardized loading, success, and error behavior.
- Supporting shared utilities include `SyncSidebar.confirmIfUnsaved()`, `SyncSidebar.confirmAndRun()`, `SyncSidebar.setStatusBadge()`, `SyncSidebar.updateQuotaDisplay()`, `SyncSidebar.markQuotaError()`, `SyncSidebar.showToast()`, and `SyncSidebar.handleError()`.

**Current Adoption Pattern**: Sync-oriented sidebars now primarily compose behavior through `initSidebar`, `runPullAction`, `runPushAction`, and `runAction` instead of hand-rolled `google.script.run` flows. This shared layer is the preferred path for all new or refactored sidebar interactions.

### 6a. Shared Sidebar Shell
`SidebarShared.html` also provides the canonical shared shell tokens and utility classes for sidebar composition:
- Shared semantic action buttons: `btn-pull`, `btn-push`
- Shared layout primitives: `sync-sidebar-action-grid`, `sync-sidebar-action-stack`, `sync-sidebar-inline-options`
- Shared shell classes for headers, cards, section labels, button groups, and status badges: `header`, `header-title`, `card`, `sync-sidebar-card`, `section-label`, `sync-sidebar-section-label`, `btn-group`, `sync-sidebar-button-group`, `status-badge`

Sync-oriented sidebars should converge on shared pull/push semantics and reuse these primitives instead of redefining the same action rail patterns locally.
- **Icons**: All sidebars must use the Lucide icon framework exclusively (<i data-lucide="..."></i>).
- **Styling Guidance**: Standard sync sidebars should avoid redefining generic shell styles and should rely on the tokens and primitives inside `SidebarShared.html`. Purpose-built dashboards such as `Settings_Sidebar.html` and `PipelineControl_Sidebar.html` may keep local, scoped layout styles when their UX is materially different, but should still reuse `SidebarShared.html` for runtime behaviors and shared overlays.

### 7. Standard Tooltip & Help Architecture
To maintain high usability without cluttering the UI, the project uses a standardized tooltip system:
- **`help-trigger`**: A CSS class applied to icons (usually `help-circle`).
- **`data-help-target`**: An attribute on the trigger that points to the `ID` of a hidden content element.
- **Hidden Content Container**: A `div` at the bottom of the HTML file (set to `display: none`) containing multiple divs with specific IDs (e.g., `help-getting-started`).
- **Global Event Handlers**: `SidebarShared.html` contains the logic to calculate tooltip positioning, handle boundary overflows, and manage transitions.
- **Guide Section**: Sidebars that need guided onboarding can include a bottom "Help & Guide" card using standardized `.sync-sidebar-help-guide-card` and `.sync-sidebar-help-guide-item` classes.

## 🌍 Global Variables & State

- **`SHEET_THEME`**: A Proxy object in `01_Config_Theme.js` that provides access to theme colors and styles.
- **`SHEET_NAMES`**: Centralized mapping of internal keys to actual tab names in `00_Config_Constants.js`.
- **`APP_PROPS`**: Metadata registry for all `PropertiesService` keys in `00_Config_Constants.js`.

## 🔌 Connection Flow (Frontend <-> Backend)
1. **Trigger**: User clicks menu or sidebar button.
2. **Launch**: `_App_openSidebar('TOOL_KEY')` handles sheet prep and sidebar rendering.
3. **Execution**: Sidebar calls `SyncSidebar.run('ToolName_publicFunc')` -> Backend function -> `Logger.run()` for a consistent execution boundary.
4. **Response**: Backend returns standardized `_App_ok(...)` / `_App_fail(...)` payloads, usually carrying row or sidebar data in `data` and optional metadata in `meta`.

### 🛠️ Developer Reporting & Error Architecture
The project employs a silent, user-centric error handling system that prioritizes spreadsheet feedback over background logs:
- **Silent Backend**: Handled by `Logger.js`, which provides `Logger.run` for public entry points and no-op logger methods for internal status calls. All technical console logging (`console.log` / `console.warn` / `console.error`) is strictly forbidden to maintain zero noise in the Apps Script execution logs.
- **System-Wide UI Errors**: Critical system-level failures are surfaced via a copyable modal overlay in the sidebar, providing technical details for support without requiring console access.
- **Row-Level Feedback**: All granular, data-specific feedback (success or failure) is reported directly in the **Status** column of the tool sheet. The `_App_BatchProcessor` facilitates this by returning structured result objects to the `onBatchComplete` hook. All statuses MUST use the standardized `SHEET_THEME.STATUS_PREFIXES` (✅, ❌, ⚠️).

### 📐 Unified Reporting Architecture
To maintain a professional and consistent user experience, all tools must adhere to the following reporting standards:
1. **Row-Level Status**:
   - **Success**: Must be prefixed with `SHEET_THEME.STATUS_PREFIXES.SUCCESS` (✅).
   - **Failure**: Must be prefixed with `SHEET_THEME.STATUS_PREFIXES.ERROR` (❌) and contain clear, actionable error details.
   - **Warning**: Must be prefixed with `SHEET_THEME.STATUS_PREFIXES.WARNING` (⚠️) for non-blocking issues.
2. **General/System Errors**:
   - All errors surfaced in the sidebar via `SyncSidebar.handleError` should specify a `severity` level:
     - `mild`: Blue "Notice" modal for informational non-errors.
     - `medium`: Amber "Warning" modal for recoverable issues.
     - `critical`: Red "System Error" modal for blocking/critical failures (Default).
3. **Unification**: These standards are enforced via the `SHEET_THEME` configuration and the `SidebarShared.html` engine.

### 🔄 Unified Batch Processing (`_App_BatchProcessor`)
To ensure consistency and performance across all tools, row-by-row operations must use the centralized processor in `03_Core_Utils.js`.
- **Automatic Retries**: Wraps each item in `_App_callWithBackoff` to handle transient API errors.
- **Progress Tracking**: Automatically updates `CacheService` with progress data for sidebar polling.
- **Time-Limit Guarding**: Monitors the Google Apps Script 6-minute limit and pauses execution at 5.5 minutes, allowing for safe partial completions.
- **Batch Updates**: Encourages the use of `SheetManager.batchPatchRows` within the `onBatchComplete` hook to minimize Spreadsheet API calls.

### ⏳ Execution Time Management
Global timing is managed via `_App_resetExecutionTimer()` and `_App_isExecutionLimitApproaching()`. Tools with long-running recursive or iterative tasks should check this limit frequently to prevent hard script timeouts.

