# 🦫 WorkspaceSync Appscripts - BLUEPRINT & Architecture Context

## Overview
This file serves as the absolute Architectural Map for AI agents working on the "WorkspaceSync Appscripts" workspace. It contains details on file connections, global state management, and module structures.

> [!NOTE]
> Workflow rules and procedural instructions for agents are strictly located in `AGENTS.md`. Refer to that file before making any changes.

## 📂 File System Structure

The project consists of `.js` (Google Apps Script server-side code) and `.html` (Sidebar interfaces) pairs for each tool.

### Core System Files
The system logic is split into sequential modules evaluated in order:
- `00_Config_Constants.js`: Global registries, `SHEET_NAMES`, `APP_PROPS`, and enum structures.
- `01_Config_Theme.js`: Default theme definitions, colors, and `SHEET_THEME` proxy.
- `02_Config_Storage.js`: Unified properties service wrappers (`setAppProp`, `getAppProp`).
- `03_Core_Utils.js`: Core utilities (`_App_throttle`, `_App_callWithBackoff`, `_App_setProgress`, etc.).
- `04_Core_Validators.js`: Validation helpers for types and constraints.
- `05_Core_State.js`: Global application state management.
- `06_Sheets_Helpers.js`: Low-level spreadsheet operations (`_App_ensureSheetExists`, etc.).
- `07_Sheets_Formatting.js`: UI/styling application to sheets (`_App_applyBodyFormatting`).
- `08_Engine_Core.js`: The `SyncEngine` plugin registration and retrieval system.
- `09_Engine_UI.js`: UI abstractions for opening sidebars and dialogs.
- `UI.js`: The central UI orchestrator. Responsible for creating the custom "🦫 WorkspaceSync Tools" menu (`onOpen`), providing the global wrapper for the Theme Editor sidebar, and connecting user actions to the tools.
- `SidebarShared.html`: Shared HTML, CSS, and JS components to eliminate redundant sidebar code and infinite spinners.
- `01_SheetManager.js`: Centralized data access object (DAO). Uses `SyncEngine` configurations to map sheet data to JavaScript objects and vice-versa.
- `Logger.js`: Unified logging system. Provides `Logger.info`, `Logger.error`, etc., using a buffered transporter architecture with `CacheService` and `LockService` for performant, concurrent-safe logging. Registers itself with `SyncEngine`.
- `SystemAudit.js`: Runs comprehensive audits across all registered tools, verifying sheet integrity, API access, and schema setup, and generates AI debug output logs.
- `Logger_SidebarController.js`: Backend controller for the Developer Log sidebar, handling client-to-server interactions like fetching logs and running system audits.
- `appsscript.json` / `.clasp.json`: Google Apps Script configuration and Clasp deployment environment details.

### Tool Modules & Connections
Each tool has a Backend file, a Frontend sidebar file, and a global Entry Function triggered by the custom menu in `UI.js`.

| Tool Name | Backend (`.js`) | Frontend (`.html`) | UI Menu Entry Function |
|---|---|---|---|
| **Calendar Sync** | `Calendar_Sync.js` | `Calender_Sidebar.html` | `Calendar_showSidebar` |
| **Contacts Sync** | `Contacts_Sync.js` | `ContactsSidebar.html` | `Contacts_showSidebar` |
| **Mail Merge** | `Mail_merge_Code.js` | `Mail_merge_HTML.html` | `MailMerge_openSidebar` |
| **Mail Sender** | `Mail_Sender.js` | `Mail_Sender-Sidebar.html` | `Mail_Sender_openSidebar` |
| **Docs Merge** | `Docs_Merge_Code.js` | `Docs_merge_Sidebar.html` | `DocsMerge_openSidebar` |
| **Task Manager** | `Task_Sync_code.js` | `Tasks_Sidebar.html` | `Tasks_showSidebar` |
| **Forms Sync** | `FormsSync_Code.js` | `FormsSync_Sidebar.html` | `FormsSync_openSidebar` |
| **Bulk Folder** | `BulkFolderCreation.js` | `BulkFolderCreationSidebar.html` | `BulkFolder_showSidebar` |
| **Drive Sync** | `DriveFileDetails.gs.js` | `DriveFileDetailsSidebar.html` | `Drive_showSidebar` |
| **Pipeline Control** | `PipelineControl.js` | `PipelineSidebar.html` | `Pipeline_showSidebar` |
| **Developer Log** | `Logger.js`, `SystemAudit.js`, `Logger_SidebarController.js` | `Logger_Sidebar.html` | `Logger_showSidebar` |
| **Theme Editor** | (Inside `UI.js`) | `ThemeEditorSidebar.html` | `UI_openThemeDialog` |

> [!CAUTION]
> **Large File Warning:** The following files are large (25KB+). Use surgical reads.
> - `DriveFileDetails.gs.js` (~32KB): Complex Drive synchronization logic.
> - `Contacts_Sync.js` (~27KB): People API integration logic.

## 🔑 Google API Scopes & Services Used

Each tool relies on specific Google APIs. Do NOT use an API in a tool that doesn't need it.

| Tool | Google APIs / Services | Advanced Service? |
|---|---|---|
| **Calendar Sync** | `CalendarApp`, `Calendar` (Advanced) | Yes — `Calendar API v3` |
| **Contacts Sync** | `People` (Advanced) | Yes — `People API v1` |
| **Mail Merge** | `GmailApp`, `DocumentApp` | No |
| **Mail Sender** | `GmailApp`, `MailApp` | No |
| **Docs Merge** | `DocumentApp`, `DriveApp` | No |
| **Task Manager** | `Tasks` (Advanced) | Yes — `Tasks API v1` |
| **Forms Sync** | `FormApp`, `DriveApp` | No |
| **Bulk Folder** | `DriveApp` | No |
| **Pipeline Sync** | `DriveApp`, `Drive` (Advanced) | Yes — `Drive API v3` |
| **Pipeline Control** | `PropertiesService`, `SpreadsheetApp`, `ScriptApp` | No |
| **Developer Log** | `CacheService`, `Session`, `Utilities` | No |
| **Theme Editor** | `PropertiesService` only | No |

## 🗝️ PropertiesService Key Registry

All keys used across the codebase. **Do NOT invent new key names** — check here first and follow the naming convention in `APP_PROPS`.

| Key | File | Store Type | Purpose |
|---|---|---|---|
| `WorkspaceSync_SHEET_THEME` | `01_Config_Theme.js` | `DocumentProperties` | Custom theme JSON overrides |
| `SYSTEM_ENABLED` | `PipelineControl.js` | `ScriptProperties` | Master on/off toggle for pipeline |
| `DOCS_MERGE_TEMPLATE_URL` | `Docs_Merge_Code.js` | `DocumentProperties` | Saved template Doc URL |
| `DOCS_MERGE_FOLDER_URL` | `Docs_Merge_Code.js` | `DocumentProperties` | Saved output folder URL |
| `DOCS_MERGE_TEMPLATE_NAME` | `Docs_Merge_Code.js` | `DocumentProperties` | Cached template file name |
| `DOCS_MERGE_FOLDER_NAME` | `Docs_Merge_Code.js` | `DocumentProperties` | Cached folder name |
| `selectedCalIds` | `Calendar_Sync.js` | `UserProperties` | JSON array of selected calendar IDs |
| `startDate` | `Calendar_Sync.js` | `UserProperties` | Saved start date filter |
| `endDate` | `Calendar_Sync.js` | `UserProperties` | Saved end date filter |
| `selectedContactGroups` | `Contacts_Sync.js` | `UserProperties` | JSON array of selected contact group IDs |
| `FORMSSYNC_CURRENT_FORM` | `FormsSync_Code.js` | `DocumentProperties` | Stores currently synced form ID |
| `FORMSSYNC_SELECTED_FORM` | `FormsSync_Code.js` | `UserProperties` | Stores user's selected form ID for sidebar auto-selection |
| `ENABLE_DEBUG_LOGGING` | `Logger.js` | `DocumentProperties` | Global on/off for debug logging ('true'/'false') |
| `LOGGER_MAX_ROWS` | `Logger.js` | `DocumentProperties` | Max rows to keep in the Log sheet |

## 🏗️ Architectural Patterns

The codebase follows a strict and predictable design pattern across all tools:

### 1. Decentralized Plugin Architecture (`SyncEngine`)
The project uses a decentralized registration pattern to manage tools.
- **`SyncEngine`**: A singleton in `08_Engine_Core.js` that handles tool registration (`registerTool`) and retrieval (`getTool`).
- **Self-Registration**: Each tool module registers its own configuration block at the top of its file.
- **Registry Metadata**: Configuration includes `SHEET_NAME`, `TITLE`, `SIDEBAR_HTML`, `COL_WIDTHS`, and a `COL_SCHEMA` for declarative column validations and types.

### 2. Unified Utilities (`_App_`)
Core logic is abstracted into `_App_` prefixed functions spread across `03_Core_Utils.js` to `09_Engine_UI.js`.

### 3. Flat Function Prefix Architecture
To avoid naming collisions and facilitate AI interaction, functions use `ToolName_` prefixes.

### 4. Global Entry Points
Entry functions for the UI menu generally look like:
```javascript
function ToolName_showSidebar() {
  _App_openSidebar('TOOL_KEY');
}
```

### 5. Background Automation & Trigger Management
Tools that require background execution should manage their own triggers programmatically.

### 6. Development Templates (The "Golden" Pattern)
For creating any new tool, or modifying existing ones, agents MUST refer to the following template files as the source of truth for design patterns:
- **Backend Pattern:** `Template_Tool_Code.js`
- **Frontend Pattern:** `Template_Tool_Sidebar.html`
These files contain the latest mandatory registration logic, column schema definitions, and UI status-handling patterns required by the core engine.


## 🌍 Global Variables & State

- **`SHEET_THEME`**: A Proxy object in `01_Config_Theme.js` that provides access to theme colors and styles.
- **`SHEET_NAMES`**: Centralized mapping of internal keys to actual tab names in `00_Config_Constants.js`.
- **`APP_PROPS`**: Metadata registry for all `PropertiesService` keys in `00_Config_Constants.js`.

## 🔌 Connection Flow (Frontend <-> Backend)
1. **Trigger**: User clicks menu or sidebar button.
2. **Launch**: `_App_openSidebar('TOOL_KEY')` handles sheet prep and sidebar rendering.
3. **Execution**: Sidebar calls `google.script.run` -> Backend function -> `Logger.run()` for automatic logging/error tracking.
4. **Response**: Backend returns `{ success: true, message: "..." }`.

## 🛠️ Developer Logging Architecture
The project employs a robust, asynchronous-style logging system.
- **Transporter Pattern**: Logs are first queued into `CacheService`.
- **Orchestration**: `Logger.run()` serves as an execution supervisor.
- **Concurrency Safety**: Uses `LockService.getDocumentLock()` during the "flush" phase.
