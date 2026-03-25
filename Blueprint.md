# 🦫 Beaver Appscripts - Blueprint & Architecture Context

## Overview
This file serves as a reference guide for AI agents and developers working on the "Beaver Appscripts" workspace. The project is a suite of Google Sheets automation tools with custom HTML sidebars. This blueprint contains details on file connections, architectural patterns, global variables, and module structures to minimize time spent understanding the codebase.

## 📂 File System Structure

The project consists of `.js` (Google Apps Script server-side code) and `.html` (Sidebar interfaces) pairs for each tool.

### Core System Files
- `00_AppConfig.js`: The central configuration hub. Evaluated first by the runtime. Defines `BeaverEngine` for tool registration, global sheet names (`SHEET_NAMES`), base themes, property registry (`APP_PROPS`), and unified utilities (`_App_openSidebar`, `_App_ensureSheetExists`, `_App_applyBodyFormatting`, `_App_throttle`, `_App_callWithBackoff`, `_App_setProgress`).
- `UI.js`: The central UI orchestrator. Responsible for creating the custom "🦫 Beaver Tools" menu (`onOpen`), providing the global wrapper for the Theme Editor sidebar, and connecting user actions to the tools.
- `01_SheetManager.js`: Centralized data access object (DAO). Uses `BeaverEngine` configurations to map sheet data to JavaScript objects and vice-versa.
- `Logger.js`: Unified logging system. Provides `Logger.info`, `Logger.error`, etc., and manages the `🛠️ Developer Log` sheet. Registers itself with `BeaverEngine`.
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
> **Large File Warning:** The following files are large (30KB+). Use surgical reads.
> - `00_AppConfig.js` (~26KB): Global registry and core utilities.
> - `DriveFileDetails.gs.js` (~32KB): Complex Drive synchronization logic.
> - `Contacts_Sync.js` (~27KB): People API integration logic.
> - `PipelineSidebar.html` (~35KB): Complex reactive UI for Pipeline Control.
> - `ThemeEditorSidebar.html` (~49KB): Massive CSS/JS for the Theme Studio.

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
| **Drive Sync** | `DriveApp`, `Drive` (Advanced) | Yes — `Drive API v3` |
| **Pipeline Control** | `PropertiesService`, `SpreadsheetApp` | No |
| **Developer Log** | `CacheService`, `Session`, `Utilities` | No |
| **Theme Editor** | `PropertiesService` only | No |

## 🗝️ PropertiesService Key Registry

All keys used across the codebase. **Do NOT invent new key names** — check here first and follow the naming convention in `APP_PROPS`.

| Key | File | Store Type | Purpose |
|---|---|---|---|
| `BEAVER_SHEET_THEME` | `00_AppConfig.js` | `DocumentProperties` | Custom theme JSON overrides |
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

### 1. Decentralized Plugin Architecture (`BeaverEngine`)
The project uses a decentralized registration pattern to manage tools.
- **`BeaverEngine`**: A singleton in `00_AppConfig.js` that handles tool registration (`registerTool`) and retrieval (`getTool`).
- **Self-Registration**: Each tool module (e.g., `Mail_Sender.js`) registers its own configuration block at the top of its file.
- **Monolith Removal**: The old `APP_REGISTRY` monolith has been replaced. A backward-compatibility Proxy is maintained in `00_AppConfig.js` for legacy code.
- **Registry Metadata**: Configuration includes `SHEET_NAME`, `TITLE`, `SIDEBAR_HTML`, `COL_WIDTHS`, and a `COL_SCHEMA` for declarative column validations and types.

### 2. Unified Utilities (`_App_`)
Core logic is abstracted into `_App_` prefixed functions in `00_AppConfig.js`:
- `_App_openSidebar(toolKey)`: Opens the sidebar and ensures the sheet exists via `BeaverEngine`.
- `_App_ensureSheetExists(toolKey)`: Scaffolds the tool sheet based on `BeaverEngine` metadata.
- `_App_applyBodyFormatting(sheet, numRows, config)`: Applies consistent styling and conditional rules.
- `_App_callWithBackoff(func)`: Standardized exponential backoff for Google API calls.
- `_App_throttle(tracker, callsMade)`: Unified rate limiting for API batches.

### 3. Flat Function Prefix Architecture
To avoid naming collisions and facilitate AI interaction, functions use `ToolName_` prefixes.
- `ToolName_publicAction()`: Exposed to sidebar.
- `_ToolName_internalHelper()`: Internal logic.

### 4. Global Entry Points
Entry functions for the UI menu generally look like:
```javascript
function ToolName_showSidebar() {
  _App_openSidebar('TOOL_KEY');
}
```

## 🌍 Global Variables & State

- **`SHEET_THEME`**: A Proxy object in `00_AppConfig.js` that provides access to theme colors and styles, lazily loading from `PropertiesService`.
- **`SHEET_NAMES`**: Centralized mapping of internal keys to actual tab names.
- **`APP_PROPS`**: Metadata registry for all `PropertiesService` keys to ensure type safety (JSON vs String) and store consistency.

## 🔌 Connection Flow (Frontend <-> Backend)
1. **Trigger**: User clicks menu or sidebar button.
2. **Launch**: `_App_openSidebar('TOOL_KEY')` handles sheet prep and sidebar rendering.
3. **Execution**: Sidebar calls `google.script.run` -> Backend function -> `Logger.run()` for automatic logging/error tracking.
4. **Response**: Backend returns `{ success: true, message: "..." }`.

## ⚠️ Standard Return Contract
All public functions called by sidebars MUST return:
```javascript
{ success: true, message: "Success message" } // Or success: false
```

## 🤖 AI Agent Workflow Rules
1. **Minimize file reads**: ONLY read the specific tool files needed.
2. **Consult `00_AppConfig.js` first**: Most configuration and UI logic is defined there.
3. **Use `Logger.run()`**: Wrap primary tool operations in `Logger.run('TOOL_KEY', 'Context', () => { ... })` for consistent logging.
4. **Follow `APP_REGISTRY`**: When modifying sheet structure, update the registry in `00_AppConfig.js` rather than hardcoding in the tool file.
