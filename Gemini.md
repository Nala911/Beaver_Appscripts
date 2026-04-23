# WorkspaceSync Appscripts - Gemini Guide

Read this before changing any `.js` or `.html` file in this repository. **This project is maintained 100% by AI agents.** Human coders do not actively edit this codebase. Therefore, all architectural rules, patterns, and conventions must be strictly preserved to maintain systematic scalability.

## 🚀 Key Commands & Workflow
This project uses **Clasp** (Command Line Apps Script Projects) for local development.

- **Pull Code:** `clasp pull`
- **Deploy/Push:** `clasp push`
- **Open Script Editor:** `clasp open`
- **Testing:** Since this is a Google Workspace add-on/script, testing is performed by running functions directly from the Apps Script editor or interacting with the "🦫 WorkspaceSync Tools" menu in the associated Google Sheet.

## 🏛️ Ground Rules & Source Of Truth

- Treat `BLUEPRINT.md` as the ultimate architectural map. It contains connection flows and the scope of external APIs.
- Treat `00_Config_Constants.js` as the source of truth for global state keys and sheet names.
- **Do not invent new patterns.** If a tool needs local storage, use the properties registry. If a tool needs to modify the spreadsheet UI, it must use the `_App_` ecosystem.
- Before adding a new tool, use `CalendarSync_Code.js` and `CalendarSync_Sidebar.html` as your primary architectural benchmark models.

## ⛔ "Do Not Touch" Core Modules

The system is split into two halves: the Core Engine and the Tool Modules. Agents maintaining specific features should treat the **Core Engine** files as immutable unless explicitly tasked to refactor the framework itself.

**Core Engine Files (Do Not Modify for Feature Work):**
- `00_Config_Constants.js` through `09_Engine_UI.js`
- `01_SheetManager.js`
- `UI.js`
- `SidebarShared.html`
- `Logger.js`, `SystemAudit.js`, and `Logger_SidebarController.js`

If you are just editing or adding a feature (like Mail Merge, Tasks Sync, etc.), stick entirely to your tool's `_Code.js` and `_Sidebar.html` files.

## ⚙️ Mandatory Code Contracts

### 1. Naming Standards
- **Tool Backend Logic:** `<ToolName>_Code.js` (e.g., `MailMerge_Code.js`). Always PascalCase for the tool name.
- **Tool Sidebar Interface:** `<ToolName>_Sidebar.html` (e.g., `MailMerge_Sidebar.html`).
- **Public Backend Functions:** `ToolName_FunctionName` (e.g., `MailMerge_openSidebar`). These are called from Sidebars or the UI menu.
- **Internal Helper Functions:** `_ToolName_InternalFunction` (e.g., `_MailMerge_validateData`).
- **Core System Utilities:** `_App_UtilityName` (e.g., `_App_launchTool`).
- **HTML Includes (Partials):** `ToolName_CSS.html` or `ToolName_JS.html`. Use `<?!= _App_include('ToolName_Type'); ?>` to include them.

### 1a. UI Naming Standards (Uniformity)
To maintain a professional and consistent user experience, the following strings MUST match exactly:
- **`TITLE`** (in `SyncEngine.registerTool`): Must match `SHEET_NAMES[KEY]` exactly (including emoji).
- **`MENU_LABEL`** (in `SyncEngine.registerTool`): Must match `SHEET_NAMES[KEY]` exactly (including emoji).
- **`Blueprint.md` (Bold Tool Name)**: Must match `SHEET_NAMES[KEY]` but without the emoji.
- **Sidebar Header (`.header-title`)**: Must match the base tool name (without emoji or suffixes like "Toolkit").

### 2. The SyncEngine Contract
Every tool backend file must register itself with the engine at the very top of the script using `SyncEngine.registerTool(key, config)`. Do not hardcode columns inside backend logic; rely on the registry's `FORMAT_CONFIG.COL_SCHEMA`.

### 3. The Logger.run Contract
Every public function called from a sidebar or the Ribbon UI must use the `Logger.run` execution wrapper to ensure errors are caught and recorded. DO NOT manually try/catch to write error strings into cells. 
```javascript
function MyTool_publicFunction() {
    return Logger.run('MY_TOOL', 'Action Context', function() {
        // ... logic
        return { success: true, message: "Done" };
    });
}
```

### 4. Return Contract
Public functions called by `google.script.run` MUST return an object: `{ success: boolean, message: string }`.

### 5. Trigger Management
Background sync tools should manage their own `ScriptApp` triggers. Use an internal `_ToolName_manageTrigger` function called from the setting update handler to ensure triggers are created/removed in sync with user preferences.

### 6. The PropertiesService Contract
Never use `PropertiesService.getDocumentProperties()` directly in a tool. 
- Define your new key strictly in `APP_PROPS` inside `00_Config_Constants.js`.
- Use `_App_getProperty` and `_App_setProperty` from `02_Config_Storage.js`.

### 8. Batch Processing & Time Limits
Row-by-row data processing must use `_App_BatchProcessor` from `03_Core_Utils.js`. This utility handles progress tracking, backoff retries, and protects against the 6-minute script timeout. Use `SheetManager.batchPatchRows` for efficient data writing within the batch complete hook.

## 🤖 Gemini Workflow Rules

1. **Minimize file reads**: ONLY read the specific tool files needed.
2. **Consult Core Modules first**: Global configuration and logic are defined in `00_Config_Constants.js` through `09_Engine_UI.js`.
3. **Use `Logger.run()`**: Wrap primary tool operations in `Logger.run('TOOL_KEY', 'Context', () => { ... })` for consistent logging.
4. **Follow `SyncEngine`**: When modifying sheet structure, update the registration metadata in the tool's backend file, which registers with `SyncEngine`.

## 📋 Gemini Pre-Flight Checklist

Before completing any task, mentally run this checklist. Do not proceed until you have verified all points:

- [ ] Does my backend file register with `SyncEngine` at the very top?
- [ ] Are all public functions prefixed with `ToolName_` (e.g., `MailMerge_doWork`)?
- [ ] Are all internal helper functions prefixed with `_ToolName_` (e.g., `_MailMerge_validate`)?
- [ ] Did I wrap my core action inside `Logger.run('KEY', 'Context', function() {...})`?
- [ ] Does my public function return exactly `{ success: boolean, message: string }`?
- [ ] Did I use `_App_callWithBackoff` around any external Google API calls?
- [ ] If I added a new setting, is it declared in `APP_PROPS` in `00_Config_Constants.js`?
- [ ] If my tool processes rows, did I use `_App_BatchProcessor` and `SheetManager.batchPatchRows`?
- [ ] Is my sidebar strictly including `<?!= _App_include('SidebarShared'); ?>` to inherit standard WorkspaceSync UI libraries?

## 🚀 Adding a New Tool

Since humans do not code here, follow the **CalendarSync Benchmark**:
1. Duplicate `CalendarSync_Code.js` and rename it to `<NewName>_Code.js`.
2. Duplicate `CalendarSync_Sidebar.html` and rename it to `<NewName>_Sidebar.html`.
3. Add the `SHEET_NAMES` entry to `00_Config_Constants.js`.
4. Update the plugin registration inside `<NewName>_Code.js` (Key, Sheet Name, Title, etc.).
5. Implement backend logic with `Logger.run`.
6. Implement frontend logic using the unified `#status-div` (handled by `SyncSidebar`).
