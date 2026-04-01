# 🦫 Beaver Appscripts - Project Context

## Project Overview
**Beaver Appscripts** is a professional-grade suite of Google Sheets automation tools. It provides a unified framework for various productivity tasks (Mail Merge, Calendar Sync, Contacts Management, etc.) through custom HTML sidebars and a centralized core engine.

- **Primary Technologies:** Google Apps Script (V8 Runtime), HTML/CSS/JS (Sidebars), Google Workspace APIs.
- **Architecture:** Modular Plugin Architecture. Core system logic is split into sequential files (`00_Config_Constants.js` through `09_Engine_UI.js`), which are evaluated in order by the GAS runtime.

## 🏗️ Core Architecture & Engine
The project uses a decentralized registration pattern to manage its toolset:

- **`BeaverEngine` (`08_Engine_Core.js`)**: The central registry. Each tool backend file registers itself with its configuration (sheet name, schema, sidebar HTML, etc.) using `BeaverEngine.registerTool()`.
- **`_App_` Core Utilities**: Shared logic for UI launching, sheet management, rate limiting, and backoff is prefixed with `_App_` and spread across the `03`-`09` core files.
- **`Logger` (`Logger.js`)**: A robust, concurrency-safe logging system using a transporter pattern (caching in `CacheService` before flushing to a sheet).

## 🚀 Key Commands & Workflow
This project uses **Clasp** (Command Line Apps Script Projects) for local development.

- **Pull Code:** `clasp pull`
- **Deploy/Push:** `clasp push`
- **Open Script Editor:** `clasp open`
- **Testing:** Since this is a Google Workspace add-on/script, testing is performed by running functions directly from the Apps Script editor or interacting with the "🦫 Beaver Tools" menu in the associated Google Sheet.

## 🛠️ Development Conventions

### 1. Naming Standards
- **Public Backend Functions:** `ToolName_FunctionName` (e.g., `MailMerge_openSidebar`). These are called from Sidebars or the UI menu.
- **Internal Helper Functions:** `_ToolName_InternalFunction` (e.g., `_MailMerge_validateData`).
- **Core System Utilities:** `_App_UtilityName` (e.g., `_App_launchTool`).

### 2. Implementation Patterns
- **Tool Registration:** Every new tool must register with `BeaverEngine` at the top of its backend file.
- **Execution Wrapper:** Primary operations should be wrapped in `Logger.run('TOOL_KEY', 'Context', () => { ... })` for automatic tracking and error handling.
- **Return Contract**: Public functions called by `google.script.run` MUST return an object: `{ success: boolean, message: string }`.
- **Trigger Management**: Background sync tools should manage their own `ScriptApp` triggers. Use an internal `_ToolName_manageTrigger` function called from the setting update handler to ensure triggers are created/removed in sync with user preferences.
- **UI Menu**: Entry points are defined in `UI.js` and dynamically built from `BeaverEngine` metadata.

### 3. Storage (`PropertiesService`)
- Never access `PropertiesService` directly. Use the registry in `00_Config_Constants.js` (`APP_PROPS`) and the wrappers in `02_Config_Storage.js` (`_App_getProperty`, `_App_setProperty`).

## 📂 Project Structure Highlights
- **`00_Config_Constants.js`**: Global registries for sheet names and property keys.
- **`01_SheetManager.js`**: Data Access Object (DAO) mapping sheet rows to objects.
- **`UI.js`**: The central UI orchestrator and menu builder.
- **`Blueprint.md`**: The primary architectural reference for developers and AI agents.
- **`appsscript.json`**: Manifest file defining API scopes and advanced services.
- **Tool Modules**: Paired `.js` (Backend) and `.html` (Sidebar) files (e.g., `Mail_Sender.js` + `Mail_Sender-Sidebar.html`).
