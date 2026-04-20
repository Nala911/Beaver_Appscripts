/**
 * MASTER CONFIGURATION & NAMESPACE INITIALIZATION
 * ==========================================
 * This file is named 00_Config_Constants so it gets evaluated first.
 * It initializes the unified 'App' namespace.
 */

var App = {
    Config: {},
    UI: {},
    Data: {},
    Engine: {},
    Utils: {},
    Storage: {},
    Log: {}
};

App.Config = {
    SHEET_NAMES: {
            CALENDAR_SYNC: '🗓️ Google Calendar',
            CONTACTS_SYNC: '☎️ Google Contacts',
            MAIL_MERGE: '📧 Mail Merge',
            MAIL_SENDER: '📩 Mail Sender',
            DOCS_MERGE: '📄 Docs Merge',
            TASKS: '✅ Google Tasks',
            FORMS_SYNC: '📝 Google Forms',
            BULK_FOLDER: '📂 Bulk Folder Creation',
            DRIVE_SYNC: '💾 Google Drive',
            PIPELINE: '⛓  Control Center',
            CHAT_SPACE_SYNC: '💬 Google Chat Spaces',
            LOGS: '🛠️ Developer Log'
        },
        STORE_TYPES: {
            DOCUMENT: 'DOCUMENT',
            USER: 'USER',
            SCRIPT: 'SCRIPT'
        },
        APP_PROPS: {
            // UI Theme
            THEME: { key: 'WorkspaceSync_SHEET_THEME', store: 'DOCUMENT', isJson: true },

            // Pipeline Control
            SYSTEM_ENABLED: { key: 'SYSTEM_ENABLED', store: 'SCRIPT', isJson: false },

            // Docs Merge
            DOCS_MERGE_TEMPLATE_URL: { key: 'DOCS_MERGE_TEMPLATE_URL', store: 'DOCUMENT', isJson: false },
            DOCS_MERGE_FOLDER_URL: { key: 'DOCS_MERGE_FOLDER_URL', store: 'DOCUMENT', isJson: false },
            DOCS_MERGE_TEMPLATE_NAME: { key: 'DOCS_MERGE_TEMPLATE_NAME', store: 'DOCUMENT', isJson: false },
            DOCS_MERGE_FOLDER_NAME: { key: 'DOCS_MERGE_FOLDER_NAME', store: 'DOCUMENT', isJson: false },

            // Calendar Sync
            CAL_SELECTED_IDS: { key: 'selectedCalIds', store: 'USER', isJson: true },
            CAL_START_DATE: { key: 'startDate', store: 'USER', isJson: false },
            CAL_END_DATE: { key: 'endDate', store: 'USER', isJson: false },

            // Chat Space Sync
            CHAT_SELECTED_SPACES: { key: 'selectedChatSpaces', store: 'USER', isJson: true },

            // Contacts Sync
            CONTACTS_SELECTED_GROUPS: { key: 'selectedContactGroups', store: 'USER', isJson: true },

            // Forms Sync
            FORMS_CURRENT_FORM: { key: 'FORMSSYNC_CURRENT_FORM', store: 'DOCUMENT', isJson: false },
            FORMS_SELECTED_FORM: { key: 'FORMSSYNC_SELECTED_FORM', store: 'USER', isJson: false },

            // Developer Settings
            ENABLE_DEBUG_LOGGING: { key: 'ENABLE_DEBUG_LOGGING', store: 'DOCUMENT', isJson: false },
            LOGGER_MAX_ROWS: { key: 'LOGGER_MAX_ROWS', store: 'DOCUMENT', isJson: false }
        },
        CACHE_KEYS: {
            LOGS: 'WorkspaceSync_DEBUG_LOGS',
            PROGRESS: '_PROGRESS'
        },
        TOOL_LAUNCH_MODES: {
            SIDEBAR: 'SIDEBAR',
            MODAL: 'MODAL'
        }
};

// Backward Compatibility Layer (Ensures legacy files still work)
var SHEET_NAMES = App.Config.SHEET_NAMES;
var STORE_TYPES = App.Config.STORE_TYPES;
var APP_PROPS = App.Config.APP_PROPS;
var CACHE_KEYS = App.Config.CACHE_KEYS;
var TOOL_LAUNCH_MODES = App.Config.TOOL_LAUNCH_MODES;
