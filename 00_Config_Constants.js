// Global Engine Initialization & Constants
// ==========================================
// This file is named 00_AppConfig so it gets evaluated first by the Google Apps Script runtime.

// Global Sheet Names Configuration
var SHEET_NAMES = {
    CALENDAR_SYNC: '🗓️ Google Calendar',
    CONTACTS_SYNC: '☎️ Google Contacts',
    MAIL_MERGE: '📧 Mail Merge',
    MAIL_SENDER: '📩 Mail Sender',
    DOCS_MERGE: '📄 Docs Merge',
    TASKS: '✅ Google Tasks',
    FORMS_SYNC: '📝 Google Forms',
    BULK_FOLDER: '📂 Bulk Folder Creation',
    DRIVE_SYNC: '💾 Google Drive',
    PIPELINE: '⛓  Pipeline',
    CHAT_SPACE_SYNC: '💬 Google Chat Spaces',
    GMAIL_FILTERS: '🗂️ Gmail Filters',
    NEWSLETTER_ZERO: '🧹 Newsletter Zero',
    LOGS: '🛠️ Developer Log'
};

// ==========================================
// Centralized Storage Registry (PropertiesService)
// ==========================================

var STORE_TYPES = {
    DOCUMENT: 'DOCUMENT',
    USER: 'USER',
    SCRIPT: 'SCRIPT'
};

var APP_PROPS = {
    // UI Theme
    THEME: { key: 'WorkspaceSync_SHEET_THEME', store: STORE_TYPES.DOCUMENT, isJson: true },

    // Pipeline
    SYSTEM_ENABLED: { key: 'SYSTEM_ENABLED', store: STORE_TYPES.SCRIPT, isJson: false },

    // Docs Merge
    DOCS_MERGE_TEMPLATE_URL: { key: 'DOCS_MERGE_TEMPLATE_URL', store: STORE_TYPES.DOCUMENT, isJson: false },
    DOCS_MERGE_FOLDER_URL: { key: 'DOCS_MERGE_FOLDER_URL', store: STORE_TYPES.DOCUMENT, isJson: false },
    DOCS_MERGE_TEMPLATE_NAME: { key: 'DOCS_MERGE_TEMPLATE_NAME', store: STORE_TYPES.DOCUMENT, isJson: false },
    DOCS_MERGE_FOLDER_NAME: { key: 'DOCS_MERGE_FOLDER_NAME', store: STORE_TYPES.DOCUMENT, isJson: false },

    // Calendar Sync
    CAL_SELECTED_IDS: { key: 'selectedCalIds', store: STORE_TYPES.USER, isJson: true },
    CAL_START_DATE: { key: 'startDate', store: STORE_TYPES.USER, isJson: false },
    CAL_END_DATE: { key: 'endDate', store: STORE_TYPES.USER, isJson: false },

    // Chat Space Sync
    CHAT_SELECTED_SPACES: { key: 'selectedChatSpaces', store: STORE_TYPES.USER, isJson: true },

    // Contacts Sync
    CONTACTS_SELECTED_GROUPS: { key: 'selectedContactGroups', store: STORE_TYPES.USER, isJson: true },

    // Forms Sync
    FORMS_CURRENT_FORM: { key: 'FORMSSYNC_CURRENT_FORM', store: STORE_TYPES.DOCUMENT, isJson: false },
    FORMS_SELECTED_FORM: { key: 'FORMSSYNC_SELECTED_FORM', store: STORE_TYPES.USER, isJson: false },

    // Newsletter Zero
    NEWSLETTER_ZERO_SCAN_DAYS: { key: 'NEWSLETTER_ZERO_SCAN_DAYS', store: STORE_TYPES.USER, isJson: false },

    // Developer Settings
    ENABLE_DEBUG_LOGGING: { key: 'ENABLE_DEBUG_LOGGING', store: STORE_TYPES.DOCUMENT, isJson: false },
    LOGGER_MAX_ROWS: { key: 'LOGGER_MAX_ROWS', store: STORE_TYPES.DOCUMENT, isJson: false }
};

var CACHE_KEYS = {
    LOGS: 'WorkspaceSync_DEBUG_LOGS',
    PROGRESS: '_PROGRESS'
};

var TOOL_LAUNCH_MODES = {
    SIDEBAR: 'SIDEBAR',
    MODAL: 'MODAL'
};
