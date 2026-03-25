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
    PIPELINE: '⛓  Control Center',
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
    THEME: { key: 'BEAVER_SHEET_THEME', store: STORE_TYPES.DOCUMENT, isJson: true },

    // Pipeline Control
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

    // Contacts Sync
    CONTACTS_SELECTED_GROUPS: { key: 'selectedContactGroups', store: STORE_TYPES.USER, isJson: true },

    // Forms Sync
    FORMS_CURRENT_FORM: { key: 'FORMSSYNC_CURRENT_FORM', store: STORE_TYPES.DOCUMENT, isJson: false },
    FORMS_SELECTED_FORM: { key: 'FORMSSYNC_SELECTED_FORM', store: STORE_TYPES.USER, isJson: false },

    // Developer Settings
    ENABLE_DEBUG_LOGGING: { key: 'ENABLE_DEBUG_LOGGING', store: STORE_TYPES.DOCUMENT, isJson: false },
    LOGGER_MAX_ROWS: { key: 'LOGGER_MAX_ROWS', store: STORE_TYPES.DOCUMENT, isJson: false }
};

var CACHE_KEYS = {
    LOGS: 'BEAVER_DEBUG_LOGS',
    PROGRESS: '_PROGRESS'
};

var TOOL_LAUNCH_MODES = {
    SIDEBAR: 'SIDEBAR',
    MODAL: 'MODAL'
};

/**
 * Helper to get the appropriate properties store.
 */
function _App_getStore_(storeType) {
    switch (storeType) {
        case STORE_TYPES.DOCUMENT: return PropertiesService.getDocumentProperties();
        case STORE_TYPES.USER: return PropertiesService.getUserProperties();
        case STORE_TYPES.SCRIPT: return PropertiesService.getScriptProperties();
        default: throw new Error("Invalid store type: " + storeType);
    }
}

/**
 * Retrieves a property from the registry. Automatically parses JSON if configured.
 * @param {Object} propConfig An entry from APP_PROPS
 * @returns {*} The value or null if not found
 */
function _App_getProperty(propConfig) {
    var store = _App_getStore_(propConfig.store);
    var valStr = store.getProperty(propConfig.key);
    if (!valStr) return null;

    if (propConfig.isJson) {
        try {
            return JSON.parse(valStr);
        } catch (e) {
            console.warn("Failed to parse JSON for property " + propConfig.key, e);
            return null;
        }
    }
    return valStr;
}

function _App_getRawProperty(propConfig) {
    return _App_getStore_(propConfig.store).getProperty(propConfig.key);
}

/**
 * Sets a property in the registry. Automatically stringifies JSON if configured.
 * @param {Object} propConfig An entry from APP_PROPS
 * @param {*} value The value to set (can be an object or primitive)
 */
function _App_setProperty(propConfig, value) {
    var store = _App_getStore_(propConfig.store);
    var valToStore = propConfig.isJson ? JSON.stringify(value) : String(value);
    store.setProperty(propConfig.key, valToStore);
}

/**
 * Deletes a property from the registry.
 * @param {Object} propConfig An entry from APP_PROPS
 */
function _App_deleteProperty(propConfig) {
    var store = _App_getStore_(propConfig.store);
    store.deleteProperty(propConfig.key);
}

function _App_ok(message, data, meta) {
    return {
        success: true,
        message: message || 'Success',
        data: data || null,
        meta: meta || null
    };
}

function _App_fail(message, data, meta) {
    return {
        success: false,
        message: message || 'Operation failed',
        data: data || null,
        meta: meta || null
    };
}

// Default theme definition
var DEFAULT_SHEET_THEME = {
    // Cell Backgrounds
    HEADER: '#424242',
    ACTION: '#2e5a70',
    EDITABLE: '#528dab',
    READ_ONLY: '#655356',

    // Status Colors (Used for conditional formatting rules)
    STATUS: {
        SUCCESS: '#10B981',    // Emerald Green
        PENDING: '#f59e0b',    // Amber/Yellow
        ERROR: '#EF4444',      // Red
        SYNCED: '#6366F1',     // Indigo
        WARNING: '#d59679'
    },

    // Text Colors
    TEXT: '#ffffff',         // Unified light text color for all backgrounds

    // Borders
    BORDER: '#ffffff',       // Soft gray borders instead of harsh black
    BORDER_STYLE: SpreadsheetApp.BorderStyle.SOLID, // Default border style

    // Typography
    FONTS: {
        PRIMARY: 'Roboto',     // Main font for all sheets
        MONOSPACE: 'Consolas'  // Used for IDs, Paths, and technical data
    },

    SIZES: {
        HEADER: 11,            // Header font size
        BODY: 10               // Data body font size
    },

    // Alignment & Layout
    LAYOUT: {
        HEADER_ALIGN_H: 'center',
        HEADER_ALIGN_V: 'middle',
        BODY_ALIGN_H: 'left',
        BODY_ALIGN_V: 'middle',
        BODY_WRAP: SpreadsheetApp.WrapStrategy.CLIP,
        HEADER_WEIGHT: 'bold',
        HEADER_FONT_STYLE: 'normal',
        HEADER_ROW_HEIGHT: 45,
        BODY_ROW_HEIGHT: 35
    }
};

/**
 * Helper to deep merge objects so we don't lose structure if keys are missing from saved theme.
 */
function deepMergeTheme_(target, source) {
    const output = Object.assign({}, target);
    if (isObject_(target) && isObject_(source)) {
        Object.keys(source).forEach(key => {
            if (isObject_(source[key])) {
                if (!(key in target))
                    Object.assign(output, { [key]: source[key] });
                else
                    output[key] = deepMergeTheme_(target[key], source[key]);
            } else {
                Object.assign(output, { [key]: source[key] });
            }
        });
    }
    return output;
}

function isObject_(item) {
    return (item && typeof item === 'object' && !Array.isArray(item));
}

function _UI_getTheme() {
    const savedTheme = _App_getProperty(APP_PROPS.THEME);
    if (savedTheme) {
        try {
            const merged = deepMergeTheme_(DEFAULT_SHEET_THEME, savedTheme);

            // Restore Enum properties that get converted to strings or empty objects during JSON serialization
            const bw = merged.LAYOUT.BODY_WRAP;
            if (typeof bw === 'string') {
                merged.LAYOUT.BODY_WRAP = SpreadsheetApp.WrapStrategy[bw] || SpreadsheetApp.WrapStrategy.CLIP;
            } else if (bw !== SpreadsheetApp.WrapStrategy.CLIP && bw !== SpreadsheetApp.WrapStrategy.WRAP && bw !== SpreadsheetApp.WrapStrategy.OVERFLOW) {
                merged.LAYOUT.BODY_WRAP = SpreadsheetApp.WrapStrategy.CLIP;
            }

            const bs = merged.BORDER_STYLE;
            if (typeof bs === 'string') {
                merged.BORDER_STYLE = SpreadsheetApp.BorderStyle[bs] || SpreadsheetApp.BorderStyle.SOLID;
            } else if (bs !== SpreadsheetApp.BorderStyle.SOLID && bs !== SpreadsheetApp.BorderStyle.SOLID_MEDIUM && bs !== SpreadsheetApp.BorderStyle.SOLID_THICK && bs !== SpreadsheetApp.BorderStyle.DASHED && bs !== SpreadsheetApp.BorderStyle.DOTTED && bs !== SpreadsheetApp.BorderStyle.DOUBLE) {
                merged.BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID;
            }

            return merged;
        } catch (e) {
            console.error('Failed to parse saved theme, falling back to defaults:', e);
        }
    }
    return DEFAULT_SHEET_THEME;
}

function _App_withDocumentLock(lockName, callback, timeoutMs) {
    var lock = LockService.getDocumentLock();
    var waitMs = timeoutMs || 30000;

    if (!lock.tryLock(waitMs)) {
        throw new Error('System is busy with another operation' + (lockName ? ' (' + lockName + ')' : '') + '. Please try again.');
    }

    try {
        return callback();
    } finally {
        lock.releaseLock();
    }
}

function _App_canScaffoldSheet(toolConfig) {
    return !!(toolConfig && toolConfig.HEADERS && toolConfig.HEADERS.length);
}

// Global Export! Scripts using SHEET_THEME will get the dynamic version lazily.
// We use a Proxy here so that PropertiesService (a slow API call) is only invoked
// when a script actively accesses the theme, preventing execution delays on all triggers.
var __ui_sheetThemeCache = null;
var SHEET_THEME = new Proxy({}, {
    get: function (target, prop) {
        if (!__ui_sheetThemeCache) {
            __ui_sheetThemeCache = _UI_getTheme(); // Load from PropertiesService only on access
        }
        return Reflect.get(__ui_sheetThemeCache, prop);
    },
    ownKeys: function () {
        if (!__ui_sheetThemeCache) __ui_sheetThemeCache = _UI_getTheme();
        return Reflect.ownKeys(__ui_sheetThemeCache);
    },
    getOwnPropertyDescriptor: function (target, prop) {
        if (!__ui_sheetThemeCache) __ui_sheetThemeCache = _UI_getTheme();
        return Reflect.getOwnPropertyDescriptor(__ui_sheetThemeCache, prop);
    }
});

/**
 * Throws an error if the active sheet is not the expected one. Useful for direct action trigger functions.
 * @param {string} expectedSheetName The globally defined sheet name from SHEET_NAMES
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet object if valid
 */
function _App_assertActiveSheet(expectedSheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== expectedSheetName) {
        throw new Error("⚠️ Please run this action from the '" + expectedSheetName + "' sheet.");
    }
    return sheet;
}

/**
 * Returns a validation object. Useful for UI-triggered functions that need to return an error shape `{success: false, message: ...}` instead of failing ungracefully.
 * @param {string} expectedSheetName 
 * @returns {Object} `{ valid: boolean, sheet: Sheet, message?: string }`
 */
function _App_validateActiveSheet(expectedSheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== expectedSheetName) {
        return { valid: false, message: "⚠️ Please run this action from the '" + expectedSheetName + "' sheet." };
    }
    return { valid: true, sheet: sheet };
}

function _App_include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); }
};

// ==========================================
// Centralized Body Formatting Utility
// ==========================================

// Extra rows formatted beyond actual data to cover manual row additions.
var FORMATTING_BUFFER_ROWS = 30;

function _App_getColumnLetter(col) {
    var temp, letter = '';
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}

function _App_applyHeaderFormatting(sheet, headers) {
    if (!headers || headers.length === 0) return;

    sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight(SHEET_THEME.LAYOUT.HEADER_WEIGHT)
        .setFontSize(SHEET_THEME.SIZES.HEADER)
        .setFontFamily(SHEET_THEME.FONTS.PRIMARY)
        .setBackground(SHEET_THEME.HEADER)
        .setFontColor(SHEET_THEME.TEXT)
        .setFontStyle(SHEET_THEME.LAYOUT.HEADER_FONT_STYLE)
        .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE)
        .setVerticalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_V)
        .setHorizontalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_H);
    sheet.setRowHeight(1, SHEET_THEME.LAYOUT.HEADER_ROW_HEIGHT);
}

function _App_cloneFormatConfig_(config) {
    if (!config) return null;

    var clone = {};
    Object.keys(config).forEach(function(key) {
        var value = config[key];
        if (key === 'COL_SCHEMA' || key === 'conditionalRules') {
            clone[key] = (value || []).map(function(item) {
                var out = {};
                Object.keys(item).forEach(function(itemKey) {
                    out[itemKey] = item[itemKey];
                });
                return out;
            });
        } else {
            clone[key] = value;
        }
    });
    return clone;
}

function _App_buildRuntimeToolShape(toolKey, dynamicHeaders, options) {
    var cfg = BeaverEngine.getTool(toolKey);
    var runtimeHeaders = (cfg.HEADERS || []).slice();
    var runtimeWidths = (cfg.COL_WIDTHS || []).slice();
    var runtimeFormat = _App_cloneFormatConfig_(cfg.FORMAT_CONFIG);
    var headersToInsert = dynamicHeaders || [];
    var dynamicSchemaFactory = options && options.dynamicSchemaFactory;
    var dynamicColWidth = options && options.dynamicColWidth !== undefined ? options.dynamicColWidth : 150;
    var anchorHeader = options && options.anchorHeader;
    var insertIndex = runtimeHeaders.length;

    if (anchorHeader) {
        insertIndex = runtimeHeaders.indexOf(anchorHeader);
        if (insertIndex === -1) {
            throw new Error("Anchor header '" + anchorHeader + "' was not found for tool '" + toolKey + "'.");
        }
    }

    var schemaItems = headersToInsert.map(function(header) {
        if (typeof dynamicSchemaFactory === 'function') {
            return dynamicSchemaFactory(header);
        }
        return { header: header, type: 'TEXT' };
    });

    if (headersToInsert.length > 0) {
        Array.prototype.splice.apply(runtimeHeaders, [insertIndex, 0].concat(headersToInsert));
        Array.prototype.splice.apply(runtimeWidths, [insertIndex, 0].concat(headersToInsert.map(function() { return dynamicColWidth; })));
        if (runtimeFormat && runtimeFormat.COL_SCHEMA) {
            Array.prototype.splice.apply(runtimeFormat.COL_SCHEMA, [insertIndex, 0].concat(schemaItems));
            runtimeFormat.totalCols = runtimeFormat.COL_SCHEMA.length;
        }
    }

    return {
        headers: runtimeHeaders,
        widths: runtimeWidths,
        formatConfig: runtimeFormat
    };
}

function _App_syncDynamicColumns(toolKey, dynamicHeaders, options) {
    var cfg = BeaverEngine.getTool(toolKey);
    var uniqueDynamicHeaders = [];
    (dynamicHeaders || []).forEach(function(header) {
        var normalized = String(header || '').trim();
        if (normalized && uniqueDynamicHeaders.indexOf(normalized) === -1) {
            uniqueDynamicHeaders.push(normalized);
        }
    });

    var sheet = _App_ensureSheetExists(toolKey);
    var runtimeShape = _App_buildRuntimeToolShape(toolKey, uniqueDynamicHeaders, options);
    var currentHeaderCount = sheet.getLastColumn();

    if (currentHeaderCount > runtimeShape.headers.length) {
        sheet.deleteColumns(runtimeShape.headers.length + 1, currentHeaderCount - runtimeShape.headers.length);
    } else if (currentHeaderCount < runtimeShape.headers.length) {
        sheet.insertColumnsAfter(Math.max(currentHeaderCount, 1), runtimeShape.headers.length - currentHeaderCount);
    }

    _App_applyHeaderFormatting(sheet, runtimeShape.headers);

    runtimeShape.widths.forEach(function(width, index) {
        if (width !== null && width !== undefined) {
            sheet.setColumnWidth(index + 1, width);
        }
    });

    if (cfg.FROZEN_ROWS > 0) sheet.setFrozenRows(cfg.FROZEN_ROWS);
    if (cfg.FROZEN_COLS > 0) sheet.setFrozenColumns(cfg.FROZEN_COLS);

    if (runtimeShape.formatConfig) {
        var numRows = Math.max(sheet.getLastRow() - 1, 0);
        _App_applyBodyFormatting(sheet, numRows, runtimeShape.formatConfig);
    }

    return {
        headers: runtimeShape.headers,
        dynamicHeaders: uniqueDynamicHeaders,
        sheet: sheet,
        formatConfig: runtimeShape.formatConfig
    };
}

/**
 * Applies standardized body formatting to a sheet's data area.
 * This enforces strict column ordering:
 * 1: Action (SHEET_THEME.ACTION) - unless config.skipActionColoring is true
 * 2 onwards: Editable data (SHEET_THEME.EDITABLE)
 * Last N columns: Read-only data (SHEET_THEME.READ_ONLY)
 */
function _App_applyBodyFormatting(sheet, numDataRows, config) {
    var rowsToFormat = numDataRows + FORMATTING_BUFFER_ROWS;
    var maxRows = sheet.getMaxRows();
    var actualRows = Math.min(rowsToFormat, maxRows - 1);
    if (actualRows < 1) return;

    var totalCols = config.COL_SCHEMA ? config.COL_SCHEMA.length : (config.totalCols || sheet.getLastColumn());
    var numReadOnlyAtEnd = config.numReadOnlyColsAtEnd || 0;

    // 1. Base formatting
    var startRow = 2;
    var endCol = Math.max(totalCols, 1);

    var dataRange = sheet.getRange(startRow, 1, actualRows, totalCols);
    dataRange
        .setFontColor(SHEET_THEME.TEXT)
        .setFontFamily(SHEET_THEME.FONTS.PRIMARY)
        .setFontSize(SHEET_THEME.SIZES.BODY)
        .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE)
        .setHorizontalAlignment(SHEET_THEME.LAYOUT.BODY_ALIGN_H)
        .setVerticalAlignment(SHEET_THEME.LAYOUT.BODY_ALIGN_V)
        .setWrapStrategy(SHEET_THEME.LAYOUT.BODY_WRAP);

    sheet.setRowHeights(startRow, actualRows, SHEET_THEME.LAYOUT.BODY_ROW_HEIGHT);

    try {
        var startColForEditable = 1;

        // 1. Action (Col 1)
        if (!config.skipActionColoring) {
            sheet.getRange(startRow, 1, actualRows, 1).setBackground(SHEET_THEME.ACTION);
            startColForEditable = 2;
        }

        // 2. Editable Columns
        var numEditable = endCol - (startColForEditable - 1) - numReadOnlyAtEnd;
        if (numEditable > 0) {
            sheet.getRange(startRow, startColForEditable, actualRows, numEditable).setBackground(SHEET_THEME.EDITABLE);
        }

        // 3. Read-Only Columns
        if (numReadOnlyAtEnd > 0) {
            var readOnlyStartCol = endCol - numReadOnlyAtEnd + 1;
            sheet.getRange(startRow, readOnlyStartCol, actualRows, numReadOnlyAtEnd).setBackground(SHEET_THEME.READ_ONLY);
        }
    } catch (e) {
        console.error("Error applying column colors:", e);
    }

    // Apply Schema-driven validations and formats
    var validationRows = maxRows - 1;
    if (config.COL_SCHEMA) {
        config.COL_SCHEMA.forEach(function(colDef, index) {
            var colNum = index + 1;
            var range = sheet.getRange(startRow, colNum, actualRows, 1);
            var valRange = sheet.getRange(startRow, colNum, validationRows, 1);
            
            // Fonts
            if (colDef.type === 'ID' || colDef.type === 'URL') {
                range.setFontFamily(SHEET_THEME.FONTS.MONOSPACE);
            }
            if (colDef.type === 'URL' || colDef.italic) {
                range.setFontStyle('italic');
            }

            // Number Formats
            if (colDef.type === 'DATETIME') {
                range.setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');
            } else if (colDef.type === 'DATE') {
                range.setNumberFormat('MM/dd/yyyy');
            } else if (colDef.type === 'ID' || colDef.type === 'TEXT') {
                range.setNumberFormat('@'); // Force Plain Text
            }

            // Validations
            var rule = null;
            if (colDef.type === 'ACTION' || colDef.type === 'DROPDOWN') {
                var opts = typeof colDef.options === 'function' ? colDef.options() : colDef.options;
                if (opts && opts.length > 0) {
                    rule = SpreadsheetApp.newDataValidation().requireValueInList(opts, true).setAllowInvalid(colDef.allowInvalid || false).build();
                }
            } else if (colDef.type === 'CHECKBOX') {
                rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
            } else if (colDef.type === 'EMAIL' || colDef.type === 'EMAIL_LIST') {
                var letter = _App_getColumnLetter(colNum);
                var re = colDef.type === 'EMAIL' ? 'ISEMAIL(' + letter + '2)' : 'REGEXMATCH(' + letter + '2, "^[\\\\w\\\\.\\\\-@\\\\s,]+$")';
                var formula = '=OR(ISBLANK(' + letter + '2), ' + re + ')';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter valid email(s).').setAllowInvalid(true).build();
            }

            if (rule) {
                valRange.setDataValidation(rule);
            }
        });
    }

    // 6. Conditional formatting rules
    _App_applyConditionalRules(sheet, actualRows, totalCols, config.conditionalRules || []);
}

/**
 * Builds and applies conditional formatting rules from a declarative descriptor array.
 * Replaces ALL existing conditional formatting rules on the sheet.
 *
 * Supported rule types: 'success', 'error', 'errorCross', 'pending', 'synced', 'custom'
 * Supported scopes: 'fullRow' (default), 'actionOnly', 'statusOnly'
 */
function _App_applyConditionalRules(sheet, numRows, totalCols, ruleDescriptors) {
    var rules = [];
    var fullRange = sheet.getRange(2, 1, numRows, totalCols);

    ruleDescriptors.forEach(function (desc) {
        var targetRange;
        if (desc.scope === 'actionOnly' && desc.actionCol) {
            var actionColNum = desc.actionCol.charCodeAt(0) - 64; // 'A' → 1
            targetRange = sheet.getRange(2, actionColNum, numRows, 1);
        } else if (desc.scope === 'statusOnly' && desc.statusCol) {
            var statusColNum = desc.statusCol.charCodeAt(0) - 64;
            targetRange = sheet.getRange(2, statusColNum, numRows, 1);
        } else if (desc.scope === 'custom_col' && desc.col) {
            targetRange = sheet.getRange(2, desc.col, numRows, 1);
        } else {
            targetRange = fullRange; // 'fullRow'
        }

        var rule;
        if (desc.type === 'success') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "✅")')
                .setBackground(SHEET_THEME.STATUS.SUCCESS)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'error') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "⚠️")')
                .setBackground(SHEET_THEME.STATUS.ERROR)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'errorCross') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "❌")')
                .setBackground(SHEET_THEME.STATUS.ERROR)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'pending') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=$' + desc.actionCol + '2<>""')
                .setBackground(SHEET_THEME.STATUS.PENDING)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'synced') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "📝")')
                .setBackground(SHEET_THEME.STATUS.SYNCED)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'custom' && desc.formula) {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(desc.formula)
                .setBackground(desc.color)
                .setRanges([targetRange]).build();
        }

        if (rule) rules.push(rule);
    });

    sheet.setConditionalFormatRules(rules);
}

// ==========================================
// BeaverEngine — Plugin Registration System
// ==========================================

var BeaverEngine = (function() {
    var registry = {};

    function _validateToolConfig(key, config) {
        var issues = [];

        if (!config.SHEET_NAME) issues.push("Missing SHEET_NAME.");
        if (!config.TITLE) issues.push("Missing TITLE.");

        if (config.MENU_LABEL && !config.MENU_ENTRYPOINT) {
            issues.push("MENU_LABEL requires MENU_ENTRYPOINT.");
        }

        if (config.LAUNCH_MODE === TOOL_LAUNCH_MODES.SIDEBAR && !config.SIDEBAR_HTML && config.MENU_ENTRYPOINT) {
            issues.push("Sidebar tools require SIDEBAR_HTML.");
        }

        if (config.LAUNCH_MODE === TOOL_LAUNCH_MODES.MODAL && !(config.MODAL_HTML || config.SIDEBAR_HTML) && config.MENU_ENTRYPOINT) {
            issues.push("Modal tools require MODAL_HTML or SIDEBAR_HTML.");
        }

        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA && !Array.isArray(config.FORMAT_CONFIG.COL_SCHEMA)) {
            issues.push("FORMAT_CONFIG.COL_SCHEMA must be an array.");
        }

        return issues;
    }

    /**
     * Registers a tool with the engine.
     * Automatically processes COL_SCHEMA to generate HEADERS and totalCols.
     */
    function registerTool(key, config) {
        config.TOOL_KEY = key;
        config.MENU_LABEL = config.MENU_LABEL || config.TITLE;
        config.MENU_ORDER = typeof config.MENU_ORDER === 'number' ? config.MENU_ORDER : 999;
        config.LAUNCH_MODE = config.LAUNCH_MODE || TOOL_LAUNCH_MODES.SIDEBAR;

        // Post-process the config (generate HEADERS and totalCols from SCHEMA)
        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA) {
            config.HEADERS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(c) { return c.header; });
            config.FORMAT_CONFIG.totalCols = config.FORMAT_CONFIG.COL_SCHEMA.length;
        }

        var issues = _validateToolConfig(key, config);
        if (issues.length > 0) {
            throw new Error("Tool '" + key + "' is misconfigured: " + issues.join(' '));
        }

        registry[key] = config;
        // console.log("BeaverEngine: Registered " + key);
    }

    /**
     * Retrieves a tool configuration by key.
     */
    function getTool(key) {
        var cfg = registry[key];
        if (!cfg) throw new Error('Unknown tool key: "' + key + '". Ensure the tool is registered via BeaverEngine.registerTool().');
        return cfg;
    }

    /**
     * Returns all registered tools.
     */
    function getAllTools() {
        return registry;
    }

    function getToolKeys() {
        return Object.keys(registry);
    }

    function auditTool(key) {
        var cfg = getTool(key);
        return _validateToolConfig(key, cfg);
    }

    return {
        registerTool: registerTool,
        getTool: getTool,
        getAllTools: getAllTools,
        getToolKeys: getToolKeys,
        auditTool: auditTool
    };
})();

/**
 * Backward compatibility Proxy for legacy scripts still referencing APP_REGISTRY directly.
 */
var APP_REGISTRY = new Proxy({}, {
    get: function(target, prop) {
        return BeaverEngine.getTool(prop);
    },
    ownKeys: function() {
        return Object.keys(BeaverEngine.getAllTools());
    },
    getOwnPropertyDescriptor: function(target, prop) {
        return { enumerable: true, configurable: true };
    }
});

// ==========================================
// _App_openSidebar — Universal Sidebar Opener
// ==========================================
/**
 * Opens a tool's sidebar, ensuring the sheet exists first.
 */
function _App_openSidebar(toolKey, postCreateCallback) {
    var cfg = BeaverEngine.getTool(toolKey);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.SHEET_NAME);

    if (!sheet) {
        sheet = _App_ensureSheetExists(toolKey, postCreateCallback);
    } else {
        sheet.activate();
    }

    var html = HtmlService.createHtmlOutputFromFile(cfg.SIDEBAR_HTML)
        .setTitle(cfg.TITLE)
        .setWidth(cfg.SIDEBAR_WIDTH || 300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function _App_launchTool(toolKey, postCreateCallback) {
    var cfg = BeaverEngine.getTool(toolKey);
    var launchMode = cfg.LAUNCH_MODE || TOOL_LAUNCH_MODES.SIDEBAR;

    if (launchMode === TOOL_LAUNCH_MODES.MODAL) {
        var html = HtmlService.createHtmlOutputFromFile(cfg.MODAL_HTML || cfg.SIDEBAR_HTML)
            .setTitle(cfg.TITLE)
            .setWidth(cfg.MODAL_WIDTH || cfg.SIDEBAR_WIDTH || 300)
            .setHeight(cfg.MODAL_HEIGHT || 600);
        SpreadsheetApp.getUi().showModalDialog(html, cfg.TITLE);
        return;
    }

    _App_openSidebar(toolKey, postCreateCallback);
}

function _App_getMenuTools() {
    return Object.keys(BeaverEngine.getAllTools())
        .map(function(key) { return BeaverEngine.getTool(key); })
        .filter(function(cfg) { return !!cfg.MENU_LABEL; })
        .sort(function(a, b) {
            if (a.MENU_ORDER !== b.MENU_ORDER) return a.MENU_ORDER - b.MENU_ORDER;
            return String(a.MENU_LABEL).localeCompare(String(b.MENU_LABEL));
        });
}

// ==========================================
// _App_ensureSheetExists — Universal Sheet Scaffolding
// ==========================================
/**
 * Creates a tool sheet if it doesn't exist, with headers, column widths,
 * frozen rows/cols, data validations, and buffer body formatting.
 */
function _App_ensureSheetExists(toolKey, postCreateCallback) {
    var cfg = BeaverEngine.getTool(toolKey);
    if (!_App_canScaffoldSheet(cfg)) {
        throw new Error("Tool '" + toolKey + "' does not define a sheet schema and cannot be scaffolded automatically.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.SHEET_NAME);
    var isNew = !sheet;

    if (isNew) {
        sheet = ss.insertSheet(cfg.SHEET_NAME);
    }

    // Always ensure headers and basic sheet setup are correct (idempotent)
    _App_applyHeaderFormatting(sheet, cfg.HEADERS);

    if (cfg.FROZEN_ROWS > 0) sheet.setFrozenRows(cfg.FROZEN_ROWS);
    if (cfg.FROZEN_COLS > 0) sheet.setFrozenColumns(cfg.FROZEN_COLS);

    if (cfg.COL_WIDTHS) {
        cfg.COL_WIDTHS.forEach(function (w, i) {
            if (w !== null && w !== undefined) sheet.setColumnWidth(i + 1, w);
        });
    }

    if (cfg.FORMAT_CONFIG) {
        _App_applyBodyFormatting(sheet, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0, cfg.FORMAT_CONFIG);
    }

    if (isNew && typeof postCreateCallback === 'function') {
        try { postCreateCallback(sheet); }
        catch (e) { console.warn('[_App_ensureSheetExists] Post-create callback failed (' + toolKey + '): ' + e.message); }
    }

    sheet.activate();
    return sheet;
}

// ==========================================
// Progress Tracking — Unified CacheService Wrappers
// ==========================================

/**
 * Stores batch operation progress for sidebar polling.
 * @param {string} toolName     - Tool key e.g. 'MAIL_SENDER'
 * @param {number} current      - Items processed so far
 * @param {number} total        - Total items queued
 * @param {number} [ttlSec=600] - Cache TTL in seconds (default 10 min)
 */
function _App_setProgress(toolName, current, total, ttlSec) {
    CacheService.getUserCache().put(
        toolName + CACHE_KEYS.PROGRESS,
        JSON.stringify({ current: current, total: total }),
        ttlSec || 600
    );
}

/**
 * Returns cached progress or null if expired/not set.
 * @param {string} toolName
 * @returns {{ current: number, total: number }|null}
 */
function _App_getProgress(toolName) {
    var data = CacheService.getUserCache().get(toolName + CACHE_KEYS.PROGRESS);
    return data ? JSON.parse(data) : null;
}

/**
 * Removes progress state after an operation completes.
 * @param {string} toolName
 */
function _App_clearProgress(toolName) {
    CacheService.getUserCache().remove(toolName + CACHE_KEYS.PROGRESS);
}

// ==========================================
// _App_throttle — Unified API Rate Limiter
// ==========================================
/**
 * Tracks cumulative API calls and sleeps (1 s) whenever a multiple of the
 * limit is crossed — preventing Google 429 "Too Many Requests" errors.
 */
function _App_throttle(tracker, callsMade, limit) {
    var _limit = limit || 10;
    var prev = tracker.apiCalls;
    tracker.apiCalls += callsMade;
    if (Math.floor(tracker.apiCalls / _limit) > Math.floor(prev / _limit)) {
        Utilities.sleep(1000);
    }
}

// ==========================================
// _App_callWithBackoff — Unified Exponential Backoff Retry
// ==========================================
/**
 * Runs a function; retries with exponential backoff on transient Google API errors.
 */
function _App_callWithBackoff(func, retries) {
    var maxRetries = (retries !== undefined) ? retries : 5;
    for (var n = 0; n <= maxRetries; n++) {
        try {
            return func();
        } catch (e) {
            var msg = (e.message || '').toLowerCase();
            var isRetriable = (
                msg.indexOf('403') !== -1 || msg.indexOf('429') !== -1 ||
                msg.indexOf('500') !== -1 || msg.indexOf('502') !== -1 ||
                msg.indexOf('503') !== -1 ||
                msg.indexOf('rate limit') !== -1 || msg.indexOf('quota') !== -1 ||
                msg.indexOf('limit exceeded') !== -1 || msg.indexOf('too many') !== -1
            );
            if (isRetriable && n < maxRetries) {
                var waitMs = (Math.pow(2, n) * 1000) + Math.round(Math.random() * 1000);
                console.warn('[_App_callWithBackoff] Retry ' + (n + 1) + '/' + maxRetries + ' in ' + waitMs + 'ms — ' + e.message);
                Utilities.sleep(waitMs);
            } else {
                throw e;
            }
        }
    }
}

// ==========================================
// _App_logClientError / _App_logClientInfo
// ==========================================
/**
 * Logs a client-side error from a sidebar (Unified for all tools).
 * @param {string|Object} err Error message or object
 * @param {string} [context] Optional context (e.g. 'Button Click')
 */
function _App_logClientError(err, context) {
    var source = 'Client UI';
    var ref = context || 'Default';
    var msg = (typeof err === 'object' && err.message) ? err.message : String(err);
    if (typeof err === 'object' && err.stack) msg += '\nStack: ' + err.stack;
    
    // Only attempt to log if Logger framework is loaded
    if (typeof Logger !== 'undefined' && typeof Logger.error === 'function') {
        Logger.error(source, ref, msg);
    } else {
        console.error("[Client Error] " + ref + ": " + msg);
    }
}

/**
 * Logs a client-side info message from a sidebar (Unified for all tools).
 * @param {string} message 
 * @param {string} [context]
 */
function _App_logClientInfo(message, context) {
    if (typeof Logger !== 'undefined' && typeof Logger.info === 'function') {
        Logger.info('Client UI', context || 'Default', message);
    } else {
        console.log("[Client Info] " + (context || "UI") + ": " + message);
    }
}
