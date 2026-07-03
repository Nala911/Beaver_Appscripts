// =========================================================
// BUNDLED GOOGLE APPS SCRIPT CODE (AUTOMATICALLY GENERATED)
// =========================================================


// --- FILE: core/00_Logger.js ---
/**
 * Developer Logging System
 * Version: 7.0 (Silent Architecture)
 */

var Logger = (function () {
    var isLoggingEnabled = false;

    return {
        setLoggingState: function (enabled) {
            isLoggingEnabled = !!enabled;
        },
        setRunId: function (id) { return id; },
        info: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.info("[" + src + "::" + ref + "] " + msg, ctx || "");
        },
        success: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.log("✅ [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        warn: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.warn("⚠️ [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        debug: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.log("🐞 [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        error: function (src, ref, err, ctx) {
            var errMsg = err ? (err.message || String(err)) : "Unknown Error";
            console.error("❌ [" + src + "::" + ref + "] " + errMsg, ctx || "", err && err.stack ? err.stack : "");
        },
        step: function (src, ref, name) {
            if (isLoggingEnabled) console.log("➔ [" + src + "::" + ref + "] Step: " + name);
        },
        flushLogs: function () {},
        clearLogs: function () {},
        isEnabled: function() { return isLoggingEnabled; },

        run: function (toolKey, reference, callback, forceLog) {
            var oldState = isLoggingEnabled;
            if (forceLog) isLoggingEnabled = true;
            try {
                if (isLoggingEnabled) console.log("➔ [" + toolKey + "] Starting execution: " + reference);
                var res = callback();
                if (isLoggingEnabled) console.log("➔ [" + toolKey + "] Finished execution: " + reference);
                return res;
            } catch (e) {
                console.error("❌ [" + toolKey + "] Failed execution: " + reference, e);
                throw e; 
            } finally {
                isLoggingEnabled = oldState;
            }
        },

        wrap: function (source, reference, func) {
            return function() {
                var oldState = isLoggingEnabled;
                try {
                    return func.apply(this, arguments);
                } catch(e) {
                    console.error("❌ [" + source + "] Function error: " + reference, e);
                    throw e;
                } finally {
                    isLoggingEnabled = oldState;
                }
            };
        }
    };
})();



// --- FILE: core/01_Config_Constants.js ---
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
    FORMS_SYNC: '📝 Google Forms',
    BULK_FOLDER: '📂 Bulk Folder Creation',
    DRIVE_SYNC: '💾 Google Drive',
    PIPELINE: '⛓  Pipeline',
    CHAT_SYNC: '💬 Google Chat Spaces',
    GMAIL_FILTERS: '🗂️ Gmail Filters',
    TASKS_SYNC: '📋 Google Tasks'
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
    // Pipeline
    SYSTEM_ENABLED: { key: 'SYSTEM_ENABLED', store: STORE_TYPES.SCRIPT, isJson: false, validate: 'BOOLEAN' },

    // Docs Merge
    DOCS_MERGE_TEMPLATE_URL: { key: 'DOCS_MERGE_TEMPLATE_URL', store: STORE_TYPES.DOCUMENT, isJson: false, validate: 'DOCS_URL' },
    DOCS_MERGE_FOLDER_URL: { key: 'DOCS_MERGE_FOLDER_URL', store: STORE_TYPES.DOCUMENT, isJson: false, validate: 'DRIVE_URL' },
    DOCS_MERGE_TEMPLATE_NAME: { key: 'DOCS_MERGE_TEMPLATE_NAME', store: STORE_TYPES.DOCUMENT, isJson: false },
    DOCS_MERGE_FOLDER_NAME: { key: 'DOCS_MERGE_FOLDER_NAME', store: STORE_TYPES.DOCUMENT, isJson: false },
    DOCS_MERGE_MASTER_DOC_ID: { key: 'DOCS_MERGE_MASTER_DOC_ID', store: STORE_TYPES.DOCUMENT, isJson: false },

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

    // Tasks Sync
    TASKS_SELECTED_LIST: { key: 'selectedTasksList', store: STORE_TYPES.USER, isJson: false }

};

var CACHE_KEYS = {
    PROGRESS: '_PROGRESS'
};

var TOOL_LAUNCH_MODES = {
    SIDEBAR: 'SIDEBAR',
    MODAL: 'MODAL'
};

var DEFAULT_COL_WIDTHS = {
    ACTION: 120,
    STATUS: 200,
    ID: 150,
    URL: 250,
    DATETIME: 180,
    TEXT: 200,
    DROPDOWN: 150,
    CHECKBOX: 100,
    EMAIL: 200,
    EMAIL_LIST: 250,
    DATE: 120,
    READ_ONLY: 200
};
// Trigger clasp push refresh


// --- FILE: core/02_Config_Theme.js ---
// Default theme definition
var DEFAULT_SHEET_THEME = {
    // Cell Backgrounds
    HEADER: '#424242',
    FIRST_COLS_COLOR: '#2e5a70',
    MIDDLE_COLS_COLOR: '#528dab',
    LAST_COLS_COLOR: '#314974',

    // Status Colors (Used for conditional formatting rules)
    STATUS: {
        SUCCESS: '#10B981',    // Emerald Green
        PENDING: '#f59e0b',    // Amber/Yellow
        ERROR: '#EF4444',      // Red
        SYNCED: '#6366F1',     // Indigo
        WARNING: '#d59679'
    },

    // Standard Status Prefixes
    STATUS_PREFIXES: {
        SUCCESS: '✅ ',
        ERROR: '❌ ',
        WARNING: '⚠️ ',
        PENDING: '⏳ ',
        INFO: 'ℹ️ '
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

var SHEET_THEME = DEFAULT_SHEET_THEME;


// --- FILE: core/03_Config_Storage.js ---
var __propsCache = {};

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
 * Helper to get the appropriate cache store.
 */
function _App_getCacheStore_(storeType) {
    switch (storeType) {
        case STORE_TYPES.DOCUMENT: return CacheService.getDocumentCache() || CacheService.getScriptCache();
        case STORE_TYPES.USER: return CacheService.getUserCache();
        case STORE_TYPES.SCRIPT: return CacheService.getScriptCache();
        default: return null;
    }
}

/**
 * Retrieves a property from the registry. Automatically parses JSON if configured.
 * @param {Object} propConfig An entry from APP_PROPS
 * @returns {*} The value or null if not found
 */
function _App_getProperty(propConfig) {
    var cacheKey = propConfig.key;
    
    // 1. Fast Memory Cache
    if (__propsCache.hasOwnProperty(cacheKey)) {
        return __propsCache[cacheKey];
    }

    // 2. CacheService Layer
    var cacheStore = _App_getCacheStore_(propConfig.store);
    var valStr = cacheStore ? cacheStore.get(cacheKey) : null;
    
    // 3. PropertiesService Fallback
    if (valStr === null) {
        var store = _App_getStore_(propConfig.store);
        valStr = store.getProperty(cacheKey);
        if (valStr && cacheStore) {
            cacheStore.put(cacheKey, valStr, 21600); // Max 6 hours
        }
    }

    if (!valStr) return null;

    var result = valStr;
    if (propConfig.isJson) {
        try {
            result = JSON.parse(valStr);
        } catch (e) {
            return null;
        }
    }

    // Save to memory cache for subsequent calls
    __propsCache[cacheKey] = result;
    return result;
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
    if (propConfig.validate && typeof SYSTEM_VALIDATORS !== 'undefined' && SYSTEM_VALIDATORS) {
        var validator = SYSTEM_VALIDATORS[propConfig.validate];
        if (validator && typeof validator === 'function') {
            if (!validator(value)) {
                throw new Error("Validation failed: Value '" + value + "' is invalid for " + propConfig.key + " (expected: " + propConfig.validate + ")");
            }
        }
    }
    var valToStore = propConfig.isJson ? JSON.stringify(value) : String(value);
    
    // Save to DB
    var store = _App_getStore_(propConfig.store);
    store.setProperty(propConfig.key, valToStore);
    
    // Update Caches
    __propsCache[propConfig.key] = value;
    var cacheStore = _App_getCacheStore_(propConfig.store);
    if (cacheStore) {
        cacheStore.put(propConfig.key, valToStore, 21600);
    }
}

/**
 * Deletes a property from the registry.
 * @param {Object} propConfig An entry from APP_PROPS
 */
function _App_deleteProperty(propConfig) {
    // Delete from DB
    var store = _App_getStore_(propConfig.store);
    store.deleteProperty(propConfig.key);
    
    // Clear Caches
    delete __propsCache[propConfig.key];
    var cacheStore = _App_getCacheStore_(propConfig.store);
    if (cacheStore) {
        cacheStore.remove(propConfig.key);
    }
}


// --- FILE: core/04_SheetManager.js ---
// ==========================================
// Centralized Sheet Manager (DAO Pattern)
// ==========================================

var SheetManager = (function() {

    var _headersCache = {};

    function _normalizeHeaderKey(header) {
        return String(header || '').toUpperCase().trim();
    }

    /**
     * Retrieves the sheet for a given toolKey from APP_REGISTRY.
     * Throws an error if the toolKey or sheet does not exist.
     */
    function getSheet(toolKey) {
        var cfg = SyncEngine.getTool(toolKey);
        if (!cfg) throw new Error("SheetManager: Unknown toolKey '" + toolKey + "'");
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.SHEET_NAME);
        if (!sheet) throw new Error("SheetManager: Sheet '" + cfg.SHEET_NAME + "' not found.");
        return sheet;
    }

    function ensureSheet(toolKey) {
        return _App_ensureSheetExists(toolKey);
    }

    /**
     * Returns the headers for a tool. 
     * Prioritizes the actual sheet headers to support dynamic columns, 
     * falls back to SyncEngine metadata if sheet is empty or missing.
     */
    function getHeaders(toolKey) {
        if (_headersCache[toolKey]) return _headersCache[toolKey];

        try {
            var cfg = SyncEngine.getTool(toolKey);
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var sheet = ss.getSheetByName(cfg.SHEET_NAME);
            if (sheet) {
                var lastCol = sheet.getLastColumn();
                if (lastCol > 0) {
                    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
                    // Filter out empty trailing headers
                    while (headers.length > 0 && (!headers[headers.length - 1] || headers[headers.length - 1] === "")) {
                        headers.pop();
                    }
                    if (headers.length > 0) {
                        _headersCache[toolKey] = headers;
                        return headers;
                    }
                }
            }
        } catch (e) {
            // Silently fallback
        }

        var cfg = SyncEngine.getTool(toolKey);
        return cfg.HEADERS || [];
    }

    function getHeaderMap(toolKey) {
        var headers = getHeaders(toolKey);
        var map = {};
        headers.forEach(function(header, index) {
            if (header) map[header] = index + 1;
        });
        return map;
    }

    function getNormalizedHeaderMap(toolKey) {
        var headers = getHeaders(toolKey);
        var map = {};
        headers.forEach(function(header, index) {
            if (header) map[_normalizeHeaderKey(header)] = index;
        });
        return map;
    }

    function getSheetHeaderMap(sheet) {
        var lastCol = sheet.getLastColumn();
        var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
        var map = {};
        headers.forEach(function(header, index) {
            if (header) map[_normalizeHeaderKey(header)] = index;
        });
        return map;
    }

    /**
     * Reads all data rows (row 2 onwards) and maps them to an array of objects
     * using the headers defined in the tool configuration or sheet.
     * @returns {Object[]} Array of row objects mapped by header names
     */
    function readObjects(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        var headers = getHeaders(toolKey);
        if (headers.length === 0) return [];

        var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
        var data = dataRange.getValues();

        return data.map(function(row) {
            var obj = {};
            for (var i = 0; i < headers.length; i++) {
                obj[headers[i]] = row[i];
            }
            return obj;
        });
    }

    /**
     * Writes an array of objects back to the sheet.
     * Automatically maps object keys to the correct columns based on tool headers.
     * @param {string} toolKey - Tool key (e.g., 'MAIL_SENDER')
     * @param {Object[]} objectsArray - Array of objects to write
     * @param {number} [startRow] - Optional start row to write from. Defaults to lastRow + 1.
     */
    function writeObjects(toolKey, objectsArray, startRow) {
        if (!objectsArray || objectsArray.length === 0) return;

        var sheet = getSheet(toolKey);
        var headers = getHeaders(toolKey);

        var data2D = objectsArray.map(function(obj) {
            var row = [];
            for (var i = 0; i < headers.length; i++) {
                row.push(obj[headers[i]] !== undefined ? obj[headers[i]] : "");
            }
            return row;
        });

        var targetRow = startRow || Math.max(2, sheet.getLastRow() + 1);

        var range = sheet.getRange(targetRow, 1, data2D.length, headers.length);
        range.setValues(data2D);
    }

    function overwriteRows(toolKey, rows, options) {
        var opts = options || {};
        var sheet = getSheet(toolKey);
        var cfg = SyncEngine.getTool(toolKey);
        var headers = getHeaders(toolKey);
        var totalCols = opts.totalCols || headers.length || sheet.getLastColumn();
        var lastRow = sheet.getLastRow();

        if (lastRow >= 2) {
            sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), totalCols)).clearContent();
        }

        if (rows && rows.length > 0) {
            sheet.getRange(2, 1, rows.length, totalCols).setValues(rows);
        }

        _App_applyBodyFormatting(sheet, rows ? rows.length : 0, opts.formatConfig || cfg.FORMAT_CONFIG);
    }

    /**
     * Overwrites all data starting from row 2 with the given objects array.
     */
    function overwriteObjects(toolKey, objectsArray) {
        clearData(toolKey);
        if (objectsArray && objectsArray.length > 0) {
            writeObjects(toolKey, objectsArray, 2);
        }
        _App_applyBodyFormatting(getSheet(toolKey), objectsArray ? objectsArray.length : 0, SyncEngine.getTool(toolKey).FORMAT_CONFIG);
    }

    /**
     * Clears all data rows (row 2 onwards) for the specified tool.
     */
    function clearData(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        var cfg = SyncEngine.getTool(toolKey);
        var headers = getHeaders(toolKey);
        if (lastRow >= 2) {
            var colCount = headers.length || sheet.getLastColumn();
            sheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
            if (cfg.FORMAT_CONFIG) {
                _App_applyBodyFormatting(sheet, 0, cfg.FORMAT_CONFIG);
            }
        }
    }

    /**
     * Returns only the values in the 'Action' column for quick scanning.
     * Returns an array of strings.
     */
    function getActions(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        var headerMap = getHeaderMap(toolKey);
        var actionColIdx = headerMap['Action'] || headerMap['ON/OFF'] || 1;

        var values = sheet.getRange(2, actionColIdx, lastRow - 1, 1).getValues();
        return values.map(function(row) { return row[0]; });
    }

    function hasPendingActions(toolKey) {
        return getActions(toolKey).some(function(action) {
            if (action === undefined || action === null || action === false) {
                return false;
            }
            if (typeof action === 'string') {
                return action.trim() !== '';
            }
            return true;
        });
    }

    function getActionStats(toolKey, actionNames) {
        var actions = getActions(toolKey);
        var stats = {};
        (actionNames || []).forEach(function(action) {
            stats[action] = 0;
        });

        actions.forEach(function(action) {
            if (stats.hasOwnProperty(action)) {
                stats[action]++;
            }
        });

        return stats;
    }

    function patchRow(toolKey, rowNumber, updates) {
        if (!updates || Object.keys(updates).length === 0) return;
        var sheet = getSheet(toolKey);
        var headerMap = getHeaderMap(toolKey);
        var lastCol = sheet.getLastColumn();
        if (lastCol === 0) return;

        var range = sheet.getRange(rowNumber, 1, 1, lastCol);
        var rowData = range.getValues()[0];

        var hasChanges = false;
        Object.keys(updates).forEach(function(header) {
            if (headerMap[header]) {
                var colIndex = headerMap[header] - 1;
                if (colIndex < lastCol && rowData[colIndex] !== updates[header]) {
                    rowData[colIndex] = updates[header];
                    hasChanges = true;
                }
            }
        });

        if (hasChanges) {
            range.setValues([rowData]);
        }
    }

    function batchPatchRows(toolKey, rowNumbers, updatesArray) {
        if (!rowNumbers || !updatesArray || rowNumbers.length === 0 || rowNumbers.length !== updatesArray.length) return;
        
        var sheet = getSheet(toolKey);
        var headerMap = getHeaderMap(toolKey);
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        
        if (lastRow < 2 || lastCol === 0) return;
        
        // Map row numbers to their updates
        var rowUpdatesMap = {};
        for (var i = 0; i < rowNumbers.length; i++) {
            rowUpdatesMap[rowNumbers[i]] = updatesArray[i];
        }

        // Sort unique row numbers
        var sortedRows = Object.keys(rowUpdatesMap).map(Number).sort(function(a, b) { return a - b; });

        // Group rows into blocks where the gap between consecutive rows is <= MAX_GAP
        var MAX_GAP = 5;
        var blocks = [];
        var currentBlock = [sortedRows[0]];

        for (var i = 1; i < sortedRows.length; i++) {
            if (sortedRows[i] - sortedRows[i - 1] <= MAX_GAP) {
                currentBlock.push(sortedRows[i]);
            } else {
                blocks.push(currentBlock);
                currentBlock = [sortedRows[i]];
            }
        }
        blocks.push(currentBlock);

        // If blocks are highly fragmented, update in one single read/write sweep
        if (blocks.length > 3) {
            var minRow = sortedRows[0];
            var maxRow = sortedRows[sortedRows.length - 1];
            var numRows = maxRow - minRow + 1;

            var range = sheet.getRange(minRow, 1, numRows, lastCol);
            var data = range.getValues();
            var hasChanges = false;

            sortedRows.forEach(function(actualRow) {
                var relativeIdx = actualRow - minRow;
                var updates = rowUpdatesMap[actualRow];
                if (updates && relativeIdx >= 0 && relativeIdx < data.length) {
                    Object.keys(updates).forEach(function(header) {
                        if (headerMap[header]) {
                            var colIndex = headerMap[header] - 1;
                            if (colIndex < lastCol && data[relativeIdx][colIndex] !== updates[header]) {
                                data[relativeIdx][colIndex] = updates[header];
                                hasChanges = true;
                            }
                        }
                    });
                }
            });

            if (hasChanges) {
                range.setValues(data);
            }
        } else {
            // Process each block separately (contiguous updates)
            blocks.forEach(function(block) {
                var minRow = block[0];
                var maxRow = block[block.length - 1];
                var numRows = maxRow - minRow + 1;
                
                var range = sheet.getRange(minRow, 1, numRows, lastCol);
                var data = range.getValues();
                var hasChanges = false;

                block.forEach(function(actualRow) {
                    var relativeIdx = actualRow - minRow;
                    var updates = rowUpdatesMap[actualRow];
                    if (updates && relativeIdx >= 0 && relativeIdx < data.length) {
                        Object.keys(updates).forEach(function(header) {
                            if (headerMap[header]) {
                                var colIndex = headerMap[header] - 1;
                                if (colIndex < lastCol && data[relativeIdx][colIndex] !== updates[header]) {
                                    data[relativeIdx][colIndex] = updates[header];
                                    hasChanges = true;
                                }
                            }
                        });
                    }
                });

                if (hasChanges) {
                    range.setValues(data);
                }
            });
        }
    }

    function assertActiveSheet(toolKey) {
        var cfg = SyncEngine.getTool(toolKey);
        return _App_assertActiveSheet(cfg.SHEET_NAME);
    }

    function syncDynamicColumns(toolKey, dynamicHeaders, options) {
        delete _headersCache[toolKey]; // Invalidate cache
        return _App_syncDynamicColumns(toolKey, dynamicHeaders, options);
    }

    /**
     * Reads only the rows where the specified 'Action' column is not empty.
     * This is significantly faster for large sheets with sparse actions.
     * @param {string} toolKey - Tool key (e.g., 'MAIL_MERGE')
     * @param {Object} [options] - { useDisplayValues: boolean, actionColName: string }
     * @returns {Object[]} Array of objects with an additional '_rowNumber' property.
     */
    function readPendingObjects(toolKey, options) {
        var opts = options || {};
        var actionColName = opts.actionColName || 'Action';
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        var headers = getHeaders(toolKey);
        if (headers.length === 0) return [];
        
        // 1. Find the Action column index dynamically from sheet headers
        var headerMap = getHeaderMap(toolKey);
        var actionColIdx = headerMap[actionColName] || 1; 

        // 2. Read only the Action column to identify pending rows
        var actionRange = sheet.getRange(2, actionColIdx, lastRow - 1, 1);
        var actionValues = actionRange.getValues();
        var pendingIndices = []; // 0-based relative to row 2
        for (var i = 0; i < actionValues.length; i++) {
            var val = actionValues[i][0];
            if (val !== undefined && val !== null && val !== "" && val !== false) {
                pendingIndices.push(i);
            }
        }

        if (pendingIndices.length === 0) return [];

        // 3. Determine if we should perform sparse range reads or a single full-sheet read
        var MAX_GAP = 5;
        var blockCount = 1;
        var lastEnd = pendingIndices[0];
        for (var k = 1; k < pendingIndices.length; k++) {
            if (pendingIndices[k] - lastEnd > MAX_GAP) {
                blockCount++;
            }
            lastEnd = pendingIndices[k];
        }

        var results = [];

        if (blockCount > 3) {
            // Full-sheet read (extremely fast for high fragmentation)
            var fullRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
            var fullData = opts.useDisplayValues ? fullRange.getDisplayValues() : fullRange.getValues();
            pendingIndices.forEach(function(idx) {
                var row = fullData[idx];
                if (row) {
                    var actionVal = row[actionColIdx - 1];
                    if (actionVal !== undefined && actionVal !== null && actionVal !== "" && actionVal !== false) {
                        var obj = { _rowNumber: idx + 2 };
                        for (var j = 0; j < headers.length; j++) {
                            obj[headers[j]] = row[j];
                        }
                        results.push(obj);
                    }
                }
            });
        } else {
            // Block read (efficient for small/contiguous changes)
            var startIdx = pendingIndices[0];
            var endIdx = startIdx;

            var processBlock = function(s, e) {
                var numRows = e - s + 1;
                var range = sheet.getRange(s + 2, 1, numRows, headers.length);
                var data = opts.useDisplayValues ? range.getDisplayValues() : range.getValues();
                data.forEach(function(row, offset) {
                    var actionVal = row[actionColIdx - 1];
                    if (actionVal !== undefined && actionVal !== null && actionVal !== "" && actionVal !== false) {
                        var obj = { _rowNumber: s + offset + 2 };
                        for (var j = 0; j < headers.length; j++) {
                            obj[headers[j]] = row[j];
                        }
                        results.push(obj);
                    }
                });
            };

            for (var k = 1; k < pendingIndices.length; k++) {
                if (pendingIndices[k] - endIdx <= MAX_GAP) {
                    endIdx = pendingIndices[k];
                } else {
                    processBlock(startIdx, endIdx);
                    startIdx = pendingIndices[k];
                    endIdx = startIdx;
                }
            }
            processBlock(startIdx, endIdx);
        }

        return results;
    }

    return {
        getSheet: getSheet,
        ensureSheet: ensureSheet,
        getHeaders: getHeaders,
        getHeaderMap: getHeaderMap,
        getNormalizedHeaderMap: getNormalizedHeaderMap,
        getSheetHeaderMap: getSheetHeaderMap,
        readObjects: readObjects,
        readPendingObjects: readPendingObjects,
        writeObjects: writeObjects,
        overwriteRows: overwriteRows,
        overwriteObjects: overwriteObjects,
        clearData: clearData,
        getActions: getActions,
        hasPendingActions: hasPendingActions,
        getActionStats: getActionStats,
        patchRow: patchRow,
        batchPatchRows: batchPatchRows,
        assertActiveSheet: assertActiveSheet,
        syncDynamicColumns: syncDynamicColumns
    };

})();


// --- FILE: core/05_Core_Utils.js ---
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

function _App_include(filename) {
    try {
        return HtmlService.createHtmlOutputFromFile(filename).getContent();
    } catch (e) {
        if (filename.indexOf('/') === -1) {
            try {
                return HtmlService.createHtmlOutputFromFile('core/' + filename).getContent();
            } catch (e2) {
                try {
                    return HtmlService.createHtmlOutputFromFile('tools/' + filename).getContent();
                } catch (e3) {
                    throw new Error("Could not find file: '" + filename + "' (tried raw, core/, tools/)");
                }
            }
        }
        throw e;
    }
}

function _App_createTemplateFromFile(filename) {
    try {
        return HtmlService.createTemplateFromFile(filename);
    } catch (e) {
        if (filename.indexOf('/') === -1) {
            try {
                return HtmlService.createTemplateFromFile('tools/' + filename);
            } catch (e2) {
                try {
                    return HtmlService.createTemplateFromFile('core/' + filename);
                } catch (e3) {
                    throw new Error("Could not find template file: '" + filename + "' (tried raw, tools/, core/)");
                }
            }
        }
        throw e;
    }
}


function _App_formatStatus(prefixKey, message) {
    var theme = (typeof SHEET_THEME !== 'undefined' && SHEET_THEME && SHEET_THEME.STATUS_PREFIXES)
        ? SHEET_THEME
        : DEFAULT_SHEET_THEME;
    var prefixes = theme.STATUS_PREFIXES || {};
    var key = String(prefixKey || '').toUpperCase();
    var prefix = prefixes[key] || '';
    var text = (message === undefined || message === null) ? '' : String(message);
    return prefix + text;
}

function _App_makeRowPatch(rowNumber, updates) {
    var patch = {};
    if (updates && typeof updates === 'object') {
        Object.keys(updates).forEach(function (key) {
            patch[key] = updates[key];
        });
    }
    patch._rowNumber = rowNumber;
    return patch;
}

function _App_makeStatusPatch(rowNumber, prefixKey, message, updates) {
    var patch = _App_makeRowPatch(rowNumber, updates);
    patch.Status = _App_formatStatus(prefixKey, message);
    return patch;
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
            var is403Retriable = msg.indexOf('403') !== -1 && (
                msg.indexOf('rate limit') !== -1 ||
                msg.indexOf('quota') !== -1 ||
                msg.indexOf('limit exceeded') !== -1 ||
                msg.indexOf('too many') !== -1
            );
            var isRetriable = (
                is403Retriable || msg.indexOf('429') !== -1 ||
                msg.indexOf('500') !== -1 || msg.indexOf('502') !== -1 ||
                msg.indexOf('503') !== -1 ||
                msg.indexOf('rate limit') !== -1 || msg.indexOf('quota') !== -1 ||
                msg.indexOf('limit exceeded') !== -1 || msg.indexOf('too many') !== -1
            );
            if (isRetriable && n < maxRetries) {
                var waitMs = (Math.pow(2, n) * 1000) + Math.round(Math.random() * 1000);
                Utilities.sleep(waitMs);
            } else {
                throw e;
            }
        }
    }
}

// ==========================================
// Execution Time Tracking — 6-min Limit Guard
// ==========================================
var _App_executionStartTime = 0;
var _App_executionLimitMs = 330 * 1000; // 5.5 minutes (330 seconds)

/**
 * Resets the global execution timer to current time.
 */
function _App_resetExecutionTimer() {
    _App_executionStartTime = Date.now();
}

/**
 * Returns true if the script is approaching the Google Apps Script 6-minute limit.
 */
function _App_isExecutionLimitApproaching() {
    if (_App_executionStartTime === 0) return false;
    return (Date.now() - _App_executionStartTime > _App_executionLimitMs);
}

// ==========================================
// ==========================================
// Centralized Google API Error Translation Engine
// ==========================================
/**
 * Parses raw JSON or Google Apps Script exception messages and returns a user-friendly, actionable description.
 *
 * @param {Error|string|Object} err - The original exception thrown by an API or system call
 * @returns {Object} { message: string, isFatal: boolean, category: string }
 */
function _App_translateApiError(err) {
    var rawMsg = '';
    if (err) {
        if (typeof err === 'string') rawMsg = err;
        else if (err.message) rawMsg = err.message;
        else rawMsg = String(err);
    }
    
    var lower = rawMsg.toLowerCase();
    var result = {
        message: rawMsg,
        originalMessage: rawMsg,
        isFatal: false,
        category: 'unknown'
    };

    // 1. Authorization / OAuth Scope Errors
    if (lower.indexOf('authentication') !== -1 ||
        lower.indexOf('authorization') !== -1 ||
        lower.indexOf('scopes') !== -1 ||
        lower.indexOf('403') !== -1 ||
        lower.indexOf('access_denied') !== -1 ||
        lower.indexOf('unauthorized') !== -1 ||
        lower.indexOf('access denied') !== -1) {
        
        result.message = "⚠️ Authorization Required: The script lacks necessary Google permissions. Please re-authorize the script and try again.";
        result.isFatal = true;
        result.category = 'auth';
    }
    // 2. Rate Limits & Quotas
    else if (lower.indexOf('quota') !== -1 ||
             lower.indexOf('limit exceeded') !== -1 ||
             lower.indexOf('too many') !== -1 ||
             lower.indexOf('429') !== -1 ||
             lower.indexOf('rate limit') !== -1) {
             
        result.message = "⚠️ Quota Exceeded: You have hit Google's daily API limits for this service. Please wait a few hours or until tomorrow to resume syncing.";
        result.isFatal = true;
        result.category = 'quota';
    }
    // 3. Document or Resource Not Found
    else if (lower.indexOf('not found') !== -1 ||
             lower.indexOf('404') !== -1 ||
             lower.indexOf('cannot find') !== -1 ||
             lower.indexOf('inaccessible') !== -1) {
             
        result.message = "⚠️ Resource Not Found: The specified event, file, calendar, or folder ID is invalid, was deleted, or is not shared with this Google Account.";
        result.isFatal = false;
        result.category = 'not_found';
    }
    // 4. Temporary / Transient Server Failure
    else if (lower.indexOf('500') !== -1 ||
             lower.indexOf('502') !== -1 ||
             lower.indexOf('503') !== -1 ||
             lower.indexOf('internal error') !== -1 ||
             lower.indexOf('service unavailable') !== -1) {
             
        result.message = "❌ Service Error: Google servers are temporarily overwhelmed. The system will auto-retry, but please wait and try again if it continues to fail.";
        result.isFatal = false;
        result.category = 'transient';
    }
    // 5. Native parameter mismatch
    else if (lower.indexOf("don't match the method signature") !== -1 ||
             lower.indexOf("invalid parameter") !== -1 ||
             lower.indexOf("400") !== -1) {
             
        result.message = "⚠️ Formatting Error: The data provided in the sheet doesn't match the required types (e.g. invalid date formats, missing compulsory fields).";
        result.isFatal = false;
        result.category = 'validation';
    }
    
    return result;
}

// ==========================================
// Centralized Sheet Row Validator
// ==========================================
/**
 * Validates a sheet row object against the tool's registered COL_SCHEMA rules.
 * @param {Object} item - Row object (usually from SheetManager.readPendingObjects)
 * @param {string} toolKey - Tool identifier (e.g. 'CALENDAR_SYNC')
 * @returns {string|null} Error string if validation fails, or null if row is fully valid.
 */
function _App_validateRowAgainstSchema(item, toolKey) {
    if (typeof SyncEngine === 'undefined' || !toolKey) return null;
    try {
        var toolCfg = SyncEngine.getTool(toolKey);
        if (!toolCfg || !toolCfg.FORMAT_CONFIG || !toolCfg.FORMAT_CONFIG.COL_SCHEMA) {
            return null;
        }
        var schema = toolCfg.FORMAT_CONFIG.COL_SCHEMA;
        for (var i = 0; i < schema.length; i++) {
            var col = schema[i];
            var colHeader = col.header;
            var colType = col.type;
            
            // Skip validation on standard structural columns
            if (colHeader === 'Action' || colHeader === 'Status') continue;
            
            if (item && Object.prototype.hasOwnProperty.call(item, colHeader)) {
                var val = item[colHeader];
                var isValid = _App_validateValueByType(colType, val, col);
                if (!isValid) {
                    return "⚠️ Data Error: Invalid format in column '" + colHeader + "' (Expected " + colType + ")";
                }
            }
        }
    } catch (e) {
        // Fallback gracefully
    }
    return null;
}



/**
 * Converts a 1-based column number into its corresponding A-Z/AA-ZZ Excel-style column letter.
 * @param {number} col - 1-based column index
 * @returns {string} The column letter (e.g. 'A', 'Z', 'AA')
 */
function _App_getColumnLetter(col) {
    var temp, letter = '';
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}

// ==========================================
// Centralized Document & Email Utilities
// ==========================================

/**
 * Escapes special regex characters in a string.
 */
function _App_escapeRegExp(string) {
  if (!string) return "";
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Consolidated utility to extract a Google Resource ID (Docs, Sheets, Slides, Forms, Drive files/folders)
 * from a full URL, or returns the input if it is already a raw ID.
 * @param {string} urlOrId - The URL or ID string.
 * @returns {string|null} The extracted ID or null if invalid/empty.
 */
function _App_extractIdFromUrl(urlOrId) {
    if (!urlOrId || typeof urlOrId !== 'string') return null;
    var trimmed = urlOrId.trim();
    if (!trimmed) return null;
    
    // Pattern matches typical Google ID formats (usually 25-100 characters of alphanumeric/hyphen/underscore)
    var idMatch = trimmed.match(/[-\w]{25,100}/);
    if (idMatch) {
        return idMatch[0];
    }
    
    // Fallback pattern for URLs containing /d/ID
    var dMatch = trimmed.match(/\/d\/([-\w]+)/);
    if (dMatch) {
        return dMatch[1];
    }
    
    // If it looks like a raw ID (no slashes), return it
    if (trimmed.indexOf('/') === -1) {
        return trimmed;
    }
    
    return null;
}

/**
 * Utility to execute multiple HTTP requests in parallel using UrlFetchApp.fetchAll.
 * Can be used by tools for parallel API queries or external webhooks.
 * @param {Object[]} requests - Array of request objects {url, method, headers, payload, etc.}
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse[]} Array of HTTPResponse objects.
 */
function _App_fetchParallel(requests) {
    if (!requests || requests.length === 0) return [];
    return UrlFetchApp.fetchAll(requests);
}

/**
 * Retrieves a Drive file blob by URL or raw File ID.
 */
function _App_getDriveAttachment(fileIdOrUrl) {
  try {
    if (!fileIdOrUrl) return null;
    var fileId = _App_extractIdFromUrl(fileIdOrUrl);
    if (!fileId) throw new Error("Invalid File ID");

    var file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    throw new Error("Cannot find attachment in Drive (" + fileIdOrUrl + ")");
  }
}

/**
 * Merges two comma-separated lists of email addresses uniquely.
 */
function _App_mergeEmails(existingStr, newStr) {
  if (!newStr) return existingStr || "";
  var existingArr = (existingStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  var newArr = (newStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  newArr.forEach(function (em) {
    if (existingArr.indexOf(em) === -1) {
      existingArr.push(em);
    }
  });
  return existingArr.join(',');
}

/**
 * Formats a Date object or parseable date-time representation safely using the 
 * spreadsheet's active timezone. Falls back to Session.getScriptTimeZone() if needed.
 * @param {Date|string|number} date - Date representation to format
 * @param {string} [format] - Target format string, defaults to "MM/dd/yyyy HH:mm:ss"
 * @returns {string} Formatted date-time string
 */
function _App_formatDateTime(date, format) {
    if (!date) return "";
    var targetDate = (date instanceof Date) ? date : new Date(date);
    if (isNaN(targetDate.getTime())) return "";
    var tz = Session.getScriptTimeZone();
    try {
        tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch (e) {
        // Fallback safely if run outside spreadsheet context
    }
    return Utilities.formatDate(targetDate, tz, format || "MM/dd/yyyy HH:mm:ss");
}







// --- FILE: core/05_Core_Utils_Batch.js ---
// ==========================================
// _App_BatchProcessor — Unified Iteration Engine
// ==========================================

/**
 * A centralized utility for processing batches of items (rows) with automated
 * progress tracking, error handling, time-limit guarding, and logging.
 *
 * @param {string} toolKey      - Tool identifier from SyncEngine (e.g. 'MAIL_MERGE')
 * @param {Array} items         - Array of items to process (usually objects with data and originalIndex)
 * @param {Function} processFn  - Callback function(item, index) that processes one item. 
 *                                 Should return a result object (e.g. row update data) or throw an error.
 * @param {Object} [options]    - Optional configuration:
 *                                 - {number} batchSize: How many items to process before updating progress/checking time (default 10)
 *                                 - {boolean} stopOnFailure: If true, stops the entire batch if one item fails (default false)
 *                                 - {Function} onBatchComplete: function(results) called after each batch segment.
 *
 * @returns {Object} { 
 *   processedCount: number, 
 *   errorCount: number, 
 *   timeLimitReached: boolean, 
 *   results: Array 
 * }
 */
function _App_BatchProcessor(toolKey, items, processFn, options) {
    var opts = options || {};
    var batchSize = opts.batchSize || 10;
    var total = items.length;
    
    var stats = {
        processedCount: 0,
        errorCount: 0,
        timeLimitReached: false,
        results: []
    };

    if (total === 0) return stats;

    // Ensure timer is running if not already set
    if (_App_executionStartTime === 0) _App_resetExecutionTimer();

    for (var i = 0; i < total; i += batchSize) {
        // 1. Time-Limit Guard
        if (_App_isExecutionLimitApproaching()) {
            stats.timeLimitReached = true;
            break;
        }

        var segment = items.slice(i, i + batchSize);
        var segmentResults = [];

        // 3. Process Segment
        for (var j = 0; j < segment.length; j++) {
            var item = segment[j];
            var globalIndex = i + j;

            try {
                // 3a. Pre-Validation Engine Check
                var schemaValidationError = _App_validateRowAgainstSchema(item, toolKey);
                if (schemaValidationError) {
                    throw new Error(schemaValidationError);
                }

                // Wrap in backoff retry for transient API issues
                var result = _App_callWithBackoff(function() {
                    return processFn(item, globalIndex);
                });
                
                segmentResults.push(result);
                stats.results.push(result);
                stats.processedCount++;
            } catch (err) {
                stats.errorCount++;
                
                var translated = _App_translateApiError(err);
                
                var statusMsg = translated.message;
                if (translated.category !== 'unknown' && translated.originalMessage && translated.originalMessage !== translated.message) {
                    var detailText = translated.originalMessage;
                    if (detailText.length > 150) {
                        detailText = detailText.substring(0, 147) + "...";
                    }
                    // Strip duplicate emojis and error labels from original message
                    detailText = detailText.replace(/^[⚠️❌⏳ℹ️✅]\s*/, '').replace(/^(?:Data|Formatting|API|System)\s+Error:\s*/i, '');
                    
                    var colonIdx = translated.message.indexOf(':');
                    var prefix = colonIdx !== -1 ? translated.message.substring(0, colonIdx + 1) : "❌ Error:";
                    statusMsg = prefix + " " + detailText;
                }
                
                // Return an error object to the tool so it can write to the Status column
                var errObj = { isError: true, error: statusMsg };
                if (item && item._rowNumber !== undefined) {
                    errObj._rowNumber = item._rowNumber;
                }
                
                segmentResults.push(errObj);
                stats.results.push(errObj);
                
                // If it is a fatal system/connection error (OAuth scopes / Rate limit), halt the execution
                if (translated.isFatal) {
                    _App_clearProgress(toolKey);
                    
                    // Flush what was successfully processed (and the current fatal item) before halting
                    if (opts.onBatchComplete && segmentResults.length > 0) {
                        try {
                            opts.onBatchComplete(segmentResults);
                        } catch (writeErr) {
                            // Suppress errors during flush write to prioritize bubbling up original fatal error
                        }
                    }
                    
                    throw new Error(translated.message);
                }
                
                if (opts.stopOnFailure) {
                    _App_clearProgress(toolKey);
                    throw err;
                }
            }
        }

        // 2. Progress Tracking (CacheService) — Update after segment
        _App_setProgress(toolKey, stats.processedCount + stats.errorCount, total);

        // 4. Batch Lifecycle Hook
        if (opts.onBatchComplete && segmentResults.length > 0) {
            opts.onBatchComplete(segmentResults);
        }
    }

    // 5. Cleanup
    _App_clearProgress(toolKey);
    return stats;
}

/**
 * Centralized utility to batch patch rows in SheetManager from BatchProcessor results.
 *
 * @param {string} toolKey - Tool identifier (e.g., 'MAIL_MERGE')
 * @param {Array} batchResults - The results array from the batch segment
 * @param {Function} [successFieldsMapper] - Optional callback function(res) returning an object
 *                                            containing the middle/read-only columns to update on success.
 *                                            (Action and Status are automatically handled if not returned).
 */
function _App_batchPatchResults(toolKey, batchResults, successFieldsMapper) {
    var rowNumbers = [];
    var patchData = [];

    batchResults.forEach(function (res) {
        if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            if (res.isError) {
                // On error, write the error status. We keep the action so user can retry.
                patchData.push(_App_makeStatusPatch(res._rowNumber, 'ERROR', res.error));
            } else {
                // On success, construct updates
                var updates = {
                    'Action': res.action !== undefined ? res.action : '',
                    'Status': res.status !== undefined ? res.status : _App_formatStatus('SUCCESS', 'Processed')
                };
                if (successFieldsMapper) {
                    var customFields = successFieldsMapper(res);
                    if (customFields && typeof customFields === 'object') {
                        Object.keys(customFields).forEach(function (k) {
                            updates[k] = customFields[k];
                        });
                    }
                }
                patchData.push(_App_makeRowPatch(res._rowNumber, updates));
            }
        }
    });

    if (rowNumbers.length > 0) {
        SheetManager.batchPatchRows(toolKey, rowNumbers, patchData);
    }
}


// --- FILE: core/05_Core_Utils_Email.js ---
// ==========================================
// Centralized Email Utilities
// ==========================================

/**
 * Unified helper to send an email or create a draft.
 * Supports thread replies and file attachments.
 * 
 * @param {Object} options
 * @param {string} options.action - "SEND" or "DRAFT"
 * @param {string} options.to - Primary recipient(s)
 * @param {string} options.cc - CC recipient(s)
 * @param {string} options.bcc - BCC recipient(s)
 * @param {string} options.subject - Email subject
 * @param {string} options.body - Email HTML body
 * @param {Blob[]} options.attachments - Attachment blobs
 * @param {string} [options.threadIdOrSubject] - Thread ID or thread subject to reply to
 * @returns {string} User-friendly result status message
 */
function _App_sendOrDraftEmail(options) {
    var action = (options.action || "").toString().trim().toUpperCase();
    var to = options.to || "";
    var cc = options.cc || "";
    var bcc = options.bcc || "";
    var subject = options.subject || "";
    var body = options.body || "";
    var attachments = options.attachments || [];
    var threadIdOrSubject = options.threadIdOrSubject || "";

    if (!to && !threadIdOrSubject) {
        throw new Error("⚠️ Missing Email To");
    }
    if (!subject && !threadIdOrSubject) {
        throw new Error("⚠️ Missing Email Subject");
    }

    var mailOptions = {
        htmlBody: body,
        attachments: attachments
    };

    if (threadIdOrSubject) {
        var thread = null;
        try { 
            thread = GmailApp.getThreadById(threadIdOrSubject); 
        } catch (ignore) {}

        if (!thread) {
            var safeSubject = threadIdOrSubject.toString().replace(/['"]/g, '');
            var query = 'subject:("' + safeSubject + '")';
            var threads = GmailApp.search(query, 0, 1);
            if (threads && threads.length > 0) thread = threads[0];
        }
        if (!thread) {
            throw new Error("⚠️ Thread not found for ID or Subject");
        }

        var messages = thread.getMessages();
        var lastMessage = messages[messages.length - 1];

        var existingTo = lastMessage.getTo();
        var existingCc = lastMessage.getCc();

        var newTo = _App_mergeEmails(existingTo, to);
        var newCc = _App_mergeEmails(existingCc, cc);

        var replyOptions = {
            htmlBody: body,
            attachments: attachments,
            cc: newCc || "",
            bcc: bcc || ""
        };

        var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
        draftReply.update(newTo || "", subject, "", replyOptions);

        if (action === "SEND") {
            draftReply.send();
            return _App_formatStatus('SUCCESS', "Sent (" + _App_formatDateTime(new Date()) + ")");
        } else {
            return _App_formatStatus('SUCCESS', "Reply Draft Created");
        }
    } else {
        mailOptions.cc = cc;
        mailOptions.bcc = bcc;

        if (action === "SEND") {
            GmailApp.sendEmail(to, subject, "", mailOptions);
            return _App_formatStatus('SUCCESS', "Sent (" + _App_formatDateTime(new Date()) + ")");
        } else {
            GmailApp.createDraft(to, subject, "", mailOptions);
            return _App_formatStatus('SUCCESS', "Draft Created");
        }
    }
}


// --- FILE: core/05_Core_Utils_Lock.js ---
// ==========================================
// Centralized Lock Utilities
// ==========================================

/**
 * Executes a callback within a document lock, ensuring concurrency safety.
 */
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


// --- FILE: core/06_Core_Validators.js ---
// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    BOOLEAN: function(val) {
        if (typeof val === 'boolean') return true;
        if (typeof val === 'string') {
            var lower = val.toLowerCase();
            return lower === 'true' || lower === 'false';
        }
        return false;
    },
    DOCS_URL: function(val) {
        if (val === '' || val === null || val === undefined) return true;
        return typeof val === 'string' && val.indexOf('docs.google.com/document') !== -1;
    },
    DRIVE_URL: function(val) {
        if (val === '' || val === null || val === undefined) return true;
        return typeof val === 'string' && (val.indexOf('drive.google.com') !== -1 || val.indexOf('docs.google.com') !== -1);
    }
};

/**
 * Validates a single email address using a robust regex.
 * @param {string} email
 * @returns {boolean}
 */
function _App_validateEmail(email) {
    if (!email || typeof email !== 'string') return false;
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
}

/**
 * Validates a comma-separated list of email addresses.
 * @param {string} emailsString
 * @param {boolean} [allowEmpty] - If true, an empty string is considered valid.
 * @returns {boolean}
 */
function _App_validateEmailList(emailsString, allowEmpty) {
    var shouldAllowEmpty = allowEmpty !== false; // default to true, since CC/BCC etc are optional
    var val = (emailsString || '').toString().trim();
    if (val === '') return shouldAllowEmpty;

    var emails = val.split(',');
    for (var i = 0; i < emails.length; i++) {
        var email = emails[i].trim();
        if (email && !_App_validateEmail(email)) {
            return false;
        }
    }
    return true;
}

/**
 * Unifies column cell validation by checking its value against schema column type definitions.
 * @param {string} type - Column schema type (e.g. 'EMAIL', 'DATE', 'BOOLEAN', 'URL', 'DROPDOWN')
 * @param {*} value - Cell value to check
 * @param {Object} [fieldConfig] - The schema field configuration object (contains header, options, etc.)
 * @returns {boolean} True if value is valid, false otherwise.
 */
function _App_validateValueByType(type, value, fieldConfig) {
    var valStr = (value === null || value === undefined) ? '' : String(value).trim();
    
    // ACTION, STATUS, ID, and READ_ONLY do not require data-type validation or are handled natively
    if (type === 'ACTION' || type === 'STATUS' || type === 'ID' || type === 'READ_ONLY') {
        return true;
    }
    
    switch (type) {
        case 'EMAIL':
            if (valStr === '') return true; // Optional fields are empty
            return _App_validateEmail(valStr);
            
        case 'EMAIL_LIST':
            return _App_validateEmailList(valStr, true);
            
        case 'DATE':
        case 'DATETIME':
            if (valStr === '') return true;
            return SYSTEM_VALIDATORS.DATE(value);
            
        case 'BOOLEAN':
            if (valStr === '') return true;
            return SYSTEM_VALIDATORS.BOOLEAN(value);
            
        case 'URL':
            if (valStr === '') return true;
            // Match URLs, allowing optional query parameters and hash fragments
            var urlRegex = /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*(\?[^\s]*)?(#[^\s]*)?$/i;
            return urlRegex.test(valStr);
            
        case 'DOCS_URL':
            return SYSTEM_VALIDATORS.DOCS_URL(valStr);
            
        case 'DRIVE_URL':
            return SYSTEM_VALIDATORS.DRIVE_URL(valStr);
            
        case 'DROPDOWN':
            if (valStr === '') return true;
            if (fieldConfig && fieldConfig.allowInvalid) return true;
            if (fieldConfig && fieldConfig.options) {
                var opts = fieldConfig.options;
                if (typeof opts === 'function') {
                    try {
                        opts = opts();
                    } catch (e) {
                        return true; // If dynamic evaluation fails, pass validation or gracefully ignore
                    }
                }
                if (Array.isArray(opts)) {
                    var lowerVal = valStr.toLowerCase();
                    return opts.some(function(opt) {
                        return String(opt).trim().toLowerCase() === lowerVal;
                    });
                }
            }
            return true;
            
        case 'TEXT':
        default:
            return true; // Standard text has no structural restrictions
    }
}





// --- FILE: core/07_Core_State.js ---
// ==========================================
// Progress Tracking — Unified CacheService Wrappers
// ==========================================

/**
 * Helper to build a unique cache key incorporating the Spreadsheet ID
 * to prevent progress cross-contamination between different open documents.
 * @param {string} toolName
 * @returns {string}
 * @private
 */
var _App_cachedSpreadsheetId = null;

function _App_getProgressKey_(toolName) {
    if (!_App_cachedSpreadsheetId) {
        try {
            _App_cachedSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
        } catch (e) {
            _App_cachedSpreadsheetId = "fallback";
        }
    }
    return _App_cachedSpreadsheetId + "_" + toolName + CACHE_KEYS.PROGRESS;
}

/**
 * Stores batch operation progress for sidebar polling.
 * @param {string} toolName     - Tool key e.g. 'MAIL_SENDER'
 * @param {number} current      - Items processed so far
 * @param {number} total        - Total items queued
 * @param {number} [ttlSec=600] - Cache TTL in seconds (default 10 min)
 */
function _App_setProgress(toolName, current, total, ttlSec) {
    CacheService.getUserCache().put(
        _App_getProgressKey_(toolName),
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
    var data = CacheService.getUserCache().get(_App_getProgressKey_(toolName));
    return data ? JSON.parse(data) : null;
}

/**
 * Removes progress state after an operation completes.
 * @param {string} toolName
 */
function _App_clearProgress(toolName) {
    CacheService.getUserCache().remove(_App_getProgressKey_(toolName));
}


// --- FILE: core/08_Sheets_Helpers.js ---
function _App_canScaffoldSheet(toolConfig) {
    return !!(toolConfig && toolConfig.HEADERS && toolConfig.HEADERS.length);
}
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


// --- FILE: core/09_Sheets_Formatting.js ---
// ==========================================
// Centralized Body Formatting Utility
// ==========================================

// Extra rows formatted beyond actual data to cover manual row additions.
var FORMATTING_BUFFER_ROWS = 30;



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
    var cfg = SyncEngine.getTool(toolKey);
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
    var cfg = SyncEngine.getTool(toolKey);
    try {
        var sheetName = cfg.SHEET_NAME;
        var cacheKey = "formatted_rows_" + sheetName.replace(/\s+/g, "_");
        CacheService.getDocumentCache().remove(cacheKey);
    } catch(e) {}
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
        _App_applyBodyFormatting(sheet, numRows, runtimeShape.formatConfig, true);
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
 * - First Columns (Action/Status): SHEET_THEME.FIRST_COLS_COLOR
 * - Middle Columns (Editable Data): SHEET_THEME.MIDDLE_COLS_COLOR
 * - Last Columns (Read-only/IDs): SHEET_THEME.LAST_COLS_COLOR
 */
function _App_applyBodyFormatting(sheet, numDataRows, config, forceConditional) {
    var rowsToFormat = numDataRows + FORMATTING_BUFFER_ROWS;
    var maxRows = sheet.getMaxRows();
    var actualRows = Math.min(rowsToFormat, maxRows - 1);
    if (actualRows < 1) return;

    var sheetName = sheet.getName();
    var cacheKey = "formatted_rows_" + sheetName.replace(/\s+/g, "_");
    var cachedValue = null;
    try {
        cachedValue = CacheService.getDocumentCache().get(cacheKey);
    } catch (e) {}
    var formattedRows = cachedValue ? parseInt(cachedValue, 10) : 0;

    if (formattedRows >= actualRows && !forceConditional) {
        return; // Skip redundant formatting calls
    }

    try {
        CacheService.getDocumentCache().put(cacheKey, String(actualRows), 21600); // 6 hours cache TTL
    } catch (e) {}

    var totalCols = config.COL_SCHEMA ? config.COL_SCHEMA.length : (config.totalCols || sheet.getLastColumn());

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

    // Apply Schema-driven validations and formats in batches
    if (config.COL_SCHEMA) {
        var colFontFamilies = [];
        var colFontStyles = [];
        var colBackgrounds = [];
        var colNumberFormats = [];
        var colValidations = [];

        config.COL_SCHEMA.forEach(function(colDef, index) {
            var colNum = index + 1;
            
            // Fonts
            var fontFamily = SHEET_THEME.FONTS.PRIMARY;
            if (colDef.type === 'ID' || colDef.type === 'URL') {
                fontFamily = SHEET_THEME.FONTS.MONOSPACE;
            }
            colFontFamilies.push(fontFamily);

            var fontStyle = 'normal';
            if (colDef.type === 'URL' || colDef.italic) {
                fontStyle = 'italic';
            }
            colFontStyles.push(fontStyle);
            
            // Background Colors (Schema-driven Categorization)
            var category = colDef.category;
            if (!category) {
                if (colDef.type === 'ACTION' || colDef.type === 'STATUS') category = 'FIRST_COLS';
                else if (colDef.type === 'READ_ONLY' || colDef.type === 'ID') category = 'LAST_COLS';
                else category = 'MIDDLE_COLS';
            }

            var bg = SHEET_THEME.MIDDLE_COLS_COLOR;
            if (category === 'FIRST_COLS') {
                bg = SHEET_THEME.FIRST_COLS_COLOR;
            } else if (category === 'LAST_COLS') {
                bg = SHEET_THEME.LAST_COLS_COLOR;
            }
            colBackgrounds.push(bg);

            // Number Formats
            var numFormat = '@'; // Force Plain Text by default
            if (colDef.type === 'DATETIME') {
                numFormat = 'MM/dd/yyyy hh:mm:ss AM/PM';
            } else if (colDef.type === 'DATE') {
                numFormat = 'MM/dd/yyyy';
            } else if (colDef.type === 'ID' || colDef.type === 'TEXT') {
                numFormat = '@';
            } else {
                numFormat = '';
            }
            colNumberFormats.push(numFormat);

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
            } else if (colDef.type === 'DATE' || colDef.type === 'DATETIME') {
                rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Enter a valid date.').build();
            } else if (colDef.type === 'URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), REGEXMATCH(' + letter + '2, "^https?:\\/\\/"))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid URL starting with http:// or https://.').setAllowInvalid(true).build();
            } else if (colDef.type === 'DOCS_URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), REGEXMATCH(' + letter + '2, "docs\\.google\\.com\\/document"))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid Google Docs URL.').setAllowInvalid(true).build();
            } else if (colDef.type === 'DRIVE_URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), OR(REGEXMATCH(' + letter + '2, "drive\\.google\\.com"), REGEXMATCH(' + letter + '2, "docs\\.google\\.com")))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid Google Drive or Docs URL.').setAllowInvalid(true).build();
            }
            colValidations.push(rule);
        });

        // Build 2D formatting grids in memory
        var fontFamilies2D = [];
        var fontStyles2D = [];
        var backgrounds2D = [];
        var numberFormats2D = [];
        var validations2D = [];

        for (var r = 0; r < actualRows; r++) {
            fontFamilies2D.push(colFontFamilies);
            fontStyles2D.push(colFontStyles);
            backgrounds2D.push(colBackgrounds);
            numberFormats2D.push(colNumberFormats);
            validations2D.push(colValidations);
        }

        // Apply formatting grids in single batch calls
        dataRange.setFontFamilies(fontFamilies2D);
        dataRange.setFontStyles(fontStyles2D);
        dataRange.setBackgrounds(backgrounds2D);
        dataRange.setNumberFormats(numberFormats2D);
        dataRange.setDataValidations(validations2D);
    }

    // 6. Conditional formatting rules
    _App_applyConditionalRules(sheet, actualRows, totalCols, config.conditionalRules || [], forceConditional);
}

/**
 * Builds and applies conditional formatting rules from a declarative descriptor array.
 * Replaces ALL existing conditional formatting rules on the sheet.
 *
 * Supported rule types: 'success', 'error', 'errorCross', 'pending', 'synced', 'custom'
 * Supported scopes: 'fullRow' (default), 'actionOnly', 'statusOnly'
 */
function _App_applyConditionalRules(sheet, numRows, totalCols, ruleDescriptors, force) {
    if (!force) {
        var existingRules = sheet.getConditionalFormatRules();
        if (existingRules && existingRules.length > 0) {
            // Bypass clearing and setting rules if already present to optimize execution speed
            return;
        }
    }

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
                .setBackground(SHEET_THEME.STATUS.WARNING)
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


// --- FILE: core/10_Engine_Core.js ---
// ==========================================
// SyncEngine — Plugin Registration System
// ==========================================

var SyncEngine = (function() {
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
        } else if (config.FORMAT_CONFIG && Array.isArray(config.FORMAT_CONFIG.COL_SCHEMA)) {
            var schema = config.FORMAT_CONFIG.COL_SCHEMA;
            if (schema.length > 0 && schema[0].type !== 'ACTION') {
                issues.push("First column must be of type 'ACTION'.");
            }
            if (schema.length > 1 && schema[1].type !== 'STATUS') {
                issues.push("Second column must be of type 'STATUS'.");
            }
        }

        if (config.HELP_ITEMS) {
            if (typeof config.HELP_ITEMS !== 'object') {
                issues.push("HELP_ITEMS must be an object.");
            } else if (config.HELP_ITEMS.items && !Array.isArray(config.HELP_ITEMS.items)) {
                issues.push("HELP_ITEMS.items must be an array.");
            }
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

        // Unifying defaults
        config.FROZEN_ROWS = typeof config.FROZEN_ROWS === 'number' ? config.FROZEN_ROWS : 1;
        config.FROZEN_COLS = typeof config.FROZEN_COLS === 'number' ? config.FROZEN_COLS : 2;

        // Post-process the config (generate HEADERS, totalCols, COL_WIDTHS from SCHEMA)
        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA) {
            config.HEADERS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(c) { return c.header; });
            config.FORMAT_CONFIG.totalCols = config.FORMAT_CONFIG.COL_SCHEMA.length;
            
            if (!config.COL_WIDTHS) {
                config.COL_WIDTHS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(col) {
                    return col.width || DEFAULT_COL_WIDTHS[col.type] || 150;
                });
            }

            // Auto-inject default conditional formatting rules
            if (!config.FORMAT_CONFIG.conditionalRules) {
                config.FORMAT_CONFIG.conditionalRules = [];
            }

            var actionColIndex = -1;
            var statusColIndex = -1;
            config.FORMAT_CONFIG.COL_SCHEMA.forEach(function(col, idx) {
                if (col.type === 'ACTION') actionColIndex = idx + 1;
                if (col.type === 'STATUS') statusColIndex = idx + 1;
            });

            var hasPendingRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'pending'; });
            var hasSuccessRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'success'; });
            var hasWarningRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'error'; });
            var hasErrorRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'errorCross'; });

            if (actionColIndex !== -1 && !hasPendingRule) {
                var actionColLetter = _App_getColumnLetter(actionColIndex);
                config.FORMAT_CONFIG.conditionalRules.push({
                    type: 'pending',
                    actionCol: actionColLetter,
                    scope: 'actionOnly'
                });
            }

            if (statusColIndex !== -1) {
                var statusColLetter = _App_getColumnLetter(statusColIndex);
                if (!hasSuccessRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'success',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
                if (!hasWarningRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'error',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
                if (!hasErrorRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'errorCross',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
            }
        }

        var issues = _validateToolConfig(key, config);
        if (issues.length > 0) {
            throw new Error("Tool '" + key + "' is misconfigured: " + issues.join(' '));
        }

        registry[key] = config;
    }

    /**
     * Retrieves a tool configuration by key.
     */
    function getTool(key) {
        var cfg = registry[key];
        if (!cfg) throw new Error('Unknown tool key: "' + key + '". Ensure the tool is registered via SyncEngine.registerTool().');
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

    function runAction(toolKey, actionName, args) {
        var cfg = getTool(toolKey);
        var action = cfg.ACTIONS && cfg.ACTIONS[actionName];
        if (typeof action !== 'function') {
            throw new Error("Action '" + actionName + "' not found on tool '" + toolKey + "'.");
        }

        var needsLock = (actionName === 'pull' || actionName === 'push');
        var execute = function() {
            try {
                if (actionName === 'pull' || actionName === 'push') {
                    _App_ensureSheetExists(toolKey);
                }
                return action.apply(cfg, args || []);
            } catch (err) {
                var translated = _App_translateApiError(err);
                return _App_fail(translated.message, null, {
                    originalError: translated.originalMessage || err.message || String(err),
                    stack: err.stack ? String(err.stack) : "",
                    toolKey: toolKey,
                    actionName: actionName,
                    timestamp: new Date().toISOString()
                });
            }
        };

        if (typeof _App_resetExecutionTimer === 'function') {
            _App_resetExecutionTimer();
        }

        return Logger.run(toolKey, actionName, function () {
            if (needsLock) {
                var lockName = toolKey + "_" + actionName.toUpperCase();
                return _App_withDocumentLock(lockName, execute);
            } else {
                return execute();
            }
        });
    }

    return {
        registerTool: registerTool,
        getTool: getTool,
        getAllTools: getAllTools,
        getToolKeys: getToolKeys,
        runAction: runAction,
        Utils: {
            ok: _App_ok,
            fail: _App_fail,
            withDocumentLock: _App_withDocumentLock,
            include: _App_include,
            createTemplateFromFile: _App_createTemplateFromFile,
            formatStatus: _App_formatStatus,
            throttle: _App_throttle,
            callWithBackoff: _App_callWithBackoff,
            isExecutionLimitApproaching: _App_isExecutionLimitApproaching,
            translateApiError: _App_translateApiError,
            validateRowAgainstSchema: _App_validateRowAgainstSchema,
            BatchProcessor: _App_BatchProcessor,
            batchPatchResults: _App_batchPatchResults,
            extractIdFromUrl: _App_extractIdFromUrl,
            fetchParallel: _App_fetchParallel,
            formatDateTime: _App_formatDateTime,
            sendOrDraftEmail: _App_sendOrDraftEmail
        }
    };
})();

/**
 * Backward compatibility Proxy for legacy scripts still referencing APP_REGISTRY directly.
 */
var APP_REGISTRY = new Proxy({}, {
    get: function(target, prop) {
        return SyncEngine.getTool(prop);
    },
    ownKeys: function() {
        return Object.keys(SyncEngine.getAllTools());
    },
    getOwnPropertyDescriptor: function(target, prop) {
        return { enumerable: true, configurable: true };
    }
});


// --- FILE: core/11_Engine_UI.js ---
// ==========================================
// _App_openSidebar — Universal Sidebar Opener
// ==========================================
/**
 * Opens a tool's sidebar, ensuring the sheet exists first.
 */
function _App_openSidebar(toolKey, postCreateCallback) {
    var cfg = SyncEngine.getTool(toolKey);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.SHEET_NAME);

    if (!sheet) {
        sheet = _App_ensureSheetExists(toolKey, postCreateCallback);
    } else {
        sheet.activate();
    }

    var html = _App_createTemplateFromFile(cfg.SIDEBAR_HTML).evaluate()
        .setTitle(cfg.TITLE)
        .setWidth(cfg.SIDEBAR_WIDTH || 300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function _App_launchTool(toolKey, postCreateCallback) {
    return Logger.run(toolKey, 'Launch Tool', function () {
        var cfg = SyncEngine.getTool(toolKey);
        var launchMode = cfg.LAUNCH_MODE || TOOL_LAUNCH_MODES.SIDEBAR;

        if (launchMode === TOOL_LAUNCH_MODES.MODAL) {
            var html = _App_createTemplateFromFile(cfg.MODAL_HTML || cfg.SIDEBAR_HTML).evaluate()
                .setTitle(cfg.TITLE)
                .setWidth(cfg.MODAL_WIDTH || cfg.SIDEBAR_WIDTH || 300)
                .setHeight(cfg.MODAL_HEIGHT || 600);
            SpreadsheetApp.getUi().showModalDialog(html, cfg.TITLE);
            return _App_ok('Modal opened successfully.');
        }

        _App_openSidebar(toolKey, postCreateCallback);
        return _App_ok('Sidebar opened successfully.');
    });
}

function _App_getMenuTools() {
    return Object.keys(SyncEngine.getAllTools())
        .map(function(key) { return SyncEngine.getTool(key); })
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
    return Logger.run(toolKey, 'Scaffold Sheet', function () {
        var cfg = SyncEngine.getTool(toolKey);
        if (!_App_canScaffoldSheet(cfg)) {
            throw new Error("Tool '" + toolKey + "' does not define a sheet schema and cannot be scaffolded automatically.");
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(cfg.SHEET_NAME);
        var isNew = !sheet;

        if (isNew) {
            sheet = ss.insertSheet(cfg.SHEET_NAME);
            Logger.info(cfg.TITLE, 'Scaffold', "Created new sheet: " + cfg.SHEET_NAME);
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
            catch (e) { 
                Logger.warn(cfg.TITLE, 'Post-Scaffold Callback', e.message);
            }
        }

        sheet.activate();
        return sheet;
    });
}


// --- FILE: core/12_UI.js ---
function onOpen() {
  return Logger.run('SYSTEM', 'Initialize UI', function () {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Workspace Sync Tools');

    _App_getMenuTools().forEach(function (cfg) {
      if (cfg.MENU_ENTRYPOINT) {
        menu.addItem(cfg.MENU_LABEL, cfg.MENU_ENTRYPOINT);
      }
    });

    menu.addToUi();
  });
}

function onInstall(e) {
  onOpen(e);
}

// Global settings and Theme Config have been moved to 00_Config_Constants.js and 01_Config_Theme.js
// so that Apps Script evaluates them before the rest of the files.


// ==========================================
// Dialog Server-Side Functions
// ==========================================

function UI_getProgress(toolKey) {
  return Logger.run('SYSTEM', 'Get Progress', function () {
    return _App_ok('Progress', _App_getProgress(toolKey));
  });
}

function UI_getHelpItems(toolKey) {
  return Logger.run('SYSTEM', 'Get Help Items', function () {
    var cfg = SyncEngine.getTool(toolKey);
    return _App_ok('Help Items retrieved', cfg.HELP_ITEMS || null);
  });
}

// ==========================================
// Centralized Validation / Status Checks
// ==========================================

/**
 * Checks for unsaved changes dynamically across any sheet.
 * Called from client sidebars via SyncSidebar wrappers.
 */
function UI_checkForUnsavedChanges(toolKey) {
  return Logger.run(toolKey, 'Check Unsaved', function () {
    var hasChanges = false;
    try {
      hasChanges = SheetManager.hasPendingActions(toolKey);
    } catch (e) {
      // Sheet does not exist or isn't scaffolded yet
    }
    var response = _App_ok('Check complete.', hasChanges);
    response.hasChanges = hasChanges;
    return response;
  });
}

/**
 * Centralized global action dispatcher to route sidebar calls to the SyncEngine.
 * Exposes a single, unified entrypoint for client-side invocations.
 */
function SyncEngine_executeAction(toolKey, actionName, argsJson) {
  var args = argsJson ? JSON.parse(argsJson) : [];
  return SyncEngine.runAction(toolKey, actionName, args);
}



// --- FILE: tools/BulkFolderCreation/Code.js ---
/**
 * Bulk Folder Creation
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('BULK_FOLDER', {
    SHEET_NAME: SHEET_NAMES.BULK_FOLDER,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.BULK_FOLDER,
    MENU_LABEL: SHEET_NAMES.BULK_FOLDER,
    MENU_ENTRYPOINT: 'BulkFolderCreation_openSidebar',
    MENU_ORDER: 80,
    SIDEBAR_HTML: 'tools/BulkFolderCreation/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Level 1', type: 'TEXT' },
            { header: 'Level 2', type: 'TEXT' },
            { header: 'Level 3', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to create folders in bulk:</p><ol><li><strong>Select Destination:</strong> Pick the parent folder in the sidebar explorer.</li><li><strong>Set Action:</strong> Set the Action column to <code>CREATE</code>.</li><li><strong>Enter Levels:</strong> Provide folder names in Level 1, Level 2, Level 3 to define nested structures.</li><li><strong>Run:</strong> Click <strong>Run Bulk Creation</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Action, Status, and Levels.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>CREATE</code> to create folders on Drive.</li><li><strong>Level 1/2/3:</strong> Folders are created nested. e.g. <code>Level 1/Level 2/Level 3</code>.</li><li><strong>Duplicate Check:</strong> If a folder structure already exists, the tool will step inside rather than duplicating it.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getDriveNavData: function(folderId) {
            try {
                var folder;
                if (!folderId || folderId === 'root') {
                    folder = DriveApp.getRootFolder();
                } else {
                    folder = DriveApp.getFolderById(folderId);
                }

                var currentId = folder.getId();
                var currentName = folder.getName();

                var breadcrumbs = [];
                var parent = folder;
                var depth = 0;
                while (depth < 5) {
                    try {
                        breadcrumbs.unshift({ id: parent.getId(), name: parent.getName() });
                        var parents = parent.getParents();
                        if (parents.hasNext()) {
                            parent = parents.next();
                        } else {
                            break;
                        }
                    } catch (e) {
                        break;
                    }
                    depth++;
                }

                var folders = folder.getFolders();
                var folderList = [];
                while (folders.hasNext()) {
                    var f = folders.next();
                    folderList.push({
                        id: f.getId(),
                        name: f.getName()
                    });
                }

                folderList.sort(function (a, b) { return a.name.localeCompare(b.name); });

                return _App_ok('Navigation data loaded', {
                    current: { id: currentId, name: currentName },
                    breadcrumbs: breadcrumbs,
                    children: folderList
                });

            } catch (e) {
                throw new Error("Error fetching Drive data: " + e.message);
            }
        },
        createFolders: function(targetFolderId) {
            return _BulkFolderCreation_runBulkCreationSequence(targetFolderId);
        }
    }
});

// Column-index aliases — kept for backward compatibility.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('BULK_FOLDER').
var BULKFOLDER_COL = {
  ACTION: 0,
  STATUS: 1
};

// --- MENU & UI HANDLERS ---

/** Opens the Bulk Folder sidebar and ensures the sheet exists. */
function BulkFolderCreation_openSidebar() {
  return Logger.run('BULK_FOLDER', 'Open Sidebar', function () {
    _App_launchTool('BULK_FOLDER');
  });
}

// --- EXPLORER LOGIC ---

function _BulkFolderCreation_runBulkCreationSequence(targetFolderId) {
  return Logger.run('BULK_FOLDER', 'Batch Creation', function () {
    return _App_withDocumentLock('BULK_FOLDER_CREATION', function () {
      _App_resetExecutionTimer();
      
      var pendingRows = SheetManager.readPendingObjects('BULK_FOLDER');

      if (pendingRows.length === 0) {
        Logger.warn(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Global', "No pending 'CREATE' actions found.");
        return _App_ok("No pending 'CREATE' actions found.");
      }

      var headers = SheetManager.getHeaders('BULK_FOLDER');
      var levelCols = headers.filter(h => h.toLowerCase().startsWith('level'));

      // --- PRE-VALIDATION START ---
      var gapErrors = [];
      var emptyErrors = [];
      for (var k = 0; k < pendingRows.length; k++) {
        var item = pendingRows[k];
        var rowNum = item._rowNumber;
        var hasEmptyLevel = false;
        var hasDataAfterEmpty = false;
        var hasAnyData = false;

        for (var c = 0; c < levelCols.length; c++) {
          var header = levelCols[c];
          var fName = String(item[header] || "").trim();

          if (fName === "") {
            hasEmptyLevel = true;
          } else {
            hasAnyData = true;
            if (hasEmptyLevel) {
              hasDataAfterEmpty = true;
              break;
            }
          }
        }

        if (!hasAnyData) {
          emptyErrors.push("Row " + rowNum);
        } else if (hasDataAfterEmpty) {
          gapErrors.push("Row " + rowNum);
        }
      }

      if (gapErrors.length > 0 || emptyErrors.length > 0) {
        var errMsgs = [];
        if (gapErrors.length > 0) {
          errMsgs.push("Missing intermediate folder names in rows: " + gapErrors.join(", ") + " (e.g., Level 1 is empty, but Level 2 has data)");
        }
        if (emptyErrors.length > 0) {
          errMsgs.push("No folder names specified in rows: " + emptyErrors.join(", "));
        }
        var fullError = "⚠️ Validation Error:\n" + errMsgs.join("\n");
        Logger.warn(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Pre-Validation', fullError);
        return _App_fail(fullError + "\nPlease fix and try again.");
      }
      // --- PRE-VALIDATION END ---

      var folderCache = {};
      var stats = _App_BatchProcessor('BULK_FOLDER', pendingRows, function (item) {
        var rowNum = item._rowNumber;
        var folderNames = [];
        for (var c = 0; c < levelCols.length; c++) {
          var header = levelCols[c];
          var fName = String(item[header] || "").trim();
          if (fName) {
            folderNames.push(fName.replace(/[\\/?*]/g, "_"));
          }
        }

        if (folderNames.length === 0) {
          throw new Error("No folder names specified in Level columns.");
        }

        _BulkFolderCreation_createFolderPath(targetFolderId, folderNames, folderCache);

        return { _rowNumber: rowNum, status: _App_formatStatus('SUCCESS', 'Created: ' + folderNames.join('/')) };

      }, {
        onBatchComplete: function (results) {
          _App_batchPatchResults('BULK_FOLDER', results);
        }
      });

      var finalMsg = "Successfully processed " + stats.processedCount + " folders.";
      if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
      if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

      return _App_ok(finalMsg);
    });
  });
}

function _BulkFolderCreation_createFolderPath(baseFolderId, folderNamesArr, folderCache) {
  var currentParentId = baseFolderId === "root" ? DriveApp.getRootFolder().getId() : baseFolderId;

  for (var i = 0; i < folderNamesArr.length; i++) {
    var fName = folderNamesArr[i];
    var cacheKey = currentParentId + "_" + fName;

    // 1. Check in-memory Cache
    if (folderCache[cacheKey]) {
      currentParentId = folderCache[cacheKey];
      continue;
    }

    // 2. Check Drive API if not in Cache
    var query = "'" + currentParentId + "' in parents and name = '" + fName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";

    var result = _App_callWithBackoff(function () {
      return Drive.Files.list({ q: query, fields: "files(id, name)", pageSize: 1 });
    });

    if (result.files && result.files.length > 0) {
      currentParentId = result.files[0].id; // Step inside existing
    } else {
      // 3. Create it
      var resource = {
        name: fName,
        parents: [currentParentId],
        mimeType: 'application/vnd.google-apps.folder'
      };
      var newFolder = _App_callWithBackoff(function () {
        return Drive.Files.create(resource, null, { fields: 'id' });
      });
      currentParentId = newFolder.id;
    }

    // Save to Cache for subsequent rows
    folderCache[cacheKey] = currentParentId;
  }
  return currentParentId;
}


// --- FILE: tools/CalendarSync/Code.js ---
/**
 * Google Calendar
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CALENDAR_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Calendar API', test: function() { return typeof Calendar !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CALENDAR_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.CALENDAR_SYNC,
    MENU_LABEL: SHEET_NAMES.CALENDAR_SYNC,
    MENU_ENTRYPOINT: 'CalendarSync_openSidebar',
    MENU_ORDER: 10,
    SIDEBAR_HTML: 'tools/CalendarSync/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Calendar Name', type: 'DROPDOWN', options: function() { try { return CalendarApp.getAllCalendars().map(function(c){return c.getName()}); } catch(e) { return []; } } },
            { header: 'Event Title', type: 'TEXT' },
            { header: 'Start Time', type: 'TEXT' },
            { header: 'End Time', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Location', type: 'TEXT' },
            { header: 'Add Meet?', type: 'CHECKBOX' },
            { header: 'Guests', type: 'EMAIL_LIST' },
            { header: 'Color', type: 'DROPDOWN', options: function() { try { return ['Default'].concat(Object.keys(CalendarApp.EventColor)); } catch(e){ return ['Default']; } } },
            { header: 'Visibility', type: 'DROPDOWN', options: ['Default', 'Public', 'Private'] },
            { header: 'Event ID', type: 'ID' },
            { header: 'Calendar ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to sync events:</p><ol><li><strong>Define Actions:</strong> Set the Action column to <code>CREATE</code>, <code>UPDATE</code>, or <code>DELETE</code>.</li><li><strong>Specify Times:</strong> Ensure Start Time and End Time use standard date/time formats (e.g. <code>MM/dd/yyyy HH:mm:ss</code>).</li><li><strong>Push:</strong> Click <strong>Push Changes</strong> in the sidebar to synchronize with Google Calendar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Understand event colors, guest emails, and Meet integration.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Calendar Name:</strong> Must match one of your accessible calendars.</li><li><strong>Add Meet?:</strong> Tick to generate a Google Meet link for the event.</li><li><strong>Guests:</strong> Comma-separated list of guest email addresses.</li><li><strong>Event ID / Calendar ID:</strong> Read-only IDs generated by Google Calendar. Do not manually edit.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "Sync range limits, importing, and moving events.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>Importing:</strong> Use <strong>Pull Events</strong> to import existing events from Google Calendar into the sheet.</li><li><strong>Moving Events:</strong> To move an event, change its <code>Calendar Name</code> and select <code>UPDATE</code> action.</li><li><strong>Time Formats:</strong> Supported formats include <code>YYYY-MM-DD HH:MM:SS</code> and standard regional formats.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getLoadData: function () {
            try {
                var allCalendars = _App_callWithBackoff(function () {
                    return CalendarApp.getAllCalendars();
                });
                var seen = {};
                var uniqueCals = [];

                (allCalendars || []).forEach(function (c) {
                    var calId = c && c.getId ? c.getId() : '';
                    if (!calId || seen[calId]) return;
                    seen[calId] = true;
                    uniqueCals.push({
                        id: calId,
                        name: c.getName(),
                        color: c.getColor()
                    });
                });

                var savedCalIds = _App_getProperty(APP_PROPS.CAL_SELECTED_IDS);
                if (!Array.isArray(savedCalIds)) savedCalIds = [];

                return _App_ok('Calendar load data ready.', {
                    calendars: uniqueCals,
                    savedCalIds: savedCalIds,
                    savedStartDate: _App_getProperty(APP_PROPS.CAL_START_DATE),
                    savedEndDate: _App_getProperty(APP_PROPS.CAL_END_DATE)
                });
            } catch (err) {
                throw new Error('Unable to load calendars. ' + err.message);
            }
        },
        savePreferences: function (calIds, startStr, endStr) {
            if (calIds) _App_setProperty(APP_PROPS.CAL_SELECTED_IDS, calIds);
            if (startStr !== undefined) _App_setProperty(APP_PROPS.CAL_START_DATE, startStr);
            if (endStr !== undefined) _App_setProperty(APP_PROPS.CAL_END_DATE, endStr);
            return _App_ok('Preferences saved.');
        },
        pull: function (request) {
            var TARGET_SHEET_NAME = SHEET_NAMES.CALENDAR_SYNC;
            var allCals = _App_callWithBackoff(function () { return CalendarApp.getAllCalendars(); });

            // Fetch Events
            var start = new Date(request.startDate);
            var end = new Date(request.endDate);
            var outputObjects = [];

            allCals.forEach(function (cal) {
                try {
                    var calName = cal.getName();
                    var calId = cal.getId();
                    var events = _App_callWithBackoff(function () { return cal.getEvents(start, end); });
                    events.forEach(function (e) {
                        outputObjects.push({
                            'Action': "",
                            'Calendar Name': calName,
                            'Event Title': e.getTitle(),
                            'Start Time': _App_formatDateTime(e.getStartTime()),
                            'End Time': _App_formatDateTime(e.getEndTime()),
                            'Description': e.getDescription(),
                            'Location': e.getLocation(),
                            'Add Meet?': false,
                            'Guests': e.getGuestList().map(function (g) { return g.getEmail(); }).join(","),
                            'Color': "Default",
                            'Visibility': "Default",
                            'Event ID': e.getId(),
                            'Calendar ID': calId
                        });
                    });
                } catch (err) {
                }
            });

            // Sort by Calendar Name
            outputObjects.sort(function (a, b) {
                var nameA = (a['Calendar Name'] || "").toLowerCase();
                var nameB = (b['Calendar Name'] || "").toLowerCase();
                if (nameA < nameB) return -1;
                if (nameA > nameB) return 1;
                return 0;
            });

            // Populate Sheet via DAO
            SheetManager.overwriteObjects('CALENDAR_SYNC', outputObjects);

            SyncEngine.getTool('CALENDAR_SYNC').ACTIONS.savePreferences(null, request.startDate, request.endDate);
            var summary = 'Successfully imported ' + outputObjects.length + " events into '" + TARGET_SHEET_NAME + "'.";
            return _App_ok(summary);
        },
        push: function () {
            var pendingItems = SheetManager.readPendingObjects('CALENDAR_SYNC');

            if (pendingItems.length === 0) return _App_ok("No pending actions found.");

            var allCals = CalendarApp.getAllCalendars();
            var calMap = new Map();
            var calObjMap = new Map();

            allCals.forEach(function (c) {
                calMap.set(c.getName(), c.getId());
                calObjMap.set(c.getId(), c);
            });

            var stats = _App_BatchProcessor('CALENDAR_SYNC', pendingItems, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    eventId: item['Event ID'] ? String(item['Event ID']) : null,
                    calId: item['Calendar ID'] ? String(item['Calendar ID']) : null,
                    status: "",
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().toUpperCase();
                var targetCalName = item['Calendar Name'];
                var targetCalId = calMap.get(targetCalName);

                var eventData = {
                    title: item['Event Title'],
                    start: item['Start Time'],
                    end: item['End Time'],
                    desc: item['Description'],
                    loc: item['Location'],
                    meet: item['Add Meet?'],
                    guests: item['Guests'],
                    color: item['Color'],
                    visibility: item['Visibility']
                };

                if (!(eventData.start instanceof Date)) eventData.start = new Date(eventData.start);
                if (!(eventData.end instanceof Date)) eventData.end = new Date(eventData.end);

                if (isNaN(eventData.start.getTime())) throw new Error("⚠️ Data Error: Invalid Start Time format");
                if (isNaN(eventData.end.getTime())) throw new Error("⚠️ Data Error: Invalid End Time format");
                if (eventData.end <= eventData.start) throw new Error("⚠️ Data Error: End Time cannot be before or equal to Start Time");
                if (!eventData.title) throw new Error("⚠️ Data Error: Missing Event Title");

                if (eventData.guests && !_App_validateEmailList(eventData.guests)) {
                    throw new Error("⚠️ Data Error: Invalid guest email address(es)");
                }

                switch (action) {
                    case "CREATE":
                        if (!targetCalName) throw new Error("⚠️ Data Error: Missing Calendar Name");
                        if (!targetCalId) throw new Error("⚠️ Data Error: Calendar '" + targetCalName + "' not found");

                        var createCal = calObjMap.get(targetCalId);
                        if (!createCal) throw new Error("❌ API Error: Target calendar object is null");

                        var newEvent = _App_callWithBackoff(function () {
                            return createCal.createEvent(eventData.title, eventData.start, eventData.end, {
                                description: eventData.desc,
                                location: eventData.loc,
                                guests: eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).join(',') : ""
                            });
                        });

                        var optionErr = _CalendarSync_applyEventOptions(newEvent, eventData);
                        if (eventData.meet === true || eventData.meet === 'TRUE') {
                            try { _CalendarSync_addMeetLinkToEvent(targetCalId, newEvent.getId()); }
                            catch (meetErr) { optionErr = optionErr ? optionErr + ", Meet: " + meetErr.message : "Meet: " + meetErr.message; }
                        }

                        rowUpdates.eventId = newEvent.getId();
                        rowUpdates.calId = targetCalId;
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Created") + (optionErr ? " (" + _App_formatStatus('WARNING', optionErr) + ")" : "");
                        rowUpdates.action = "";
                        break;

                    case "UPDATE":
                        if (!rowUpdates.eventId) throw new Error("⚠️ Data Error: Missing Event ID");

                        // Identity Check: If target calendar name doesn't match Calendar ID, perform MOVE
                        if (targetCalId && rowUpdates.calId && targetCalId !== rowUpdates.calId) {
                            rowUpdates = _CalendarSync_processMove(rowUpdates, calObjMap, targetCalId, eventData);
                            break;
                        }

                        var updateCal = calObjMap.get(rowUpdates.calId);
                        if (!updateCal) throw new Error("⚠️ Data Error: Calendar ID not found on your account");

                        var eventToUpdate = null;
                        try { eventToUpdate = _App_callWithBackoff(function() { return updateCal.getEventById(rowUpdates.eventId); }); } catch(e) {}

                        if (!eventToUpdate) throw new Error("⚠️ Data Error: Event ID not found on calendar");

                        _App_callWithBackoff(function () {
                            if (eventToUpdate.getTitle() !== eventData.title) {
                                eventToUpdate.setTitle(eventData.title);
                            }
                            var currentStart = eventToUpdate.getStartTime();
                            var currentEnd = eventToUpdate.getEndTime();
                            if (currentStart.getTime() !== eventData.start.getTime() || currentEnd.getTime() !== eventData.end.getTime()) {
                                eventToUpdate.setTime(eventData.start, eventData.end);
                            }
                            if (eventToUpdate.getDescription() !== eventData.desc) {
                                eventToUpdate.setDescription(eventData.desc);
                            }
                            if (eventToUpdate.getLocation() !== eventData.loc) {
                                eventToUpdate.setLocation(eventData.loc);
                            }
                        });

                        var updateOptionErr = _CalendarSync_applyEventOptions(eventToUpdate, eventData);
                        if (eventData.meet === true || eventData.meet === 'TRUE') {
                            try { _CalendarSync_addMeetLinkToEvent(rowUpdates.calId || eventToUpdate.getOriginalCalendarId(), rowUpdates.eventId); }
                            catch (meetErr) { updateOptionErr = updateOptionErr ? updateOptionErr + ", Meet: " + meetErr.message : "Meet: " + meetErr.message; }
                        }

                        var currentGuests = eventToUpdate.getGuestList();
                        var targetGuests = eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).filter(function (g) { return g !== ""; }) : [];
                        var currentEmails = currentGuests.map(function(g) { return g.getEmail().toLowerCase(); });

                        currentGuests.forEach(function (guestObj) {
                            var email = guestObj.getEmail();
                            if (targetGuests.map(function(t) { return t.toLowerCase(); }).indexOf(email.toLowerCase()) === -1) {
                                eventToUpdate.removeGuest(email);
                            }
                        });
                        targetGuests.forEach(function (email) {
                            if (currentEmails.indexOf(email.toLowerCase()) === -1) {
                                eventToUpdate.addGuest(email);
                            }
                        });

                        rowUpdates.status = _App_formatStatus('SUCCESS', "Updated") + (updateOptionErr ? " (" + _App_formatStatus('WARNING', updateOptionErr) + ")" : "");
                        rowUpdates.action = "";
                        break;

                    case "DELETE":
                        if (!rowUpdates.eventId) throw new Error("⚠️ Data Error: Missing Event ID");
                        var delCal = calObjMap.get(rowUpdates.calId);
                        if (!delCal) throw new Error("⚠️ Data Error: Original Calendar inaccessible");

                        var eventToDel = null;
                        try { eventToDel = _App_callWithBackoff(function() { return delCal.getEventById(rowUpdates.eventId); }); } catch(e) {}

                        if (eventToDel) {
                            _App_callWithBackoff(function () { eventToDel.deleteEvent(); });
                            rowUpdates.status = _App_formatStatus('SUCCESS', "Deleted");
                            rowUpdates.action = "";
                        } else {
                            rowUpdates.status = _App_formatStatus('WARNING', "Already Deleted (Event not found)");
                            rowUpdates.action = "";
                        }
                        break;

                    default:
                        rowUpdates.status = _App_formatStatus('WARNING', "Unknown Action '" + action + "'");
                }

                return rowUpdates;

            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('CALENDAR_SYNC', batchResults, function (res) {
                        return {
                            'Event ID': res.eventId,
                            'Calendar ID': res.calId
                        };
                    });
                }
            });

            return _App_ok("Sync Complete. Processed: " + stats.processedCount);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Calendar sidebar and ensures the sheet exists. */
function CalendarSync_openSidebar() {
  return Logger.run('CALENDAR_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CALENDAR_SYNC');
  });
}

// --- HELPER VALIDATORS ---

function _CalendarSync_processMove(rowUpdates, calObjMap, targetCalId, eventData) {
  var oldCal = calObjMap.get(rowUpdates.calId);
  var newCal = calObjMap.get(targetCalId);

  if (!newCal) throw new Error("⚠️ Data Error: Target calendar not accessible");

  var newEvent = newCal.createEvent(eventData.title, eventData.start, eventData.end, {
    description: eventData.desc,
    location: eventData.loc,
    guests: eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).join(',') : ""
  });
  _CalendarSync_applyEventOptions(newEvent, eventData);

  var meetWarning = "";
  if (eventData.meet === true || eventData.meet === 'TRUE') {
    try { _CalendarSync_addMeetLinkToEvent(targetCalId, newEvent.getId()); }
    catch (meetErr) { meetWarning = " (⚠️ Meet: " + meetErr.message + ")"; }
  }

  var deleteWarning = "";
  if (oldCal && rowUpdates.eventId) {
    try {
      var oldEvent = null;
      try { oldEvent = _App_callWithBackoff(function() { return oldCal.getEventById(rowUpdates.eventId); }); } catch(e) {}
      if (oldEvent) oldEvent.deleteEvent();
      else deleteWarning = " (⚠️ Old event not found)";
    } catch (delErr) {
      deleteWarning = " (⚠️ Could not delete old event: " + delErr.message + ")";
    }
  } else {
    deleteWarning = " (⚠️ Old calendar inaccessible)";
  }

  rowUpdates.eventId = newEvent.getId();
  rowUpdates.calId = targetCalId;
  rowUpdates.status = _App_formatStatus('SUCCESS', "Moved") + deleteWarning + meetWarning;
  rowUpdates.action = "";
  return rowUpdates;
}

function _CalendarSync_applyEventOptions(event, data) {
  var warning = null;
  if (data.color && data.color !== 'Default') {
    if (CalendarApp.EventColor[data.color]) {
      try {
        var targetColor = CalendarApp.EventColor[data.color];
        if (event.getColor() !== targetColor) {
          event.setColor(targetColor);
        }
      } catch (e) { warning = "Color set failed"; }
    } else {
      warning = "Invalid Color";
    }
  }
  if (data.visibility) {
    try {
      var targetVisibility = null;
      if (data.visibility === 'Public') targetVisibility = CalendarApp.Visibility.PUBLIC;
      else if (data.visibility === 'Private') targetVisibility = CalendarApp.Visibility.PRIVATE;
      
      if (targetVisibility !== null && event.getVisibility() !== targetVisibility) {
        event.setVisibility(targetVisibility);
      }
    } catch (e) {
      warning = warning ? warning + ", Visibility set failed" : "Visibility set failed";
    }
  }
  return warning;
}

// --- UTILITIES ---

function _CalendarSync_addMeetLinkToEvent(calendarId, eventId) {
  if (typeof Calendar === 'undefined') {
    throw new Error("Enable 'Google Calendar API' in Services");
  }
  Calendar.Events.patch({
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    }
  }, calendarId, eventId, { conferenceDataVersion: 1 });
}


// --- FILE: tools/ChatSpaceSync/Code.js ---
/**
 * Google Chat Space Sync Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CHAT_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Chat API', test: function() { return typeof Chat !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CHAT_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.CHAT_SYNC,
    MENU_LABEL: SHEET_NAMES.CHAT_SYNC,
    MENU_ENTRYPOINT: 'ChatSpaceSync_openSidebar',
    MENU_ORDER: 15,
    SIDEBAR_HTML: 'tools/ChatSpaceSync/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['ADD_MEMBER', 'REMOVE_MEMBER'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Space Name', type: 'TEXT' },
            { header: 'Member Email', type: 'TEXT' },
            { header: 'Role', type: 'DROPDOWN', options: ['ROLE_MEMBER', 'ROLE_MANAGER'] },
            { header: 'Type', type: 'TEXT' }, // User or Group
            { header: 'Space ID', type: 'ID' },
            { header: 'Membership ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to sync Chat Space members:</p><ol><li><strong>Select Spaces:</strong> Check target spaces in the sidebar.</li><li><strong>Pull:</strong> Click <strong>Pull Members</strong> to import current members.</li><li><strong>Modify:</strong> Set action to <code>ADD_MEMBER</code> or <code>REMOVE_MEMBER</code>.</li><li><strong>Push:</strong> Click <strong>Push Changes</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Action, Role, and Member Email.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>ADD_MEMBER</code> to invite a user or <code>REMOVE_MEMBER</code> to evict.</li><li><strong>Role:</strong> Choose <code>ROLE_MEMBER</code> or <code>ROLE_MANAGER</code>.</li><li><strong>Type:</strong> Read-only user category (User, Group, Bot).</li><li><strong>IDs:</strong> System-generated IDs. Do not manually edit.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getLoadData: function () {
            try {
                var spacesList = [];
                var pageToken = null;
                do {
                    var response = _App_callWithBackoff(function() {
                        return Chat.Spaces.list({ pageToken: pageToken });
                    });
                    if (response.spaces) {
                        spacesList = spacesList.concat(response.spaces);
                    }
                    pageToken = response.nextPageToken;
                } while (pageToken);

                var uniqueSpaces = spacesList.map(function (s) {
                    return {
                        id: s.name,
                        name: s.displayName || s.name
                    };
                });

                var savedSpaceIds = _App_getProperty(APP_PROPS.CHAT_SELECTED_SPACES);
                if (!Array.isArray(savedSpaceIds)) savedSpaceIds = [];

                return _App_ok('Chat spaces load data ready.', {
                    spaces: uniqueSpaces,
                    savedSpaceIds: savedSpaceIds
                });
            } catch (err) {
                throw new Error('Unable to load chat spaces: ' + err.message);
            }
        },
        savePreferences: function (spaceIds) {
            if (spaceIds) _App_setProperty(APP_PROPS.CHAT_SELECTED_SPACES, spaceIds);
            return _App_ok('Preferences saved.');
        },
        pull: function () {
            return _ChatSpaceSync_pullMembers();
        },
        push: function () {
            return _ChatSpaceSync_pushChanges();
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Chat Sync sidebar and ensures the sheet exists. */
function ChatSpaceSync_openSidebar() {
  return Logger.run('CHAT_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CHAT_SYNC');
  });
}

// --- PREFERENCES & STATE ---

function _ChatSpaceSync_pullMembers() {
  return Logger.run('CHAT_SYNC', 'Pull Members', function () {
    return _App_withDocumentLock('CHAT_SYNC_PULL', function () {
      var TARGET_SHEET_NAME = SHEET_NAMES.CHAT_SYNC;
      var sheet = _App_ensureSheetExists('CHAT_SYNC');

    var outputObjects = [];
    var spacesList = [];
    var pageToken = null;

    var savedSpaceIds = _App_getProperty(APP_PROPS.CHAT_SELECTED_SPACES);
    if (!Array.isArray(savedSpaceIds)) savedSpaceIds = [];

    if (savedSpaceIds.length > 0) {
      // Pull only selected spaces
      savedSpaceIds.forEach(function (spaceId) {
        try {
          var space = _App_callWithBackoff(function () {
            return Chat.Spaces.get(spaceId);
          });
          if (space) spacesList.push(space);
        } catch (err) {
          Logger.warn('CHAT_SYNC', 'Fetch Space Error', 'Space ' + spaceId + ': ' + err.message);
        }
      });
    } else {
      // Fetch all spaces the user is a member of
      do {
        var response = _App_callWithBackoff(function() {
            return Chat.Spaces.list({
              pageToken: pageToken
            });
        });
        
        if (response.spaces) {
          spacesList = spacesList.concat(response.spaces);
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
    }

    spacesList.forEach(function (space) {
      try {
        var spaceNameId = space.name;
        var spaceDisplayName = space.displayName || space.name;
        var memberPageToken = null;
        var members = [];

        do {
            var memberResponse = _App_callWithBackoff(function() {
                return Chat.Spaces.Members.list(spaceNameId, {
                    pageToken: memberPageToken
                });
            });
            if (memberResponse.memberships) {
                members = members.concat(memberResponse.memberships);
            }
            memberPageToken = memberResponse.nextPageToken;
        } while (memberPageToken);

        members.forEach(function (m) {
          var memberEmail = "";
          var memberType = "Unknown";
          
          if (m.member && m.member.type === "HUMAN") {
              memberEmail = m.member.displayName || m.member.name;
              memberType = "User";
          } else if (m.groupMember) {
              memberEmail = m.groupMember.id;
              memberType = "Group";
          } else if (m.member && m.member.type === "BOT") {
              memberEmail = m.member.displayName || "Bot";
              memberType = "Bot";
          }

          outputObjects.push({
            'Action': "",
            'Status': "",
            'Space Name': spaceDisplayName,
            'Member Email': memberEmail,
            'Role': m.role === 'ROLE_MANAGER' ? 'ROLE_MANAGER' : 'ROLE_MEMBER',
            'Type': memberType,
            'Space ID': spaceNameId,
            'Membership ID': m.name
          });
        });
      } catch (err) {
        throw new Error('Pull Members failed for ' + space.name + ': ' + err.message);
      }
    });

    // Sort by Space Name alphabetically
    outputObjects.sort(function(a, b) {
        return a['Space Name'].localeCompare(b['Space Name']);
    });

      SheetManager.overwriteObjects('CHAT_SYNC', outputObjects);
      
      var summary = 'Successfully imported ' + outputObjects.length + " members into '" + TARGET_SHEET_NAME + "'.";
      return _App_ok(summary);
    });
  });
}

// --- THE "PUSH" WORKFLOW ---

function _ChatSpaceSync_pushChanges() {
  return Logger.run('CHAT_SYNC', 'Push Changes', function () {
    return _App_withDocumentLock('CHAT_SYNC_PUSH', function () {
      var pendingItems = SheetManager.readPendingObjects('CHAT_SYNC');

      if (pendingItems.length === 0) return _App_ok("No pending actions found.");

    var stats = _App_BatchProcessor('CHAT_SYNC', pendingItems, function (item) {
      var rowUpdates = {
        action: item['Action'],
        membershipId: item['Membership ID'] ? String(item['Membership ID']) : null,
        status: "",
        _rowNumber: item._rowNumber
      };

      var action = rowUpdates.action.toString().toUpperCase();
        var targetEmail = item['Member Email'];
        var targetRole = item['Role'] || 'ROLE_MEMBER';
        var spaceId = item['Space ID'];

        if (!spaceId) throw new Error("⚠️ Data Error: Missing Space ID");

        switch (action) {
          case "ADD_MEMBER":
            if (!targetEmail) throw new Error("⚠️ Data Error: Missing Member Email");
            
            var membership = {
              member: {
                name: "users/" + targetEmail,
                type: "HUMAN"
              },
              role: targetRole
            };

            var newMembership = _App_callWithBackoff(function () {
              return Chat.Spaces.Members.create(membership, spaceId);
            });

            rowUpdates.membershipId = newMembership.name;
            rowUpdates.status = _App_formatStatus('SUCCESS', "Added");
            rowUpdates.action = "";
            break;

          case "REMOVE_MEMBER":
            if (!rowUpdates.membershipId) throw new Error("⚠️ Data Error: Missing Membership ID for REMOVE");
            
            _App_callWithBackoff(function () {
               Chat.Spaces.Members.remove(rowUpdates.membershipId);
            });
            
            rowUpdates.status = _App_formatStatus('SUCCESS', "Removed");
            rowUpdates.action = "";
            break;

          default:
            throw new Error("❓ Unknown Action '" + action + "'");
        }

        return rowUpdates;

    }, {
      onBatchComplete: function (batchResults) {
        _App_batchPatchResults('CHAT_SYNC', batchResults, function (res) {
          return {
            'Membership ID': res.membershipId
          };
        });
      }
    });

      return _App_ok("Sync Complete. Processed: " + stats.processedCount);
    });
  });
}


// --- FILE: tools/ContactsSync/Code.js ---
/**
 * Google Contacts
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CONTACTS_SYNC', {
    REQUIRED_SERVICES: [{ name: 'People API', test: function () { return typeof People !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.CONTACTS_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.CONTACTS_SYNC,
    MENU_LABEL: SHEET_NAMES.CONTACTS_SYNC,
    MENU_ENTRYPOINT: 'ContactsSync_openSidebar',
    MENU_ORDER: 20,
    SIDEBAR_HTML: 'tools/ContactsSync/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [
            { type: 'pending', actionCol: 'A', scope: 'actionOnly' },
            { type: 'custom', formula: '=AND($E2<>\'\', COUNTIF($E:$E, $E2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 5 },
            { type: 'custom', formula: '=AND($F2<>\'\', COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'First Name', type: 'TEXT' },
            { header: 'Last Name', type: 'TEXT' },
            { header: 'Email', type: 'EMAIL' },
            { header: 'Phone', type: 'TEXT' },
            { header: 'Company', type: 'TEXT' },
            { header: 'Job Title', type: 'TEXT' },
            { header: 'Starred', type: 'CHECKBOX' },
            { header: 'Street', type: 'TEXT' },
            { header: 'City', type: 'TEXT' },
            { header: 'State', type: 'TEXT' },
            { header: 'Zip', type: 'TEXT' },
            {
                header: 'Groups/Labels', type: 'DROPDOWN', allowInvalid: true, options: function () {
                    var groups = [];
                    try {
                        var response = _App_callWithBackoff(function () {
                            return People.ContactGroups.list({ pageSize: 1000 });
                        });
                        var excluded = ['Friends', 'Family', 'Coworkers', 'All Contacts', 'Starred'];
                        (response.contactGroups || []).forEach(function (g) {
                            var name = g.formattedName || g.name;
                            if (name && !excluded.includes(name)) {
                                groups.push(name);
                            }
                        });
                        groups.sort();
                    } catch (e) { }
                    return groups.length ? groups.slice(0, 499) : ['None'];
                }
            },
            { header: 'Notes', type: 'TEXT' },
            { header: 'Contact ID', type: 'ID', italic: true }
        ]
    },
    ACTIONS: {
        getMissingGroups: function () {
            var pendingItems = SheetManager.readPendingObjects('CONTACTS_SYNC');
            if (pendingItems.length === 0) return _App_ok('No pending actions.', []);

            var groupsInSheet = [];
            pendingItems.forEach(function (item) {
                var action = (item['Action'] || '').toString().toUpperCase();
                if (action === 'CREATE' || action === 'UPDATE') {
                    var groupsStr = item['Groups/Labels'] ? String(item['Groups/Labels']) : '';
                    if (groupsStr) {
                        var split = groupsStr.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
                        split.forEach(function (g) {
                            if (groupsInSheet.indexOf(g) === -1) groupsInSheet.push(g);
                        });
                    }
                }
            });

            if (groupsInSheet.length === 0) return _App_ok('No groups to check.', []);

            var groupsResponse = People.ContactGroups.list();
            var allGroups = groupsResponse.contactGroups || [];
            var existingNames = allGroups.map(function (g) { return g.formattedName || g.name; });

            var missing = groupsInSheet.filter(function (name) {
                return existingNames.indexOf(name) === -1;
            });

            return _App_ok('Missing groups identified.', missing);
        },
        savePreferences: function (groupIds) {
            if (groupIds) _App_setProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS, groupIds);
            return _App_ok('Preferences saved.');
        },
        pull: function (request) {
            if (typeof People === 'undefined') {
                throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
            }

            var sheet = SheetManager.ensureSheet('CONTACTS_SYNC');

            var outputData = [];
            var groupIds = request.groupIds || [];
            var pullAll = groupIds.includes('all');

            var groupsResponse = People.ContactGroups.list();
            var allGroups = groupsResponse.contactGroups || [];
            var groupMap = {};
            allGroups.forEach(function (g) {
                groupMap[g.resourceName] = g.formattedName || g.name;
            });

            var pageToken = null;
            var personFields = 'names,emailAddresses,phoneNumbers,organizations,memberships,biographies,addresses';

            do {
                var options = { pageSize: 1000, personFields: personFields };
                if (pageToken) options.pageToken = pageToken;

                var response;
                try {
                    response = _App_callWithBackoff(function () {
                        return People.People.Connections.list('people/me', options);
                    });
                } catch (err) {
                    throw new Error("API Error: " + err.message);
                }

                var connections = response.connections || [];

                connections.forEach(function (person) {
                    var pGroups = person.memberships ? person.memberships.map(function (m) {
                        return m.contactGroupMembership ? m.contactGroupMembership.contactGroupResourceName : null;
                    }).filter(function (g) { return g; }) : [];

                    var isInSelectedGroup = pullAll || pGroups.some(function (g) { return groupIds.includes(g); });

                    if (isInSelectedGroup) {
                        var firstName = "";
                        var lastName = "";
                        if (person.names && person.names.length > 0) {
                            var primaryName = person.names.find(function (n) { return n.metadata && n.metadata.primary; }) || person.names[0];
                            firstName = primaryName.givenName || "";
                            lastName = primaryName.familyName || "";
                        }

                        var email = _ContactsSync_getPrimary(person.emailAddresses);
                        var phone = _ContactsSync_getPrimary(person.phoneNumbers);

                        var company = "";
                        var title = "";
                        if (person.organizations && person.organizations.length > 0) {
                            var primaryOrg = person.organizations.find(function (o) { return o.metadata && o.metadata.primary; }) || person.organizations[0];
                            company = primaryOrg.name || "";
                            title = primaryOrg.title || "";
                        }

                        var notes = person.biographies && person.biographies.length > 0 ? (person.biographies[0].value || "") : "";

                        var isStarred = pGroups.includes('contactGroups/starred');

                        var street = "", city = "", state = "", zip = "";
                        if (person.addresses && person.addresses.length > 0) {
                            var primaryAddress = person.addresses.find(function (a) { return a.metadata && a.metadata.primary; }) || person.addresses[0];
                            street = primaryAddress.streetAddress || "";
                            city = primaryAddress.city || "";
                            state = primaryAddress.region || "";
                            zip = primaryAddress.postalCode || "";
                        }

                        var groupNames = pGroups.map(function (gId) { return groupMap[gId] || "Unknown Group"; }).join(", ");

                        outputData.push([
                            "", // Action
                            "", // Status
                            firstName,
                            lastName,
                            email,
                            phone,
                            company,
                            title,
                            isStarred,
                            street,
                            city,
                            state,
                            zip,
                            groupNames,
                            notes,
                            person.resourceName // Contact ID
                        ]);
                    }
                });
                pageToken = response.nextPageToken;
            } while (pageToken);

            // Apply body formatting using the registered tool config directly
            SheetManager.overwriteRows('CONTACTS_SYNC', outputData, {
                totalCols: SyncEngine.getTool('CONTACTS_SYNC').HEADERS.length,
                formatConfig: SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG
            });

            SyncEngine.getTool('CONTACTS_SYNC').ACTIONS.savePreferences(groupIds);
            return _App_ok('Successfully imported ' + outputData.length + " contacts.");
        },
        push: function () {
            if (typeof People === 'undefined') {
                throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
            }

            var groupsResponse = People.ContactGroups.list();
            var allGroups = groupsResponse.contactGroups || [];
            var groupNameToId = {};
            allGroups.forEach(function (g) {
                groupNameToId[g.formattedName || g.name] = g.resourceName;
            });

            var pendingRows = SheetManager.readPendingObjects('CONTACTS_SYNC');

            if (pendingRows.length === 0) return _App_ok("No pending actions found.");

            var groupAdditions = {};

            var stats = _App_BatchProcessor('CONTACTS_SYNC', pendingRows, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    contactId: item['Contact ID'],
                    status: "",
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().toUpperCase();
                var contactData = {
                    firstName: item['First Name'] !== "" ? String(item['First Name']) : "",
                    lastName: item['Last Name'] !== "" ? String(item['Last Name']) : "",
                    email: item['Email'] !== "" ? String(item['Email']) : "",
                    phone: item['Phone'] !== "" ? String(item['Phone']) : "",
                    company: item['Company'] !== "" ? String(item['Company']) : "",
                    title: item['Job Title'] !== "" ? String(item['Job Title']) : "",
                    starred: item['Starred'],
                    street: item['Street'] !== "" ? String(item['Street']) : "",
                    city: item['City'] !== "" ? String(item['City']) : "",
                    state: item['State'] !== "" ? String(item['State']) : "",
                    zip: item['Zip'] !== "" ? String(item['Zip']) : "",
                    groupsStr: item['Groups/Labels'] !== "" ? String(item['Groups/Labels']) : "",
                    notes: item['Notes'] !== "" ? String(item['Notes']) : ""
                };

                var person = { names: [], emailAddresses: [], phoneNumbers: [], organizations: [], biographies: [], addresses: [] };

                if (action !== "DELETE") {
                    if (contactData.firstName || contactData.lastName) {
                        person.names.push({ givenName: contactData.firstName || "", familyName: contactData.lastName || "" });
                    } else {
                        throw new Error("⚠️ Name is required to push.");
                    }

                    if (contactData.email) person.emailAddresses.push({ value: contactData.email });
                    if (contactData.phone) person.phoneNumbers.push({ value: contactData.phone });
                    if (contactData.company || contactData.title) person.organizations.push({ name: contactData.company || "", title: contactData.title || "" });
                    if (contactData.notes) person.biographies.push({ value: contactData.notes });
                    if (contactData.street || contactData.city || contactData.state || contactData.zip) {
                        person.addresses.push({
                            streetAddress: contactData.street || "",
                            city: contactData.city || "",
                            region: contactData.state || "",
                            postalCode: contactData.zip || ""
                        });
                    }
                }

                switch (action) {
                    case "CREATE":
                        var createdPerson = People.People.createContact(person);
                        rowUpdates.contactId = createdPerson.resourceName;
                        if (contactData.groupsStr) {
                            contactData.groupsStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (gName) {
                                var gId = _ContactsSync_getOrCreateGroup(gName, groupNameToId);
                                if (gId) {
                                    if (!groupAdditions[gId]) groupAdditions[gId] = [];
                                    groupAdditions[gId].push(createdPerson.resourceName);
                                }
                            });
                        }
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            if (!groupAdditions['contactGroups/starred']) groupAdditions['contactGroups/starred'] = [];
                            groupAdditions['contactGroups/starred'].push(createdPerson.resourceName);
                        }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Created");
                        rowUpdates.action = "";
                        break;

                    case "UPDATE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        var existingPerson = People.People.get(rowUpdates.contactId, { personFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses' });
                        person.etag = existingPerson.etag;

                        if (existingPerson.emailAddresses && existingPerson.emailAddresses.length > 0) {
                            var primaryMailIndex = existingPerson.emailAddresses.findIndex(function (e) { return e.metadata && e.metadata.primary; });
                            if (primaryMailIndex === -1) primaryMailIndex = 0;
                            var existingMails = existingPerson.emailAddresses;
                            if (contactData.email) existingMails[primaryMailIndex].value = contactData.email;
                            person.emailAddresses = existingMails;
                        }

                        if (existingPerson.phoneNumbers && existingPerson.phoneNumbers.length > 0) {
                            var primaryPhoneIndex = existingPerson.phoneNumbers.findIndex(function (p) { return p.metadata && p.metadata.primary; });
                            if (primaryPhoneIndex === -1) primaryPhoneIndex = 0;
                            var existingPhones = existingPerson.phoneNumbers;
                            if (contactData.phone) existingPhones[primaryPhoneIndex].value = contactData.phone;
                            person.phoneNumbers = existingPhones;
                        }

                        if (existingPerson.addresses && existingPerson.addresses.length > 0) {
                            var primaryAddressIndex = existingPerson.addresses.findIndex(function (a) { return a.metadata && a.metadata.primary; });
                            if (primaryAddressIndex === -1) primaryAddressIndex = 0;
                            var existingAddresses = existingPerson.addresses;
                            if (contactData.street || contactData.city || contactData.state || contactData.zip) {
                                var newAddr = { streetAddress: contactData.street || "", city: contactData.city || "", region: contactData.state || "", postalCode: contactData.zip || "" };
                                if (primaryAddressIndex > -1) { newAddr.metadata = existingAddresses[primaryAddressIndex].metadata; existingAddresses[primaryAddressIndex] = newAddr; }
                                else { newAddr.metadata = { primary: true }; existingAddresses.push(newAddr); }
                            }
                            person.addresses = existingAddresses;
                        }

                        People.People.updateContact(person, rowUpdates.contactId, { updatePersonFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses' });
                        if (contactData.groupsStr) {
                            contactData.groupsStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (gName) {
                                var gId = _ContactsSync_getOrCreateGroup(gName, groupNameToId);
                                if (gId) {
                                    if (!groupAdditions[gId]) groupAdditions[gId] = [];
                                    groupAdditions[gId].push(rowUpdates.contactId);
                                }
                            });
                        }
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            if (!groupAdditions['contactGroups/starred']) groupAdditions['contactGroups/starred'] = [];
                            groupAdditions['contactGroups/starred'].push(rowUpdates.contactId);
                        }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Updated");
                        rowUpdates.action = "";
                        break;

                    case "DELETE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        try { People.People.deleteContact(rowUpdates.contactId); } catch (e) { }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Deleted");
                        rowUpdates.action = "";
                        break;

                    default:
                        rowUpdates.status = _App_formatStatus('WARNING', "Unknown Action '" + action + "'");
                }

                return rowUpdates;

            }, {
                onBatchComplete: function (batchResults) {
                    Object.keys(groupAdditions).forEach(function (gId) {
                        var members = groupAdditions[gId];
                        if (members && members.length > 0) {
                            try {
                                _App_callWithBackoff(function () {
                                    People.ContactGroups.Members.modify({ resourceNamesToAdd: members }, gId);
                                });
                            } catch (e) {
                                Logger.warn('CONTACTS_SYNC', 'Batch Group Modify', 'Failed to add members to ' + gId + ': ' + e.message);
                            }
                        }
                    });
                    groupAdditions = {};

                    _App_batchPatchResults('CONTACTS_SYNC', batchResults, function (res) {
                        return {
                            'Contact ID': res.contactId
                        };
                    });
                }
            });

            return _App_ok("Sync Complete. Processed: " + stats.processedCount);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Contacts sidebar and ensures the sheet exists. */
function ContactsSync_openSidebar() {
    return Logger.run('CONTACTS_SYNC', 'Open Sidebar', function () {
        _App_launchTool('CONTACTS_SYNC');
    });
}


function _ContactsSync_getOrCreateGroup(gName, groupNameToId) {
    var id = groupNameToId[gName];
    if (!id) {
        try {
            var newGroup = _App_callWithBackoff(function () {
                return People.ContactGroups.create({
                    contactGroup: { name: gName }
                });
            });
            id = newGroup.resourceName;
            groupNameToId[gName] = id;
        } catch (e) {
            // Silently return null if creation fails
        }
    }
    return id;
}


/** Retrieves the primary field value, or falls back to the first item. */
function _ContactsSync_getPrimary(arr) {
    if (!arr || arr.length === 0) return "";
    var primaryItem = arr.find(function (item) { return item.metadata && item.metadata.primary; }) || arr[0];
    return primaryItem.value || "";
}


// --- FILE: tools/DocsMerge/Code.js ---
/**
 * Docs Merge
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('DOCS_MERGE', {
    SHEET_NAME: SHEET_NAMES.DOCS_MERGE,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.DOCS_MERGE,
    MENU_LABEL: SHEET_NAMES.DOCS_MERGE,
    MENU_ENTRYPOINT: 'DocsMerge_openSidebar',
    MENU_ORDER: 50,
    SIDEBAR_HTML: 'tools/DocsMerge/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Generate PDF', 'Generate Doc'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Document Name', type: 'TEXT' },
            { header: 'Merged File Link', type: 'URL' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to merge your first document:</p><ol><li><strong>Prepare Doc:</strong> Create a Google Doc with <code>{{placeholder}}</code> tags.</li><li><strong>Setup URLs:</strong> Search and select your Template Doc and Destination Folder.</li><li><strong>Sync & Run:</strong> Click <strong>Pull Placeholders</strong> to add columns, then click <strong>Run</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about Action, Document Name, and Merged File Link.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> <code>Generate PDF</code> or <code>Generate Doc</code>.</li><li><strong>Document Name:</strong> The filename for the generated document.</li><li><strong>Merged File Link:</strong> This cell will be updated with a link to the final file.</li><li><strong>Dynamic Columns:</strong> These are created automatically based on your <code>{{tags}}</code>.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "Single vs Individual mode and placeholder sync.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>Individual Mode:</strong> Creates a separate file for every row in the folder.</li><li><strong>Single Mode:</strong> Merges all rows into one giant document (with page breaks).</li><li><strong>Permissions:</strong> Ensure you have 'Editor' access to the template and destination folder.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getConfig: function () {
            var templateUrl = _App_getProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_URL) || "";
            var folderUrl = _App_getProperty(APP_PROPS.DOCS_MERGE_FOLDER_URL) || "";

            var templateName = _App_getProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME) || "";
            if (templateUrl && !templateName) {
                try {
                    templateName = DriveApp.getFileById(_App_extractIdFromUrl(templateUrl)).getName();
                    _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME, templateName);
                } catch (e) { }
            }

            var folderName = _App_getProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME) || "";
            if (folderUrl && !folderName) {
                try {
                    folderName = DriveApp.getFolderById(_App_extractIdFromUrl(folderUrl)).getName();
                    _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME, folderName);
                } catch (e) { }
            }

            return _App_ok('Configuration loaded.', {
                templateUrl: templateUrl,
                folderUrl: folderUrl,
                templateName: templateName,
                folderName: folderName
            });
        },
        saveConfig: function (config) {
            if (config.templateUrl !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_URL, config.templateUrl);
            if (config.folderUrl !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_URL, config.folderUrl);
            if (config.templateName !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME, config.templateName);
            if (config.folderName !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME, config.folderName);
            return _App_ok('Config saved.');
        },
        searchFolders: function (query) {
            if (!query || query.length < 2) return _App_ok('Folder search skipped.', { results: [] });
            var results = [];
            try {
                var folders = DriveApp.searchFolders("title contains '" + query.replace(/'/g, "\\'") + "'");
                var count = 0;
                while (folders.hasNext() && count < 10) {
                    var f = folders.next();
                    results.push({
                        id: f.getId(),
                        name: f.getName(),
                        url: f.getUrl()
                    });
                    count++;
                }
            } catch (e) { }
            return _App_ok('Folder search complete.', { results: results });
        },
        searchDocs: function (query) {
            if (!query || query.length < 2) return _App_ok('Document search skipped.', { results: [] });
            var results = [];
            try {
                var files = DriveApp.searchFiles("mimeType = 'application/vnd.google-apps.document' and title contains '" + query.replace(/'/g, "\\'") + "'");
                var count = 0;
                while (files.hasNext() && count < 10) {
                    var f = files.next();
                    results.push({
                        id: f.getId(),
                        name: f.getName(),
                        url: f.getUrl()
                    });
                    count++;
                }
            } catch (e) { }
            return _App_ok('Document search complete.', { results: results });
        },
        syncPlaceholders: function (templateUrl) {
            var templateId = _App_extractIdFromUrl(templateUrl);
            if (!templateId) throw new Error("Could not extract Template ID. Paste the full Doc URL.");

            try {
                var docText = DocumentApp.openById(templateId).getBody().getText();
                var placeholders = [];
                var regex = /\{\{([^{}]+)\}\}/g;
                var match;

                while ((match = regex.exec(docText)) !== null) {
                    if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
                }

                var syncResult = SheetManager.syncDynamicColumns('DOCS_MERGE', placeholders, {
                    anchorHeader: 'Merged File Link',
                    dynamicColWidth: 150
                });

                return _App_ok('Synced ' + placeholders.length + ' placeholders.', {
                    placeholders: placeholders,
                    headers: syncResult.headers
                });
            } catch (e) {
                return _App_fail("Sync failed: " + e.message + ". Ensure you have editor access to the Doc.");
            }
        },
        executeBatch: function (config) {
            var mode = config.mode || "INDIVIDUAL";

            var pendingRows = SheetManager.readPendingObjects('DOCS_MERGE', { useDisplayValues: true });

            if (pendingRows.length === 0) {
                return _App_ok("Nothing to do! No 'Generate PDF' or 'Generate Doc' actions pending.");
            }

            var templateId = _App_extractIdFromUrl(config.templateUrl);
            var folderId = _App_extractIdFromUrl(config.folderUrl);

            if (!templateId || !folderId) {
                throw new Error("Could not extract IDs from URLs. Please provide full valid URLs.");
            }

            // Save config for future use
            _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_URL, config.templateUrl);
            _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_URL, config.folderUrl);

            var templateFile = null;
            var targetFolder = null;
            try {
                templateFile = DriveApp.getFileById(templateId);
                targetFolder = DriveApp.getFolderById(folderId);
            } catch (e) {
                throw new Error("Permission Error: I can't access the Doc or Folder. Make sure you have 'Editor' access.");
            }

            var masterDocId = null;
            if (mode === "SINGLE") {
                var dateStr = _App_formatDateTime(new Date(), "yyyy-MM-dd HH:mm");
                var masterDoc = templateFile.makeCopy('Merged_Doc_' + dateStr);
                masterDocId = masterDoc.getId();
                _App_setProperty(APP_PROPS.DOCS_MERGE_MASTER_DOC_ID, masterDocId);

                var masterDocOpen = DocumentApp.openById(masterDocId);
                masterDocOpen.getBody().clear();
                masterDocOpen.saveAndClose();
            }

            var headers = SheetManager.getHeaders('DOCS_MERGE');
            var rowLinkColName = "Merged File Link";
            var linkColIndex = headers.indexOf(rowLinkColName) + 1; // 1-based index

            var stats = _App_BatchProcessor('DOCS_MERGE', pendingRows, function (item, index) {
                var rowUpdates = {
                    action: item['Action'],
                    _rowNumber: item._rowNumber,
                    status: "",
                    linkUrl: null
                };

                var isFirstInWholeProcess = (index === 0);
                var isLastInWholeProcess = (index === (pendingRows.length - 1));
                var outputFormat = item['Action'] === "Generate PDF" ? "PDF" : "DOC";

                if (mode === "SINGLE") {
                    var tempId = templateFile.makeCopy('Temp_' + item._rowNumber).getId();
                    var tempDoc = DocumentApp.openById(tempId);
                    var tempBody = tempDoc.getBody();

                    _DocsMerge_replacePlaceholders(tempBody, headers, item);
                    tempDoc.saveAndClose();

                    var masterOpened = DocumentApp.openById(masterDocId);
                    var masterBody = masterOpened.getBody();

                    var tempOpened = DocumentApp.openById(tempId);
                    var tempBodyOpened = tempOpened.getBody();
                    var numChildren = tempBodyOpened.getNumChildren();

                    for (var j = 0; j < numChildren; j++) {
                        var child = tempBodyOpened.getChild(j).copy();
                        var type = child.getType();

                        if (isFirstInWholeProcess && j === 0) {
                            if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child.asParagraph());
                            else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child.asTable());
                            else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child.asListItem());
                            if (masterBody.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && masterBody.getChild(0).getText() === "") {
                                masterBody.removeChild(masterBody.getChild(0)); // Remove default empty paragraph safely
                            }
                        } else {
                            if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child.asParagraph());
                            else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child.asTable());
                            else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child.asListItem());
                        }
                    }

                    if (!isLastInWholeProcess) {
                        masterBody.appendPageBreak();
                    }

                    masterOpened.saveAndClose();
                    DriveApp.getFileById(tempId).setTrashed(true);

                    rowUpdates.status = _App_formatStatus('SUCCESS', "Appended to Master");
                    rowUpdates.action = "";
                } else {
                    // INDIVIDUAL
                    var fileName = item['Document Name'] || 'Document_' + item._rowNumber;
                    var tempFile = templateFile.makeCopy(fileName);
                    var tempDoc = DocumentApp.openById(tempFile.getId());
                    _DocsMerge_replacePlaceholders(tempDoc.getBody(), headers, item);
                    tempDoc.saveAndClose();

                    var finalUrl = "";
                    if (outputFormat === "PDF") {
                        var pdfBlob = tempFile.getAs(MimeType.PDF);
                        var newPdf = targetFolder.createFile(pdfBlob);
                        finalUrl = newPdf.getUrl();
                        tempFile.setTrashed(true);
                    } else {
                        tempFile.moveTo(targetFolder);
                        finalUrl = tempFile.getUrl();
                    }

                    rowUpdates.status = _App_formatStatus('SUCCESS', outputFormat + ' Created');
                    rowUpdates.linkUrl = finalUrl;
                    rowUpdates.action = "";
                }

                return rowUpdates;

            }, {
                onBatchComplete: function (results) {
                    var sheet = SheetManager.getSheet('DOCS_MERGE');

                    results.forEach(function (res) {
                        if (res && res._rowNumber !== undefined && !res.isError) {
                            if (res.linkUrl && linkColIndex > 0) {
                                var richText = SpreadsheetApp.newRichTextValue()
                                    .setText("View File")
                                    .setLinkUrl(res.linkUrl)
                                    .build();
                                sheet.getRange(res._rowNumber, linkColIndex).setRichTextValue(richText);
                            }
                        }
                    });

                    _App_batchPatchResults('DOCS_MERGE', results);
                }
            });

            // Handle finish step for SINGLE mode
            if (mode === "SINGLE" && masterDocId) {
                try {
                    var masterFile = DriveApp.getFileById(masterDocId);
                    var finalMasterUrl = "";

                    // Find if we should output PDF based on the first pending row's original action
                    var isPdf = pendingRows[0] && pendingRows[0]['Action'] === "Generate PDF";

                    if (isPdf) {
                        var masterPdfBlob = masterFile.getAs(MimeType.PDF);
                        var masterPdfFile = targetFolder.createFile(masterPdfBlob);
                        finalMasterUrl = masterPdfFile.getUrl();
                        masterFile.setTrashed(true);
                    } else {
                        masterFile.moveTo(targetFolder);
                        finalMasterUrl = masterFile.getUrl();
                    }

                    var sheet = SheetManager.getSheet('DOCS_MERGE');
                    var richTextMaster = SpreadsheetApp.newRichTextValue()
                        .setText("View Master File")
                        .setLinkUrl(finalMasterUrl)
                        .build();

                    if (linkColIndex > 0) {
                        pendingRows.forEach(function (item) {
                            sheet.getRange(item._rowNumber, linkColIndex).setRichTextValue(richTextMaster);
                        });
                    }
                } catch (err) {
                    throw new Error("Finish Export Failed: " + err.message);
                } finally {
                    _App_deleteProperty(APP_PROPS.DOCS_MERGE_MASTER_DOC_ID);
                }
            }

            var finalMsg = "Successfully processed " + stats.processedCount + " documents.";
            if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
            if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

            return _App_ok(finalMsg);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Docs Merge sidebar and ensures the sheet exists. */
function DocsMerge_openSidebar() {
  return Logger.run('DOCS_MERGE', 'Open Sidebar', function () {
    _App_launchTool('DOCS_MERGE');
  });
}

function _DocsMerge_replacePlaceholders(body, headers, rowObj) {
  // Headers start from Action(0), Doc Name(1), Merged Link(2). Dynamic placeholders start from index 3.
  for (var h = 3; h < headers.length; h++) {
    var key = headers[h];
    body.replaceText('{{' + key + '}}', rowObj[key] !== undefined ? String(rowObj[key]) : "");
  }
}


// --- FILE: tools/DriveFileDetails/Code.js ---
/**
 * Google Drive
 * Server-side Logic — Version 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('DRIVE_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Drive API', test: function() { return typeof Drive !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.DRIVE_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.DRIVE_SYNC,
    MENU_LABEL: SHEET_NAMES.DRIVE_SYNC,
    MENU_ENTRYPOINT: 'DriveFileDetails_openSidebar',
    MENU_ORDER: 90,
    SIDEBAR_HTML: 'tools/DriveFileDetails/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Item Name', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Starred', type: 'CHECKBOX' },
            { header: 'Type', type: 'DROPDOWN', options: ['Folder', 'Google Doc', 'Google Sheet', 'Google Slide', 'Google Form', 'PDF', 'Image', 'Video', 'Audio', 'Zip', 'Text', 'Code', 'File'] },
            { header: 'Editors', type: 'TEXT' },
            { header: 'Viewers', type: 'TEXT' },
            { header: 'Is Public?', type: 'CHECKBOX' },
            { header: 'Parent Path', type: 'TEXT' },
            { header: 'Item Path', type: 'TEXT', italic: true },
            { header: 'Size', type: 'TEXT' },
            { header: 'Owner', type: 'TEXT' },
            { header: 'URL', type: 'URL' },
            { header: 'Item ID', type: 'ID' },
            { header: 'Parent ID', type: 'ID' }
        ]
    },
    ACTIONS: {
        getFolderContent: function(folderId) {
            try {
                var parentId = folderId || "root";
                var query = "'" + parentId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
                var folders = [];
                var pageToken = null;

                var currentFolder = null;
                try {
                    currentFolder = Drive.Files.get(parentId, { fields: "id, name", supportsAllDrives: true });
                } catch (e) {
                    try {
                        var drv = Drive.Drives.get(parentId);
                        currentFolder = { id: drv.id, name: drv.name };
                    } catch (e2) {
                        currentFolder = { id: parentId, name: parentId === "root" ? "Root" : "Unknown" };
                    }
                }

                do {
                    var result = Drive.Files.list({
                        q: query,
                        fields: "nextPageToken, files(id, name)",
                        orderBy: "name",
                        pageToken: pageToken,
                        includeItemsFromAllDrives: true,
                        supportsAllDrives: true
                    });
                    if (result.files) folders = folders.concat(result.files);
                    pageToken = result.nextPageToken;
                } while (pageToken);

                return _App_ok('Folder content loaded', {
                    current: { id: currentFolder.id, name: currentFolder.name },
                    children: folders
                });
            } catch (e) {
                return _App_fail(e.message);
            }
        },
        getDrivesList: function() {
            try {
                var drives = [];
                var pageToken = null;
                do {
                    var result = _DriveFileDetails_safeListDrives({
                        pageToken: pageToken,
                        fields: "nextPageToken, drives(id, name)"
                    });
                    if (result && result.drives) drives = drives.concat(result.drives);
                    pageToken = result ? result.nextPageToken : null;
                } while (pageToken);

                drives.sort(function (a, b) {
                    return a.name.localeCompare(b.name);
                });

                return _App_ok('Drives loaded', { drives: drives });
            } catch (e) {
                return _App_fail(e.message);
            }
        },
        getFolderHierarchy: function() {
            try {
                var query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
                var folders = [];
                var pageToken = null;

                // 1. Fetch all folders
                do {
                    var result = _App_callWithBackoff(function () {
                        return Drive.Files.list({
                            q: query,
                            fields: "nextPageToken, files(id, name, parents)",
                            orderBy: "name",
                            pageToken: pageToken,
                            includeItemsFromAllDrives: true,
                            supportsAllDrives: true,
                            pageSize: 1000
                        });
                    });
                    if (result.files) folders = folders.concat(result.files);
                    pageToken = result.nextPageToken;
                } while (pageToken);

                // 2. Fetch all shared drives
                var drives = [];
                var drivePageToken = null;
                do {
                    var dResult = _DriveFileDetails_safeListDrives({
                        pageToken: drivePageToken,
                        fields: "nextPageToken, drives(id, name)"
                    });
                    if (dResult && dResult.drives) drives = drives.concat(dResult.drives);
                    drivePageToken = dResult ? dResult.nextPageToken : null;
                } while (drivePageToken);

                drives.sort(function (a, b) {
                    return a.name.localeCompare(b.name);
                });

                var topology = {
                    rootDrives: drives,
                    dict: {},
                    myDriveId: "root"
                };

                // 3. Resolve the actual ID of "My Drive"
                try {
                    var actualRoot = Drive.Files.get("root", { fields: "id", supportsAllDrives: true });
                    if (actualRoot && actualRoot.id) {
                        topology.myDriveId = actualRoot.id;
                    }
                } catch (e) { }

                // 4. Group folders by parentId
                for (var i = 0; i < folders.length; i++) {
                    var f = folders[i];
                    if (f.parents && f.parents.length > 0) {
                        var parentId = f.parents[0];
                        if (!topology.dict[parentId]) {
                            topology.dict[parentId] = [];
                        }
                        topology.dict[parentId].push({ id: f.id, name: f.name });
                    }
                }

                return _App_ok('Hierarchy loaded', { topology: topology });
            } catch (e) {
                return _App_fail(e.message);
            }
        },
        getPendingStats: function() {
            SheetManager.assertActiveSheet('DRIVE_SYNC');
            var stats = SheetManager.getActionStats('DRIVE_SYNC', ['CREATE', 'UPDATE', 'DELETE']);
            stats.total = (stats.CREATE || 0) + (stats.UPDATE || 0) + (stats.DELETE || 0);
            return _App_ok('Pending stats loaded.', {
                creates: stats.CREATE || 0,
                updates: stats.UPDATE || 0,
                deletes: stats.DELETE || 0,
                total: stats.total
            });
        },
        pull: function(targetFolderId, isShallow) {
            return _DriveFileDetails_pullFromDrive(targetFolderId, isShallow);
        },
        push: function() {
            return _DriveFileDetails_runPushSequence();
        },
        fillActivePath: function(folderId, pathString) {
            var sheet = _App_assertActiveSheet(SHEET_NAMES.DRIVE_SYNC);

            var cell = sheet.getActiveCell();
            var row = cell.getRow();

            if (row < 2) throw new Error("Please select a row in the data area (Row 2 or below).");

            sheet.getRange(row, DRIVE_SYNC_COL.PARENT_PATH + 1).setValue(pathString);
            sheet.getRange(row, DRIVE_SYNC_COL.PARENT_ID + 1).setValue(folderId);

            return "Updated Row " + row + " with path: " + pathString;
        }
    }
});

/* ==========================================================================
   CONFIGURATION
   ========================================================================== */

// Column-index aliases — kept for backward-compat; metadata now in SyncEngine.getTool('DRIVE_SYNC').
var DRIVE_SYNC_COL = {
  ACTION: 0, STATUS: 1, NAME: 2, DESC: 3, STARRED: 4, TYPE: 5,
  EDITORS: 6, VIEWERS: 7, IS_PUBLIC: 8, PARENT_PATH: 9, ITEM_PATH: 10,
  SIZE: 11, OWNER: 12, URL: 13, ITEM_ID: 14, PARENT_ID: 15
};

// --- SIDEBAR & SHEET SETUP ---
function _DriveFileDetails_InternalFunction() { } 

/** Opens the Drive Sync sidebar and ensures the sheet exists. */
function DriveFileDetails_openSidebar() {
  return Logger.run('DRIVE_SYNC', 'Open Sidebar', function () {
    _App_launchTool('DRIVE_SYNC');
  });
}

/* ==========================================================================
   CORE LOGIC
   ========================================================================== */

function _DriveFileDetails_pullFromDrive(targetFolderId, isShallow) {
  return Logger.run('DRIVE_SYNC', 'Pull from Drive', function () {
    return _App_withDocumentLock('DRIVE_SYNC_PULL', function () {
      _App_resetExecutionTimer();
      targetFolderId = targetFolderId || "root";

      var sheet = _App_ensureSheetExists('DRIVE_SYNC');

      SheetManager.clearData('DRIVE_SYNC');

      var allItems = [];
      var folderMap = new Map();

      // Recursive Fetch with Error Guard
      function recursiveFetch(parentId) {
        if (_App_isExecutionLimitApproaching()) {
          throw new Error("⏳ Time limit approaching. Operation paused safely.");
        }
        try {
          var query = "'" + parentId + "' in parents and trashed = false";
          var fields = "files(id, name, description, starred, mimeType, parents, webViewLink, size, permissions(type, role, emailAddress))";
          var items = _DriveFileDetails_fetchAllItems(query, fields);

          items.forEach(function (item) {
            item._traversalParentId = parentId; // Track which folder we found this in
            allItems.push(item);
            if (item.mimeType === 'application/vnd.google-apps.folder') {
              folderMap.set(item.id, { name: item.name, parentId: parentId });
              if (!isShallow) {
                recursiveFetch(item.id);
              }
            }
          });
        } catch (e) {
          if (e.message && e.message.indexOf("Time limit") !== -1) throw e;
          Logger.warn(SyncEngine.getTool('DRIVE_SYNC').TITLE, 'recursiveFetch', "Error fetching folder " + parentId + ": " + e.message);
        }
      }

      var rootObj = { id: targetFolderId, name: "Root", parents: [] };
      try {
        var drv = Drive.Drives.get(targetFolderId);
        if (drv && drv.name) rootObj.name = drv.name;
        else rootObj = Drive.Files.get(targetFolderId, { fields: "id, name, parents", supportsAllDrives: true });
      } catch (e) {
        rootObj = Drive.Files.get(targetFolderId, { fields: "id, name, parents", supportsAllDrives: true });
      }

      var rootParent = (rootObj.parents && rootObj.parents.length > 0) ? rootObj.parents[0] : null;
      folderMap.set(rootObj.id, { name: rootObj.name, parentId: rootParent });

      // Resolve full path of the target folder (start point)
      var targetFolderFullPath = "";
      var rootFolderName = "My Drive"; // Default fallback
      var foundSharedDriveRoot = false;

      try {
        if (targetFolderId !== "root") {
          var driveObj = Drive.Drives.get(targetFolderId);
          if (driveObj && driveObj.name) {
            rootFolderName = driveObj.name;
            foundSharedDriveRoot = true;
          }
        }
      } catch (e) { }

      if (!foundSharedDriveRoot) {
        try {
          var actualRoot = Drive.Files.get("root", { fields: "name", supportsAllDrives: true });
          if (actualRoot && actualRoot.name) rootFolderName = actualRoot.name;
        } catch (e) { Logger.warn(SyncEngine.getTool('DRIVE_SYNC').TITLE, 'Path Resolution', "Could not fetch root name, using default."); }
      }

      if (targetFolderId === "root" || foundSharedDriveRoot) {
        targetFolderFullPath = rootFolderName;
      } else {
        var parts = [];
        var curr = targetFolderId;
        var depth = 0;
        var foundRoot = false;
        while (curr && depth < 50) {
          if (curr === "root") {
            foundRoot = true;
            break;
          }
          try {
            var isDrive = false;
            try {
              var drv2 = Drive.Drives.get(curr);
              if (drv2 && drv2.name) {
                rootFolderName = drv2.name;
                foundRoot = true;
                isDrive = true;
              }
            } catch (e2) { }
            if (isDrive) break;

            var f = Drive.Files.get(curr, { fields: "name, parents", supportsAllDrives: true });
            parts.unshift(f.name);
            curr = (f.parents && f.parents.length) ? f.parents[0] : null;
            depth++;
          } catch (e) { break; }
        }
        var rootPrefix = foundRoot ? rootFolderName : "";
        var partsStr = parts.join("/");
        if (rootPrefix && partsStr) {
          targetFolderFullPath = rootPrefix + "/" + partsStr;
        } else if (rootPrefix) {
          targetFolderFullPath = rootPrefix;
        } else if (partsStr) {
          targetFolderFullPath = partsStr;
        } else {
          targetFolderFullPath = "";
        }
      }

      var isPartialPull = false;
      try {
        recursiveFetch(targetFolderId);
      } catch (timeoutEx) {
        if (timeoutEx.message && timeoutEx.message.indexOf("Time limit") !== -1) {
          isPartialPull = true;
        } else {
          throw timeoutEx;
        }
      }

      var rows = [];

      var getPath = function (itemId, currentPath) {
        if (!currentPath) currentPath = [];
        var item = folderMap.get(itemId);
        if (!item || itemId === targetFolderId) {
          var relative = currentPath.join("/");
          if (targetFolderFullPath && relative) {
            return targetFolderFullPath + "/" + relative;
          } else if (targetFolderFullPath) {
            return targetFolderFullPath;
          } else if (relative) {
            return relative;
          } else {
            return "";
          }
        }
        currentPath.unshift(item.name);
        return getPath(item.parentId, currentPath);
      };

      var headers = SyncEngine.getTool('DRIVE_SYNC').HEADERS;
      for (var i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        var parentId = item._traversalParentId || ((item.parents && item.parents.length > 0) ? item.parents[0] : "");
        var path = parentId ? getPath(parentId) : targetFolderFullPath;

        var perms = _DriveFileDetails_parsePermissions(item.permissions);

        var row = new Array(headers.length);
        row[DRIVE_SYNC_COL.ACTION] = "";
        row[DRIVE_SYNC_COL.STATUS] = "";
        row[DRIVE_SYNC_COL.NAME] = item.name;
        row[DRIVE_SYNC_COL.DESC] = item.description || "";
        row[DRIVE_SYNC_COL.STARRED] = item.starred || false;
        row[DRIVE_SYNC_COL.TYPE] = _DriveFileDetails_getFriendlyType(item.mimeType);

        row[DRIVE_SYNC_COL.SIZE] = _DriveFileDetails_formatBytes(item.size);
        row[DRIVE_SYNC_COL.OWNER] = perms.owners.join(", ");
        row[DRIVE_SYNC_COL.EDITORS] = perms.editors.join(", ");
        row[DRIVE_SYNC_COL.VIEWERS] = perms.viewers.join(", ");
        row[DRIVE_SYNC_COL.IS_PUBLIC] = perms.isPublic;

        row[DRIVE_SYNC_COL.PARENT_PATH] = path;
        row[DRIVE_SYNC_COL.ITEM_PATH] = path ? path + "/" + item.name : item.name;
        row[DRIVE_SYNC_COL.ITEM_ID] = item.id;
        row[DRIVE_SYNC_COL.PARENT_ID] = parentId;
        row[DRIVE_SYNC_COL.URL] = item.webViewLink;
        rows.push(row);
      }

      if (rows.length > 0) {
        var rowParams = { start: 2, total: rows.length };
        var range = sheet.getRange(rowParams.start, 1, rowParams.total, rows[0].length);
        range.setValues(rows);

        _App_applyBodyFormatting(sheet, rows.length, SyncEngine.getTool('DRIVE_SYNC').FORMAT_CONFIG);

        var msg = "Successfully pulled " + rows.length + " items from " + (targetFolderId === "root" ? "Root Drive" : rootObj.name) + ".";
        if (isPartialPull) {
          msg = "⚠️ Partial Pull: " + msg + " (Execution Time Limit Reached. Run again to continue)";
        }
        Logger.info(SyncEngine.getTool('DRIVE_SYNC').TITLE, 'Pull Complete', msg);
        return _App_ok(msg);
      } else {
        return _App_ok("Target folder is empty.");
      }

    });
  });
}

function _DriveFileDetails_runPushSequence() {
  return Logger.run('DRIVE_SYNC', 'Push Sequence', function () {
    return _App_withDocumentLock('DRIVE_SYNC_PUSH', function () {
      var logs = [];
      function log(msg) { logs.push("[" + _App_formatDateTime(new Date(), "HH:mm:ss") + "] " + msg); }
      log("Starting Push Sequence...");

      var pendingRows = SheetManager.readPendingObjects('DRIVE_SYNC');

      if (pendingRows.length === 0) {
        log("No pending actions found.");
        return logs;
      }

      log("Found " + pendingRows.length + " pending actions.");

      pendingRows.sort(function (a, b) {
        var score = function (obj) {
          var act = obj['Action'];
          var type = obj['Type'];
          if (act === 'CREATE' && type === 'Folder') return 1;
          if (act === 'CREATE') return 2;
          if (act === 'UPDATE') return 3;
          return 4;
        };
        return score(a) - score(b);
      });

      var stats = _App_BatchProcessor('DRIVE_SYNC', pendingRows, function (item) {
          var statusMsg = "";
          var action = item['Action'];
          var resultValues = {};

          if (action === 'CREATE') statusMsg = _DriveFileDetails_handleCreate(item, resultValues);
          else if (action === 'UPDATE') statusMsg = _DriveFileDetails_handleUpdate(item);
          else if (action === 'DELETE') statusMsg = _DriveFileDetails_handleDelete(item);

          log("Row " + item._rowNumber + ": " + statusMsg);

          return {
            _rowNumber: item._rowNumber,
            action: "",
            status: statusMsg,
            itemId: resultValues.id,
            url: resultValues.url,
            size: resultValues.size
          };

      }, {
        onBatchComplete: function (results) {
          _App_batchPatchResults('DRIVE_SYNC', results, function (res) {
            var fields = {};
            if (res.itemId) fields['Item ID'] = res.itemId;
            if (res.url) fields['URL'] = res.url;
            if (res.size) fields['Size'] = res.size;
            return fields;
          });
        }
      });

      if (stats.timeLimitReached) {
        log("Sequence Paused: 5.5-minute time limit approached. Please run again to process remaining items.");
      } else {
        log("Sequence Complete.");
      }
      return logs;

    });
  });
}

/* ==========================================================================
   PRIVATE HELPER FUNCTIONS
   ========================================================================== */

function _DriveFileDetails_validateHeaders(sheet) {
  var headers = SyncEngine.getTool('DRIVE_SYNC').HEADERS;
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (currentHeaders[0] !== headers[0] || currentHeaders[DRIVE_SYNC_COL.ITEM_ID] !== headers[DRIVE_SYNC_COL.ITEM_ID]) {
    _App_applyHeaderFormatting(sheet, headers);
  }
}

function _DriveFileDetails_fetchAllItems(query, fields) {
  var items = [];
  var pageToken = null;
  do {
    var result = _App_callWithBackoff(function () {
      return Drive.Files.list({
        q: query,
        fields: "nextPageToken, " + fields,
        pageToken: pageToken,
        pageSize: 1000,
        includeItemsFromAllDrives: true,
        supportsAllDrives: true
      });
    });
    if (result.files) items = items.concat(result.files);
    pageToken = result.nextPageToken;
  } while (pageToken);
  return items;
}

/**
 * Safely lists shared drives with error shielding for cases where they are not supported or enabled.
 * @param {Object} params - API parameters for Drive.Drives.list
 * @returns {Object|null} Result object or null on failure.
 */
function _DriveFileDetails_safeListDrives(params) {
  try {
    return _App_callWithBackoff(function () {
      return Drive.Drives.list(params);
    });
  } catch (e) {
    Logger.warn(SyncEngine.getTool('DRIVE_SYNC').TITLE, 'safeListDrives', "Shared Drives listing failed or unsupported: " + e.message);
    return null;
  }
}

function _DriveFileDetails_formatBytes(bytes) {
  if (!bytes || bytes == 0) return "-";
  var k = 1024;
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function _DriveFileDetails_parsePermissions(permissions) {
  var res = { owners: [], editors: [], viewers: [], isPublic: false };
  if (!permissions) return res;

  permissions.forEach(function (p) {
    if (p.type === 'anyone') res.isPublic = true;
    if (p.emailAddress) {
      if (p.role === 'owner') res.owners.push(p.emailAddress);
      else if (p.role === 'writer' || p.role === 'fileOrganizer') res.editors.push(p.emailAddress);
      else if (p.role === 'reader') res.viewers.push(p.emailAddress);
    }
  });
  return res;
}

function _DriveFileDetails_parseEmailList(str) {
  if (!str) return [];
  return str.toString().split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return s !== ""; });
}

function _DriveFileDetails_getFriendlyType(mimeType) {
  if (!mimeType) return 'File';
  if (mimeType === 'application/vnd.google-apps.folder') return 'Folder';
  if (mimeType === 'application/vnd.google-apps.spreadsheet') return 'Google Sheet';
  if (mimeType === 'application/vnd.google-apps.document') return 'Google Doc';
  if (mimeType === 'application/vnd.google-apps.presentation') return 'Google Slide';
  if (mimeType === 'application/vnd.google-apps.form') return 'Google Form';
  if (mimeType === 'application/pdf') return 'PDF';
  return 'File';
}

function _DriveFileDetails_getMimeTypeFromFriendly(friendlyType) {
  switch (friendlyType) {
    case 'Folder': return 'application/vnd.google-apps.folder';
    case 'Google Sheet': return 'application/vnd.google-apps.spreadsheet';
    case 'Google Doc': return 'application/vnd.google-apps.document';
    case 'Google Slide': return 'application/vnd.google-apps.presentation';
    case 'Google Form': return 'application/vnd.google-apps.form';
    case 'PDF': return 'application/pdf';
    default: return 'application/vnd.google-apps.folder';
  }
}

function _DriveFileDetails_handleCreate(rowObj, res) {
  var name = rowObj['Item Name'];
  if (!name) throw new Error("Name is required");

  var pathStr = rowObj['Parent Path'];
  if (!pathStr || pathStr.trim() === "") {
    throw new Error("Parent Path is required for creating an item.");
  }

  var desc = rowObj['Description'];
  var starred = rowObj['Starred'] === true || rowObj['Starred'] === 'TRUE';
  var friendlyType = rowObj['Type'] || 'Folder';

  var parentId = rowObj['Parent ID'];

  // Priority: Path > ParentID
  try {
    parentId = _DriveFileDetails_resolveFolderIdFromPath(pathStr);
  } catch (e) {
    throw new Error("Path resolution failed: " + e.message);
  }

  var mimeType = _DriveFileDetails_getMimeTypeFromFriendly(friendlyType);

  var resource = { name: name, description: desc, starred: starred, parents: [parentId], mimeType: mimeType };
  var file = _App_callWithBackoff(function () { return Drive.Files.create(resource, null, { fields: 'id, webViewLink, mimeType', supportsAllDrives: true }); });

  res.id = file.id;
  res.url = file.webViewLink;
  res.mime = file.mimeType;

  return _App_formatStatus('SUCCESS', "Created (" + (friendlyType || 'Folder') + ")");
}

function _DriveFileDetails_resolveFolderIdFromPath(pathString) {
  if (!pathString || pathString === "/" || pathString.trim() === "") return "root";

  // Normalize path: Remove leading/trailing slashes and split
  var parts = pathString.split("/").filter(function (p) { return p.trim() !== ""; });

  var possibleSharedDrive = null;
  if (parts.length > 0) {
    var first = parts[0].toLowerCase();
    if (first === "my drive" || first === "drive") {
      parts.shift();
    } else {
      try {
        var drivesResult = Drive.Drives.list({ fields: "drives(id, name)" });
        if (drivesResult && drivesResult.drives) {
          for (var d = 0; d < drivesResult.drives.length; d++) {
            if (drivesResult.drives[d].name.toLowerCase() === first) {
              possibleSharedDrive = drivesResult.drives[d];
              parts.shift();
              break;
            }
          }
        }
      } catch (e) { }
    }
  }

  var currentId = possibleSharedDrive ? possibleSharedDrive.id : "root";
  var resolvedSoFar = [];
  if (possibleSharedDrive) resolvedSoFar.push(possibleSharedDrive.name);

  for (var i = 0; i < parts.length; i++) {
    var folderName = parts[i];

    // Search for existing folder in current parent
    var query = "'" + currentId + "' in parents and name = '" + folderName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
    var folders = [];
    try {
      var result = Drive.Files.list({ q: query, fields: "files(id, name)", pageSize: 1, includeItemsFromAllDrives: true, supportsAllDrives: true });
      if (result.files && result.files.length > 0) {
        folders = result.files;
      }
    } catch (e) {
      throw new Error("Drive API error while searching for '" + folderName + "': " + e.message);
    }

    if (folders.length > 0) {
      currentId = folders[0].id;
      resolvedSoFar.push(folderName);
    } else {
      // Folder not found — inform the user instead of auto-creating
      var resolvedPath = resolvedSoFar.length > 0 ? resolvedSoFar.join("/") : "(root)";
      throw new Error(
        "Folder not found: '" + folderName + "' does not exist. " +
        "Resolved up to: " + resolvedPath + ". " +
        "Remaining path: " + parts.slice(i).join("/") + ". " +
        "Please create the folder first or correct the path."
      );
    }
  }

  return currentId;
}

function _DriveFileDetails_handleUpdate(rowObj) {
  var fileId = rowObj['Item ID'];
  if (!fileId) throw new Error("Cannot Update: Item ID is missing.");

  var newName = rowObj['Item Name'];
  var newDesc = rowObj['Description'];
  var newStarred = rowObj['Starred'] === true || rowObj['Starred'] === 'TRUE';

  var currentFile;
  try {
    currentFile = _App_callWithBackoff(function () {
      return Drive.Files.get(fileId, { fields: 'name, description, starred, permissions(id, role, emailAddress, type)', supportsAllDrives: true });
    });
  } catch (e) {
    // Fallback if permissions field access is forbidden
    currentFile = _App_callWithBackoff(function () {
      return Drive.Files.get(fileId, { fields: 'name, description, starred', supportsAllDrives: true });
    });
  }

  var changes = [];
  var resource = {};
  if (newName && newName !== currentFile.name) resource.name = newName;
  if (newDesc !== (currentFile.description || "")) resource.description = newDesc;
  if (newStarred !== (currentFile.starred || false)) resource.starred = newStarred;

  var optionalArgs = { supportsAllDrives: true };

  if (Object.keys(resource).length > 0) {
    _App_callWithBackoff(function () { Drive.Files.update(resource, fileId, null, optionalArgs); });
    changes.push("Properties");
  }

  // Permissions
  try {
    var newEditors = _DriveFileDetails_parseEmailList(rowObj['Editors']);
    var newViewers = _DriveFileDetails_parseEmailList(rowObj['Viewers']);
    var targetIsPublic = rowObj['Is Public?'] === true || rowObj['Is Public?'] === 'TRUE';

    var currentEmailPerms = {};
    var publicPermId = null;

    if (currentFile.permissions) {
      currentFile.permissions.forEach(function (p) {
        if (p.type === 'anyone') publicPermId = p.id;
        else if (p.emailAddress) currentEmailPerms[p.emailAddress.toLowerCase()] = p;
      });
    }

    if (targetIsPublic && !publicPermId) {
      _App_callWithBackoff(function () { Drive.Permissions.create({ role: 'reader', type: 'anyone' }, fileId, { supportsAllDrives: true }); });
      changes.push("Made Public");
    } else if (!targetIsPublic && publicPermId) {
      _App_callWithBackoff(function () { Drive.Permissions.remove(fileId, publicPermId, { supportsAllDrives: true }); });
      changes.push("Made Private");
    }

    var permChanges = false;
    Object.keys(currentEmailPerms).forEach(function (email) {
      var p = currentEmailPerms[email];
      if (p.role === 'owner' || p.role === 'organizer') return;

      var shouldBeEditor = newEditors.indexOf(email) !== -1;
      var shouldBeViewer = newViewers.indexOf(email) !== -1;

      if (!shouldBeEditor && !shouldBeViewer) {
        _App_callWithBackoff(function () { Drive.Permissions.remove(fileId, p.id, { supportsAllDrives: true }); });
        permChanges = true;
      } else if (shouldBeEditor && p.role !== 'writer' && p.role !== 'fileOrganizer') {
        _App_callWithBackoff(function () { Drive.Permissions.update({ role: 'writer' }, fileId, p.id, { supportsAllDrives: true }); });
        permChanges = true;
      } else if (shouldBeViewer && p.role !== 'reader') {
        _App_callWithBackoff(function () { Drive.Permissions.update({ role: 'reader' }, fileId, p.id, { supportsAllDrives: true }); });
        permChanges = true;
      }
    });

    var allTargetEmails = newEditors.concat(newViewers);
    allTargetEmails.forEach(function (email) {
      if (currentEmailPerms[email]) return;
      var role = newEditors.indexOf(email) !== -1 ? 'writer' : 'reader';
      _App_callWithBackoff(function () {
        Drive.Permissions.create({ role: role, type: 'user', emailAddress: email }, fileId, { sendNotificationEmails: false, supportsAllDrives: true });
      });
      permChanges = true;
    });

    if (permChanges) changes.push("Permissions");
  } catch (permErr) {
    Logger.warn("DRIVE_SYNC", "handleUpdate", "Could not fully sync permissions (may lack permissions management rights): " + permErr.message);
  }

  return changes.length > 0 ? _App_formatStatus('SUCCESS', "Updated: " + changes.join(", ")) : _App_formatStatus('INFO', "No Changes Needed");
}

function _DriveFileDetails_handleDelete(rowObj) {
  var fileId = rowObj['Item ID'];
  if (!fileId) throw new Error("Cannot Delete: Item ID is missing.");
  _App_callWithBackoff(function () { Drive.Files.update({ trashed: true }, fileId, null, { supportsAllDrives: true }); });
  return _App_formatStatus('SUCCESS', "Deleted (Trashed)");
}


// --- FILE: tools/FormsSync/Code.js ---
/**
 * Forms Sync Tool
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 * Syncs questions and options between Google Sheets and Google Forms
 */

SyncEngine.registerTool('FORMS_SYNC', {
    SHEET_NAME: SHEET_NAMES.FORMS_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.FORMS_SYNC,
    MENU_LABEL: SHEET_NAMES.FORMS_SYNC,
    MENU_ENTRYPOINT: 'FormsSync_openSidebar',
    MENU_ORDER: 70,
    SIDEBAR_HTML: 'tools/FormsSync/Sidebar',
    SIDEBAR_WIDTH: 300,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Question Title', type: 'TEXT' },
            { header: 'Type', type: 'DROPDOWN', options: ['MULTIPLE_CHOICE', 'CHECKBOX', 'LIST', 'TEXT', 'PARAGRAPH_TEXT', 'DATE', 'TIME', 'DATETIME', 'DURATION', 'SCALE', 'GRID', 'CHECKBOX_GRID', 'FILE_UPLOAD', 'PAGE_BREAK', 'SECTION_HEADER', 'IMAGE', 'VIDEO'] },
            { header: 'Options', type: 'TEXT' },
            { header: 'Help Text', type: 'TEXT' },
            { header: 'Required', type: 'CHECKBOX' },
            { header: 'Item ID', type: 'ID' }
        ]
    },
    ACTIONS: {
        getForms: function () {
            try {
                var files = DriveApp.searchFiles("mimeType='application/vnd.google-apps.form' and trashed=false");
                var forms = [];
                var count = 0;
                var MAX_FORMS = 10;

                while (files.hasNext() && count < MAX_FORMS) {
                    var file = files.next();
                    forms.push({
                        id: file.getId(),
                        title: file.getName() || "Untitled Form",
                        lastUpdated: file.getLastUpdated().getTime()
                    });
                    count++;
                }

                forms.sort(function (a, b) {
                    return b.lastUpdated - a.lastUpdated;
                });

                var mappedForms = forms.map(function (f) {
                    return { id: f.id, title: f.title };
                });

                var savedFormId = _App_getProperty(APP_PROPS.FORMS_SELECTED_FORM);

                return _App_ok('Forms loaded', { forms: mappedForms, savedFormId: savedFormId });
            } catch (e) {
                Logger.error(SyncEngine.getTool('FORMS_SYNC').TITLE, 'Get Forms', e);
                throw new Error("Failed to fetch forms: " + e.toString());
            }
        },
        pull: function (formInput) {
            return _FormsSync_pullForm(formInput);
        },
        push: function () {
            return _FormsSync_syncToForm();
        },
        getFormLinks: function () {
            var formId = _App_getProperty(APP_PROPS.FORMS_CURRENT_FORM);
            if (!formId) return _App_ok('No form selected.', null);
            try {
                var form = FormApp.openById(formId);
                return _App_ok('Links loaded', {
                    editUrl: form.getEditUrl(),
                    responsesUrl: form.getSummaryUrl()
                });
            } catch (e) {
                return _App_ok('Links loaded (fallback)', {
                    editUrl: 'https://docs.google.com/forms/d/' + formId + '/edit',
                    responsesUrl: 'https://docs.google.com/forms/d/' + formId + '/edit#responses'
                });
            }
        }
    }
});

/** Opens the Forms Sync sidebar and ensures the sheet exists. */
function FormsSync_openSidebar() {
    return Logger.run('FORMS_SYNC', 'Open Sidebar', function () {
        _App_launchTool('FORMS_SYNC');
    });
}

// --- CORE HELPER LOGIC ---

function _FormsSync_extractFormId(inputUrlOrId) {
    return _App_extractIdFromUrl(inputUrlOrId);
}

function _FormsSync_applyItemProperties(targetItem, type, required, optionsArr, gridRows, gridCols) {
    try {
        if (type === "MULTIPLE_CHOICE") {
            var mcItem = targetItem.asMultipleChoiceItem();
            mcItem.setRequired(required);
            if (optionsArr.length > 0) _FormsSync_setChoicesSafe(mcItem, optionsArr, "MULTIPLE_CHOICE");
        } else if (type === "CHECKBOX") {
            var cbItem = targetItem.asCheckboxItem();
            cbItem.setRequired(required);
            if (optionsArr.length > 0) _FormsSync_setChoicesSafe(cbItem, optionsArr, "CHECKBOX");
        } else if (type === "LIST") {
            var liItem = targetItem.asListItem();
            liItem.setRequired(required);
            if (optionsArr.length > 0) _FormsSync_setChoicesSafe(liItem, optionsArr, "LIST");
        } else if (type === "TEXT") {
            targetItem.asTextItem().setRequired(required);
        } else if (type === "PARAGRAPH_TEXT") {
            targetItem.asParagraphTextItem().setRequired(required);
        } else if (type === "DATE") {
            targetItem.asDateItem().setRequired(required);
        } else if (type === "TIME") {
            targetItem.asTimeItem().setRequired(required);
        } else if (type === "DATETIME") {
            targetItem.asDateTimeItem().setRequired(required);
        } else if (type === "DURATION") {
            targetItem.asDurationItem().setRequired(required);
        } else if (type === "GRID") {
            var gridItem = targetItem.asGridItem();
            gridItem.setRequired(required);
            if (gridRows && gridRows.length > 0) gridItem.setRows(gridRows);
            if (gridCols && gridCols.length > 0) gridItem.setColumns(gridCols);
        } else if (type === "CHECKBOX_GRID") {
            var cbGridItem = targetItem.asCheckboxGridItem();
            cbGridItem.setRequired(required);
            if (gridRows && gridRows.length > 0) cbGridItem.setRows(gridRows);
            if (gridCols && gridCols.length > 0) cbGridItem.setColumns(gridCols);
        }
    } catch (e) {
        Logger.warn(SyncEngine.getTool('FORMS_SYNC').TITLE, 'Apply Properties', "Failed to apply properties", e);
    }
}

function _FormsSync_setChoicesSafe(item, optionsArr, type) {
    if (!optionsArr || optionsArr.length === 0) return;

    try {
        var hasOther = false;
        if (typeof item.hasOtherOption === 'function') {
            hasOther = item.hasOtherOption();
        }

        var existingChoices = [];
        if (typeof item.getChoices === 'function') {
            existingChoices = item.getChoices();
        }

        var choices = [];
        for (var i = 0; i < optionsArr.length; i++) {
            var optString = optionsArr[i];
            var added = false;

            if (i < existingChoices.length) {
                var ec = existingChoices[i];
                if (typeof ec.getPageNavigationType === 'function') {
                    var navType = ec.getPageNavigationType();
                    if (navType === FormApp.PageNavigationType.GO_TO_PAGE) {
                        var gotoPage = ec.getGotoPage();
                        if (gotoPage && typeof item.createChoice === 'function') {
                            try {
                                if (type === "MULTIPLE_CHOICE" || type === "LIST") {
                                    choices.push(item.createChoice(optString, gotoPage));
                                    added = true;
                                }
                            } catch (e) { }
                        }
                    } else if (navType) {
                        try {
                            if (type === "MULTIPLE_CHOICE" || type === "LIST") {
                                choices.push(item.createChoice(optString, navType));
                                added = true;
                            }
                        } catch (e) { }
                    }
                }
            }
            if (!added && typeof item.createChoice === 'function') {
                choices.push(item.createChoice(optString));
            }
        }

        if (choices.length > 0 && typeof item.setChoices === 'function') {
            item.setChoices(choices);
        }

        if (hasOther && typeof item.showOtherOption === 'function') {
            item.showOtherOption(true);
        }
    } catch (e) {
        Logger.warn(SyncEngine.getTool('FORMS_SYNC').TITLE, 'Set Choices', "Failed to set choices", e);
    }
}

function _FormsSync_pullForm(formInput) {
    return Logger.run('FORMS_SYNC', 'Pull Form', function () {
        return _App_withDocumentLock('FORMS_SYNC_PULL', function () {
            var formId = _FormsSync_extractFormId(formInput);
            if (!formId) return _App_fail("Invalid Form URL or ID");

            try {
                var form = _App_callWithBackoff(function () { return FormApp.openById(formId); });
                var items = _App_callWithBackoff(function () { return form.getItems(); });

                var sheetData = [];

                for (var i = 0; i < items.length; i++) {
                    var item = items[i];
                    var id = item.getId().toString();
                    var title = item.getTitle() || "";
                    var type = item.getType().toString();
                    var options = "";
                    var helpText = item.getHelpText() || "";
                    var required = false;

                    // Extract type-specific properties (options, required)
                    try {
                        if (type === "MULTIPLE_CHOICE") {
                            var mcItem = item.asMultipleChoiceItem();
                            required = mcItem.isRequired();
                            options = mcItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                        } else if (type === "CHECKBOX") {
                            var cbItem = item.asCheckboxItem();
                            required = cbItem.isRequired();
                            options = cbItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                        } else if (type === "LIST") {
                            var liItem = item.asListItem();
                            required = liItem.isRequired();
                            options = liItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                        } else if (type === "TEXT") {
                            required = item.asTextItem().isRequired();
                        } else if (type === "PARAGRAPH_TEXT") {
                            required = item.asParagraphTextItem().isRequired();
                        } else if (type === "DATE") {
                            required = item.asDateItem().isRequired();
                        } else if (type === "TIME") {
                            required = item.asTimeItem().isRequired();
                        } else if (type === "DATETIME") {
                            required = item.asDateTimeItem().isRequired();
                        } else if (type === "DURATION") {
                            required = item.asDurationItem().isRequired();
                        } else if (type === "SCALE") {
                            required = item.asScaleItem().isRequired();
                        } else if (type === "GRID") {
                            var gridItem = item.asGridItem();
                            required = gridItem.isRequired();
                            var gridRows = gridItem.getRows() || [];
                            var gridCols = gridItem.getColumns() || [];
                            options = gridRows.join("\n") + "\n||\n" + gridCols.join("\n");
                        } else if (type === "CHECKBOX_GRID") {
                            var cbGridItem = item.asCheckboxGridItem();
                            required = cbGridItem.isRequired();
                            var cbGridRows = cbGridItem.getRows() || [];
                            var cbGridCols = cbGridItem.getColumns() || [];
                            options = cbGridRows.join("\n") + "\n||\n" + cbGridCols.join("\n");
                        }
                    } catch (propErr) {
                        Logger.warn(SyncEngine.getTool('FORMS_SYNC').TITLE, 'Property Error', "Error reading item properties for ID " + id + ": " + propErr);
                    }

                    sheetData.push(["", "", title, type, options, helpText, required, id]);
                }

                var sheet = _App_ensureSheetExists('FORMS_SYNC');

                // Clear old data
                var lastRow = sheet.getLastRow();
                if (lastRow > 1) {
                    var headers = SheetManager.getHeaders('FORMS_SYNC');
                    var reqColIndex = headers.indexOf('Required') + 1;
                    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
                    if (reqColIndex > 0) sheet.getRange(2, reqColIndex, lastRow - 1, 1).removeCheckboxes();
                }

                // Set New Data
                if (sheetData.length > 0) {
                    var targetRange = sheet.getRange(2, 1, sheetData.length, sheetData[0].length);
                    targetRange.setValues(sheetData);
                }

                // Apply body formatting via shared utility
                _App_applyBodyFormatting(sheet, sheetData.length, SyncEngine.getTool('FORMS_SYNC').FORMAT_CONFIG);

                // Save Form ID to PropertiesService for syncing back
                _App_setProperty(APP_PROPS.FORMS_CURRENT_FORM, formId);
                // Save to UserProperties for sidebar auto-selection
                _App_setProperty(APP_PROPS.FORMS_SELECTED_FORM, formId);

                return _App_ok("Successfully pulled " + sheetData.length + " items.");
            } catch (e) {
                throw e;
            }
        });
    });
}

function _FormsSync_syncToForm() {
    return Logger.run('FORMS_SYNC', 'Sync to Form', function () {
        return _App_withDocumentLock('FORMS_SYNC_PUSH', function () {
            var formId = _App_getProperty(APP_PROPS.FORMS_CURRENT_FORM);
            if (!formId) return _App_fail("No form connected. Please Pull data first.");

            try {
                var form = _App_callWithBackoff(function () { return FormApp.openById(formId); });
                
                var pendingRows = SheetManager.readPendingObjects('FORMS_SYNC', { useDisplayValues: true });

                if (pendingRows.length === 0) return _App_ok("No pending actions found.");

                var stats = _App_BatchProcessor('FORMS_SYNC', pendingRows, function (item) {
                    var action = (item['Action'] || "").toString().trim().toUpperCase();
                    var id = (item['Item ID'] || "").toString().trim();
                    var title = (item['Question Title'] || "").toString();
                    var type = (item['Type'] || "").toString();
                    var optionsRaw = (item['Options'] || "").toString();
                    var helpText = (item['Help Text'] || "").toString();
                    var required = item['Required'] === true || item['Required'] === 'TRUE';

                    var updateObj = {
                        action: action,
                        id: id,
                        status: "",
                        _rowNumber: item._rowNumber
                    };

                    var optionsArr = [];
                    var gridRows = [];
                    var gridCols = [];

                    if (type === "GRID" || type === "CHECKBOX_GRID") {
                        var gridParts = optionsRaw.split("||");
                        gridRows = (gridParts[0] || "").split("\n").map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 0; });
                        gridCols = (gridParts[1] || "").split("\n").map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 0; });
                    } else {
                        optionsArr = optionsRaw ? optionsRaw.split("\n").map(function (o) { return o.trim(); }).filter(function (o) { return o.length > 0; }) : [];
                    }

                    if (action === "CREATE") {
                        if (!title) throw new Error("Missing Title");
                        var targetItem = null;

                        _App_callWithBackoff(function () {
                            if (type === "MULTIPLE_CHOICE") targetItem = form.addMultipleChoiceItem();
                            else if (type === "CHECKBOX") targetItem = form.addCheckboxItem();
                            else if (type === "LIST") targetItem = form.addListItem();
                            else if (type === "TEXT") targetItem = form.addTextItem();
                            else if (type === "PARAGRAPH_TEXT") targetItem = form.addParagraphTextItem();
                            else if (type === "DATE") targetItem = form.addDateItem();
                            else if (type === "TIME") targetItem = form.addTimeItem();
                            else if (type === "DATETIME") targetItem = form.addDateTimeItem();
                            else if (type === "DURATION") targetItem = form.addDurationItem();
                            else if (type === "SCALE") targetItem = form.addScaleItem();
                            else if (type === "GRID") targetItem = form.addGridItem();
                            else if (type === "CHECKBOX_GRID") targetItem = form.addCheckboxGridItem();
                            else { targetItem = form.addTextItem(); type = "TEXT"; }
                        });

                        _App_callWithBackoff(function () {
                            targetItem.setTitle(title);
                            targetItem.setHelpText(helpText);
                            _FormsSync_applyItemProperties(targetItem, type, required, optionsArr, gridRows, gridCols);
                        });

                        updateObj.id = targetItem.getId().toString();
                        updateObj.action = "";
                        updateObj.status = _App_formatStatus('SUCCESS', "Created");
                    }
                    else if (action === "UPDATE") {
                        if (!id) throw new Error("Missing ID");
                        var updItem = _App_callWithBackoff(function () { return form.getItemById(parseInt(id, 10)); });
                        if (!updItem) throw new Error("Item ID not found");

                        var currentType = updItem.getType().toString();

                        if (currentType === type) {
                            _App_callWithBackoff(function () {
                                updItem.setTitle(title);
                                updItem.setHelpText(helpText);
                                _FormsSync_applyItemProperties(updItem, type, required, optionsArr, gridRows, gridCols);
                            });
                            updateObj.status = _App_formatStatus('SUCCESS', "Updated");
                        } else {
                            var targetIndex = updItem.getIndex();
                            _App_callWithBackoff(function () { form.deleteItem(updItem); });

                            var newItem = null;
                            _App_callWithBackoff(function () {
                                if (type === "MULTIPLE_CHOICE") newItem = form.addMultipleChoiceItem();
                                else if (type === "CHECKBOX") newItem = form.addCheckboxItem();
                                else if (type === "LIST") newItem = form.addListItem();
                                else if (type === "TEXT") newItem = form.addTextItem();
                                else if (type === "PARAGRAPH_TEXT") newItem = form.addParagraphTextItem();
                                else if (type === "DATE") newItem = form.addDateItem();
                                else if (type === "TIME") newItem = form.addTimeItem();
                                else if (type === "DATETIME") newItem = form.addDateTimeItem();
                                else if (type === "DURATION") newItem = form.addDurationItem();
                                else if (type === "SCALE") newItem = form.addScaleItem();
                                else if (type === "GRID") newItem = form.addGridItem();
                                else if (type === "CHECKBOX_GRID") newItem = form.addCheckboxGridItem();
                                else { newItem = form.addTextItem(); type = "TEXT"; }
                            });

                            _App_callWithBackoff(function () {
                                newItem.setTitle(title);
                                newItem.setHelpText(helpText);
                                _FormsSync_applyItemProperties(newItem, type, required, optionsArr, gridRows, gridCols);
                                form.moveItem(newItem.getIndex(), targetIndex);
                            });

                            updateObj.id = newItem.getId().toString();
                            updateObj.status = _App_formatStatus('SUCCESS', "Updated (Type Recreated)");
                        }
                        updateObj.action = "";
                    }
                    else if (action === "DELETE") {
                        if (!id) throw new Error("Missing ID");
                        var delItem = _App_callWithBackoff(function () { return form.getItemById(parseInt(id, 10)); });
                        if (delItem) {
                            _App_callWithBackoff(function () { form.deleteItem(delItem); });
                            updateObj.status = _App_formatStatus('SUCCESS', "Deleted");
                        } else {
                            updateObj.status = _App_formatStatus('WARNING', "Already Deleted");
                        }
                        updateObj.action = "";
                    }

                    return updateObj;
                }, {
                    onBatchComplete: function (batchResults) {
                        _App_batchPatchResults('FORMS_SYNC', batchResults, function (res) {
                            return {
                                'Item ID': res.id
                            };
                        });
                    }
                });

                return _App_ok("Sync Complete. Processed: " + stats.processedCount);
            } catch (e) {
                throw e;
            }
        });
    });
}


// --- FILE: tools/GmailFilters/Code.js ---
/**
 * Gmail Filters Tool
 * Version: 1.0 (Plugin Architecture)
 * 
 * Allows users to manage Gmail filters directly from the spreadsheet.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('GMAIL_FILTERS', {
    REQUIRED_SERVICES: [{ name: 'Gmail API', test: function () { return typeof Gmail !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.GMAIL_FILTERS,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.GMAIL_FILTERS,
    MENU_LABEL: SHEET_NAMES.GMAIL_FILTERS,
    MENU_ENTRYPOINT: 'GmailFilters_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'tools/GmailFilters/Sidebar',
    SIDEBAR_WIDTH: 320,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Criteria: From', type: 'TEXT' },
            { header: 'Criteria: To', type: 'TEXT' },
            { header: 'Criteria: Subject', type: 'TEXT' },
            { header: 'Criteria: Includes Words', type: 'TEXT' },
            { header: 'Criteria: Excludes Words', type: 'TEXT' },
            { header: 'Criteria: Has Attachment', type: 'CHECKBOX' },
            { header: 'Action: Skip the Inbox (Archive it)', type: 'CHECKBOX' },
            { header: 'Action: Mark as read', type: 'CHECKBOX' },
            { header: 'Action: Star it', type: 'CHECKBOX' },
            {
                header: 'Action: Labels', type: 'DROPDOWN', allowInvalid: true, options: function () {
                    var labels = [];
                    try {
                        var response = _App_callWithBackoff(function () { return Gmail.Users.Labels.list('me'); });
                        var systemLabelIds = ['INBOX', 'UNREAD', 'STARRED', 'TRASH', 'SPAM', 'IMPORTANT', 'CHAT', 'DRAFT', 'GREEN_CIRCLE', 'SENT', 'YELLOW_STAR'];
                        (response.labels || []).forEach(function (l) {
                            if (systemLabelIds.indexOf(l.id) === -1 && !l.id.startsWith('CATEGORY_')) {
                                labels.push(l.name);
                            }
                        });
                        labels.sort();
                    } catch (e) { }
                    return labels.length ? labels.slice(0, 499) : ['None'];
                }
            },
            { header: 'Action: Forward to', type: 'EMAIL' },
            { header: 'Action: Delete it', type: 'CHECKBOX' },
            { header: 'Action: Never send it to Spam', type: 'CHECKBOX' },
            { header: 'Action: Always mark it as important', type: 'CHECKBOX' },
            { header: 'Action: Never mark it as important', type: 'CHECKBOX' },
            { header: 'Action: Also apply filter to previous mails', type: 'CHECKBOX' },
            { header: 'Filter ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to manage Gmail filters:</p><ol><li><strong>Define Criteria:</strong> Set filter matching rules under <code>Criteria:</code> columns (From, Subject, etc.).</li><li><strong>Set Action:</strong> Set output settings (Skip Inbox, Labels, Delete, etc.).</li><li><strong>Push:</strong> Click <strong>Push Changes</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Understand filter criteria and action flags.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Criteria:</strong> The filter rules that match incoming messages (e.g. <code>Criteria: From</code>).</li><li><strong>Action: Labels:</strong> Custom label to apply to matching messages.</li><li><strong>Retroactive:</strong> Tick <code>Also apply filter to previous mails</code> to run the filter matching on past emails.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        pull: function () {
            var labels = _GmailFilters_getLabelMap();
            var filtersResponse = _App_callWithBackoff(function () {
                return Gmail.Users.Settings.Filters.list('me');
            });

            var filters = (filtersResponse && filtersResponse.filter) || [];
            var rows = filters.map(function (f) {
                var criteria = f.criteria || {};
                var action = f.action || {};
                var addLabelIds = action.addLabelIds || [];
                var removeLabelIds = action.removeLabelIds || [];

                // Filter out system labels for the "Labels" column and get the first user label
                var systemLabelIds = ['INBOX', 'UNREAD', 'STARRED', 'TRASH', 'SPAM', 'IMPORTANT'];
                var userLabelIds = addLabelIds.filter(function (id) {
                    return systemLabelIds.indexOf(id) === -1 && !id.startsWith('CATEGORY_');
                });
                var labelsList = userLabelIds.length > 0 ? (labels.idToName[userLabelIds[0]] || userLabelIds[0]) : '';

                return {
                    'Action': '',
                    'Status': 'Synced',
                    'Criteria: From': criteria.from || '',
                    'Criteria: To': criteria.to || '',
                    'Criteria: Subject': criteria.subject || '',
                    'Criteria: Includes Words': criteria.query || '',
                    'Criteria: Excludes Words': criteria.negatedQuery || '',
                    'Criteria: Has Attachment': !!criteria.hasAttachment,
                    'Action: Skip the Inbox (Archive it)': removeLabelIds.indexOf('INBOX') !== -1,
                    'Action: Mark as read': removeLabelIds.indexOf('UNREAD') !== -1,
                    'Action: Star it': addLabelIds.indexOf('STARRED') !== -1,
                    'Action: Labels': labelsList,
                    'Action: Forward to': action.forward || '',
                    'Action: Delete it': addLabelIds.indexOf('TRASH') !== -1,
                    'Action: Never send it to Spam': removeLabelIds.indexOf('SPAM') !== -1,
                    'Action: Always mark it as important': addLabelIds.indexOf('IMPORTANT') !== -1,
                    'Action: Never mark it as important': removeLabelIds.indexOf('IMPORTANT') !== -1,
                    'Action: Also apply filter to previous mails': false,
                    'Filter ID': f.id
                };
            });

            // Clear existing and write new
            SheetManager.overwriteObjects('GMAIL_FILTERS', rows);

            return _App_ok("Successfully pulled " + rows.length + " filters.");
        },
        push: function () {
            var pendingItems = SheetManager.readPendingObjects('GMAIL_FILTERS');
            if (pendingItems.length === 0) {
                return _App_ok("No pending actions to process.");
            }

            var labelMap = _GmailFilters_getLabelMap();

            var stats = _App_BatchProcessor('GMAIL_FILTERS', pendingItems, function (item) {
                var actionType = item['Action'];
                var filterId = item['Filter ID'];
                var resultStatus = _App_formatStatus('SUCCESS', "Success");
                var resultAction = "";
                var newFilterId = filterId;

                if (actionType === 'DELETE' || actionType === 'UPDATE') {
                    if (!filterId) throw new Error("Missing Filter ID for " + actionType);
                    try {
                        _App_callWithBackoff(function () {
                            Gmail.Users.Settings.Filters.remove('me', filterId);
                        });
                    } catch (e) {
                        if (e.message.indexOf("Empty response") !== -1) {
                            // Ignored "Empty response"
                        } else {
                            throw e;
                        }
                    }
                    if (actionType === 'DELETE') {
                        return { action: "", status: _App_formatStatus('SUCCESS', "Deleted"), _rowNumber: item._rowNumber };
                    }
                }

                if (actionType === 'CREATE' || actionType === 'UPDATE') {
                    var filterResource = _GmailFilters_constructFilterResource(item, labelMap.nameToId);
                    var createdFilter = _App_callWithBackoff(function () {
                        return Gmail.Users.Settings.Filters.create(filterResource, 'me');
                    });
                    newFilterId = createdFilter.id;
                    resultStatus = (actionType === 'UPDATE') ? _App_formatStatus('SUCCESS', "Updated") : _App_formatStatus('SUCCESS', "Created");

                    // Handle retroactive application
                    if (item['Action: Also apply filter to previous mails']) {
                        var searchQuery = _GmailFilters_buildSearchQuery(filterResource.criteria);
                        _GmailFilters_applyToExistingMessages(searchQuery, filterResource.action.addLabelIds || [], filterResource.action.removeLabelIds || []);
                        resultStatus += " (+ Applied to existing)";
                    }
                }

                return {
                    action: resultAction,
                    status: resultStatus,
                    'Filter ID': newFilterId,
                    _rowNumber: item._rowNumber
                };
            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('GMAIL_FILTERS', batchResults, function (res) {
                        var fields = {};
                        if (res['Filter ID'] !== undefined) {
                            fields['Filter ID'] = res['Filter ID'];
                        }
                        return fields;
                    });
                }
            });

            return _App_ok("Processed " + stats.processedCount + " filters.");
        },
        getMissingLabels: function () {
            var pendingItems = SheetManager.readPendingObjects('GMAIL_FILTERS');
            if (pendingItems.length === 0) return _App_ok('No pending actions.', []);

            var labelsInSheet = [];
            pendingItems.forEach(function (item) {
                var action = (item['Action'] || '').toString().toUpperCase();
                if (action === 'CREATE' || action === 'UPDATE') {
                    var label = item['Action: Labels'] ? item['Action: Labels'].toString().trim() : '';
                    if (label && labelsInSheet.indexOf(label) === -1) {
                        labelsInSheet.push(label);
                    }
                }
            });

            if (labelsInSheet.length === 0) return _App_ok('No labels to check.', []);

            var labelMap = _GmailFilters_getLabelMap();
            var missing = labelsInSheet.filter(function (name) {
                return !labelMap.nameToId[name.toLowerCase()];
            });

            return _App_ok('Missing labels identified.', missing);
        }
    }
});

// --- PUBLIC ENTRY POINTS ---

/**
 * Opens the Sidebar and prepares the sheet.
 */
function GmailFilters_openSidebar() {
    return Logger.run('GMAIL_FILTERS', 'Open Sidebar', function () {
        _App_launchTool('GMAIL_FILTERS');
    });
}

// --- INTERNAL HELPERS ---

/**
 * Fetches Gmail labels and creates mapping objects.
 */
function _GmailFilters_getLabelMap() {
    var response = _App_callWithBackoff(function () {
        return Gmail.Users.Labels.list('me');
    });

    var nameToId = {};
    var idToName = {};

    (response.labels || []).forEach(function (l) {
        nameToId[l.name.toLowerCase()] = l.id;
        idToName[l.id] = l.name;
    });

    return { nameToId: nameToId, idToName: idToName };
}

/**
 * Resolves Label ID to Name.
 */
function _GmailFilters_resolveLabelIds(ids, idToName) {
    if (!ids || !Array.isArray(ids) || ids.length === 0) return '';
    return idToName[ids[0]] || ids[0];
}

/**
 * Constructs a Gmail Filter resource from sheet data.
 */
function _GmailFilters_constructFilterResource(item, nameToId) {
    var criteria = {};
    if (item['Criteria: From']) criteria.from = item['Criteria: From'];
    if (item['Criteria: To']) criteria.to = item['Criteria: To'];
    if (item['Criteria: Subject']) criteria.subject = item['Criteria: Subject'];
    if (item['Criteria: Includes Words']) criteria.query = item['Criteria: Includes Words'];
    if (item['Criteria: Excludes Words']) criteria.negatedQuery = item['Criteria: Excludes Words'];
    if (item['Criteria: Has Attachment']) criteria.hasAttachment = true;

    // Item 5: Validate minimum criteria
    var hasCriteria = !!(criteria.from || criteria.to || criteria.subject || criteria.query || criteria.negatedQuery || criteria.hasAttachment);
    if (!hasCriteria) {
        throw new Error("A filter must specify at least one criteria (e.g., Criteria: From, Criteria: Subject, etc.).");
    }

    var action = {};
    var addLabelIds = [];
    var removeLabelIds = [];

    // Process Boolean Flags -> addLabelIds
    if (item['Action: Star it']) addLabelIds.push('STARRED');
    if (item['Action: Delete it']) addLabelIds.push('TRASH');
    if (item['Action: Always mark it as important']) addLabelIds.push('IMPORTANT');

    // Process Boolean Flags -> removeLabelIds
    if (item['Action: Skip the Inbox (Archive it)']) removeLabelIds.push('INBOX');
    if (item['Action: Mark as read']) removeLabelIds.push('UNREAD');
    if (item['Action: Never send it to Spam']) removeLabelIds.push('SPAM');
    if (item['Action: Never mark it as important']) removeLabelIds.push('IMPORTANT');

    // Item 5: Conflict checking
    if (item['Action: Always mark it as important'] && item['Action: Never mark it as important']) {
        throw new Error("Cannot set a filter to both 'Always mark it as important' and 'Never mark it as important'.");
    }

    // Process Label with Auto-Creation (Case-Insensitive)
    if (item['Action: Labels']) {
        var labelName = item['Action: Labels'].toString().trim();
        if (labelName) {
            var lookupName = labelName.toLowerCase();
            var id = nameToId[lookupName];
            if (!id) {
                // Auto-create missing label
                var newLabel = _App_callWithBackoff(function () {
                    return Gmail.Users.Labels.create({ name: labelName }, 'me');
                });
                id = newLabel.id;
                nameToId[lookupName] = id; // Update map for subsequent rows in same batch
            }
            if (addLabelIds.indexOf(id) === -1) addLabelIds.push(id);
        }
    }

    if (item['Action: Forward to']) action.forward = item['Action: Forward to'];
    
    // Item 5: Validate minimum action
    var hasAction = !!(addLabelIds.length > 0 || removeLabelIds.length > 0 || action.forward);
    if (!hasAction) {
        throw new Error("A filter must specify at least one action (e.g., Star it, Skip Inbox, Labels, etc.).");
    }

    if (addLabelIds.length > 0) action.addLabelIds = addLabelIds;
    if (removeLabelIds.length > 0) action.removeLabelIds = removeLabelIds;

    return { criteria: criteria, action: action };
}

/**
 * Builds a Gmail search query string from filter criteria.
 */
function _GmailFilters_buildSearchQuery(criteria) {
    var queryParts = [];
    if (criteria.from) queryParts.push('from:(' + criteria.from + ')');
    if (criteria.to) queryParts.push('to:(' + criteria.to + ')');
    if (criteria.subject) queryParts.push('subject:(' + criteria.subject + ')');
    if (criteria.query) queryParts.push(criteria.query);
    if (criteria.negatedQuery) queryParts.push('-(' + criteria.negatedQuery + ')');
    if (criteria.hasAttachment) queryParts.push('has:attachment');
    return queryParts.join(' ').trim();
}

/**
 * Applies labels to up to 1000 existing messages matching the query.
 */
function _GmailFilters_applyToExistingMessages(query, addLabelIds, removeLabelIds) {
    if (!query) return;

    var response = _App_callWithBackoff(function () {
        return Gmail.Users.Messages.list('me', { q: query, maxResults: 1000 });
    });

    if (response.messages && response.messages.length > 0) {
        var messageIds = response.messages.map(function (m) { return m.id; });
        _App_callWithBackoff(function () {
            Gmail.Users.Messages.batchModify({
                ids: messageIds,
                addLabelIds: addLabelIds.length > 0 ? addLabelIds : undefined,
                removeLabelIds: removeLabelIds.length > 0 ? removeLabelIds : undefined
            }, 'me');
        });
    }
}


// --- FILE: tools/MailMerge/Code.js ---
/**
 * Mail Merge
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_MERGE', {
    SHEET_NAME: SHEET_NAMES.MAIL_MERGE,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.MAIL_MERGE,
    MENU_LABEL: SHEET_NAMES.MAIL_MERGE,
    MENU_ENTRYPOINT: 'MailMerge_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'tools/MailMerge/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [
            { type: 'pending', actionCol: 'A', scope: 'actionOnly' },
            { type: 'success', statusCol: 'B', scope: 'fullRow' },
            { type: 'error', statusCol: 'B', scope: 'fullRow' }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'To', type: 'EMAIL_LIST' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these 3 steps to run your first merge:</p><ol><li><strong>Gmail Draft:</strong> Create a draft with tags like <code>{{First Name}}</code>.</li><li><strong>Select & Sync:</strong> Choose the draft above and click <strong>Pull Placeholders</strong>.</li><li><strong>Execute:</strong> Fill the data, set Action to <strong>SEND</strong>, and click <strong>Run</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about Action, To, Thread ID, and Attachments.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>SEND</code> for immediate mailing or <code>DRAFT</code> to create Gmail drafts for review.</li><li><strong>To:</strong> Primary recipient email address.</li><li><strong>Thread ID or Subject:</strong> Paste a Gmail Thread ID to reply to a specific conversation.</li><li><strong>Attachments:</strong> Comma-separated list of Google Drive File IDs or URLs.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "Case-sensitivity, formatting, and quota limits.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>Case Sensitive:</strong> Tags like <code>{{Name}}</code> must match the column header exactly.</li><li><strong>HTML Support:</strong> Any formatting (bold, links, images) in your Gmail draft is preserved.</li><li><strong>Status:</strong> The tool updates the Status column automatically after each row.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getQuota: function () {
            return _App_ok('Quota loaded.', { remaining: MailApp.getRemainingDailyQuota() });
        },
        getGmailDrafts: function () {
            try {
                var drafts = GmailApp.getDrafts();
                var validDrafts = [];
                var regex = /\{\{[^{}]+\}\}/;

                for (var i = 0; i < drafts.length; i++) {
                    var msg = drafts[i].getMessage();
                    var subject = msg.getSubject() || "";
                    var body = msg.getBody() || "";

                    if (regex.test(subject) || regex.test(body)) {
                        validDrafts.push({
                            id: drafts[i].getId(),
                            subject: subject || "(No Subject)"
                        });

                        if (validDrafts.length >= 10) {
                            break;
                        }
                    }
                }
                return _App_ok('Drafts loaded.', { drafts: validDrafts });
            } catch (e) {
                return _App_ok('No drafts available.', { drafts: [] });
            }
        },
        syncPlaceholders: function (draftId) {
            if (!draftId) return _App_fail("No draft selected.");
            try {
                var draft = GmailApp.getDraft(draftId);
                if (!draft) throw new Error("Draft not found.");
                var msg = draft.getMessage();
                var subject = msg.getSubject() || "";
                var body = msg.getBody() || "";

                var placeholders = [];
                var regex = /\{\{([^{}]+)\}\}/g;

                var match;
                while ((match = regex.exec(subject)) !== null) {
                    if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
                }
                while ((match = regex.exec(body)) !== null) {
                    if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
                }

                var syncResult = SheetManager.syncDynamicColumns('MAIL_MERGE', placeholders, {
                    dynamicColWidth: 150
                });

                return _App_ok('Synced ' + placeholders.length + ' placeholders.', {
                    placeholders: placeholders,
                    headers: syncResult.headers
                });
            } catch (e) {
                var toolConfig = SyncEngine.getTool('MAIL_MERGE') || { TITLE: 'MAIL_MERGE' };
                Logger.error(toolConfig.TITLE, 'Sync Placeholders', e);
                return _App_fail("Sync failed: " + e.message + (e.stack ? "\nTrace:\n" + e.stack : ""));
            }
        },
        executeActions: function (draftId) {
            var pendingRows = SheetManager.readPendingObjects('MAIL_MERGE', { useDisplayValues: true });

            if (pendingRows.length === 0) {
                return _App_ok("Nothing to do! No 'SEND' or 'DRAFT' actions pending.");
            }

            var template = null;
            try {
                var draft = GmailApp.getDraft(draftId);
                if (!draft) throw new Error("Draft not found.");
                var msg = draft.getMessage();
                template = {
                    subject: msg.getSubject(),
                    body: msg.getBody(),
                    attachments: msg.getAttachments()
                };
            } catch (e) {
                throw new Error("⚠️ Failed to load Draft: " + e.message);
            }

            var headers = SheetManager.getHeaders('MAIL_MERGE');
            var compiledPlaceholders = [];
            for (var colIndex = 2; colIndex < headers.length; colIndex++) {
                var header = headers[colIndex];
                if (!header) continue;
                compiledPlaceholders.push({
                    header: header,
                    regex: new RegExp('{{' + _App_escapeRegExp(header) + '}}', 'g')
                });
            }

            var stats = _App_BatchProcessor('MAIL_MERGE', pendingRows, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    status: "",
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().trim().toUpperCase();
                if (action !== "SEND" && action !== "DRAFT") return null;

                var targetTo = item['To'];
                var targetCc = item['CC'];
                var targetBcc = item['BCC'];
                var targetThreadId = item['Thread ID or Subject'];
                var targetAttachments = item['Attachments'];

                if (!targetTo && !targetThreadId) throw new Error("Missing Email To");

                var emailBody = template.body;
                var emailSubject = template.subject;

                compiledPlaceholders.forEach(function (pDef) {
                    var value = item[pDef.header];
                    var valStr = (value === undefined || value === null || value === "") ? "" : String(value);
                    var bodyVal = valStr.replace(/\r?\n/g, '<br>');

                    emailBody = emailBody.replace(pDef.regex, () => bodyVal);
                    emailSubject = emailSubject.replace(pDef.regex, () => valStr);
                });

                var remainingPlaceholders = [];
                var unmatched;
                var regexExtract = /\{\{([^{}]+)\}\}/g;
                while ((unmatched = regexExtract.exec(emailBody)) !== null) {
                    remainingPlaceholders.push(unmatched[1]);
                }
                while ((unmatched = regexExtract.exec(emailSubject)) !== null) {
                    remainingPlaceholders.push(unmatched[1]);
                }
                var allRemaining = [...new Set(remainingPlaceholders)];
                if (allRemaining.length > 0) {
                    throw new Error("Missing columns for: " + allRemaining.join(', '));
                }

                var finalAttachments = [...template.attachments];
                if (targetAttachments) {
                    var files = targetAttachments.split(',');
                    for (var f = 0; f < files.length; f++) {
                        var blob = _App_getDriveAttachment(files[f].trim());
                        if (blob) finalAttachments.push(blob);
                    }
                }

                rowUpdates.status = _App_sendOrDraftEmail({
                    action: action,
                    to: targetTo,
                    cc: targetCc,
                    bcc: targetBcc,
                    subject: emailSubject,
                    body: emailBody,
                    attachments: finalAttachments,
                    threadIdOrSubject: targetThreadId
                });
                rowUpdates.action = "";

                Logger.info(SyncEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
                return rowUpdates;
            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('MAIL_MERGE', batchResults);
                }
            });

            var finalMsg = "Successfully processed " + stats.processedCount + " emails.";
            if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
            if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

            return _App_ok(finalMsg);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Mail Merge sidebar and ensures the sheet exists. */
function MailMerge_openSidebar() {
  return Logger.run('MAIL_MERGE', 'Open Sidebar', function () {
    _App_launchTool('MAIL_MERGE');
  });
}


// --- FILE: tools/MailSender/Code.js ---
/**
 * Mail Sender
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_SENDER', {
    SHEET_NAME: SHEET_NAMES.MAIL_SENDER,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.MAIL_SENDER,
    MENU_LABEL: SHEET_NAMES.MAIL_SENDER,
    MENU_ENTRYPOINT: 'MailSender_openSidebar',
    MENU_ORDER: 40,
    SIDEBAR_HTML: 'tools/MailSender/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'To', type: 'EMAIL_LIST' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'TEXT' },
            { header: 'Email Subject', type: 'TEXT' },
            { header: 'Email Body', type: 'TEXT' },
            { header: 'PDF HTML', type: 'TEXT' },
            { header: 'PDF Name', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to send custom emails:</p><ol><li><strong>Fill Content:</strong> Enter subjects and bodies directly into the sheet.</li><li><strong>Action:</strong> Set the Action column to <code>SEND</code> or <code>DRAFT</code>.</li><li><strong>Run:</strong> Click <strong>Send Custom Mail</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about PDF HTML, Attachments, and BCC.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Email Subject / Body:</strong> The actual text sent to recipients.</li><li><strong>PDF HTML:</strong> Raw HTML that will be converted to a PDF attachment.</li><li><strong>PDF Name:</strong> The name of the generated PDF file.</li><li><strong>Attachments:</strong> Comma-separated Drive File IDs or URLs.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "HTML-to-PDF conversion and bulk sending.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>HTML to PDF:</strong> Use <code>&lt;table&gt;</code>, <code>&lt;h1&gt;</code>, and inline CSS for professional PDF attachments.</li><li><strong>Bulk Sending:</strong> Ideal for unique, one-off messages that don't follow a fixed template.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getQuota: function () {
            var quota = MailApp.getRemainingDailyQuota();
            return _App_ok('Remaining quota: ' + quota, { remaining: quota });
        },
        executeActions: function () {
            var pendingRows = SheetManager.readPendingObjects('MAIL_SENDER', { useDisplayValues: true });

            if (pendingRows.length === 0) return _App_ok("Nothing to do! No 'SEND' or 'DRAFT' actions pending.");

            var stats = _App_BatchProcessor('MAIL_SENDER', pendingRows, function (item, index) {
                var rowUpdates = {
                    action: item['Action'],
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().trim().toUpperCase();
                if (action !== "SEND" && action !== "DRAFT") return null;

                var targetTo = item['To'];
                var targetCc = item['CC'];
                var targetBcc = item['BCC'];
                var targetThreadId = item['Thread ID or Subject'];
                var targetAttachments = item['Attachments'];
                var targetPdfHtml = item['PDF HTML'];
                var targetPdfName = item['PDF Name'];

                if (!targetTo && !targetThreadId) throw new Error("⚠️ Missing Email To");

                var emailSubject = item['Email Subject'];
                var emailBody = item['Email Body'] ? String(item['Email Body']).replace(/\r?\n/g, '<br>') : "";

                if (!emailSubject && !targetThreadId) throw new Error("⚠️ Missing Email Subject");
                if (!emailBody) throw new Error("⚠️ Missing Email Body");

                var finalAttachments = [];
                if (targetAttachments) {
                    var files = targetAttachments.split(',');
                    for (var f = 0; f < files.length; f++) {
                        var blob = _App_getDriveAttachment(files[f].trim());
                        if (blob) finalAttachments.push(blob);
                    }
                }

                if (targetPdfHtml) {
                    var defaultFileName = "document.pdf";
                    var fileName = targetPdfName ? targetPdfName.toString().trim() : defaultFileName;
                    if (!fileName.toLowerCase().endsWith(".pdf")) {
                        fileName += ".pdf";
                    }
                    var pdfBlob = Utilities.newBlob(targetPdfHtml, 'text/html', fileName).getAs('application/pdf');
                    finalAttachments.push(pdfBlob);
                }

                rowUpdates.status = _App_sendOrDraftEmail({
                    action: action,
                    to: targetTo,
                    cc: targetCc,
                    bcc: targetBcc,
                    subject: emailSubject,
                    body: emailBody,
                    attachments: finalAttachments,
                    threadIdOrSubject: targetThreadId
                });
                rowUpdates.action = "";
                return rowUpdates;

            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('MAIL_SENDER', batchResults);
                }
            });

            var finalResult = stats.processedCount + " actions processed!";
            return _App_ok(finalResult);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Mail Sender sidebar and ensures the sheet exists. */
function MailSender_openSidebar() {
  return Logger.run('MAIL_SENDER', 'Open Sidebar', function () {
    _App_launchTool('MAIL_SENDER');
  });
}


// --- FILE: tools/PipelineControl/Code.js ---
/**
 * Pipeline
 * Version: 4.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('PIPELINE', {
    SHEET_NAME: SHEET_NAMES.PIPELINE,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.PIPELINE,
    MENU_LABEL: SHEET_NAMES.PIPELINE,
    MENU_ENTRYPOINT: 'PipelineControl_openSidebar',
    MENU_ORDER: 100,
    SIDEBAR_HTML: 'tools/PipelineControl/Sidebar',
    SIDEBAR_WIDTH: 300,
    FORMAT_CONFIG: {
        conditionalRules: [
            { type: 'custom', formula: '=$A2="Enabled"', color: SHEET_THEME.STATUS.SUCCESS, scope: 'actionOnly', actionCol: 'A' }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Enabled', 'Disabled'], width: 100 },
            { header: 'Status', type: 'STATUS' },
            { header: 'Pipeline Name', type: 'TEXT' },
            { header: 'Source URL', type: 'URL' },
            { header: 'Source Range', type: 'TEXT' },
            { header: 'Destination URL', type: 'URL' },
            { header: 'Destination Cell', type: 'TEXT' },
            { header: 'Sync Interval', type: 'DROPDOWN', options: ['Manual Only', '15 min', '30 min', '1 hour', '4 hours', '12 hours', '1 day'] },
            { header: 'Last Run Time', type: 'DATETIME' }
        ]
    },
    ACTIONS: {
        getSystemStatus: function () {
            var enabled = _App_getProperty(APP_PROPS.SYSTEM_ENABLED);
            var status = enabled === null ? 'false' : enabled;
            return _App_ok('System status retrieved.', status);
        },
        setSystemStatus: function (isEnabled) {
            _App_setProperty(APP_PROPS.SYSTEM_ENABLED, isEnabled.toString());
            _PipelineControl_manageTrigger(isEnabled);
            return _App_ok('System status updated.', isEnabled);
        },
        processScheduled: function () {
            return _PipelineControl_processPipelinesInternal();
        },
        runAll: function () {
            return _PipelineControl_runAllPipelinesInternal();
        },
        getDashboardData: function () {
            var pendingPipelines = SheetManager.readPendingObjects('PIPELINE', { actionColName: 'Action' });
            var allPipelines = SheetManager.readAllObjects('PIPELINE');

            var enabledCount = 0;
            var disabledCount = 0;
            var list = [];

            allPipelines.forEach(function (p) {
                var action = (p['Action'] || '').toString().trim();
                var isEnabled = action === 'Enabled';
                if (isEnabled) enabledCount++;
                else disabledCount++;

                list.push({
                    rowIndex: p._rowNumber,
                    name: p['Pipeline Name'] || ('Row ' + p._rowNumber),
                    lastStatus: p['Status'] || '',
                    lastRun: p['Last Run Time'] || '',
                    isEnabled: isEnabled
                });
            });

            return _App_ok('Pipeline dashboard data retrieved.', {
                enabledCount: enabledCount,
                disabledCount: disabledCount,
                pipelines: list
            });
        },
        runSelected: function (rowIndexes) {
            if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) {
                return _App_fail("No pipelines selected.");
            }

            var allPipelines = SheetManager.readAllObjects('PIPELINE');
            var selectedPipelines = allPipelines.filter(function (p) {
                return rowIndexes.indexOf(p._rowNumber) !== -1;
            });

            if (selectedPipelines.length === 0) {
                return _App_fail("No matching pipelines found.");
            }

            var stats = _App_BatchProcessor('PIPELINE', selectedPipelines, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    status: "",
                    lastRun: "",
                    _rowNumber: item._rowNumber
                };

                try {
                    _PipelineControl_runPipeline(item);
                    rowUpdates.status = _App_formatStatus('SUCCESS', 'Success');
                    rowUpdates.lastRun = _App_formatDateTime(new Date());
                } catch (e) {
                    throw e;
                }

                return rowUpdates;
            }, {
                onBatchComplete: function (results) {
                    _App_batchPatchResults('PIPELINE', results);
                }
            });

            return _App_ok("Successfully executed " + stats.processedCount + " pipelines.");
        },
        formatSheet: function () {
            var sheet = SheetManager.getSheet('PIPELINE');
            if (sheet) {
                _App_applyBodyFormatting(sheet, sheet.getLastRow(), SyncEngine.getTool('PIPELINE').FORMAT_CONFIG);
                return _App_ok("Sheet formatted successfully.");
            }
            return _App_fail("Sheet not found.");
        }
    }
});
/**
 * Column Mappings (0-indexed):
 * 0: Action (Enabled/Disabled)     1: Status
 * 2: Pipeline Name                3: Source URL
 * 4: Source Range                 5: Destination URL
 * 6: Destination Cell             7: Sync Interval
 * 8: Last Run Time
 */

var PIPELINE_NON_DATA_ROWS = 1;

// --- SIDEBAR ---

/** Opens the Pipeline sidebar, creating the sheet if needed. */
function PipelineControl_openSidebar() {
    return Logger.run('PIPELINE', 'Open Sidebar', function () {
        _App_launchTool('PIPELINE');
    });
}

/**
 * Manages the background execution trigger for Pipeline sync.
 * @param {boolean} isEnabled Whether the system should be active.
 */
function _PipelineControl_manageTrigger(isEnabled) {
    var functionName = 'PipelineControl_processPipelines';
    var triggers = ScriptApp.getProjectTriggers();
    
    // Remove existing triggers to avoid duplicates or when disabling
    for (var i = 0; i < triggers.length; i++) {
        var handler = triggers[i].getHandlerFunction();
        if (handler === functionName) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
    
    // Create new trigger if enabled
    if (isEnabled) {
        ScriptApp.newTrigger(functionName)
            .timeBased()
            .everyMinutes(15) // Check every 15 mins (minimum interval supported)
            .create();
        Logger.info('PIPELINE', 'System', 'Background sync trigger created (15m interval).');
    } else {
        Logger.info('PIPELINE', 'System', 'Background sync trigger removed.');
    }
}

function _PipelineControl_processPipelinesInternal() {
    return Logger.run('PIPELINE', 'Scheduled Execution', function () {
        return _App_withDocumentLock('PIPELINE_PROCESS', function () {
            if (_App_getProperty(APP_PROPS.SYSTEM_ENABLED) !== 'true') {
                Logger.info('PIPELINE', 'Global', "System is globally disabled. Skipping execution.");
                return _App_ok('System is globally disabled. Skipping execution.');
            }

            var pendingPipelines = SheetManager.readPendingObjects('PIPELINE', { actionColName: 'Action' });

            var activeScheduled = pendingPipelines.filter(function(p) {
                var isEnabled = (String(p['Action']).toLowerCase() === 'enabled') || (p['Action'] === true);
                return isEnabled && _PipelineControl_shouldRun(p);
            });

            if (activeScheduled.length === 0) {
                Logger.info('PIPELINE', 'Global', "No pipelines scheduled to run.");
                return _App_ok('No pipelines scheduled to run.');
            }

            var sheet = SheetManager.getSheet('PIPELINE');

            _App_BatchProcessor('PIPELINE', activeScheduled, function (item) {
                var statusMsg = _PipelineControl_runPipeline(item);
                return { _rowNumber: item._rowNumber, status: statusMsg };
            }, {
                onBatchComplete: function(results) {
                    var rowNumbers = [];
                    var patchData = [];
                    results.forEach(function (res) {
                        if (res && res._rowNumber) {
                            rowNumbers.push(res._rowNumber);
                            if (res.isError) {
                                patchData.push(_App_makeStatusPatch(res._rowNumber, 'ERROR', res.error, { 'Last Run Time': new Date() }));
                            } else {
                                patchData.push(_App_makeRowPatch(res._rowNumber, { 'Status': _App_formatStatus('SUCCESS', res.status), 'Last Run Time': new Date() }));
                            }
                        }
                    });
                    if (rowNumbers.length > 0) SheetManager.batchPatchRows('PIPELINE', rowNumbers, patchData);
                }
            });
            return _App_ok('Scheduled pipelines processed.', { processedCount: activeScheduled.length });
        });
    });
}

function _PipelineControl_runAllPipelinesInternal() {
    return Logger.run('PIPELINE', 'Run All', function () {
        return _App_withDocumentLock('PIPELINE_RUN_ALL', function () {
            var pendingPipelines = SheetManager.readPendingObjects('PIPELINE', { actionColName: 'Action' });
            var enabledPipelines = pendingPipelines.filter(function(p) {
                return (String(p['Action']).toLowerCase() === 'enabled') || (p['Action'] === true);
            });

            if (enabledPipelines.length === 0) return _App_ok('No enabled pipelines to run.');

            var sheet = SheetManager.getSheet('PIPELINE');

            var stats = _App_BatchProcessor('PIPELINE', enabledPipelines, function (item) {
                var statusMsg = _PipelineControl_runPipeline(item);
                return { _rowNumber: item._rowNumber, status: statusMsg };
            }, {
                onBatchComplete: function(results) {
                    var rowNumbers = [];
                    var patchData = [];
                    results.forEach(function (res) {
                        if (res && res._rowNumber) {
                            rowNumbers.push(res._rowNumber);
                            if (res.isError) {
                                patchData.push(_App_makeStatusPatch(res._rowNumber, 'ERROR', res.error, { 'Last Run Time': new Date() }));
                            } else {
                                patchData.push(_App_makeRowPatch(res._rowNumber, { 'Status': _App_formatStatus('SUCCESS', res.status), 'Last Run Time': new Date() }));
                            }
                        }
                    });
                    if (rowNumbers.length > 0) SheetManager.batchPatchRows('PIPELINE', rowNumbers, patchData);
                }
            });

            var resultMsg = 'Execution complete. Processed ' + stats.processedCount + ' pipelines.';
            if (stats.timeLimitReached) resultMsg = '⏳ Time limit reached. ' + resultMsg;
            return _App_ok(resultMsg);
        });
    });
}

function _PipelineControl_shouldRun(item) {
    var intervalStr = String(item['Sync Interval']);
    var lastRun = item['Last Run Time'];

    if (intervalStr === "Manual Only") return false;
    if (!lastRun || lastRun === "") return true;

    var lastRunTime = new Date(lastRun).getTime();
    var now = new Date().getTime();
    var diffMs = now - lastRunTime;
    var diffMins = diffMs / (1000 * 60);
    var diffHours = diffMins / 60;

    if (intervalStr.includes("hour")) {
        var hours = parseInt(intervalStr.match(/\d+/)) || 1;
        return diffHours >= hours;
    } else if (intervalStr.includes("min")) {
        var mins = parseInt(intervalStr.match(/\d+/)) || 15;
        return diffMins >= mins;
    } else if (intervalStr.includes("day") || intervalStr.includes("24 hours")) {
        return diffHours >= 24;
    }

    return false;
}

function _PipelineControl_runPipeline(rowObj) {
    var pipelineName = rowObj['Pipeline Name'];

    function getSheetFromUrl(url) {
        var match = url.match(/gid=([0-9]+)/);
        var ss = SpreadsheetApp.openByUrl(url);
        if (match) {
            var gid = parseInt(match[1], 10);
            var sheets = ss.getSheets();
            for (var i = 0; i < sheets.length; i++) {
                if (sheets[i].getSheetId() === gid) return sheets[i];
            }
        }
        return ss.getSheets()[0];
    }

    var sourceUrl = rowObj['Source URL'];
    var sourceRangeA1 = rowObj['Source Range'];
    var destUrl = rowObj['Destination URL'];
    var destStartCell = rowObj['Destination Cell'];

    if (!destStartCell || destStartCell.toString().trim() === "") {
        destStartCell = "A1";
    }

    if (!sourceUrl || !destUrl) {
        throw new Error(`Missing details -> Source URL: ${sourceUrl ? 'OK' : 'Blank'}, Dest URL: ${destUrl ? 'OK' : 'Blank'}`);
    }

    var sSheet;
    try {
        sSheet = getSheetFromUrl(sourceUrl);
    } catch (e) {
        throw new Error("Cannot access Source URL (Check permissions or URL validity)");
    }
    if (!sSheet) throw new Error("Source sheet not found");

    var values;
    var isSheetLevelSync = false;
    if (sourceRangeA1 && String(sourceRangeA1).trim() !== "") {
        values = sSheet.getRange(String(sourceRangeA1).trim()).getValues();
    } else {
        isSheetLevelSync = true;
        values = sSheet.getDataRange().getValues();
    }
    
    if (values.length === 0) throw new Error("Source range empty");

    var dSheet;
    try {
        dSheet = getSheetFromUrl(destUrl);
    } catch (e) {
        throw new Error("Cannot access Destination URL (Check permissions or URL validity)");
    }
    if (!dSheet) throw new Error("Destination sheet not found");

    var numRows = values.length;
    var numCols = values[0].length;

    if (numRows > 0 && numCols > 0) {
        if (isSheetLevelSync) {
            dSheet.clearContents();
        }

        var destRange = dSheet.getRange(destStartCell);
        var startRow = destRange.getRow();
        var startCol = destRange.getColumn();

        var reqRows = startRow + numRows - 1;
        var reqCols = startCol + numCols - 1;

        if (dSheet.getMaxRows() < reqRows) {
            dSheet.insertRowsAfter(dSheet.getMaxRows(), reqRows - dSheet.getMaxRows());
        }
        if (dSheet.getMaxColumns() < reqCols) {
            dSheet.insertColumnsAfter(dSheet.getMaxColumns(), reqCols - dSheet.getMaxColumns());
        }

        dSheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
        if (isSheetLevelSync) SpreadsheetApp.flush();
        
        return "Synced " + numRows + " rows.";
    } else {
        return "No data found in source range.";
    }
}

function PipelineControl_processPipelines() {
    return SyncEngine.runAction('PIPELINE', 'processScheduled');
}


// --- FILE: tools/TasksSync/Code.js ---
/**
 * Google Tasks Sync
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('TASKS_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Tasks API', test: function() { return typeof Tasks !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.TASKS_SYNC,
    TITLE: SHEET_NAMES.TASKS_SYNC,
    MENU_LABEL: SHEET_NAMES.TASKS_SYNC,
    MENU_ENTRYPOINT: 'TasksSync_openSidebar',
    MENU_ORDER: 45,
    SIDEBAR_HTML: 'tools/TasksSync/Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 2,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Task List Name', type: 'DROPDOWN', options: function() {
                try {
                  var lists = _App_callWithBackoff(function () {
                    return Tasks.Tasklists.list().items;
                  });
                  return (lists || []).map(function(t) { return t.title; });
                } catch(e) {
                  return [];
                }
              }
            },
            { header: 'Task Title', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Due Date', type: 'DATE' },
            { header: 'Completed?', type: 'DROPDOWN', options: ['Completed', 'Not Completed'] },
            { header: 'Task ID', type: 'ID' },
            { header: 'Task List ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to sync tasks:</p><ol><li><strong>Define Actions:</strong> Set the Action column to <code>CREATE</code>, <code>UPDATE</code>, or <code>DELETE</code>.</li><li><strong>Fill Details:</strong> Provide Task Title, Task List Name, and optional Due Date/Description.</li><li><strong>Push:</strong> Click <strong>Push Changes</strong> in the sidebar to execute actions on Google Tasks.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Bulk Operations",
                shortDesc: "Set Action column to CREATE, UPDATE, or DELETE.",
                tooltipId: "help-bulk-ops",
                tooltipContent: "<p><strong>Bulk Operations</strong></p><ul><li><strong>CREATE:</strong> Add a new task to a task list.</li><li><strong>UPDATE:</strong> Modify title, description, due date, or list.</li><li><strong>DELETE:</strong> Permanent deletion of the task.</li></ul>"
            },
            {
                icon: "check-circle",
                color: "var(--primary-light)",
                label: "Completed Status",
                shortDesc: "Choose UPDATE and change Completed? column.",
                tooltipId: "help-completed",
                tooltipContent: "<p><strong>Marking Completed</strong></p><p>To change completion status, select <code>UPDATE</code> in the Action column and set <code>Completed?</code> to <code>Completed</code> or <code>Not Completed</code>.</p>"
            },
            {
                icon: "move",
                color: "var(--warning)",
                label: "Move Tasks",
                shortDesc: "Change Task List Name and set Action to UPDATE.",
                tooltipId: "help-move-tasks",
                tooltipContent: "<p><strong>Move Tasks</strong></p><p>You can migrate tasks between lists by editing the <code>Task List Name</code> and choosing <code>UPDATE</code> in the Action column.</p>"
            }
        ]
    },
    ACTIONS: {
        pull: function () {
            var TARGET_SHEET_NAME = SHEET_NAMES.TASKS_SYNC;
            var allLists = _App_callWithBackoff(function () {
                return Tasks.Tasklists.list().items;
            }) || [];

            var outputObjects = [];

            allLists.forEach(function (list) {
                try {
                    var tasksResult = _App_callWithBackoff(function () {
                        return Tasks.Tasks.list(list.id, { showCompleted: true, showHidden: true });
                    });
                    var tasks = tasksResult.items || [];
                    tasks.forEach(function (t) {
                        var formattedDue = "";
                        if (t.due) {
                            var d = new Date(t.due);
                            if (!isNaN(d.getTime())) {
                                formattedDue = _App_formatDateTime(d, "MM/dd/yyyy");
                            }
                        }

                        outputObjects.push({
                            'Action': "",
                            'Status': "",
                            'Task List Name': list.title,
                            'Task Title': t.title || "",
                            'Description': t.notes || "",
                            'Due Date': formattedDue,
                            'Completed?': t.status === 'completed' ? 'Completed' : 'Not Completed',
                            'Task ID': t.id,
                            'Task List ID': list.id
                        });
                    });
                } catch (err) {
                    Logger.warn('TASKS_SYNC', 'Pull List Error', 'List ' + list.title + ': ' + err.message);
                }
            });

            // Sort by Task List Name, then Task Title
            outputObjects.sort(function (a, b) {
                var nameA = (a['Task List Name'] || "").toLowerCase();
                var nameB = (b['Task List Name'] || "").toLowerCase();
                if (nameA < nameB) return -1;
                if (nameA > nameB) return 1;

                var titleA = (a['Task Title'] || "").toLowerCase();
                var titleB = (b['Task Title'] || "").toLowerCase();
                if (titleA < titleB) return -1;
                if (titleA > titleB) return 1;
                return 0;
            });

            // Populate sheet
            SheetManager.overwriteObjects('TASKS_SYNC', outputObjects);

            var summary = 'Successfully imported ' + outputObjects.length + " tasks into '" + TARGET_SHEET_NAME + "'.";
            return _App_ok(summary);
        },
        push: function () {
            var pendingItems = SheetManager.readPendingObjects('TASKS_SYNC');

            if (pendingItems.length === 0) return _App_ok("No pending actions found.");

            var allLists = _App_callWithBackoff(function () {
                return Tasks.Tasklists.list().items;
            }) || [];

            var listMapByName = new Map();
            var listMapById = new Map();
            allLists.forEach(function (l) {
                listMapByName.set(l.title, l.id);
                listMapById.set(l.id, l);
            });

            var stats = _App_BatchProcessor('TASKS_SYNC', pendingItems, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    taskId: item['Task ID'] ? String(item['Task ID']) : null,
                    listId: item['Task List ID'] ? String(item['Task List ID']) : null,
                    status: "",
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().toUpperCase();
                var targetListName = item['Task List Name'];
                var targetListId = listMapByName.get(targetListName);

                var taskData = {
                    title: item['Task Title'],
                    description: item['Description'],
                    dueDate: item['Due Date'],
                    completedVal: item['Completed?']
                };

                if (action !== "DELETE" && !taskData.title) throw new Error("⚠️ Data Error: Missing Task Title");

                switch (action) {
                    case "CREATE":
                        if (!targetListName) throw new Error("⚠️ Data Error: Missing Task List Name");
                        if (!targetListId) throw new Error("⚠️ Data Error: Task List '" + targetListName + "' not found");

                        var dueStr = null;
                        if (taskData.dueDate) {
                            var d = new Date(taskData.dueDate);
                            if (isNaN(d.getTime())) throw new Error("⚠️ Data Error: Invalid Due Date format");
                            d.setUTCHours(0, 0, 0, 0);
                            dueStr = d.toISOString();
                        }

                        var taskResource = {
                            title: taskData.title,
                            notes: taskData.description || "",
                            status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction'
                        };
                        if (dueStr) taskResource.due = dueStr;

                        var newTask = _App_callWithBackoff(function () {
                            return Tasks.Tasks.insert(taskResource, targetListId);
                        });

                        rowUpdates.taskId = newTask.id;
                        rowUpdates.listId = targetListId;
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Created");
                        rowUpdates.action = "";
                        break;

                    case "UPDATE":
                        if (!rowUpdates.taskId) throw new Error("⚠️ Data Error: Missing Task ID");
                        if (!rowUpdates.listId) throw new Error("⚠️ Data Error: Missing Task List ID");

                        if (targetListName && !targetListId) throw new Error("⚠️ Data Error: Task List '" + targetListName + "' not found");

                        // Identity Check: If target task list name doesn't match current Task List ID, perform MOVE
                        if (targetListId && rowUpdates.listId && targetListId !== rowUpdates.listId) {
                            rowUpdates = _TasksSync_processMove(rowUpdates, targetListId, taskData);
                            break;
                        }

                        var dueStrUpdate = null;
                        if (taskData.dueDate) {
                            var dUpdate = new Date(taskData.dueDate);
                            if (isNaN(dUpdate.getTime())) throw new Error("⚠️ Data Error: Invalid Due Date format");
                            dUpdate.setUTCHours(0, 0, 0, 0);
                            dueStrUpdate = dUpdate.toISOString();
                        }

                        var updateResource = {
                            title: taskData.title,
                            notes: taskData.description || "",
                            status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction',
                            due: dueStrUpdate
                        };

                        _App_callWithBackoff(function () {
                            return Tasks.Tasks.patch(updateResource, rowUpdates.listId, rowUpdates.taskId);
                        });

                        rowUpdates.status = _App_formatStatus('SUCCESS', "Updated");
                        rowUpdates.action = "";
                        break;

                    case "DELETE":
                        if (!rowUpdates.taskId) throw new Error("⚠️ Data Error: Missing Task ID");
                        if (!rowUpdates.listId) throw new Error("⚠️ Data Error: Missing Task List ID");

                        try {
                            _App_callWithBackoff(function () {
                                Tasks.Tasks.remove(rowUpdates.listId, rowUpdates.taskId);
                            });
                            rowUpdates.status = _App_formatStatus('SUCCESS', "Deleted");
                            rowUpdates.action = "";
                        } catch (e) {
                            if (e.message.indexOf('404') !== -1 || e.message.indexOf('not found') !== -1) {
                                rowUpdates.status = _App_formatStatus('WARNING', "Already Deleted");
                                rowUpdates.action = "";
                            } else {
                                throw e;
                            }
                        }
                        break;

                    default:
                        rowUpdates.status = _App_formatStatus('WARNING', "Unknown Action '" + action + "'");
                }

                return rowUpdates;
            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('TASKS_SYNC', batchResults, function (res) {
                        return {
                            'Task ID': res.taskId,
                            'Task List ID': res.listId
                        };
                    });
                }
            });

            return _App_ok("Sync Complete. Processed: " + stats.processedCount);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Tasks sidebar and ensures the sheet exists. */
function TasksSync_openSidebar() {
  return Logger.run('TASKS_SYNC', 'Open Sidebar', function () {
    _App_launchTool('TASKS_SYNC');
  });
}

// --- INTERNAL HELPERS ---

/** Handles moving a task from one list to another by copying then deleting the old one */
function _TasksSync_processMove(rowUpdates, targetListId, taskData) {
  var dueStr = null;
  if (taskData.dueDate) {
    var d = new Date(taskData.dueDate);
    if (!isNaN(d.getTime())) {
      d.setUTCHours(0, 0, 0, 0);
      dueStr = d.toISOString();
    }
  }

  var taskResource = {
    title: taskData.title,
    notes: taskData.description || "",
    status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction'
  };
  if (dueStr) taskResource.due = dueStr;

  // Insert into target list
  var newTask = _App_callWithBackoff(function () {
    return Tasks.Tasks.insert(taskResource, targetListId);
  });

  var deleteWarning = "";
  // Delete from original list
  if (rowUpdates.listId && rowUpdates.taskId) {
    try {
      _App_callWithBackoff(function () {
        Tasks.Tasks.remove(rowUpdates.listId, rowUpdates.taskId);
      });
    } catch (e) {
      deleteWarning = " (⚠️ Could not delete old task)";
    }
  }

  rowUpdates.taskId = newTask.id;
  rowUpdates.listId = targetListId;
  rowUpdates.status = _App_formatStatus('SUCCESS', "Moved") + deleteWarning;
  rowUpdates.action = "";

  return rowUpdates;
}

