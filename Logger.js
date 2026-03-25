/**
 * Developer Logging System
 * Version: 5.0 (Plugin Architecture — registers with BeaverEngine)
 */

BeaverEngine.registerTool('LOGS', {
    SHEET_NAME: SHEET_NAMES.LOGS,
    TITLE: '🛠️ Developer Log',
    MENU_LABEL: '🛠️ Developer Log',
    MENU_ENTRYPOINT: 'Logger_showSidebar',
    MENU_ORDER: 110,
    SIDEBAR_HTML: 'Logger_Sidebar',
    SIDEBAR_WIDTH: 360,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [150, 200, 85, 150, 150, 400, 150],
    FORMAT_CONFIG: {
        skipActionColoring: true,
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [
            { type: 'custom', formula: '=OR($C2="ERROR", $C2="FAIL")', color: SHEET_THEME.STATUS.ERROR, scope: 'fullRow' },
            { type: 'custom', formula: '=$C2="SUCCESS"', color: SHEET_THEME.STATUS.SUCCESS, scope: 'fullRow' },
            { type: 'custom', formula: '=$C2="WARN"', color: SHEET_THEME.STATUS.WARNING, scope: 'fullRow' },
            { type: 'custom', formula: '=$C2="INFO"', color: SHEET_THEME.ACTION, scope: 'fullRow' }
        ],
        COL_SCHEMA: [
            { header: 'Timestamp', type: 'TEXT' },
            { header: 'Run ID', type: 'ID' },
            { header: 'Level', type: 'TEXT' },
            { header: 'Source', type: 'TEXT' },
            { header: 'Reference', type: 'TEXT' },
            { header: 'Message', type: 'TEXT' },
            { header: 'User', type: 'TEXT' }
        ]
    }
});

var Logger = (function () {
    var CACHE_KEY = typeof CACHE_KEYS !== 'undefined' ? CACHE_KEYS.LOGS : 'BEAVER_DEBUG_LOGS';
    var currentRunId = null;

    var currentSteps = []; // Breadcrumbs for current run

    /**
     * @typedef {Object} LogEntry
     * @property {string} timestamp
     * @property {string} runId
     * @property {string} level
     * @property {string} source
     * @property {string} reference
     * @property {string} message
     * @property {string} user
     * @property {string} [context]
     */

    /**
     * Safely stringifies complex or circular objects for deep context capture.
     */
    function _safeStringify(obj, maxDepth = 4) {
        if (obj === undefined) return "undefined";
        if (obj === null) return "null";
        if (typeof obj !== 'object') return String(obj);

        const cache = new Set();
        return JSON.stringify(obj, (key, value) => {
            if (typeof value === 'object' && value !== null) {
                if (cache.has(value)) {
                    return '[Circular]'; // Discard circular reference
                }
                cache.add(value);
            }
            if (value instanceof Error) {
                 return { name: value.name, message: value.message, stack: value.stack };
            }
            return value;
        }, 2); // 2 spaces for readability
    }


    function isEnabled() {
        return _App_getProperty(APP_PROPS.ENABLE_DEBUG_LOGGING) === 'true';
    }

    function setRunId(id) {
        currentRunId = id || Utilities.getUuid();
        currentSteps = []; // Reset breadcrumbs for new run
        return currentRunId;
    }

    /**
     * Initializes or removes the Log sheet based on state.
     */
    function setLoggingState(enabled) {
        _App_setProperty(APP_PROPS.ENABLE_DEBUG_LOGGING, enabled ? 'true' : 'false');

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetName = SHEET_NAMES.LOGS;
        var sheet = ss.getSheetByName(sheetName);

        if (enabled) {
            if (!sheet) {
                sheet = _App_ensureSheetExists('LOGS');
            }
            if (sheet && sheet.isSheetHidden()) sheet.showSheet();
        } else {
            CacheService.getDocumentCache().remove(CACHE_KEY);
        }
    }

    function _queueLog(level, source, reference, message, ctx) {
        if (!isEnabled()) return;

        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        var user = Session.getActiveUser().getEmail() || 'Unknown';
        var runId = currentRunId || 'N/A';
        
        // Enhance message with deep context if provided
        var contextStr = ctx ? "\n\n--- Deep Context ---\n" + _safeStringify(ctx) : "";
        var msgStr = (typeof message === 'object' ? _safeStringify(message) : String(message)) + contextStr;
        
        var logEntry = [timestamp, runId, level, source, reference || '', msgStr, user];

        var cache = CacheService.getDocumentCache();
        var existing = cache.get(CACHE_KEY);
        var logs = existing ? JSON.parse(existing) : [];

        logs.push(logEntry);

        // Keep cache size manageable (CacheService has 100KB limit per key).
        // When we overflow, flush the batch (minus the new entry) and keep the new entry in cache.
        var maxCacheItems = 50;
        if (logs.length > maxCacheItems) {
            var overflow = logs.slice(0, logs.length - 1); // everything except the newest entry
            _flushToSheet(overflow);
            logs = [logEntry]; // keep only the newest entry in cache for the next flush
        }

        cache.put(CACHE_KEY, JSON.stringify(logs), 21600);
    }

    function info(source, reference, message, ctx) { _queueLog('INFO', source, reference, message, ctx); }
    function success(source, reference, message, ctx) { _queueLog('SUCCESS', source, reference, message, ctx); }
    function warn(source, reference, message, ctx) { _queueLog('WARN', source, reference, message, ctx); }
    function debug(source, reference, message, ctx) { _queueLog('DEBUG', source, reference, message, ctx); }
    
    /**
     * Records a breadcrumb step for the current execution.
     */
    function step(source, reference, stepName) {
        currentSteps.push(`[${new Date().toISOString()}] ${source} - ${stepName}`);
        debug(source, reference, "STEP: " + stepName);
    }
    
    function error(source, reference, err, context) {
        var msg = typeof err === 'string' ? err : (err.message + (err.stack ? '\n' + err.stack : ''));

        var fullContext = Object.assign({}, context || {});
        
        // Auto-extract deep context explicitly attached to error objects
        if (typeof err === 'object' && err !== null && err.localContext) {
            fullContext.errorDetails = err.localContext;
        }

        if (currentSteps.length > 0) fullContext.executionBreadcrumbs = currentSteps.slice();
        
        _queueLog('ERROR', source, reference, msg, fullContext);
    }

    /**
     * Internal function to write logs to the sheet.
     */
    function _flushToSheet(logs) {
        if (!logs || logs.length === 0) return;
        
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
        if (!sheet) {
            // Try to create it if it disappeared
            try { sheet = _App_ensureSheetExists('LOGS'); } catch(e) { return; }
        }

        // Append rows to the bottom of the data range
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, logs.length, logs[0].length).setValues(logs);

        // Enforce LOGGER_MAX_ROWS — prune oldest rows from top if configured
        var maxRowsRaw = _App_getProperty(APP_PROPS.LOGGER_MAX_ROWS);
        var maxRows = maxRowsRaw ? parseInt(maxRowsRaw, 10) : 0;
        if (maxRows > 0) {
            var currentDataRows = sheet.getLastRow() - 1;
            if (currentDataRows > maxRows) {
                var excess = currentDataRows - maxRows;
                sheet.deleteRows(2, excess); // Delete from row 2 (oldest logs) downwards
            }
        }

        // Apply formatting (pass the row where we started inserting)
        var logConfig = BeaverEngine.getTool('LOGS');
        if (logConfig && logConfig.FORMAT_CONFIG) {
            _App_applyBodyFormatting(sheet, startRow - 1, logConfig.FORMAT_CONFIG);
        }
    }

    /**
     * Flushes cached logs to the sheet.
     */
    function flushLogs() {
        if (!isEnabled()) return;

        var cache = CacheService.getDocumentCache();
        var existing = cache.get(CACHE_KEY);
        if (!existing) return;

        var logs = [];
        try { logs = JSON.parse(existing); } catch (e) { return; }
        if (logs.length === 0) return;

        cache.remove(CACHE_KEY);
        _flushToSheet(logs);
    }

    /**
     * Orchestrates a tool execution with automatic logging and error handling.
     * @param {string} toolKey - Key from APP_REGISTRY
     * @param {string} reference - Contextual reference (e.g. 'Manual Run')
     * @param {Function} callback - Function containing the tool logic
     */
    function run(toolKey, reference, callback) {
        setRunId();
        var cfg;
        try {
            cfg = BeaverEngine.getTool(toolKey);
        } catch (e) {
            cfg = { TITLE: toolKey };
        }
        var source = cfg.TITLE;

        info(source, reference, "🚀 Execution started");
        
        try {
            var result = callback();
            success(source, reference, "✅ Execution completed successfully");
            return result;
        } catch (e) {
            error(source, reference, e, { toolConfig: cfg, arguments: arguments });
            // Optionally notify user via Toast if this was a manual trigger
            throw e; 
        } finally {
            flushLogs();
        }
    }

    /**
     * Utility to safely execute an entry-point function and guarantee log flushing.
     */
    function wrap(source, reference, func) {
        return function() {
            setRunId();
            try {
                return func.apply(this, arguments);
            } catch(e) {
                error(source, reference, e, { args: arguments });
                throw e;
            } finally {
                flushLogs();
            }
        };
    }

    return {
        setLoggingState: setLoggingState,
        setRunId: setRunId,
        info: info,
        success: success,
        warn: warn,
        debug: debug,
        error: error,
        step: step,
        flushLogs: flushLogs,
        isEnabled: isEnabled,
        run: run,
        wrap: wrap
    };
})();

