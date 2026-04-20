/**
 * UNIFIED LOGGING SYSTEM
 * ==========================================
 * Provides buffered, concurrent-safe logging with document-level locking.
 */

Object.assign(App.Log, (function () {
    var CACHE_KEY = App.Config.CACHE_KEYS.LOGS;
    var MAX_CACHE_ITEMS = 25;
    
    var _state = {
        currentRunId: null,
        currentSteps: [],
        depth: 0
    };

    var _transporter = {
        queue: function (entry) {
            var lock = LockService.getDocumentLock();
            if (!lock.tryLock(10000)) return;

            try {
                var cache = CacheService.getDocumentCache();
                var existing = cache.get(CACHE_KEY);
                var logs = existing ? JSON.parse(existing) : [];
                logs.push(entry);

                var serialized = JSON.stringify(logs);
                if (logs.length >= MAX_CACHE_ITEMS || serialized.length > 90000) {
                    this.flush(logs, true);
                    cache.remove(CACHE_KEY); 
                } else {
                    cache.put(CACHE_KEY, serialized, 21600);
                }
            } finally {
                lock.releaseLock();
            }
        },

        flush: function (manualLogs, skipLock) {
            var cache = CacheService.getDocumentCache();
            var logs = manualLogs;
            if (!logs) {
                var existing = cache.get(CACHE_KEY);
                if (!existing) return;
                try { logs = JSON.parse(existing); } catch (e) { return; }
            }
            if (!logs || logs.length === 0) return;

            var lock = null;
            if (!skipLock) {
                lock = LockService.getDocumentLock();
                if (!lock.tryLock(15000)) return; 
            }

            try {
                var ss = SpreadsheetApp.getActiveSpreadsheet();
                var logSheetName = App.Config.SHEET_NAMES.LOGS;
                var sheet = ss.getSheetByName(logSheetName) || ss.insertSheet(logSheetName);
                
                if (sheet.getLastRow() === 0) {
                    var headers = ['Timestamp', 'Run ID', 'Level', 'Source', 'Reference', 'Message', 'User'];
                    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
                    sheet.setFrozenRows(1);
                }

                sheet.getRange(sheet.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);

                var systemPrefs = App.Engine.getPrefs('SYSTEM');
                var maxRows = parseInt(systemPrefs.loggerMaxRows || "0", 10);
                if (maxRows > 0 && (sheet.getLastRow() - 1) > maxRows) {
                    sheet.deleteRows(2, (sheet.getLastRow() - 1) - maxRows);
                }

                if (!manualLogs) cache.remove(CACHE_KEY);
            } catch (err) {
                // Silent catch to prevent infinite error loops during logging
                console.error("Logger Flush Error: " + err.message);
            } finally {
                if (lock) lock.releaseLock();
            }
        },

        clear: function () {
            var lock = LockService.getDocumentLock();
            if (lock.tryLock(15000)) {
                try {
                    CacheService.getDocumentCache().remove(CACHE_KEY);
                    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(App.Config.SHEET_NAMES.LOGS);
                    if (sheet && sheet.getLastRow() > 1) {
                        sheet.deleteRows(2, sheet.getLastRow() - 1);
                        var logConfig = App.Engine.getTool('LOGS');
                        if (logConfig) _App_applyBodyFormatting(sheet, 0, logConfig.FORMAT_CONFIG);
                    }
                } finally {
                    lock.releaseLock();
                }
            }
        }
    };

    function _safeStringify(obj) {
        if (obj === undefined) return "undefined";
        if (obj === null) return "null";
        if (typeof obj !== 'object') return String(obj);
        const cache = new Set();
        return JSON.stringify(obj, (key, value) => {
            if (typeof value === 'object' && value !== null) {
                if (cache.has(value)) return '[Circular]';
                cache.add(value);
            }
            if (value instanceof Error) return { name: value.name, message: value.message, stack: value.stack };
            return value;
        }, 2);
    }

    function isEnabled() {
        return App.Engine.getPrefs('SYSTEM').enableDebugLogging === true;
    }

    function _log(level, source, reference, message, ctx) {
        if (!isEnabled()) return;
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        var user = Session.getActiveUser().getEmail() || 'Unknown';
        var runId = _state.currentRunId || 'N/A';
        var msgStr = (typeof message === 'object' ? _safeStringify(message) : String(message)) + (ctx ? "\n\n--- Context ---\n" + _safeStringify(ctx) : "");
        _transporter.queue([timestamp, runId, level, source, reference || '', msgStr, user]);
    }

    return {
        setRunId: function (id) {
            _state.currentRunId = id || Utilities.getUuid();
            _state.currentSteps = [];
            return _state.currentRunId;
        },
        info: function (src, ref, msg, ctx) { _log('INFO', src, ref, msg, ctx); },
        success: function (src, ref, msg, ctx) { _log('SUCCESS', src, ref, msg, ctx); },
        warn: function (src, ref, msg, ctx) { _log('WARN', src, ref, msg, ctx); },
        error: function (src, ref, err, ctx) {
            var msg = typeof err === 'string' ? err : (err.message + (err.stack ? '\n' + err.stack : ''));
            var fullCtx = Object.assign({ breadcrumbs: _state.currentSteps.slice() }, ctx || {});
            _log('ERROR', src, ref, msg, fullCtx);
        },
        step: function (src, ref, name) {
            _state.currentSteps.push(`[${new Date().toISOString()}] ${src}: ${name}`);
            _log('DEBUG', src, ref, "STEP: " + name);
        },
        flush: function () { _transporter.flush(); },
        clear: function () { _transporter.clear(); },
        isEnabled: isEnabled,
        setLoggingState: function(enabled) {
            var prefs = App.Engine.getPrefs('SYSTEM');
            prefs.enableDebugLogging = !!enabled;
            App.Engine.setPrefs('SYSTEM', prefs);
        },
        flushLogs: function() { this.flush(); },
        clearLogs: function() { _transporter.clear(); },
        run: function (toolKey, reference, callback) {
            _state.depth++;
            this.setRunId();
            var source = toolKey;
            try { source = App.Engine.getTool(toolKey).TITLE; } catch (e) {}
            this.info(source, reference, "🚀 Execution started");
            try {
                var result = callback();
                this.success(source, reference, "✅ Execution completed");
                return result;
            } catch (e) {
                this.error(source, reference, e);
                throw e; 
            } finally {
                this.flush();
                _state.depth--;
                if (_state.depth <= 0) { _state.currentRunId = null; _state.currentSteps = []; _state.depth = 0; }
            }
        }
    };
})());

// Backward Compatibility Layer
var Logger = App.Log;

// Registration with Engine
App.Engine.registerTool('LOGS', {
    SHEET_NAME: App.Config.SHEET_NAMES.LOGS,
    TITLE: '🛠️ Developer Log',
    MENU_LABEL: '🛠️ Developer Log',
    MENU_ENTRYPOINT: 'Logger_showSidebar',
    MENU_ORDER: 110,
    SIDEBAR_HTML: 'Logger_Sidebar',
    SIDEBAR_WIDTH: 360,
    FROZEN_ROWS: 1,
    FORMAT_CONFIG: {
        skipActionColoring: true,
        conditionalRules: [
            { type: 'custom', formula: '=OR($C2="ERROR", $C2="FAIL")', color: '#F4CCCC', scope: 'fullRow' },
            { type: 'custom', formula: '=$C2="SUCCESS"', color: '#D9EAD3', scope: 'fullRow' }
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
