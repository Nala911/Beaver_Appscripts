/**
 * Developer Logging System
 * Version: 6.0 (Modular Transporter Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('LOGS', {
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
    var CACHE_KEY = typeof CACHE_KEYS !== 'undefined' ? CACHE_KEYS.LOGS : 'WorkspaceSync_DEBUG_LOGS';
    var MAX_CACHE_ITEMS = 25;
    
    var _state = {
        currentRunId: null,
        currentSteps: [],
        depth: 0
    };

    /**
     * Internal Transporter — Abstracts the storage and persistence logic.
     */
    var _transporter = {
        /**
         * Queues a log entry to the document cache.
         * Uses a lock to prevent race conditions during cache read/write.
         */
        queue: function (entry) {
            var lock = LockService.getDocumentLock();
            if (!lock.tryLock(10000)) {
                console.warn("Logger: Queue lock timeout. Log may be lost.");
                return;
            }

            try {
                var cache = CacheService.getDocumentCache();
                var existing = cache.get(CACHE_KEY);
                var logs = existing ? JSON.parse(existing) : [];

                logs.push(entry);

                var serialized = JSON.stringify(logs);
                // Flush if batch is full OR if string size approaches 100KB CacheService limit
                if (logs.length >= MAX_CACHE_ITEMS || serialized.length > 90000) {
                    this.flush(logs, true); // skipLock = true because we already hold the lock
                    cache.remove(CACHE_KEY); 
                } else {
                    cache.put(CACHE_KEY, serialized, 21600); // 6h TTL
                }
            } catch (e) {
                console.error("Logger Queue Error: " + e.message);
            } finally {
                lock.releaseLock();
            }
        },

        /**
         * Flushes logs to the spreadsheet with document-level locking.
         * @param {Array} [manualLogs] Optional logs to flush directly
         * @param {boolean} [skipLock] If true, bypasses lock acquisition (used when caller already holds it)
         */
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
                if (!lock.tryLock(15000)) {
                    console.warn("Logger: Transport lock timeout. Logs retained in cache.");
                    return; 
                }
            }

            try {
                var ss = SpreadsheetApp.getActiveSpreadsheet();
                var sheet = ss.getSheetByName(SHEET_NAMES.LOGS) || _App_ensureSheetExists('LOGS');
                
                var startRow = sheet.getLastRow() + 1;
                sheet.getRange(startRow, 1, logs.length, logs[0].length).setValues(logs);

                // Auto-Pruning (keep sheet size manageable)
                var maxRows = parseInt(_App_getProperty(APP_PROPS.LOGGER_MAX_ROWS) || "0", 10);
                if (maxRows > 0) {
                    var currentDataRows = sheet.getLastRow() - 1;
                    if (currentDataRows > maxRows) {
                        sheet.deleteRows(2, currentDataRows - maxRows);
                    }
                }

                // Batch Formatting
                var logConfig = SyncEngine.getTool('LOGS');
                if (logConfig && logConfig.FORMAT_CONFIG) {
                    _App_applyBodyFormatting(sheet, sheet.getLastRow() - 1, logConfig.FORMAT_CONFIG);
                }

                if (!manualLogs) cache.remove(CACHE_KEY);
            } catch (err) {
                console.error("Logger Transporter Error: " + err.message);
            } finally {
                if (lock) lock.releaseLock();
            }
        },

        /**
         * Clears all logs from the sheet and cache.
         */
        clear: function () {
            var lock = LockService.getDocumentLock();
            if (lock.tryLock(15000)) {
                try {
                    CacheService.getDocumentCache().remove(CACHE_KEY);
                    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
                    if (sheet && sheet.getLastRow() > 1) {
                        sheet.deleteRows(2, sheet.getLastRow() - 1);
                        var logConfig = SyncEngine.getTool('LOGS');
                        if (logConfig) _App_applyBodyFormatting(sheet, 0, logConfig.FORMAT_CONFIG);
                    }
                } finally {
                    lock.releaseLock();
                }
            }
        }
    };

    /**
     * Safely stringifies complex or circular objects for deep context capture.
     */
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
            if (value instanceof Error) {
                 return { name: value.name, message: value.message, stack: value.stack };
            }
            return value;
        }, 2);
    }

    function isEnabled() {
        return _App_getProperty(APP_PROPS.ENABLE_DEBUG_LOGGING) === 'true';
    }

    /**
     * Internal factory for log entries.
     */
    function _log(level, source, reference, message, ctx) {
        if (!isEnabled()) return;

        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        var user = Session.getActiveUser().getEmail() || 'Unknown';
        var runId = _state.currentRunId || 'N/A';
        
        var contextStr = ctx ? "\n\n--- Deep Context ---\n" + _safeStringify(ctx) : "";
        var msgStr = (typeof message === 'object' ? _safeStringify(message) : String(message)) + contextStr;
        
        _transporter.queue([timestamp, runId, level, source, reference || '', msgStr, user]);
    }

    return {
        /**
         * Global toggle for the logging system.
         */
        setLoggingState: function (enabled) {
            _App_setProperty(APP_PROPS.ENABLE_DEBUG_LOGGING, enabled ? 'true' : 'false');
            if (enabled) {
                var sheet = _App_ensureSheetExists('LOGS');
                if (sheet.isSheetHidden()) sheet.showSheet();
            } else {
                _transporter.clear();
            }
        },

        setRunId: function (id) {
            if (id && id !== _state.currentRunId) {
                _state.currentRunId = id;
                _state.currentSteps = []; // Reset steps if ID is explicitly changed
            } else if (!_state.currentRunId) {
                _state.currentRunId = Utilities.getUuid();
                _state.currentSteps = [];
            }
            return _state.currentRunId;
        },

        info: function (src, ref, msg, ctx) { _log('INFO', src, ref, msg, ctx); },
        success: function (src, ref, msg, ctx) { _log('SUCCESS', src, ref, msg, ctx); },
        warn: function (src, ref, msg, ctx) { _log('WARN', src, ref, msg, ctx); },
        debug: function (src, ref, msg, ctx) { _log('DEBUG', src, ref, msg, ctx); },
        
        error: function (src, ref, err, ctx) {
            var msg = typeof err === 'string' ? err : (err.message + (err.stack ? '\n' + err.stack : ''));
            var fullCtx = Object.assign({}, ctx || {});
            
            // Auto-extract deep context explicitly attached to error objects
            if (typeof err === 'object' && err !== null && err.localContext) {
                fullCtx.errorDetails = err.localContext;
            }
            if (_state.currentSteps.length > 0) {
                fullCtx.executionBreadcrumbs = _state.currentSteps.slice();
            }
            
            _log('ERROR', src, ref, msg, fullCtx);
        },

        /**
         * Records a breadcrumb step for the current execution.
         */
        step: function (src, ref, name) {
            _state.currentSteps.push(`[${new Date().toISOString()}] ${src} - ${name}`);
            _log('DEBUG', src, ref, "STEP: " + name);
        },

        flushLogs: function () { _transporter.flush(); },
        clearLogs: function () { _transporter.clear(); },
        isEnabled: isEnabled,

        /**
         * High-level orchestrator for tool execution.
         */
        run: function (toolKey, reference, callback) {
            _state.depth++;
            this.setRunId();
            var cfg = { TITLE: toolKey };
            try { cfg = SyncEngine.getTool(toolKey); } catch (e) {}
            var source = cfg.TITLE;

            this.info(source, reference, "🚀 Execution started");
            try {
                var result = callback();
                this.success(source, reference, "✅ Execution completed successfully");
                return result;
            } catch (e) {
                this.error(source, reference, e, { toolConfig: cfg });
                throw e; 
            } finally {
                this.flushLogs();
                _state.depth--;
                if (_state.depth <= 0) {
                    _state.currentRunId = null;
                    _state.currentSteps = [];
                    _state.depth = 0;
                }
            }
        },

        /**
         * Wraps an entry-point function with automatic logging and error handling.
         */
        wrap: function (source, reference, func) {
            var self = this;
            return function() {
                _state.depth++;
                self.setRunId();
                try {
                    return func.apply(this, arguments);
                } catch(e) {
                    self.error(source, reference, e, { args: arguments });
                    throw e;
                } finally {
                    self.flushLogs();
                    _state.depth--;
                    if (_state.depth <= 0) {
                        _state.currentRunId = null;
                        _state.currentSteps = [];
                        _state.depth = 0;
                    }
                }
            };
        }
    };
})();
