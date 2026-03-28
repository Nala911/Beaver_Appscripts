// ==========================================
// Client-side wrappers
// ==========================================

function Logger_getLoggingState() { return Logger.isEnabled(); }

/**
 * Activates (navigates to) the Developer Log sheet.
 * Called by the sidebar's "View Logs Sheet" button.
 */
function Logger_activateSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
    if (sheet) {
        sheet.activate();
    } else {
        throw new Error("Developer Log sheet not found. Enable logging first.");
    }
}

function Logger_getSettings() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
    var rowCount = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

    return {
        enabled: Logger.isEnabled(),
        currentRows: rowCount,
        theme: SHEET_THEME 
    };
}

function Logger_saveSettings(enabled) {
    try {
        Logger.setLoggingState(enabled);
        return _App_ok(enabled ? 'Logging enabled.' : 'Logging disabled.');
    } catch (e) {
        return _App_fail('Error: ' + e.message);
    }
}

function Logger_clearLogs() {
    try {
        return _App_withDocumentLock('LOGGER_CLEAR', function () {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
            if (sheet && sheet.getLastRow() > 1) {
                sheet.deleteRows(2, sheet.getLastRow() - 1);
                if (BeaverEngine.getTool('LOGS')) _App_applyBodyFormatting(sheet, 0, BeaverEngine.getTool('LOGS').FORMAT_CONFIG);
            }
            return _App_ok('Logs cleared successfully.');
        });
    } catch (e) {
        return _App_fail('Clear failed: ' + e.message);
    }
}


/**
 * Retrieves a summary of the most recent execution (grouped by Run ID).
 * Used by the sidebar to show "Last Run" status.
 */
function Logger_getLastRunSummary() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
        if (!sheet || sheet.getLastRow() < 2) return null;

        // Logs are now appended to bottom, so search bottom-up
        var lastRow = sheet.getLastRow();
        var startRow = Math.max(2, lastRow - 99);
        var numRows = lastRow - startRow + 1;
        var numCols = BeaverEngine.getTool('LOGS').HEADERS.length;
        var data = sheet.getRange(startRow, 1, numRows, numCols).getValues();

        var lastRunId = data[data.length - 1][1]; // Col B (Run ID) of bottom row
        if (!lastRunId || lastRunId === 'N/A') {
            return {
                timestamp: data[data.length - 1][0],
                source: data[data.length - 1][3],
                level: data[data.length - 1][2],
                message: data[data.length - 1][5]
            };
        }

        var summary = {
            runId: lastRunId,
            timestamp: data[data.length - 1][0],
            sources: {},
            totalErrors: 0,
            totalInfos: 0,
            totalSuccess: 0,
            totalWarnings: 0,
            recentErrors: []
        };

        // Traverse backwards to collect logs for the last run
        for (var i = data.length - 1; i >= 0; i--) {
            if (data[i][1] !== lastRunId) break; 

            var level = data[i][2];
            var src = data[i][3];
            var msg = data[i][5];

            if (!summary.sources[src]) summary.sources[src] = { ok: 0, err: 0, warn: 0 };

            if (level === 'ERROR') {
                summary.sources[src].err++;
                summary.totalErrors++;
                if (summary.recentErrors.length < 5) summary.recentErrors.push({ source: src, msg: msg });
            } else if (level === 'SUCCESS') {
                summary.sources[src].ok++;
                summary.totalSuccess++;
            } else if (level === 'WARN') {
                summary.sources[src].warn++;
                summary.totalWarnings++;
            } else {
                summary.sources[src].ok++;
                summary.totalInfos++;
            }
        }

        return summary;
    } catch (e) {
        return { error: e.message };
    }
}

/**
 * Logs a client-side error from a sidebar.
 * @param {string|Object} err Error message or object
 * @param {string} [context] Optional context (e.g. 'Button Click')
 */
function Logger_logClientError(err, context) {
    var source = 'Client Sidebar';
    var ref = context || 'UI';
    var msg = (typeof err === 'object' && err.message) ? err.message : String(err);
    if (typeof err === 'object' && err.stack) msg += '\nStack: ' + err.stack;
    
    Logger.error(source, ref, msg);
    // Removed immediate Logger.flushLogs() to fix performance bottleneck
}

/**
 * Logs a client-side info message from a sidebar.
 * @param {string} message 
 * @param {string} [context]
 */
function Logger_logClientInfo(message, context) {
    Logger.info('Client Sidebar', context || 'UI', message);
    // Removed immediate Logger.flushLogs() to fix performance bottleneck
}
