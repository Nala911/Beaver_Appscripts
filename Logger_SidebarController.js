// ==========================================
// Client-side wrappers
// ==========================================

function Logger_getLoggingState() { return Logger.isEnabled(); }

/**
 * Activates (navigates to) the Developer Log sheet.
 * Called by the sidebar's "View Logs Sheet" button.
 */
function Logger_activateSheet() {
    return Logger.run('LOGS', 'Activate Sheet', function () {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
        if (sheet) {
            sheet.activate();
        } else {
            throw new Error("Developer Log sheet not found. Enable logging first.");
        }
    });
}

function Logger_getSettings() {
    return Logger.run('LOGS', 'Get Settings', function () {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
        var rowCount = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

        return {
            enabled: Logger.isEnabled(),
            currentRows: rowCount,
            theme: SHEET_THEME 
        };
    });
}

function Logger_saveSettings(enabled) {
    return Logger.run('LOGS', 'Save Settings', function () {
        try {
            Logger.setLoggingState(enabled);
            return _App_ok(enabled ? 'Logging enabled.' : 'Logging disabled.');
        } catch (e) {
            return _App_fail('Error: ' + e.message);
        }
    });
}

function Logger_clearLogs() {
    return Logger.run('LOGS', 'Clear Logs', function () {
        try {
            Logger.clearLogs();
            return _App_ok('Logs cleared successfully.');
        } catch (e) {
            return _App_fail('Clear failed: ' + e.message);
        }
    });
}


/**
 * Retrieves a summary of the most recent execution (grouped by Run ID).
 * Used by the sidebar to show "Last Run" status.
 */
function Logger_getLastRunSummary() {
    return Logger.run('LOGS', 'Last Run Summary', function () {
        try {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
            if (!sheet || sheet.getLastRow() < 2) return null;

            // Logs are now appended to bottom, so search bottom-up
            var lastRow = sheet.getLastRow();
            var startRow = Math.max(2, lastRow - 99);
            var numRows = lastRow - startRow + 1;
            var numCols = SyncEngine.getTool('LOGS').HEADERS.length;
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
    });
}
