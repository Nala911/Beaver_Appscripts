/**
 * Pipeline
 * Version: 4.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('PIPELINE', {
    SHEET_NAME: SHEET_NAMES.PIPELINE,
    TITLE: SHEET_NAMES.PIPELINE,
    MENU_LABEL: SHEET_NAMES.PIPELINE,
    MENU_ENTRYPOINT: 'PipelineControl_openSidebar',
    MENU_ORDER: 100,
    SIDEBAR_HTML: 'PipelineControl_Sidebar',
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 2,
    COL_WIDTHS: [60, 200, 200, 200, 200, 200, 200, 200],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [
            { type: 'custom', formula: '=$A2=TRUE', color: SHEET_THEME.STATUS.SUCCESS, scope: 'actionOnly', actionCol: 'A' }
        ],
        COL_SCHEMA: [
            { header: 'ON/OFF', type: 'CHECKBOX' },
            { header: 'Pipeline Name', type: 'TEXT' },
            { header: 'Source URL', type: 'URL' },
            { header: 'Source Range', type: 'TEXT' },
            { header: 'Destination URL', type: 'URL' },
            { header: 'Destination Cell', type: 'TEXT' },
            { header: 'Sync Interval', type: 'DROPDOWN', options: ['Manual Only', '15 min', '30 min', '1 hour', '4 hours', '12 hours', '1 day'] },
            { header: 'Last Run Time', type: 'DATETIME' }
        ]
    }
});
/**
 *
 * Column Mappings (0-indexed):
 * 0: ON/OFF
 * 1: Pipeline Name                2: Source URL
 * 3: Source Range                 4: Destination URL
 * 5: Destination Cell             6: Sync Interval
 * 7: Last Run Time
 */

var PIPELINE_NON_DATA_ROWS = 1;

// --- SIDEBAR ---

/** Opens the Pipeline sidebar, creating the sheet if needed. */
function PipelineControl_openSidebar() {
    return Logger.run('PIPELINE', 'Open Sidebar', function () {
        _App_launchTool('PIPELINE');
    });
}

// --- GLOBAL CONTROLS ---

function PipelineControl_getSystemStatus() {
    return Logger.run('PIPELINE', 'Get Status', function () {
        var enabled = _App_getProperty(APP_PROPS.SYSTEM_ENABLED);
        var status = enabled === null ? 'false' : enabled;
        return _App_ok('System status retrieved.', status);
    });
}

function PipelineControl_setSystemStatus(isEnabled) {
    return Logger.run('PIPELINE', 'Set Status', function () {
        _App_setProperty(APP_PROPS.SYSTEM_ENABLED, isEnabled.toString());
        _PipelineControl_manageTrigger(isEnabled);
        return _App_ok('System status updated.', isEnabled);
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

// --- PIPELINE EXECUTION ---

function PipelineControl_processPipelines() {
    return Logger.run('PIPELINE', 'Scheduled Execution', function () {
        return _App_withDocumentLock('PIPELINE_PROCESS', function () {
            if (_App_getProperty(APP_PROPS.SYSTEM_ENABLED) !== 'true') {
                Logger.info('PIPELINE', 'Global', "System is globally disabled. Skipping execution.");
                return;
            }

            var pendingPipelines = SheetManager.readPendingObjects('PIPELINE', { actionColName: 'ON/OFF' });

            var activeScheduled = pendingPipelines.filter(function(p) {
                return _PipelineControl_shouldRun(p);
            });

            if (activeScheduled.length === 0) {
                Logger.info('PIPELINE', 'Global', "No pipelines scheduled to run.");
                return;
            }

            var sheet = SheetManager.getSheet('PIPELINE');

            _App_BatchProcessor('PIPELINE', activeScheduled, function (item) {
                _PipelineControl_runPipeline(sheet, item._rowNumber, item);
                return { success: true };
            });
        });
    });
}

function PipelineControl_runAllPipelines() {
    return Logger.run('PIPELINE', 'Run All', function () {
        return _App_withDocumentLock('PIPELINE_RUN_ALL', function () {
            var pendingPipelines = SheetManager.readPendingObjects('PIPELINE', { actionColName: 'ON/OFF' });

            if (pendingPipelines.length === 0) return _App_ok('No enabled pipelines to run.');

            var sheet = SheetManager.getSheet('PIPELINE');

            var stats = _App_BatchProcessor('PIPELINE', pendingPipelines, function (item) {
                _PipelineControl_runPipeline(sheet, item._rowNumber, item);
                return { success: true };
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

function PipelineControl_getPipelineDashboardData() {
    return Logger.run('PIPELINE', 'Dashboard Data', function () {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PIPELINE);
        if (!sheet) return null;

        var dataObjects = SheetManager.readObjects('PIPELINE');

        var summary = {
            total: 0,
            active: 0,
            success: 0,
            failed: 0
        };
        var pipelines = [];

        dataObjects.forEach(function (obj, idx) {
            var statusVal = obj['ON/OFF'];
            var isEnabled = (String(statusVal).toLowerCase() === 'enabled') || (statusVal === true);
            var name = obj['Pipeline Name'];
            var lastRun = obj['Last Run Time'];

            if (name) { 
                summary.total++;
                if (isEnabled) summary.active++;

                var formattedDate = "";
                if (lastRun && lastRun instanceof Date) {
                    var opts = { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
                    formattedDate = lastRun.toLocaleDateString(undefined, opts);
                } else if (lastRun) {
                    formattedDate = String(lastRun);
                }

                pipelines.push({
                    rowIndex: idx + 2,
                    name: name,
                    isEnabled: isEnabled,
                    lastRun: formattedDate,
                    lastStatus: "Check Logs" 
                });
            }
        });

        return {
            summary: summary,
            pipelines: pipelines
        };
    });
}

function PipelineControl_runSelectedPipelines(rowIndexes) {
    return Logger.run('PIPELINE', 'Run Selected', function () {
        return _App_withDocumentLock('PIPELINE_RUN_SELECTED', function () {
            var sheet = SheetManager.getSheet('PIPELINE');
            var headers = SheetManager.getHeaders('PIPELINE');
            
            var items = rowIndexes.map(function(idx) {
                return { rowIndex: idx };
            });

            var stats = _App_BatchProcessor('PIPELINE', items, function (item) {
                var range = sheet.getRange(item.rowIndex, 1, 1, headers.length);
                var rowArray = range.getValues()[0];
                var rowObj = {};
                for(var i=0; i<headers.length; i++) {
                    rowObj[headers[i]] = rowArray[i];
                }
                _PipelineControl_runPipeline(sheet, item.rowIndex, rowObj);

                var updatedLastRun = sheet.getRange(item.rowIndex, 8).getValue();
                return {
                    rowIndex: item.rowIndex,
                    lastStatus: "Check Logs",
                    lastRun: updatedLastRun ? updatedLastRun.toString() : ""
                };
            });

            return _App_ok('Selected pipelines completed', stats.results);
        });
    });
}

function _PipelineControl_runPipeline(sheet, rowIdx, rowObj) {
    var logMessage = "";
    var isSuccess = false;
    var errorObj = null;
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

    try {
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
            
            logMessage = "Synced " + numRows + " rows.";
            isSuccess = true;
        } else {
            logMessage = "No data found in source range.";
        }

    } catch (e) {
        logMessage = "Error: " + e.message;
        errorObj = e;
        isSuccess = false;
    }

    var timestamp = new Date();
    sheet.getRange(rowIdx, 8).setValue(timestamp);

    var reference = pipelineName ? pipelineName : ('Row ' + rowIdx);
    if (isSuccess) {
        Logger.info(SyncEngine.getTool('PIPELINE').TITLE, reference, logMessage);
    } else {
        Logger.error(SyncEngine.getTool('PIPELINE').TITLE, reference, errorObj || logMessage);
    }
}

function PipelineControl_formatControlCenter() {
    return Logger.run('PIPELINE', 'Format Center', function () {
        var sheet = _App_ensureSheetExists('PIPELINE');

        // Re-apply standard setup from Registry.
        // Note: _App_applyBodyFormatting already replaces all conditional format rules and
        // data validations via setConditionalFormatRules, so no pre-clear is needed.
        var cfg = SyncEngine.getTool('PIPELINE');
        _App_applyBodyFormatting(sheet, 0, cfg.FORMAT_CONFIG);
        // Schema-driven validation now handles this within _App_applyBodyFormatting

        // Column group separators (Pipeline-specific visual grouping)
        sheet.getRange("C:C").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
        sheet.getRange("E:E").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
        sheet.getRange("G:G").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
        sheet.getRange("H:H").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);

        return "Formatted Pipeline with Dark Theme & Elegant Groups!";
    });
}
