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
    SIDEBAR_HTML: 'tools/PipelineControl_Sidebar',
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
                    row: p._rowNumber,
                    name: p['Pipeline Name'] || ('Row ' + p._rowNumber),
                    status: p['Status'] || '',
                    lastRun: p['Last Run Time'] || '',
                    enabled: isEnabled
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
