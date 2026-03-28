/**
 * Pipeline Control Center
 * Version: 4.0 (Plugin Architecture — registers with BeaverEngine)
 */

BeaverEngine.registerTool('PIPELINE', {
    SHEET_NAME: SHEET_NAMES.PIPELINE,
    TITLE: '⛓ Pipeline Control Center',
    MENU_LABEL: '⛓  Control Center',
    MENU_ENTRYPOINT: 'Pipeline_showSidebar',
    MENU_ORDER: 100,
    SIDEBAR_HTML: 'PipelineSidebar',
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 2,
    COL_WIDTHS: [60, 200, 200, 200, 200, 200, 200, 200, 200, 200],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [
            { type: 'custom', formula: '=$A2=TRUE', color: SHEET_THEME.STATUS.SUCCESS, scope: 'actionOnly', actionCol: 'A' },
            { type: 'custom', formula: '=$A2=FALSE', color: SHEET_THEME.STATUS.ERROR, scope: 'actionOnly', actionCol: 'A' }
        ],
        COL_SCHEMA: [
            { header: 'ON/OFF', type: 'CHECKBOX' },
            { header: 'Pipeline Name', type: 'TEXT' },
            { header: 'Source URL', type: 'URL' },
            { header: 'Source Sheet Name', type: 'TEXT' },
            { header: 'Source Range', type: 'TEXT' },
            { header: 'Destination URL', type: 'URL' },
            { header: 'Destination Sheet Name', type: 'TEXT' },
            { header: 'Destination Cell', type: 'TEXT' },
            { header: 'Sync Interval', type: 'TEXT' },
            { header: 'Last Run Time', type: 'DATETIME' }
        ]
    }
});
/**
 *
 * Column Mappings (0-indexed):
 * 0: ON/OFF
 * 1: Pipeline Name                2: Source URL
 * 3: Source Sheet Name            4: Source Range
 * 5: Destination URL              6: Destination Sheet Name
 * 7: Destination Cell             8: Sync Interval
 * 9: Last Run Time
 */

var PIPELINE_NON_DATA_ROWS = 1;

// --- SIDEBAR ---

/** @deprecated — Use _App_ensureSheetExists('PIPELINE') instead. */
function _Pipeline_ensureSheetExistsAndActivate() {
    return _App_ensureSheetExists('PIPELINE');
}

/** Opens the Pipeline sidebar, creating the sheet if needed. */
function Pipeline_showSidebar() {
    _App_launchTool('PIPELINE');
}

// --- GLOBAL CONTROLS ---

function Pipeline_getSystemStatus() {
    var enabled = _App_getProperty(APP_PROPS.SYSTEM_ENABLED);
    return enabled === null ? 'true' : enabled;
}

function Pipeline_setSystemStatus(isEnabled) {
    _App_setProperty(APP_PROPS.SYSTEM_ENABLED, isEnabled.toString());
    return isEnabled;
}

// --- PIPELINE EXECUTION ---

function Pipeline_processPipelines() {
    return Logger.run('PIPELINE', 'Scheduled Execution', function () {
        return _App_withDocumentLock('PIPELINE_PROCESS', function () {
            if (Pipeline_getSystemStatus() !== 'true') {
                Logger.info('PIPELINE', 'Global', "System is globally disabled. Skipping execution.");
                return;
            }

            var sheet = SheetManager.getSheet('PIPELINE');
            var data = sheet.getDataRange().getValues();

            for (var i = PIPELINE_NON_DATA_ROWS; i < data.length; i++) {
                var row = data[i];
                var statusVal = row[0];
                var isEnabled = (String(statusVal).toLowerCase() === 'enabled') || (statusVal === true);

                if (!isEnabled) continue;

                if (_Pipeline_shouldRun(row)) {
                    _Pipeline_runPipeline(sheet, i + 1, row);
                }
            }
        });
    });
}

function Pipeline_runAllPipelines() {
    return Logger.run('PIPELINE', 'Run All', function () {
        return _App_withDocumentLock('PIPELINE_RUN_ALL', function () {
            var sheet = SheetManager.getSheet('PIPELINE');
            var data = sheet.getDataRange().getValues();

            for (var i = PIPELINE_NON_DATA_ROWS; i < data.length; i++) {
                var row = data[i];
                var statusVal = row[0];
                var isEnabled = (String(statusVal).toLowerCase() === 'enabled') || (statusVal === true);

                if (isEnabled) {
                    _Pipeline_runPipeline(sheet, i + 1, row);
                }
            }
            return _App_ok('Execution complete');
        });
    });
}

function _Pipeline_shouldRun(row) {
    var intervalStr = String(row[8]);
    var lastRun = row[9];

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

function Pipeline_getPipelineDashboardData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PIPELINE);
    if (!sheet) return null;

    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

    var summary = {
        total: 0,
        active: 0,
        success: 0,
        failed: 0
    };
    var pipelines = [];

    for (var i = PIPELINE_NON_DATA_ROWS; i < data.length; i++) {
        var row = data[i];
        var statusVal = row[0];
        var isEnabled = (String(statusVal).toLowerCase() === 'enabled') || (statusVal === true);
        var name = row[1];
        var lastRun = row[9];

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
                rowIndex: i + 1,
                name: name,
                isEnabled: isEnabled,
                lastRun: formattedDate,
                lastStatus: "Check Logs" 
            });
        }
    }

    return {
        summary: summary,
        pipelines: pipelines
    };
}

function Pipeline_runSelectedPipelines(rowIndexes) {
    return Logger.run('PIPELINE', 'Run Selected', function () {
        return _App_withDocumentLock('PIPELINE_RUN_SELECTED', function () {
            var sheet = SheetManager.getSheet('PIPELINE');
            var colCount = BeaverEngine.getTool('PIPELINE').HEADERS.length;
            var results = [];

            for (var i = 0; i < rowIndexes.length; i++) {
                var rowIdx = rowIndexes[i];
                var data = sheet.getRange(rowIdx, 1, 1, colCount).getValues()[0];
                _Pipeline_runPipeline(sheet, rowIdx, data);

                var updatedData = sheet.getRange(rowIdx, 1, 1, colCount).getValues()[0];
                results.push({
                    rowIndex: rowIdx,
                    lastStatus: "Check Logs",
                    lastRun: updatedData[9] ? updatedData[9].toString() : ""
                });
            }
            return _App_ok('Selected pipelines completed', results);
        });
    });
}

function _Pipeline_runPipeline(sheet, rowIdx, rowData) {
    var logMessage = "";
    var isSuccess = false;
    var pipelineName = rowData[1];

    try {
        var sourceUrl = rowData[2];
        var sourceSheetName = rowData[3];
        var sourceRangeA1 = rowData[4];
        var destUrl = rowData[5];
        var destSheetName = rowData[6];
        var destStartCell = rowData[7];

        if (!sourceUrl || !sourceSheetName || !sourceRangeA1 || !destUrl || !destSheetName || !destStartCell) {
            throw new Error("Missing required config (URL, Sheet Name, Range, or Destination)");
        }

        var sSs;
        try {
            sSs = SpreadsheetApp.openByUrl(sourceUrl);
        } catch (e) {
            throw new Error("Cannot access Source URL (Check permissions or URL validity)");
        }
        var sSheet = sSs.getSheetByName(sourceSheetName);
        if (!sSheet) throw new Error("Source sheet '" + sourceSheetName + "' not found");

        var values = sSheet.getRange(sourceRangeA1).getValues();
        if (values.length === 0) throw new Error("Source range empty");

        var dSs;
        try {
            dSs = SpreadsheetApp.openByUrl(destUrl);
        } catch (e) {
            throw new Error("Cannot access Destination URL (Check permissions or URL validity)");
        }
        var dSheet = dSs.getSheetByName(destSheetName);
        if (!dSheet) throw new Error("Destination sheet '" + destSheetName + "' not found");

        var numRows = values.length;
        var numCols = values[0].length;

        if (numRows > 0 && numCols > 0) {
            var destRange = dSheet.getRange(destStartCell);
            var startRow = destRange.getRow();
            var startCol = destRange.getColumn();

            dSheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
            logMessage = "Synced " + numRows + " rows.";
            isSuccess = true;
        } else {
            logMessage = "No data found in source range.";
        }

    } catch (e) {
        logMessage = "Error: " + e.message;
        isSuccess = false;
    }

    var timestamp = new Date();
    sheet.getRange(rowIdx, 10).setValue(timestamp);

    var reference = pipelineName ? pipelineName : ('Row ' + rowIdx);
    if (isSuccess) {
        Logger.info(BeaverEngine.getTool('PIPELINE').TITLE, reference, logMessage);
    } else {
        Logger.error(BeaverEngine.getTool('PIPELINE').TITLE, reference, logMessage);
    }
}

function Pipeline_formatControlCenter() {
    var sheet = _App_ensureSheetExists('PIPELINE');

    // Re-apply standard setup from Registry.
    // Note: _App_applyBodyFormatting already replaces all conditional format rules and
    // data validations via setConditionalFormatRules, so no pre-clear is needed.
    var cfg = BeaverEngine.getTool('PIPELINE');
    _App_applyBodyFormatting(sheet, 0, cfg.FORMAT_CONFIG);
    // Schema-driven validation now handles this within _App_applyBodyFormatting

    // Column group separators (Pipeline-specific visual grouping)
    sheet.getRange("C:C").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
    sheet.getRange("F:F").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
    sheet.getRange("I:I").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);
    sheet.getRange("J:J").setBorder(null, true, null, null, null, null, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);

    return "Formatted Control Center with Dark Theme & Elegant Groups!";
}


