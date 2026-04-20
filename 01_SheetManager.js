// ==========================================
// Centralized Sheet Manager (DAO Pattern)
// ==========================================

var SheetManager = (function() {

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

    function getHeaders(toolKey) {
        return SyncEngine.getTool(toolKey).HEADERS || [];
    }

    function getHeaderMap(toolKey) {
        var headers = getHeaders(toolKey);
        var map = {};
        headers.forEach(function(header, index) {
            map[header] = index + 1;
        });
        return map;
    }

    function getNormalizedHeaderMap(toolKey) {
        var headers = getHeaders(toolKey);
        var map = {};
        headers.forEach(function(header, index) {
            map[_normalizeHeaderKey(header)] = index;
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
     * using the headers defined in the tool configuration.
     * @returns {Object[]} Array of row objects mapped by header names
     */
    function readObjects(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        var cfg = SyncEngine.getTool(toolKey);
        var headers = cfg.HEADERS;
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
     * Reads only rows with a pending 'Action' in Column A.
     * Memory-efficient alternative to readObjects for large sheets.
     * @returns {Object[]} Array of row objects with an additional '_rowNumber' property.
     */
    function readPendingActions(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        var cfg = SyncEngine.getTool(toolKey);
        var headers = cfg.HEADERS;
        
        // Step 1: Read ONLY the Action column (Column A)
        var actionValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var rowsToRead = [];
        for (var i = 0; i < actionValues.length; i++) {
            if (actionValues[i][0] && actionValues[i][0] !== "") {
                rowsToRead.push(i + 2); // 1-based index, row 2 is index 0
            }
        }
        
        if (rowsToRead.length === 0) return [];

        // Step 2: For efficiency, read full range and filter in memory if row count is small,
        // or fetch specific rows if the list is sparse. 
        // For Google Sheets, reading a single block is usually faster than many getRange calls.
        var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
        var allData = dataRange.getValues();
        
        var filteredData = [];
        for (var j = 0; j < rowsToRead.length; j++) {
            var rowIdx = rowsToRead[j] - 2;
            var row = allData[rowIdx];
            var obj = { _rowNumber: rowsToRead[j] };
            for (var k = 0; k < headers.length; k++) {
                obj[headers[k]] = row[k];
            }
            filteredData.push(obj);
        }
        
        return filteredData;
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
        var cfg = SyncEngine.getTool(toolKey);
        var headers = cfg.HEADERS;

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
        var totalCols = opts.totalCols || (cfg.HEADERS ? cfg.HEADERS.length : sheet.getLastColumn());
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
        var headers = cfg.HEADERS;
        if (lastRow >= 2) {
            sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
            if (cfg.FORMAT_CONFIG) {
                _App_applyBodyFormatting(sheet, 0, cfg.FORMAT_CONFIG);
            }
        }
    }

    /**
     * Returns only the values in the 'Action' column (Column A) for quick scanning.
     * Returns an array of strings.
     */
    function getActions(toolKey) {
        var sheet = getSheet(toolKey);
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];
        var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        return values.map(function(row) { return row[0]; });
    }

    function hasPendingActions(toolKey) {
        return getActions(toolKey).some(function(action) {
            return action !== '';
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
        
        // Find min and max row to fetch a single block
        var minRow = Math.min.apply(null, rowNumbers);
        var maxRow = Math.max.apply(null, rowNumbers);
        var numRows = maxRow - minRow + 1;
        
        var range = sheet.getRange(minRow, 1, numRows, lastCol);
        var data = range.getValues();
        var hasChanges = false;
        
        for (var i = 0; i < rowNumbers.length; i++) {
            var actualRow = rowNumbers[i];
            var relativeIdx = actualRow - minRow; // index in 'data' array
            var updates = updatesArray[i];
            
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
        }
        
        if (hasChanges) {
            range.setValues(data);
        }
    }

    function assertActiveSheet(toolKey) {
        var cfg = SyncEngine.getTool(toolKey);
        return _App_assertActiveSheet(cfg.SHEET_NAME);
    }

    function syncDynamicColumns(toolKey, dynamicHeaders, options) {
        return _App_syncDynamicColumns(toolKey, dynamicHeaders, options);
    }

    return {
        getSheet: getSheet,
        ensureSheet: ensureSheet,
        getHeaders: getHeaders,
        getHeaderMap: getHeaderMap,
        getNormalizedHeaderMap: getNormalizedHeaderMap,
        getSheetHeaderMap: getSheetHeaderMap,
        readObjects: readObjects,
        readPendingActions: readPendingActions,
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
