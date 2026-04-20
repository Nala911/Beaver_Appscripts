/**
 * DATA LAYER (DAO & MAPPER)
 * ==========================================
 * Manages spreadsheet interactions and data transformation.
 */

Object.assign(App.Data, (function() {

    function _normalizeHeaderKey(header) {
        return String(header || '').toUpperCase().trim();
    }

    /**
     * Retrieves the sheet for a given toolKey.
     */
    function getSheet(toolKey) {
        var cfg = App.Engine.getTool(toolKey);
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.SHEET_NAME);
        if (!sheet) throw new Error("App.Data: Sheet '" + cfg.SHEET_NAME + "' not found.");
        return sheet;
    }

    /**
     * Intelligent Data Mapper
     * Converts sheet rows (arrays) to objects and vice-versa with type casting.
     */
    var Mapper = {
        _castValue: function(value, type) {
            if (value === undefined || value === null || (typeof value === 'string' && value.trim() === '')) return null;
            switch (type) {
                case 'DATETIME':
                case 'DATE':
                    if (value instanceof Date) return value;
                    var d = new Date(value);
                    return isNaN(d.getTime()) ? value : d;
                case 'NUMBER':
                    return isNaN(parseFloat(value)) ? value : Number(value);
                case 'BOOLEAN':
                case 'CHECKBOX':
                    return String(value).toUpperCase() === 'TRUE' || value === true || value === 'Yes' || value === 'CHECKED';
                case 'EMAIL_LIST':
                case 'LIST':
                    return typeof value === 'string' ? value.split(',').map(function(e) { return e.trim(); }).filter(Boolean) : (Array.isArray(value) ? value : [value]);
                default:
                    return value;
            }
        },

        toObject: function(toolKey, rowData) {
            var cfg = App.Engine.getTool(toolKey);
            var schema = (cfg.FORMAT_CONFIG && cfg.FORMAT_CONFIG.COL_SCHEMA) ? cfg.FORMAT_CONFIG.COL_SCHEMA : [];
            var obj = {};
            for (var i = 0; i < cfg.HEADERS.length; i++) {
                obj[cfg.HEADERS[i]] = this._castValue(rowData[i], schema[i] ? schema[i].type : 'TEXT');
            }
            return obj;
        },

        castRow: function(toolKey, rawRow) {
            var cfg = App.Engine.getTool(toolKey);
            var schema = (cfg.FORMAT_CONFIG && cfg.FORMAT_CONFIG.COL_SCHEMA) ? cfg.FORMAT_CONFIG.COL_SCHEMA : [];
            var casted = {};
            Object.keys(rawRow).forEach(function(k) { if (k.indexOf('_') === 0) casted[k] = rawRow[k]; });
            schema.forEach(function(col) {
                casted[col.header] = Mapper._castValue(rawRow[col.header], col.type);
            });
            return casted;
        },

        toRow: function(toolKey, obj) {
            var cfg = App.Engine.getTool(toolKey);
            return cfg.HEADERS.map(function(h) {
                var val = obj[h];
                if (Array.isArray(val)) return val.join(',');
                if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                return val === null || val === undefined ? "" : val;
            });
        }
    };

    return {
        Mapper: Mapper,
        getSheet: getSheet,
        
        readObjects: function(toolKey) {
            var sheet = getSheet(toolKey);
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return [];

            var cfg = App.Engine.getTool(toolKey);
            var data = sheet.getRange(2, 1, lastRow - 1, cfg.HEADERS.length).getValues();

            return data.map(function(row) {
                return Mapper.toObject(toolKey, row);
            });
        },

        readPendingActions: function(toolKey) {
            var sheet = getSheet(toolKey);
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return [];

            var cfg = App.Engine.getTool(toolKey);
            var actionValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            var allData = sheet.getRange(2, 1, lastRow - 1, cfg.HEADERS.length).getValues();
            
            var results = [];
            for (var i = 0; i < actionValues.length; i++) {
                if (actionValues[i][0]) {
                    var obj = Mapper.toObject(toolKey, allData[i]);
                    obj._rowNumber = i + 2;
                    results.push(obj);
                }
            }
            return results;
        },

        writeObjects: function(toolKey, objectsArray, startRow) {
            if (!objectsArray || objectsArray.length === 0) return;
            var sheet = getSheet(toolKey);
            var data2D = objectsArray.map(function(obj) {
                return Mapper.toRow(toolKey, obj);
            });
            var targetRow = startRow || Math.max(2, sheet.getLastRow() + 1);
            sheet.getRange(targetRow, 1, data2D.length, data2D[0].length).setValues(data2D);
        },

        patchRow: function(toolKey, rowNumber, updates) {
            var sheet = getSheet(toolKey);
            var cfg = App.Engine.getTool(toolKey);
            var headers = cfg.HEADERS;
            var range = sheet.getRange(rowNumber, 1, 1, headers.length);
            var rowData = range.getValues()[0];
            var obj = Mapper.toObject(toolKey, rowData);

            var changed = false;
            Object.keys(updates).forEach(function(k) {
                if (obj.hasOwnProperty(k) && obj[k] !== updates[k]) {
                    obj[k] = updates[k];
                    changed = true;
                }
            });

            if (changed) {
                range.setValues([Mapper.toRow(toolKey, obj)]);
            }
        },

        clearData: function(toolKey) {
            var sheet = getSheet(toolKey);
            var lastRow = sheet.getLastRow();
            var cfg = App.Engine.getTool(toolKey);
            if (lastRow >= 2) {
                sheet.getRange(2, 1, lastRow - 1, cfg.HEADERS.length).clearContent();
            }
        },

        hasPendingActions: function(toolKey) {
            var sheet = getSheet(toolKey);
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return false;
            var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            return values.some(function(r) { return r[0] !== ""; });
        },

        getSheetHeaderMap: function(sheet) {
            var lastCol = sheet.getLastColumn();
            if (lastCol === 0) return {};
            var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
            var map = {};
            headers.forEach(function(h, i) {
                if (h) map[_normalizeHeaderKey(h)] = i;
            });
            return map;
        },

        // Backward Compatibility
        ensureSheet: function(toolKey, callback) { return _App_ensureSheetExists(toolKey, callback); }
    };
})());

// Backward Compatibility Layer
var SheetManager = App.Data;
