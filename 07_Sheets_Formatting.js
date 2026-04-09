// ==========================================
// Centralized Body Formatting Utility
// ==========================================

// Extra rows formatted beyond actual data to cover manual row additions.
var FORMATTING_BUFFER_ROWS = 30;

function _App_getColumnLetter(col) {
    var temp, letter = '';
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}

function _App_applyHeaderFormatting(sheet, headers) {
    if (!headers || headers.length === 0) return;

    sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight(SHEET_THEME.LAYOUT.HEADER_WEIGHT)
        .setFontSize(SHEET_THEME.SIZES.HEADER)
        .setFontFamily(SHEET_THEME.FONTS.PRIMARY)
        .setBackground(SHEET_THEME.HEADER)
        .setFontColor(SHEET_THEME.TEXT)
        .setFontStyle(SHEET_THEME.LAYOUT.HEADER_FONT_STYLE)
        .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE)
        .setVerticalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_V)
        .setHorizontalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_H);
    sheet.setRowHeight(1, SHEET_THEME.LAYOUT.HEADER_ROW_HEIGHT);
}

function _App_cloneFormatConfig_(config) {
    if (!config) return null;

    var clone = {};
    Object.keys(config).forEach(function(key) {
        var value = config[key];
        if (key === 'COL_SCHEMA' || key === 'conditionalRules') {
            clone[key] = (value || []).map(function(item) {
                var out = {};
                Object.keys(item).forEach(function(itemKey) {
                    out[itemKey] = item[itemKey];
                });
                return out;
            });
        } else {
            clone[key] = value;
        }
    });
    return clone;
}

function _App_buildRuntimeToolShape(toolKey, dynamicHeaders, options) {
    var cfg = SyncEngine.getTool(toolKey);
    var runtimeHeaders = (cfg.HEADERS || []).slice();
    var runtimeWidths = (cfg.COL_WIDTHS || []).slice();
    var runtimeFormat = _App_cloneFormatConfig_(cfg.FORMAT_CONFIG);
    var headersToInsert = dynamicHeaders || [];
    var dynamicSchemaFactory = options && options.dynamicSchemaFactory;
    var dynamicColWidth = options && options.dynamicColWidth !== undefined ? options.dynamicColWidth : 150;
    var anchorHeader = options && options.anchorHeader;
    var insertIndex = runtimeHeaders.length;

    if (anchorHeader) {
        insertIndex = runtimeHeaders.indexOf(anchorHeader);
        if (insertIndex === -1) {
            throw new Error("Anchor header '" + anchorHeader + "' was not found for tool '" + toolKey + "'.");
        }
    }

    var schemaItems = headersToInsert.map(function(header) {
        if (typeof dynamicSchemaFactory === 'function') {
            return dynamicSchemaFactory(header);
        }
        return { header: header, type: 'TEXT' };
    });

    if (headersToInsert.length > 0) {
        Array.prototype.splice.apply(runtimeHeaders, [insertIndex, 0].concat(headersToInsert));
        Array.prototype.splice.apply(runtimeWidths, [insertIndex, 0].concat(headersToInsert.map(function() { return dynamicColWidth; })));
        if (runtimeFormat && runtimeFormat.COL_SCHEMA) {
            Array.prototype.splice.apply(runtimeFormat.COL_SCHEMA, [insertIndex, 0].concat(schemaItems));
            runtimeFormat.totalCols = runtimeFormat.COL_SCHEMA.length;
        }
    }

    return {
        headers: runtimeHeaders,
        widths: runtimeWidths,
        formatConfig: runtimeFormat
    };
}

function _App_syncDynamicColumns(toolKey, dynamicHeaders, options) {
    var cfg = SyncEngine.getTool(toolKey);
    var uniqueDynamicHeaders = [];
    (dynamicHeaders || []).forEach(function(header) {
        var normalized = String(header || '').trim();
        if (normalized && uniqueDynamicHeaders.indexOf(normalized) === -1) {
            uniqueDynamicHeaders.push(normalized);
        }
    });

    var sheet = _App_ensureSheetExists(toolKey);
    var runtimeShape = _App_buildRuntimeToolShape(toolKey, uniqueDynamicHeaders, options);
    var currentHeaderCount = sheet.getLastColumn();

    if (currentHeaderCount > runtimeShape.headers.length) {
        sheet.deleteColumns(runtimeShape.headers.length + 1, currentHeaderCount - runtimeShape.headers.length);
    } else if (currentHeaderCount < runtimeShape.headers.length) {
        sheet.insertColumnsAfter(Math.max(currentHeaderCount, 1), runtimeShape.headers.length - currentHeaderCount);
    }

    _App_applyHeaderFormatting(sheet, runtimeShape.headers);

    runtimeShape.widths.forEach(function(width, index) {
        if (width !== null && width !== undefined) {
            sheet.setColumnWidth(index + 1, width);
        }
    });

    if (cfg.FROZEN_ROWS > 0) sheet.setFrozenRows(cfg.FROZEN_ROWS);
    if (cfg.FROZEN_COLS > 0) sheet.setFrozenColumns(cfg.FROZEN_COLS);

    if (runtimeShape.formatConfig) {
        var numRows = Math.max(sheet.getLastRow() - 1, 0);
        _App_applyBodyFormatting(sheet, numRows, runtimeShape.formatConfig);
    }

    return {
        headers: runtimeShape.headers,
        dynamicHeaders: uniqueDynamicHeaders,
        sheet: sheet,
        formatConfig: runtimeShape.formatConfig
    };
}

/**
 * Applies standardized body formatting to a sheet's data area.
 * This enforces strict column ordering:
 * 1: Action (SHEET_THEME.ACTION) - unless config.skipActionColoring is true
 * 2 onwards: Editable data (SHEET_THEME.EDITABLE)
 * Last N columns: Read-only data (SHEET_THEME.READ_ONLY)
 */
function _App_applyBodyFormatting(sheet, numDataRows, config) {
    var rowsToFormat = numDataRows + FORMATTING_BUFFER_ROWS;
    var maxRows = sheet.getMaxRows();
    var actualRows = Math.min(rowsToFormat, maxRows - 1);
    if (actualRows < 1) return;

    var totalCols = config.COL_SCHEMA ? config.COL_SCHEMA.length : (config.totalCols || sheet.getLastColumn());
    var numReadOnlyAtEnd = config.numReadOnlyColsAtEnd || 0;

    // 1. Base formatting
    var startRow = 2;
    var endCol = Math.max(totalCols, 1);

    var dataRange = sheet.getRange(startRow, 1, actualRows, totalCols);
    dataRange
        .setFontColor(SHEET_THEME.TEXT)
        .setFontFamily(SHEET_THEME.FONTS.PRIMARY)
        .setFontSize(SHEET_THEME.SIZES.BODY)
        .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE)
        .setHorizontalAlignment(SHEET_THEME.LAYOUT.BODY_ALIGN_H)
        .setVerticalAlignment(SHEET_THEME.LAYOUT.BODY_ALIGN_V)
        .setWrapStrategy(SHEET_THEME.LAYOUT.BODY_WRAP);

    sheet.setRowHeights(startRow, actualRows, SHEET_THEME.LAYOUT.BODY_ROW_HEIGHT);

    try {
        var startColForEditable = 1;

        // 1. Action (Col 1)
        if (!config.skipActionColoring) {
            sheet.getRange(startRow, 1, actualRows, 1).setBackground(SHEET_THEME.ACTION);
            startColForEditable = 2;
        }

        // 2. Editable Columns
        var numEditable = endCol - (startColForEditable - 1) - numReadOnlyAtEnd;
        if (numEditable > 0) {
            sheet.getRange(startRow, startColForEditable, actualRows, numEditable).setBackground(SHEET_THEME.EDITABLE);
        }

        // 3. Read-Only Columns
        if (numReadOnlyAtEnd > 0) {
            var readOnlyStartCol = endCol - numReadOnlyAtEnd + 1;
            sheet.getRange(startRow, readOnlyStartCol, actualRows, numReadOnlyAtEnd).setBackground(SHEET_THEME.READ_ONLY);
        }
    } catch (e) {
        console.error("Error applying column colors:", e);
    }

    // Apply Schema-driven validations and formats
    var validationRows = maxRows - 1;
    if (config.COL_SCHEMA) {
        config.COL_SCHEMA.forEach(function(colDef, index) {
            var colNum = index + 1;
            var range = sheet.getRange(startRow, colNum, actualRows, 1);
            var valRange = sheet.getRange(startRow, colNum, validationRows, 1);
            
            // Fonts
            if (colDef.type === 'ID' || colDef.type === 'URL') {
                range.setFontFamily(SHEET_THEME.FONTS.MONOSPACE);
            }
            if (colDef.type === 'URL' || colDef.italic) {
                range.setFontStyle('italic');
            }

            // Number Formats
            if (colDef.type === 'DATETIME') {
                range.setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');
            } else if (colDef.type === 'DATE') {
                range.setNumberFormat('MM/dd/yyyy');
            } else if (colDef.type === 'ID' || colDef.type === 'TEXT') {
                range.setNumberFormat('@'); // Force Plain Text
            }

            // Validations
            var rule = null;
            if (colDef.type === 'ACTION' || colDef.type === 'DROPDOWN') {
                var opts = typeof colDef.options === 'function' ? colDef.options() : colDef.options;
                if (opts && opts.length > 0) {
                    rule = SpreadsheetApp.newDataValidation().requireValueInList(opts, true).setAllowInvalid(colDef.allowInvalid || false).build();
                }
            } else if (colDef.type === 'CHECKBOX') {
                rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
            } else if (colDef.type === 'EMAIL' || colDef.type === 'EMAIL_LIST') {
                var letter = _App_getColumnLetter(colNum);
                var re = colDef.type === 'EMAIL' ? 'ISEMAIL(' + letter + '2)' : 'REGEXMATCH(' + letter + '2, "^[\\\\w\\\\.\\\\-@\\\\s,]+$")';
                var formula = '=OR(ISBLANK(' + letter + '2), ' + re + ')';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter valid email(s).').setAllowInvalid(true).build();
            }

            if (rule) {
                valRange.setDataValidation(rule);
            }
        });
    }

    // 6. Conditional formatting rules
    _App_applyConditionalRules(sheet, actualRows, totalCols, config.conditionalRules || []);
}

/**
 * Builds and applies conditional formatting rules from a declarative descriptor array.
 * Replaces ALL existing conditional formatting rules on the sheet.
 *
 * Supported rule types: 'success', 'error', 'errorCross', 'pending', 'synced', 'custom'
 * Supported scopes: 'fullRow' (default), 'actionOnly', 'statusOnly'
 */
function _App_applyConditionalRules(sheet, numRows, totalCols, ruleDescriptors) {
    var rules = [];
    var fullRange = sheet.getRange(2, 1, numRows, totalCols);

    ruleDescriptors.forEach(function (desc) {
        var targetRange;
        if (desc.scope === 'actionOnly' && desc.actionCol) {
            var actionColNum = desc.actionCol.charCodeAt(0) - 64; // 'A' → 1
            targetRange = sheet.getRange(2, actionColNum, numRows, 1);
        } else if (desc.scope === 'statusOnly' && desc.statusCol) {
            var statusColNum = desc.statusCol.charCodeAt(0) - 64;
            targetRange = sheet.getRange(2, statusColNum, numRows, 1);
        } else if (desc.scope === 'custom_col' && desc.col) {
            targetRange = sheet.getRange(2, desc.col, numRows, 1);
        } else {
            targetRange = fullRange; // 'fullRow'
        }

        var rule;
        if (desc.type === 'success') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "✅")')
                .setBackground(SHEET_THEME.STATUS.SUCCESS)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'error') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "⚠️")')
                .setBackground(SHEET_THEME.STATUS.ERROR)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'errorCross') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "❌")')
                .setBackground(SHEET_THEME.STATUS.ERROR)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'pending') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=$' + desc.actionCol + '2<>""')
                .setBackground(SHEET_THEME.STATUS.PENDING)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'synced') {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied('=REGEXMATCH($' + desc.statusCol + '2, "📝")')
                .setBackground(SHEET_THEME.STATUS.SYNCED)
                .setRanges([targetRange]).build();
        } else if (desc.type === 'custom' && desc.formula) {
            rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(desc.formula)
                .setBackground(desc.color)
                .setRanges([targetRange]).build();
        }

        if (rule) rules.push(rule);
    });

    sheet.setConditionalFormatRules(rules);
}
