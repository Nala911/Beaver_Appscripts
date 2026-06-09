// ==========================================
// Centralized Body Formatting Utility
// ==========================================

// Extra rows formatted beyond actual data to cover manual row additions.
var FORMATTING_BUFFER_ROWS = 30;



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
        _App_applyBodyFormatting(sheet, numRows, runtimeShape.formatConfig, true);
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
 * - First Columns (Action/Status): SHEET_THEME.FIRST_COLS_COLOR
 * - Middle Columns (Editable Data): SHEET_THEME.MIDDLE_COLS_COLOR
 * - Last Columns (Read-only/IDs): SHEET_THEME.LAST_COLS_COLOR
 */
function _App_applyBodyFormatting(sheet, numDataRows, config, forceConditional) {
    var rowsToFormat = numDataRows + FORMATTING_BUFFER_ROWS;
    var maxRows = sheet.getMaxRows();
    var actualRows = Math.min(rowsToFormat, maxRows - 1);
    if (actualRows < 1) return;

    var totalCols = config.COL_SCHEMA ? config.COL_SCHEMA.length : (config.totalCols || sheet.getLastColumn());

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

    // Apply Schema-driven validations and formats in batches
    if (config.COL_SCHEMA) {
        var colFontFamilies = [];
        var colFontStyles = [];
        var colBackgrounds = [];
        var colNumberFormats = [];
        var colValidations = [];

        config.COL_SCHEMA.forEach(function(colDef, index) {
            var colNum = index + 1;
            
            // Fonts
            var fontFamily = SHEET_THEME.FONTS.PRIMARY;
            if (colDef.type === 'ID' || colDef.type === 'URL') {
                fontFamily = SHEET_THEME.FONTS.MONOSPACE;
            }
            colFontFamilies.push(fontFamily);

            var fontStyle = 'normal';
            if (colDef.type === 'URL' || colDef.italic) {
                fontStyle = 'italic';
            }
            colFontStyles.push(fontStyle);
            
            // Background Colors (Schema-driven Categorization)
            var category = colDef.category;
            if (!category) {
                if (colDef.type === 'ACTION' || colDef.type === 'STATUS') category = 'FIRST_COLS';
                else if (colDef.type === 'READ_ONLY' || colDef.type === 'ID') category = 'LAST_COLS';
                else category = 'MIDDLE_COLS';
            }

            var bg = SHEET_THEME.MIDDLE_COLS_COLOR;
            if (category === 'FIRST_COLS') {
                bg = SHEET_THEME.FIRST_COLS_COLOR;
            } else if (category === 'LAST_COLS') {
                bg = SHEET_THEME.LAST_COLS_COLOR;
            }
            colBackgrounds.push(bg);

            // Number Formats
            var numFormat = '@'; // Force Plain Text by default
            if (colDef.type === 'DATETIME') {
                numFormat = 'MM/dd/yyyy hh:mm:ss AM/PM';
            } else if (colDef.type === 'DATE') {
                numFormat = 'MM/dd/yyyy';
            } else if (colDef.type === 'ID' || colDef.type === 'TEXT') {
                numFormat = '@';
            } else {
                numFormat = '';
            }
            colNumberFormats.push(numFormat);

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
            } else if (colDef.type === 'DATE' || colDef.type === 'DATETIME') {
                rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Enter a valid date.').build();
            } else if (colDef.type === 'URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), REGEXMATCH(' + letter + '2, "^https?:\\/\\/"))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid URL starting with http:// or https://.').setAllowInvalid(true).build();
            } else if (colDef.type === 'DOCS_URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), REGEXMATCH(' + letter + '2, "docs\\.google\\.com\\/document"))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid Google Docs URL.').setAllowInvalid(true).build();
            } else if (colDef.type === 'DRIVE_URL') {
                var letter = _App_getColumnLetter(colNum);
                var formula = '=OR(ISBLANK(' + letter + '2), OR(REGEXMATCH(' + letter + '2, "drive\\.google\\.com"), REGEXMATCH(' + letter + '2, "docs\\.google\\.com")))';
                rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setHelpText('Enter a valid Google Drive or Docs URL.').setAllowInvalid(true).build();
            }
            colValidations.push(rule);
        });

        // Build 2D formatting grids in memory
        var fontFamilies2D = [];
        var fontStyles2D = [];
        var backgrounds2D = [];
        var numberFormats2D = [];
        var validations2D = [];

        for (var r = 0; r < actualRows; r++) {
            fontFamilies2D.push(colFontFamilies);
            fontStyles2D.push(colFontStyles);
            backgrounds2D.push(colBackgrounds);
            numberFormats2D.push(colNumberFormats);
            validations2D.push(colValidations);
        }

        // Apply formatting grids in single batch calls
        dataRange.setFontFamilies(fontFamilies2D);
        dataRange.setFontStyles(fontStyles2D);
        dataRange.setBackgrounds(backgrounds2D);
        dataRange.setNumberFormats(numberFormats2D);
        dataRange.setDataValidations(validations2D);
    }

    // 6. Conditional formatting rules
    _App_applyConditionalRules(sheet, actualRows, totalCols, config.conditionalRules || [], forceConditional);
}

/**
 * Builds and applies conditional formatting rules from a declarative descriptor array.
 * Replaces ALL existing conditional formatting rules on the sheet.
 *
 * Supported rule types: 'success', 'error', 'errorCross', 'pending', 'synced', 'custom'
 * Supported scopes: 'fullRow' (default), 'actionOnly', 'statusOnly'
 */
function _App_applyConditionalRules(sheet, numRows, totalCols, ruleDescriptors, force) {
    if (!force) {
        var existingRules = sheet.getConditionalFormatRules();
        if (existingRules && existingRules.length > 0) {
            // Bypass clearing and setting rules if already present to optimize execution speed
            return;
        }
    }

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
                .setBackground(SHEET_THEME.STATUS.WARNING)
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
