/**
 * UI FORMATTING LAYER
 * ==========================================
 * Handles spreadsheet styling, headers, and conditional formatting.
 */

App.UI.Formatting = (function() {
    var BUFFER_ROWS = 30;

    function getColumnLetter(col) {
        var temp, letter = '';
        while (col > 0) {
            temp = (col - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            col = (col - temp - 1) / 26;
        }
        return letter;
    }

    return {
        getColumnLetter: getColumnLetter,

        applyHeader: function(sheet, headers) {
            if (!sheet || !headers || headers.length === 0) return;
            
            try {
                var theme = globalThis.SHEET_THEME;
                var range = sheet.getRange(1, 1, 1, headers.length);
                
                range.setValues([headers])
                    .setFontWeight(theme.LAYOUT.HEADER_WEIGHT || 'bold')
                    .setFontColor(theme.TEXT || '#ffffff')
                    .setBackground(theme.HEADER || '#424242')
                    .setHorizontalAlignment(theme.LAYOUT.HEADER_ALIGN_H || 'center')
                    .setVerticalAlignment(theme.LAYOUT.HEADER_ALIGN_V || 'middle')
                    .setFontFamily(theme.FONTS.PRIMARY || 'Roboto')
                    .setFontSize(theme.SIZES.HEADER || 11);

                range.setBorder(true, true, true, true, true, true, theme.BORDER || '#ffffff', theme.BORDER_STYLE || SpreadsheetApp.BorderStyle.SOLID);
                
                sheet.setRowHeight(1, theme.LAYOUT.HEADER_ROW_HEIGHT || 45);
            } catch (err) {
                console.error("Error applying header formatting:", err, err.stack);
            }
        },

        applyBody: function(sheet, numDataRows, config) {
            if (!sheet || !config) return;

            try {
                var theme = globalThis.SHEET_THEME;
                var maxRows = sheet.getMaxRows();
                var actualRows = Math.min(numDataRows + BUFFER_ROWS, maxRows - 1);
                
                if (actualRows < 1) return;

                var totalCols = (config.COL_SCHEMA && config.COL_SCHEMA.length) ? config.COL_SCHEMA.length : sheet.getLastColumn();
                if (totalCols < 1) return;

                var range = sheet.getRange(2, 1, actualRows, totalCols);
                
                range.setFontFamily(theme.FONTS.PRIMARY || 'Roboto')
                     .setFontSize(theme.SIZES.BODY || 10)
                     .setVerticalAlignment(theme.LAYOUT.BODY_ALIGN_V || 'middle')
                     .setWrapStrategy(theme.LAYOUT.BODY_WRAP || SpreadsheetApp.WrapStrategy.CLIP);

                if (config.COL_SCHEMA) {
                    config.COL_SCHEMA.forEach(function(col, i) {
                        if (i >= totalCols) return;
                        var colRange = sheet.getRange(2, i + 1, actualRows, 1);
                        
                        if (col.type === 'DATETIME') colRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
                        else if (col.type === 'DATE') colRange.setNumberFormat('MM/dd/yyyy');
                        else if (col.type === 'ID' || col.type === 'URL') colRange.setFontFamily(theme.FONTS.MONOSPACE || 'Consolas');
                        
                        var bgColor = theme.EDITABLE || '#528dab';
                        if (col.type === 'ACTION') bgColor = theme.ACTION || '#2e5a70';
                        else if (col.type === 'READ_ONLY' || i >= (totalCols - (config.numReadOnlyColsAtEnd || 0))) bgColor = theme.READ_ONLY || '#655356';
                        colRange.setBackground(bgColor).setFontColor(theme.TEXT || '#ffffff');
                        
                        if (col.type === 'ACTION' || col.type === 'DROPDOWN') {
                            try {
                                var opts = typeof col.options === 'function' ? col.options() : col.options;
                                if (opts && Array.isArray(opts) && opts.length > 0) {
                                    var rule = SpreadsheetApp.newDataValidation().requireValueInList(opts).build();
                                    sheet.getRange(2, i + 1, maxRows - 1, 1).setDataValidation(rule);
                                }
                            } catch (optErr) {
                                console.warn("Validation error in column " + (i+1) + ":", optErr);
                            }
                        }
                        
                        if (col.type === 'CHECKBOX') {
                            sheet.getRange(2, i + 1, maxRows - 1, 1).insertCheckboxes();
                        }
                    });
                }
                
                if (actualRows > 0) {
                    sheet.setRowHeights(2, actualRows, theme.LAYOUT.BODY_ROW_HEIGHT || 35);
                }
                
                // Optional: Action Column Highlight
                if (config.conditionalRules) {
                   this.applyConditionalRules(sheet, config.conditionalRules, totalCols);
                }

            } catch (err) {
                console.error("Error applying body formatting:", err, err.stack);
            }
        },

        applyConditionalRules: function(sheet, rules, totalCols) {
            if (!sheet || !rules || rules.length === 0) return;
            var theme = globalThis.SHEET_THEME;
            var maxRows = sheet.getMaxRows();
            var sheetRules = sheet.getConditionalFormatRules();

            rules.forEach(function(ruleDef) {
                var range = null;
                if (ruleDef.scope === 'actionOnly' && ruleDef.actionCol) {
                    var colIdx = (typeof ruleDef.actionCol === 'number') ? ruleDef.actionCol : 1; 
                    range = sheet.getRange(2, colIdx, maxRows - 1, 1);
                } else {
                    range = sheet.getRange(2, 1, maxRows - 1, totalCols);
                }

                var rule = null;
                if (ruleDef.type === 'pending') {
                    rule = SpreadsheetApp.newConditionalFormatRule()
                        .whenFormulaSatisfied("=LEN($A2)>0")
                        .setBackground(theme.STATUS.PENDING || '#f59e0b')
                        .setFontColor('#ffffff')
                        .setRanges([range])
                        .build();
                } else if (ruleDef.type === 'custom' && ruleDef.formula) {
                     rule = SpreadsheetApp.newConditionalFormatRule()
                        .whenFormulaSatisfied(ruleDef.formula)
                        .setBackground(ruleDef.color || '#eeeeee')
                        .setRanges([range])
                        .build();
                }

                if (rule) sheetRules.push(rule);
            });

            sheet.setConditionalFormatRules(sheetRules);
        }
    };
})();

// Backward Compatibility Aliases
function _App_applyHeaderFormatting(s, h) { return App.UI.Formatting.applyHeader(s, h); }
function _App_applyBodyFormatting(s, n, c) { return App.UI.Formatting.applyBody(s, n, c); }
function _App_getColumnLetter(c) { return App.UI.Formatting.getColumnLetter(c); }
