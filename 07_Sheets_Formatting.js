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
        applyHeader: function(sheet, headers) {
            if (!headers || headers.length === 0) return;
            var theme = SHEET_THEME;
            var range = sheet.getRange(1, 1, 1, headers.length);
            
            range.setValues([headers])
                .setFontWeight(theme.LAYOUT.HEADER_WEIGHT || 'bold')
                .setFontColor(theme.TEXT)
                .setBackground(theme.HEADER)
                .setHorizontalAlignment(theme.LAYOUT.HEADER_ALIGN_H || 'center')
                .setVerticalAlignment(theme.LAYOUT.HEADER_ALIGN_V || 'middle')
                .setFontFamily(theme.FONTS.PRIMARY)
                .setFontSize(theme.SIZES.HEADER || 11);

            range.setBorder(true, true, true, true, true, true, theme.BORDER, theme.BORDER_STYLE);
            
            sheet.setRowHeight(1, theme.LAYOUT.HEADER_ROW_HEIGHT || 45);
        },

        applyBody: function(sheet, numDataRows, config) {
            var theme = SHEET_THEME;
            var actualRows = Math.min(numDataRows + BUFFER_ROWS, sheet.getMaxRows() - 1);
            if (actualRows < 1) return;

            var totalCols = config.HEADERS ? config.HEADERS.length : sheet.getLastColumn();
            var range = sheet.getRange(2, 1, actualRows, totalCols);
            
            range.setFontFamily(theme.FONTS.PRIMARY)
                 .setFontSize(theme.SIZES.BODY || 10)
                 .setVerticalAlignment(theme.LAYOUT.BODY_ALIGN_V || 'middle')
                 .setWrapStrategy(theme.LAYOUT.BODY_WRAP);

            if (config.COL_SCHEMA) {
                config.COL_SCHEMA.forEach(function(col, i) {
                    var colRange = sheet.getRange(2, i + 1, actualRows, 1);
                    
                    if (col.type === 'DATETIME') colRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
                    if (col.type === 'DATE') colRange.setNumberFormat('MM/dd/yyyy');
                    if (col.type === 'ID') colRange.setFontFamily(theme.FONTS.MONOSPACE);
                    
                    if (col.type === 'ACTION' || col.type === 'DROPDOWN') {
                        var opts = typeof col.options === 'function' ? col.options() : col.options;
                        if (opts) {
                            var rule = SpreadsheetApp.newDataValidation().requireValueInList(opts).build();
                            sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
                        }
                    }
                });
            }
            
            if (actualRows > 0) {
                sheet.setRowHeights(2, actualRows, theme.LAYOUT.BODY_ROW_HEIGHT || 35);
            }
        }
    };
})();

// Backward Compatibility Aliases
function _App_applyHeaderFormatting(s, h) { return App.UI.Formatting.applyHeader(s, h); }
function _App_applyBodyFormatting(s, n, c) { return App.UI.Formatting.applyBody(s, n, c); }
function _App_getColumnLetter(c) { return App.UI.Formatting.getColumnLetter(c); }
