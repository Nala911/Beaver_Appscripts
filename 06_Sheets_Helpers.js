/**
 * DATA HELPERS (Sheet Scaffolding)
 * ==========================================
 * Low-level spreadsheet operations and validation.
 */

App.Data.Helpers = (function() {
    return {
        canScaffold: function(toolConfig) {
            return !!(toolConfig && toolConfig.HEADERS && toolConfig.HEADERS.length);
        },
        assertActive: function(expectedName) {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            if (sheet.getName() !== expectedName) {
                throw new Error("⚠️ Please run from '" + expectedName + "' sheet.");
            }
            return sheet;
        },
        validateActive: function(expectedName) {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            if (sheet.getName() !== expectedName) {
                return { valid: false, message: "⚠️ Please run from '" + expectedName + "'." };
            }
            return { valid: true, sheet: sheet };
        }
    };
})();

// Backward Compatibility Aliases
function _App_canScaffoldSheet(c) { return App.Data.Helpers.canScaffold(c); }
function _App_assertActiveSheet(n) { return App.Data.Helpers.assertActive(n); }
function _App_validateActiveSheet(n) { return App.Data.Helpers.validateActive(n); }
